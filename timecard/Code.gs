/**
 * 従業員タイムカードアプリ(Google Apps Script)
 *
 * 勤務ルール:
 *  - 定時: 8:00〜17:00
 *  - 休憩: 12:00〜13:00(1時間)→ 所定内労働は最大8時間
 *  - 8:00より前の出勤 = 早出(15分単位・切り捨て)
 *  - 17:00以降の退勤 = 残業(15分単位・切り捨て)
 *  - 毎月1日に前月分の就業時間を自動集計
 */

var TIMEZONE = 'Asia/Tokyo';
var SHEET_RECORDS = '打刻記録';
var SHEET_SUMMARY = '月次集計';

var WORK_START_MIN = 8 * 60;    // 8:00
var WORK_END_MIN = 17 * 60;     // 17:00
var BREAK_START_MIN = 12 * 60;  // 12:00
var BREAK_END_MIN = 13 * 60;    // 13:00
var ROUND_UNIT_MIN = 15;        // 残業・早出の丸め単位(分)

var RECORD_HEADERS = ['日付', '氏名', '出勤時刻', '退勤時刻', '所定内(時間)', '早出(時間)', '残業(時間)', '就業時間(時間)'];
var SUMMARY_HEADERS = ['年月', '氏名', '出勤日数', '所定内(時間)', '早出(時間)', '残業(時間)', '就業時間合計(時間)'];

/**
 * 初回セットアップ:シートの作成と月次集計トリガーの登録。
 * Apps Scriptエディタからこの関数を一度実行してください。
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var records = ss.getSheetByName(SHEET_RECORDS);
  if (!records) {
    records = ss.insertSheet(SHEET_RECORDS);
    records.appendRow(RECORD_HEADERS);
    records.setFrozenRows(1);
  }

  var summary = ss.getSheetByName(SHEET_SUMMARY);
  if (!summary) {
    summary = ss.insertSheet(SHEET_SUMMARY);
    summary.appendRow(SUMMARY_HEADERS);
    summary.setFrozenRows(1);
  }

  // 既存の月次トリガーを削除してから登録(二重登録防止)
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'monthlyAggregate') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('monthlyAggregate')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();
}

/** Webアプリのエントリポイント */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('タイムカード')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 出勤打刻 */
function punchIn(name) {
  name = (name || '').trim();
  if (!name) return { ok: false, message: '氏名を入力してください。' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getRecordsSheet_();
    var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
    var now = Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm');

    if (findTodayRow_(sheet, today, name) !== -1) {
      return { ok: false, message: '本日はすでに出勤打刻済みです。' };
    }

    sheet.appendRow([today, name, now, '', '', '', '', '']);
    return { ok: true, message: name + ' さん、出勤打刻しました(' + now + ')。' };
  } finally {
    lock.releaseLock();
  }
}

/** 退勤打刻(早出・残業・就業時間を計算してシートに記録) */
function punchOut(name) {
  name = (name || '').trim();
  if (!name) return { ok: false, message: '氏名を入力してください。' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getRecordsSheet_();
    var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
    var now = Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm');

    var row = findTodayRow_(sheet, today, name);
    if (row === -1) {
      return { ok: false, message: '本日の出勤打刻が見つかりません。先に出勤打刻をしてください。' };
    }
    if (sheet.getRange(row, 4).getValue() !== '') {
      return { ok: false, message: '本日はすでに退勤打刻済みです。' };
    }

    var clockIn = String(sheet.getRange(row, 3).getDisplayValue());
    var result = calcWorkHours_(clockIn, now);

    sheet.getRange(row, 4, 1, 5).setValues([[
      now, result.regularHours, result.earlyHours, result.overtimeHours, result.totalHours
    ]]);

    var msg = name + ' さん、退勤打刻しました(' + now + ')。\n' +
      '所定内 ' + result.regularHours + 'h / 早出 ' + result.earlyHours + 'h / 残業 ' + result.overtimeHours + 'h / 合計 ' + result.totalHours + 'h';
    return { ok: true, message: msg };
  } finally {
    lock.releaseLock();
  }
}

/** 本日の打刻状況を取得(画面表示用) */
function getStatus(name) {
  name = (name || '').trim();
  if (!name) return { punchedIn: false, punchedOut: false };

  var sheet = getRecordsSheet_();
  var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  var row = findTodayRow_(sheet, today, name);
  if (row === -1) return { punchedIn: false, punchedOut: false };

  return {
    punchedIn: true,
    punchedOut: sheet.getRange(row, 4).getDisplayValue() !== '',
    clockIn: sheet.getRange(row, 3).getDisplayValue(),
    clockOut: sheet.getRange(row, 4).getDisplayValue()
  };
}

/**
 * 月次集計(毎月1日のトリガーで実行)。
 * 前月分の打刻記録を従業員ごとに集計して「月次集計」シートへ書き込む。
 */
function monthlyAggregate() {
  var now = new Date();
  var prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  aggregateMonth(prev.getFullYear(), prev.getMonth() + 1);
}

/**
 * 指定した年月を集計する(手動実行も可能)。
 * 例: aggregateMonth(2026, 5) → 2026年5月分を集計
 */
function aggregateMonth(year, month) {
  var prefix = year + '-' + ('0' + month).slice(-2); // 'yyyy-MM'
  var records = getRecordsSheet_();
  var data = records.getDataRange().getDisplayValues();

  var byName = {};
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0];
    if (rowDate.indexOf(prefix) !== 0) continue;
    if (data[i][3] === '') continue; // 退勤打刻なしの行は対象外

    var nm = data[i][1];
    if (!byName[nm]) byName[nm] = { days: 0, regular: 0, early: 0, overtime: 0, total: 0 };
    byName[nm].days += 1;
    byName[nm].regular += Number(data[i][4]) || 0;
    byName[nm].early += Number(data[i][5]) || 0;
    byName[nm].overtime += Number(data[i][6]) || 0;
    byName[nm].total += Number(data[i][7]) || 0;
  }

  var summary = getSummarySheet_();

  // 同じ年月の既存集計行を削除(再実行時の重複防止)
  var sData = summary.getDataRange().getDisplayValues();
  for (var r = sData.length - 1; r >= 1; r--) {
    if (sData[r][0] === prefix) summary.deleteRow(r + 1);
  }

  Object.keys(byName).sort().forEach(function (nm) {
    var s = byName[nm];
    summary.appendRow([
      prefix, nm, s.days,
      round2_(s.regular), round2_(s.early), round2_(s.overtime), round2_(s.total)
    ]);
  });
}

/**
 * 出退勤時刻から各労働時間を計算する。
 *  - 所定内: [出勤, 退勤] と [8:00, 17:00] の重なりから休憩(12:00〜13:00)の重なりを引く
 *  - 早出: 8:00より前の分数を15分単位で切り捨て
 *  - 残業: 17:00以降の分数を15分単位で切り捨て
 */
function calcWorkHours_(clockIn, clockOut) {
  var inMin = toMinutes_(clockIn);
  var outMin = toMinutes_(clockOut);
  if (outMin < inMin) outMin = inMin;

  var regularMin = overlap_(inMin, outMin, WORK_START_MIN, WORK_END_MIN) -
    overlap_(inMin, outMin, BREAK_START_MIN, BREAK_END_MIN);
  if (regularMin < 0) regularMin = 0;

  var earlyRaw = Math.max(0, WORK_START_MIN - inMin);
  var earlyMin = Math.floor(earlyRaw / ROUND_UNIT_MIN) * ROUND_UNIT_MIN;

  var overtimeRaw = Math.max(0, outMin - WORK_END_MIN);
  var overtimeMin = Math.floor(overtimeRaw / ROUND_UNIT_MIN) * ROUND_UNIT_MIN;

  return {
    regularHours: round2_(regularMin / 60),
    earlyHours: round2_(earlyMin / 60),
    overtimeHours: round2_(overtimeMin / 60),
    totalHours: round2_((regularMin + earlyMin + overtimeMin) / 60)
  };
}

// ---- 内部ユーティリティ ----

function getRecordsSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECORDS);
  if (!sheet) throw new Error('「' + SHEET_RECORDS + '」シートがありません。setup() を実行してください。');
  return sheet;
}

function getSummarySheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SUMMARY);
  if (!sheet) throw new Error('「' + SHEET_SUMMARY + '」シートがありません。setup() を実行してください。');
  return sheet;
}

function findTodayRow_(sheet, today, name) {
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === today && data[i][1] === name) return i + 1;
  }
  return -1;
}

function toMinutes_(hhmm) {
  var parts = String(hhmm).split(':');
  return Number(parts[0]) * 60 + Number(parts[1]);
}

function overlap_(aStart, aEnd, bStart, bEnd) {
  return Math.max(0, Math.min(aEnd, bEnd) - Math.max(aStart, bStart));
}

function round2_(n) {
  return Math.round(n * 100) / 100;
}
