/* ============================================
   ケアプラン彩心 LP - JavaScript
   ============================================ */

document.addEventListener('DOMContentLoaded', () => {
  'use strict';

  // --- Scroll Animation (Intersection Observer) ---
  const animElements = document.querySelectorAll('.fade-in-up');

  if ('IntersectionObserver' in window) {
    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            entry.target.classList.add('visible');
            observer.unobserve(entry.target);
          }
        });
      },
      {
        threshold: 0.15,
        rootMargin: '0px 0px -40px 0px',
      }
    );

    animElements.forEach((el) => observer.observe(el));
  } else {
    // Fallback: show all
    animElements.forEach((el) => el.classList.add('visible'));
  }

  // --- Header scroll state ---
  const header = document.getElementById('header');
  const fixedFooterBar = document.getElementById('fixedFooterBar');
  const scrollTopBtn = document.getElementById('scrollTop');

  function onScroll() {
    const scrollY = window.scrollY;

    // Header shadow
    if (scrollY > 10) {
      header.classList.add('scrolled');
    } else {
      header.classList.remove('scrolled');
    }

    // Fixed footer bar: 常に表示（電話・LINEへいつでもアクセスできるように）
    fixedFooterBar.classList.add('visible');

    // Scroll top button
    if (scrollY > 600) {
      scrollTopBtn.classList.add('visible');
    } else {
      scrollTopBtn.classList.remove('visible');
    }
  }

  window.addEventListener('scroll', onScroll, { passive: true });
  onScroll(); // Initial call

  // --- Scroll to top ---
  scrollTopBtn.addEventListener('click', () => {
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  // --- Hamburger Menu（1518版ヘッダー） ---
  const menuToggle = document.getElementById('menu-toggle');

  menuToggle.addEventListener('click', () => {
    const isOpen = header.classList.toggle('nav-open');
    menuToggle.setAttribute('aria-label', isOpen ? 'メニューを閉じる' : 'メニューを開く');
    document.body.style.overflow = isOpen ? 'hidden' : '';
  });

  // Close menu on link click
  document.querySelectorAll('.nav a').forEach((link) => {
    link.addEventListener('click', () => {
      header.classList.remove('nav-open');
      menuToggle.setAttribute('aria-label', 'メニューを開く');
      document.body.style.overflow = '';
    });
  });

  // --- FAQ Accordion ---
  const faqQuestions = document.querySelectorAll('.faq-question');

  faqQuestions.forEach((question) => {
    question.addEventListener('click', () => {
      const expanded = question.getAttribute('aria-expanded') === 'true';
      const answer = question.nextElementSibling;

      // Close all other answers
      faqQuestions.forEach((q) => {
        if (q !== question) {
          q.setAttribute('aria-expanded', 'false');
          q.nextElementSibling.classList.remove('open');
          q.nextElementSibling.setAttribute('aria-hidden', 'true');
        }
      });

      // Toggle current
      question.setAttribute('aria-expanded', !expanded);
      answer.classList.toggle('open');
      answer.setAttribute('aria-hidden', expanded);
    });

    // Keyboard accessibility
    question.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        question.click();
      }
    });
  });

  // --- Smooth scroll for anchor links ---
  document.querySelectorAll('a[href^="#"]').forEach((anchor) => {
    anchor.addEventListener('click', (e) => {
      const targetId = anchor.getAttribute('href');
      if (targetId === '#') return;

      const target = document.querySelector(targetId);
      if (!target) return;

      e.preventDefault();

      const headerHeight = header.offsetHeight;
      const targetPosition = target.getBoundingClientRect().top + window.scrollY - headerHeight - 16;

      window.scrollTo({
        top: targetPosition,
        behavior: 'smooth',
      });
    });
  });
});
