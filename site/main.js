/* ============================================================
   HANOVA CONSULTANCY — MAIN JAVASCRIPT
   Handles: nav scroll, mobile menu, scroll animations
   ============================================================ */

// --- NAV: add .scrolled class on scroll ---
const nav = document.getElementById('nav');
if (nav) {
  // Always scrolled on inner pages (not hero)
  if (!document.querySelector('.hero')) {
    nav.classList.add('scrolled');
  }
  window.addEventListener('scroll', () => {
    nav.classList.toggle('scrolled', window.scrollY > 40);
  }, { passive: true });
}

// --- MOBILE MENU ---
function toggleMenu() {
  const links = document.getElementById('navLinks');
  const btn = document.querySelector('.nav-toggle');
  if (links) {
    links.classList.toggle('open');
    btn.classList.toggle('open');
    document.body.style.overflow = links.classList.contains('open') ? 'hidden' : '';
  }
}

// Close menu on link click
document.querySelectorAll('.nav-links a').forEach(link => {
  link.addEventListener('click', () => {
    const links = document.getElementById('navLinks');
    if (links) {
      links.classList.remove('open');
      document.body.style.overflow = '';
    }
  });
});

// --- INTERSECTION OBSERVER: fade-up animations ---
const fadeEls = document.querySelectorAll('.fade-up');

if ('IntersectionObserver' in window && fadeEls.length) {
  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.12, rootMargin: '0px 0px -40px 0px' });

  fadeEls.forEach(el => observer.observe(el));
} else {
  // Fallback: just show everything
  fadeEls.forEach(el => el.classList.add('visible'));
}

// --- CONTACT FORM: basic validation ---
const form = document.getElementById('contactForm');
if (form) {
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const name = form.querySelector('#name').value.trim();
    const email = form.querySelector('#email').value.trim();
    const message = form.querySelector('#message').value.trim();

    if (!name || !email || !message) {
      alert('Please fill in all required fields.');
      return;
    }

    // Simulate submission
    const btn = form.querySelector('button[type="submit"]');
    btn.textContent = 'Sending...';
    btn.disabled = true;

    setTimeout(() => {
      btn.textContent = 'Message Sent ✓';
      form.reset();
    }, 1200);
  });
}

// --- ACTIVE NAV LINK ---
const currentPath = window.location.pathname.split('/').pop() || 'index.html';
document.querySelectorAll('.nav-links a').forEach(link => {
  link.classList.remove('active');
  const href = link.getAttribute('href');
  if (href === currentPath || (currentPath === '' && href === 'index.html')) {
    link.classList.add('active');
  }
});
