// =========================================================================
// APP.JS - INTERACTIVE TELECOM EBOOK ENGINE
// Author: Engr. Mohammad Abdul Mannan | BTRC Spectrum Division
// =========================================================================

document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initReadingProgress();
  initSidebar();
  initSearch();
  initScrollSpy();
  initBackToTop();
});

// 1. Dark / Light Theme Toggle with LocalStorage
function initTheme() {
  const themeToggle = document.getElementById('themeToggle');
  const savedTheme = localStorage.getItem('telecom_ebook_theme') || 'dark';
  document.documentElement.setAttribute('data-theme', savedTheme);

  if (themeToggle) {
    themeToggle.addEventListener('click', () => {
      const currentTheme = document.documentElement.getAttribute('data-theme');
      const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', newTheme);
      localStorage.setItem('telecom_ebook_theme', newTheme);
    });
  }
}

// 2. Reading Progress Bar
function initReadingProgress() {
  const bar = document.getElementById('progressBar');
  window.addEventListener('scroll', () => {
    const totalHeight = document.documentElement.scrollHeight - window.innerHeight;
    if (totalHeight > 0) {
      const progress = (window.scrollY / totalHeight) * 100;
      bar.style.width = `${progress}%`;
    }
  });
}

// 3. Mobile Sidebar Toggle & Backdrop
function initSidebar() {
  const toggleBtn = document.getElementById('menuToggle');
  const sidebar = document.getElementById('sidebar');
  const backdrop = document.getElementById('sidebarBackdrop');

  if (toggleBtn && sidebar && backdrop) {
    toggleBtn.addEventListener('click', () => {
      sidebar.classList.toggle('open');
      backdrop.classList.toggle('active');
    });

    backdrop.addEventListener('click', () => {
      sidebar.classList.remove('open');
      backdrop.classList.remove('active');
    });

    // Close on link click on mobile
    const links = sidebar.querySelectorAll('.toc-link');
    links.forEach(l => {
      l.addEventListener('click', () => {
        if (window.innerWidth <= 768) {
          sidebar.classList.remove('open');
          backdrop.classList.remove('active');
        }
      });
    });
  }
}

// 4. Real-time Full Text Search
function initSearch() {
  const searchInput = document.getElementById('globalSearch');
  const mobileSearch = document.getElementById('mobileSearch');
  const countDisplay = document.getElementById('searchCount');

  function handleSearch(query) {
    const q = query.trim().toLowerCase();
    const tocLinks = document.querySelectorAll('.sidebar .toc-link');
    const sections = document.querySelectorAll('.chapter-block, .section-container, .formula-card, .math-solution-card, .bullet-card');

    if (!q) {
      tocLinks.forEach(l => l.parentElement.style.display = '');
      sections.forEach(s => s.style.display = '');
      if (countDisplay) countDisplay.textContent = '';
      return;
    }

    let matchCount = 0;

    // Filter content blocks
    sections.forEach(sec => {
      const text = sec.textContent.toLowerCase();
      if (text.includes(q)) {
        sec.style.display = '';
        matchCount++;
      } else {
        sec.style.display = 'none';
      }
    });

    // Filter sidebar links
    tocLinks.forEach(l => {
      const text = l.textContent.toLowerCase();
      if (text.includes(q)) {
        l.parentElement.style.display = '';
      } else {
        l.parentElement.style.display = 'none';
      }
    });

    if (countDisplay) {
      countDisplay.textContent = `${matchCount} ফলাফল`;
    }
  }

  if (searchInput) {
    searchInput.addEventListener('input', (e) => handleSearch(e.target.value));
    // Ctrl + K shortcut
    window.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        searchInput.focus();
        searchInput.select();
      }
    });
  }

  if (mobileSearch) {
    mobileSearch.addEventListener('input', (e) => handleSearch(e.target.value));
  }
}

// 5. ScrollSpy for Sidebar Active State
function initScrollSpy() {
  const headings = document.querySelectorAll('.chapter-block, .section-container, .author-bio-card, .book-hero-card, .calculators-hub-section');
  const navLinks = document.querySelectorAll('.sidebar .toc-link');

  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        const id = entry.target.getAttribute('id');
        if (id) {
          navLinks.forEach(link => {
            if (link.getAttribute('data-target') === id) {
              link.classList.add('active');
            } else {
              link.classList.remove('active');
            }
          });
        }
      }
    });
  }, { rootMargin: '-20% 0px -70% 0px' });

  headings.forEach(h => observer.observe(h));
}

// 6. Back to Top Button
function initBackToTop() {
  const btn = document.getElementById('backToTop');
  if (btn) {
    window.addEventListener('scroll', () => {
      if (window.scrollY > 400) {
        btn.classList.add('visible');
      } else {
        btn.classList.remove('visible');
      }
    });

    btn.addEventListener('click', () => {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  }
}

// 7. Font Size Adjuster
let currentFontSize = 16;
function toggleFontSize() {
  currentFontSize = currentFontSize >= 19 ? 15 : currentFontSize + 1;
  document.documentElement.style.fontSize = `${currentFontSize}px`;
}

// 8. Interactive Calculators Functions
function switchCalc(calcName) {
  const cards = document.querySelectorAll('.calc-card');
  const tabs = document.querySelectorAll('.calc-tab-btn');

  cards.forEach(c => c.classList.remove('active'));
  tabs.forEach(t => t.classList.remove('active'));

  const activeCard = document.getElementById(`calc_${calcName}`);
  if (activeCard) activeCard.classList.add('active');

  const activeTab = Array.from(tabs).find(t => t.getAttribute('onclick').includes(calcName));
  if (activeTab) activeTab.classList.add('active');
}

// Shannon Calculator
function calculateShannon() {
  const bw = parseFloat(document.getElementById('shannon_bw').value) || 0;
  const snrDb = parseFloat(document.getElementById('shannon_snr').value) || 0;
  const snrLinear = Math.pow(10, snrDb / 10);
  const capacity = bw * (Math.log(1 + snrLinear) / Math.LN2);
  const resElem = document.getElementById('shannon_result').querySelector('.res-value');
  resElem.textContent = `${capacity.toFixed(2)} Mbps`;
}

// FSPL Calculator
function calculateFSPL() {
  const f = parseFloat(document.getElementById('fspl_f').value) || 0;
  const d = parseFloat(document.getElementById('fspl_d').value) || 0;
  if (f <= 0 || d <= 0) return;
  const fspl = 92.45 + (20 * Math.log10(f)) + (20 * Math.log10(d));
  const resElem = document.getElementById('fspl_result').querySelector('.res-value');
  resElem.textContent = `${fspl.toFixed(2)} dB`;
}

// Fresnel Zone Calculator
function calculateFresnel() {
  const d1 = parseFloat(document.getElementById('fresnel_d1').value) || 0;
  const d2 = parseFloat(document.getElementById('fresnel_d2').value) || 0;
  const f = parseFloat(document.getElementById('fresnel_f').value) || 0;
  const totalD = d1 + d2;
  if (f <= 0 || totalD <= 0) return;
  const r1 = 17.32 * Math.sqrt((d1 * d2) / (f * totalD));
  const r60 = 0.60 * r1;
  const resElem = document.getElementById('fresnel_result').querySelector('.res-value');
  resElem.textContent = `${r1.toFixed(2)} মিটার (৬০% ক্লিয়ারেন্স: ${r60.toFixed(2)} মি.)`;
}

// Fiber Reach Calculator
function calculateFiber() {
  const ptx = parseFloat(document.getElementById('fiber_ptx').value) || 0;
  const psens = parseFloat(document.getElementById('fiber_psens').value) || 0;
  const alpha = parseFloat(document.getElementById('fiber_alpha').value) || 0.225;
  const margin = parseFloat(document.getElementById('fiber_margin').value) || 4.0;
  const allowed = (ptx - psens) - margin;
  if (alpha <= 0) return;
  const lmax = allowed / alpha;
  const resElem = document.getElementById('fiber_result').querySelector('.res-value');
  resElem.textContent = `${lmax.toFixed(2)} কিমি`;
}
