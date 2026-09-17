// ── CHATORI CHAAT – MAIN JS v2 ───────────────────────────────────────────────

// Scroll reveal
const observer = new IntersectionObserver(entries => {
  entries.forEach(e => { if (e.isIntersecting) e.target.classList.add('visible'); });
}, { threshold: 0.08 });
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('.reveal').forEach(el => observer.observe(el));

  // Active nav
  const path = window.location.pathname;
  document.querySelectorAll('.nav-link').forEach(link => {
    const href = link.getAttribute('href');
    if (href === path || (href !== '/' && path.startsWith(href))) {
      link.style.color = 'var(--turmeric)';
    }
  });

  // Navbar shrink
  window.addEventListener('scroll', () => {
    const nb = document.querySelector('.cc-navbar');
    if (nb) nb.style.padding = window.scrollY > 50 ? '5px 0' : '10px 0';
  });

  // Pull notifications if on admin
  if (document.getElementById('notifBell')) loadNotifications();
});

// ── LIGHTBOX ─────────────────────────────────────────────────────────────────
let lightboxImages = [];
let lightboxIndex = 0;

function openLightbox(src, caption, groupImages) {
  lightboxImages = groupImages || [{ src, caption }];
  lightboxIndex = lightboxImages.findIndex(i => i.src === src);
  if (lightboxIndex === -1) lightboxIndex = 0;
  showLightboxImage();
  document.getElementById('lightbox').classList.add('active');
  document.body.style.overflow = 'hidden';
}

function showLightboxImage() {
  const img = lightboxImages[lightboxIndex];
  document.getElementById('lightbox-img').src = img.src;
  document.getElementById('lightbox-caption').textContent = img.caption || '';
}

function lightboxNav(dir) {
  lightboxIndex = (lightboxIndex + dir + lightboxImages.length) % lightboxImages.length;
  showLightboxImage();
}

function closeLightbox() {
  document.getElementById('lightbox').classList.remove('active');
  document.body.style.overflow = '';
}

document.addEventListener('keydown', e => {
  if (!document.getElementById('lightbox').classList.contains('active')) return;
  if (e.key === 'Escape') closeLightbox();
  if (e.key === 'ArrowLeft') lightboxNav(-1);
  if (e.key === 'ArrowRight') lightboxNav(1);
});

// ── NOTIFICATIONS ─────────────────────────────────────────────────────────────
function loadNotifications() {
  fetch('/api/notifications').then(r => r.json()).then(data => {
    const badge = document.getElementById('notifCount');
    const list = document.getElementById('notifList');
    if (badge) badge.textContent = data.unread || '';
    if (list && data.notifications) {
      list.innerHTML = data.notifications.map(n => `
        <div class="notif-item ${n.type}">
          <div class="fw-600 small">${n.title}</div>
          <div class="text-muted" style="font-size:0.78rem;">${n.message}</div>
          <div class="text-muted" style="font-size:0.7rem;">${n.created_at?.slice(0,16) || ''}</div>
        </div>`).join('') || '<div class="p-3 text-muted small text-center">No notifications</div>';
    }
  }).catch(() => {});
  setTimeout(loadNotifications, 30000); // refresh every 30s
}

function toggleNotifPopup() {
  const popup = document.getElementById('notifPopup');
  if (popup) popup.classList.toggle('show');
}

// Close popup on outside click
document.addEventListener('click', e => {
  const popup = document.getElementById('notifPopup');
  const bell = document.getElementById('notifBell');
  if (popup && !popup.contains(e.target) && bell && !bell.contains(e.target)) {
    popup.classList.remove('show');
  }
});
