/* Happy Hoppy – main.js */

document.addEventListener('DOMContentLoaded', () => {

  // ── Global search: sync type radio to hidden field in nav form ──────────
  const globalTypeInput = document.getElementById('global-search-type');
  document.querySelectorAll('input[name="type"]').forEach(radio => {
    radio.addEventListener('change', () => {
      if (globalTypeInput) globalTypeInput.value = radio.value;
    });
  });

  // ── Auto-resize email iframe once loaded ─────────────────────────────────
  const iframe = document.getElementById('email-html-frame');
  if (iframe) {
    const resize = () => {
      try {
        const h = iframe.contentDocument.documentElement.scrollHeight;
        iframe.style.height = Math.min(Math.max(h + 20, 200), 1400) + 'px';
      } catch (e) { /* cross-origin guard */ }
    };
    iframe.addEventListener('load', resize);
    // Retry once after a short delay (some HTML bodies render slowly)
    setTimeout(resize, 600);
  }

  // ── Keyboard shortcut: '/' focuses search bar ────────────────────────────
  document.addEventListener('keydown', e => {
    if (e.key === '/' && document.activeElement.tagName !== 'INPUT'
                      && document.activeElement.tagName !== 'TEXTAREA') {
      e.preventDefault();
      const input = document.getElementById('global-search-input');
      if (input) { input.focus(); input.select(); }
    }
  });

  // ── Highlight search term in page title on results page ──────────────────
  const q = new URLSearchParams(window.location.search).get('q');
  if (q) document.title = `"${q}" – Happy Hoppy`;

});
