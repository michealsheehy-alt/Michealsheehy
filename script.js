(() => {
  const main = document.querySelector('main');
  if (main && !main.id) main.id = 'main-content';

  if (main && !document.querySelector('.skip-link')) {
    const skip = document.createElement('a');
    skip.className = 'skip-link';
    skip.href = '#main-content';
    skip.textContent = 'Skip to main content';
    document.body.prepend(skip);
  }

  if (!document.querySelector('#accessibility-enhancements')) {
    const style = document.createElement('style');
    style.id = 'accessibility-enhancements';
    style.textContent = `
      .skip-link{position:fixed;left:14px;top:14px;z-index:9999;padding:10px 14px;border-radius:999px;background:#192720;color:#fff;text-decoration:none;transform:translateY(-180%);transition:transform .15s ease}
      .skip-link:focus{transform:translateY(0)}
      :focus-visible{outline:3px solid #7b8f27;outline-offset:3px}
    `;
    document.head.appendChild(style);
  }

  const toggle = document.querySelector('.mobile-toggle');
  const links = document.querySelector('.nav-links');
  if (!toggle || !links) return;

  if (!links.id) links.id = 'primary-navigation';
  toggle.setAttribute('aria-controls', links.id);
  toggle.setAttribute('aria-expanded', 'false');

  const setOpen = (open) => {
    links.classList.toggle('open', open);
    toggle.setAttribute('aria-expanded', String(open));
    toggle.setAttribute('aria-label', open ? 'Close navigation' : 'Open navigation');
  };

  toggle.addEventListener('click', () => setOpen(!links.classList.contains('open')));
  links.addEventListener('click', (event) => {
    if (event.target.closest('a')) setOpen(false);
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && links.classList.contains('open')) {
      setOpen(false);
      toggle.focus();
    }
  });
  document.addEventListener('click', (event) => {
    if (links.classList.contains('open') && !links.contains(event.target) && !toggle.contains(event.target)) setOpen(false);
  });
  window.addEventListener('resize', () => {
    if (window.innerWidth > 980) setOpen(false);
  });
})();
