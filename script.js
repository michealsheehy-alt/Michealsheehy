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

  let home = links.querySelector('a[data-home-link]');
  if (!home) {
    home = document.createElement('a');
    home.href = '/';
    home.textContent = 'Home';
    home.dataset.homeLink = 'true';
    links.prepend(home);
  }
  if (location.pathname === '/' || location.pathname.endsWith('/index.html')) {
    home.setAttribute('aria-current', 'page');
  }

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

  document.querySelectorAll('[data-filter-target]').forEach((bar) => {
    const target = document.querySelector(bar.dataset.filterTarget);
    if (!target) return;
    const cards = [...target.querySelectorAll('[data-category]')];
    bar.querySelectorAll('button[data-filter]').forEach((button) => {
      button.addEventListener('click', () => {
        const selected = button.dataset.filter;
        bar.querySelectorAll('button[data-filter]').forEach((item) => item.setAttribute('aria-pressed', String(item === button)));
        cards.forEach((card) => {
          const categories = (card.dataset.category || '').split(/\s+/);
          card.hidden = selected !== 'all' && !categories.includes(selected);
        });
      });
    });
  });

  const canonical = document.querySelector('link[rel="canonical"]')?.href;
  const article = document.querySelector('main article');
  const path = location.pathname.replace(/\/+$/, '');
  if (canonical && article && path.includes('/articles/') && !document.querySelector('.article-utility')) {
    const sequences = [
      [
        ['/articles/chapter-1.html', 'The Industry Solved the Wrong Problem'],
        ['/articles/chapter-2.html', 'The Evidence Has Been There All Along'],
        ['/articles/chapter-3.html', 'Every Model Begins Drifting on Day One'],
        ['/articles/chapter-4.html', 'The Adaptive Compliance Maturity Model'],
        ['/articles/chapter-5.html', 'AI Won’t Replace Compliance'],
        ['/articles/chapter-6.html', 'Designing an Adaptive Compliance Organization'],
        ['/articles/chapter-7.html', 'Measuring What Matters'],
        ['/articles/chapter-8.html', 'Adaptive Compliance as a Strategic Capability']
      ],
      [
        ['/articles/compliance-advantage.html', 'The Compliance Advantage'],
        ['/articles/trust-at-the-speed-of-money.html', 'Trust at the Speed of Money'],
        ['/articles/the-message-is-the-control.html', 'The Message Is the Control'],
        ['/articles/the-criminal-is-already-in-production.html', 'The Criminal Is Already in Production'],
        ['/articles/global-compliance-without-the-global-bottleneck.html', 'Global Compliance Without the Global Bottleneck'],
        ['/articles/your-new-platform-is-not-a-transformation.html', 'Your New Platform Is Not a Transformation'],
        ['/articles/the-safest-customer-is-not-the-customer-you-refused-to-understand.html', 'The Safest Customer Is Not the Customer You Refused to Understand']
      ],
      [
        ['/articles/the-loneliness-of-senior-accountability.html', 'The Loneliness of Senior Accountability'],
        ['/articles/the-vendor-decision-i-got-wrong.html', 'The Vendor Decision I Got Wrong'],
        ['/articles/your-leaders-should-eventually-stop-needing-you.html', 'Your Leaders Should Eventually Stop Needing You'],
        ['/articles/escalation-is-evidence-of-a-healthy-culture.html', 'Why Escalation Is Evidence of a Healthy Culture'],
        ['/articles/a-cco-must-sometimes-be-unpopular.html', 'A CCO Must Sometimes Be Comfortable Being Unpopular']
      ]
    ];
    const currentSequence = sequences.find((items) => items.some(([url]) => path.endsWith(url)));
    const index = currentSequence?.findIndex(([url]) => path.endsWith(url)) ?? -1;
    const previous = index > 0 ? currentSequence[index - 1] : null;
    const next = currentSequence && index < currentSequence.length - 1 ? currentSequence[index + 1] : null;
    const related = path.includes('leadership') || path.includes('loneliness') || path.includes('vendor-decision') || path.includes('leaders-should') || path.includes('escalation') || path.includes('unpopular')
      ? [
          ['/articles/personal-side-of-being-a-cco.html', 'Finding My Voice'],
          ['/articles/c-suite-leadership.html', 'What Nobody Tells You About Joining the C-Suite']
        ]
      : path.endsWith('/articles/one-global-standard-does-not-mean-one-global-answer.html')
        ? [
            ['/articles/global-compliance-without-the-global-bottleneck.html', 'Global Compliance Without the Global Bottleneck'],
            ['/articles/global-standards-local-proof.html', 'Global Standards, Local Proof']
          ]
      : [
          ['/tools.html', 'Practical tools and frameworks'],
          ['/articles/one-global-standard-does-not-mean-one-global-answer.html', 'One Global Standard Does Not Mean One Global Answer']
        ];
    const utility = document.createElement('section');
    utility.className = 'article-utility';
    utility.setAttribute('aria-label', 'Article actions and related reading');
    const navItems = [
      previous ? `<a href="..${previous[0]}"><span>Previous</span><strong>${previous[1]}</strong></a>` : '',
      next ? `<a href="..${next[0]}"><span>Next</span><strong>${next[1]}</strong></a>` : ''
    ].join('');
    utility.innerHTML = `
      <div class="article-utility__share">
        <strong>Share this perspective</strong>
        <a href="https://www.linkedin.com/sharing/share-offsite/?url=${encodeURIComponent(canonical)}" target="_blank" rel="noopener noreferrer">Share on LinkedIn ↗</a>
        <a href="mailto:?subject=${encodeURIComponent(document.title)}&body=${encodeURIComponent(canonical)}">Share by email</a>
      </div>
      ${navItems ? `<nav class="article-utility__nav" aria-label="Previous and next article">${navItems}</nav>` : ''}
      <div class="article-utility__related"><h2>Continue reading</h2><ul>${related.map(([url,title]) => `<li><a href="..${url}">${title}</a></li>`).join('')}</ul></div>
    `;
    const host = article.closest('.container') || article.parentElement;
    host.appendChild(utility);
  }
})();


// Google Analytics 4 and conversion measurement
(() => {
  const measurementId = 'G-P9LKM1MH8H';
  window.dataLayer = window.dataLayer || [];
  window.gtag = window.gtag || function gtag(){ window.dataLayer.push(arguments); };
  window.gtag('js', new Date());
  window.gtag('config', measurementId, {
    anonymize_ip: true,
    transport_type: 'beacon'
  });

  const analyticsScript = document.createElement('script');
  analyticsScript.async = true;
  analyticsScript.src = `https://www.googletagmanager.com/gtag/js?id=${measurementId}`;
  document.head.appendChild(analyticsScript);

  document.addEventListener('click', (event) => {
    const link = event.target.closest('a[href]');
    if (!link) return;
    const href = link.getAttribute('href') || '';
    const absoluteUrl = link.href || href;
    const label = (link.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 120);

    if (/^(mailto:|tel:)/i.test(href)) {
      window.gtag('event', 'contact_click', {
        contact_method: href.startsWith('mailto:') ? 'email' : 'phone',
        link_text: label
      });
      return;
    }

    if (link.hasAttribute('download') || /\.(pdf|docx|pptx|xlsx|zip)(\?|#|$)/i.test(href)) {
      window.gtag('event', 'file_download', {
        file_name: decodeURIComponent(href.split('/').pop()?.split(/[?#]/)[0] || ''),
        link_text: label,
        link_url: absoluteUrl
      });
    }

    if (link.classList.contains('button')) {
      window.gtag('event', 'cta_click', {
        link_text: label,
        link_url: absoluteUrl,
        page_path: window.location.pathname
      });
    }

    try {
      const target = new URL(absoluteUrl, window.location.href);
      if (/^https?:$/.test(target.protocol) && target.hostname !== window.location.hostname) {
        window.gtag('event', 'outbound_click', {
          link_domain: target.hostname,
          link_text: label,
          link_url: target.href
        });
      }
    } catch (_) {}
  });

  const videoFrames = [...document.querySelectorAll('iframe[src*="youtube.com/embed/"], iframe[src*="youtube-nocookie.com/embed/"]')];
  if (!videoFrames.length) return;

  const existingReady = window.onYouTubeIframeAPIReady;
  window.onYouTubeIframeAPIReady = () => {
    if (typeof existingReady === 'function') existingReady();
    videoFrames.forEach((frame) => {
      const videoId = (frame.src.match(/\/embed\/([^?&]+)/) || [])[1] || 'unknown';
      let started = false;
      try {
        new window.YT.Player(frame, {
          events: {
            onStateChange: ({ data }) => {
              if (data === window.YT.PlayerState.PLAYING && !started) {
                started = true;
                window.gtag('event', 'video_start', {
                  video_id: videoId,
                  video_title: frame.title || 'Embedded video'
                });
              }
              if (data === window.YT.PlayerState.ENDED) {
                window.gtag('event', 'video_complete', {
                  video_id: videoId,
                  video_title: frame.title || 'Embedded video'
                });
              }
            }
          }
        });
      } catch (_) {}
    });
  };

  const youtubeApi = document.createElement('script');
  youtubeApi.src = 'https://www.youtube.com/iframe_api';
  youtubeApi.async = true;
  document.head.appendChild(youtubeApi);
})();
