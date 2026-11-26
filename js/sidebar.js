window.Sidebar = (() => {
  const STORAGE_KEY = 'aoi-sidebar-state';
  const mobileQuery = window.matchMedia('(max-width: 1023px)');
  let initialized = false;
  let appShell;
  let sidebar;
  let toggleButton;
  let mobileButton;
  let mainShell;
  let backdrop;
  let tooltip;

  function isMobile() {
    return mobileQuery.matches;
  }

  function savedState() {
    try {
      return localStorage.getItem(STORAGE_KEY) === 'collapsed' ? 'collapsed' : 'open';
    } catch (error) {
      return 'open';
    }
  }

  function saveState(state) {
    try {
      localStorage.setItem(STORAGE_KEY, state);
    } catch (error) {
      // The sidebar remains usable when storage is unavailable.
    }
  }

  function updateDesktopState(state, persist = true) {
    const changed = appShell.dataset.sidebar !== state;
    appShell.dataset.sidebar = state;
    document.documentElement.dataset.sidebar = state;
    toggleButton.setAttribute('aria-expanded', String(state === 'open'));
    toggleButton.setAttribute('aria-label', state === 'open' ? 'Collapse sidebar' : 'Expand sidebar');
    if (persist) saveState(state);
    hideTooltip();
    if (changed && window.matchMedia('(prefers-reduced-motion: reduce)').matches) dispatchResize();
  }

  function dispatchResize() {
    window.dispatchEvent(new Event('resize'));
  }

  function closeDrawer(returnFocus = false) {
    appShell.dataset.sidebarDrawer = 'closed';
    sidebar.inert = true;
    sidebar.setAttribute('aria-hidden', 'true');
    mainShell.inert = false;
    mobileButton.setAttribute('aria-expanded', 'false');
    mobileButton.setAttribute('aria-label', 'Open navigation menu');
    if (returnFocus) mobileButton.focus();
  }

  function openDrawer() {
    appShell.dataset.sidebarDrawer = 'open';
    sidebar.inert = false;
    sidebar.setAttribute('aria-hidden', 'false');
    mainShell.inert = true;
    mobileButton.setAttribute('aria-expanded', 'true');
    mobileButton.setAttribute('aria-label', 'Close navigation menu');
    const firstItem = sidebar.querySelector('.nav-item');
    if (firstItem) firstItem.focus();
  }

  function open() {
    if (isMobile()) {
      openDrawer();
      return;
    }
    updateDesktopState('open');
  }

  function close() {
    if (isMobile()) {
      closeDrawer();
      return;
    }
    updateDesktopState('collapsed');
  }

  function toggle() {
    if (isMobile()) {
      if (appShell.dataset.sidebarDrawer === 'open') closeDrawer();
      else openDrawer();
      return;
    }
    updateDesktopState(appShell.dataset.sidebar === 'open' ? 'collapsed' : 'open');
  }

  function hideTooltip() {
    if (!tooltip) return;
    tooltip.classList.remove('visible');
    tooltip.textContent = '';
  }

  function showTooltip(target) {
    if (isMobile() || appShell.dataset.sidebar !== 'collapsed') {
      hideTooltip();
      return;
    }
    const label = target.dataset.tooltip;
    if (!label) return;
    const bounds = target.getBoundingClientRect();
    tooltip.textContent = label;
    tooltip.style.top = `${Math.round(bounds.top + bounds.height / 2)}px`;
    const tooltipWidth = tooltip.getBoundingClientRect().width;
    const left = Math.max(8, Math.min(Math.round(bounds.right + 8), window.innerWidth - tooltipWidth - 8));
    tooltip.style.left = `${left}px`;
    tooltip.classList.add('visible');
  }

  function updateResponsiveState() {
    hideTooltip();
    if (isMobile()) {
      const returnFocus = sidebar.contains(document.activeElement);
      closeDrawer();
      if (returnFocus) mobileButton.focus();
      return;
    }
    appShell.dataset.sidebarDrawer = 'closed';
    sidebar.inert = false;
    mainShell.inert = false;
    sidebar.setAttribute('aria-hidden', 'false');
    mobileButton.setAttribute('aria-expanded', 'false');
    mobileButton.setAttribute('aria-label', 'Open navigation menu');
  }

  function init() {
    if (initialized) return;
    appShell = document.getElementById('appShell');
    sidebar = document.getElementById('sidebar');
    toggleButton = document.getElementById('sidebarToggle');
    mobileButton = document.getElementById('mobileMenu');
    mainShell = document.querySelector('.main-shell');
    backdrop = document.getElementById('sidebarBackdrop');
    if (!appShell || !sidebar || !toggleButton || !mobileButton || !mainShell || !backdrop) return;

    initialized = true;
    tooltip = document.createElement('div');
    tooltip.className = 'sidebar-tooltip';
    tooltip.setAttribute('role', 'tooltip');
    document.body.appendChild(tooltip);

    updateDesktopState(savedState(), false);
    updateResponsiveState();

    toggleButton.addEventListener('click', toggle);
    mobileButton.addEventListener('click', toggle);
    backdrop.addEventListener('click', () => closeDrawer(true));
    sidebar.querySelector('#navList').addEventListener('click', event => {
      if (event.target.closest('.nav-item') && isMobile()) closeDrawer(true);
    });

    sidebar.addEventListener('pointerover', event => {
      const target = event.target.closest('[data-tooltip]');
      if (target && sidebar.contains(target)) showTooltip(target);
    });
    sidebar.addEventListener('pointerout', event => {
      const target = event.target.closest('[data-tooltip]');
      if (target && (!event.relatedTarget || !target.contains(event.relatedTarget))) hideTooltip();
    });
    sidebar.addEventListener('focusin', event => {
      const target = event.target.closest('[data-tooltip]');
      if (target && sidebar.contains(target)) showTooltip(target);
    });
    sidebar.addEventListener('focusout', event => {
      if (!event.relatedTarget || !sidebar.contains(event.relatedTarget)) hideTooltip();
    });

    sidebar.addEventListener('transitionend', event => {
      if (event.target !== sidebar || event.propertyName !== 'width') return;
      dispatchResize();
    });

    document.addEventListener('keydown', event => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'b') {
        event.preventDefault();
        toggle();
        return;
      }
      if (event.key === 'Escape' && isMobile() && appShell.dataset.sidebarDrawer === 'open') {
        event.preventDefault();
        closeDrawer(true);
        return;
      }
      if (event.key === 'Tab' && isMobile() && appShell.dataset.sidebarDrawer === 'open') {
        const focusable = [...sidebar.querySelectorAll('button:not(:disabled)')].filter(element => element.getClientRects().length);
        const first = focusable[0];
        const last = focusable[focusable.length - 1];
        if (event.shiftKey && document.activeElement === first) {
          event.preventDefault();
          last.focus();
        } else if (!event.shiftKey && document.activeElement === last) {
          event.preventDefault();
          first.focus();
        }
      }
    });

    mobileQuery.addEventListener('change', updateResponsiveState);
    window.addEventListener('resize', hideTooltip);
    window.addEventListener('scroll', hideTooltip, true);

    requestAnimationFrame(() => {
      delete document.documentElement.dataset.sidebarInitializing;
    });
  }

  init();
  return { init, open, close, toggle };
})();
