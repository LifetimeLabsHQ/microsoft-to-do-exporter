export type ThemeMode = 'auto' | 'light' | 'dark';

const STORAGE_KEY = 'pn_exporter_theme';

function readStoredMode(): ThemeMode {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw === 'light' || raw === 'dark' || raw === 'auto') return raw;
  } catch {
    // Ignore storage access errors.
  }
  return 'auto';
}

function resolveEffective(mode: ThemeMode): 'light' | 'dark' {
  if (mode === 'light' || mode === 'dark') return mode;
  return window.matchMedia?.('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

export function applyTheme(mode: ThemeMode): void {
  const root = document.documentElement;
  if (mode === 'auto') {
    root.removeAttribute('data-theme');
  } else {
    root.setAttribute('data-theme', mode);
  }
  root.setAttribute('data-theme-mode', mode);
  root.style.colorScheme = resolveEffective(mode);
  try {
    localStorage.setItem(STORAGE_KEY, mode);
  } catch {
    // Ignore storage access errors.
  }
}

export function initThemeSwitcher(): void {
  let current = readStoredMode();
  applyTheme(current);

  const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>('[data-theme-choice]'));
  const syncButtons = () => {
    for (const btn of buttons) {
      const choice = btn.dataset.themeChoice as ThemeMode | undefined;
      const active = choice === current;
      btn.setAttribute('aria-pressed', String(active));
      btn.classList.toggle('is-active', active);
    }
  };
  syncButtons();

  for (const btn of buttons) {
    btn.addEventListener('click', () => {
      const choice = btn.dataset.themeChoice as ThemeMode | undefined;
      if (!choice) return;
      current = choice;
      applyTheme(current);
      syncButtons();
    });
  }

  const mq = window.matchMedia?.('(prefers-color-scheme: dark)');
  mq?.addEventListener?.('change', () => {
    if (current === 'auto') applyTheme('auto');
  });
}
