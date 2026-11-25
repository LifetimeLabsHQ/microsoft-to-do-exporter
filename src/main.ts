import type { AccountInfo } from '@azure/msal-browser';
import { acquireGraphToken, initAuth, signInWithMicrosoft, signOutLocal } from './auth';
import { buildDemoBundle } from './demoData';
import { exportMicrosoftToDo } from './graph';
import { buildExportZip, exportDateStamp, renderExportCsv, summarizeBundle } from './packager';
import { initThemeSwitcher } from './theme';
import type { ExportBundle, FetchProgress } from './types';

interface AppState {
  phase: 'init' | 'signed-out' | 'signed-in' | 'exporting' | 'ready';
  account: AccountInfo | null;
  demoMode: boolean;
  includeCompleted: boolean;
  includeAttachments: boolean;
  progress: FetchProgress | null;
  bundle: ExportBundle | null;
  error: string | null;
  busyDownload: boolean;
}

const state: AppState = {
  phase: 'init',
  account: null,
  demoMode: false,
  includeCompleted: true,
  includeAttachments: true,
  progress: null,
  bundle: null,
  error: null,
  busyDownload: false,
};

let abortController: AbortController | null = null;
let progressDom: {
  msgEl: HTMLElement;
  fillEl: HTMLElement;
  countsEl: HTMLElement;
} | null = null;

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  attrs: Record<string, string> = {},
  children: Array<Node | string> = [],
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    node.setAttribute(k, v);
  }
  for (const child of children) {
    node.append(typeof child === 'string' ? document.createTextNode(child) : child);
  }
  return node;
}

function createMicrosoftLogo(): SVGSVGElement {
  const ns = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(ns, 'svg');
  svg.setAttribute('width', '15');
  svg.setAttribute('height', '15');
  svg.setAttribute('viewBox', '0 0 21 21');
  svg.setAttribute('aria-hidden', 'true');
  const squares: Array<[string, string, string]> = [
    ['1', '1', '#f25022'],
    ['11', '1', '#7fba00'],
    ['1', '11', '#00a4ef'],
    ['11', '11', '#ffb900'],
  ];
  for (const [x, y, fill] of squares) {
    const rect = document.createElementNS(ns, 'rect');
    rect.setAttribute('x', x);
    rect.setAttribute('y', y);
    rect.setAttribute('width', '9');
    rect.setAttribute('height', '9');
    rect.setAttribute('fill', fill);
    svg.appendChild(rect);
  }
  return svg;
}

function createCheckIcon(): SVGSVGElement {
  const ns = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(ns, 'svg');
  svg.setAttribute('width', '11');
  svg.setAttribute('height', '11');
  svg.setAttribute('viewBox', '0 0 12 12');
  svg.setAttribute('fill', 'none');
  svg.setAttribute('aria-hidden', 'true');
  const path = document.createElementNS(ns, 'path');
  path.setAttribute('d', 'M2 6.5 4.8 9.2 10 3');
  path.setAttribute('stroke', 'currentColor');
  path.setAttribute('stroke-width', '2');
  path.setAttribute('stroke-linecap', 'round');
  path.setAttribute('stroke-linejoin', 'round');
  svg.appendChild(path);
  return svg;
}

function triggerDownload(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

function progressPercent(p: FetchProgress): number {
  return p.listsFound > 0
    ? Math.max(8, Math.min(98, Math.round((p.listsDone / p.listsFound) * 100)))
    : 12;
}

function updateProgressInPlace(p: FetchProgress): void {
  if (!progressDom) {
    render();
    return;
  }
  progressDom.msgEl.textContent = p.message;
  progressDom.fillEl.style.width = `${progressPercent(p)}%`;
  progressDom.countsEl.textContent =
    `Lists ${p.listsDone}/${p.listsFound} · Tasks ${p.tasksFound} · Files ${p.attachmentsDone}/${p.attachmentsFound}`;
}

function activeStep(): 1 | 2 | 3 {
  if (state.phase === 'ready') return 3;
  if (state.phase === 'signed-in' || state.phase === 'exporting') return 2;
  return 1;
}

function renderSteps(): HTMLElement {
  const current = activeStep();
  const labels: Array<{ n: 1 | 2 | 3; label: string }> = [
    { n: 1, label: 'Sign in' },
    { n: 2, label: 'Export' },
    { n: 3, label: 'Download' },
  ];

  const items: Node[] = [];
  labels.forEach((s, idx) => {
    const done = s.n < current;
    const active = s.n === current;
    const cls = `step-item${active ? ' is-active' : ''}${done ? ' is-done' : ''}`;
    items.push(
      el('div', { class: cls, role: 'listitem', ...(active ? { 'aria-current': 'step' } : {}) }, [
        el('span', { class: 'step-dot' }, [done ? createCheckIcon() : String(s.n)]),
        el('span', {}, [s.label]),
      ]),
    );
    if (idx < labels.length - 1) {
      items.push(el('span', { class: 'step-sep', 'aria-hidden': 'true' }));
    }
  });

  return el('div', { class: 'steps', role: 'list', 'aria-label': 'Export steps' }, items);
}

async function handleSignIn(): Promise<void> {
  state.error = null;
  render();
  try {
    await signInWithMicrosoft();
  } catch (err) {
    state.error = err instanceof Error ? err.message : 'Could not start Microsoft sign-in.';
    render();
  }
}

async function handleSignOut(): Promise<void> {
  abortController?.abort();
  abortController = null;
  await signOutLocal();
  state.account = null;
  state.demoMode = false;
  state.bundle = null;
  state.progress = null;
  state.error = null;
  state.phase = 'signed-out';
  render();
}

function handleStartDemo(): void {
  state.demoMode = true;
  state.error = null;
  state.bundle = buildDemoBundle();
  state.phase = 'ready';
  render();
}

async function handleRunExport(): Promise<void> {
  state.error = null;
  state.bundle = null;
  state.phase = 'exporting';
  state.progress = {
    phase: 'lists',
    message: 'Connecting to Microsoft To Do...',
    listsFound: 0,
    listsDone: 0,
    tasksFound: 0,
    attachmentsFound: 0,
    attachmentsDone: 0,
  };
  progressDom = null;
  render();

  abortController = new AbortController();

  try {
    const initialToken = await acquireGraphToken({ forceRefresh: false, allowRedirect: true });
    if (!initialToken) return;
    const bundle = await exportMicrosoftToDo({
      accessToken: initialToken,
      getAccessToken: async (forceRefresh = false) => {
        const refreshed = await acquireGraphToken({ forceRefresh, allowRedirect: false });
        if (!refreshed) {
          throw new Error('Your Microsoft session expired during export. Please sign in again.');
        }
        return refreshed;
      },
      includeCompleted: state.includeCompleted,
      includeAttachments: state.includeAttachments,
      signal: abortController.signal,
      onProgress: (p) => {
        state.progress = p;
        updateProgressInPlace(p);
      },
    });
    state.bundle = bundle;
    state.phase = 'ready';
  } catch (err) {
    if (err instanceof DOMException && err.name === 'AbortError') {
      state.error = 'Export canceled.';
    } else {
      state.error = err instanceof Error ? err.message : 'Export failed.';
    }
    state.phase = state.account ? 'signed-in' : 'signed-out';
  } finally {
    abortController = null;
    progressDom = null;
    render();
  }
}

async function handleDownloadZip(): Promise<void> {
  if (!state.bundle || state.busyDownload) return;
  state.busyDownload = true;
  render();
  try {
    const blob = await buildExportZip(state.bundle);
    const stamp = exportDateStamp(state.bundle.document.exportedAt);
    triggerDownload(blob, `microsoft-to-do-export-${stamp}.zip`);
  } catch (err) {
    state.error = err instanceof Error ? err.message : 'Could not build ZIP file.';
  } finally {
    state.busyDownload = false;
    render();
  }
}

function handleDownloadJson(): void {
  if (!state.bundle) return;
  const stamp = exportDateStamp(state.bundle.document.exportedAt);
  const blob = new Blob([JSON.stringify(state.bundle.document, null, 2)], {
    type: 'application/json;charset=utf-8',
  });
  triggerDownload(blob, `microsoft-to-do-export-${stamp}.json`);
}

function handleDownloadCsv(): void {
  if (!state.bundle) return;
  const stamp = exportDateStamp(state.bundle.document.exportedAt);
  const blob = new Blob([renderExportCsv(state.bundle)], {
    type: 'text/csv;charset=utf-8',
  });
  triggerDownload(blob, `microsoft-to-do-export-${stamp}.csv`);
}

function renderOptions(): HTMLElement {
  const completedInput = el('input', { type: 'checkbox' });
  completedInput.checked = state.includeCompleted;
  completedInput.addEventListener('change', () => {
    state.includeCompleted = completedInput.checked;
  });

  const attInput = el('input', { type: 'checkbox' });
  attInput.checked = state.includeAttachments;
  attInput.addEventListener('change', () => {
    state.includeAttachments = attInput.checked;
  });

  return el('div', { class: 'options' }, [
    el('label', { class: 'option-check' }, [
      completedInput,
      el('span', {}, [
        'Include completed tasks',
        el('span', { class: 'option-hint' }, ['Keep finished tasks checked off.']),
      ]),
    ]),
    el('label', { class: 'option-check' }, [
      attInput,
      el('span', {}, [
        'Include task attachments',
        el('span', { class: 'option-hint' }, ['Download files attached to tasks into the ZIP.']),
      ]),
    ]),
  ]);
}

function renderErrorBanner(): HTMLElement | null {
  if (!state.error) return null;
  return el('div', { class: 'banner banner-err', role: 'alert' }, [state.error]);
}

function render(): void {
  const root = document.getElementById('app-root');
  if (!root) return;
  progressDom = null;
  root.replaceChildren();

  root.append(renderSteps());
  const body = el('div', { class: 'panel-body' });
  root.append(body);

  const errBanner = renderErrorBanner();
  if (errBanner) body.append(errBanner);

  if (state.phase === 'init') {
    body.append(el('p', { class: 'muted' }, ['Checking sign-in status...']));
    return;
  }

  if (state.phase === 'signed-out') {
    const signInBtn = el('button', { class: 'btn btn-primary', type: 'button' }, [
      createMicrosoftLogo(),
      'Sign in with Microsoft',
    ]);
    signInBtn.addEventListener('click', () => void handleSignIn());

    const demoBtn = el('button', { class: 'btn btn-ghost', type: 'button' }, ['Preview with sample data']);
    demoBtn.addEventListener('click', handleStartDemo);

    body.append(
      el('p', { class: 'muted' }, [
        'Sign in with a personal Microsoft account (outlook.com, hotmail.com, live.com) or Microsoft 365. ',
        'Your browser fetches directly from Microsoft Graph and builds the ZIP locally - tokens live only in temporary tab storage and clear when you sign out or close the tab.',
      ]),
      el('div', { class: 'actions' }, [signInBtn, demoBtn]),
    );
    return;
  }

  if (state.phase === 'signed-in') {
    const accountLabel = state.account?.username || state.account?.name || 'Microsoft account';
    const signOutBtn = el('button', { class: 'btn btn-secondary btn-sm', type: 'button' }, ['Sign out']);
    signOutBtn.addEventListener('click', () => void handleSignOut());

    const exportBtn = el('button', { class: 'btn btn-primary', type: 'button' }, ['Fetch my Microsoft To Do data']);
    exportBtn.addEventListener('click', () => void handleRunExport());

    body.append(
      el('div', { class: 'row-between' }, [
        el('span', { class: 'account-label' }, [accountLabel]),
        signOutBtn,
      ]),
      renderOptions(),
      el('div', { class: 'actions' }, [exportBtn]),
    );
    return;
  }

  if (state.phase === 'exporting') {
    const p = state.progress!;
    const msgEl = el('p', { class: 'muted', 'aria-live': 'polite' }, [p.message]);
    const fillEl = el('div', { class: 'progress-fill' });
    fillEl.style.width = `${progressPercent(p)}%`;
    const countsEl = el('p', { class: 'note-line' }, [
      `Lists ${p.listsDone}/${p.listsFound} · Tasks ${p.tasksFound} · Files ${p.attachmentsDone}/${p.attachmentsFound}`,
    ]);
    progressDom = { msgEl, fillEl, countsEl };

    const cancelBtn = el('button', { class: 'btn btn-secondary btn-sm', type: 'button' }, ['Cancel']);
    cancelBtn.addEventListener('click', () => abortController?.abort());

    body.append(
      el('div', { class: 'progress-wrap' }, [
        msgEl,
        el('div', { class: 'progress-bar' }, [fillEl]),
        countsEl,
      ]),
      el('div', { class: 'actions' }, [cancelBtn]),
    );
    return;
  }

  if (state.phase === 'ready' && state.bundle) {
    const summary = summarizeBundle(state.bundle);
    const label = state.demoMode
      ? 'Sample export'
      : state.account?.username || state.account?.name || 'Microsoft account';

    const resetBtn = el('button', { class: 'btn btn-secondary btn-sm', type: 'button' }, [
      state.demoMode ? 'Exit sample' : 'Sign out',
    ]);
    resetBtn.addEventListener('click', () => void handleSignOut());

    const zipBtn = el('button', { class: 'btn btn-primary', type: 'button' }, [
      state.busyDownload ? 'Building ZIP...' : 'Download ZIP (recommended)',
    ]);
    zipBtn.disabled = state.busyDownload;
    zipBtn.addEventListener('click', () => void handleDownloadZip());

    const jsonBtn = el('button', { class: 'btn btn-secondary', type: 'button' }, ['JSON']);
    jsonBtn.addEventListener('click', handleDownloadJson);

    const csvBtn = el('button', { class: 'btn btn-secondary', type: 'button' }, ['CSV']);
    csvBtn.addEventListener('click', handleDownloadCsv);

    const listItems = state.bundle.document.lists.map((list) => {
      const done = list.tasks.filter((t) => t.status === 'completed').length;
      const countText =
        done > 0
          ? `${list.tasks.length} (${done} done)`
          : `${list.tasks.length}`;
      return el('li', {}, [
        el('span', {}, [list.displayName]),
        el('span', { class: 'list-count' }, [countText]),
      ]);
    });

    const nodes: HTMLElement[] = [
      el('div', { class: 'row-between' }, [
        el('span', { class: 'account-label' }, [label]),
        resetBtn,
      ]),
      el('p', { class: 'summary-line' }, [
        el('strong', {}, [String(summary.lists)]),
        ` list${summary.lists === 1 ? '' : 's'} · `,
        el('strong', {}, [String(summary.tasks)]),
        ` task${summary.tasks === 1 ? '' : 's'} · `,
        el('strong', {}, [String(summary.subtasks)]),
        ` subtask${summary.subtasks === 1 ? '' : 's'} · `,
        el('strong', {}, [String(summary.attachments)]),
        ` file${summary.attachments === 1 ? '' : 's'}`,
      ]),
      el('div', { class: 'actions' }, [zipBtn, jsonBtn, csvBtn]),
      el('p', { class: 'note-line' }, [
        'Importing into ',
        el('span', { class: 'pn-mark' }, [el('span', { class: 'pn-mark-a' }, ['Privacy']), 'Notes']),
        '? Open ',
        el('a', { class: 'brand-link', href: 'https://use.privacynotes.app', target: '_blank', rel: 'noopener noreferrer' }, [
          el('span', { class: 'pn-mark' }, [el('span', { class: 'pn-mark-a' }, ['Privacy']), 'Notes']),
        ]),
        ' > Settings > Transfer > Import > Microsoft To Do and drop this ZIP in.',
      ]),
    ];

    if (state.bundle.document.warnings.length > 0) {
      nodes.push(
        el('div', { class: 'banner banner-warn' }, [
          el('strong', {}, [
            `${state.bundle.document.warnings.length} note${state.bundle.document.warnings.length === 1 ? '' : 's'} during export:`,
          ]),
          el(
            'ul',
            {},
            state.bundle.document.warnings.slice(0, 10).map((w) => el('li', {}, [w])),
          ),
        ]),
      );
    }

    if (listItems.length > 0) {
      nodes.push(el('ul', { class: 'list-preview' }, listItems));
    }

    if (!state.demoMode) {
      const refetchBtn = el('button', { class: 'btn btn-ghost btn-sm', type: 'button' }, ['Change options and fetch again']);
      refetchBtn.addEventListener('click', () => {
        state.phase = 'signed-in';
        state.bundle = null;
        render();
      });
      nodes.push(el('div', { class: 'actions' }, [refetchBtn]));
    }

    body.append(...nodes);
  }
}

async function boot(): Promise<void> {
  if (window !== window.parent) {
    try {
      await initAuth();
    } catch {
      // Ignore iframe silent-auth errors; parent window handles fallback.
    }
    return;
  }

  initThemeSwitcher();
  render();
  const params = new URLSearchParams(window.location.search);
  if (params.get('demo') === '1') {
    handleStartDemo();
    return;
  }

  try {
    const account = await initAuth();
    state.account = account;
    state.phase = account ? 'signed-in' : 'signed-out';
  } catch (err) {
    state.error = err instanceof Error ? err.message : 'Could not initialize Microsoft sign-in.';
    state.phase = 'signed-out';
  }
  render();
}

void boot();
