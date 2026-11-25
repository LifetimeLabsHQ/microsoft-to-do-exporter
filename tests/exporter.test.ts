import type { PublicClientApplication } from '@azure/msal-browser';
import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { _resetAuthForTests, GRAPH_SCOPES, signInWithMicrosoft, signOutLocal } from '../src/auth';
import { buildDemoBundle } from '../src/demoData';
import {
  defaultSleep,
  exportMicrosoftToDo,
  GraphApiError,
  parseRetryAfterMs,
  sanitizeFileName,
} from '../src/graph';
import {
  buildExportZip,
  exportDateStamp,
  formatDateOnly,
  formatRecurrence,
  renderExportCsv,
  renderListMarkdown,
  stripHtmlForMarkdown,
  summarizeBundle,
} from '../src/packager';

function jsonResponse(body: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json', ...headers },
  });
}

afterEach(() => {
  _resetAuthForTests(null);
});

describe('Microsoft To Do exporter', () => {
  it('paginates lists and tasks, retries 429, falls back on 400 $expand, downloads attachments, handles referenceAttachments, and warns on non-auth attachment errors', async () => {
    const calls: string[] = [];
    let rateLimitedOnce = false;

    const fakeFetch: typeof fetch = async (input) => {
      const url = typeof input === 'string' ? input : input.toString();
      calls.push(url);

      if (url === 'https://graph.microsoft.com/v1.0/me/todo/lists?$top=100') {
        if (!rateLimitedOnce) {
          rateLimitedOnce = true;
          return jsonResponse({ error: { code: 'TooManyRequests', message: 'Slow down' } }, 429, {
            'Retry-After': '1',
          });
        }
        return jsonResponse({
          value: [
            { id: 'L1', displayName: 'Tasks', wellknownListName: 'defaultList', isOwner: true, isShared: false },
          ],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/todo/lists?$skip=1',
        });
      }

      if (url === 'https://graph.microsoft.com/v1.0/me/todo/lists?$skip=1') {
        return jsonResponse({
          value: [
            { id: 'L2', displayName: 'Work / Q4', wellknownListName: 'none', isOwner: true, isShared: true },
          ],
        });
      }

      if (url.includes('/lists/L1/tasks?') && url.includes('$expand')) {
        return jsonResponse({
          value: [
            {
              id: 'T1',
              title: 'Finish taxes',
              status: 'notStarted',
              importance: 'high',
              isReminderOn: true,
              hasAttachments: true,
              categories: ['Finance'],
              dueDateTime: { dateTime: '2026-04-15T00:00:00.0000000', timeZone: 'UTC' },
              body: { content: 'Bring W-2 form', contentType: 'text' },
              checklistItems: [{ id: 'C1', displayName: 'Collect receipts', isChecked: true }],
              linkedResources: [],
            },
            {
              id: 'T2',
              title: 'Done item',
              status: 'completed',
              importance: 'normal',
              isReminderOn: false,
              hasAttachments: false,
              categories: [],
              checklistItems: [],
              linkedResources: [],
            },
          ],
        });
      }

      if (url.includes('/lists/L2/tasks?') && url.includes('$expand')) {
        return jsonResponse({ error: { code: 'BadRequest', message: 'Expand not supported' } }, 400);
      }

      if (url === 'https://graph.microsoft.com/v1.0/me/todo/lists/L2/tasks?$top=100') {
        return jsonResponse({
          value: [
            {
              id: 'T3',
              title: 'Prepare roadmap',
              status: 'inProgress',
              importance: 'normal',
              isReminderOn: false,
              hasAttachments: false,
            },
          ],
        });
      }

      if (url.includes('/lists/L2/tasks/T3/checklistItems')) {
        return jsonResponse({
          value: [{ id: 'C2', displayName: 'Draft slides', isChecked: false }],
        });
      }

      if (url.includes('/lists/L2/tasks/T3/linkedResources')) {
        return jsonResponse({ value: [] });
      }

      if (url.includes('/lists/L1/tasks/T1/attachments?$top=100')) {
        return jsonResponse({
          value: [
            { id: 'A1', name: '../receipt.txt', contentType: 'text/plain', size: 5 },
            { id: 'A2', name: 'broken.pdf', contentType: 'application/pdf', size: 10 },
            {
              '@odata.type': '#microsoft.graph.referenceAttachment',
              id: 'A3',
              name: 'OneDrive Doc.docx',
              sourceUrl: 'https://onedrive.live.com/view?id=123',
            },
          ],
        });
      }

      if (url.endsWith('/attachments/A1')) {
        return jsonResponse({
          id: 'A1',
          name: '../receipt.txt',
          contentType: 'text/plain',
          size: 5,
          contentBytes: btoa('hello'),
        });
      }

      if (url.endsWith('/attachments/A2')) {
        return jsonResponse({ error: { code: 'Forbidden', message: 'Access denied' } }, 403);
      }

      return jsonResponse({ error: { message: `Unhandled URL ${url}` } }, 404);
    };

    const bundle = await exportMicrosoftToDo({
      accessToken: 'test-token',
      includeCompleted: false,
      includeAttachments: true,
      fetchImpl: fakeFetch,
      sleepImpl: async () => {},
    });

    expect(bundle.document.schemaVersion).toBe(1);
    expect(bundle.document.source).toBe('microsoft-to-do');
    expect(bundle.document.lists).toHaveLength(2);
    expect(bundle.document.lists[0]!.tasks).toHaveLength(1);
    expect(bundle.document.lists[0]!.tasks[0]!.checklistItems).toHaveLength(1);
    expect(bundle.document.lists[1]!.tasks[0]!.checklistItems[0]!.displayName).toBe('Draft slides');
    expect(bundle.attachmentFiles).toHaveLength(1);
    expect(bundle.attachmentFiles[0]!.zipPath).toBe('attachments/001-receipt.txt');
    expect(new TextDecoder().decode(bundle.attachmentFiles[0]!.bytes)).toBe('hello');
    expect(bundle.document.warnings).toHaveLength(2);
    expect(bundle.document.warnings[0]).toContain('broken.pdf');
    expect(bundle.document.warnings[1]).toContain('cloud reference link');
    // Reference attachment did NOT waste a GET /attachments/A3 call
    expect(calls.some((u) => u.endsWith('/attachments/A3'))).toBe(false);
    expect(bundle.document.lists[0]!.tasks[0]!.attachments[2]!.sourceUrl).toBe('https://onedrive.live.com/view?id=123');
  });

  it('refreshes an expired token once on 401 mid-export and aborts fatally if refresh fails', async () => {
    let refreshCount = 0;
    const seenTokens: string[] = [];

    const fakeFetch: typeof fetch = async (input, init) => {
      const url = typeof input === 'string' ? input : input.toString();
      const auth = (init?.headers as Record<string, string> | undefined)?.Authorization ?? '';
      seenTokens.push(auth);

      if (url.includes('/lists?$top=100')) {
        return jsonResponse({
          value: [{ id: 'L1', displayName: 'Tasks', wellknownListName: 'defaultList' }],
        });
      }
      if (auth === 'Bearer old-token') {
        return jsonResponse({ error: { code: 'InvalidAuthenticationToken', message: 'Expired' } }, 401);
      }
      return jsonResponse({
        value: [{ id: 'T1', title: 'Survived refresh', status: 'notStarted' }],
      });
    };

    const bundle = await exportMicrosoftToDo({
      accessToken: 'old-token',
      getAccessToken: async (forceRefresh) => {
        if (forceRefresh) refreshCount++;
        return 'fresh-token';
      },
      includeCompleted: true,
      includeAttachments: false,
      fetchImpl: fakeFetch,
      sleepImpl: async () => {},
    });

    expect(refreshCount).toBe(1);
    expect(bundle.document.lists[0]!.tasks[0]!.title).toBe('Survived refresh');
    expect(seenTokens).toContain('Bearer fresh-token');

    // If silent refresh fails during an attachment download, 401 is fatal (never produces partial ready ZIP)
    const failingRefreshFetch: typeof fetch = async (input) => {
      const url = typeof input === 'string' ? input : input.toString();
      if (url.includes('/lists?$top=100')) {
        return jsonResponse({ value: [{ id: 'L1', displayName: 'Tasks' }] });
      }
      if (url.includes('/tasks?')) {
        return jsonResponse({ value: [{ id: 'T1', title: 'Has file', status: 'notStarted', hasAttachments: true }] });
      }
      return jsonResponse({ error: { code: 'InvalidAuthenticationToken', message: 'Expired' } }, 401);
    };

    await expect(
      exportMicrosoftToDo({
        accessToken: 'dead-token',
        getAccessToken: async () => {
          throw new Error('Refresh failed');
        },
        includeCompleted: true,
        includeAttachments: true,
        fetchImpl: failingRefreshFetch,
        sleepImpl: async () => {},
      }),
    ).rejects.toBeInstanceOf(GraphApiError);
  });

  it('rejects pagination nextLink pointing outside graph.microsoft.com', async () => {
    const fakeFetch: typeof fetch = async () =>
      jsonResponse({
        value: [],
        '@odata.nextLink': 'https://evil.example.com/steal-token',
      });

    await expect(
      exportMicrosoftToDo({
        accessToken: 'secret-token',
        includeCompleted: true,
        includeAttachments: false,
        fetchImpl: fakeFetch,
        sleepImpl: async () => {},
      }),
    ).rejects.toThrow(/outside Microsoft Graph/i);
  });

  it('parses Retry-After as seconds (>30s up to 120s) and HTTP-date, and cancels backoff sleep immediately on abort', async () => {
    expect(parseRetryAfterMs('60')).toBe(60_000);
    expect(parseRetryAfterMs('300')).toBe(120_000);
    const future = new Date(Date.now() + 45_000).toUTCString();
    const parsed = parseRetryAfterMs(future);
    expect(parsed).not.toBeNull();
    expect(parsed!).toBeGreaterThan(35_000);
    expect(parsed!).toBeLessThanOrEqual(50_000);

    const controller = new AbortController();
    const sleepPromise = defaultSleep(60_000, controller.signal);
    controller.abort();
    await expect(sleepPromise).rejects.toThrow(/canceled/i);
  });

  it('enforces maxTotalAttachmentBytes safety cap', async () => {
    const fakeFetch: typeof fetch = async (input) => {
      const url = typeof input === 'string' ? input : input.toString();
      if (url.includes('/lists?$top=100')) return jsonResponse({ value: [{ id: 'L1', displayName: 'Tasks' }] });
      if (url.includes('/tasks?')) {
        return jsonResponse({ value: [{ id: 'T1', title: 'Big file', status: 'notStarted', hasAttachments: true }] });
      }
      if (url.includes('/attachments?$top=100')) {
        return jsonResponse({ value: [{ id: 'A1', name: 'huge.bin', size: 1000, contentBytes: btoa('0123456789') }] });
      }
      return jsonResponse({}, 404);
    };

    const bundle = await exportMicrosoftToDo({
      accessToken: 'token',
      includeCompleted: true,
      includeAttachments: true,
      maxTotalAttachmentBytes: 5,
      fetchImpl: fakeFetch,
      sleepImpl: async () => {},
    });

    expect(bundle.attachmentFiles).toHaveLength(0);
    expect(bundle.document.warnings[0]).toContain('safe browser memory limit');
  });

  it('throws GraphApiError after exhausting retries on persistent 503', async () => {
    let attempts = 0;
    const fakeFetch: typeof fetch = async () => {
      attempts++;
      return jsonResponse({ error: { code: 'ServiceUnavailable', message: 'Busy' } }, 503);
    };

    await expect(
      exportMicrosoftToDo({
        accessToken: 'token',
        includeCompleted: true,
        includeAttachments: false,
        fetchImpl: fakeFetch,
        sleepImpl: async () => {},
      }),
    ).rejects.toBeInstanceOf(GraphApiError);
    expect(attempts).toBe(5);
  });

  it('builds a valid ZIP, neutralizes CSV formula injection, strips <style>/<script> safely, and encodes Markdown links', async () => {
    const bundle = buildDemoBundle();
    bundle.document.lists.push({
      id: 'list-dup',
      displayName: 'Tasks',
      wellknownListName: 'none',
      isOwner: true,
      isShared: false,
      tasks: [
        {
          id: 't-sec',
          title: '=HYPERLINK("https://evil.example","Click")',
          status: 'notStarted',
          importance: 'normal',
          isReminderOn: false,
          categories: [],
          body: {
            content: '<style>body{color:red}</style><script>alert(1)</script><p>- Bullet &amp;lt;literal&amp;gt;</p>',
            contentType: 'html',
          },
          checklistItems: [],
          linkedResources: [
            { id: 'lr-ok', displayName: 'Invoice [2026]', webUrl: 'https://example.com/path_(1)' },
            { id: 'lr-bad', displayName: 'Bad JS', webUrl: 'javascript:alert(1)' },
          ],
          attachments: [
            {
              id: 'a-space',
              name: 'Receipt [Final] (1).png',
              contentType: 'image/png',
              size: 10,
              zipPath: 'attachments/003-Receipt (Final).png',
            },
          ],
        },
      ],
    });

    const summary = summarizeBundle(bundle);
    expect(summary).toEqual({
      lists: 3,
      tasks: 6,
      completed: 2,
      subtasks: 2,
      attachments: 2,
    });

    const blob = await buildExportZip(bundle);
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const names = Object.keys(zip.files).sort();

    expect(names).toContain('microsoft-to-do-export.json');
    expect(names).toContain('tasks.csv');
    expect(names).toContain('README.txt');
    expect(names).toContain('markdown/Tasks.md');
    expect(names).toContain('markdown/Tasks (2).md');
    expect(names).toContain('markdown/Groceries.md');
    expect(names).toContain('attachments/001-receipt-pixel.png');
    expect(names).toContain('attachments/002-warranty.txt');

    const md = renderListMarkdown(bundle.document.lists[0]!);
    expect(md).toContain('- [ ] Move tasks off Microsoft To Do before June 2027 [Important] (due: 2026-12-01) #Migration #Personal');
    expect(md).toContain('  - [x] Export all lists as ZIP');
    expect(md).toContain('../attachments/001-receipt-pixel.png');
    expect(md).toContain('(repeats: every 3 months)');
    expect(md.indexOf('Pay quarterly internet bill')).toBeLessThan(md.indexOf('Renew domain registration'));

    const secMd = renderListMarkdown(bundle.document.lists[2]!);
    expect(secMd).not.toContain('color:red');
    expect(secMd).not.toContain('alert(1)');
    expect(secMd).not.toContain('javascript:');
    expect(secMd).toContain('&lt;literal&gt;');
    expect(secMd).toContain('[Invoice (2026)](https://example.com/path_%281%29)');
    expect(secMd).toContain('[Receipt (Final) (1).png](../attachments/003-Receipt%20%28Final%29.png)');

    const csv = renderExportCsv(bundle);
    expect(csv).toContain('List,Title,Status,Importance');
    expect(csv).toContain(`"'=HYPERLINK(""https://evil.example"",""Click"")"`);
    expect(csv).toContain(`'- Bullet &lt;literal&gt;`);
    expect(csv).not.toContain('color:red');

    // Direct stripHtmlForMarkdown check
    expect(stripHtmlForMarkdown('<style>.a{}</style><p>Hello &amp;lt;world&amp;gt;</p>')).toBe('Hello &lt;world&gt;');
  });

  it('formats recurrence patterns human-readably, resolves UTC due dates cleanly, and formats local exportDateStamp', () => {
    expect(formatRecurrence({ pattern: { type: 'daily', interval: 1 } })).toBe('daily');
    expect(formatRecurrence({ pattern: { type: 'weekly', interval: 2, daysOfWeek: ['monday', 'friday'] } })).toBe(
      'every 2 weeks (monday, friday)',
    );
    // Live Microsoft Graph sets default index: "first" on weekly patterns; only relativeMonthly/relativeYearly should show it.
    expect(
      formatRecurrence({
        pattern: { type: 'weekly', interval: 1, daysOfWeek: ['tuesday'], index: 'first' },
      }),
    ).toBe('weekly (tuesday)');
    expect(
      formatRecurrence({
        pattern: { type: 'weekly', interval: 2, daysOfWeek: ['monday', 'thursday'], index: 'first' },
      }),
    ).toBe('every 2 weeks (monday, thursday)');
    expect(formatRecurrence({ pattern: { type: 'AbsoluteMonthly', interval: 3 } })).toBe('every 3 months');
    expect(formatRecurrence({ pattern: { type: 'relativeMonthly', interval: 1, index: 'first', daysOfWeek: ['monday'] } })).toBe(
      'monthly (first monday)',
    );
    expect(formatRecurrence(null)).toBe('');

    expect(formatDateOnly({ dateTime: '2026-11-24T16:00:00.0000000', timeZone: 'UTC' }, 'Asia/Manila')).toBe('2026-11-25');
    expect(formatDateOnly({ dateTime: '2026-11-25T05:00:00.0000000', timeZone: 'UTC' }, 'America/New_York')).toBe('2026-11-25');
    expect(formatDateOnly({ dateTime: '2026-12-01T00:00:00.0000000', timeZone: 'UTC' }, 'Asia/Manila')).toBe('2026-12-01');
    expect(formatDateOnly({ dateTime: '2026-12-01T00:00', timeZone: 'UTC' }, 'America/New_York')).toBe('2026-12-01');
    expect(formatDateOnly({ dateTime: '2026-11-24T16:00:00+08:00' }, 'Asia/Manila')).toBe('2026-11-24');

    const localMidday = new Date(2026, 10, 24, 12, 0, 0);
    expect(exportDateStamp(localMidday)).toBe('2026-11-24');
  });

  it('sanitizes unsafe attachment and list filenames', () => {
    expect(sanitizeFileName('../../evil:file*name?.pdf')).toBe('evil_file_name_.pdf');
    expect(sanitizeFileName('...')).toBe('...');
    expect(sanitizeFileName('..')).toBe('attachment');
    expect(sanitizeFileName('   ')).toBe('attachment');
  });

  it('auth uses Tasks.Read only and passes prompt: select_account after local sign-out', async () => {
    expect(GRAPH_SCOPES).toEqual(['Tasks.Read']);

    const storage = new Map<string, string>();
    const fakeSessionStorage = {
      getItem: (k: string) => storage.get(k) ?? null,
      setItem: (k: string, v: string) => { storage.set(k, v); },
      removeItem: (k: string) => { storage.delete(k); },
      clear: () => { storage.clear(); },
    };
    vi.stubGlobal('sessionStorage', fakeSessionStorage);

    const loginRedirect = vi.fn(async () => {});
    const setActiveAccount = vi.fn();
    const clearCache = vi.fn(async () => {});
    const fakeMsal = {
      loginRedirect,
      setActiveAccount,
      clearCache,
      getActiveAccount: () => ({ username: 'user@example.com' }),
      getAllAccounts: () => [],
    } as unknown as PublicClientApplication;

    _resetAuthForTests(fakeMsal);

    // First sign-in: no prompt so active personal MSA sessions go straight to consent
    await signInWithMicrosoft();
    expect(loginRedirect).toHaveBeenLastCalledWith({ scopes: ['Tasks.Read'] });

    // Sign out locally: clears storage and marks next sign-in to show account picker
    storage.set('msal.token', 'secret');
    await signOutLocal();
    expect(storage.has('msal.token')).toBe(false);
    expect(setActiveAccount).toHaveBeenCalledWith(null);
    expect(clearCache).toHaveBeenCalled();

    // Next sign-in: passes prompt: 'select_account' so user can switch accounts
    await signInWithMicrosoft();
    expect(loginRedirect).toHaveBeenLastCalledWith({ scopes: ['Tasks.Read'], prompt: 'select_account' });

    vi.unstubAllGlobals();
  });
});
