import type {
  AttachmentPayload,
  ExportBundle,
  ExportedAttachment,
  ExportedTask,
  ExportedTaskList,
  FetchProgress,
  GraphAttachmentMeta,
  GraphChecklistItem,
  GraphLinkedResource,
} from './types';

export const GRAPH_TODO_BASE = 'https://graph.microsoft.com/v1.0/me/todo';
export const MAX_RETRY_AFTER_MS = 120_000;
export const MAX_TOTAL_ATTACHMENT_BYTES = 750 * 1024 * 1024;

export interface FetchOptions {
  /** Static token (handy for tests) or initial token alongside `getAccessToken`. */
  accessToken?: string;
  /** Token provider called initially and with `forceRefresh = true` on HTTP 401. */
  getAccessToken?: (forceRefresh?: boolean) => Promise<string>;
  includeCompleted: boolean;
  includeAttachments: boolean;
  signal?: AbortSignal;
  onProgress?: (progress: FetchProgress) => void;
  fetchImpl?: typeof fetch;
  sleepImpl?: (ms: number, signal?: AbortSignal) => Promise<void>;
  maxTotalAttachmentBytes?: number;
}

interface GraphCollection<T> {
  value?: T[];
  '@odata.nextLink'?: string;
}

interface RawGraphTaskList {
  id: string;
  displayName?: string;
  wellknownListName?: string;
  isOwner?: boolean;
  isShared?: boolean;
}

interface RawGraphTask {
  id: string;
  title?: string;
  status?: string;
  importance?: string;
  isReminderOn?: boolean;
  hasAttachments?: boolean;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  completedDateTime?: { dateTime: string; timeZone: string } | null;
  dueDateTime?: { dateTime: string; timeZone: string } | null;
  startDateTime?: { dateTime: string; timeZone: string } | null;
  reminderDateTime?: { dateTime: string; timeZone: string } | null;
  recurrence?: ExportedTask['recurrence'];
  categories?: string[];
  body?: { content?: string; contentType?: string } | null;
  checklistItems?: GraphChecklistItem[];
  linkedResources?: GraphLinkedResource[];
}

export class GraphApiError extends Error {
  constructor(
    public readonly status: number,
    message: string,
  ) {
    super(message);
    this.name = 'GraphApiError';
  }
}

const RETRYABLE_STATUSES = new Set([429, 500, 502, 503, 504]);
const MAX_RETRIES = 4;

export function defaultSleep(ms: number, signal?: AbortSignal): Promise<void> {
  if (signal?.aborted) {
    return Promise.reject(new DOMException('Export canceled', 'AbortError'));
  }
  return new Promise((resolve, reject) => {
    const onAbort = () => {
      clearTimeout(timer);
      reject(new DOMException('Export canceled', 'AbortError'));
    };
    const timer = setTimeout(() => {
      signal?.removeEventListener('abort', onAbort);
      resolve();
    }, ms);
    signal?.addEventListener('abort', onAbort, { once: true });
  });
}

export function parseRetryAfterMs(header: string | null, nowMs = Date.now()): number | null {
  if (!header) return null;
  const trimmed = header.trim();
  if (!trimmed) return null;
  if (/^\d+$/.test(trimmed)) {
    const secs = Number.parseInt(trimmed, 10);
    if (Number.isFinite(secs) && secs > 0) {
      return Math.min(secs * 1000, MAX_RETRY_AFTER_MS);
    }
    return null;
  }
  const dateMs = Date.parse(trimmed);
  if (!Number.isNaN(dateMs)) {
    const delta = dateMs - nowMs;
    if (delta > 0) {
      return Math.min(delta, MAX_RETRY_AFTER_MS);
    }
  }
  return null;
}

export function sanitizeFileName(raw: string, fallback = 'attachment'): string {
  const base = raw
    .replace(/\\/g, '/')
    .split('/')
    .pop()
    ?.replace(/[\u0000-\u001f<>:"/\\|?*]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();
  if (!base || base === '.' || base === '..') return fallback;
  return base.slice(0, 120);
}

export function decodeBase64Bytes(base64: string): Uint8Array {
  const cleaned = base64.replace(/\s+/g, '');
  const uint8WithFromBase64 = Uint8Array as unknown as {
    fromBase64?: (input: string) => Uint8Array;
  };
  if (typeof uint8WithFromBase64.fromBase64 === 'function') {
    return uint8WithFromBase64.fromBase64(cleaned);
  }
  const binary = atob(cleaned);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function isAbortError(err: unknown): boolean {
  return err instanceof DOMException && err.name === 'AbortError';
}

function isFatalAuthError(err: unknown): boolean {
  return err instanceof GraphApiError && err.status === 401;
}

interface TokenBox {
  current: string;
  refresh?: () => Promise<string>;
}

async function graphFetchJson<T>(
  url: string,
  tokenBox: TokenBox,
  fetchImpl: typeof fetch,
  sleepImpl: (ms: number, signal?: AbortSignal) => Promise<void>,
  signal?: AbortSignal,
): Promise<T> {
  let attempt = 0;
  let refreshedAuth = false;

  while (true) {
    if (signal?.aborted) {
      throw new DOMException('Export canceled', 'AbortError');
    }
    const res = await fetchImpl(url, {
      method: 'GET',
      headers: {
        Authorization: `Bearer ${tokenBox.current}`,
        Accept: 'application/json',
      },
      signal,
    });

    if (res.ok) {
      return (await res.json()) as T;
    }

    if (res.status === 401 && !refreshedAuth && tokenBox.refresh) {
      refreshedAuth = true;
      try {
        tokenBox.current = await tokenBox.refresh();
      } catch {
        throw new GraphApiError(401, 'Your Microsoft session expired during export. Please sign in again.');
      }
      continue;
    }

    if (RETRYABLE_STATUSES.has(res.status) && attempt < MAX_RETRIES) {
      attempt++;
      const fallbackMs = Math.min(1000 * 2 ** (attempt - 1), 15_000);
      const retryAfterMs = parseRetryAfterMs(res.headers.get('Retry-After'));
      const waitMs = retryAfterMs ?? fallbackMs;
      await sleepImpl(waitMs, signal);
      continue;
    }

    let detail = res.statusText || `HTTP ${res.status}`;
    try {
      const errBody = (await res.json()) as { error?: { message?: string; code?: string } };
      if (errBody?.error?.message) {
        detail = `${errBody.error.code ? `${errBody.error.code}: ` : ''}${errBody.error.message}`;
      }
    } catch {
      // Keep status text.
    }
    throw new GraphApiError(res.status, detail);
  }
}

async function fetchAllPages<T>(
  initialUrl: string,
  tokenBox: TokenBox,
  fetchImpl: typeof fetch,
  sleepImpl: (ms: number, signal?: AbortSignal) => Promise<void>,
  signal?: AbortSignal,
): Promise<T[]> {
  const items: T[] = [];
  let nextUrl: string | undefined = initialUrl;
  const seenUrls = new Set<string>();

  while (nextUrl) {
    if (seenUrls.has(nextUrl)) break;
    seenUrls.add(nextUrl);
    // Guard against unexpected redirect/nextLink pointing outside Graph.
    if (!nextUrl.startsWith('https://graph.microsoft.com/')) {
      throw new Error('Unexpected pagination link outside Microsoft Graph.');
    }
    const page: GraphCollection<T> = await graphFetchJson<GraphCollection<T>>(
      nextUrl,
      tokenBox,
      fetchImpl,
      sleepImpl,
      signal,
    );
    if (Array.isArray(page.value)) {
      items.push(...page.value);
    }
    nextUrl = page['@odata.nextLink'];
  }
  return items;
}

export async function exportMicrosoftToDo(options: FetchOptions): Promise<ExportBundle> {
  const fetchImpl = options.fetchImpl ?? fetch;
  const sleepImpl = options.sleepImpl ?? defaultSleep;
  const { includeCompleted, includeAttachments, signal, onProgress } = options;
  const maxAttachmentBytes = options.maxTotalAttachmentBytes ?? MAX_TOTAL_ATTACHMENT_BYTES;

  const initialToken =
    options.accessToken ?? (options.getAccessToken ? await options.getAccessToken(false) : '');
  if (!initialToken) {
    throw new GraphApiError(401, 'Missing Microsoft Graph access token.');
  }
  const tokenBox: TokenBox = {
    current: initialToken,
    refresh: options.getAccessToken ? () => options.getAccessToken!(true) : undefined,
  };

  const warnings: string[] = [];
  const attachmentFiles: AttachmentPayload[] = [];
  let attachmentSeq = 0;
  let totalAttachmentBytes = 0;

  const progress: FetchProgress = {
    phase: 'lists',
    message: 'Loading your Microsoft To Do lists...',
    listsFound: 0,
    listsDone: 0,
    tasksFound: 0,
    attachmentsFound: 0,
    attachmentsDone: 0,
  };
  onProgress?.({ ...progress });

  const rawLists = await fetchAllPages<RawGraphTaskList>(
    `${GRAPH_TODO_BASE}/lists?$top=100`,
    tokenBox,
    fetchImpl,
    sleepImpl,
    signal,
  );

  progress.listsFound = rawLists.length;
  progress.phase = 'tasks';
  progress.message = `Found ${rawLists.length} list${rawLists.length === 1 ? '' : 's'}. Fetching tasks...`;
  onProgress?.({ ...progress });

  const exportedLists: ExportedTaskList[] = [];

  for (const rawList of rawLists) {
    const listName = rawList.displayName?.trim() || 'Untitled list';
    progress.message = `Reading "${listName}" (${progress.listsDone + 1}/${rawLists.length})...`;
    onProgress?.({ ...progress });

    const encodedListId = encodeURIComponent(rawList.id);
    const expandedUrl = `${GRAPH_TODO_BASE}/lists/${encodedListId}/tasks?$top=100&$expand=checklistItems,linkedResources`;
    const plainUrl = `${GRAPH_TODO_BASE}/lists/${encodedListId}/tasks?$top=100`;

    let rawTasks: RawGraphTask[];
    let usedExpand = true;
    try {
      rawTasks = await fetchAllPages<RawGraphTask>(expandedUrl, tokenBox, fetchImpl, sleepImpl, signal);
    } catch (err) {
      if (isAbortError(err) || isFatalAuthError(err)) throw err;
      // Only fall back when Graph explicitly rejects $expand (400 / 501).
      if (!(err instanceof GraphApiError && (err.status === 400 || err.status === 501))) {
        throw err;
      }
      usedExpand = false;
      rawTasks = await fetchAllPages<RawGraphTask>(plainUrl, tokenBox, fetchImpl, sleepImpl, signal);
    }

    const filteredTasks = includeCompleted
      ? rawTasks
      : rawTasks.filter((t) => (t.status ?? '').toLowerCase() !== 'completed');

    progress.tasksFound += filteredTasks.length;
    onProgress?.({ ...progress });

    const tasksOut: ExportedTask[] = [];

    for (const rawTask of filteredTasks) {
      const encodedTaskId = encodeURIComponent(rawTask.id);
      let checklistItems: GraphChecklistItem[] = Array.isArray(rawTask.checklistItems)
        ? rawTask.checklistItems
        : [];
      let linkedResources: GraphLinkedResource[] = Array.isArray(rawTask.linkedResources)
        ? rawTask.linkedResources
        : [];

      if (!usedExpand) {
        try {
          checklistItems = await fetchAllPages<GraphChecklistItem>(
            `${GRAPH_TODO_BASE}/lists/${encodedListId}/tasks/${encodedTaskId}/checklistItems?$top=100`,
            tokenBox,
            fetchImpl,
            sleepImpl,
            signal,
          );
        } catch (err) {
          if (isAbortError(err) || isFatalAuthError(err)) throw err;
          warnings.push(`Could not load subtasks for "${rawTask.title || 'Untitled task'}".`);
        }

        try {
          linkedResources = await fetchAllPages<GraphLinkedResource>(
            `${GRAPH_TODO_BASE}/lists/${encodedListId}/tasks/${encodedTaskId}/linkedResources?$top=100`,
            tokenBox,
            fetchImpl,
            sleepImpl,
            signal,
          );
        } catch (err) {
          if (isAbortError(err) || isFatalAuthError(err)) throw err;
          // Linked resources are optional; skip when unsupported.
        }
      }

      const exportedAttachments: ExportedAttachment[] = [];
      if (includeAttachments && rawTask.hasAttachments) {
        progress.phase = 'attachments';
        progress.message = `Downloading attachments for "${rawTask.title || 'Untitled task'}"...`;
        onProgress?.({ ...progress });

        try {
          const attMetas = await fetchAllPages<GraphAttachmentMeta>(
            `${GRAPH_TODO_BASE}/lists/${encodedListId}/tasks/${encodedTaskId}/attachments?$top=100`,
            tokenBox,
            fetchImpl,
            sleepImpl,
            signal,
          );
          progress.attachmentsFound += attMetas.length;
          onProgress?.({ ...progress });

          for (const meta of attMetas) {
            const safeName = sanitizeFileName(meta.name || 'attachment');
            try {
              const odataType = (meta['@odata.type'] ?? '').toLowerCase();
              if (odataType.includes('referenceattachment') || (!meta.contentBytes && meta.sourceUrl)) {
                const msg = `Attachment "${meta.name || safeName}" is a cloud reference link rather than an embedded file.`;
                warnings.push(msg);
                exportedAttachments.push({
                  id: meta.id,
                  name: meta.name || safeName,
                  contentType: meta.contentType || 'application/octet-stream',
                  size: meta.size ?? 0,
                  lastModifiedDateTime: meta.lastModifiedDateTime,
                  ...(meta.sourceUrl ? { sourceUrl: meta.sourceUrl } : {}),
                  error: msg,
                });
                continue;
              }

              const estimatedSize = meta.size ?? 0;
              if (estimatedSize > 0 && totalAttachmentBytes + estimatedSize > maxAttachmentBytes) {
                const msg = `Skipped "${meta.name || safeName}" because total downloaded attachments exceeded the safe browser memory limit.`;
                warnings.push(msg);
                exportedAttachments.push({
                  id: meta.id,
                  name: meta.name || safeName,
                  contentType: meta.contentType || 'application/octet-stream',
                  size: estimatedSize,
                  lastModifiedDateTime: meta.lastModifiedDateTime,
                  error: msg,
                });
                continue;
              }

              let contentBytes = meta.contentBytes;
              let contentType = meta.contentType || 'application/octet-stream';
              let size = estimatedSize;

              if (!contentBytes) {
                const fullAtt = await graphFetchJson<GraphAttachmentMeta>(
                  `${GRAPH_TODO_BASE}/lists/${encodedListId}/tasks/${encodedTaskId}/attachments/${encodeURIComponent(meta.id)}`,
                  tokenBox,
                  fetchImpl,
                  sleepImpl,
                  signal,
                );
                contentBytes = fullAtt.contentBytes;
                contentType = fullAtt.contentType || contentType;
                size = fullAtt.size ?? size;
              }

              if (contentBytes) {
                const bytes = decodeBase64Bytes(contentBytes);
                if (totalAttachmentBytes + bytes.byteLength > maxAttachmentBytes) {
                  const msg = `Skipped "${meta.name || safeName}" because total downloaded attachments exceeded the safe browser memory limit.`;
                  warnings.push(msg);
                  exportedAttachments.push({
                    id: meta.id,
                    name: meta.name || safeName,
                    contentType,
                    size: bytes.byteLength,
                    lastModifiedDateTime: meta.lastModifiedDateTime,
                    error: msg,
                  });
                  continue;
                }
                totalAttachmentBytes += bytes.byteLength;
                attachmentSeq++;
                const prefix = String(attachmentSeq).padStart(3, '0');
                const zipPath = `attachments/${prefix}-${safeName}`;
                attachmentFiles.push({ zipPath, bytes });
                exportedAttachments.push({
                  id: meta.id,
                  name: meta.name || safeName,
                  contentType,
                  size: bytes.byteLength || size,
                  lastModifiedDateTime: meta.lastModifiedDateTime,
                  zipPath,
                });
              } else {
                const msg = `Attachment "${meta.name || safeName}" had no file content.`;
                warnings.push(msg);
                exportedAttachments.push({
                  id: meta.id,
                  name: meta.name || safeName,
                  contentType,
                  size,
                  lastModifiedDateTime: meta.lastModifiedDateTime,
                  error: msg,
                });
              }
            } catch (err) {
              if (isAbortError(err) || isFatalAuthError(err)) throw err;
              const reason = err instanceof Error ? err.message : 'Download failed';
              warnings.push(`Skipped attachment "${meta.name || safeName}" on "${rawTask.title || 'Untitled task'}": ${reason}`);
              exportedAttachments.push({
                id: meta.id,
                name: meta.name || safeName,
                contentType: meta.contentType || 'application/octet-stream',
                size: meta.size ?? 0,
                lastModifiedDateTime: meta.lastModifiedDateTime,
                error: reason,
              });
            } finally {
              progress.attachmentsDone++;
              onProgress?.({ ...progress });
            }
          }
        } catch (err) {
          if (isAbortError(err) || isFatalAuthError(err)) throw err;
          const reason = err instanceof Error ? err.message : 'Could not list attachments';
          warnings.push(`Could not fetch attachments for "${rawTask.title || 'Untitled task'}": ${reason}`);
        } finally {
          progress.phase = 'tasks';
        }
      }

      tasksOut.push({
        id: rawTask.id,
        title: rawTask.title?.trim() || 'Untitled task',
        status: rawTask.status || 'notStarted',
        importance: rawTask.importance || 'normal',
        isReminderOn: Boolean(rawTask.isReminderOn),
        hasAttachments: Boolean(rawTask.hasAttachments),
        createdDateTime: rawTask.createdDateTime,
        lastModifiedDateTime: rawTask.lastModifiedDateTime,
        completedDateTime: rawTask.completedDateTime ?? null,
        dueDateTime: rawTask.dueDateTime ?? null,
        startDateTime: rawTask.startDateTime ?? null,
        reminderDateTime: rawTask.reminderDateTime ?? null,
        recurrence: rawTask.recurrence ?? null,
        categories: Array.isArray(rawTask.categories) ? rawTask.categories.filter((c) => typeof c === 'string' && c.trim() !== '') : [],
        body: rawTask.body && typeof rawTask.body.content === 'string'
          ? { content: rawTask.body.content, contentType: rawTask.body.contentType || 'text' }
          : null,
        checklistItems,
        linkedResources,
        attachments: exportedAttachments,
      });
    }

    exportedLists.push({
      id: rawList.id,
      displayName: listName,
      wellknownListName: rawList.wellknownListName || 'none',
      isOwner: rawList.isOwner ?? true,
      isShared: rawList.isShared ?? false,
      tasks: tasksOut,
    });

    progress.listsDone++;
    onProgress?.({ ...progress });
  }

  progress.phase = 'done';
  progress.message = 'Export ready to download.';
  onProgress?.({ ...progress });

  let exportedTimeZone: string | undefined;
  try {
    exportedTimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone || undefined;
  } catch {
    exportedTimeZone = undefined;
  }

  return {
    document: {
      schemaVersion: 1,
      source: 'microsoft-to-do',
      exporter: 'privacynotes-microsoft-to-do-exporter',
      exportedAt: new Date().toISOString(),
      ...(exportedTimeZone ? { exportedTimeZone } : {}),
      options: {
        includeCompleted,
        includeAttachments,
      },
      lists: exportedLists,
      warnings,
    },
    attachmentFiles,
  };
}
