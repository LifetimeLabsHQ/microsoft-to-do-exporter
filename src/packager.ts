import JSZip from 'jszip';
import { sanitizeFileName } from './graph';
import type { ExportBundle, ExportedTaskList, GraphPatternedRecurrence } from './types';

const SAFE_URL_RE = /^(?:https?:\/\/|mailto:)/i;

const RECURRENCE_SINGULAR: Record<string, string> = {
  daily: 'daily',
  weekly: 'weekly',
  absolutemonthly: 'monthly',
  relativemonthly: 'monthly',
  absoluteyearly: 'yearly',
  relativeyearly: 'yearly',
};

const RECURRENCE_UNIT: Record<string, string> = {
  daily: 'days',
  weekly: 'weeks',
  absolutemonthly: 'months',
  relativemonthly: 'months',
  absoluteyearly: 'years',
  relativeyearly: 'years',
};

/**
 * Microsoft Graph stores To Do due/completed dates as local midnight converted
 * to UTC (e.g. Nov 25 00:00 in UTC+8 comes back as "2026-11-24T16:00:00.0000000"
 * with timeZone "UTC"). When the field is in UTC with a non-midnight time and we
 * know the user's IANA time zone, convert it back to their local calendar date.
 */
export function formatDateOnly(
  dt?: { dateTime?: string; timeZone?: string } | null,
  targetTimeZone?: string,
): string {
  if (!dt?.dateTime) return '';
  const raw = dt.dateTime.trim();
  const datePart = raw.match(/^(\d{4}-\d{2}-\d{2})/)?.[1];
  if (!datePart) return raw;

  const isUtc = !dt.timeZone || dt.timeZone.toUpperCase() === 'UTC';
  const hasNonMidnightTime = /T(?!00:00(?::00(?:\.0+)?)?(?:Z|$))/.test(raw);
  if (isUtc && hasNonMidnightTime && targetTimeZone) {
    const normalized = raw.replace(/(\.\d{3})\d+/, '$1');
    const hasExplicitZone = /Z|[+-]\d{2}:?\d{2}$/i.test(normalized);
    const ms = Date.parse(hasExplicitZone ? normalized : `${normalized}Z`);
    if (!Number.isNaN(ms)) {
      try {
        return new Intl.DateTimeFormat('en-CA', {
          timeZone: targetTimeZone,
          year: 'numeric',
          month: '2-digit',
          day: '2-digit',
        }).format(new Date(ms));
      } catch {
        return datePart;
      }
    }
  }
  return datePart;
}

export function formatRecurrence(rec?: GraphPatternedRecurrence | null): string {
  const rawType = rec?.pattern?.type?.trim();
  if (!rawType) return '';
  const key = rawType.toLowerCase();
  const interval =
    typeof rec?.pattern?.interval === 'number' && Number.isFinite(rec.pattern.interval) && rec.pattern.interval > 1
      ? Math.floor(rec.pattern.interval)
      : 1;
  const days = (rec?.pattern?.daysOfWeek ?? []).filter((d) => typeof d === 'string' && d.trim() !== '');
  const index = rec?.pattern?.index?.trim();
  const isRelative = key === 'relativemonthly' || key === 'relativeyearly';
  const daysDetail =
    days.length > 0 ? (isRelative && index ? `${index} ${days.join(', ')}` : days.join(', ')) : '';
  const daysSuffix = daysDetail ? ` (${daysDetail})` : '';

  if (interval === 1) {
    const label = Object.hasOwn(RECURRENCE_SINGULAR, key) ? RECURRENCE_SINGULAR[key]! : rawType;
    return `${label}${daysSuffix}`;
  }
  const unit = Object.hasOwn(RECURRENCE_UNIT, key) ? RECURRENCE_UNIT[key]! : rawType;
  return `every ${interval} ${unit}${daysSuffix}`;
}

export function stripHtmlForMarkdown(raw: string): string {
  return raw
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/\r\n?/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function escapeMdLinkLabel(raw: string): string {
  return raw.replace(/\r?\n/g, ' ').replace(/\[/g, '(').replace(/\]/g, ')').trim() || 'item';
}

function encodeMdLinkTarget(target: string): string {
  return target
    .trim()
    .replace(/[\u0000-\u001f\u007f]/g, '')
    .replace(/%/g, '%25')
    .replace(/\s/g, '%20')
    .replace(/\(/g, '%28')
    .replace(/\)/g, '%29')
    .replace(/</g, '%3C')
    .replace(/>/g, '%3E')
    .replace(/"/g, '%22')
    .replace(/\\/g, '%5C');
}

function encodeMdUrl(url: string): string {
  return url
    .trim()
    .replace(/[\u0000-\u001f\u007f]/g, '')
    .replace(/\s/g, '%20')
    .replace(/\(/g, '%28')
    .replace(/\)/g, '%29')
    .replace(/</g, '%3C')
    .replace(/>/g, '%3E')
    .replace(/"/g, '%22')
    .replace(/\\/g, '%5C');
}

export function renderListMarkdown(list: ExportedTaskList, timeZone?: string): string {
  const lines: string[] = [`# ${list.displayName}`, ''];

  if (list.tasks.length === 0) {
    lines.push('_No tasks in this list._', '');
    return lines.join('\n');
  }

  const ordered = [
    ...list.tasks.filter((t) => t.status !== 'completed'),
    ...list.tasks.filter((t) => t.status === 'completed'),
  ];

  for (const task of ordered) {
    const done = task.status === 'completed';
    const meta: string[] = [];
    if (task.importance === 'high') meta.push('[Important]');
    const due = formatDateOnly(task.dueDateTime, timeZone);
    if (due) meta.push(`(due: ${due})`);
    const rec = formatRecurrence(task.recurrence);
    if (rec) meta.push(`(repeats: ${rec})`);
    if (task.categories.length > 0) {
      meta.push(task.categories.map((c) => `#${c.replace(/\s+/g, '-')}`).join(' '));
    }

    const suffix = meta.length > 0 ? ` ${meta.join(' ')}` : '';
    lines.push(`- [${done ? 'x' : ' '}] ${task.title.replace(/\r?\n/g, ' ')}${suffix}`);

    for (const step of task.checklistItems) {
      lines.push(`  - [${step.isChecked ? 'x' : ' '}] ${step.displayName.replace(/\r?\n/g, ' ')}`);
    }

    const bodyText = task.body?.content
      ? (task.body.contentType === 'html' ? stripHtmlForMarkdown(task.body.content) : task.body.content.replace(/\r\n?/g, '\n').trim())
      : '';
    if (bodyText) {
      for (const noteLine of bodyText.split('\n')) {
        lines.push(noteLine ? `  > ${noteLine}` : '  >');
      }
    }

    for (const lr of task.linkedResources) {
      const label = lr.displayName || lr.applicationName || lr.webUrl || 'Linked item';
      if (lr.webUrl && SAFE_URL_RE.test(lr.webUrl.trim())) {
        lines.push(`  - Link: [${escapeMdLinkLabel(label)}](${encodeMdUrl(lr.webUrl)})`);
      } else {
        lines.push(`  - Link: ${label.replace(/\r?\n/g, ' ')}`);
      }
    }

    for (const att of task.attachments) {
      if (att.zipPath) {
        lines.push(`  - Attachment: [${escapeMdLinkLabel(att.name)}](../${encodeMdLinkTarget(att.zipPath)})`);
      } else if (att.sourceUrl && SAFE_URL_RE.test(att.sourceUrl.trim())) {
        lines.push(`  - Attachment (cloud link): [${escapeMdLinkLabel(att.name)}](${encodeMdUrl(att.sourceUrl)})`);
      } else {
        lines.push(`  - Attachment (not downloaded): ${att.name.replace(/\r?\n/g, ' ')}`);
      }
    }
  }

  lines.push('');
  return lines.join('\n');
}

function csvCell(value: string): string {
  // Neutralize spreadsheet formula injection (CWE-1236) and avoid #NAME? on bulleted notes.
  const safe = /^[\t\r ]*[=+\-@]/.test(value) ? `'${value}` : value;
  if (/[",\r\n]/.test(safe)) {
    return `"${safe.replace(/"/g, '""')}"`;
  }
  return safe;
}

export function renderExportCsv(bundle: ExportBundle): string {
  const header = [
    'List',
    'Title',
    'Status',
    'Importance',
    'Due Date',
    'Completed Date',
    'Recurrence',
    'Categories',
    'Subtasks',
    'Notes',
    'Attachments',
    'Created',
  ];
  const rows: string[] = [header.join(',')];

  const tz = bundle.document.exportedTimeZone;
  for (const list of bundle.document.lists) {
    for (const task of list.tasks) {
      const subtasks = task.checklistItems
        .map((c) => `[${c.isChecked ? 'x' : ' '}] ${c.displayName}`)
        .join(' | ');
      const notes = task.body?.content
        ? (task.body.contentType === 'html' ? stripHtmlForMarkdown(task.body.content) : task.body.content.trim())
        : '';
      const attachments = task.attachments.map((a) => a.zipPath || a.name).join(' | ');

      rows.push(
        [
          list.displayName,
          task.title,
          task.status,
          task.importance,
          formatDateOnly(task.dueDateTime, tz),
          formatDateOnly(task.completedDateTime, tz),
          formatRecurrence(task.recurrence),
          task.categories.join('; '),
          subtasks,
          notes,
          attachments,
          task.createdDateTime || '',
        ]
          .map(csvCell)
          .join(','),
      );
    }
  }

  return `${rows.join('\r\n')}\r\n`;
}

function uniqueMarkdownName(list: ExportedTaskList, used: Set<string>): string {
  const base = sanitizeFileName(list.displayName, 'Untitled list');
  let candidate = `${base}.md`;
  let n = 2;
  while (used.has(candidate.toLowerCase())) {
    candidate = `${base} (${n}).md`;
    n++;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

function buildReadme(bundle: ExportBundle): string {
  const totalTasks = bundle.document.lists.reduce((sum, l) => sum + l.tasks.length, 0);
  return [
    'Microsoft To Do Export',
    '======================',
    '',
    `Exported at: ${bundle.document.exportedAt}`,
    `Lists: ${bundle.document.lists.length}`,
    `Tasks: ${totalTasks}`,
    `Attachments: ${bundle.attachmentFiles.length}`,
    '',
    'Files included in this ZIP:',
    '- microsoft-to-do-export.json   Full structured backup (recommended for PrivacyNotes import)',
    '- markdown/*.md                 One readable Markdown checklist per list',
    '- tasks.csv                     Spreadsheet-friendly CSV of all tasks',
    '- attachments/*                 Task file attachments (when available)',
    '',
    'Import into PrivacyNotes:',
    'Open PrivacyNotes > Settings > Transfer > Import > Microsoft To Do, and drop this ZIP file in.',
    'Website: https://privacynotes.app/',
    'Try live demo: https://try.privacynotes.app/',
    'Import guide: https://privacynotes.app/help/import/microsoft-to-do',
    '',
    'This export was created locally in your browser by the open-source',
    'PrivacyNotes Microsoft To Do Exporter (https://todo.privacynotes.app).',
    '',
  ].join('\n');
}

export async function buildExportZip(bundle: ExportBundle): Promise<Blob> {
  const zip = new JSZip();
  zip.file('microsoft-to-do-export.json', JSON.stringify(bundle.document, null, 2), {
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  zip.file('tasks.csv', renderExportCsv(bundle), {
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  zip.file('README.txt', buildReadme(bundle), {
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  const usedMdNames = new Set<string>();
  for (const list of bundle.document.lists) {
    const fileName = uniqueMarkdownName(list, usedMdNames);
    zip.file(`markdown/${fileName}`, renderListMarkdown(list, bundle.document.exportedTimeZone), {
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
  }

  // Store attachments uncompressed to avoid main-thread freezes and memory spikes on images/PDFs.
  for (const file of bundle.attachmentFiles) {
    zip.file(file.zipPath, file.bytes, { compression: 'STORE' });
  }

  return zip.generateAsync({ type: 'blob' });
}

export function exportDateStamp(dateInput: Date | string = new Date()): string {
  const d = typeof dateInput === 'string' ? new Date(dateInput) : dateInput;
  if (Number.isNaN(d.getTime())) {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
  }
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

export function summarizeBundle(bundle: ExportBundle): {
  lists: number;
  tasks: number;
  completed: number;
  subtasks: number;
  attachments: number;
} {
  let tasks = 0;
  let completed = 0;
  let subtasks = 0;
  for (const list of bundle.document.lists) {
    tasks += list.tasks.length;
    for (const task of list.tasks) {
      if (task.status === 'completed') completed++;
      subtasks += task.checklistItems.length;
    }
  }
  return {
    lists: bundle.document.lists.length,
    tasks,
    completed,
    subtasks,
    attachments: bundle.attachmentFiles.length,
  };
}
