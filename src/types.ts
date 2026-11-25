export interface GraphDateTimeTimeZone {
  dateTime: string;
  timeZone: string;
}

export interface GraphItemBody {
  content: string;
  contentType: 'text' | 'html' | string;
}

export interface GraphChecklistItem {
  id: string;
  displayName: string;
  isChecked: boolean;
  createdDateTime?: string;
  checkedDateTime?: string;
}

export interface GraphLinkedResource {
  id: string;
  webUrl?: string;
  applicationName?: string;
  displayName?: string;
  externalId?: string;
}

export interface GraphRecurrencePattern {
  type?: string;
  interval?: number;
  month?: number;
  dayOfMonth?: number;
  daysOfWeek?: string[];
  firstDayOfWeek?: string;
  index?: string;
}

export interface GraphRecurrenceRange {
  type?: string;
  startDate?: string;
  endDate?: string;
  recurrenceTimeZone?: string;
  numberOfOccurrences?: number;
}

export interface GraphPatternedRecurrence {
  pattern?: GraphRecurrencePattern;
  range?: GraphRecurrenceRange;
}

export interface GraphAttachmentMeta {
  '@odata.type'?: string;
  id: string;
  name: string;
  contentType?: string;
  size?: number;
  lastModifiedDateTime?: string;
  contentBytes?: string;
  sourceUrl?: string;
}

export interface ExportedAttachment {
  id: string;
  name: string;
  contentType: string;
  size: number;
  lastModifiedDateTime?: string;
  /** Relative path inside the exported ZIP, e.g. "attachments/001-receipt.pdf". */
  zipPath?: string;
  /** External reference URL when Graph returns a referenceAttachment instead of file bytes. */
  sourceUrl?: string;
  /** Present when the attachment could not be downloaded from Microsoft Graph. */
  error?: string;
}

export interface ExportedTask {
  id: string;
  title: string;
  status: 'notStarted' | 'inProgress' | 'completed' | 'waitingOnOthers' | 'deferred' | string;
  importance: 'low' | 'normal' | 'high' | string;
  isReminderOn: boolean;
  hasAttachments?: boolean;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  completedDateTime?: GraphDateTimeTimeZone | null;
  dueDateTime?: GraphDateTimeTimeZone | null;
  startDateTime?: GraphDateTimeTimeZone | null;
  reminderDateTime?: GraphDateTimeTimeZone | null;
  recurrence?: GraphPatternedRecurrence | null;
  categories: string[];
  body: GraphItemBody | null;
  checklistItems: GraphChecklistItem[];
  linkedResources: GraphLinkedResource[];
  attachments: ExportedAttachment[];
}

export interface ExportedTaskList {
  id: string;
  displayName: string;
  wellknownListName: 'none' | 'defaultList' | 'flaggedEmails' | 'unknownFutureValue' | string;
  isOwner: boolean;
  isShared: boolean;
  tasks: ExportedTask[];
}

export interface MicrosoftToDoExportDocument {
  schemaVersion: 1;
  source: 'microsoft-to-do';
  exporter: 'privacynotes-microsoft-to-do-exporter';
  exportedAt: string;
  /** Browser IANA time zone at export time (e.g. "America/New_York"), used to resolve UTC date-only fields. */
  exportedTimeZone?: string;
  options: {
    includeCompleted: boolean;
    includeAttachments: boolean;
  };
  lists: ExportedTaskList[];
  warnings: string[];
}

export interface AttachmentPayload {
  zipPath: string;
  bytes: Uint8Array;
}

export interface ExportBundle {
  document: MicrosoftToDoExportDocument;
  attachmentFiles: AttachmentPayload[];
}

export interface FetchProgress {
  phase: 'lists' | 'tasks' | 'attachments' | 'done';
  message: string;
  listsFound: number;
  listsDone: number;
  tasksFound: number;
  attachmentsFound: number;
  attachmentsDone: number;
}
