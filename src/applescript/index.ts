/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * AppleScript module for Outlook integration.
 *
 * Provides AppleScript-based access to Microsoft Outlook for Mac,
 * enabling support for both classic and new Outlook versions.
 */

// Executor
export {
  executeAppleScript,
  executeAppleScriptOrThrow,
  escapeForAppleScript,
  isOutlookRunning,
  launchOutlook,
  getOutlookVersion,
  AppleScriptExecutionError,
  type AppleScriptResult,
  type ExecuteOptions,
} from './executor.js';

// Scripts
export { DELIMITERS } from './scripts.js';

// Parser
export {
  parseFolders,
  parseEmails,
  parseEmail,
  parseCalendars,
  parseEvents,
  parseEvent,
  parseContacts,
  parseContact,
  parseTasks,
  parseTask,
  parseNotes,
  parseNote,
  parseCount,
  parseAccounts,
  parseDefaultAccountId,
  parseFoldersWithAccount,
  parseRespondToEventResult,
  parseDeleteEventResult,
  parseUpdateEventResult,
  parseSendEmailResult,
  parseAttachments,
  parseSaveAttachmentResult,
  type AppleScriptFolderRow,
  type AppleScriptEmailRow,
  type AppleScriptCalendarRow,
  type AppleScriptEventRow,
  type AppleScriptContactRow,
  type AppleScriptTaskRow,
  type AppleScriptNoteRow,
  type AppleScriptAccountRow,
  type AppleScriptFolderWithAccountRow,
  type AppleScriptAttachmentRow,
  type RespondToEventResult,
  type DeleteEventResult,
  type UpdateEventResult,
  type SendEmailResult,
  type SaveAttachmentResult,
} from './parser.js';

// Repository
export {
  AppleScriptRepository,
  createAppleScriptRepository,
} from './repository.js';

// Account Repository
export {
  AccountRepository,
  createAccountRepository,
  type IAccountRepository,
} from './account-repository.js';

// Calendar Writer
export {
  AppleScriptCalendarWriter,
  createCalendarWriter,
  type ICalendarWriter,
  type CreateEventParams as CalendarWriterCreateEventParams,
  type CreatedEvent,
  type RecurrenceConfig,
} from './calendar-writer.js';

// Calendar Manager
export {
  AppleScriptCalendarManager,
  createCalendarManager,
  type ICalendarManager,
  type ResponseType,
  type ApplyToScope,
  type EventUpdates,
  type UpdatedEvent,
} from './calendar-manager.js';

// Content Readers
export {
  AppleScriptEmailContentReader,
  AppleScriptEventContentReader,
  AppleScriptContactContentReader,
  AppleScriptTaskContentReader,
  AppleScriptNoteContentReader,
  AppleScriptAttachmentReader,
  createAppleScriptContentReaders,
  createEmailPath,
  createEventPath,
  createContactPath,
  createTaskPath,
  createNotePath,
  EMAIL_PATH_PREFIX,
  EVENT_PATH_PREFIX,
  CONTACT_PATH_PREFIX,
  TASK_PATH_PREFIX,
  NOTE_PATH_PREFIX,
  type AppleScriptContentReaders,
  type IAttachmentReader,
} from './content-readers.js';

// Mail Sender
export {
  AppleScriptMailSender,
  createMailSender,
  type IMailSender,
  type SendEmailParams as MailSenderSendEmailParams,
  type SentEmail,
  type Attachment,
  type InlineImage,
} from './mail-sender.js';
