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
  type AppleScriptFolderRow,
  type AppleScriptEmailRow,
  type AppleScriptCalendarRow,
  type AppleScriptEventRow,
  type AppleScriptContactRow,
  type AppleScriptTaskRow,
  type AppleScriptNoteRow,
} from './parser.js';

// Repository
export {
  AppleScriptRepository,
  createAppleScriptRepository,
} from './repository.js';

// Content Readers
export {
  AppleScriptEmailContentReader,
  AppleScriptEventContentReader,
  AppleScriptContactContentReader,
  AppleScriptTaskContentReader,
  AppleScriptNoteContentReader,
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
} from './content-readers.js';
