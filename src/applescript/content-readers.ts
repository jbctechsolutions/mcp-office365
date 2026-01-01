/**
 * AppleScript-based content readers.
 *
 * Implements content reader interfaces for fetching detailed content
 * directly from Outlook via AppleScript.
 *
 * Since AppleScript doesn't use file paths like SQLite does, these
 * content readers extract IDs from special path formats and fetch
 * content via AppleScript.
 */

import type { IContentReader } from '../tools/mail.js';
import type { IEventContentReader, EventDetails } from '../tools/calendar.js';
import type { IContactContentReader, ContactDetails } from '../tools/contacts.js';
import type { ITaskContentReader, TaskDetails } from '../tools/tasks.js';
import type { INoteContentReader, NoteDetails } from '../tools/notes.js';
import { executeAppleScript } from './executor.js';
import * as scripts from './scripts.js';
import * as parser from './parser.js';

// =============================================================================
// Path Format Constants
// =============================================================================

/**
 * Prefix for AppleScript email content paths.
 * Format: "applescript-email:123" where 123 is the message ID.
 */
export const EMAIL_PATH_PREFIX = 'applescript-email:';

/**
 * Prefix for AppleScript event content paths.
 * Format: "applescript-event:123" where 123 is the event ID.
 */
export const EVENT_PATH_PREFIX = 'applescript-event:';

/**
 * Prefix for AppleScript contact content paths.
 * Format: "applescript-contact:123" where 123 is the contact ID.
 */
export const CONTACT_PATH_PREFIX = 'applescript-contact:';

/**
 * Prefix for AppleScript task content paths.
 * Format: "applescript-task:123" where 123 is the task ID.
 */
export const TASK_PATH_PREFIX = 'applescript-task:';

/**
 * Prefix for AppleScript note content paths.
 * Format: "applescript-note:123" where 123 is the note ID.
 */
export const NOTE_PATH_PREFIX = 'applescript-note:';

// =============================================================================
// Utility Functions
// =============================================================================

/**
 * Extracts the ID from an AppleScript path.
 */
function extractId(path: string | null, prefix: string): number | null {
  if (path == null || !path.startsWith(prefix)) {
    return null;
  }
  const idStr = path.substring(prefix.length);
  const id = parseInt(idStr, 10);
  return isNaN(id) ? null : id;
}

/**
 * Creates an AppleScript path from an entity ID.
 */
export function createEmailPath(id: number): string {
  return `${EMAIL_PATH_PREFIX}${id}`;
}

export function createEventPath(id: number): string {
  return `${EVENT_PATH_PREFIX}${id}`;
}

export function createContactPath(id: number): string {
  return `${CONTACT_PATH_PREFIX}${id}`;
}

export function createTaskPath(id: number): string {
  return `${TASK_PATH_PREFIX}${id}`;
}

export function createNotePath(id: number): string {
  return `${NOTE_PATH_PREFIX}${id}`;
}

// =============================================================================
// Email Content Reader
// =============================================================================

/**
 * AppleScript-based email content reader.
 */
export class AppleScriptEmailContentReader implements IContentReader {
  readEmailBody(dataFilePath: string | null): string | null {
    const id = extractId(dataFilePath, EMAIL_PATH_PREFIX);
    if (id == null) {
      return null;
    }

    try {
      const script = scripts.getMessage(id);
      const result = executeAppleScript(script);
      if (!result.success) {
        return null;
      }

      const email = parser.parseEmail(result.output);
      if (email == null) {
        return null;
      }

      // Return HTML content if available, otherwise plain text
      return email.htmlContent ?? email.plainContent;
    } catch {
      return null;
    }
  }
}

// =============================================================================
// Event Content Reader
// =============================================================================

/**
 * AppleScript-based event content reader.
 */
export class AppleScriptEventContentReader implements IEventContentReader {
  readEventDetails(dataFilePath: string | null): EventDetails | null {
    const id = extractId(dataFilePath, EVENT_PATH_PREFIX);
    if (id == null) {
      return null;
    }

    try {
      const script = scripts.getEvent(id);
      const result = executeAppleScript(script);
      if (!result.success) {
        return null;
      }

      const event = parser.parseEvent(result.output);
      if (event == null) {
        return null;
      }

      return {
        title: event.subject,
        location: event.location,
        description: event.htmlContent ?? event.plainContent,
        organizer: event.organizer,
        attendees: event.attendees.map((a) => ({
          email: a.email,
          name: a.name,
          status: 'unknown' as const,
        })),
      };
    } catch {
      return null;
    }
  }
}

// =============================================================================
// Contact Content Reader
// =============================================================================

/**
 * AppleScript-based contact content reader.
 */
export class AppleScriptContactContentReader implements IContactContentReader {
  readContactDetails(dataFilePath: string | null): ContactDetails | null {
    const id = extractId(dataFilePath, CONTACT_PATH_PREFIX);
    if (id == null) {
      return null;
    }

    try {
      const script = scripts.getContact(id);
      const result = executeAppleScript(script);
      if (!result.success) {
        return null;
      }

      const contact = parser.parseContact(result.output);
      if (contact == null) {
        return null;
      }

      // Build emails array
      const emails: { type: string; address: string }[] = contact.emails.map((e) => ({
        type: 'work',
        address: e,
      }));

      // Build phones array
      const phones: { type: string; number: string }[] = [];
      if (contact.homePhone != null) {
        phones.push({ type: 'home', number: contact.homePhone });
      }
      if (contact.workPhone != null) {
        phones.push({ type: 'work', number: contact.workPhone });
      }
      if (contact.mobilePhone != null) {
        phones.push({ type: 'mobile', number: contact.mobilePhone });
      }

      // Build addresses array
      const addresses: {
        type: string;
        street: string | null;
        city: string | null;
        state: string | null;
        postalCode: string | null;
        country: string | null;
      }[] = [];

      if (
        contact.homeStreet != null ||
        contact.homeCity != null ||
        contact.homeState != null ||
        contact.homeZip != null ||
        contact.homeCountry != null
      ) {
        addresses.push({
          type: 'home',
          street: contact.homeStreet,
          city: contact.homeCity,
          state: contact.homeState,
          postalCode: contact.homeZip,
          country: contact.homeCountry,
        });
      }

      return {
        firstName: contact.firstName,
        lastName: contact.lastName,
        middleName: contact.middleName,
        nickname: contact.nickname,
        company: contact.company,
        jobTitle: contact.jobTitle,
        department: contact.department,
        emails,
        phones,
        addresses,
        notes: contact.notes,
      };
    } catch {
      return null;
    }
  }
}

// =============================================================================
// Task Content Reader
// =============================================================================

/**
 * AppleScript-based task content reader.
 */
export class AppleScriptTaskContentReader implements ITaskContentReader {
  readTaskDetails(dataFilePath: string | null): TaskDetails | null {
    const id = extractId(dataFilePath, TASK_PATH_PREFIX);
    if (id == null) {
      return null;
    }

    try {
      const script = scripts.getTask(id);
      const result = executeAppleScript(script);
      if (!result.success) {
        return null;
      }

      const task = parser.parseTask(result.output);
      if (task == null) {
        return null;
      }

      return {
        body: task.htmlContent ?? task.plainContent,
        completedDate: task.completedDate,
        reminderDate: null,
        categories: [],
      };
    } catch {
      return null;
    }
  }
}

// =============================================================================
// Note Content Reader
// =============================================================================

/**
 * AppleScript-based note content reader.
 */
export class AppleScriptNoteContentReader implements INoteContentReader {
  readNoteDetails(dataFilePath: string | null): NoteDetails | null {
    const id = extractId(dataFilePath, NOTE_PATH_PREFIX);
    if (id == null) {
      return null;
    }

    try {
      const script = scripts.getNote(id);
      const result = executeAppleScript(script);
      if (!result.success) {
        return null;
      }

      const note = parser.parseNote(result.output);
      if (note == null) {
        return null;
      }

      const body = note.htmlContent ?? note.plainContent ?? '';
      const preview = body.substring(0, 200);

      return {
        title: note.name,
        body,
        preview,
        createdDate: note.createdDate,
        categories: [],
      };
    } catch {
      return null;
    }
  }
}

// =============================================================================
// Factory Functions
// =============================================================================

/**
 * All AppleScript content readers bundled together.
 */
export interface AppleScriptContentReaders {
  readonly email: IContentReader;
  readonly event: IEventContentReader;
  readonly contact: IContactContentReader;
  readonly task: ITaskContentReader;
  readonly note: INoteContentReader;
}

/**
 * Creates all AppleScript content readers.
 */
export function createAppleScriptContentReaders(): AppleScriptContentReaders {
  return {
    email: new AppleScriptEmailContentReader(),
    event: new AppleScriptEventContentReader(),
    contact: new AppleScriptContactContentReader(),
    task: new AppleScriptTaskContentReader(),
    note: new AppleScriptNoteContentReader(),
  };
}
