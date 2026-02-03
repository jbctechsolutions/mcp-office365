/**
 * Parser for AppleScript output.
 *
 * Converts delimiter-based AppleScript output into typed objects
 * compatible with the repository row types.
 */

import { DELIMITERS } from './scripts.js';

// =============================================================================
// Types
// =============================================================================

/**
 * A parsed record from AppleScript output.
 */
type ParsedRecord = Record<string, string>;

// =============================================================================
// Row Types (matching repository interface)
// =============================================================================

export interface AppleScriptFolderRow {
  readonly id: number;
  readonly name: string | null;
  readonly unreadCount: number;
}

export interface AppleScriptEmailRow {
  readonly id: number;
  readonly folderId: number | null;
  readonly subject: string | null;
  readonly senderName: string | null;
  readonly senderEmail: string | null;
  readonly toRecipients: string | null;
  readonly ccRecipients: string | null;
  readonly preview: string | null;
  readonly isRead: boolean;
  readonly dateReceived: string | null;
  readonly dateSent: string | null;
  readonly priority: string;
  readonly htmlContent: string | null;
  readonly plainContent: string | null;
  readonly hasHtml: boolean;
  readonly attachments: string[];
}

export interface AppleScriptCalendarRow {
  readonly id: number;
  readonly name: string | null;
}

export interface AppleScriptEventRow {
  readonly id: number;
  readonly calendarId: number | null;
  readonly subject: string | null;
  readonly startTime: string | null;
  readonly endTime: string | null;
  readonly location: string | null;
  readonly isAllDay: boolean;
  readonly isRecurring: boolean;
  readonly organizer: string | null;
  readonly htmlContent: string | null;
  readonly plainContent: string | null;
  readonly attendees: Array<{ email: string; name: string }>;
}

export interface AppleScriptContactRow {
  readonly id: number;
  readonly displayName: string | null;
  readonly firstName: string | null;
  readonly lastName: string | null;
  readonly middleName: string | null;
  readonly nickname: string | null;
  readonly company: string | null;
  readonly jobTitle: string | null;
  readonly department: string | null;
  readonly notes: string | null;
  readonly emails: string[];
  readonly homePhone: string | null;
  readonly workPhone: string | null;
  readonly mobilePhone: string | null;
  readonly homeStreet: string | null;
  readonly homeCity: string | null;
  readonly homeState: string | null;
  readonly homeZip: string | null;
  readonly homeCountry: string | null;
}

export interface AppleScriptTaskRow {
  readonly id: number;
  readonly folderId: number | null;
  readonly name: string | null;
  readonly isCompleted: boolean;
  readonly dueDate: string | null;
  readonly startDate: string | null;
  readonly completedDate: string | null;
  readonly priority: string;
  readonly htmlContent: string | null;
  readonly plainContent: string | null;
}

export interface AppleScriptNoteRow {
  readonly id: number;
  readonly folderId: number | null;
  readonly name: string | null;
  readonly createdDate: string | null;
  readonly modifiedDate: string | null;
  readonly preview: string | null;
  readonly htmlContent: string | null;
  readonly plainContent: string | null;
}

export interface RespondToEventResult {
  readonly success: boolean;
  readonly eventId?: number;
  readonly error?: string;
}

export interface DeleteEventResult {
  readonly success: boolean;
  readonly eventId?: number;
  readonly error?: string;
}

export interface AppleScriptAccountRow {
  readonly id: number;
  readonly name: string | null;
  readonly email: string | null;
  readonly type: string;
}

export interface AppleScriptFolderWithAccountRow extends AppleScriptFolderRow {
  readonly accountId: number;
  readonly messageCount: number;
}

// =============================================================================
// Core Parsing Functions
// =============================================================================

/**
 * Parses the raw AppleScript output into an array of records.
 */
function parseRawOutput(output: string): ParsedRecord[] {
  if (output.trim().length === 0) {
    return [];
  }

  const records: ParsedRecord[] = [];
  const recordStrings = output.split(DELIMITERS.RECORD).filter((s) => s.length > 0);

  for (const recordStr of recordStrings) {
    const record: ParsedRecord = {};
    const fieldStrings = recordStr.split(DELIMITERS.FIELD);

    for (const fieldStr of fieldStrings) {
      const [key, value] = fieldStr.split(DELIMITERS.EQUALS);
      if (key !== undefined && value !== undefined) {
        record[key] = value;
      }
    }

    if (Object.keys(record).length > 0) {
      records.push(record);
    }
  }

  return records;
}

/**
 * Safely parses a string to a number.
 */
function parseNumber(value: string | undefined): number {
  if (value === undefined || value === '' || value === DELIMITERS.NULL) {
    return 0;
  }
  const num = parseInt(value, 10);
  return isNaN(num) ? 0 : num;
}

/**
 * Safely parses a string to a number or null.
 */
function parseNumberOrNull(value: string | undefined): number | null {
  if (value === undefined || value === '' || value === DELIMITERS.NULL) {
    return null;
  }
  const num = parseInt(value, 10);
  return isNaN(num) ? null : num;
}

/**
 * Parses a boolean value from AppleScript output.
 */
function parseBoolean(value: string | undefined): boolean {
  if (value === undefined) {
    return false;
  }
  return value.toLowerCase() === 'true';
}

/**
 * Parses a string value, returning null for empty or missing values.
 */
function parseString(value: string | undefined): string | null {
  if (value === undefined || value === '' || value === DELIMITERS.NULL) {
    return null;
  }
  return value;
}

/**
 * Parses a comma-separated list into an array.
 */
function parseList(value: string | undefined): string[] {
  if (value === undefined || value === '' || value === DELIMITERS.NULL) {
    return [];
  }
  return value.split(',').filter((s) => s.length > 0).map((s) => s.trim());
}

/**
 * Parses an attendee list (format: "email|name,email|name").
 */
function parseAttendees(value: string | undefined): Array<{ email: string; name: string }> {
  if (value === undefined || value === '' || value === DELIMITERS.NULL) {
    return [];
  }

  const attendees: Array<{ email: string; name: string }> = [];
  const items = value.split(',').filter((s) => s.length > 0);

  for (const item of items) {
    const parts = item.split('|');
    if (parts[0] !== undefined) {
      attendees.push({
        email: parts[0].trim(),
        name: parts[1]?.trim() ?? '',
      });
    }
  }

  return attendees;
}

// =============================================================================
// Type-Specific Parsers
// =============================================================================

/**
 * Parses folder output from AppleScript.
 */
export function parseFolders(output: string): AppleScriptFolderRow[] {
  const records = parseRawOutput(output);
  return records.map((r) => ({
    id: parseNumber(r['id']),
    name: parseString(r['name']),
    unreadCount: parseNumber(r['unreadCount']),
  }));
}

/**
 * Parses email list output from AppleScript.
 */
export function parseEmails(output: string): AppleScriptEmailRow[] {
  const records = parseRawOutput(output);
  return records.map((r) => ({
    id: parseNumber(r['id']),
    folderId: parseNumberOrNull(r['folderId']),
    subject: parseString(r['subject']),
    senderName: parseString(r['senderName']),
    senderEmail: parseString(r['senderEmail']),
    toRecipients: parseString(r['toRecipients']),
    ccRecipients: parseString(r['ccRecipients']),
    preview: parseString(r['preview']),
    isRead: parseBoolean(r['isRead']),
    dateReceived: parseString(r['dateReceived']),
    dateSent: parseString(r['dateSent']),
    priority: r['priority'] ?? 'normal',
    htmlContent: parseString(r['htmlContent']),
    plainContent: parseString(r['plainContent']),
    hasHtml: parseBoolean(r['hasHtml']),
    attachments: parseList(r['attachments']),
  }));
}

/**
 * Parses a single email output from AppleScript.
 */
export function parseEmail(output: string): AppleScriptEmailRow | null {
  const emails = parseEmails(output);
  return emails[0] ?? null;
}

/**
 * Parses calendar list output from AppleScript.
 */
export function parseCalendars(output: string): AppleScriptCalendarRow[] {
  const records = parseRawOutput(output);
  return records.map((r) => ({
    id: parseNumber(r['id']),
    name: parseString(r['name']),
  }));
}

/**
 * Parses event list output from AppleScript.
 */
export function parseEvents(output: string): AppleScriptEventRow[] {
  const records = parseRawOutput(output);
  return records.map((r) => ({
    id: parseNumber(r['id']),
    calendarId: parseNumberOrNull(r['calendarId']),
    subject: parseString(r['subject']),
    startTime: parseString(r['startTime']),
    endTime: parseString(r['endTime']),
    location: parseString(r['location']),
    isAllDay: parseBoolean(r['isAllDay']),
    isRecurring: parseBoolean(r['isRecurring']),
    organizer: parseString(r['organizer']),
    htmlContent: parseString(r['htmlContent']),
    plainContent: parseString(r['plainContent']),
    attendees: parseAttendees(r['attendees']),
  }));
}

/**
 * Parses a single event output from AppleScript.
 */
export function parseEvent(output: string): AppleScriptEventRow | null {
  const events = parseEvents(output);
  return events[0] ?? null;
}

/**
 * Parses contact list output from AppleScript.
 */
export function parseContacts(output: string): AppleScriptContactRow[] {
  const records = parseRawOutput(output);
  return records.map((r) => ({
    id: parseNumber(r['id']),
    displayName: parseString(r['displayName']),
    firstName: parseString(r['firstName']),
    lastName: parseString(r['lastName']),
    middleName: parseString(r['middleName']),
    nickname: parseString(r['nickname']),
    company: parseString(r['company']),
    jobTitle: parseString(r['jobTitle']),
    department: parseString(r['department']),
    notes: parseString(r['notes']),
    emails: parseList(r['emails']),
    homePhone: parseString(r['homePhone']),
    workPhone: parseString(r['workPhone']),
    mobilePhone: parseString(r['mobilePhone']),
    homeStreet: parseString(r['homeStreet']),
    homeCity: parseString(r['homeCity']),
    homeState: parseString(r['homeState']),
    homeZip: parseString(r['homeZip']),
    homeCountry: parseString(r['homeCountry']),
  }));
}

/**
 * Parses a single contact output from AppleScript.
 */
export function parseContact(output: string): AppleScriptContactRow | null {
  const contacts = parseContacts(output);
  return contacts[0] ?? null;
}

/**
 * Parses task list output from AppleScript.
 */
export function parseTasks(output: string): AppleScriptTaskRow[] {
  const records = parseRawOutput(output);
  return records.map((r) => ({
    id: parseNumber(r['id']),
    folderId: parseNumberOrNull(r['folderId']),
    name: parseString(r['name']),
    isCompleted: parseBoolean(r['isCompleted']),
    dueDate: parseString(r['dueDate']),
    startDate: parseString(r['startDate']),
    completedDate: parseString(r['completedDate']),
    priority: r['priority'] ?? 'normal',
    htmlContent: parseString(r['htmlContent']),
    plainContent: parseString(r['plainContent']),
  }));
}

/**
 * Parses a single task output from AppleScript.
 */
export function parseTask(output: string): AppleScriptTaskRow | null {
  const tasks = parseTasks(output);
  return tasks[0] ?? null;
}

/**
 * Parses note list output from AppleScript.
 */
export function parseNotes(output: string): AppleScriptNoteRow[] {
  const records = parseRawOutput(output);
  return records.map((r) => ({
    id: parseNumber(r['id']),
    folderId: parseNumberOrNull(r['folderId']),
    name: parseString(r['name']),
    createdDate: parseString(r['createdDate']),
    modifiedDate: parseString(r['modifiedDate']),
    preview: parseString(r['preview']),
    htmlContent: parseString(r['htmlContent']),
    plainContent: parseString(r['plainContent']),
  }));
}

/**
 * Parses a single note output from AppleScript.
 */
export function parseNote(output: string): AppleScriptNoteRow | null {
  const notes = parseNotes(output);
  return notes[0] ?? null;
}

/**
 * Parses a simple count value from AppleScript output.
 */
export function parseCount(output: string): number {
  const trimmed = output.trim();
  const num = parseInt(trimmed, 10);
  return isNaN(num) ? 0 : num;
}

/**
 * Parses account list output from AppleScript.
 */
export function parseAccounts(output: string): AppleScriptAccountRow[] {
  const records = parseRawOutput(output);
  return records.map((r) => ({
    id: parseNumber(r['id']),
    name: parseString(r['name']),
    email: parseString(r['email']),
    type: r['type'] ?? 'exchange',
  }));
}

/**
 * Parses default account output from AppleScript.
 * Returns the account ID or null if not found.
 */
export function parseDefaultAccountId(output: string): number | null {
  const trimmed = output.trim();
  if (trimmed.startsWith('error')) {
    return null;
  }
  const parts = trimmed.split(DELIMITERS.EQUALS);
  if (parts[0] === 'id' && parts[1] !== undefined) {
    return parseNumberOrNull(parts[1]);
  }
  return null;
}

/**
 * Parses the result of a create event AppleScript.
 * Returns the created event's ID and calendar ID.
 */
export function parseCreateEventResult(output: string): { id: number; calendarId: number | null } | null {
  const records = parseRawOutput(output);
  if (records.length === 0) return null;
  const r = records[0]!;
  const id = parseNumber(r['id']);
  if (id === 0) return null; // 0 indicates malformed output, not a real event ID
  return {
    id,
    calendarId: parseNumberOrNull(r['calendarId']),
  };
}

/**
 * Parses folders with account information from AppleScript.
 */
export function parseFoldersWithAccount(output: string): AppleScriptFolderWithAccountRow[] {
  const records = parseRawOutput(output);
  return records.map((r) => ({
    id: parseNumber(r['id']),
    name: parseString(r['name']),
    unreadCount: parseNumber(r['unreadCount']),
    messageCount: parseNumber(r['messageCount']),
    accountId: parseNumber(r['accountId']),
  }));
}

/**
 * Parses the result of a respond-to-event operation.
 */
export function parseRespondToEventResult(output: string): RespondToEventResult | null {
  const records = parseRawOutput(output);
  if (records.length === 0) return null;

  const record = records[0];
  if (!record) return null;

  const success = record['success'] === 'true';

  if (success) {
    return {
      success: true,
      eventId: parseNumber(record['eventId']),
    };
  } else {
    return {
      success: false,
      error: record['error'] ?? 'Unknown error',
    };
  }
}

/**
 * Parses the result of a delete-event operation.
 */
export function parseDeleteEventResult(output: string): DeleteEventResult | null {
  const records = parseRawOutput(output);
  if (records.length === 0) return null;

  const record = records[0];
  if (!record) return null;

  const success = record['success'] === 'true';

  if (success) {
    return {
      success: true,
      eventId: parseNumber(record['eventId']),
    };
  } else {
    return {
      success: false,
      error: record['error'] ?? 'Unknown error',
    };
  }
}
