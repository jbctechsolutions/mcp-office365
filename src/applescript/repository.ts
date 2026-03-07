/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * AppleScript-based repository implementation.
 *
 * Implements IRepository using AppleScript to communicate with Outlook,
 * enabling support for both classic and new Outlook for Mac.
 */

import type {
  IWriteableRepository,
  FolderRow,
  EmailRow,
  EventRow,
  ContactRow,
  TaskRow,
  NoteRow,
} from '../database/repository.js';
import { executeAppleScriptOrThrow } from './executor.js';
import * as scripts from './scripts.js';
import * as parser from './parser.js';
import { isoToAppleTimestamp } from '../utils/dates.js';
import {
  createEmailPath,
  createEventPath,
  createContactPath,
  createTaskPath,
  createNotePath,
} from './content-readers.js';

// =============================================================================
// Utility Functions
// =============================================================================

/**
 * Converts a priority string to a numeric value.
 */
function priorityToNumber(priority: string): number {
  switch (priority.toLowerCase()) {
    case 'high':
      return 1;
    case 'low':
      return -1;
    default:
      return 0;
  }
}

// =============================================================================
// Date Conversion
// =============================================================================

/**
 * Parses an ISO 8601 date string into individual UTC components
 * for locale-safe AppleScript date construction.
 */
function isoToDateComponents(isoString: string): {
  year: number;
  month: number;
  day: number;
  hours: number;
  minutes: number;
} {
  const date = new Date(isoString);
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    hours: date.getUTCHours(),
    minutes: date.getUTCMinutes(),
  };
}

// =============================================================================
// Row Converters
// =============================================================================

/**
 * Converts AppleScript folder output to FolderRow.
 */
function toFolderRow(asFolder: parser.AppleScriptFolderRow): FolderRow {
  return {
    id: asFolder.id,
    name: asFolder.name,
    parentId: null, // AppleScript doesn't provide parent info easily
    specialType: 0,
    folderType: 1, // Mail folder
    accountId: 1, // Default account
    messageCount: 0, // Not available via AppleScript directly
    unreadCount: asFolder.unreadCount,
  };
}

/**
 * Converts AppleScript calendar output to FolderRow (calendars use FolderRow).
 */
function calendarToFolderRow(asCal: parser.AppleScriptCalendarRow): FolderRow {
  return {
    id: asCal.id,
    name: asCal.name,
    parentId: null,
    specialType: 0,
    folderType: 2, // Calendar folder
    accountId: 1,
    messageCount: 0,
    unreadCount: 0,
  };
}

/**
 * Converts AppleScript email output to EmailRow.
 */
function toEmailRow(asEmail: parser.AppleScriptEmailRow): EmailRow {
  return {
    id: asEmail.id,
    folderId: asEmail.folderId ?? 0,
    subject: asEmail.subject,
    sender: asEmail.senderName,
    senderAddress: asEmail.senderEmail,
    recipients: asEmail.toRecipients,
    displayTo: asEmail.toRecipients,
    toAddresses: asEmail.toRecipients,
    ccAddresses: asEmail.ccRecipients,
    preview: asEmail.preview,
    isRead: asEmail.isRead ? 1 : 0,
    timeReceived: isoToAppleTimestamp(asEmail.dateReceived),
    timeSent: isoToAppleTimestamp(asEmail.dateSent),
    hasAttachment: asEmail.attachments.length > 0 ? 1 : 0,
    size: 0,
    priority: priorityToNumber(asEmail.priority),
    flagStatus: 0,
    categories: null,
    messageId: null,
    conversationId: null,
    dataFilePath: createEmailPath(asEmail.id), // Special path for AppleScript content reader
  };
}

/**
 * Converts AppleScript event output to EventRow.
 */
function toEventRow(asEvent: parser.AppleScriptEventRow): EventRow {
  return {
    id: asEvent.id,
    folderId: asEvent.calendarId ?? 0,
    subject: asEvent.subject ?? null,
    startDate: isoToAppleTimestamp(asEvent.startTime),
    endDate: isoToAppleTimestamp(asEvent.endTime),
    isRecurring: asEvent.isRecurring ? 1 : 0,
    hasReminder: 0,
    attendeeCount: asEvent.attendees.length,
    uid: null,
    masterRecordId: null,
    recurrenceId: null,
    dataFilePath: createEventPath(asEvent.id),
  };
}

/**
 * Converts AppleScript contact output to ContactRow.
 */
function toContactRow(asContact: parser.AppleScriptContactRow): ContactRow {
  return {
    id: asContact.id,
    folderId: 0,
    displayName: asContact.displayName,
    sortName: asContact.lastName ?? asContact.displayName,
    contactType: null,
    dataFilePath: createContactPath(asContact.id),
  };
}

/**
 * Converts AppleScript task output to TaskRow.
 */
function toTaskRow(asTask: parser.AppleScriptTaskRow): TaskRow {
  return {
    id: asTask.id,
    folderId: asTask.folderId ?? 0,
    name: asTask.name,
    isCompleted: asTask.isCompleted ? 1 : 0,
    dueDate: isoToAppleTimestamp(asTask.dueDate),
    startDate: isoToAppleTimestamp(asTask.startDate),
    priority: priorityToNumber(asTask.priority),
    hasReminder: null,
    dataFilePath: createTaskPath(asTask.id),
  };
}

/**
 * Converts AppleScript note output to NoteRow.
 */
function toNoteRow(asNote: parser.AppleScriptNoteRow): NoteRow {
  return {
    id: asNote.id,
    folderId: asNote.folderId ?? 0,
    modifiedDate: isoToAppleTimestamp(asNote.modifiedDate),
    dataFilePath: createNotePath(asNote.id),
  };
}

// =============================================================================
// Repository Implementation
// =============================================================================

/**
 * Repository implementation using AppleScript.
 *
 * Communicates with Microsoft Outlook via osascript to fetch data.
 * Works with both classic and new Outlook for Mac.
 */
export class AppleScriptRepository implements IWriteableRepository {
  // Cache for folder ID lookup
  private readonly folderCache: Map<number, FolderRow> = new Map();
  private folderCacheExpiry: number = 0;
  private readonly CACHE_TTL_MS = 30000; // 30 seconds

  // ---------------------------------------------------------------------------
  // Folders
  // ---------------------------------------------------------------------------

  listFolders(): FolderRow[] {
    const output = executeAppleScriptOrThrow(scripts.LIST_MAIL_FOLDERS);
    const folders = parser.parseFolders(output).map(toFolderRow);

    // Update cache
    this.folderCache.clear();
    for (const folder of folders) {
      this.folderCache.set(folder.id, folder);
    }
    this.folderCacheExpiry = Date.now() + this.CACHE_TTL_MS;

    return folders;
  }

  getFolder(id: number): FolderRow | undefined {
    // Check cache first
    if (Date.now() < this.folderCacheExpiry) {
      const cached = this.folderCache.get(id);
      if (cached != null) {
        return cached;
      }
    }

    // Refresh cache
    const folders = this.listFolders();
    return folders.find((f) => f.id === id);
  }

  // ---------------------------------------------------------------------------
  // Emails
  // ---------------------------------------------------------------------------

  listEmails(folderId: number, limit: number, offset: number): EmailRow[] {
    const script = scripts.listMessages(folderId, limit, offset, false);
    const output = executeAppleScriptOrThrow(script);
    return parser.parseEmails(output).map(toEmailRow);
  }

  listUnreadEmails(folderId: number, limit: number, offset: number): EmailRow[] {
    const script = scripts.listMessages(folderId, limit, offset, true);
    const output = executeAppleScriptOrThrow(script);
    return parser.parseEmails(output).map(toEmailRow);
  }

  searchEmails(query: string, limit: number): EmailRow[] {
    const script = scripts.searchMessages(query, null, limit);
    const output = executeAppleScriptOrThrow(script);
    return parser.parseEmails(output).map(toEmailRow);
  }

  searchEmailsInFolder(folderId: number, query: string, limit: number): EmailRow[] {
    const script = scripts.searchMessages(query, folderId, limit);
    const output = executeAppleScriptOrThrow(script);
    return parser.parseEmails(output).map(toEmailRow);
  }

  getEmail(id: number): EmailRow | undefined {
    try {
      const script = scripts.getMessage(id);
      const output = executeAppleScriptOrThrow(script);
      const email = parser.parseEmail(output);
      return email != null ? toEmailRow(email) : undefined;
    } catch {
      return undefined;
    }
  }

  getUnreadCount(): number {
    // Sum unread counts from all folders
    const folders = this.listFolders();
    return folders.reduce((sum, f) => sum + f.unreadCount, 0);
  }

  getUnreadCountByFolder(folderId: number): number {
    try {
      const script = scripts.getUnreadCount(folderId);
      const output = executeAppleScriptOrThrow(script);
      return parser.parseCount(output);
    } catch {
      return 0;
    }
  }

  // ---------------------------------------------------------------------------
  // Calendar
  // ---------------------------------------------------------------------------

  listCalendars(): FolderRow[] {
    const output = executeAppleScriptOrThrow(scripts.LIST_CALENDARS);
    return parser.parseCalendars(output).map(calendarToFolderRow);
  }

  listEvents(limit: number): EventRow[] {
    const script = scripts.listEvents(null, null, null, limit);
    const output = executeAppleScriptOrThrow(script);
    return parser.parseEvents(output).map(toEventRow);
  }

  listEventsByFolder(folderId: number, limit: number): EventRow[] {
    const script = scripts.listEvents(folderId, null, null, limit);
    const output = executeAppleScriptOrThrow(script);
    return parser.parseEvents(output).map(toEventRow);
  }

  listEventsByDateRange(startDate: number, endDate: number, limit: number): EventRow[] {
    // Convert Apple timestamps to ISO for filtering
    // Note: AppleScript filtering by date is done client-side for simplicity
    const allEvents = this.listEvents(1000); // Get more events for filtering
    return allEvents
      .filter((e) => {
        if (e.startDate == null) return false;
        return e.startDate >= startDate && e.startDate <= endDate;
      })
      .slice(0, limit);
  }

  searchEvents(query: string | null, startDate: string | null, endDate: string | null, limit: number): EventRow[] {
    const startComponents = startDate != null ? isoToDateComponents(startDate) : null;
    const endComponents = endDate != null ? isoToDateComponents(endDate) : null;
    const script = scripts.searchEvents(query, startComponents, endComponents, limit);
    const output = executeAppleScriptOrThrow(script);
    return parser.parseEvents(output).map(toEventRow);
  }

  getEvent(id: number): EventRow | undefined {
    try {
      const script = scripts.getEvent(id);
      const output = executeAppleScriptOrThrow(script);
      const event = parser.parseEvent(output);
      return event != null ? toEventRow(event) : undefined;
    } catch {
      return undefined;
    }
  }

  // ---------------------------------------------------------------------------
  // Contacts
  // ---------------------------------------------------------------------------

  listContacts(limit: number, offset: number): ContactRow[] {
    const script = scripts.listContacts(limit, offset);
    const output = executeAppleScriptOrThrow(script);
    return parser.parseContacts(output).map(toContactRow);
  }

  searchContacts(query: string, limit: number): ContactRow[] {
    const script = scripts.searchContacts(query, limit);
    const output = executeAppleScriptOrThrow(script);
    return parser.parseContacts(output).map(toContactRow);
  }

  getContact(id: number): ContactRow | undefined {
    try {
      const script = scripts.getContact(id);
      const output = executeAppleScriptOrThrow(script);
      const contact = parser.parseContact(output);
      return contact != null ? toContactRow(contact) : undefined;
    } catch {
      return undefined;
    }
  }

  // ---------------------------------------------------------------------------
  // Tasks
  // ---------------------------------------------------------------------------

  listTasks(limit: number, offset: number): TaskRow[] {
    const script = scripts.listTasks(limit, offset, true);
    const output = executeAppleScriptOrThrow(script);
    return parser.parseTasks(output).map(toTaskRow);
  }

  listIncompleteTasks(limit: number, offset: number): TaskRow[] {
    const script = scripts.listTasks(limit, offset, false);
    const output = executeAppleScriptOrThrow(script);
    return parser.parseTasks(output).map(toTaskRow);
  }

  searchTasks(query: string, limit: number): TaskRow[] {
    const script = scripts.searchTasks(query, limit);
    const output = executeAppleScriptOrThrow(script);
    return parser.parseTasks(output).map(toTaskRow);
  }

  getTask(id: number): TaskRow | undefined {
    try {
      const script = scripts.getTask(id);
      const output = executeAppleScriptOrThrow(script);
      const task = parser.parseTask(output);
      return task != null ? toTaskRow(task) : undefined;
    } catch {
      return undefined;
    }
  }

  // ---------------------------------------------------------------------------
  // Notes
  // ---------------------------------------------------------------------------

  listNotes(limit: number, offset: number): NoteRow[] {
    const script = scripts.listNotes(limit, offset);
    const output = executeAppleScriptOrThrow(script);
    return parser.parseNotes(output).map(toNoteRow);
  }

  getNote(id: number): NoteRow | undefined {
    try {
      const script = scripts.getNote(id);
      const output = executeAppleScriptOrThrow(script);
      const note = parser.parseNote(output);
      return note != null ? toNoteRow(note) : undefined;
    } catch {
      return undefined;
    }
  }

  // ---------------------------------------------------------------------------
  // Write Operations
  // ---------------------------------------------------------------------------

  moveEmail(emailId: number, destinationFolderId: number): void {
    const script = scripts.moveMessage(emailId, destinationFolderId);
    executeAppleScriptOrThrow(script);
  }

  deleteEmail(emailId: number): void {
    const script = scripts.deleteMessage(emailId);
    executeAppleScriptOrThrow(script);
  }

  archiveEmail(emailId: number): void {
    const script = scripts.archiveMessage(emailId);
    executeAppleScriptOrThrow(script);
  }

  junkEmail(emailId: number): void {
    const script = scripts.junkMessage(emailId);
    executeAppleScriptOrThrow(script);
  }

  markEmailRead(emailId: number, isRead: boolean): void {
    const script = scripts.setMessageReadStatus(emailId, isRead);
    executeAppleScriptOrThrow(script);
  }

  setEmailFlag(emailId: number, flagStatus: number): void {
    const script = scripts.setMessageFlag(emailId, flagStatus);
    executeAppleScriptOrThrow(script);
  }

  setEmailCategories(emailId: number, categories: string[]): void {
    const script = scripts.setMessageCategories(emailId, categories);
    executeAppleScriptOrThrow(script);
  }

  setEmailImportance(_emailId: number, _importance: string): void {
    throw new Error('setEmailImportance is only supported via Graph API');
  }

  createFolder(name: string, parentFolderId?: number): FolderRow {
    const script = scripts.createMailFolder(name, parentFolderId);
    const output = executeAppleScriptOrThrow(script);
    const newFolderId = parseInt(output.trim(), 10);

    return {
      id: newFolderId,
      name,
      parentId: parentFolderId ?? null,
      specialType: 0,
      folderType: 1,
      accountId: 1,
      messageCount: 0,
      unreadCount: 0,
    };
  }

  deleteFolder(folderId: number): void {
    const script = scripts.deleteMailFolder(folderId);
    executeAppleScriptOrThrow(script);
  }

  renameFolder(folderId: number, newName: string): void {
    const script = scripts.renameMailFolder(folderId, newName);
    executeAppleScriptOrThrow(script);
  }

  moveFolder(folderId: number, destinationParentId: number): void {
    const script = scripts.moveMailFolder(folderId, destinationParentId);
    executeAppleScriptOrThrow(script);
  }

  emptyFolder(folderId: number): void {
    const script = scripts.emptyMailFolder(folderId);
    executeAppleScriptOrThrow(script);
  }
}

/**
 * Creates an AppleScript-based repository.
 */
export function createAppleScriptRepository(): IWriteableRepository {
  return new AppleScriptRepository();
}
