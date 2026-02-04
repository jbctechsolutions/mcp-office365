/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Repository for accessing Outlook data.
 *
 * Provides a data access layer over the SQLite database.
 */

import type { IConnection } from './connection.js';
import * as queries from './queries.js';

// =============================================================================
// Row Types (raw database rows)
// =============================================================================

export interface FolderRow {
  readonly id: number;
  readonly name: string | null;
  readonly parentId: number | null;
  readonly specialType: number;
  readonly folderType: number;
  readonly accountId: number;
  readonly messageCount: number;
  readonly unreadCount: number;
}

export interface EmailRow {
  readonly id: number;
  readonly folderId: number;
  readonly subject: string | null;
  readonly sender: string | null;
  readonly senderAddress: string | null;
  readonly recipients: string | null;
  readonly displayTo: string | null;
  readonly toAddresses: string | null;
  readonly ccAddresses: string | null;
  readonly preview: string | null;
  readonly isRead: number;
  readonly timeReceived: number | null;
  readonly timeSent: number | null;
  readonly hasAttachment: number;
  readonly size: number;
  readonly priority: number;
  readonly flagStatus: number;
  readonly categories: Buffer | null;
  readonly messageId: string | null;
  readonly conversationId: number | null;
  readonly dataFilePath: string | null;
}

export interface EventRow {
  readonly id: number;
  readonly folderId: number;
  readonly startDate: number | null;
  readonly endDate: number | null;
  readonly isRecurring: number;
  readonly hasReminder: number;
  readonly attendeeCount: number;
  readonly uid: string | null;
  readonly masterRecordId: number | null;
  readonly recurrenceId: number | null;
  readonly dataFilePath: string | null;
}

export interface ContactRow {
  readonly id: number;
  readonly folderId: number;
  readonly displayName: string | null;
  readonly sortName: string | null;
  readonly contactType: number | null;
  readonly dataFilePath: string | null;
}

export interface TaskRow {
  readonly id: number;
  readonly folderId: number;
  readonly name: string | null;
  readonly isCompleted: number;
  readonly dueDate: number | null;
  readonly startDate: number | null;
  readonly priority: number;
  readonly hasReminder: number | null;
  readonly dataFilePath: string | null;
}

export interface NoteRow {
  readonly id: number;
  readonly folderId: number;
  readonly modifiedDate: number | null;
  readonly dataFilePath: string | null;
}

export interface CountRow {
  readonly count: number;
}

// =============================================================================
// Repository Interface
// =============================================================================

/**
 * Interface for the Outlook data repository (for dependency injection).
 */
export interface IRepository {
  // Folders
  listFolders(): FolderRow[];
  getFolder(id: number): FolderRow | undefined;

  // Emails
  listEmails(folderId: number, limit: number, offset: number): EmailRow[];
  listUnreadEmails(folderId: number, limit: number, offset: number): EmailRow[];
  searchEmails(query: string, limit: number): EmailRow[];
  searchEmailsInFolder(folderId: number, query: string, limit: number): EmailRow[];
  getEmail(id: number): EmailRow | undefined;
  getUnreadCount(): number;
  getUnreadCountByFolder(folderId: number): number;

  // Calendar
  listCalendars(): FolderRow[];
  listEvents(limit: number): EventRow[];
  listEventsByFolder(folderId: number, limit: number): EventRow[];
  listEventsByDateRange(startDate: number, endDate: number, limit: number): EventRow[];
  getEvent(id: number): EventRow | undefined;

  // Contacts
  listContacts(limit: number, offset: number): ContactRow[];
  searchContacts(query: string, limit: number): ContactRow[];
  getContact(id: number): ContactRow | undefined;

  // Tasks
  listTasks(limit: number, offset: number): TaskRow[];
  listIncompleteTasks(limit: number, offset: number): TaskRow[];
  searchTasks(query: string, limit: number): TaskRow[];
  getTask(id: number): TaskRow | undefined;

  // Notes
  listNotes(limit: number, offset: number): NoteRow[];
  getNote(id: number): NoteRow | undefined;
}

// =============================================================================
// Writeable Repository Interface
// =============================================================================

/**
 * Interface for writable Outlook data operations.
 * Extends IRepository with mutation methods for mailbox organization.
 */
export interface IWriteableRepository extends IRepository {
  // Email organization
  moveEmail(emailId: number, destinationFolderId: number): void;
  deleteEmail(emailId: number): void;
  archiveEmail(emailId: number): void;
  junkEmail(emailId: number): void;
  markEmailRead(emailId: number, isRead: boolean): void;
  setEmailFlag(emailId: number, flagStatus: number): void;
  setEmailCategories(emailId: number, categories: string[]): void;

  // Folder management
  createFolder(name: string, parentFolderId?: number): FolderRow;
  deleteFolder(folderId: number): void;
  renameFolder(folderId: number, newName: string): void;
  moveFolder(folderId: number, destinationParentId: number): void;
  emptyFolder(folderId: number): void;
}

// =============================================================================
// Async-Compatible Repository Interface
// =============================================================================

/**
 * A value that may be synchronous or wrapped in a Promise.
 */
export type MaybePromise<T> = T | Promise<T>;

/**
 * Async-compatible repository interface for mailbox organization tools.
 *
 * Both sync (AppleScript) and async (Graph) repositories satisfy this
 * interface. AppleScript repos return plain values; Graph repos return
 * Promises. MailboxOrganizationTools awaits all calls uniformly.
 */
export interface IMailboxRepository {
  // Read
  getEmail(id: number): MaybePromise<EmailRow | undefined>;
  getFolder(id: number): MaybePromise<FolderRow | undefined>;

  // Email organization
  moveEmail(emailId: number, destinationFolderId: number): MaybePromise<void>;
  deleteEmail(emailId: number): MaybePromise<void>;
  archiveEmail(emailId: number): MaybePromise<void>;
  junkEmail(emailId: number): MaybePromise<void>;
  markEmailRead(emailId: number, isRead: boolean): MaybePromise<void>;
  setEmailFlag(emailId: number, flagStatus: number): MaybePromise<void>;
  setEmailCategories(emailId: number, categories: string[]): MaybePromise<void>;

  // Folder management
  createFolder(name: string, parentFolderId?: number): MaybePromise<FolderRow>;
  deleteFolder(folderId: number): MaybePromise<void>;
  renameFolder(folderId: number, newName: string): MaybePromise<void>;
  moveFolder(folderId: number, destinationParentId: number): MaybePromise<void>;
  emptyFolder(folderId: number): MaybePromise<void>;
}

// =============================================================================
// Repository Implementation
// =============================================================================

/**
 * Repository implementation using better-sqlite3.
 */
export class OutlookRepository implements IRepository {
  constructor(private readonly connection: IConnection) {}

  // ---------------------------------------------------------------------------
  // Folders
  // ---------------------------------------------------------------------------

  listFolders(): FolderRow[] {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.LIST_FOLDERS);
      return stmt.all() as FolderRow[];
    });
  }

  getFolder(id: number): FolderRow | undefined {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.GET_FOLDER);
      return stmt.get(id) as FolderRow | undefined;
    });
  }

  // ---------------------------------------------------------------------------
  // Emails
  // ---------------------------------------------------------------------------

  listEmails(folderId: number, limit: number, offset: number): EmailRow[] {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.LIST_EMAILS);
      return stmt.all(folderId, limit, offset) as EmailRow[];
    });
  }

  listUnreadEmails(folderId: number, limit: number, offset: number): EmailRow[] {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.LIST_UNREAD_EMAILS);
      return stmt.all(folderId, limit, offset) as EmailRow[];
    });
  }

  searchEmails(query: string, limit: number): EmailRow[] {
    return this.connection.execute((db) => {
      const pattern = `%${query}%`;
      const stmt = db.prepare(queries.SEARCH_EMAILS);
      return stmt.all(pattern, pattern, pattern, limit) as EmailRow[];
    });
  }

  searchEmailsInFolder(folderId: number, query: string, limit: number): EmailRow[] {
    return this.connection.execute((db) => {
      const pattern = `%${query}%`;
      const stmt = db.prepare(queries.SEARCH_EMAILS_IN_FOLDER);
      return stmt.all(folderId, pattern, pattern, pattern, limit) as EmailRow[];
    });
  }

  getEmail(id: number): EmailRow | undefined {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.GET_EMAIL);
      return stmt.get(id) as EmailRow | undefined;
    });
  }

  getUnreadCount(): number {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.GET_UNREAD_COUNT);
      const row = stmt.get() as CountRow | undefined;
      return row?.count ?? 0;
    });
  }

  getUnreadCountByFolder(folderId: number): number {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.GET_UNREAD_COUNT_BY_FOLDER);
      const row = stmt.get(folderId) as CountRow | undefined;
      return row?.count ?? 0;
    });
  }

  // ---------------------------------------------------------------------------
  // Calendar
  // ---------------------------------------------------------------------------

  listCalendars(): FolderRow[] {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.LIST_CALENDARS);
      return stmt.all() as FolderRow[];
    });
  }

  listEvents(limit: number): EventRow[] {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.LIST_EVENTS);
      return stmt.all(limit) as EventRow[];
    });
  }

  listEventsByFolder(folderId: number, limit: number): EventRow[] {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.LIST_EVENTS_BY_FOLDER);
      return stmt.all(folderId, limit) as EventRow[];
    });
  }

  listEventsByDateRange(startDate: number, endDate: number, limit: number): EventRow[] {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.LIST_EVENTS_BY_DATE_RANGE);
      return stmt.all(startDate, endDate, limit) as EventRow[];
    });
  }

  getEvent(id: number): EventRow | undefined {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.GET_EVENT);
      return stmt.get(id) as EventRow | undefined;
    });
  }

  // ---------------------------------------------------------------------------
  // Contacts
  // ---------------------------------------------------------------------------

  listContacts(limit: number, offset: number): ContactRow[] {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.LIST_CONTACTS);
      return stmt.all(limit, offset) as ContactRow[];
    });
  }

  searchContacts(query: string, limit: number): ContactRow[] {
    return this.connection.execute((db) => {
      const pattern = `%${query}%`;
      const stmt = db.prepare(queries.SEARCH_CONTACTS);
      return stmt.all(pattern, pattern, limit) as ContactRow[];
    });
  }

  getContact(id: number): ContactRow | undefined {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.GET_CONTACT);
      return stmt.get(id) as ContactRow | undefined;
    });
  }

  // ---------------------------------------------------------------------------
  // Tasks
  // ---------------------------------------------------------------------------

  listTasks(limit: number, offset: number): TaskRow[] {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.LIST_TASKS);
      return stmt.all(limit, offset) as TaskRow[];
    });
  }

  listIncompleteTasks(limit: number, offset: number): TaskRow[] {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.LIST_INCOMPLETE_TASKS);
      return stmt.all(limit, offset) as TaskRow[];
    });
  }

  searchTasks(query: string, limit: number): TaskRow[] {
    return this.connection.execute((db) => {
      const pattern = `%${query}%`;
      const stmt = db.prepare(queries.SEARCH_TASKS);
      return stmt.all(pattern, limit) as TaskRow[];
    });
  }

  getTask(id: number): TaskRow | undefined {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.GET_TASK);
      return stmt.get(id) as TaskRow | undefined;
    });
  }

  // ---------------------------------------------------------------------------
  // Notes
  // ---------------------------------------------------------------------------

  listNotes(limit: number, offset: number): NoteRow[] {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.LIST_NOTES);
      return stmt.all(limit, offset) as NoteRow[];
    });
  }

  getNote(id: number): NoteRow | undefined {
    return this.connection.execute((db) => {
      const stmt = db.prepare(queries.GET_NOTE);
      return stmt.get(id) as NoteRow | undefined;
    });
  }
}

/**
 * Creates a repository with the given connection.
 */
export function createRepository(connection: IConnection): IRepository {
  return new OutlookRepository(connection);
}
