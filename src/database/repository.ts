/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Row and repository shapes shared across the Graph tool layer.
 *
 * These are types only. The SQLite-backed implementation that once lived here
 * belonged to the AppleScript/local-Outlook backend removed in the Graph-only
 * change, and its runtime half was deleted with the better-sqlite3 dependency
 * (#108) — it had been unreachable since, constructed by nothing. Every
 * importer of this module uses `import type`.
 */


// =============================================================================
// Row Types (raw database rows)
// =============================================================================

export interface FolderRow {
  // Durable self-encoding string token (`fd_…`) carrying the immutable Graph
  // folder/calendar id (U5). Graph-only — string, not a legacy numeric union.
  readonly id: string;
  readonly name: string | null;
  readonly parentId: string | null;
  readonly specialType: number;
  readonly folderType: number;
  readonly accountId: number;
  readonly messageCount: number;
  readonly unreadCount: number;
}

export interface EmailRow {
  // Durable string token (`em_…`) on the Graph backend (U5); numeric on the
  // AppleScript/SQLite backend (D4).
  readonly id: string;
  // Durable self-encoding `fd_…` token (U5), Graph-only — string.
  readonly folderId: string;
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
  readonly conversationId: string | null;
  readonly dataFilePath: string | null;
}

export interface EventRow {
  // Durable string token (`ev_…`) on the Graph backend (U5); numeric on the
  // AppleScript/SQLite backend (D4).
  readonly id: string;
  // Durable self-encoding `fd_…` token (U5), Graph-only — string.
  readonly folderId: string;
  readonly subject: string | null;
  readonly startDate: number | null;
  readonly endDate: number | null;
  readonly isRecurring: number;
  readonly hasReminder: number;
  readonly attendeeCount: number;
  readonly uid: string | null;
  readonly masterRecordId: number | null;
  readonly recurrenceId: number | null;
  readonly dataFilePath: string | null;
  readonly onlineMeetingUrl: string | null;
}

export interface ContactRow {
  // Durable string token (`ct_…`) on the Graph backend (U5); numeric on the
  // AppleScript/SQLite backend (D4).
  readonly id: string;
  readonly folderId: number;
  readonly displayName: string | null;
  readonly sortName: string | null;
  readonly contactType: number | null;
  readonly dataFilePath: string | null;
}

export interface TaskRow {
  // Durable composite `td_…` token on the Graph backend (U5); numeric on the
  // AppleScript/SQLite backend (D4).
  readonly id: string;
  // Durable alias-backed `tl_…` token on the Graph backend (U5); numeric on the
  // AppleScript/SQLite backend (D4).
  readonly folderId: string;
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
  getFolder(id: string): FolderRow | undefined;

  // Emails
  listEmails(folderId: number, limit: number, offset: number): EmailRow[];
  listUnreadEmails(folderId: number, limit: number, offset: number): EmailRow[];
  searchEmails(query: string, limit: number): EmailRow[];
  searchEmailsInFolder(folderId: number, query: string, limit: number): EmailRow[];
  getEmail(id: string): EmailRow | undefined;
  getUnreadCount(): number;
  getUnreadCountByFolder(folderId: number): number;

  // Calendar
  listCalendars(): FolderRow[];
  listEvents(limit: number): EventRow[];
  listEventsByFolder(folderId: number, limit: number): EventRow[];
  listEventsByDateRange(startDate: number, endDate: number, limit: number): EventRow[];
  searchEvents(query: string | null, startDate: string | null, endDate: string | null, limit: number): EventRow[];
  getEvent(id: string): EventRow | undefined;

  // Contacts
  listContacts(limit: number, offset: number): ContactRow[];
  searchContacts(query: string, limit: number): ContactRow[];
  getContact(id: string): ContactRow | undefined;

  // Tasks
  listTasks(limit: number, offset: number): TaskRow[];
  listIncompleteTasks(limit: number, offset: number): TaskRow[];
  searchTasks(query: string, limit: number): TaskRow[];
  getTask(id: string): TaskRow | undefined;

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
  moveEmail(emailId: string, destinationFolderId: string): void;
  deleteEmail(emailId: string): void;
  archiveEmail(emailId: string): void;
  junkEmail(emailId: string): void;
  markEmailRead(emailId: string, isRead: boolean): void;
  setEmailFlag(emailId: string, flagStatus: number): void;
  setEmailCategories(emailId: string, categories: string[]): void;
  setEmailImportance(emailId: string, importance: string): void;

  // Folder management
  createFolder(name: string, parentFolderId?: string): FolderRow;
  deleteFolder(folderId: string): void;
  renameFolder(folderId: string, newName: string): void;
  moveFolder(folderId: string, destinationParentId: string): void;
  emptyFolder(folderId: string): void;
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
  getEmail(id: string): MaybePromise<EmailRow | undefined>;
  getFolder(id: string): MaybePromise<FolderRow | undefined>;

  // Email organization
  moveEmail(emailId: string, destinationFolderId: string): MaybePromise<void>;
  deleteEmail(emailId: string): MaybePromise<void>;
  archiveEmail(emailId: string): MaybePromise<void>;
  junkEmail(emailId: string): MaybePromise<void>;
  markEmailRead(emailId: string, isRead: boolean): MaybePromise<void>;
  setEmailFlag(emailId: string, flagStatus: number): MaybePromise<void>;
  setEmailCategories(emailId: string, categories: string[]): MaybePromise<void>;
  setEmailImportance(emailId: string, importance: string): MaybePromise<void>;

  // Folder management
  createFolder(name: string, parentFolderId?: string): MaybePromise<FolderRow>;
  deleteFolder(folderId: string): MaybePromise<void>;
  renameFolder(folderId: string, newName: string): MaybePromise<void>;
  moveFolder(folderId: string, destinationParentId: string): MaybePromise<void>;
  emptyFolder(folderId: string): MaybePromise<void>;
}

// =============================================================================
// Repository Implementation
// =============================================================================

/**
 * Repository implementation using better-sqlite3.
 */
