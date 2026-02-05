/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Test database fixtures and helpers.
 */

import Database from 'better-sqlite3';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { randomBytes } from 'node:crypto';
import { unlinkSync, existsSync } from 'node:fs';

/**
 * Creates a temporary test database with sample data.
 */
export function createTestDatabase(): { path: string; cleanup: () => void } {
  const filename = `outlook-test-${randomBytes(8).toString('hex')}.sqlite`;
  const path = join(tmpdir(), filename);

  const db = new Database(path);

  // Create tables matching Outlook schema
  db.exec(`
    CREATE TABLE Folders (
      Record_RecordID INTEGER PRIMARY KEY AUTOINCREMENT,
      PathToDataFile TEXT NOT NULL,
      Record_ModDate DATETIME NOT NULL,
      Record_AccountUID INTEGER DEFAULT 0,
      Folder_ParentID INTEGER,
      Folder_FolderClass INTEGER NOT NULL DEFAULT 0,
      Folder_FolderType INTEGER NOT NULL,
      Folder_SpecialFolderType INTEGER DEFAULT 0,
      Folder_Name TEXT,
      Folder_FolderOrder INTEGER DEFAULT 0
    );

    CREATE TABLE Mail (
      Record_RecordID INTEGER PRIMARY KEY AUTOINCREMENT,
      PathToDataFile TEXT NOT NULL,
      Record_ModDate DATETIME NOT NULL,
      Record_FolderID INTEGER NOT NULL,
      Record_AccountUID INTEGER DEFAULT 0,
      Message_HasAttachment BOOLEAN DEFAULT 0,
      Message_MessageID TEXT,
      Message_NormalizedSubject TEXT,
      Message_Preview TEXT,
      Message_ReadFlag BOOLEAN DEFAULT 0,
      Message_RecipientList TEXT,
      Message_DisplayTo TEXT,
      Message_SenderList TEXT,
      Message_SenderAddressList TEXT,
      Message_ToRecipientAddressList TEXT,
      Message_CCRecipientAddressList TEXT,
      Message_Size INTEGER DEFAULT 0,
      Conversation_ConversationID INTEGER DEFAULT 0,
      Message_TimeReceived DATETIME,
      Message_TimeSent DATETIME,
      Record_Categories BLOB,
      Record_Priority INTEGER DEFAULT 3,
      Record_FlagStatus INTEGER DEFAULT 0
    );

    CREATE TABLE CalendarEvents (
      Record_RecordID INTEGER PRIMARY KEY AUTOINCREMENT,
      PathToDataFile TEXT NOT NULL,
      Record_ModDate DATETIME NOT NULL,
      Record_FolderID INTEGER NOT NULL,
      Record_AccountUID INTEGER DEFAULT 0,
      Calendar_AttendeeCount INTEGER DEFAULT 0,
      Calendar_EndDateUTC DATETIME,
      Calendar_HasReminder BOOLEAN DEFAULT 0,
      Calendar_IsRecurring BOOLEAN DEFAULT 0,
      Calendar_MasterRecordID INTEGER DEFAULT 0,
      Calendar_RecurrenceID INTEGER DEFAULT 0,
      Calendar_StartDateUTC DATETIME,
      Calendar_UID TEXT NOT NULL
    );

    CREATE TABLE Contacts (
      Record_RecordID INTEGER PRIMARY KEY AUTOINCREMENT,
      PathToDataFile TEXT NOT NULL,
      Record_ModDate DATETIME NOT NULL,
      Record_FolderID INTEGER NOT NULL,
      Record_AccountUID INTEGER DEFAULT 0,
      Contact_ContactRecType INTEGER NOT NULL DEFAULT 0,
      Contact_DisplayName TEXT,
      Contact_DisplayNameSort TEXT
    );

    CREATE TABLE Tasks (
      Record_RecordID INTEGER PRIMARY KEY AUTOINCREMENT,
      PathToDataFile TEXT NOT NULL,
      Record_ModDate DATETIME NOT NULL,
      Record_FolderID INTEGER NOT NULL,
      Task_Name TEXT,
      Task_Completed BOOLEAN DEFAULT 0,
      Record_DueDate DATETIME,
      Record_StartDate DATETIME,
      Record_Priority INTEGER DEFAULT 3,
      Record_HasReminder BOOLEAN DEFAULT 0
    );

    CREATE TABLE Notes (
      Record_RecordID INTEGER PRIMARY KEY AUTOINCREMENT,
      PathToDataFile TEXT NOT NULL,
      Record_ModDate DATETIME NOT NULL,
      Record_FolderID INTEGER NOT NULL
    );
  `);

  // Insert sample folders
  const insertFolder = db.prepare(`
    INSERT INTO Folders (PathToDataFile, Record_ModDate, Folder_FolderClass, Folder_FolderType, Folder_SpecialFolderType, Folder_Name, Folder_FolderOrder)
    VALUES (?, 1700000000, 0, 0, ?, ?, ?)
  `);

  insertFolder.run('Folders/1.folder', 1, 'Inbox', 1);
  insertFolder.run('Folders/2.folder', 8, 'Sent Items', 2);
  insertFolder.run('Folders/3.folder', 10, 'Drafts', 3);
  insertFolder.run('Folders/4.folder', 9, 'Deleted Items', 4);
  insertFolder.run('Folders/5.folder', 4, 'Calendar', 5); // Calendar folder

  // Insert sample emails (folderId 1 = Inbox)
  const insertMail = db.prepare(`
    INSERT INTO Mail (PathToDataFile, Record_ModDate, Record_FolderID, Message_NormalizedSubject, Message_SenderList, Message_Preview, Message_ReadFlag, Message_TimeReceived, Message_HasAttachment)
    VALUES (?, 1700000000, ?, ?, ?, ?, ?, ?, ?)
  `);

  insertMail.run('Mail/1.msg', 1, 'Welcome to Outlook', 'Microsoft <no-reply@microsoft.com>', 'Thank you for using Outlook...', 1, 739584645, 0);
  insertMail.run('Mail/2.msg', 1, 'Meeting Tomorrow', 'John Doe <john@example.com>', 'Hi, can we meet tomorrow at 10am?', 0, 739584700, 0);
  insertMail.run('Mail/3.msg', 1, 'Project Update', 'Jane Smith <jane@example.com>', 'The project is on track...', 0, 739584800, 1);
  insertMail.run('Mail/4.msg', 2, 'RE: Meeting Tomorrow', 'You <me@example.com>', 'Sure, 10am works for me.', 1, 739584750, 0);

  // Insert sample calendar events (folderId 5 = Calendar)
  const insertEvent = db.prepare(`
    INSERT INTO CalendarEvents (PathToDataFile, Record_ModDate, Record_FolderID, Calendar_StartDateUTC, Calendar_EndDateUTC, Calendar_UID, Calendar_IsRecurring, Calendar_AttendeeCount)
    VALUES (?, 1700000000, 5, ?, ?, ?, ?, ?)
  `);

  insertEvent.run('Events/1.event', 739670400, 739674000, 'event-uid-1', 0, 2);
  insertEvent.run('Events/2.event', 739756800, 739760400, 'event-uid-2', 1, 0);

  // Insert sample contacts
  const insertContact = db.prepare(`
    INSERT INTO Contacts (PathToDataFile, Record_ModDate, Record_FolderID, Contact_DisplayName, Contact_DisplayNameSort)
    VALUES (?, 1700000000, 1, ?, ?)
  `);

  insertContact.run('Contacts/1.contact', 'John Doe', 'Doe, John');
  insertContact.run('Contacts/2.contact', 'Jane Smith', 'Smith, Jane');
  insertContact.run('Contacts/3.contact', 'Alice Johnson', 'Johnson, Alice');

  // Insert sample tasks
  const insertTask = db.prepare(`
    INSERT INTO Tasks (PathToDataFile, Record_ModDate, Record_FolderID, Task_Name, Task_Completed, Record_DueDate, Record_Priority)
    VALUES (?, 1700000000, 1, ?, ?, ?, ?)
  `);

  insertTask.run('Tasks/1.task', 'Complete report', 0, 739843200, 2);
  insertTask.run('Tasks/2.task', 'Review PRs', 1, 739756800, 3);

  // Insert sample notes
  const insertNote = db.prepare(`
    INSERT INTO Notes (PathToDataFile, Record_ModDate, Record_FolderID)
    VALUES (?, ?, 1)
  `);

  insertNote.run('Notes/1.note', 1700000000);
  insertNote.run('Notes/2.note', 1700001000);

  db.close();

  return {
    path,
    cleanup: () => {
      if (existsSync(path)) {
        unlinkSync(path);
      }
    },
  };
}

/**
 * Sample data counts for verification.
 */
export const SAMPLE_COUNTS = {
  folders: 5,
  mailFolders: 5, // All folders with FolderClass = 0
  emails: 4,
  unreadEmails: 2,
  inboxEmails: 3,
  events: 2,
  contacts: 3,
  tasks: 2,
  incompleteTasks: 1,
  notes: 2,
} as const;
