import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { createTestDatabase, SAMPLE_COUNTS } from '../../fixtures/database.js';
import { createConnection, type IConnection } from '../../../src/database/connection.js';
import { OutlookRepository, createRepository, type IRepository } from '../../../src/database/repository.js';

describe('OutlookRepository', () => {
  let testDb: { path: string; cleanup: () => void };
  let connection: IConnection;
  let repository: IRepository;

  beforeEach(() => {
    testDb = createTestDatabase();
    connection = createConnection(testDb.path);
    repository = createRepository(connection);
  });

  afterEach(() => {
    connection.close();
    testDb.cleanup();
  });

  // ---------------------------------------------------------------------------
  // Folders
  // ---------------------------------------------------------------------------

  describe('listFolders', () => {
    it('returns all mail folders', () => {
      const folders = repository.listFolders();
      expect(folders.length).toBe(SAMPLE_COUNTS.mailFolders);
    });

    it('includes folder details', () => {
      const folders = repository.listFolders();
      const inbox = folders.find((f) => f.name === 'Inbox');

      expect(inbox).toBeDefined();
      expect(inbox?.specialType).toBe(1);
      expect(inbox?.messageCount).toBe(SAMPLE_COUNTS.inboxEmails);
      expect(inbox?.unreadCount).toBe(SAMPLE_COUNTS.unreadEmails);
    });
  });

  describe('getFolder', () => {
    it('returns folder by ID', () => {
      const folders = repository.listFolders();
      const firstFolder = folders[0];

      if (firstFolder) {
        const folder = repository.getFolder(firstFolder.id);
        expect(folder).toBeDefined();
        expect(folder?.id).toBe(firstFolder.id);
      }
    });

    it('returns undefined for non-existent ID', () => {
      const folder = repository.getFolder(99999);
      expect(folder).toBeUndefined();
    });
  });

  // ---------------------------------------------------------------------------
  // Emails
  // ---------------------------------------------------------------------------

  describe('listEmails', () => {
    it('returns emails in folder', () => {
      const folders = repository.listFolders();
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const emails = repository.listEmails(inbox.id, 50, 0);
        expect(emails.length).toBe(SAMPLE_COUNTS.inboxEmails);
      }
    });

    it('respects limit parameter', () => {
      const folders = repository.listFolders();
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const emails = repository.listEmails(inbox.id, 1, 0);
        expect(emails.length).toBe(1);
      }
    });

    it('respects offset parameter', () => {
      const folders = repository.listFolders();
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const allEmails = repository.listEmails(inbox.id, 50, 0);
        const offsetEmails = repository.listEmails(inbox.id, 50, 1);
        expect(offsetEmails.length).toBe(allEmails.length - 1);
      }
    });
  });

  describe('listUnreadEmails', () => {
    it('returns only unread emails', () => {
      const folders = repository.listFolders();
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const emails = repository.listUnreadEmails(inbox.id, 50, 0);
        expect(emails.length).toBe(SAMPLE_COUNTS.unreadEmails);
        expect(emails.every((e) => e.isRead === 0)).toBe(true);
      }
    });
  });

  describe('searchEmails', () => {
    it('finds emails by subject', () => {
      const emails = repository.searchEmails('Meeting', 50);
      expect(emails.length).toBeGreaterThan(0);
      expect(emails.some((e) => e.subject?.includes('Meeting'))).toBe(true);
    });

    it('finds emails by sender', () => {
      const emails = repository.searchEmails('John', 50);
      expect(emails.length).toBeGreaterThan(0);
    });

    it('returns empty array for no matches', () => {
      const emails = repository.searchEmails('xyznonexistent', 50);
      expect(emails.length).toBe(0);
    });
  });

  describe('getEmail', () => {
    it('returns email by ID', () => {
      const folders = repository.listFolders();
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const emails = repository.listEmails(inbox.id, 1, 0);
        const firstEmail = emails[0];

        if (firstEmail) {
          const email = repository.getEmail(firstEmail.id);
          expect(email).toBeDefined();
          expect(email?.id).toBe(firstEmail.id);
        }
      }
    });

    it('returns undefined for non-existent ID', () => {
      const email = repository.getEmail(99999);
      expect(email).toBeUndefined();
    });
  });

  describe('getUnreadCount', () => {
    it('returns total unread count', () => {
      const count = repository.getUnreadCount();
      expect(count).toBe(SAMPLE_COUNTS.unreadEmails);
    });
  });

  describe('getUnreadCountByFolder', () => {
    it('returns unread count for specific folder', () => {
      const folders = repository.listFolders();
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const count = repository.getUnreadCountByFolder(inbox.id);
        expect(count).toBe(SAMPLE_COUNTS.unreadEmails);
      }
    });

    it('returns 0 for folder with no unread emails', () => {
      const folders = repository.listFolders();
      const sent = folders.find((f) => f.name === 'Sent Items');

      if (sent) {
        const count = repository.getUnreadCountByFolder(sent.id);
        expect(count).toBe(0);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Calendar
  // ---------------------------------------------------------------------------

  describe('listCalendars', () => {
    it('returns calendar folders', () => {
      const calendars = repository.listCalendars();
      expect(calendars.length).toBe(1);
      expect(calendars[0]?.name).toBe('Calendar');
    });
  });

  describe('listEvents', () => {
    it('returns events', () => {
      const events = repository.listEvents(50);
      expect(events.length).toBe(SAMPLE_COUNTS.events);
    });
  });

  describe('getEvent', () => {
    it('returns event by ID', () => {
      const events = repository.listEvents(1);
      const firstEvent = events[0];

      if (firstEvent) {
        const event = repository.getEvent(firstEvent.id);
        expect(event).toBeDefined();
        expect(event?.id).toBe(firstEvent.id);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Contacts
  // ---------------------------------------------------------------------------

  describe('listContacts', () => {
    it('returns contacts with pagination', () => {
      const contacts = repository.listContacts(50, 0);
      expect(contacts.length).toBe(SAMPLE_COUNTS.contacts);
    });

    it('returns contacts sorted by sortName', () => {
      const contacts = repository.listContacts(50, 0);
      const sortNames = contacts.map((c) => c.sortName);
      const sorted = [...sortNames].sort();
      expect(sortNames).toEqual(sorted);
    });
  });

  describe('searchContacts', () => {
    it('finds contacts by name', () => {
      const contacts = repository.searchContacts('John', 50);
      expect(contacts.length).toBeGreaterThan(0);
    });

    it('returns empty array for no matches', () => {
      const contacts = repository.searchContacts('xyznonexistent', 50);
      expect(contacts.length).toBe(0);
    });
  });

  describe('getContact', () => {
    it('returns contact by ID', () => {
      const contacts = repository.listContacts(1, 0);
      const firstContact = contacts[0];

      if (firstContact) {
        const contact = repository.getContact(firstContact.id);
        expect(contact).toBeDefined();
        expect(contact?.id).toBe(firstContact.id);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Tasks
  // ---------------------------------------------------------------------------

  describe('listTasks', () => {
    it('returns all tasks', () => {
      const tasks = repository.listTasks(50, 0);
      expect(tasks.length).toBe(SAMPLE_COUNTS.tasks);
    });
  });

  describe('listIncompleteTasks', () => {
    it('returns only incomplete tasks', () => {
      const tasks = repository.listIncompleteTasks(50, 0);
      expect(tasks.length).toBe(SAMPLE_COUNTS.incompleteTasks);
      expect(tasks.every((t) => t.isCompleted === 0)).toBe(true);
    });
  });

  describe('searchTasks', () => {
    it('finds tasks by name', () => {
      const tasks = repository.searchTasks('report', 50);
      expect(tasks.length).toBeGreaterThan(0);
    });
  });

  describe('getTask', () => {
    it('returns task by ID', () => {
      const tasks = repository.listTasks(1, 0);
      const firstTask = tasks[0];

      if (firstTask) {
        const task = repository.getTask(firstTask.id);
        expect(task).toBeDefined();
        expect(task?.id).toBe(firstTask.id);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Notes
  // ---------------------------------------------------------------------------

  describe('listNotes', () => {
    it('returns notes with pagination', () => {
      const notes = repository.listNotes(50, 0);
      expect(notes.length).toBe(SAMPLE_COUNTS.notes);
    });
  });

  describe('getNote', () => {
    it('returns note by ID', () => {
      const notes = repository.listNotes(1, 0);
      const firstNote = notes[0];

      if (firstNote) {
        const note = repository.getNote(firstNote.id);
        expect(note).toBeDefined();
        expect(note?.id).toBe(firstNote.id);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Factory function
  // ---------------------------------------------------------------------------

  describe('createRepository', () => {
    it('creates a repository instance', () => {
      const repo = createRepository(connection);
      expect(repo).toBeInstanceOf(OutlookRepository);
    });
  });
});
