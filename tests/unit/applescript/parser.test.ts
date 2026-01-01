/**
 * Unit tests for AppleScript output parser.
 */

import { describe, it, expect } from 'vitest';
import {
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
} from '../../../src/applescript/parser.js';
import { DELIMITERS } from '../../../src/applescript/scripts.js';

describe('AppleScript Parser', () => {
  describe('parseFolders', () => {
    it('should parse folder output', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}123${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}Inbox${DELIMITERS.FIELD}unreadCount${DELIMITERS.EQUALS}5`;
      const result = parseFolders(output);

      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({
        id: 123,
        name: 'Inbox',
        unreadCount: 5,
      });
    });

    it('should parse multiple folders', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}1${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}Inbox${DELIMITERS.FIELD}unreadCount${DELIMITERS.EQUALS}5${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}2${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}Sent${DELIMITERS.FIELD}unreadCount${DELIMITERS.EQUALS}0`;
      const result = parseFolders(output);

      expect(result).toHaveLength(2);
      expect(result[0]?.name).toBe('Inbox');
      expect(result[1]?.name).toBe('Sent');
    });

    it('should return empty array for empty output', () => {
      expect(parseFolders('')).toEqual([]);
      expect(parseFolders('   ')).toEqual([]);
    });
  });

  describe('parseEmails', () => {
    it('should parse email output', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}456${DELIMITERS.FIELD}subject${DELIMITERS.EQUALS}Test Subject${DELIMITERS.FIELD}senderEmail${DELIMITERS.EQUALS}test@example.com${DELIMITERS.FIELD}senderName${DELIMITERS.EQUALS}Test User${DELIMITERS.FIELD}isRead${DELIMITERS.EQUALS}true${DELIMITERS.FIELD}priority${DELIMITERS.EQUALS}high`;
      const result = parseEmails(output);

      expect(result).toHaveLength(1);
      expect(result[0]).toMatchObject({
        id: 456,
        subject: 'Test Subject',
        senderEmail: 'test@example.com',
        senderName: 'Test User',
        isRead: true,
        priority: 'high',
      });
    });

    it('should handle null values', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}1${DELIMITERS.FIELD}subject${DELIMITERS.EQUALS}${DELIMITERS.NULL}${DELIMITERS.FIELD}preview${DELIMITERS.EQUALS}`;
      const result = parseEmails(output);

      expect(result[0]?.subject).toBeNull();
      expect(result[0]?.preview).toBeNull();
    });

    it('should parse attachments list', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}1${DELIMITERS.FIELD}attachments${DELIMITERS.EQUALS}file1.pdf,file2.doc,file3.xlsx`;
      const result = parseEmails(output);

      expect(result[0]?.attachments).toEqual(['file1.pdf', 'file2.doc', 'file3.xlsx']);
    });
  });

  describe('parseEmail', () => {
    it('should return single email or null', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}1${DELIMITERS.FIELD}subject${DELIMITERS.EQUALS}Test`;
      const result = parseEmail(output);

      expect(result).not.toBeNull();
      expect(result?.subject).toBe('Test');
    });

    it('should return null for empty output', () => {
      expect(parseEmail('')).toBeNull();
    });
  });

  describe('parseCalendars', () => {
    it('should parse calendar output', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}10${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}Work Calendar`;
      const result = parseCalendars(output);

      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({
        id: 10,
        name: 'Work Calendar',
      });
    });
  });

  describe('parseEvents', () => {
    it('should parse event output', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}100${DELIMITERS.FIELD}subject${DELIMITERS.EQUALS}Team Meeting${DELIMITERS.FIELD}startTime${DELIMITERS.EQUALS}2024-01-15T10:00:00Z${DELIMITERS.FIELD}endTime${DELIMITERS.EQUALS}2024-01-15T11:00:00Z${DELIMITERS.FIELD}isAllDay${DELIMITERS.EQUALS}false${DELIMITERS.FIELD}isRecurring${DELIMITERS.EQUALS}true`;
      const result = parseEvents(output);

      expect(result).toHaveLength(1);
      expect(result[0]).toMatchObject({
        id: 100,
        subject: 'Team Meeting',
        startTime: '2024-01-15T10:00:00Z',
        endTime: '2024-01-15T11:00:00Z',
        isAllDay: false,
        isRecurring: true,
      });
    });

    it('should parse attendees', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}1${DELIMITERS.FIELD}attendees${DELIMITERS.EQUALS}john@test.com|John Doe,jane@test.com|Jane Smith`;
      const result = parseEvents(output);

      expect(result[0]?.attendees).toEqual([
        { email: 'john@test.com', name: 'John Doe' },
        { email: 'jane@test.com', name: 'Jane Smith' },
      ]);
    });
  });

  describe('parseEvent', () => {
    it('should return single event', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}1${DELIMITERS.FIELD}subject${DELIMITERS.EQUALS}Meeting`;
      const result = parseEvent(output);

      expect(result).not.toBeNull();
      expect(result?.subject).toBe('Meeting');
    });
  });

  describe('parseContacts', () => {
    it('should parse contact output', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}200${DELIMITERS.FIELD}displayName${DELIMITERS.EQUALS}John Doe${DELIMITERS.FIELD}firstName${DELIMITERS.EQUALS}John${DELIMITERS.FIELD}lastName${DELIMITERS.EQUALS}Doe${DELIMITERS.FIELD}company${DELIMITERS.EQUALS}Acme Inc${DELIMITERS.FIELD}emails${DELIMITERS.EQUALS}john@acme.com,john.doe@gmail.com`;
      const result = parseContacts(output);

      expect(result).toHaveLength(1);
      expect(result[0]).toMatchObject({
        id: 200,
        displayName: 'John Doe',
        firstName: 'John',
        lastName: 'Doe',
        company: 'Acme Inc',
        emails: ['john@acme.com', 'john.doe@gmail.com'],
      });
    });
  });

  describe('parseContact', () => {
    it('should parse full contact details', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}1${DELIMITERS.FIELD}displayName${DELIMITERS.EQUALS}Test${DELIMITERS.FIELD}homePhone${DELIMITERS.EQUALS}123-456-7890${DELIMITERS.FIELD}homeStreet${DELIMITERS.EQUALS}123 Main St`;
      const result = parseContact(output);

      expect(result).not.toBeNull();
      expect(result?.homePhone).toBe('123-456-7890');
      expect(result?.homeStreet).toBe('123 Main St');
    });
  });

  describe('parseTasks', () => {
    it('should parse task output', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}300${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}Complete report${DELIMITERS.FIELD}isCompleted${DELIMITERS.EQUALS}false${DELIMITERS.FIELD}dueDate${DELIMITERS.EQUALS}2024-01-20T17:00:00Z${DELIMITERS.FIELD}priority${DELIMITERS.EQUALS}high`;
      const result = parseTasks(output);

      expect(result).toHaveLength(1);
      expect(result[0]).toMatchObject({
        id: 300,
        name: 'Complete report',
        isCompleted: false,
        dueDate: '2024-01-20T17:00:00Z',
        priority: 'high',
      });
    });
  });

  describe('parseTask', () => {
    it('should return single task', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}1${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}Task${DELIMITERS.FIELD}isCompleted${DELIMITERS.EQUALS}true`;
      const result = parseTask(output);

      expect(result).not.toBeNull();
      expect(result?.isCompleted).toBe(true);
    });
  });

  describe('parseNotes', () => {
    it('should parse note output', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}400${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}Meeting Notes${DELIMITERS.FIELD}createdDate${DELIMITERS.EQUALS}2024-01-10T09:00:00Z${DELIMITERS.FIELD}modifiedDate${DELIMITERS.EQUALS}2024-01-15T14:30:00Z`;
      const result = parseNotes(output);

      expect(result).toHaveLength(1);
      expect(result[0]).toMatchObject({
        id: 400,
        name: 'Meeting Notes',
        createdDate: '2024-01-10T09:00:00Z',
        modifiedDate: '2024-01-15T14:30:00Z',
      });
    });
  });

  describe('parseNote', () => {
    it('should return single note', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}1${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}Note`;
      const result = parseNote(output);

      expect(result).not.toBeNull();
      expect(result?.name).toBe('Note');
    });
  });

  describe('parseCount', () => {
    it('should parse count value', () => {
      expect(parseCount('42')).toBe(42);
      expect(parseCount('  10  ')).toBe(10);
    });

    it('should return 0 for invalid input', () => {
      expect(parseCount('')).toBe(0);
      expect(parseCount('abc')).toBe(0);
    });
  });
});
