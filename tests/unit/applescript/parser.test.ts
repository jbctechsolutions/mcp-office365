/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

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
  parseRespondToEventResult,
  parseDeleteEventResult,
  parseUpdateEventResult,
  parseSendEmailResult,
  parseAttachments,
  parseSaveAttachmentResult,
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

  describe('parseRespondToEventResult', () => {
    it('should parse successful response', () => {
      const output = `${DELIMITERS.RECORD}success${DELIMITERS.EQUALS}true${DELIMITERS.FIELD}eventId${DELIMITERS.EQUALS}123`;
      const result = parseRespondToEventResult(output);
      expect(result).toEqual({ success: true, eventId: 123 });
    });

    it('should parse failure response', () => {
      const output = `${DELIMITERS.RECORD}success${DELIMITERS.EQUALS}false${DELIMITERS.FIELD}error${DELIMITERS.EQUALS}Permission denied`;
      const result = parseRespondToEventResult(output);
      expect(result).toEqual({ success: false, error: 'Permission denied' });
    });

    it('should handle empty output', () => {
      const result = parseRespondToEventResult('');
      expect(result).toBeNull();
    });
  });

  describe('parseDeleteEventResult', () => {
    it('should parse successful delete', () => {
      const output = `${DELIMITERS.RECORD}success${DELIMITERS.EQUALS}true${DELIMITERS.FIELD}eventId${DELIMITERS.EQUALS}123`;
      const result = parseDeleteEventResult(output);
      expect(result).toEqual({ success: true, eventId: 123 });
    });

    it('should parse failure', () => {
      const output = `${DELIMITERS.RECORD}success${DELIMITERS.EQUALS}false${DELIMITERS.FIELD}error${DELIMITERS.EQUALS}Event not found`;
      const result = parseDeleteEventResult(output);
      expect(result).toEqual({ success: false, error: 'Event not found' });
    });

    it('should handle empty output', () => {
      const result = parseDeleteEventResult('');
      expect(result).toBeNull();
    });
  });

  describe('parseUpdateEventResult', () => {
    it('should parse successful update with multiple fields', () => {
      const output = `${DELIMITERS.RECORD}success${DELIMITERS.EQUALS}true${DELIMITERS.FIELD}eventId${DELIMITERS.EQUALS}123${DELIMITERS.FIELD}updatedFields${DELIMITERS.EQUALS}title,location,description`;
      const result = parseUpdateEventResult(output);
      expect(result).toEqual({
        success: true,
        id: 123,
        updatedFields: ['title', 'location', 'description'],
      });
    });

    it('should parse successful update with single field', () => {
      const output = `${DELIMITERS.RECORD}success${DELIMITERS.EQUALS}true${DELIMITERS.FIELD}eventId${DELIMITERS.EQUALS}456${DELIMITERS.FIELD}updatedFields${DELIMITERS.EQUALS}title`;
      const result = parseUpdateEventResult(output);
      expect(result).toEqual({
        success: true,
        id: 456,
        updatedFields: ['title'],
      });
    });

    it('should parse successful update with no fields', () => {
      const output = `${DELIMITERS.RECORD}success${DELIMITERS.EQUALS}true${DELIMITERS.FIELD}eventId${DELIMITERS.EQUALS}789${DELIMITERS.FIELD}updatedFields${DELIMITERS.EQUALS}`;
      const result = parseUpdateEventResult(output);
      expect(result).toEqual({
        success: true,
        id: 789,
        updatedFields: [],
      });
    });

    it('should parse failure', () => {
      const output = `${DELIMITERS.RECORD}success${DELIMITERS.EQUALS}false${DELIMITERS.FIELD}error${DELIMITERS.EQUALS}Event not found`;
      const result = parseUpdateEventResult(output);
      expect(result).toEqual({ success: false, error: 'Event not found' });
    });

    it('should parse failure with missing error field', () => {
      const output = `${DELIMITERS.RECORD}success${DELIMITERS.EQUALS}false`;
      const result = parseUpdateEventResult(output);
      expect(result).toEqual({ success: false, error: 'Unknown error' });
    });

    it('should handle empty output', () => {
      const result = parseUpdateEventResult('');
      expect(result).toBeNull();
    });

    it('should handle missing record', () => {
      const result = parseUpdateEventResult('invalid');
      expect(result).toBeNull();
    });
  });

  describe('parseSendEmailResult', () => {
    it('should parse successful send', () => {
      const output = `${DELIMITERS.RECORD}success${DELIMITERS.EQUALS}true${DELIMITERS.FIELD}messageId${DELIMITERS.EQUALS}12345${DELIMITERS.FIELD}sentAt${DELIMITERS.EQUALS}2024-01-15T10:30:00Z`;
      const result = parseSendEmailResult(output);
      expect(result).toEqual({
        success: true,
        messageId: '12345',
        sentAt: '2024-01-15T10:30:00Z',
      });
    });

    it('should parse failure', () => {
      const output = `${DELIMITERS.RECORD}success${DELIMITERS.EQUALS}false${DELIMITERS.FIELD}error${DELIMITERS.EQUALS}Recipient not found`;
      const result = parseSendEmailResult(output);
      expect(result).toEqual({ success: false, error: 'Recipient not found' });
    });

    it('should parse failure with missing error field', () => {
      const output = `${DELIMITERS.RECORD}success${DELIMITERS.EQUALS}false`;
      const result = parseSendEmailResult(output);
      expect(result).toEqual({ success: false, error: 'Unknown error' });
    });

    it('should handle empty output', () => {
      const result = parseSendEmailResult('');
      expect(result).toBeNull();
    });

    it('should handle missing record', () => {
      const result = parseSendEmailResult('invalid');
      expect(result).toBeNull();
    });
  });

  // ===========================================================================
  // Attachment Parsers
  // ===========================================================================

  describe('parseAttachments', () => {
    it('should parse multiple attachment records', () => {
      const output =
        `${DELIMITERS.RECORD}index${DELIMITERS.EQUALS}1${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}report.pdf${DELIMITERS.FIELD}fileSize${DELIMITERS.EQUALS}102400${DELIMITERS.FIELD}contentType${DELIMITERS.EQUALS}application/pdf` +
        `${DELIMITERS.RECORD}index${DELIMITERS.EQUALS}2${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}image.png${DELIMITERS.FIELD}fileSize${DELIMITERS.EQUALS}51200${DELIMITERS.FIELD}contentType${DELIMITERS.EQUALS}image/png`;
      const result = parseAttachments(output);

      expect(result).toHaveLength(2);
      expect(result[0]).toEqual({
        index: 1,
        name: 'report.pdf',
        fileSize: 102400,
        contentType: 'application/pdf',
      });
      expect(result[1]).toEqual({
        index: 2,
        name: 'image.png',
        fileSize: 51200,
        contentType: 'image/png',
      });
    });

    it('should handle empty output', () => {
      expect(parseAttachments('')).toEqual([]);
      expect(parseAttachments('   ')).toEqual([]);
    });

    it('should default to application/octet-stream for missing contentType', () => {
      const output = `${DELIMITERS.RECORD}index${DELIMITERS.EQUALS}1${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}file.bin${DELIMITERS.FIELD}fileSize${DELIMITERS.EQUALS}1024`;
      const result = parseAttachments(output);

      expect(result).toHaveLength(1);
      expect(result[0]!.contentType).toBe('application/octet-stream');
    });

    it('should default to 0 for missing fileSize', () => {
      const output = `${DELIMITERS.RECORD}index${DELIMITERS.EQUALS}1${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}file.txt${DELIMITERS.FIELD}contentType${DELIMITERS.EQUALS}text/plain`;
      const result = parseAttachments(output);

      expect(result).toHaveLength(1);
      expect(result[0]!.fileSize).toBe(0);
    });
  });

  describe('parseSaveAttachmentResult', () => {
    it('should parse success result', () => {
      const output = `${DELIMITERS.RECORD}success${DELIMITERS.EQUALS}true${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}report.pdf${DELIMITERS.FIELD}savedTo${DELIMITERS.EQUALS}/tmp/report.pdf${DELIMITERS.FIELD}fileSize${DELIMITERS.EQUALS}102400`;
      const result = parseSaveAttachmentResult(output);

      expect(result).toEqual({
        success: true,
        name: 'report.pdf',
        savedTo: '/tmp/report.pdf',
        fileSize: 102400,
      });
    });

    it('should parse failure result', () => {
      const output = `${DELIMITERS.RECORD}success${DELIMITERS.EQUALS}false${DELIMITERS.FIELD}error${DELIMITERS.EQUALS}Permission denied`;
      const result = parseSaveAttachmentResult(output);

      expect(result).toEqual({
        success: false,
        error: 'Permission denied',
      });
    });

    it('should return null for empty output', () => {
      expect(parseSaveAttachmentResult('')).toBeNull();
      expect(parseSaveAttachmentResult('   ')).toBeNull();
    });
  });

  // ===========================================================================
  // parseEmail with attachmentDetails
  // ===========================================================================

  describe('parseEmail with attachmentDetails', () => {
    it('should include attachmentDetails array', () => {
      const output = `${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}1${DELIMITERS.FIELD}subject${DELIMITERS.EQUALS}Test${DELIMITERS.FIELD}attachmentDetails${DELIMITERS.EQUALS}1|report.pdf|102400|application/pdf,2|image.png|51200|image/png`;
      const result = parseEmail(output);

      expect(result).not.toBeNull();
      expect(result!.attachmentDetails).toHaveLength(2);
      expect(result!.attachmentDetails[0]).toEqual({
        index: 1,
        name: 'report.pdf',
        fileSize: 102400,
        contentType: 'application/pdf',
      });
      expect(result!.attachmentDetails[1]).toEqual({
        index: 2,
        name: 'image.png',
        fileSize: 51200,
        contentType: 'image/png',
      });
    });
  });
});
