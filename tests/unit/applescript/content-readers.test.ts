import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('../../../src/applescript/executor.js', () => ({
  executeAppleScript: vi.fn(),
  executeAppleScriptOrThrow: vi.fn(),
}));

vi.mock('../../../src/applescript/scripts.js', () => ({
  getMessage: vi.fn(() => 'mock-email-script'),
  getEvent: vi.fn(() => 'mock-event-script'),
  getContact: vi.fn(() => 'mock-contact-script'),
  getTask: vi.fn(() => 'mock-task-script'),
  getNote: vi.fn(() => 'mock-note-script'),
  listAttachments: vi.fn(() => 'mock-list-attachments-script'),
  saveAttachment: vi.fn(() => 'mock-save-attachment-script'),
}));

vi.mock('../../../src/applescript/parser.js', () => ({
  parseEmail: vi.fn(),
  parseEvent: vi.fn(),
  parseContact: vi.fn(),
  parseTask: vi.fn(),
  parseNote: vi.fn(),
  parseAttachments: vi.fn(),
  parseSaveAttachmentResult: vi.fn(),
}));

import {
  createEmailPath,
  createEventPath,
  createContactPath,
  createTaskPath,
  createNotePath,
  AppleScriptEmailContentReader,
  AppleScriptEventContentReader,
  AppleScriptContactContentReader,
  AppleScriptTaskContentReader,
  AppleScriptNoteContentReader,
  AppleScriptAttachmentReader,
  createAppleScriptContentReaders,
} from '../../../src/applescript/content-readers.js';
import { executeAppleScript, executeAppleScriptOrThrow } from '../../../src/applescript/executor.js';
import * as parser from '../../../src/applescript/parser.js';

const mockedExecute = vi.mocked(executeAppleScript);
const mockedExecuteOrThrow = vi.mocked(executeAppleScriptOrThrow);
const mockedParseEmail = vi.mocked(parser.parseEmail);
const mockedParseEvent = vi.mocked(parser.parseEvent);
const mockedParseContact = vi.mocked(parser.parseContact);
const mockedParseTask = vi.mocked(parser.parseTask);
const mockedParseNote = vi.mocked(parser.parseNote);
const mockedParseAttachments = vi.mocked(parser.parseAttachments);
const mockedParseSaveAttachmentResult = vi.mocked(parser.parseSaveAttachmentResult);

// =============================================================================
// Path Helpers
// =============================================================================

describe('path helpers', () => {
  it('createEmailPath returns correct format', () => {
    expect(createEmailPath(123)).toBe('applescript-email:123');
  });

  it('createEventPath returns correct format', () => {
    expect(createEventPath(456)).toBe('applescript-event:456');
  });

  it('createContactPath returns correct format', () => {
    expect(createContactPath(789)).toBe('applescript-contact:789');
  });

  it('createTaskPath returns correct format', () => {
    expect(createTaskPath(101)).toBe('applescript-task:101');
  });

  it('createNotePath returns correct format', () => {
    expect(createNotePath(202)).toBe('applescript-note:202');
  });
});

// =============================================================================
// Email Content Reader
// =============================================================================

describe('AppleScriptEmailContentReader', () => {
  let reader: AppleScriptEmailContentReader;

  beforeEach(() => {
    vi.clearAllMocks();
    reader = new AppleScriptEmailContentReader();
  });

  it('returns html content for valid email path', () => {
    mockedExecute.mockReturnValue({ success: true, output: 'raw' });
    mockedParseEmail.mockReturnValue({
      id: 123,
      subject: 'Test',
      sender: 'a@b.com',
      senderName: 'A',
      recipients: '',
      date: '2024-01-01',
      isRead: true,
      htmlContent: '<p>Hello</p>',
      plainContent: 'Hello',
      hasAttachments: false,
      categories: '',
      folderId: 1,
    });

    const result = reader.readEmailBody('applescript-email:123');
    expect(result).toBe('<p>Hello</p>');
  });

  it('returns null for null path', () => {
    expect(reader.readEmailBody(null)).toBeNull();
  });

  it('returns null for wrong prefix', () => {
    expect(reader.readEmailBody('wrong-prefix:123')).toBeNull();
  });

  it('returns null when AppleScript fails', () => {
    mockedExecute.mockReturnValue({ success: false, output: '', error: 'failed' });
    expect(reader.readEmailBody('applescript-email:123')).toBeNull();
  });

  it('returns null when parser returns null', () => {
    mockedExecute.mockReturnValue({ success: true, output: 'raw' });
    mockedParseEmail.mockReturnValue(null);
    expect(reader.readEmailBody('applescript-email:123')).toBeNull();
  });
});

// =============================================================================
// Event Content Reader
// =============================================================================

describe('AppleScriptEventContentReader', () => {
  let reader: AppleScriptEventContentReader;

  beforeEach(() => {
    vi.clearAllMocks();
    reader = new AppleScriptEventContentReader();
  });

  it('returns event details for valid path', () => {
    mockedExecute.mockReturnValue({ success: true, output: 'raw' });
    mockedParseEvent.mockReturnValue({
      id: 456,
      subject: 'Meeting',
      startDate: '2024-01-01T10:00:00',
      endDate: '2024-01-01T11:00:00',
      location: 'Room 1',
      organizer: 'org@test.com',
      isAllDay: false,
      recurrence: null,
      calendarId: 1,
      htmlContent: '<p>Notes</p>',
      plainContent: 'Notes',
      attendees: [{ email: 'a@b.com', name: 'A', status: 'accepted' }],
    });

    const result = reader.readEventDetails('applescript-event:456');
    expect(result).toEqual({
      title: 'Meeting',
      location: 'Room 1',
      description: '<p>Notes</p>',
      organizer: 'org@test.com',
      attendees: [{ email: 'a@b.com', name: 'A', status: 'unknown' }],
    });
  });

  it('returns null for null path', () => {
    expect(reader.readEventDetails(null)).toBeNull();
  });

  it('returns null when AppleScript fails', () => {
    mockedExecute.mockReturnValue({ success: false, output: '', error: 'err' });
    expect(reader.readEventDetails('applescript-event:456')).toBeNull();
  });
});

// =============================================================================
// Contact Content Reader
// =============================================================================

describe('AppleScriptContactContentReader', () => {
  let reader: AppleScriptContactContentReader;

  beforeEach(() => {
    vi.clearAllMocks();
    reader = new AppleScriptContactContentReader();
  });

  it('returns contact details for valid path', () => {
    mockedExecute.mockReturnValue({ success: true, output: 'raw' });
    mockedParseContact.mockReturnValue({
      id: 789,
      firstName: 'John',
      lastName: 'Doe',
      middleName: null,
      nickname: null,
      company: 'Acme',
      jobTitle: 'Dev',
      department: null,
      emails: ['john@acme.com'],
      homePhone: '555-1234',
      workPhone: null,
      mobilePhone: '555-5678',
      homeStreet: '123 Main St',
      homeCity: 'Springfield',
      homeState: 'IL',
      homeZip: '62701',
      homeCountry: 'US',
      notes: 'A note',
    });

    const result = reader.readContactDetails('applescript-contact:789');
    expect(result).not.toBeNull();
    expect(result!.firstName).toBe('John');
    expect(result!.emails).toEqual([{ type: 'work', address: 'john@acme.com' }]);
    expect(result!.phones).toHaveLength(2);
    expect(result!.addresses).toHaveLength(1);
  });

  it('returns null for null path', () => {
    expect(reader.readContactDetails(null)).toBeNull();
  });

  it('returns null when AppleScript fails', () => {
    mockedExecute.mockReturnValue({ success: false, output: '', error: 'err' });
    expect(reader.readContactDetails('applescript-contact:789')).toBeNull();
  });
});

// =============================================================================
// Task Content Reader
// =============================================================================

describe('AppleScriptTaskContentReader', () => {
  let reader: AppleScriptTaskContentReader;

  beforeEach(() => {
    vi.clearAllMocks();
    reader = new AppleScriptTaskContentReader();
  });

  it('returns task details for valid path', () => {
    mockedExecute.mockReturnValue({ success: true, output: 'raw' });
    mockedParseTask.mockReturnValue({
      id: 101,
      name: 'My Task',
      dueDate: '2024-01-15',
      startDate: null,
      completedDate: '2024-01-14',
      priority: 5,
      isComplete: true,
      htmlContent: '<p>Details</p>',
      plainContent: 'Details',
    });

    const result = reader.readTaskDetails('applescript-task:101');
    expect(result).toEqual({
      body: '<p>Details</p>',
      completedDate: '2024-01-14',
      reminderDate: null,
      categories: [],
    });
  });

  it('returns null for null path', () => {
    expect(reader.readTaskDetails(null)).toBeNull();
  });

  it('returns null when AppleScript fails', () => {
    mockedExecute.mockReturnValue({ success: false, output: '', error: 'err' });
    expect(reader.readTaskDetails('applescript-task:101')).toBeNull();
  });
});

// =============================================================================
// Note Content Reader
// =============================================================================

describe('AppleScriptNoteContentReader', () => {
  let reader: AppleScriptNoteContentReader;

  beforeEach(() => {
    vi.clearAllMocks();
    reader = new AppleScriptNoteContentReader();
  });

  it('returns note details for valid path', () => {
    mockedExecute.mockReturnValue({ success: true, output: 'raw' });
    mockedParseNote.mockReturnValue({
      id: 202,
      name: 'My Note',
      createdDate: '2024-01-01',
      modifiedDate: '2024-01-02',
      htmlContent: '<p>Note body</p>',
      plainContent: 'Note body',
    });

    const result = reader.readNoteDetails('applescript-note:202');
    expect(result).not.toBeNull();
    expect(result!.title).toBe('My Note');
    expect(result!.body).toBe('<p>Note body</p>');
    expect(result!.preview).toBe('<p>Note body</p>');
  });

  it('returns null for null path', () => {
    expect(reader.readNoteDetails(null)).toBeNull();
  });

  it('returns null when AppleScript fails', () => {
    mockedExecute.mockReturnValue({ success: false, output: '', error: 'err' });
    expect(reader.readNoteDetails('applescript-note:202')).toBeNull();
  });
});

// =============================================================================
// Factory
// =============================================================================

describe('createAppleScriptContentReaders', () => {
  it('returns all content readers', () => {
    const readers = createAppleScriptContentReaders();
    expect(readers.email).toBeInstanceOf(AppleScriptEmailContentReader);
    expect(readers.event).toBeInstanceOf(AppleScriptEventContentReader);
    expect(readers.contact).toBeInstanceOf(AppleScriptContactContentReader);
    expect(readers.task).toBeInstanceOf(AppleScriptTaskContentReader);
    expect(readers.note).toBeInstanceOf(AppleScriptNoteContentReader);
  });

  it('includes attachment reader', () => {
    const readers = createAppleScriptContentReaders();
    expect(readers.attachment).toBeInstanceOf(AppleScriptAttachmentReader);
  });
});

// =============================================================================
// Attachment Reader
// =============================================================================

describe('AppleScriptAttachmentReader', () => {
  let reader: AppleScriptAttachmentReader;

  beforeEach(() => {
    vi.clearAllMocks();
    reader = new AppleScriptAttachmentReader();
  });

  describe('listAttachments', () => {
    it('returns correct AttachmentInfo[] when AppleScript succeeds', () => {
      mockedExecute.mockReturnValue({ success: true, output: 'raw-attachments' });
      mockedParseAttachments.mockReturnValue([
        { index: 1, name: 'report.pdf', fileSize: 102400, contentType: 'application/pdf' },
        { index: 2, name: 'image.png', fileSize: 51200, contentType: 'image/png' },
      ]);

      const result = reader.listAttachments(123);

      expect(result).toHaveLength(2);
      expect(result[0]).toEqual({
        index: 1,
        name: 'report.pdf',
        size: 102400,
        contentType: 'application/pdf',
      });
      expect(result[1]).toEqual({
        index: 2,
        name: 'image.png',
        size: 51200,
        contentType: 'image/png',
      });
    });

    it('returns empty array when AppleScript fails', () => {
      mockedExecute.mockReturnValue({ success: false, output: '', error: 'failed' });

      const result = reader.listAttachments(123);
      expect(result).toEqual([]);
    });
  });

  describe('saveAttachment', () => {
    it('returns success result', () => {
      mockedExecuteOrThrow.mockReturnValue('raw-save-output');
      mockedParseSaveAttachmentResult.mockReturnValue({
        success: true,
        name: 'report.pdf',
        savedTo: '/tmp/report.pdf',
        fileSize: 102400,
      });

      const result = reader.saveAttachment(123, 1, '/tmp/report.pdf');

      expect(result).toEqual({
        success: true,
        name: 'report.pdf',
        savedTo: '/tmp/report.pdf',
        fileSize: 102400,
      });
    });

    it('throws on null parse result', () => {
      mockedExecuteOrThrow.mockReturnValue('invalid-output');
      mockedParseSaveAttachmentResult.mockReturnValue(null);

      expect(() => {
        reader.saveAttachment(123, 1, '/tmp/report.pdf');
      }).toThrow('Failed to parse save attachment response');
    });
  });
});
