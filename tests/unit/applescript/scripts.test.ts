/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Unit tests for AppleScript template generation.
 */

import { describe, it, expect } from 'vitest';
import {
  respondToEvent,
  deleteEvent,
  sendEmail,
  DELIMITERS,
  LIST_MAIL_FOLDERS,
  LIST_CALENDARS,
  listMessages,
  searchMessages,
  getMessage,
  getUnreadCount,
  listAttachments,
  saveAttachment,
  listEvents,
  getEvent,
  searchEvents,
  updateEvent,
  createEvent,
  listContacts,
  searchContacts,
  getContact,
  listTasks,
  searchTasks,
  getTask,
  listNotes,
  searchNotes,
  getNote,
  moveMessage,
  deleteMessage,
  archiveMessage,
  junkMessage,
  setMessageReadStatus,
  setMessageFlag,
  setMessageCategories,
  createMailFolder,
  deleteMailFolder,
  renameMailFolder,
  moveMailFolder,
  emptyMailFolder,
} from '../../../src/applescript/scripts.js';

describe('respondToEvent', () => {
  it('should generate accept script with comment', () => {
    const script = respondToEvent({
      eventId: 123,
      response: 'accept',
      sendResponse: true,
      comment: 'I will be there',
    });

    expect(script).toContain('calendar event id 123');
    expect(script).toContain('accept');
    expect(script).toContain('I will be there');
  });

  it('should generate decline script without sending response', () => {
    const script = respondToEvent({
      eventId: 456,
      response: 'decline',
      sendResponse: false,
    });

    expect(script).toContain('calendar event id 456');
    expect(script).toContain('decline');
  });

  it('should generate tentative accept script', () => {
    const script = respondToEvent({
      eventId: 789,
      response: 'tentative',
      sendResponse: true,
    });

    expect(script).toContain('calendar event id 789');
    expect(script).toContain('tentative');
  });
});

describe('deleteEvent', () => {
  it('should generate script for single instance', () => {
    const script = deleteEvent({ eventId: 123, applyTo: 'this_instance' });
    expect(script).toContain('calendar event id 123');
    expect(script).toContain('delete');
    expect(script).toContain('Deleting single instance');
  });

  it('should generate script for all in series', () => {
    const script = deleteEvent({ eventId: 456, applyTo: 'all_in_series' });
    expect(script).toContain('calendar event id 456');
    expect(script).toContain('delete');
    expect(script).toContain('Deleting entire series');
  });

  it('should include success output format', () => {
    const script = deleteEvent({ eventId: 789, applyTo: 'this_instance' });
    expect(script).toContain('success{{=}}true');
    expect(script).toContain('eventId{{=}}');
  });
});

describe('sendEmail', () => {
  it('should generate plain text email with single recipient', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test Subject',
      body: 'Test body',
      bodyType: 'plain',
    });

    expect(script).toContain('Test Subject');
    expect(script).toContain('Test body');
    expect(script).toContain('plain text content');
    expect(script).toContain('test@example.com');
  });

  it('should generate HTML email', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'HTML Test',
      body: '<p>HTML body</p>',
      bodyType: 'html',
    });

    expect(script).toContain('HTML Test');
    expect(script).toContain('html content');
    expect(script).toContain('HTML body');
  });

  it('should include CC and BCC recipients', () => {
    const script = sendEmail({
      to: ['to@example.com'],
      subject: 'Test',
      body: 'Body',
      bodyType: 'plain',
      cc: ['cc1@example.com', 'cc2@example.com'],
      bcc: ['bcc@example.com'],
    });

    expect(script).toContain('cc1@example.com');
    expect(script).toContain('cc2@example.com');
    expect(script).toContain('bcc@example.com');
    expect(script).toContain('recipient cc');
    expect(script).toContain('recipient bcc');
  });

  it('should include reply-to address', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test',
      body: 'Body',
      bodyType: 'plain',
      replyTo: 'reply@example.com',
    });

    expect(script).toContain('reply to of newMessage to "reply@example.com"');
  });

  it('should include attachments', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test',
      body: 'Body',
      bodyType: 'plain',
      attachments: [
        { path: '/path/to/file.pdf' },
        { path: '/path/to/image.png', name: 'screenshot.png' },
      ],
    });

    expect(script).toContain('POSIX file "/path/to/file.pdf"');
    expect(script).toContain('POSIX file "/path/to/image.png"');
    expect(script).toContain('make new attachment');
  });

  it('should include account ID', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test',
      body: 'Body',
      bodyType: 'plain',
      accountId: 123,
    });

    expect(script).toContain('account id 123');
  });

  it('should handle special characters in subject and body', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test "quotes" and \\backslash',
      body: 'Body with "quotes"',
      bodyType: 'plain',
    });

    expect(script).toContain('Test');
    expect(script).toContain('Body with');
  });

  it('should include success output format', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test',
      body: 'Body',
      bodyType: 'plain',
    });

    expect(script).toContain('success{{=}}true');
    expect(script).toContain('messageId{{=}}');
    expect(script).toContain('sentAt{{=}}');
  });

  it('should include error handling', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test',
      body: 'Body',
      bodyType: 'plain',
    });

    expect(script).toContain('on error errMsg');
    expect(script).toContain('success{{=}}false');
    expect(script).toContain('error{{=}}');
  });
});

// =============================================================================
// Constants
// =============================================================================

describe('DELIMITERS', () => {
  it('has expected delimiter values', () => {
    expect(DELIMITERS.RECORD).toBe('{{RECORD}}');
    expect(DELIMITERS.FIELD).toBe('{{FIELD}}');
    expect(DELIMITERS.EQUALS).toBe('{{=}}');
    expect(DELIMITERS.NULL).toBe('{{NULL}}');
  });
});

describe('LIST_MAIL_FOLDERS', () => {
  it('is a valid AppleScript template', () => {
    expect(LIST_MAIL_FOLDERS).toContain('tell application "Microsoft Outlook"');
    expect(LIST_MAIL_FOLDERS).toContain('mail folders');
  });
});

describe('LIST_CALENDARS', () => {
  it('is a valid AppleScript template', () => {
    expect(LIST_CALENDARS).toContain('tell application "Microsoft Outlook"');
    expect(LIST_CALENDARS).toContain('calendars');
  });
});

// =============================================================================
// Mail Scripts
// =============================================================================

describe('listMessages', () => {
  it('generates script with folder ID and pagination', () => {
    const script = listMessages(42, 10, 0, false);
    expect(script).toContain('mail folder id 42');
    expect(script).toContain('set startIdx to 1');
    expect(script).toContain('set endIdx to 10');
  });

  it('applies unread filter when unreadOnly is true', () => {
    const script = listMessages(42, 10, 0, true);
    expect(script).toContain('whose is read is false');
  });

  it('does not apply unread filter when unreadOnly is false', () => {
    const script = listMessages(42, 10, 0, false);
    expect(script).not.toContain('whose is read is false');
  });

  it('calculates offset correctly', () => {
    const script = listMessages(42, 10, 5, false);
    expect(script).toContain('set startIdx to 6');
    expect(script).toContain('set endIdx to 15');
  });
});

describe('searchMessages', () => {
  it('generates search script with query', () => {
    const script = searchMessages('test query', null, 20);
    expect(script).toContain('test query');
    expect(script).toContain('set maxResults to 20');
  });

  it('includes folder clause when folderId is provided', () => {
    const script = searchMessages('test', 42, 10);
    expect(script).toContain('of mail folder id 42');
  });

  it('omits folder clause when folderId is null', () => {
    const script = searchMessages('test', null, 10);
    expect(script).not.toContain('of mail folder id');
  });

  it('does not use address of sender in WHERE clause', () => {
    const script = searchMessages('test', null, 10);
    // The WHERE clause should only filter by subject, not sender
    expect(script).toContain('whose subject contains');
    expect(script).not.toMatch(/whose.*address of sender/);
  });

  it('includes sender scan phase with try/catch protection', () => {
    const script = searchMessages('test', null, 10);
    // Phase 2 should scan messages and access sender safely
    expect(script).toContain('Phase 2');
    expect(script).toContain('if mSender contains');
    // Sender access must be inside try/catch
    expect(script).toMatch(/try\s+set mSender to address of sender of m\s+end try/);
  });

  it('includes deduplication via matchedIds', () => {
    const script = searchMessages('test', null, 10);
    expect(script).toContain('set matchedIds to {}');
    expect(script).toContain('matchedIds does not contain mId');
  });

  it('applies sender scan limit', () => {
    const script = searchMessages('test', null, 10);
    expect(script).toContain('scanLimit');
    // Should cap at SENDER_SCAN_LIMIT (500)
    expect(script).toContain('500');
  });
});

describe('getMessage', () => {
  it('generates script for message ID', () => {
    const script = getMessage(123);
    expect(script).toContain('message id 123');
    expect(script).toContain('subject of m');
    expect(script).toContain('to recipients');
    expect(script).toContain('attachments');
  });
});

describe('getUnreadCount', () => {
  it('generates script for folder ID', () => {
    const script = getUnreadCount(42);
    expect(script).toContain('mail folder id 42');
    expect(script).toContain('unread count');
  });
});

// =============================================================================
// Calendar Scripts
// =============================================================================

describe('listEvents', () => {
  it('generates script with calendar filter', () => {
    const script = listEvents(5, null, null, 50);
    expect(script).toContain('of calendar id 5');
    expect(script).toContain('set maxEvents to 50');
  });

  it('omits calendar clause when calendarId is null', () => {
    const script = listEvents(null, null, null, 50);
    expect(script).not.toContain('of calendar id');
  });
});

describe('getEvent', () => {
  it('generates script for event ID', () => {
    const script = getEvent(789);
    expect(script).toContain('calendar event id 789');
    expect(script).toContain('attendees');
    expect(script).toContain('organizer');
  });
});

describe('searchEvents', () => {
  it('generates search script with query', () => {
    const script = searchEvents('standup', null, null, 10);
    expect(script).toContain('standup');
    expect(script).toContain('set resultCount to count of searchResults');
  });
});

describe('updateEvent', () => {
  it('generates update script with title', () => {
    const script = updateEvent({
      eventId: 123,
      applyTo: 'this_instance',
      updates: { title: 'New Title' },
    });
    expect(script).toContain('calendar event id 123');
    expect(script).toContain('set subject of myEvent to "New Title"');
    expect(script).toContain('Updating single instance');
  });

  it('includes all_in_series comment', () => {
    const script = updateEvent({
      eventId: 123,
      applyTo: 'all_in_series',
      updates: { title: 'Test' },
    });
    expect(script).toContain('Updating entire series');
  });

  it('generates update with location and description', () => {
    const script = updateEvent({
      eventId: 123,
      applyTo: 'this_instance',
      updates: { location: 'Room A', description: 'Meeting notes' },
    });
    expect(script).toContain('set location of myEvent to "Room A"');
    expect(script).toContain('set content of myEvent to "Meeting notes"');
  });

  it('generates update with isAllDay', () => {
    const script = updateEvent({
      eventId: 123,
      applyTo: 'this_instance',
      updates: { isAllDay: true },
    });
    expect(script).toContain('set all day flag of myEvent to true');
  });

  it('tracks updated fields in output', () => {
    const script = updateEvent({
      eventId: 123,
      applyTo: 'this_instance',
      updates: { title: 'T', location: 'L' },
    });
    expect(script).toContain('updatedFields{{=}}title,location');
  });
});

describe('createEvent', () => {
  it('generates create script with required fields', () => {
    const script = createEvent({
      title: 'Team Meeting',
      startYear: 2024, startMonth: 6, startDay: 15, startHours: 10, startMinutes: 0,
      endYear: 2024, endMonth: 6, endDay: 15, endHours: 11, endMinutes: 0,
    });
    expect(script).toContain('Team Meeting');
    expect(script).toContain('set year of theStartDate to 2024');
    expect(script).toContain('set month of theStartDate to 6');
  });

  it('includes calendar ID when provided', () => {
    const script = createEvent({
      title: 'Test',
      startYear: 2024, startMonth: 1, startDay: 1, startHours: 9, startMinutes: 0,
      endYear: 2024, endMonth: 1, endDay: 1, endHours: 10, endMinutes: 0,
      calendarId: 42,
    });
    expect(script).toContain('at calendar id 42');
  });

  it('includes location when provided', () => {
    const script = createEvent({
      title: 'Test',
      startYear: 2024, startMonth: 1, startDay: 1, startHours: 9, startMinutes: 0,
      endYear: 2024, endMonth: 1, endDay: 1, endHours: 10, endMinutes: 0,
      location: 'Room 101',
    });
    expect(script).toContain('location:"Room 101"');
  });

  it('includes isAllDay flag when true', () => {
    const script = createEvent({
      title: 'Day Off',
      startYear: 2024, startMonth: 1, startDay: 1, startHours: 0, startMinutes: 0,
      endYear: 2024, endMonth: 1, endDay: 1, endHours: 23, endMinutes: 59,
      isAllDay: true,
    });
    expect(script).toContain('all day flag:true');
  });

  it('includes description when provided', () => {
    const script = createEvent({
      title: 'Test',
      startYear: 2024, startMonth: 1, startDay: 1, startHours: 9, startMinutes: 0,
      endYear: 2024, endMonth: 1, endDay: 1, endHours: 10, endMinutes: 0,
      description: 'A meeting about things',
    });
    expect(script).toContain('A meeting about things');
  });

  it('includes recurrence when provided', () => {
    const script = createEvent({
      title: 'Weekly Standup',
      startYear: 2024, startMonth: 1, startDay: 1, startHours: 9, startMinutes: 0,
      endYear: 2024, endMonth: 1, endDay: 1, endHours: 9, endMinutes: 30,
      recurrence: { frequency: 'weekly', interval: 1, daysOfWeek: ['monday', 'wednesday'] },
    });
    expect(script).toContain('is recurring of newEvent to true');
    expect(script).toContain('weekly recurrence');
    expect(script).toContain('Monday, Wednesday');
  });
});

// =============================================================================
// Contact Scripts
// =============================================================================

describe('listContacts', () => {
  it('generates script with pagination', () => {
    const script = listContacts(10, 5);
    expect(script).toContain('set startIdx to 6');
    expect(script).toContain('set endIdx to 15');
    expect(script).toContain('contacts');
  });
});

describe('searchContacts', () => {
  it('generates search script with query', () => {
    const script = searchContacts('John', 20);
    expect(script).toContain('John');
    expect(script).toContain('set maxResults to 20');
    expect(script).toContain('display name contains');
  });
});

describe('getContact', () => {
  it('generates script for contact ID', () => {
    const script = getContact(555);
    expect(script).toContain('contact id 555');
    expect(script).toContain('first name');
    expect(script).toContain('email addresses');
    expect(script).toContain('home phone number');
  });
});

// =============================================================================
// Task Scripts
// =============================================================================

describe('listTasks', () => {
  it('generates script excluding completed tasks', () => {
    const script = listTasks(10, 0, false);
    expect(script).toContain('whose is completed is false');
  });

  it('generates script including completed tasks', () => {
    const script = listTasks(10, 0, true);
    expect(script).not.toContain('whose is completed is false');
  });

  it('handles pagination', () => {
    const script = listTasks(10, 5, false);
    expect(script).toContain('set startIdx to 6');
    expect(script).toContain('set endIdx to 15');
  });
});

describe('searchTasks', () => {
  it('generates search script', () => {
    const script = searchTasks('report', 10);
    expect(script).toContain('report');
    expect(script).toContain('name contains');
  });
});

describe('getTask', () => {
  it('generates script for task ID', () => {
    const script = getTask(101);
    expect(script).toContain('task id 101');
    expect(script).toContain('due date');
    expect(script).toContain('is completed');
  });
});

// =============================================================================
// Note Scripts
// =============================================================================

describe('listNotes', () => {
  it('generates script with pagination', () => {
    const script = listNotes(10, 0);
    expect(script).toContain('set startIdx to 1');
    expect(script).toContain('notes');
  });
});

describe('searchNotes', () => {
  it('generates search script', () => {
    const script = searchNotes('meeting', 5);
    expect(script).toContain('meeting');
    expect(script).toContain('name contains');
  });
});

describe('getNote', () => {
  it('generates script for note ID', () => {
    const script = getNote(202);
    expect(script).toContain('note id 202');
    expect(script).toContain('content of n');
  });
});

// =============================================================================
// Write Operation Scripts
// =============================================================================

describe('moveMessage', () => {
  it('generates move script', () => {
    const script = moveMessage(100, 200);
    expect(script).toContain('message id 100');
    expect(script).toContain('mail folder id 200');
    expect(script).toContain('move m to targetFolder');
  });
});

describe('deleteMessage', () => {
  it('generates delete script', () => {
    const script = deleteMessage(100);
    expect(script).toContain('message id 100');
    expect(script).toContain('deleted items');
  });
});

describe('archiveMessage', () => {
  it('generates archive script', () => {
    const script = archiveMessage(100);
    expect(script).toContain('message id 100');
    expect(script).toContain('Archive');
  });
});

describe('junkMessage', () => {
  it('generates junk script', () => {
    const script = junkMessage(100);
    expect(script).toContain('message id 100');
    expect(script).toContain('junk mail');
  });
});

describe('setMessageReadStatus', () => {
  it('generates set read script', () => {
    const script = setMessageReadStatus(100, true);
    expect(script).toContain('message id 100');
    expect(script).toContain('set is read of m to true');
  });

  it('generates set unread script', () => {
    const script = setMessageReadStatus(100, false);
    expect(script).toContain('set is read of m to false');
  });
});

describe('setMessageFlag', () => {
  it('generates flagged script for status 1', () => {
    const script = setMessageFlag(100, 1);
    expect(script).toContain('flag marked');
  });

  it('generates completed script for status 2', () => {
    const script = setMessageFlag(100, 2);
    expect(script).toContain('flag complete');
  });

  it('generates not flagged script for status 0', () => {
    const script = setMessageFlag(100, 0);
    expect(script).toContain('flag not flagged');
  });
});

describe('setMessageCategories', () => {
  it('generates categories script', () => {
    const script = setMessageCategories(100, ['Important', 'Work']);
    expect(script).toContain('message id 100');
    expect(script).toContain('"Important"');
    expect(script).toContain('"Work"');
    expect(script).toContain('set category of m');
  });
});

describe('createMailFolder', () => {
  it('generates script without parent', () => {
    const script = createMailFolder('New Folder', undefined);
    expect(script).toContain('New Folder');
    expect(script).toContain('make new mail folder');
    expect(script).not.toContain('parentFolder');
  });

  it('generates script with parent folder', () => {
    const script = createMailFolder('Subfolder', 42);
    expect(script).toContain('Subfolder');
    expect(script).toContain('mail folder id 42');
    expect(script).toContain('at parentFolder');
  });
});

describe('deleteMailFolder', () => {
  it('generates delete folder script', () => {
    const script = deleteMailFolder(42);
    expect(script).toContain('mail folder id 42');
    expect(script).toContain('delete f');
  });
});

describe('renameMailFolder', () => {
  it('generates rename script', () => {
    const script = renameMailFolder(42, 'Renamed');
    expect(script).toContain('mail folder id 42');
    expect(script).toContain('set name of f to "Renamed"');
  });
});

describe('moveMailFolder', () => {
  it('generates move folder script', () => {
    const script = moveMailFolder(42, 99);
    expect(script).toContain('mail folder id 42');
    expect(script).toContain('mail folder id 99');
    expect(script).toContain('move f to targetParent');
  });
});

describe('emptyMailFolder', () => {
  it('generates empty folder script', () => {
    const script = emptyMailFolder(42);
    expect(script).toContain('mail folder id 42');
    expect(script).toContain('messages of targetFolder');
    expect(script).toContain('deleted items');
  });
});

// =============================================================================
// Attachment Scripts
// =============================================================================

describe('listAttachments', () => {
  it('generates AppleScript referencing the correct message id', () => {
    const script = listAttachments(123);
    expect(script).toContain('message id 123');
  });

  it('iterates attachments of m and collects name, file size, content type', () => {
    const script = listAttachments(123);
    expect(script).toContain('attachments of m');
    expect(script).toContain('name of a');
    expect(script).toContain('file size of a');
    expect(script).toContain('content type of a');
  });
});

describe('saveAttachment', () => {
  it('generates save command with item index and POSIX file', () => {
    const script = saveAttachment(123, 2, '/tmp/doc.pdf');
    expect(script).toContain('message id 123');
    expect(script).toContain('item 2');
    expect(script).toContain('POSIX file "/tmp/doc.pdf"');
    expect(script).toContain('save a in');
  });

  it('escapes paths with special characters', () => {
    const script = saveAttachment(123, 1, '/tmp/My "Documents"/file (1).pdf');
    expect(script).toContain('POSIX file');
    // The path should be present (escaped) in the output
    expect(script).toContain('My');
    expect(script).toContain('Documents');
    expect(script).toContain('file');
  });
});

// =============================================================================
// sendEmail with inlineImages
// =============================================================================

describe('sendEmail with inlineImages', () => {
  it('generates content id assignment for inline images', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test',
      body: '<p>Hello</p>',
      bodyType: 'html',
      inlineImages: [
        { path: '/path/to/logo.png', contentId: 'logo123' },
      ],
    });

    expect(script).toContain('content id');
    expect(script).toContain('logo123');
    expect(script).toContain('POSIX file "/path/to/logo.png"');
    expect(script).toContain('make new attachment');
  });
});

// =============================================================================
// getMessage with attachmentDetails
// =============================================================================

describe('getMessage with attachmentDetails', () => {
  it('includes attachmentDetails field in output', () => {
    const script = getMessage(456);
    expect(script).toContain('message id 456');
    expect(script).toContain('attachmentDetails{{=}}');
    expect(script).toContain('attachDetailList');
  });
});
