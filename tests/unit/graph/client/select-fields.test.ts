/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests that all $select fields in Graph API calls are valid Microsoft Graph properties.
 *
 * These tests validate the field names used in .select() calls against the
 * @microsoft/microsoft-graph-types definitions. Invalid field names cause
 * runtime errors like "Could not find a property named 'X' on type 'Y'".
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

// ---------------------------------------------------------------------------
// Valid fields per entity, derived from @microsoft/microsoft-graph-types
// ---------------------------------------------------------------------------

// OutlookItem base fields (inherited by Event, Message, Contact)
const OUTLOOK_ITEM_FIELDS = ['id', 'categories', 'changeKey', 'createdDateTime', 'lastModifiedDateTime'];

const VALID_EVENT_FIELDS = new Set([
  ...OUTLOOK_ITEM_FIELDS,
  'allowNewTimeProposals', 'attendees', 'body', 'bodyPreview',
  'cancelledOccurrences', 'end', 'hasAttachments', 'hideAttendees',
  'iCalUId', 'importance', 'isAllDay', 'isCancelled', 'isDraft',
  'isOnlineMeeting', 'isOrganizer', 'isReminderOn', 'location',
  'locations', 'onlineMeeting', 'onlineMeetingProvider', 'onlineMeetingUrl',
  'organizer', 'originalEndTimeZone', 'originalStart', 'originalStartTimeZone',
  'recurrence', 'reminderMinutesBeforeStart', 'responseRequested',
  'responseStatus', 'sensitivity', 'seriesMasterId', 'showAs',
  'start', 'subject', 'transactionId', 'type', 'webLink',
  // Navigation properties
  'attachments', 'calendar', 'exceptionOccurrences', 'extensions',
  'instances', 'multiValueExtendedProperties', 'singleValueExtendedProperties',
]);

const VALID_MESSAGE_FIELDS = new Set([
  ...OUTLOOK_ITEM_FIELDS,
  'bccRecipients', 'body', 'bodyPreview', 'ccRecipients',
  'conversationId', 'conversationIndex', 'flag', 'from',
  'hasAttachments', 'importance', 'inferenceClassification',
  'internetMessageHeaders', 'internetMessageId',
  'isDeliveryReceiptRequested', 'isDraft', 'isRead',
  'isReadReceiptRequested', 'parentFolderId', 'receivedDateTime',
  'replyTo', 'sender', 'sentDateTime', 'subject', 'toRecipients',
  'uniqueBody', 'webLink',
  // Navigation properties
  'attachments', 'extensions', 'multiValueExtendedProperties',
  'singleValueExtendedProperties',
]);

const VALID_MAIL_FOLDER_FIELDS = new Set([
  'id', 'childFolderCount', 'displayName', 'isHidden',
  'parentFolderId', 'totalItemCount', 'unreadItemCount',
  // Navigation properties
  'childFolders', 'messageRules', 'messages',
  'multiValueExtendedProperties', 'singleValueExtendedProperties',
]);

const VALID_CALENDAR_FIELDS = new Set([
  'id', 'allowedOnlineMeetingProviders', 'canEdit', 'canShare',
  'canViewPrivateItems', 'changeKey', 'color',
  'defaultOnlineMeetingProvider', 'hexColor', 'isDefaultCalendar',
  'isRemovable', 'isTallyingResponses', 'name', 'owner',
  // Navigation properties
  'calendarPermissions', 'calendarView', 'events',
  'multiValueExtendedProperties', 'singleValueExtendedProperties',
]);

const VALID_CONTACT_FIELDS = new Set([
  ...OUTLOOK_ITEM_FIELDS,
  'assistantName', 'birthday', 'businessAddress', 'businessHomePage',
  'businessPhones', 'children', 'companyName', 'department',
  'displayName', 'emailAddresses', 'fileAs', 'generation',
  'givenName', 'homeAddress', 'homePhones', 'imAddresses',
  'initials', 'jobTitle', 'manager', 'middleName', 'mobilePhone',
  'nickName', 'officeLocation', 'otherAddress', 'parentFolderId',
  'personalNotes', 'profession', 'spouseName', 'surname', 'title',
  // Navigation properties
  'extensions', 'multiValueExtendedProperties',
  'singleValueExtendedProperties', 'photo',
]);

const VALID_TODO_TASK_LIST_FIELDS = new Set([
  'id', 'displayName', 'isOwner', 'isShared', 'wellknownListName',
  // Navigation properties
  'extensions', 'tasks',
]);

const VALID_TODO_TASK_FIELDS = new Set([
  'id', 'body', 'bodyLastModifiedDateTime', 'categories',
  'completedDateTime', 'createdDateTime', 'dueDateTime',
  'hasAttachments', 'importance', 'isReminderOn',
  'lastModifiedDateTime', 'recurrence', 'reminderDateTime',
  'startDateTime', 'status', 'title',
  // Navigation properties
  'attachments', 'attachmentSessions', 'checklistItems',
  'extensions', 'linkedResources',
]);

// ---------------------------------------------------------------------------
// Capture $select calls from the Graph client
// ---------------------------------------------------------------------------

// Track all select() calls with their context
const selectCalls: Array<{ url: string; fields: string }> = [];

const createTrackingRequestBuilder = (mockResponse: any) => {
  const builder: any = {
    select: vi.fn().mockImplementation(function (this: any, fields: string) {
      // Store the select fields along with the URL that was called
      selectCalls.push({ url: builder._url ?? 'unknown', fields });
      return this;
    }),
    top: vi.fn().mockReturnThis(),
    skip: vi.fn().mockReturnThis(),
    orderby: vi.fn().mockReturnThis(),
    filter: vi.fn().mockReturnThis(),
    search: vi.fn().mockReturnThis(),
    query: vi.fn().mockReturnThis(),
    get: vi.fn().mockResolvedValue(mockResponse),
    post: vi.fn().mockResolvedValue(mockResponse),
    patch: vi.fn().mockResolvedValue(mockResponse),
    delete: vi.fn().mockResolvedValue(undefined),
    _url: null as string | null,
  };
  return builder;
};

const mockApi = vi.fn().mockImplementation((url: string) => {
  const builder = createTrackingRequestBuilder({ value: [] });
  builder._url = url;
  return builder;
});

const mockGraphClient = { api: mockApi };

vi.mock('@microsoft/microsoft-graph-client', () => ({
  Client: {
    init: vi.fn(function () { return mockGraphClient; }),
  },
}));

vi.mock('../../../../src/graph/auth/index.js', () => ({
  getAccessToken: vi.fn().mockResolvedValue('test-access-token'),
}));

vi.mock('isomorphic-fetch', () => ({ default: vi.fn() }));

import { GraphClient } from '../../../../src/graph/client/graph-client.js';

// ---------------------------------------------------------------------------
// Helper
// ---------------------------------------------------------------------------

function validateSelectFields(fields: string, validFields: Set<string>, entityName: string): void {
  const fieldList = fields.split(',');
  for (const field of fieldList) {
    expect(
      validFields.has(field),
      `Invalid $select field '${field}' for ${entityName}. ` +
      `This field does not exist in Microsoft Graph API. ` +
      `Valid fields include: ${[...validFields].sort().join(', ')}`
    ).toBe(true);
  }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('Graph API $select field validation', () => {
  let client: GraphClient;

  beforeEach(() => {
    vi.clearAllMocks();
    selectCalls.length = 0;
    client = new GraphClient();
  });

  describe('MailFolder endpoints', () => {
    it('listMailFolders uses only valid MailFolder fields', async () => {
      await client.listMailFolders();

      const folderSelects = selectCalls.filter(c =>
        c.url.includes('mailFolders') && !c.url.includes('/messages')
      );
      expect(folderSelects.length).toBeGreaterThan(0);

      for (const call of folderSelects) {
        validateSelectFields(call.fields, VALID_MAIL_FOLDER_FIELDS, 'MailFolder');
      }
    });

    it('getMailFolder uses only valid MailFolder fields', async () => {
      await client.getMailFolder('folder-1');

      const folderSelects = selectCalls.filter(c => c.url.includes('mailFolders/folder-1'));
      expect(folderSelects.length).toBeGreaterThan(0);

      for (const call of folderSelects) {
        validateSelectFields(call.fields, VALID_MAIL_FOLDER_FIELDS, 'MailFolder');
      }
    });
  });

  describe('Message endpoints', () => {
    it('listMessages uses only valid Message fields', async () => {
      await client.listMessages('folder-1');

      const msgSelects = selectCalls.filter(c => c.url.includes('/messages'));
      expect(msgSelects.length).toBeGreaterThan(0);

      for (const call of msgSelects) {
        validateSelectFields(call.fields, VALID_MESSAGE_FIELDS, 'Message');
      }
    });

    it('listUnreadMessages uses only valid Message fields', async () => {
      await client.listUnreadMessages('folder-1');

      const msgSelects = selectCalls.filter(c => c.url.includes('/messages'));
      expect(msgSelects.length).toBeGreaterThan(0);

      for (const call of msgSelects) {
        validateSelectFields(call.fields, VALID_MESSAGE_FIELDS, 'Message');
      }
    });

    it('searchMessages uses only valid Message fields', async () => {
      await client.searchMessages('test query');

      const msgSelects = selectCalls.filter(c => c.url.includes('/messages'));
      expect(msgSelects.length).toBeGreaterThan(0);

      for (const call of msgSelects) {
        validateSelectFields(call.fields, VALID_MESSAGE_FIELDS, 'Message');
      }
    });

    it('searchMessagesInFolder uses only valid Message fields', async () => {
      await client.searchMessagesInFolder('folder-1', 'test query');

      const msgSelects = selectCalls.filter(c => c.url.includes('/messages'));
      expect(msgSelects.length).toBeGreaterThan(0);

      for (const call of msgSelects) {
        validateSelectFields(call.fields, VALID_MESSAGE_FIELDS, 'Message');
      }
    });

    it('getMessage uses only valid Message fields', async () => {
      mockApi.mockImplementation((url: string) => {
        const builder = createTrackingRequestBuilder({ id: 'msg-1', subject: 'Test' });
        builder._url = url;
        return builder;
      });

      await client.getMessage('msg-1');

      const msgSelects = selectCalls.filter(c => c.url.includes('/messages/'));
      expect(msgSelects.length).toBeGreaterThan(0);

      for (const call of msgSelects) {
        validateSelectFields(call.fields, VALID_MESSAGE_FIELDS, 'Message');
      }
    });
  });

  describe('Calendar endpoints', () => {
    it('listCalendars uses only valid Calendar fields', async () => {
      await client.listCalendars();

      const calSelects = selectCalls.filter(c =>
        c.url.includes('/calendars') && !c.url.includes('calendarView')
      );
      expect(calSelects.length).toBeGreaterThan(0);

      for (const call of calSelects) {
        validateSelectFields(call.fields, VALID_CALENDAR_FIELDS, 'Calendar');
      }
    });
  });

  describe('Event endpoints', () => {
    it('listEvents without date range uses only valid Event fields', async () => {
      await client.listEvents(50);

      const eventSelects = selectCalls.filter(c => c.url.includes('/events'));
      expect(eventSelects.length).toBeGreaterThan(0);

      for (const call of eventSelects) {
        validateSelectFields(call.fields, VALID_EVENT_FIELDS, 'Event');
      }
    });

    it('listEvents with date range (calendarView) uses only valid Event fields', async () => {
      const start = new Date('2026-02-24T00:00:00Z');
      const end = new Date('2026-02-24T23:59:59Z');

      await client.listEvents(50, undefined, start, end);

      const eventSelects = selectCalls.filter(c =>
        c.url.includes('calendarView') || c.url.includes('/events')
      );
      expect(eventSelects.length).toBeGreaterThan(0);

      for (const call of eventSelects) {
        validateSelectFields(call.fields, VALID_EVENT_FIELDS, 'Event');
      }
    });

    it('listEvents with calendar ID and date range uses only valid Event fields', async () => {
      const start = new Date('2026-02-24T00:00:00Z');
      const end = new Date('2026-02-24T23:59:59Z');

      await client.listEvents(50, 'cal-1', start, end);

      const eventSelects = selectCalls.filter(c =>
        c.url.includes('calendarView') || c.url.includes('/events')
      );
      expect(eventSelects.length).toBeGreaterThan(0);

      for (const call of eventSelects) {
        validateSelectFields(call.fields, VALID_EVENT_FIELDS, 'Event');
      }
    });

    it('getEvent uses only valid Event fields', async () => {
      mockApi.mockImplementation((url: string) => {
        const builder = createTrackingRequestBuilder({ id: 'evt-1', subject: 'Meeting' });
        builder._url = url;
        return builder;
      });

      await client.getEvent('evt-1');

      const eventSelects = selectCalls.filter(c => c.url.includes('/events/'));
      expect(eventSelects.length).toBeGreaterThan(0);

      for (const call of eventSelects) {
        validateSelectFields(call.fields, VALID_EVENT_FIELDS, 'Event');
      }
    });

    it('rejects isRecurrence as an invalid Event field', () => {
      // This field caused the original GRAPH_ERROR. Ensure it's never valid.
      expect(VALID_EVENT_FIELDS.has('isRecurrence')).toBe(false);
    });
  });

  describe('Contact endpoints', () => {
    it('listContacts uses only valid Contact fields', async () => {
      await client.listContacts();

      const contactSelects = selectCalls.filter(c => c.url.includes('/contacts'));
      expect(contactSelects.length).toBeGreaterThan(0);

      for (const call of contactSelects) {
        validateSelectFields(call.fields, VALID_CONTACT_FIELDS, 'Contact');
      }
    });

    it('searchContacts uses only valid Contact fields', async () => {
      await client.searchContacts('john');

      const contactSelects = selectCalls.filter(c => c.url.includes('/contacts'));
      expect(contactSelects.length).toBeGreaterThan(0);

      for (const call of contactSelects) {
        validateSelectFields(call.fields, VALID_CONTACT_FIELDS, 'Contact');
      }
    });

    it('getContact uses only valid Contact fields', async () => {
      mockApi.mockImplementation((url: string) => {
        const builder = createTrackingRequestBuilder({ id: 'c-1', displayName: 'John' });
        builder._url = url;
        return builder;
      });

      await client.getContact('c-1');

      const contactSelects = selectCalls.filter(c => c.url.includes('/contacts/'));
      expect(contactSelects.length).toBeGreaterThan(0);

      for (const call of contactSelects) {
        validateSelectFields(call.fields, VALID_CONTACT_FIELDS, 'Contact');
      }
    });
  });

  describe('TodoTaskList endpoints', () => {
    it('listTaskLists uses only valid TodoTaskList fields', async () => {
      await client.listTaskLists();

      const listSelects = selectCalls.filter(c =>
        c.url.includes('/todo/lists') && !c.url.includes('/tasks')
      );
      expect(listSelects.length).toBeGreaterThan(0);

      for (const call of listSelects) {
        validateSelectFields(call.fields, VALID_TODO_TASK_LIST_FIELDS, 'TodoTaskList');
      }
    });
  });

  describe('TodoTask endpoints', () => {
    it('listTasks uses only valid TodoTask fields', async () => {
      await client.listTasks('list-1');

      const taskSelects = selectCalls.filter(c => c.url.includes('/tasks'));
      expect(taskSelects.length).toBeGreaterThan(0);

      for (const call of taskSelects) {
        validateSelectFields(call.fields, VALID_TODO_TASK_FIELDS, 'TodoTask');
      }
    });

    it('getTask uses only valid TodoTask fields', async () => {
      mockApi.mockImplementation((url: string) => {
        const builder = createTrackingRequestBuilder({ id: 't-1', title: 'Test Task' });
        builder._url = url;
        return builder;
      });

      await client.getTask('list-1', 'task-1');

      const taskSelects = selectCalls.filter(c => c.url.includes('/tasks/'));
      expect(taskSelects.length).toBeGreaterThan(0);

      for (const call of taskSelects) {
        validateSelectFields(call.fields, VALID_TODO_TASK_FIELDS, 'TodoTask');
      }
    });
  });
});
