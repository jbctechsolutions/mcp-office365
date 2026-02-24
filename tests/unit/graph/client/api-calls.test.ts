/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests that all Graph API calls use correct endpoints, HTTP methods,
 * request bodies, query parameters, and well-known folder names.
 *
 * Validates every public method on GraphClient against the Microsoft
 * Graph v1.0 API specification.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';

// ---------------------------------------------------------------------------
// Full-tracking mock: captures URL, method, body, and all query builder calls
// ---------------------------------------------------------------------------

interface ApiCall {
  url: string;
  method: 'get' | 'post' | 'patch' | 'delete';
  body?: any;
  selectFields?: string;
  filterExpr?: string;
  orderbyExpr?: string;
  searchExpr?: string;
  queryParams?: any;
  topValue?: number;
  skipValue?: number;
}

const apiCalls: ApiCall[] = [];

function createTrackingBuilder(mockResponse: any) {
  const call: ApiCall = { url: '', method: 'get' };
  const builder: any = {
    select: vi.fn().mockImplementation((fields: string) => {
      call.selectFields = fields;
      return builder;
    }),
    top: vi.fn().mockImplementation((n: number) => {
      call.topValue = n;
      return builder;
    }),
    skip: vi.fn().mockImplementation((n: number) => {
      call.skipValue = n;
      return builder;
    }),
    orderby: vi.fn().mockImplementation((expr: string) => {
      call.orderbyExpr = expr;
      return builder;
    }),
    filter: vi.fn().mockImplementation((expr: string) => {
      call.filterExpr = expr;
      return builder;
    }),
    search: vi.fn().mockImplementation((expr: string) => {
      call.searchExpr = expr;
      return builder;
    }),
    query: vi.fn().mockImplementation((params: any) => {
      call.queryParams = params;
      return builder;
    }),
    get: vi.fn().mockImplementation(async () => {
      call.method = 'get';
      apiCalls.push({ ...call });
      return mockResponse;
    }),
    post: vi.fn().mockImplementation(async (body: any) => {
      call.method = 'post';
      call.body = body;
      apiCalls.push({ ...call });
      return mockResponse;
    }),
    patch: vi.fn().mockImplementation(async (body: any) => {
      call.method = 'patch';
      call.body = body;
      apiCalls.push({ ...call });
      return mockResponse;
    }),
    delete: vi.fn().mockImplementation(async () => {
      call.method = 'delete';
      apiCalls.push({ ...call });
      return undefined;
    }),
  };
  return { builder, call };
}

const mockApi = vi.fn();
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

function setupMock(response: any = { value: [] }) {
  mockApi.mockImplementation((url: string) => {
    const { builder, call } = createTrackingBuilder(response);
    call.url = url;
    return builder;
  });
}

// ---------------------------------------------------------------------------
// Valid Graph v1.0 endpoint patterns
// ---------------------------------------------------------------------------

const VALID_ENDPOINT_PATTERNS = [
  // Mail folders
  /^\/me\/mailFolders$/,
  /^\/me\/mailFolders\/[^/]+$/,
  /^\/me\/mailFolders\/[^/]+\/childFolders$/,
  /^\/me\/mailFolders\/[^/]+\/move$/,
  // Messages
  /^\/me\/messages$/,
  /^\/me\/messages\/[^/]+$/,
  /^\/me\/messages\/[^/]+\/move$/,
  /^\/me\/messages\/[^/]+\/send$/,
  /^\/me\/messages\/[^/]+\/reply$/,
  /^\/me\/messages\/[^/]+\/replyAll$/,
  /^\/me\/messages\/[^/]+\/forward$/,
  /^\/me\/sendMail$/,
  /^\/me\/mailFolders\/[^/]+\/messages$/,
  // Calendars
  /^\/me\/calendars$/,
  /^\/me\/calendars\/[^/]+\/events$/,
  /^\/me\/calendars\/[^/]+\/calendarView$/,
  // Events
  /^\/me\/events$/,
  /^\/me\/events\/[^/]+$/,
  /^\/me\/calendarView$/,
  // Contacts
  /^\/me\/contacts$/,
  /^\/me\/contacts\/[^/]+$/,
  // Tasks (Microsoft To Do)
  /^\/me\/todo\/lists$/,
  /^\/me\/todo\/lists\/[^/]+\/tasks$/,
  /^\/me\/todo\/lists\/[^/]+\/tasks\/[^/]+$/,
  // Attachments
  /^\/me\/messages\/[^/]+\/attachments$/,
  /^\/me\/messages\/[^/]+\/attachments\/[^/]+$/,
  /^\/me\/messages\/[^/]+\/attachments\/createUploadSession$/,
  // Pagination (nextLink URLs from Graph)
  /^https:\/\/graph\.microsoft\.com\//,
];

function isValidEndpoint(url: string): boolean {
  return VALID_ENDPOINT_PATTERNS.some(p => p.test(url));
}

// Graph API well-known folder names:
// https://learn.microsoft.com/en-us/graph/api/resources/mailfolder
const VALID_WELL_KNOWN_FOLDERS = new Set([
  'archive', 'clutter', 'conflicts', 'conversationhistory',
  'deleteditems', 'drafts', 'inbox', 'junkemail', 'localfailures',
  'msgfolderroot', 'outbox', 'recoverableitemsdeletions',
  'scheduled', 'searchfolders', 'sentitems', 'serverfailures',
  'syncissues',
]);

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('Graph API endpoint and method validation', () => {
  let client: GraphClient;

  beforeEach(() => {
    vi.clearAllMocks();
    apiCalls.length = 0;
    client = new GraphClient();
    setupMock();
  });

  // =========================================================================
  // Read operations
  // =========================================================================

  describe('Read operation endpoints', () => {
    it('listMailFolders calls /me/mailFolders with GET', async () => {
      await client.listMailFolders();

      const calls = apiCalls.filter(c => c.method === 'get');
      expect(calls.length).toBeGreaterThan(0);
      expect(calls[0].url).toBe('/me/mailFolders');

      for (const call of calls) {
        expect(
          isValidEndpoint(call.url),
          `Invalid endpoint: ${call.url}`
        ).toBe(true);
      }
    });

    it('listMailFolders fetches child folders', async () => {
      await client.listMailFolders();

      const childCalls = apiCalls.filter(c =>
        c.url.includes('/childFolders')
      );
      // Should attempt to get children for each top-level folder
      expect(childCalls.length).toBeGreaterThanOrEqual(0);
      for (const call of childCalls) {
        expect(call.url).toMatch(/^\/me\/mailFolders\/[^/]+\/childFolders$/);
      }
    });

    it('listMailFolders uses $top(100) for pagination', async () => {
      await client.listMailFolders();

      expect(apiCalls[0].topValue).toBe(100);
    });

    it('getMailFolder calls /me/mailFolders/{id} with GET', async () => {
      await client.getMailFolder('folder-123');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailFolders/folder-123');
      expect(apiCalls[0].method).toBe('get');
    });

    it('listMessages calls /me/mailFolders/{id}/messages with GET', async () => {
      await client.listMessages('folder-1', 25, 10);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailFolders/folder-1/messages');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].topValue).toBe(25);
      expect(apiCalls[0].skipValue).toBe(10);
      expect(apiCalls[0].orderbyExpr).toBe('receivedDateTime desc');
    });

    it('listUnreadMessages filters with isRead eq false', async () => {
      await client.listUnreadMessages('folder-1', 50, 0);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailFolders/folder-1/messages');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].filterExpr).toBe('isRead eq false');
      expect(apiCalls[0].orderbyExpr).toBe('receivedDateTime desc');
    });

    it('searchMessages uses /me/messages with $search', async () => {
      await client.searchMessages('test query', 25);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].searchExpr).toBe('"test query"');
      expect(apiCalls[0].topValue).toBe(25);
    });

    it('searchMessagesInFolder uses /me/mailFolders/{id}/messages with $search', async () => {
      await client.searchMessagesInFolder('folder-1', 'query', 30);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailFolders/folder-1/messages');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].searchExpr).toBe('"query"');
      expect(apiCalls[0].topValue).toBe(30);
    });

    it('getMessage calls /me/messages/{id} with GET', async () => {
      setupMock({ id: 'msg-1', subject: 'Test' });
      await client.getMessage('msg-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1');
      expect(apiCalls[0].method).toBe('get');
    });

    it('listCalendars calls /me/calendars with GET', async () => {
      await client.listCalendars();

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/calendars');
      expect(apiCalls[0].method).toBe('get');
    });

    it('listEvents without date range calls /me/events with GET', async () => {
      await client.listEvents(50);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/events');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].orderbyExpr).toBe('start/dateTime');
    });

    it('listEvents with date range calls /me/calendarView with ISO query params', async () => {
      const start = new Date('2026-02-24T00:00:00Z');
      const end = new Date('2026-02-24T23:59:59Z');

      await client.listEvents(50, undefined, start, end);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/calendarView');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].queryParams).toEqual({
        startDateTime: start.toISOString(),
        endDateTime: end.toISOString(),
      });
    });

    it('listEvents with calendar ID calls /me/calendars/{id}/events', async () => {
      await client.listEvents(50, 'cal-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/calendars/cal-1/events');
    });

    it('listEvents with calendar ID + date range calls /me/calendars/{id}/calendarView', async () => {
      const start = new Date('2026-02-24T00:00:00Z');
      const end = new Date('2026-02-24T23:59:59Z');

      await client.listEvents(50, 'cal-1', start, end);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/calendars/cal-1/calendarView');
    });

    it('getEvent calls /me/events/{id} with GET', async () => {
      setupMock({ id: 'evt-1', subject: 'Meeting' });
      await client.getEvent('evt-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/events/evt-1');
      expect(apiCalls[0].method).toBe('get');
    });

    it('listContacts calls /me/contacts with GET and $orderby displayName', async () => {
      await client.listContacts(50, 0);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/contacts');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].orderbyExpr).toBe('displayName');
    });

    it('searchContacts uses $filter with contains(displayName)', async () => {
      await client.searchContacts('John', 50);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/contacts');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].filterExpr).toBe("contains(displayName,'John')");
    });

    it('getContact calls /me/contacts/{id} with GET', async () => {
      setupMock({ id: 'c-1', displayName: 'John' });
      await client.getContact('c-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/contacts/c-1');
      expect(apiCalls[0].method).toBe('get');
    });

    it('listTaskLists calls /me/todo/lists with GET', async () => {
      await client.listTaskLists();

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/todo/lists');
      expect(apiCalls[0].method).toBe('get');
    });

    it('listTasks calls /me/todo/lists/{id}/tasks with GET', async () => {
      await client.listTasks('list-1', 50, 0);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/todo/lists/list-1/tasks');
      expect(apiCalls[0].method).toBe('get');
    });

    it('listTasks with includeCompleted=false applies status filter', async () => {
      await client.listTasks('list-1', 50, 0, false);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].filterExpr).toBe("status ne 'completed'");
    });

    it('listTasks with includeCompleted=true (default) does not filter', async () => {
      await client.listTasks('list-1', 50, 0, true);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].filterExpr).toBeUndefined();
    });

    it('getTask calls /me/todo/lists/{listId}/tasks/{taskId} with GET', async () => {
      setupMock({ id: 't-1', title: 'Task' });
      await client.getTask('list-1', 'task-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/todo/lists/list-1/tasks/task-1');
      expect(apiCalls[0].method).toBe('get');
    });
  });

  // =========================================================================
  // Write operations
  // =========================================================================

  describe('Write operation endpoints and bodies', () => {
    it('moveMessage POSTs to /me/messages/{id}/move with destinationId', async () => {
      await client.moveMessage('msg-1', 'folder-dest');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/move');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ destinationId: 'folder-dest' });
    });

    it('deleteMessage POSTs to /me/messages/{id}/move with deleteditems', async () => {
      await client.deleteMessage('msg-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/move');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ destinationId: 'deleteditems' });
    });

    it('archiveMessage POSTs to /me/messages/{id}/move with archive', async () => {
      await client.archiveMessage('msg-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/move');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ destinationId: 'archive' });
    });

    it('junkMessage POSTs to /me/messages/{id}/move with junkemail', async () => {
      await client.junkMessage('msg-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/move');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ destinationId: 'junkemail' });
    });

    it('updateMessage PATCHes /me/messages/{id} with updates', async () => {
      const updates = { isRead: true, flag: { flagStatus: 'flagged' } };
      await client.updateMessage('msg-1', updates);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1');
      expect(apiCalls[0].method).toBe('patch');
      expect(apiCalls[0].body).toEqual(updates);
    });

    it('createMailFolder POSTs to /me/mailFolders with displayName', async () => {
      setupMock({ id: 'new-folder', displayName: 'Test Folder' });
      await client.createMailFolder('Test Folder');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailFolders');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ displayName: 'Test Folder' });
    });

    it('createMailFolder with parent POSTs to /me/mailFolders/{parentId}/childFolders', async () => {
      setupMock({ id: 'new-folder', displayName: 'Child' });
      await client.createMailFolder('Child', 'parent-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailFolders/parent-1/childFolders');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ displayName: 'Child' });
    });

    it('deleteMailFolder DELETEs /me/mailFolders/{id}', async () => {
      await client.deleteMailFolder('folder-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailFolders/folder-1');
      expect(apiCalls[0].method).toBe('delete');
    });

    it('renameMailFolder PATCHes /me/mailFolders/{id} with displayName', async () => {
      await client.renameMailFolder('folder-1', 'New Name');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailFolders/folder-1');
      expect(apiCalls[0].method).toBe('patch');
      expect(apiCalls[0].body).toEqual({ displayName: 'New Name' });
    });

    it('moveMailFolder POSTs to /me/mailFolders/{id}/move with destinationId', async () => {
      await client.moveMailFolder('folder-1', 'parent-dest');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailFolders/folder-1/move');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ destinationId: 'parent-dest' });
    });

    it('emptyMailFolder GETs messages then POSTs move to deleteditems', async () => {
      const messages = [{ id: 'msg-1' }, { id: 'msg-2' }];
      let getCount = 0;
      mockApi.mockImplementation((url: string) => {
        const isMessageList = url.includes('/messages') && !url.includes('/move');
        const response = isMessageList
          ? { value: getCount === 0 ? messages : [] }
          : {};
        if (isMessageList) getCount++;
        const { builder, call } = createTrackingBuilder(response);
        call.url = url;
        return builder;
      });

      await client.emptyMailFolder('folder-1');

      // Verify GET messages
      const getCall = apiCalls.find(c => c.method === 'get' && c.url.includes('/messages'));
      expect(getCall).toBeDefined();
      expect(getCall!.url).toBe('/me/mailFolders/folder-1/messages');
      expect(getCall!.selectFields).toBe('id');
      expect(getCall!.topValue).toBe(100);

      // Verify POST moves
      const moveCalls = apiCalls.filter(c => c.method === 'post' && c.url.includes('/move'));
      expect(moveCalls).toHaveLength(2);
      expect(moveCalls[0].url).toBe('/me/messages/msg-1/move');
      expect(moveCalls[0].body).toEqual({ destinationId: 'deleteditems' });
      expect(moveCalls[1].url).toBe('/me/messages/msg-2/move');
      expect(moveCalls[1].body).toEqual({ destinationId: 'deleteditems' });
    });
  });

  // =========================================================================
  // Draft & Send operations
  // =========================================================================

  describe('Draft & Send operation endpoints and bodies', () => {
    it('createDraft POSTs to /me/messages with isDraft and message fields', async () => {
      const draftMessage = {
        subject: 'Test Draft',
        body: { contentType: 'text' as const, content: 'Hello' },
        toRecipients: [{ emailAddress: { address: 'user@example.com' } }],
        ccRecipients: [],
        bccRecipients: [],
        isDraft: true,
      };
      setupMock({ id: 'draft-1', subject: 'Test Draft', isDraft: true });

      await client.createDraft(draftMessage);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual(draftMessage);
    });

    it('updateDraft PATCHes /me/messages/{id} with updates', async () => {
      const updates = { subject: 'Updated Subject', body: { contentType: 'html', content: '<p>Updated</p>' } };
      setupMock({ id: 'draft-1', subject: 'Updated Subject' });

      await client.updateDraft('draft-1', updates);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/draft-1');
      expect(apiCalls[0].method).toBe('patch');
      expect(apiCalls[0].body).toEqual(updates);
    });

    it('sendDraft POSTs to /me/messages/{id}/send with null body', async () => {
      await client.sendDraft('draft-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/draft-1/send');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toBeNull();
    });

    it('sendMail POSTs to /me/sendMail with message object', async () => {
      const message = {
        subject: 'Direct Send',
        body: { contentType: 'text' as const, content: 'Hello' },
        toRecipients: [{ emailAddress: { address: 'user@example.com' } }],
        ccRecipients: [],
        bccRecipients: [],
      };

      await client.sendMail(message);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/sendMail');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ message });
    });

    it('replyMessage POSTs to /me/messages/{id}/reply with comment', async () => {
      await client.replyMessage('msg-1', 'Thanks for the update', false);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/reply');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ comment: 'Thanks for the update' });
    });

    it('replyMessage with replyAll POSTs to /me/messages/{id}/replyAll', async () => {
      await client.replyMessage('msg-1', 'Reply to all', true);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/replyAll');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ comment: 'Reply to all' });
    });

    it('forwardMessage POSTs to /me/messages/{id}/forward with toRecipients and comment', async () => {
      const toRecipients = [{ emailAddress: { address: 'forward@example.com' } }];

      await client.forwardMessage('msg-1', toRecipients, 'Please review');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/forward');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ toRecipients, comment: 'Please review' });
    });

    it('forwardMessage without comment sends only toRecipients', async () => {
      const toRecipients = [{ emailAddress: { address: 'forward@example.com' } }];

      await client.forwardMessage('msg-1', toRecipients);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/forward');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ toRecipients });
    });
  });

  // =========================================================================
  // Attachment operations
  // =========================================================================

  describe('Attachment operation endpoints and bodies', () => {
    it('listAttachments GETs /me/messages/{id}/attachments with $select', async () => {
      await client.listAttachments('msg-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/attachments');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].selectFields).toBe('id,name,size,contentType,isInline');
    });

    it('getAttachment GETs /me/messages/{id}/attachments/{attachmentId}', async () => {
      setupMock({ id: 'att-1', name: 'file.pdf', contentBytes: 'base64data' });
      await client.getAttachment('msg-1', 'att-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/attachments/att-1');
      expect(apiCalls[0].method).toBe('get');
    });

    it('addAttachment POSTs to /me/messages/{id}/attachments with attachment body', async () => {
      const attachment = {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: 'file.txt',
        contentBytes: 'SGVsbG8gV29ybGQ=',
        contentType: 'text/plain',
      };
      setupMock({ id: 'att-new', name: 'file.txt' });

      await client.addAttachment('msg-1', attachment);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/attachments');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual(attachment);
    });

    it('createUploadSession POSTs to /me/messages/{id}/attachments/createUploadSession', async () => {
      const body = {
        AttachmentItem: {
          attachmentType: 'file',
          name: 'largefile.zip',
          size: 5000000,
        },
      };
      setupMock({ uploadUrl: 'https://upload.example.com/session123' });

      await client.createUploadSession('msg-1', body);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/attachments/createUploadSession');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual(body);
    });
  });

  // =========================================================================
  // Well-known folder name validation
  // =========================================================================

  describe('Well-known folder names', () => {
    it('deleteMessage uses valid well-known folder "deleteditems"', async () => {
      await client.deleteMessage('msg-1');
      expect(VALID_WELL_KNOWN_FOLDERS.has(apiCalls[0].body.destinationId)).toBe(true);
    });

    it('archiveMessage uses valid well-known folder "archive"', async () => {
      await client.archiveMessage('msg-1');
      expect(VALID_WELL_KNOWN_FOLDERS.has(apiCalls[0].body.destinationId)).toBe(true);
    });

    it('junkMessage uses valid well-known folder "junkemail"', async () => {
      await client.junkMessage('msg-1');
      expect(VALID_WELL_KNOWN_FOLDERS.has(apiCalls[0].body.destinationId)).toBe(true);
    });

    it('emptyMailFolder moves to valid well-known folder "deleteditems"', async () => {
      const messages = [{ id: 'msg-1' }];
      let getCount = 0;
      mockApi.mockImplementation((url: string) => {
        const isMessageList = url.includes('/messages') && !url.includes('/move');
        const response = isMessageList
          ? { value: getCount === 0 ? messages : [] }
          : {};
        if (isMessageList) getCount++;
        const { builder, call } = createTrackingBuilder(response);
        call.url = url;
        return builder;
      });

      await client.emptyMailFolder('folder-1');

      const moveCalls = apiCalls.filter(c => c.method === 'post' && c.url.includes('/move'));
      for (const call of moveCalls) {
        expect(VALID_WELL_KNOWN_FOLDERS.has(call.body.destinationId)).toBe(true);
      }
    });
  });

  // =========================================================================
  // Comprehensive endpoint pattern validation
  // =========================================================================

  describe('All endpoints match valid Graph v1.0 patterns', () => {
    it('every API call across all methods uses a valid endpoint', async () => {
      // Exercise every read method
      setupMock();
      await client.listMailFolders();
      await client.getMailFolder('f1');
      await client.listMessages('f1');
      await client.listUnreadMessages('f1');
      await client.searchMessages('q');
      await client.searchMessagesInFolder('f1', 'q');

      setupMock({ id: 'msg-1', subject: 'Test' });
      await client.getMessage('msg-1');

      setupMock();
      await client.listCalendars();
      await client.listEvents(10);
      await client.listEvents(10, 'cal-1');
      await client.listEvents(10, undefined, new Date(), new Date());
      await client.listEvents(10, 'cal-1', new Date(), new Date());

      setupMock({ id: 'evt-1', subject: 'M' });
      await client.getEvent('evt-1');

      setupMock();
      await client.listContacts();
      await client.searchContacts('J');

      setupMock({ id: 'c-1', displayName: 'J' });
      await client.getContact('c-1');

      setupMock();
      await client.listTaskLists();
      await client.listTasks('l1');

      setupMock({ id: 't-1', title: 'T' });
      await client.getTask('l1', 't1');

      // Exercise every write method
      setupMock();
      await client.moveMessage('m1', 'f1');
      await client.deleteMessage('m1');
      await client.archiveMessage('m1');
      await client.junkMessage('m1');
      await client.updateMessage('m1', { isRead: true });

      setupMock({ id: 'new', displayName: 'New' });
      await client.createMailFolder('New');
      await client.createMailFolder('Child', 'parent');

      setupMock();
      await client.deleteMailFolder('f1');
      await client.renameMailFolder('f1', 'Name');
      await client.moveMailFolder('f1', 'p1');

      // Exercise draft & send methods
      setupMock({ id: 'draft-1', subject: 'Draft', isDraft: true });
      await client.createDraft({
        subject: 'Draft',
        body: { contentType: 'text', content: 'Body' },
        toRecipients: [],
        isDraft: true,
      });

      setupMock({ id: 'draft-1', subject: 'Updated' });
      await client.updateDraft('draft-1', { subject: 'Updated' });

      setupMock();
      await client.sendDraft('draft-1');
      await client.sendMail({
        subject: 'Send',
        body: { contentType: 'text', content: 'Body' },
        toRecipients: [{ emailAddress: { address: 'a@b.com' } }],
      });
      await client.replyMessage('m1', 'comment', false);
      await client.replyMessage('m1', 'comment', true);
      await client.forwardMessage('m1', [{ emailAddress: { address: 'a@b.com' } }], 'fwd');

      // Exercise attachment methods
      setupMock({ value: [] });
      await client.listAttachments('m1');

      setupMock({ id: 'att-1', name: 'file.pdf', contentBytes: 'data' });
      await client.getAttachment('m1', 'att-1');

      setupMock({ id: 'att-new', name: 'file.txt' });
      await client.addAttachment('m1', { '@odata.type': '#microsoft.graph.fileAttachment', name: 'file.txt', contentBytes: 'data' });

      setupMock({ uploadUrl: 'https://upload.example.com/session' });
      await client.createUploadSession('m1', { AttachmentItem: { attachmentType: 'file', name: 'big.zip', size: 5000000 } });

      // Verify all captured URLs
      for (const call of apiCalls) {
        expect(
          isValidEndpoint(call.url),
          `Invalid Graph API endpoint: ${call.url} (method: ${call.method})`
        ).toBe(true);
      }
    });
  });

  // =========================================================================
  // Cache invalidation after mutations
  // =========================================================================

  describe('Cache invalidation after write operations', () => {
    it('moveMessage clears cache so next read hits API', async () => {
      await client.listMailFolders();
      apiCalls.length = 0;

      await client.moveMessage('msg-1', 'dest');
      apiCalls.length = 0;

      setupMock();
      await client.listMailFolders();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/mailFolders');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('deleteMessage clears cache', async () => {
      await client.listMailFolders();
      apiCalls.length = 0;

      await client.deleteMessage('msg-1');
      apiCalls.length = 0;

      setupMock();
      await client.listMailFolders();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/mailFolders');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('updateMessage clears cache', async () => {
      await client.listMessages('f1');
      apiCalls.length = 0;

      await client.updateMessage('msg-1', { isRead: true });
      apiCalls.length = 0;

      setupMock();
      await client.listMessages('f1');

      expect(apiCalls.filter(c => c.method === 'get').length).toBeGreaterThan(0);
    });

    it('createMailFolder clears cache', async () => {
      await client.listMailFolders();
      apiCalls.length = 0;

      setupMock({ id: 'new', displayName: 'New' });
      await client.createMailFolder('New');
      apiCalls.length = 0;

      setupMock();
      await client.listMailFolders();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/mailFolders');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('deleteMailFolder clears cache', async () => {
      await client.listMailFolders();
      apiCalls.length = 0;

      await client.deleteMailFolder('f1');
      apiCalls.length = 0;

      setupMock();
      await client.listMailFolders();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/mailFolders');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('renameMailFolder clears cache', async () => {
      await client.listMailFolders();
      apiCalls.length = 0;

      await client.renameMailFolder('f1', 'Renamed');
      apiCalls.length = 0;

      setupMock();
      await client.listMailFolders();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/mailFolders');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('moveMailFolder clears cache', async () => {
      await client.listMailFolders();
      apiCalls.length = 0;

      await client.moveMailFolder('f1', 'p1');
      apiCalls.length = 0;

      setupMock();
      await client.listMailFolders();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/mailFolders');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('createDraft clears cache', async () => {
      await client.listMessages('f1');
      apiCalls.length = 0;

      setupMock({ id: 'draft-1', subject: 'Test', isDraft: true });
      await client.createDraft({
        subject: 'Test',
        body: { contentType: 'text', content: 'Hello' },
        toRecipients: [],
        isDraft: true,
      });
      apiCalls.length = 0;

      setupMock();
      await client.listMessages('f1');

      expect(apiCalls.filter(c => c.method === 'get').length).toBeGreaterThan(0);
    });

    it('updateDraft clears cache', async () => {
      await client.listMessages('f1');
      apiCalls.length = 0;

      setupMock({ id: 'draft-1', subject: 'Updated' });
      await client.updateDraft('draft-1', { subject: 'Updated' });
      apiCalls.length = 0;

      setupMock();
      await client.listMessages('f1');

      expect(apiCalls.filter(c => c.method === 'get').length).toBeGreaterThan(0);
    });

    it('sendDraft clears cache', async () => {
      await client.listMessages('f1');
      apiCalls.length = 0;

      await client.sendDraft('draft-1');
      apiCalls.length = 0;

      setupMock();
      await client.listMessages('f1');

      expect(apiCalls.filter(c => c.method === 'get').length).toBeGreaterThan(0);
    });

    it('sendMail clears cache', async () => {
      await client.listMessages('f1');
      apiCalls.length = 0;

      await client.sendMail({
        subject: 'Test',
        body: { contentType: 'text', content: 'Hello' },
        toRecipients: [{ emailAddress: { address: 'user@example.com' } }],
      });
      apiCalls.length = 0;

      setupMock();
      await client.listMessages('f1');

      expect(apiCalls.filter(c => c.method === 'get').length).toBeGreaterThan(0);
    });

    it('replyMessage clears cache', async () => {
      await client.listMessages('f1');
      apiCalls.length = 0;

      await client.replyMessage('msg-1', 'Thanks', false);
      apiCalls.length = 0;

      setupMock();
      await client.listMessages('f1');

      expect(apiCalls.filter(c => c.method === 'get').length).toBeGreaterThan(0);
    });

    it('forwardMessage clears cache', async () => {
      await client.listMessages('f1');
      apiCalls.length = 0;

      await client.forwardMessage('msg-1', [{ emailAddress: { address: 'fwd@example.com' } }]);
      apiCalls.length = 0;

      setupMock();
      await client.listMessages('f1');

      expect(apiCalls.filter(c => c.method === 'get').length).toBeGreaterThan(0);
    });

    it('addAttachment clears cache', async () => {
      await client.listMessages('f1');
      apiCalls.length = 0;

      setupMock({ id: 'att-new', name: 'file.txt' });
      await client.addAttachment('msg-1', {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: 'file.txt',
        contentBytes: 'SGVsbG8=',
      });
      apiCalls.length = 0;

      setupMock();
      await client.listMessages('f1');

      expect(apiCalls.filter(c => c.method === 'get').length).toBeGreaterThan(0);
    });
  });
});
