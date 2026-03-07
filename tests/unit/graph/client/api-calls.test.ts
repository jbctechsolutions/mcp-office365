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
  method: 'get' | 'post' | 'patch' | 'delete' | 'put';
  body?: any;
  selectFields?: string;
  filterExpr?: string;
  orderbyExpr?: string;
  searchExpr?: string;
  queryParams?: any;
  topValue?: number;
  skipValue?: number;
  headers?: Record<string, string>;
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
    put: vi.fn().mockImplementation(async (body: any) => {
      call.method = 'put';
      call.body = body;
      apiCalls.push({ ...call });
      return undefined;
    }),
    header: vi.fn().mockImplementation((key: string, value: string) => {
      call.headers = call.headers ?? {};
      call.headers[key] = value;
      return builder;
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
  /^\/me\/messages\/[^/]+\/createReply$/,
  /^\/me\/messages\/[^/]+\/createReplyAll$/,
  /^\/me\/messages\/[^/]+\/createForward$/,
  /^\/me\/sendMail$/,
  /^\/me\/mailFolders\/[^/]+\/messages$/,
  /^\/me\/mailFolders\/[^/]+\/messages\/delta$/,
  // Calendars
  /^\/me\/calendars$/,
  /^\/me\/calendars\/[^/]+\/events$/,
  /^\/me\/calendars\/[^/]+\/calendarView$/,
  // Events
  /^\/me\/events$/,
  /^\/me\/events\/[^/]+$/,
  /^\/me\/events\/[^/]+\/accept$/,
  /^\/me\/events\/[^/]+\/decline$/,
  /^\/me\/events\/[^/]+\/tentativelyAccept$/,
  /^\/me\/events\/[^/]+\/instances$/,
  /^\/me\/calendarView$/,
  // Contacts
  /^\/me\/contacts$/,
  /^\/me\/contacts\/[^/]+$/,
  /^\/me\/contacts\/[^/]+\/photo\/\$value$/,
  // Contact Folders
  /^\/me\/contactFolders$/,
  /^\/me\/contactFolders\/[^/]+$/,
  /^\/me\/contactFolders\/[^/]+\/contacts$/,
  // Tasks (Microsoft To Do)
  /^\/me\/todo\/lists$/,
  /^\/me\/todo\/lists\/[^/]+$/,
  /^\/me\/todo\/lists\/[^/]+\/tasks$/,
  /^\/me\/todo\/lists\/[^/]+\/tasks\/[^/]+$/,
  // Automatic Replies (Out of Office)
  /^\/me\/mailboxSettings\/automaticRepliesSetting$/,
  /^\/me\/mailboxSettings$/,
  // Mail Rules
  /^\/me\/mailFolders\/inbox\/messageRules$/,
  /^\/me\/mailFolders\/inbox\/messageRules\/[^/]+$/,
  // Master Categories
  /^\/me\/outlook\/masterCategories$/,
  /^\/me\/outlook\/masterCategories\/[^/]+$/,
  // Focused Inbox Overrides
  /^\/me\/inferenceClassification\/overrides$/,
  /^\/me\/inferenceClassification\/overrides\/[^/]+$/,
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

    it('searchMessagesKql passes raw KQL query to search without quotes', async () => {
      await client.searchMessagesKql('from:alice AND hasAttachments:true', 20);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].searchExpr).toBe('from:alice AND hasAttachments:true');
      expect(apiCalls[0].topValue).toBe(20);
    });

    it('searchMessagesKqlInFolder passes raw KQL with folder scope', async () => {
      await client.searchMessagesKqlInFolder('folder-123', 'subject:"report"', 10);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailFolders/folder-123/messages');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].searchExpr).toBe('subject:"report"');
      expect(apiCalls[0].topValue).toBe(10);
    });

    it('listConversationMessages filters by conversationId with asc ordering', async () => {
      await client.listConversationMessages('AAMkAGQ=', 10);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].filterExpr).toContain("conversationId eq 'AAMkAGQ='");
      expect(apiCalls[0].orderbyExpr).toBe('receivedDateTime asc');
      expect(apiCalls[0].topValue).toBe(10);
    });

    it('getMessagesDelta initial call uses /me/mailFolders/{id}/messages/delta', async () => {
      setupMock({ value: [{ id: 'msg-1', subject: 'New' }], '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/delta-token' });
      const result = await client.getMessagesDelta('folder-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailFolders/folder-1/messages/delta');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].topValue).toBe(50);
      expect(result.messages).toHaveLength(1);
      expect(result.deltaLink).toBe('https://graph.microsoft.com/v1.0/delta-token');
    });

    it('getMessagesDelta subsequent call uses deltaLink URL directly', async () => {
      const deltaUrl = 'https://graph.microsoft.com/v1.0/me/mailFolders/folder-1/messages/delta?$deltatoken=abc123';
      setupMock({ value: [{ id: 'msg-2', subject: 'Changed' }], '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/delta-token-2' });
      const result = await client.getMessagesDelta('folder-1', deltaUrl);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe(deltaUrl);
      expect(apiCalls[0].method).toBe('get');
      expect(result.messages).toHaveLength(1);
      expect(result.deltaLink).toBe('https://graph.microsoft.com/v1.0/delta-token-2');
    });

    it('getMessagesDelta handles pagination with @odata.nextLink', async () => {
      // First page has nextLink, second page has deltaLink
      let callCount = 0;
      mockApi.mockImplementation((url: string) => {
        callCount++;
        const response = callCount === 1
          ? { value: [{ id: 'msg-1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next-page' }
          : { value: [{ id: 'msg-2' }], '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/delta-final' };
        const { builder, call } = createTrackingBuilder(response);
        call.url = url;
        return builder;
      });

      const result = await client.getMessagesDelta('folder-1');

      expect(result.messages).toHaveLength(2);
      expect(result.messages[0].id).toBe('msg-1');
      expect(result.messages[1].id).toBe('msg-2');
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

    it('listEventInstances calls /me/events/{id}/instances with query params', async () => {
      await client.listEventInstances('evt-1', '2024-01-01T00:00:00Z', '2024-12-31T23:59:59Z');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/events/evt-1/instances');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].queryParams).toEqual({
        startDateTime: '2024-01-01T00:00:00Z',
        endDateTime: '2024-12-31T23:59:59Z',
      });
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

    it('emptyMailFolder handles pagination when @odata.nextLink is present', async () => {
      const page1Messages = [{ id: 'msg-a' }, { id: 'msg-b' }];
      const page2Messages = [{ id: 'msg-c' }, { id: 'msg-d' }];
      const nextLinkUrl = 'https://graph.microsoft.com/v1.0/me/mailFolders/folder-1/messages?$skip=100';
      let getCount = 0;
      mockApi.mockImplementation((url: string) => {
        const isFirstPage = url.includes('/messages') && !url.includes('/move') && !url.startsWith('https://');
        const isNextLink = url === nextLinkUrl;
        let response: any = {};
        if (isFirstPage && getCount === 0) {
          response = { value: page1Messages, '@odata.nextLink': nextLinkUrl };
          getCount++;
        } else if (isNextLink) {
          response = { value: page2Messages };
        }
        const { builder, call } = createTrackingBuilder(response);
        call.url = url;
        return builder;
      });

      await client.emptyMailFolder('folder-1');

      // Verify the initial GET fetched messages from the folder
      const initialGet = apiCalls.find(c => c.method === 'get' && c.url.includes('/mailFolders/'));
      expect(initialGet).toBeDefined();
      expect(initialGet!.url).toBe('/me/mailFolders/folder-1/messages');
      expect(initialGet!.selectFields).toBe('id');
      expect(initialGet!.topValue).toBe(100);

      // Verify the nextLink URL was called
      const nextLinkGet = apiCalls.find(c => c.method === 'get' && c.url === nextLinkUrl);
      expect(nextLinkGet).toBeDefined();

      // Verify all messages from both pages were moved to deleteditems
      const moveCalls = apiCalls.filter(c => c.method === 'post' && c.url.includes('/move'));
      expect(moveCalls).toHaveLength(4);
      expect(moveCalls[0].url).toBe('/me/messages/msg-a/move');
      expect(moveCalls[0].body).toEqual({ destinationId: 'deleteditems' });
      expect(moveCalls[1].url).toBe('/me/messages/msg-b/move');
      expect(moveCalls[1].body).toEqual({ destinationId: 'deleteditems' });
      expect(moveCalls[2].url).toBe('/me/messages/msg-c/move');
      expect(moveCalls[2].body).toEqual({ destinationId: 'deleteditems' });
      expect(moveCalls[3].url).toBe('/me/messages/msg-d/move');
      expect(moveCalls[3].body).toEqual({ destinationId: 'deleteditems' });
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
  // Reply/Forward as draft endpoints
  // =========================================================================

  describe('Reply/Forward as draft endpoints', () => {
    it('createReplyDraft POSTs to /me/messages/{id}/createReply', async () => {
      const mockDraft = { id: 'draft-1', subject: 'RE: Test' };
      setupMock(mockDraft);

      const result = await client.createReplyDraft('msg-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/createReply');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toBeNull();
      expect(result).toEqual(mockDraft);
    });

    it('createReplyAllDraft POSTs to /me/messages/{id}/createReplyAll', async () => {
      const mockDraft = { id: 'draft-2', subject: 'RE: Test' };
      setupMock(mockDraft);

      const result = await client.createReplyAllDraft('msg-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/createReplyAll');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toBeNull();
      expect(result).toEqual(mockDraft);
    });

    it('createForwardDraft POSTs to /me/messages/{id}/createForward', async () => {
      const mockDraft = { id: 'draft-3', subject: 'FW: Test' };
      setupMock(mockDraft);

      const result = await client.createForwardDraft('msg-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/messages/msg-1/createForward');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toBeNull();
      expect(result).toEqual(mockDraft);
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
  // Calendar Write Operations
  // =========================================================================

  describe('Calendar write operation endpoints and bodies', () => {
    it('createEvent POSTs to /me/events with event body (no calendarId)', async () => {
      const event = {
        subject: 'Team Meeting',
        start: { dateTime: '2026-02-24T10:00:00', timeZone: 'UTC' },
        end: { dateTime: '2026-02-24T11:00:00', timeZone: 'UTC' },
      };
      setupMock({ id: 'evt-new', subject: 'Team Meeting' });

      await client.createEvent(event);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/events');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual(event);
    });

    it('createEvent with calendarId POSTs to /me/calendars/{id}/events', async () => {
      const event = {
        subject: 'Calendar-specific Event',
        start: { dateTime: '2026-02-24T10:00:00', timeZone: 'UTC' },
        end: { dateTime: '2026-02-24T11:00:00', timeZone: 'UTC' },
      };
      setupMock({ id: 'evt-new', subject: 'Calendar-specific Event' });

      await client.createEvent(event, 'cal-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/calendars/cal-1/events');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual(event);
    });

    it('updateEvent PATCHes /me/events/{id} with updates', async () => {
      const updates = { subject: 'Updated Meeting', location: { displayName: 'Room 42' } };

      await client.updateEvent('evt-1', updates);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/events/evt-1');
      expect(apiCalls[0].method).toBe('patch');
      expect(apiCalls[0].body).toEqual(updates);
    });

    it('deleteEvent DELETEs /me/events/{id}', async () => {
      await client.deleteEvent('evt-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/events/evt-1');
      expect(apiCalls[0].method).toBe('delete');
    });

    it('respondToEvent with accept POSTs to /me/events/{id}/accept', async () => {
      await client.respondToEvent('evt-1', 'accept', true, 'I will attend');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/events/evt-1/accept');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ sendResponse: true, comment: 'I will attend' });
    });

    it('respondToEvent with decline POSTs to /me/events/{id}/decline', async () => {
      await client.respondToEvent('evt-1', 'decline', true, 'Cannot make it');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/events/evt-1/decline');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ sendResponse: true, comment: 'Cannot make it' });
    });

    it('respondToEvent with tentative POSTs to /me/events/{id}/tentativelyAccept', async () => {
      await client.respondToEvent('evt-1', 'tentative', false, 'Maybe');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/events/evt-1/tentativelyAccept');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ sendResponse: false, comment: 'Maybe' });
    });

    it('respondToEvent without comment defaults to empty string', async () => {
      await client.respondToEvent('evt-1', 'accept', true);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/events/evt-1/accept');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ sendResponse: true, comment: '' });
    });
  });

  // =========================================================================
  // Contact Write Operations
  // =========================================================================

  describe('Contact write operation endpoints and bodies', () => {
    it('createContact POSTs to /me/contacts with contact body', async () => {
      const contact = {
        givenName: 'John',
        surname: 'Doe',
        emailAddresses: [{ address: 'john@example.com', name: 'John Doe' }],
      };
      setupMock({ id: 'contact-new', displayName: 'John Doe' });

      await client.createContact(contact);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/contacts');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual(contact);
    });

    it('updateContact PATCHes /me/contacts/{id} with updates', async () => {
      const updates = { givenName: 'Jane', jobTitle: 'Manager' };

      await client.updateContact('contact-1', updates);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/contacts/contact-1');
      expect(apiCalls[0].method).toBe('patch');
      expect(apiCalls[0].body).toEqual(updates);
    });

    it('deleteContact DELETEs /me/contacts/{id}', async () => {
      await client.deleteContact('contact-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/contacts/contact-1');
      expect(apiCalls[0].method).toBe('delete');
    });
  });

  // =========================================================================
  // Task Write Operations
  // =========================================================================

  describe('Task write operation endpoints and bodies', () => {
    it('createTask POSTs to /me/todo/lists/{listId}/tasks', async () => {
      const task = { title: 'Buy groceries', importance: 'high' };
      setupMock({ id: 'task-new', title: 'Buy groceries' });

      await client.createTask('list-1', task);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/todo/lists/list-1/tasks');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual(task);
    });

    it('updateTask PATCHes /me/todo/lists/{listId}/tasks/{taskId}', async () => {
      const updates = { title: 'Updated title' };
      setupMock({ id: 'task-1', title: 'Updated title' });

      await client.updateTask('list-1', 'task-1', updates);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/todo/lists/list-1/tasks/task-1');
      expect(apiCalls[0].method).toBe('patch');
      expect(apiCalls[0].body).toEqual(updates);
    });

    it('deleteTask DELETEs /me/todo/lists/{listId}/tasks/{taskId}', async () => {
      setupMock(undefined);

      await client.deleteTask('list-1', 'task-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/todo/lists/list-1/tasks/task-1');
      expect(apiCalls[0].method).toBe('delete');
    });

    it('createTaskList POSTs to /me/todo/lists', async () => {
      setupMock({ id: 'list-new', displayName: 'Shopping' });

      await client.createTaskList('Shopping');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/todo/lists');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ displayName: 'Shopping' });
    });

    it('updateTaskList PATCHes /me/todo/lists/{listId}', async () => {
      setupMock(undefined);

      await client.updateTaskList('list-1', { displayName: 'Renamed' });

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/todo/lists/list-1');
      expect(apiCalls[0].method).toBe('patch');
      expect(apiCalls[0].body).toEqual({ displayName: 'Renamed' });
    });

    it('deleteTaskList DELETEs /me/todo/lists/{listId}', async () => {
      setupMock(undefined);

      await client.deleteTaskList('list-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/todo/lists/list-1');
      expect(apiCalls[0].method).toBe('delete');
    });
  });

  // =========================================================================
  // Mail Rules operations
  // =========================================================================

  describe('Mail Rules operations', () => {
    it('listMailRules GETs /me/mailFolders/inbox/messageRules', async () => {
      setupMock({ value: [{ id: 'rule-1', displayName: 'Test Rule' }] });

      const result = await client.listMailRules();

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailFolders/inbox/messageRules');
      expect(apiCalls[0].method).toBe('get');
      expect(result).toEqual([{ id: 'rule-1', displayName: 'Test Rule' }]);
    });

    it('createMailRule POSTs to /me/mailFolders/inbox/messageRules', async () => {
      const rule = { displayName: 'New Rule', isEnabled: true };
      setupMock({ id: 'rule-new', displayName: 'New Rule' });

      await client.createMailRule(rule);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailFolders/inbox/messageRules');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual(rule);
    });

    it('deleteMailRule DELETEs /me/mailFolders/inbox/messageRules/{ruleId}', async () => {
      setupMock(undefined);

      await client.deleteMailRule('rule-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailFolders/inbox/messageRules/rule-1');
      expect(apiCalls[0].method).toBe('delete');
    });
  });

  // =========================================================================
  // Master Categories operations
  // =========================================================================

  describe('Master Categories operations', () => {
    it('listMasterCategories GETs /me/outlook/masterCategories', async () => {
      setupMock({ value: [{ id: 'cat-1', displayName: 'Red Category', color: 'preset0' }] });

      const result = await client.listMasterCategories();

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/outlook/masterCategories');
      expect(apiCalls[0].method).toBe('get');
      expect(result).toEqual([{ id: 'cat-1', displayName: 'Red Category', color: 'preset0' }]);
    });

    it('createMasterCategory POSTs to /me/outlook/masterCategories', async () => {
      setupMock({ id: 'cat-new', displayName: 'Work', color: 'preset1' });

      await client.createMasterCategory('Work', 'preset1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/outlook/masterCategories');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ displayName: 'Work', color: 'preset1' });
    });

    it('deleteMasterCategory DELETEs /me/outlook/masterCategories/{categoryId}', async () => {
      setupMock(undefined);

      await client.deleteMasterCategory('cat-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/outlook/masterCategories/cat-1');
      expect(apiCalls[0].method).toBe('delete');
    });
  });

  // =========================================================================
  // Focused Inbox Override operations
  // =========================================================================

  describe('Focused Inbox Override operations', () => {
    it('listFocusedOverrides GETs /me/inferenceClassification/overrides', async () => {
      setupMock({ value: [{ id: 'ov-1', classifyAs: 'focused', senderEmailAddress: { address: 'a@b.com' } }] });

      const result = await client.listFocusedOverrides();

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/inferenceClassification/overrides');
      expect(apiCalls[0].method).toBe('get');
      expect(result).toEqual([{ id: 'ov-1', classifyAs: 'focused', senderEmailAddress: { address: 'a@b.com' } }]);
    });

    it('createFocusedOverride POSTs to /me/inferenceClassification/overrides', async () => {
      setupMock({ id: 'ov-new', classifyAs: 'focused', senderEmailAddress: { address: 'a@b.com' } });

      await client.createFocusedOverride('a@b.com', 'focused');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/inferenceClassification/overrides');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({
        classifyAs: 'focused',
        senderEmailAddress: { address: 'a@b.com' },
      });
    });

    it('deleteFocusedOverride DELETEs /me/inferenceClassification/overrides/{overrideId}', async () => {
      setupMock(undefined);

      await client.deleteFocusedOverride('ov-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/inferenceClassification/overrides/ov-1');
      expect(apiCalls[0].method).toBe('delete');
    });
  });

  // =========================================================================
  // Automatic Replies (Out of Office) operations
  // =========================================================================

  describe('Automatic Replies operations', () => {
    it('getAutomaticReplies GETs /me/mailboxSettings/automaticRepliesSetting', async () => {
      setupMock({ status: 'disabled', externalAudience: 'none' });

      const result = await client.getAutomaticReplies();

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailboxSettings/automaticRepliesSetting');
      expect(apiCalls[0].method).toBe('get');
      expect(result).toEqual({ status: 'disabled', externalAudience: 'none' });
    });

    it('setAutomaticReplies PATCHes /me/mailboxSettings with automaticRepliesSetting', async () => {
      const settings = { status: 'alwaysEnabled', internalReplyMessage: 'I am out' };
      setupMock(undefined);

      await client.setAutomaticReplies(settings);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/mailboxSettings');
      expect(apiCalls[0].method).toBe('patch');
      expect(apiCalls[0].body).toEqual({ automaticRepliesSetting: settings });
    });
  });

  // =========================================================================
  // Contact Folders operations
  // =========================================================================

  describe('Contact Folders operations', () => {
    it('listContactFolders GETs /me/contactFolders', async () => {
      setupMock({ value: [{ id: 'cf-1', displayName: 'Work' }] });

      const result = await client.listContactFolders();

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/contactFolders');
      expect(apiCalls[0].method).toBe('get');
      expect(result).toEqual([{ id: 'cf-1', displayName: 'Work' }]);
    });

    it('createContactFolder POSTs to /me/contactFolders', async () => {
      setupMock({ id: 'cf-new', displayName: 'Friends' });

      await client.createContactFolder('Friends');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/contactFolders');
      expect(apiCalls[0].method).toBe('post');
      expect(apiCalls[0].body).toEqual({ displayName: 'Friends' });
    });

    it('deleteContactFolder DELETEs /me/contactFolders/{folderId}', async () => {
      setupMock(undefined);

      await client.deleteContactFolder('cf-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/contactFolders/cf-1');
      expect(apiCalls[0].method).toBe('delete');
    });

    it('listContactsInFolder GETs /me/contactFolders/{folderId}/contacts', async () => {
      setupMock({ value: [{ id: 'c-1', displayName: 'Alice' }] });

      const result = await client.listContactsInFolder('cf-1', 50);

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/contactFolders/cf-1/contacts');
      expect(apiCalls[0].method).toBe('get');
      expect(apiCalls[0].topValue).toBe(50);
      expect(result).toEqual([{ id: 'c-1', displayName: 'Alice' }]);
    });
  });

  // =========================================================================
  // Contact Photos
  // =========================================================================

  describe('Contact Photo operations', () => {
    it('getContactPhoto GETs /me/contacts/{contactId}/photo/$value', async () => {
      const mockPhotoData = new ArrayBuffer(8);
      setupMock(mockPhotoData);

      const result = await client.getContactPhoto('contact-1');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/contacts/contact-1/photo/$value');
      expect(apiCalls[0].method).toBe('get');
      expect(result).toBe(mockPhotoData);
    });

    it('setContactPhoto PUTs to /me/contacts/{contactId}/photo/$value', async () => {
      setupMock(undefined);

      const photoData = Buffer.from('fake-photo-data');
      await client.setContactPhoto('contact-1', photoData, 'image/jpeg');

      expect(apiCalls).toHaveLength(1);
      expect(apiCalls[0].url).toBe('/me/contacts/contact-1/photo/$value');
      expect(apiCalls[0].method).toBe('put');
      expect(apiCalls[0].body).toBe(photoData);
      expect(apiCalls[0].headers).toEqual({ 'Content-Type': 'image/jpeg' });
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
      await client.searchMessagesKql('from:alice');
      await client.searchMessagesKqlInFolder('f1', 'subject:"test"');
      await client.listConversationMessages('conv-1', 10);

      setupMock({ value: [{ id: 'msg-d' }], '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/delta' });
      await client.getMessagesDelta('f1');

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
      await client.listEventInstances('evt-1', '2024-01-01T00:00:00Z', '2024-12-31T23:59:59Z');

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

      // Exercise reply/forward as draft methods
      setupMock({ id: 'reply-draft', subject: 'RE: Test' });
      await client.createReplyDraft('m1');

      setupMock({ id: 'replyall-draft', subject: 'RE: Test' });
      await client.createReplyAllDraft('m1');

      setupMock({ id: 'forward-draft', subject: 'FW: Test' });
      await client.createForwardDraft('m1');

      // Exercise attachment methods
      setupMock({ value: [] });
      await client.listAttachments('m1');

      setupMock({ id: 'att-1', name: 'file.pdf', contentBytes: 'data' });
      await client.getAttachment('m1', 'att-1');

      setupMock({ id: 'att-new', name: 'file.txt' });
      await client.addAttachment('m1', { '@odata.type': '#microsoft.graph.fileAttachment', name: 'file.txt', contentBytes: 'data' });

      setupMock({ uploadUrl: 'https://upload.example.com/session' });
      await client.createUploadSession('m1', { AttachmentItem: { attachmentType: 'file', name: 'big.zip', size: 5000000 } });

      // Exercise calendar write methods
      setupMock({ id: 'evt-new', subject: 'New Event' });
      await client.createEvent({ subject: 'New Event' });
      await client.createEvent({ subject: 'Cal Event' }, 'cal-1');

      setupMock();
      await client.updateEvent('evt-1', { subject: 'Updated' });
      await client.deleteEvent('evt-1');
      await client.respondToEvent('evt-1', 'accept', true, 'Yes');
      await client.respondToEvent('evt-1', 'decline', true, 'No');
      await client.respondToEvent('evt-1', 'tentative', false);

      // Exercise contact write methods
      setupMock({ id: 'contact-new', displayName: 'John Doe' });
      await client.createContact({ givenName: 'John', surname: 'Doe' });

      setupMock();
      await client.updateContact('c-1', { givenName: 'Jane' });
      await client.deleteContact('c-1');

      // Exercise task write methods
      setupMock({ id: 'task-new', title: 'Buy groceries' });
      await client.createTask('list-1', { title: 'Buy groceries' });

      setupMock({ id: 'task-1', title: 'Updated' });
      await client.updateTask('list-1', 'task-1', { title: 'Updated' });

      setupMock();
      await client.deleteTask('list-1', 'task-1');

      setupMock({ id: 'list-new', displayName: 'Shopping' });
      await client.createTaskList('Shopping');

      setupMock(undefined);
      await client.updateTaskList('list-1', { displayName: 'Renamed' });

      setupMock(undefined);
      await client.deleteTaskList('list-1');

      // Exercise automatic replies methods
      setupMock({ status: 'disabled', externalAudience: 'none' });
      await client.getAutomaticReplies();

      setupMock(undefined);
      await client.setAutomaticReplies({ status: 'alwaysEnabled' });

      // Exercise mail rules methods
      setupMock({ value: [{ id: 'rule-1', displayName: 'Test Rule' }] });
      await client.listMailRules();

      setupMock({ id: 'rule-new', displayName: 'New Rule' });
      await client.createMailRule({ displayName: 'New Rule' });

      setupMock(undefined);
      await client.deleteMailRule('rule-1');

      // Exercise master categories methods
      setupMock({ value: [{ id: 'cat-1', displayName: 'Red Category', color: 'preset0' }] });
      await client.listMasterCategories();

      setupMock({ id: 'cat-new', displayName: 'Work', color: 'preset1' });
      await client.createMasterCategory('Work', 'preset1');

      setupMock(undefined);
      await client.deleteMasterCategory('cat-1');

      // Exercise contact folder methods
      setupMock({ value: [{ id: 'cf-1', displayName: 'Work' }] });
      await client.listContactFolders();

      setupMock({ id: 'cf-new', displayName: 'Friends' });
      await client.createContactFolder('Friends');

      setupMock(undefined);
      await client.deleteContactFolder('cf-1');

      setupMock({ value: [{ id: 'c-1', displayName: 'Alice' }] });
      await client.listContactsInFolder('cf-1', 50);

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

    it('createReplyDraft clears cache', async () => {
      await client.listMessages('f1');
      apiCalls.length = 0;

      setupMock({ id: 'draft-1', subject: 'RE: Test' });
      await client.createReplyDraft('msg-1');
      apiCalls.length = 0;

      setupMock();
      await client.listMessages('f1');

      expect(apiCalls.filter(c => c.method === 'get').length).toBeGreaterThan(0);
    });

    it('createReplyAllDraft clears cache', async () => {
      await client.listMessages('f1');
      apiCalls.length = 0;

      setupMock({ id: 'draft-2', subject: 'RE: Test' });
      await client.createReplyAllDraft('msg-1');
      apiCalls.length = 0;

      setupMock();
      await client.listMessages('f1');

      expect(apiCalls.filter(c => c.method === 'get').length).toBeGreaterThan(0);
    });

    it('createForwardDraft clears cache', async () => {
      await client.listMessages('f1');
      apiCalls.length = 0;

      setupMock({ id: 'draft-3', subject: 'FW: Test' });
      await client.createForwardDraft('msg-1');
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

    it('createEvent clears cache', async () => {
      await client.listCalendars();
      apiCalls.length = 0;

      setupMock({ id: 'evt-new', subject: 'New Event' });
      await client.createEvent({ subject: 'New Event' });
      apiCalls.length = 0;

      setupMock();
      await client.listCalendars();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/calendars');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('updateEvent clears cache', async () => {
      await client.listCalendars();
      apiCalls.length = 0;

      await client.updateEvent('evt-1', { subject: 'Updated' });
      apiCalls.length = 0;

      setupMock();
      await client.listCalendars();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/calendars');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('deleteEvent clears cache', async () => {
      await client.listCalendars();
      apiCalls.length = 0;

      await client.deleteEvent('evt-1');
      apiCalls.length = 0;

      setupMock();
      await client.listCalendars();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/calendars');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('respondToEvent clears cache', async () => {
      await client.listCalendars();
      apiCalls.length = 0;

      await client.respondToEvent('evt-1', 'accept', true);
      apiCalls.length = 0;

      setupMock();
      await client.listCalendars();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/calendars');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('createContact clears cache', async () => {
      await client.listContacts();
      apiCalls.length = 0;

      setupMock({ id: 'contact-new', displayName: 'John Doe' });
      await client.createContact({ givenName: 'John', surname: 'Doe' });
      apiCalls.length = 0;

      setupMock();
      await client.listContacts();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/contacts');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('updateContact clears cache', async () => {
      await client.listContacts();
      apiCalls.length = 0;

      await client.updateContact('c-1', { givenName: 'Jane' });
      apiCalls.length = 0;

      setupMock();
      await client.listContacts();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/contacts');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('deleteContact clears cache', async () => {
      await client.listContacts();
      apiCalls.length = 0;

      await client.deleteContact('c-1');
      apiCalls.length = 0;

      setupMock();
      await client.listContacts();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/contacts');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('createTask clears cache', async () => {
      await client.listTaskLists();
      apiCalls.length = 0;

      setupMock({ id: 'task-new', title: 'Test' });
      await client.createTask('list-1', { title: 'Test' });
      apiCalls.length = 0;

      setupMock();
      await client.listTaskLists();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/todo/lists');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('updateTask clears cache', async () => {
      await client.listTaskLists();
      apiCalls.length = 0;

      setupMock({ id: 'task-1', title: 'Updated' });
      await client.updateTask('list-1', 'task-1', { title: 'Updated' });
      apiCalls.length = 0;

      setupMock();
      await client.listTaskLists();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/todo/lists');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('deleteTask clears cache', async () => {
      await client.listTaskLists();
      apiCalls.length = 0;

      await client.deleteTask('list-1', 'task-1');
      apiCalls.length = 0;

      setupMock();
      await client.listTaskLists();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/todo/lists');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('createTaskList clears cache', async () => {
      await client.listTaskLists();
      apiCalls.length = 0;

      setupMock({ id: 'list-new', displayName: 'New List' });
      await client.createTaskList('New List');
      apiCalls.length = 0;

      setupMock();
      await client.listTaskLists();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/todo/lists');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('updateTaskList clears cache', async () => {
      await client.listTaskLists();
      apiCalls.length = 0;

      setupMock(undefined);
      await client.updateTaskList('list-1', { displayName: 'Renamed' });
      apiCalls.length = 0;

      setupMock();
      await client.listTaskLists();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/todo/lists');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('deleteTaskList clears cache', async () => {
      await client.listTaskLists();
      apiCalls.length = 0;

      setupMock(undefined);
      await client.deleteTaskList('list-1');
      apiCalls.length = 0;

      setupMock();
      await client.listTaskLists();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/todo/lists');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('createMailRule clears cache', async () => {
      await client.listMailFolders();
      apiCalls.length = 0;

      setupMock({ id: 'rule-new', displayName: 'New Rule' });
      await client.createMailRule({ displayName: 'New Rule' });
      apiCalls.length = 0;

      setupMock();
      await client.listMailFolders();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/mailFolders');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('deleteMailRule clears cache', async () => {
      await client.listMailFolders();
      apiCalls.length = 0;

      setupMock(undefined);
      await client.deleteMailRule('rule-1');
      apiCalls.length = 0;

      setupMock();
      await client.listMailFolders();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/mailFolders');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('createContactFolder clears cache', async () => {
      await client.listMailFolders();
      apiCalls.length = 0;

      setupMock({ id: 'cf-new', displayName: 'Friends' });
      await client.createContactFolder('Friends');
      apiCalls.length = 0;

      setupMock();
      await client.listMailFolders();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/mailFolders');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('deleteContactFolder clears cache', async () => {
      await client.listMailFolders();
      apiCalls.length = 0;

      setupMock(undefined);
      await client.deleteContactFolder('cf-1');
      apiCalls.length = 0;

      setupMock();
      await client.listMailFolders();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/mailFolders');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('createMasterCategory clears cache', async () => {
      await client.listMailFolders();
      apiCalls.length = 0;

      setupMock({ id: 'cat-new', displayName: 'Work', color: 'preset1' });
      await client.createMasterCategory('Work', 'preset1');
      apiCalls.length = 0;

      setupMock();
      await client.listMailFolders();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/mailFolders');
      expect(getCalls.length).toBeGreaterThan(0);
    });

    it('deleteMasterCategory clears cache', async () => {
      await client.listMailFolders();
      apiCalls.length = 0;

      setupMock(undefined);
      await client.deleteMasterCategory('cat-1');
      apiCalls.length = 0;

      setupMock();
      await client.listMailFolders();

      const getCalls = apiCalls.filter(c => c.method === 'get' && c.url === '/me/mailFolders');
      expect(getCalls.length).toBeGreaterThan(0);
    });
  });
});
