/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Graph API client wrapper.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

// Mock request builder that chains all methods
const createMockRequestBuilder = (mockResponse: any) => {
  const builder: any = {
    select: vi.fn().mockReturnThis(),
    top: vi.fn().mockReturnThis(),
    skip: vi.fn().mockReturnThis(),
    orderby: vi.fn().mockReturnThis(),
    filter: vi.fn().mockReturnThis(),
    search: vi.fn().mockReturnThis(),
    query: vi.fn().mockReturnThis(),
    get: vi.fn().mockResolvedValue(mockResponse),
  };
  return builder;
};

// Mock Graph client
const mockApi = vi.fn();
const mockGraphClient = {
  api: mockApi,
};

vi.mock('@microsoft/microsoft-graph-client', () => ({
  Client: {
    init: vi.fn(function() { return mockGraphClient; }),
  },
}));

// Mock auth module
vi.mock('../../../../src/graph/auth/index.js', () => ({
  getAccessToken: vi.fn().mockResolvedValue('test-access-token'),
}));

// Mock isomorphic-fetch
vi.mock('isomorphic-fetch', () => ({
  default: vi.fn(),
}));

import { GraphClient, createGraphClient } from '../../../../src/graph/client/graph-client.js';

describe('graph/client/graph-client', () => {
  let graphClient: GraphClient;

  beforeEach(() => {
    vi.clearAllMocks();
    graphClient = new GraphClient();
  });

  describe('createGraphClient', () => {
    it('creates a GraphClient instance', () => {
      const client = createGraphClient();
      expect(client).toBeInstanceOf(GraphClient);
    });

    it('accepts an optional deviceCodeCallback', () => {
      const callback = vi.fn();
      const client = createGraphClient(callback);
      expect(client).toBeInstanceOf(GraphClient);
    });
  });

  describe('clearCache', () => {
    it('clears the response cache', () => {
      // This just tests that the method exists and doesn't throw
      expect(() => graphClient.clearCache()).not.toThrow();
    });
  });

  describe('Mail Folders', () => {
    describe('listMailFolders', () => {
      it('returns mail folders', async () => {
        const mockFolders: MicrosoftGraph.MailFolder[] = [
          { id: 'folder-1', displayName: 'Inbox' },
          { id: 'folder-2', displayName: 'Sent' },
        ];

        // First call for top-level folders
        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: mockFolders,
          })
        );

        // Child folder calls
        mockApi.mockReturnValue(
          createMockRequestBuilder({
            value: [],
          })
        );

        const result = await graphClient.listMailFolders();

        expect(result).toHaveLength(2);
        expect(result[0].displayName).toBe('Inbox');
      });

      it('handles pagination', async () => {
        const firstPage: MicrosoftGraph.MailFolder[] = [
          { id: 'folder-1', displayName: 'Inbox' },
        ];

        const secondPage: MicrosoftGraph.MailFolder[] = [
          { id: 'folder-2', displayName: 'Sent' },
        ];

        // First page with nextLink
        const firstBuilder = createMockRequestBuilder({
          value: firstPage,
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/mailFolders?$skip=1',
        });

        // Second page (no nextLink)
        const secondBuilder = createMockRequestBuilder({
          value: secondPage,
        });

        // Child folder calls return empty
        const childBuilder = createMockRequestBuilder({
          value: [],
        });

        mockApi
          .mockReturnValueOnce(firstBuilder)  // First page
          .mockReturnValueOnce(secondBuilder)  // Second page (nextLink)
          .mockReturnValue(childBuilder);  // Child folders

        const result = await graphClient.listMailFolders();

        expect(result.length).toBeGreaterThanOrEqual(2);
      });

      it('returns cached results on subsequent calls', async () => {
        // Create a fresh client for this specific test
        const testClient = new GraphClient();

        const mockFolders: MicrosoftGraph.MailFolder[] = [
          { id: 'folder-1', displayName: 'Inbox' },
        ];

        // First call for top-level folders
        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: mockFolders,
          })
        );

        // Child folder calls return empty
        mockApi.mockReturnValue(
          createMockRequestBuilder({
            value: [],
          })
        );

        // First call populates cache
        await testClient.listMailFolders();

        // Reset mock to track if API is called again
        mockApi.mockClear();

        // Second call should use cache
        const result = await testClient.listMailFolders();

        expect(result).toHaveLength(1);
        // API should not be called again (cached)
        expect(mockApi).not.toHaveBeenCalled();
      });

      it('handles child folder errors gracefully', async () => {
        const mockFolders: MicrosoftGraph.MailFolder[] = [
          { id: 'folder-1', displayName: 'Inbox' },
        ];

        // First call returns folders
        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: mockFolders,
          })
        );

        // Child folder call throws
        const errorBuilder = createMockRequestBuilder({});
        errorBuilder.get.mockRejectedValue(new Error('Access denied'));
        mockApi.mockReturnValue(errorBuilder);

        // Should not throw
        const result = await graphClient.listMailFolders();

        expect(result).toHaveLength(1);
      });
    });

    describe('getMailFolder', () => {
      it('returns a mail folder by ID', async () => {
        const mockFolder: MicrosoftGraph.MailFolder = {
          id: 'folder-1',
          displayName: 'Inbox',
        };

        mockApi.mockReturnValue(createMockRequestBuilder(mockFolder));

        const result = await graphClient.getMailFolder('folder-1');

        expect(result?.displayName).toBe('Inbox');
      });

      it('returns null when folder not found', async () => {
        const errorBuilder = createMockRequestBuilder(null);
        errorBuilder.get.mockRejectedValue(new Error('Not found'));
        mockApi.mockReturnValue(errorBuilder);

        const result = await graphClient.getMailFolder('invalid-id');

        expect(result).toBeNull();
      });
    });
  });

  describe('Messages (Emails)', () => {
    describe('listMessages', () => {
      it('returns messages from a folder', async () => {
        const mockMessages: MicrosoftGraph.Message[] = [
          { id: 'msg-1', subject: 'Test Email' },
          { id: 'msg-2', subject: 'Another Email' },
        ];

        mockApi.mockReturnValue(
          createMockRequestBuilder({
            value: mockMessages,
          })
        );

        const result = await graphClient.listMessages('folder-1', 50, 0);

        expect(result).toHaveLength(2);
        expect(result[0].subject).toBe('Test Email');
      });

      it('uses pagination parameters', async () => {
        const mockMessages: MicrosoftGraph.Message[] = [];

        const builder = createMockRequestBuilder({
          value: mockMessages,
        });

        mockApi.mockReturnValue(builder);

        await graphClient.listMessages('folder-1', 25, 50);

        expect(builder.top).toHaveBeenCalledWith(25);
        expect(builder.skip).toHaveBeenCalledWith(50);
      });

      it('returns cached results', async () => {
        mockApi.mockReturnValue(
          createMockRequestBuilder({
            value: [{ id: 'msg-1' }],
          })
        );

        // First call
        await graphClient.listMessages('folder-1', 50, 0);

        mockApi.mockClear();

        // Second call should use cache
        const result = await graphClient.listMessages('folder-1', 50, 0);

        expect(result).toHaveLength(1);
        expect(mockApi).not.toHaveBeenCalled();
      });
    });

    describe('listUnreadMessages', () => {
      it('filters for unread messages', async () => {
        const mockMessages: MicrosoftGraph.Message[] = [
          { id: 'msg-1', subject: 'Unread Email', isRead: false },
        ];

        const builder = createMockRequestBuilder({
          value: mockMessages,
        });

        mockApi.mockReturnValue(builder);

        const result = await graphClient.listUnreadMessages('folder-1', 50, 0);

        expect(result).toHaveLength(1);
        expect(builder.filter).toHaveBeenCalledWith('isRead eq false');
      });
    });

    describe('searchMessages', () => {
      it('searches messages across all folders', async () => {
        const mockMessages: MicrosoftGraph.Message[] = [
          { id: 'msg-1', subject: 'Matching Email' },
        ];

        const builder = createMockRequestBuilder({
          value: mockMessages,
        });

        mockApi.mockReturnValue(builder);

        const result = await graphClient.searchMessages('test query', 50);

        expect(result).toHaveLength(1);
        expect(builder.search).toHaveBeenCalledWith('"test query"');
      });
    });

    describe('searchMessagesInFolder', () => {
      it('searches messages in a specific folder', async () => {
        const mockMessages: MicrosoftGraph.Message[] = [
          { id: 'msg-1', subject: 'Folder Match' },
        ];

        const builder = createMockRequestBuilder({
          value: mockMessages,
        });

        mockApi.mockReturnValue(builder);

        const result = await graphClient.searchMessagesInFolder('folder-1', 'query', 50);

        expect(result).toHaveLength(1);
        expect(builder.search).toHaveBeenCalledWith('"query"');
        expect(mockApi).toHaveBeenCalledWith('/me/mailFolders/folder-1/messages');
      });
    });

    describe('getMessage', () => {
      it('returns a message by ID with full body', async () => {
        const mockMessage: MicrosoftGraph.Message = {
          id: 'msg-1',
          subject: 'Test Email',
          body: { content: '<p>Body content</p>' },
        };

        mockApi.mockReturnValue(createMockRequestBuilder(mockMessage));

        const result = await graphClient.getMessage('msg-1');

        expect(result?.subject).toBe('Test Email');
        expect(result?.body?.content).toBe('<p>Body content</p>');
      });

      it('returns null when message not found', async () => {
        const errorBuilder = createMockRequestBuilder(null);
        errorBuilder.get.mockRejectedValue(new Error('Not found'));
        mockApi.mockReturnValue(errorBuilder);

        const result = await graphClient.getMessage('invalid-id');

        expect(result).toBeNull();
      });
    });
  });

  describe('Calendars', () => {
    describe('listCalendars', () => {
      it('returns calendars', async () => {
        const mockCalendars: MicrosoftGraph.Calendar[] = [
          { id: 'cal-1', name: 'Personal' },
          { id: 'cal-2', name: 'Work' },
        ];

        mockApi.mockReturnValue(
          createMockRequestBuilder({
            value: mockCalendars,
          })
        );

        const result = await graphClient.listCalendars();

        expect(result).toHaveLength(2);
        expect(result[0].name).toBe('Personal');
      });

      it('returns cached results', async () => {
        mockApi.mockReturnValue(
          createMockRequestBuilder({
            value: [{ id: 'cal-1' }],
          })
        );

        await graphClient.listCalendars();
        mockApi.mockClear();

        const result = await graphClient.listCalendars();

        expect(result).toHaveLength(1);
        expect(mockApi).not.toHaveBeenCalled();
      });
    });
  });

  describe('Events', () => {
    describe('listEvents', () => {
      it('returns upcoming events', async () => {
        const mockEvents: MicrosoftGraph.Event[] = [
          { id: 'evt-1', subject: 'Meeting' },
        ];

        mockApi.mockReturnValue(
          createMockRequestBuilder({
            value: mockEvents,
          })
        );

        const result = await graphClient.listEvents(50);

        expect(result).toHaveLength(1);
        expect(result[0].subject).toBe('Meeting');
      });

      it('uses calendarView for date ranges', async () => {
        const mockEvents: MicrosoftGraph.Event[] = [
          { id: 'evt-1', subject: 'Meeting' },
        ];

        const builder = createMockRequestBuilder({
          value: mockEvents,
        });

        mockApi.mockReturnValue(builder);

        const startDate = new Date('2024-01-01');
        const endDate = new Date('2024-01-31');

        await graphClient.listEvents(50, undefined, startDate, endDate);

        expect(mockApi).toHaveBeenCalledWith('/me/calendarView');
        expect(builder.query).toHaveBeenCalled();
      });

      it('uses specific calendar when calendarId is provided', async () => {
        const mockEvents: MicrosoftGraph.Event[] = [];

        mockApi.mockReturnValue(
          createMockRequestBuilder({
            value: mockEvents,
          })
        );

        await graphClient.listEvents(50, 'cal-1');

        expect(mockApi).toHaveBeenCalledWith('/me/calendars/cal-1/events');
      });

      it('uses specific calendar with date range', async () => {
        const mockEvents: MicrosoftGraph.Event[] = [];

        mockApi.mockReturnValue(
          createMockRequestBuilder({
            value: mockEvents,
          })
        );

        const startDate = new Date('2024-01-01');
        const endDate = new Date('2024-01-31');

        await graphClient.listEvents(50, 'cal-1', startDate, endDate);

        expect(mockApi).toHaveBeenCalledWith('/me/calendars/cal-1/calendarView');
      });
    });

    describe('getEvent', () => {
      it('returns an event by ID', async () => {
        const mockEvent: MicrosoftGraph.Event = {
          id: 'evt-1',
          subject: 'Meeting',
        };

        mockApi.mockReturnValue(createMockRequestBuilder(mockEvent));

        const result = await graphClient.getEvent('evt-1');

        expect(result?.subject).toBe('Meeting');
      });

      it('returns null when event not found', async () => {
        const errorBuilder = createMockRequestBuilder(null);
        errorBuilder.get.mockRejectedValue(new Error('Not found'));
        mockApi.mockReturnValue(errorBuilder);

        const result = await graphClient.getEvent('invalid-id');

        expect(result).toBeNull();
      });
    });
  });

  describe('Contacts', () => {
    describe('listContacts', () => {
      it('returns contacts', async () => {
        const mockContacts: MicrosoftGraph.Contact[] = [
          { id: 'contact-1', displayName: 'John Doe' },
        ];

        mockApi.mockReturnValue(
          createMockRequestBuilder({
            value: mockContacts,
          })
        );

        const result = await graphClient.listContacts(50, 0);

        expect(result).toHaveLength(1);
        expect(result[0].displayName).toBe('John Doe');
      });

      it('returns cached results', async () => {
        mockApi.mockReturnValue(
          createMockRequestBuilder({
            value: [{ id: 'contact-1' }],
          })
        );

        await graphClient.listContacts(50, 0);
        mockApi.mockClear();

        const result = await graphClient.listContacts(50, 0);

        expect(result).toHaveLength(1);
        expect(mockApi).not.toHaveBeenCalled();
      });
    });

    describe('searchContacts', () => {
      it('searches contacts by display name', async () => {
        const mockContacts: MicrosoftGraph.Contact[] = [
          { id: 'contact-1', displayName: 'John Doe' },
        ];

        const builder = createMockRequestBuilder({
          value: mockContacts,
        });

        mockApi.mockReturnValue(builder);

        const result = await graphClient.searchContacts('John', 50);

        expect(result).toHaveLength(1);
        expect(builder.filter).toHaveBeenCalledWith("contains(displayName,'John')");
      });
    });

    describe('getContact', () => {
      it('returns a contact by ID', async () => {
        const mockContact: MicrosoftGraph.Contact = {
          id: 'contact-1',
          displayName: 'John Doe',
        };

        mockApi.mockReturnValue(createMockRequestBuilder(mockContact));

        const result = await graphClient.getContact('contact-1');

        expect(result?.displayName).toBe('John Doe');
      });

      it('returns null when contact not found', async () => {
        const errorBuilder = createMockRequestBuilder(null);
        errorBuilder.get.mockRejectedValue(new Error('Not found'));
        mockApi.mockReturnValue(errorBuilder);

        const result = await graphClient.getContact('invalid-id');

        expect(result).toBeNull();
      });
    });
  });

  describe('Tasks', () => {
    describe('listTaskLists', () => {
      it('returns task lists', async () => {
        const mockLists: MicrosoftGraph.TodoTaskList[] = [
          { id: 'list-1', displayName: 'Tasks' },
        ];

        mockApi.mockReturnValue(
          createMockRequestBuilder({
            value: mockLists,
          })
        );

        const result = await graphClient.listTaskLists();

        expect(result).toHaveLength(1);
        expect(result[0].displayName).toBe('Tasks');
      });

      it('returns cached results', async () => {
        mockApi.mockReturnValue(
          createMockRequestBuilder({
            value: [{ id: 'list-1' }],
          })
        );

        await graphClient.listTaskLists();
        mockApi.mockClear();

        const result = await graphClient.listTaskLists();

        expect(result).toHaveLength(1);
        expect(mockApi).not.toHaveBeenCalled();
      });
    });

    describe('listTasks', () => {
      it('returns tasks from a task list', async () => {
        const mockTasks: MicrosoftGraph.TodoTask[] = [
          { id: 'task-1', title: 'Task 1' },
        ];

        mockApi.mockReturnValue(
          createMockRequestBuilder({
            value: mockTasks,
          })
        );

        const result = await graphClient.listTasks('list-1', 50, 0);

        expect(result).toHaveLength(1);
        expect(result[0].title).toBe('Task 1');
      });

      it('filters out completed tasks when requested', async () => {
        const builder = createMockRequestBuilder({
          value: [],
        });

        mockApi.mockReturnValue(builder);

        await graphClient.listTasks('list-1', 50, 0, false);

        expect(builder.filter).toHaveBeenCalledWith("status ne 'completed'");
      });

      it('includes completed tasks by default', async () => {
        const builder = createMockRequestBuilder({
          value: [],
        });

        mockApi.mockReturnValue(builder);

        await graphClient.listTasks('list-1', 50, 0);

        expect(builder.filter).not.toHaveBeenCalled();
      });
    });

    describe('listAllTasks', () => {
      it('returns tasks from all task lists', async () => {
        // First call gets task lists
        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [{ id: 'list-1' }, { id: 'list-2' }],
          })
        );

        // Tasks from list-1
        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [{ id: 'task-1', title: 'Task 1' }],
          })
        );

        // Tasks from list-2
        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [{ id: 'task-2', title: 'Task 2' }],
          })
        );

        const result = await graphClient.listAllTasks(50, 0);

        expect(result).toHaveLength(2);
        expect(result[0].taskListId).toBeDefined();
      });

      it('skips task lists with no ID', async () => {
        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [{ id: 'list-1' }, { id: undefined }],
          })
        );

        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [{ id: 'task-1', title: 'Task 1' }],
          })
        );

        const result = await graphClient.listAllTasks(50, 0);

        expect(result).toHaveLength(1);
      });

      it('sorts tasks by due date', async () => {
        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [{ id: 'list-1' }],
          })
        );

        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [
              { id: 'task-1', title: 'Later', dueDateTime: { dateTime: '2024-12-31' } },
              { id: 'task-2', title: 'Sooner', dueDateTime: { dateTime: '2024-01-01' } },
              { id: 'task-3', title: 'No due date' },
            ],
          })
        );

        const result = await graphClient.listAllTasks(50, 0);

        expect(result[0].title).toBe('Sooner');
        expect(result[1].title).toBe('Later');
        expect(result[2].title).toBe('No due date');
      });

      it('applies pagination after sorting', async () => {
        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [{ id: 'list-1' }],
          })
        );

        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [
              { id: 'task-1' },
              { id: 'task-2' },
              { id: 'task-3' },
            ],
          })
        );

        const result = await graphClient.listAllTasks(2, 1);

        expect(result).toHaveLength(2);
      });
    });

    describe('getTask', () => {
      it('returns a task by ID', async () => {
        const mockTask: MicrosoftGraph.TodoTask = {
          id: 'task-1',
          title: 'Task 1',
        };

        mockApi.mockReturnValue(createMockRequestBuilder(mockTask));

        const result = await graphClient.getTask('list-1', 'task-1');

        expect(result?.title).toBe('Task 1');
      });

      it('returns null when task not found', async () => {
        const errorBuilder = createMockRequestBuilder(null);
        errorBuilder.get.mockRejectedValue(new Error('Not found'));
        mockApi.mockReturnValue(errorBuilder);

        const result = await graphClient.getTask('list-1', 'invalid-id');

        expect(result).toBeNull();
      });
    });

    describe('searchTasks', () => {
      it('searches tasks by title', async () => {
        // Create fresh client
        const testClient = new GraphClient();

        // Get all tasks first
        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [{ id: 'list-1' }],
          })
        );

        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [
              { id: 'task-1', title: 'Important task' },
              { id: 'task-2', title: 'Other stuff' },
            ],
          })
        );

        const result = await testClient.searchTasks('Important', 50);

        expect(result).toHaveLength(1);
        expect(result[0].title).toBe('Important task');
      });

      it('performs case-insensitive search', async () => {
        // Create fresh client
        const testClient = new GraphClient();

        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [{ id: 'list-1' }],
          })
        );

        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [
              { id: 'task-1', title: 'UPPERCASE TASK' },
            ],
          })
        );

        const result = await testClient.searchTasks('uppercase', 50);

        expect(result).toHaveLength(1);
      });

      it('limits search results', async () => {
        // Create fresh client
        const testClient = new GraphClient();

        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [{ id: 'list-1' }],
          })
        );

        mockApi.mockReturnValueOnce(
          createMockRequestBuilder({
            value: [
              { id: 'task-1', title: 'Match 1' },
              { id: 'task-2', title: 'Match 2' },
              { id: 'task-3', title: 'Match 3' },
            ],
          })
        );

        const result = await testClient.searchTasks('Match', 2);

        expect(result).toHaveLength(2);
      });
    });
  });

  describe('getSchedule', () => {
    it('calls POST /me/calendar/getSchedule with correct body and returns value array', async () => {
      const mockResponse = {
        value: [
          {
            scheduleId: 'bob@example.com',
            availabilityView: '0120',
            scheduleItems: [
              { status: 'busy', start: { dateTime: '2026-02-24T10:00:00' }, end: { dateTime: '2026-02-24T11:00:00' } },
            ],
          },
        ],
      };
      const postBuilder = { post: vi.fn().mockResolvedValue(mockResponse) };
      mockApi.mockReturnValue(postBuilder);

      const result = await graphClient.getSchedule({
        schedules: ['bob@example.com'],
        startTime: { dateTime: '2026-02-24T08:00:00', timeZone: 'UTC' },
        endTime: { dateTime: '2026-02-24T18:00:00', timeZone: 'UTC' },
        availabilityViewInterval: 30,
      });

      expect(mockApi).toHaveBeenCalledWith('/me/calendar/getSchedule');
      expect(postBuilder.post).toHaveBeenCalledWith({
        schedules: ['bob@example.com'],
        startTime: { dateTime: '2026-02-24T08:00:00', timeZone: 'UTC' },
        endTime: { dateTime: '2026-02-24T18:00:00', timeZone: 'UTC' },
        availabilityViewInterval: 30,
      });
      expect(result).toEqual(mockResponse.value);
    });
  });

  describe('findMeetingTimes', () => {
    it('calls POST /me/findMeetingTimes with correct body and returns full response', async () => {
      const mockResponse = {
        meetingTimeSuggestions: [
          {
            confidence: 100,
            meetingTimeSlot: {
              start: { dateTime: '2026-02-24T14:00:00', timeZone: 'UTC' },
              end: { dateTime: '2026-02-24T15:00:00', timeZone: 'UTC' },
            },
            attendeeAvailability: [
              { attendee: { emailAddress: { address: 'bob@example.com' } }, availability: 'free' },
            ],
          },
        ],
        emptySuggestionsReason: '',
      };
      const postBuilder = { post: vi.fn().mockResolvedValue(mockResponse) };
      mockApi.mockReturnValue(postBuilder);

      const result = await graphClient.findMeetingTimes({
        attendees: [{ emailAddress: { address: 'bob@example.com' }, type: 'required' }],
        meetingDuration: 'PT1H',
        timeConstraint: {
          timeslots: [{
            start: { dateTime: '2026-02-24T08:00:00', timeZone: 'UTC' },
            end: { dateTime: '2026-02-24T18:00:00', timeZone: 'UTC' },
          }],
        },
        maxCandidates: 5,
      });

      expect(mockApi).toHaveBeenCalledWith('/me/findMeetingTimes');
      expect(postBuilder.post).toHaveBeenCalledWith({
        attendees: [{ emailAddress: { address: 'bob@example.com' }, type: 'required' }],
        meetingDuration: 'PT1H',
        timeConstraint: {
          timeslots: [{
            start: { dateTime: '2026-02-24T08:00:00', timeZone: 'UTC' },
            end: { dateTime: '2026-02-24T18:00:00', timeZone: 'UTC' },
          }],
        },
        maxCandidates: 5,
      });
      expect(result).toEqual(mockResponse);
    });
  });
});
