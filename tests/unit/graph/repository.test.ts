/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Graph API repository.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { GraphRepository, createGraphRepository } from '../../../src/graph/repository.js';
import { hashStringToNumber } from '../../../src/graph/mappers/utils.js';
import { downloadAttachment } from '../../../src/graph/attachments.js';

vi.mocked(downloadAttachment);

// Mock the GraphClient
vi.mock('../../../src/graph/client/index.js', () => ({
  GraphClient: vi.fn().mockImplementation(function() {
    return {
      listMailFolders: vi.fn(),
      getMailFolder: vi.fn(),
      listMessages: vi.fn(),
      listUnreadMessages: vi.fn(),
      searchMessages: vi.fn(),
      searchMessagesInFolder: vi.fn(),
      getMessage: vi.fn(),
      listCalendars: vi.fn(),
      listEvents: vi.fn(),
      getEvent: vi.fn(),
      listContacts: vi.fn(),
      searchContacts: vi.fn(),
      getContact: vi.fn(),
      listTaskLists: vi.fn(),
      listAllTasks: vi.fn(),
      searchTasks: vi.fn(),
      getTask: vi.fn(),
      // Write operations
      moveMessage: vi.fn(),
      deleteMessage: vi.fn(),
      archiveMessage: vi.fn(),
      junkMessage: vi.fn(),
      updateMessage: vi.fn(),
      createMailFolder: vi.fn(),
      deleteMailFolder: vi.fn(),
      renameMailFolder: vi.fn(),
      moveMailFolder: vi.fn(),
      emptyMailFolder: vi.fn(),
      // Draft & send operations
      createDraft: vi.fn(),
      updateDraft: vi.fn(),
      sendDraft: vi.fn(),
      sendMail: vi.fn(),
      replyMessage: vi.fn(),
      forwardMessage: vi.fn(),
      // Reply/Forward as draft operations
      createReplyDraft: vi.fn(),
      createReplyAllDraft: vi.fn(),
      createForwardDraft: vi.fn(),
      // Attachment operations
      listAttachments: vi.fn(),
      // Calendar write operations
      createEvent: vi.fn(),
      updateEvent: vi.fn(),
      deleteEvent: vi.fn(),
      respondToEvent: vi.fn(),
      // Contact write operations
      createContact: vi.fn(),
      updateContact: vi.fn(),
      deleteContact: vi.fn(),
      // Task write operations
      createTask: vi.fn(),
      updateTask: vi.fn(),
      deleteTask: vi.fn(),
      createTaskList: vi.fn(),
      // Calendar scheduling operations
      getSchedule: vi.fn(),
      findMeetingTimes: vi.fn(),
    };
  }),
}));

// Mock the downloadAttachment helper
vi.mock('../../../src/graph/attachments.js', () => ({
  downloadAttachment: vi.fn(),
}));

describe('graph/repository', () => {
  let repository: GraphRepository;
  let mockClient: any;

  beforeEach(async () => {
    vi.clearAllMocks();
    repository = createGraphRepository();
    // Access the internal client for mocking
    mockClient = (repository as any).client;
  });

  describe('createGraphRepository', () => {
    it('creates a repository instance', () => {
      const repo = createGraphRepository();
      expect(repo).toBeInstanceOf(GraphRepository);
    });
  });

  describe('Folders', () => {
    describe('listFolders (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.listFolders()).toThrow('Use listFoldersAsync()');
      });
    });

    describe('listFoldersAsync', () => {
      it('returns mapped folder rows', async () => {
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-1', displayName: 'Inbox', totalItemCount: 100, unreadItemCount: 5 },
          { id: 'folder-2', displayName: 'Sent', totalItemCount: 50, unreadItemCount: 0 },
        ]);

        const result = await repository.listFoldersAsync();

        expect(result).toHaveLength(2);
        expect(result[0].id).toBe(hashStringToNumber('folder-1'));
        expect(result[0].name).toBe('Inbox');
        expect(result[1].id).toBe(hashStringToNumber('folder-2'));
      });

      it('caches folder IDs for later retrieval', async () => {
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-1', displayName: 'Inbox' },
        ]);

        await repository.listFoldersAsync();

        const graphId = repository.getGraphId('folder', hashStringToNumber('folder-1'));
        expect(graphId).toBe('folder-1');
      });
    });

    describe('getFolder (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.getFolder(123)).toThrow('Use getFolderAsync()');
      });
    });

    describe('getFolderAsync', () => {
      it('returns folder by numeric ID', async () => {
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-1', displayName: 'Inbox' },
        ]);
        mockClient.getMailFolder.mockResolvedValue({
          id: 'folder-1',
          displayName: 'Inbox',
          totalItemCount: 100,
        });

        // First populate the cache
        await repository.listFoldersAsync();

        const result = await repository.getFolderAsync(hashStringToNumber('folder-1'));

        expect(result?.name).toBe('Inbox');
      });

      it('returns undefined when folder not found', async () => {
        mockClient.listMailFolders.mockResolvedValue([]);
        mockClient.getMailFolder.mockResolvedValue(null);

        const result = await repository.getFolderAsync(99999);

        expect(result).toBeUndefined();
      });
    });
  });

  describe('Emails', () => {
    describe('listEmails (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.listEmails(1, 50, 0)).toThrow('Use listEmailsAsync()');
      });
    });

    describe('listEmailsAsync', () => {
      beforeEach(async () => {
        // Set up folder cache
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-1', displayName: 'Inbox' },
        ]);
        await repository.listFoldersAsync();
      });

      it('returns mapped email rows', async () => {
        mockClient.listMessages.mockResolvedValue([
          { id: 'msg-1', subject: 'Test Email', isRead: true },
          { id: 'msg-2', subject: 'Another Email', isRead: false },
        ]);

        const result = await repository.listEmailsAsync(hashStringToNumber('folder-1'), 50, 0);

        expect(result).toHaveLength(2);
        expect(result[0].subject).toBe('Test Email');
        expect(result[1].subject).toBe('Another Email');
      });

      it('returns empty array when folder not found', async () => {
        const result = await repository.listEmailsAsync(99999, 50, 0);
        expect(result).toEqual([]);
      });
    });

    describe('listUnreadEmails (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.listUnreadEmails(1, 50, 0)).toThrow('Use listUnreadEmailsAsync()');
      });
    });

    describe('searchEmails (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.searchEmails('query', 50)).toThrow('Use searchEmailsAsync()');
      });
    });

    describe('searchEmailsAsync', () => {
      it('returns search results', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-1', subject: 'Matching Email' },
        ]);

        const result = await repository.searchEmailsAsync('test query', 50);

        expect(result).toHaveLength(1);
        expect(mockClient.searchMessages).toHaveBeenCalledWith('test query', 50);
      });
    });

    describe('getEmail (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.getEmail(123)).toThrow('Use getEmailAsync()');
      });
    });

    describe('getEmailAsync', () => {
      it('returns email by numeric ID', async () => {
        // First populate the message cache
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-1', subject: 'Test' },
        ]);
        await repository.searchEmailsAsync('test', 50);

        mockClient.getMessage.mockResolvedValue({
          id: 'msg-1',
          subject: 'Test Email',
          body: { content: 'Body content' },
        });

        const result = await repository.getEmailAsync(hashStringToNumber('msg-1'));

        expect(result?.subject).toBe('Test Email');
      });

      it('returns undefined when message ID not in cache', async () => {
        const result = await repository.getEmailAsync(99999);
        expect(result).toBeUndefined();
      });
    });

    describe('getUnreadCount (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.getUnreadCount()).toThrow('Use getUnreadCountAsync()');
      });
    });

    describe('getUnreadCountAsync', () => {
      it('returns total unread count', async () => {
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'f1', unreadItemCount: 5 },
          { id: 'f2', unreadItemCount: 3 },
          { id: 'f3', unreadItemCount: 0 },
        ]);

        const result = await repository.getUnreadCountAsync();

        expect(result).toBe(8);
      });
    });

    describe('getUnreadCountByFolder (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.getUnreadCountByFolder(1)).toThrow('Use getUnreadCountByFolderAsync()');
      });
    });

    describe('getUnreadCountByFolderAsync', () => {
      it('returns unread count for cached folder', async () => {
        // Populate folder cache
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-1', displayName: 'Inbox', unreadItemCount: 10 },
        ]);
        await repository.listFoldersAsync();

        mockClient.getMailFolder.mockResolvedValue({
          id: 'folder-1',
          displayName: 'Inbox',
          unreadItemCount: 10,
        });

        const result = await repository.getUnreadCountByFolderAsync(hashStringToNumber('folder-1'));

        expect(result).toBe(10);
        expect(mockClient.getMailFolder).toHaveBeenCalledWith('folder-1');
      });

      it('refreshes folders if folder not in cache', async () => {
        // First call doesn't have the folder
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-1', displayName: 'Inbox', unreadItemCount: 5 },
        ]);

        mockClient.getMailFolder.mockResolvedValue({
          id: 'folder-1',
          displayName: 'Inbox',
          unreadItemCount: 5,
        });

        const result = await repository.getUnreadCountByFolderAsync(hashStringToNumber('folder-1'));

        expect(result).toBe(5);
      });

      it('returns 0 when folder not found after refresh', async () => {
        mockClient.listMailFolders.mockResolvedValue([]);

        const result = await repository.getUnreadCountByFolderAsync(99999);

        expect(result).toBe(0);
      });

      it('returns 0 when getMailFolder returns null', async () => {
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-1', displayName: 'Inbox' },
        ]);
        await repository.listFoldersAsync();

        mockClient.getMailFolder.mockResolvedValue(null);

        const result = await repository.getUnreadCountByFolderAsync(hashStringToNumber('folder-1'));

        expect(result).toBe(0);
      });
    });

    describe('searchEmailsInFolder (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.searchEmailsInFolder(1, 'query', 50)).toThrow('Use searchEmailsInFolderAsync()');
      });
    });

    describe('searchEmailsInFolderAsync', () => {
      beforeEach(async () => {
        // Set up folder cache
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-1', displayName: 'Inbox' },
        ]);
        await repository.listFoldersAsync();
      });

      it('returns search results for cached folder', async () => {
        mockClient.searchMessagesInFolder.mockResolvedValue([
          { id: 'msg-1', subject: 'Match' },
        ]);

        const result = await repository.searchEmailsInFolderAsync(
          hashStringToNumber('folder-1'),
          'Match',
          50
        );

        expect(result).toHaveLength(1);
        expect(mockClient.searchMessagesInFolder).toHaveBeenCalledWith('folder-1', 'Match', 50);
      });

      it('refreshes folders if folder not in cache', async () => {
        // Create fresh repository
        const freshRepo = createGraphRepository();
        const freshClient = (freshRepo as any).client;

        freshClient.listMailFolders.mockResolvedValue([
          { id: 'folder-new', displayName: 'New Folder' },
        ]);
        freshClient.searchMessagesInFolder.mockResolvedValue([
          { id: 'msg-1', subject: 'Found' },
        ]);

        const result = await freshRepo.searchEmailsInFolderAsync(
          hashStringToNumber('folder-new'),
          'query',
          50
        );

        expect(result).toHaveLength(1);
      });

      it('returns empty array when folder not found', async () => {
        mockClient.listMailFolders.mockResolvedValue([]);

        const result = await repository.searchEmailsInFolderAsync(99999, 'query', 50);

        expect(result).toEqual([]);
      });
    });
  });

  describe('Calendar', () => {
    describe('listCalendars (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.listCalendars()).toThrow('Use listCalendarsAsync()');
      });
    });

    describe('listCalendarsAsync', () => {
      it('returns mapped calendar rows', async () => {
        mockClient.listCalendars.mockResolvedValue([
          { id: 'cal-1', name: 'Personal' },
          { id: 'cal-2', name: 'Work' },
        ]);

        const result = await repository.listCalendarsAsync();

        expect(result).toHaveLength(2);
        expect(result[0].name).toBe('Personal');
      });
    });

    describe('listEvents (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.listEvents(50)).toThrow('Use listEventsAsync()');
      });
    });

    describe('listEventsAsync', () => {
      it('returns mapped event rows', async () => {
        mockClient.listEvents.mockResolvedValue([
          { id: 'evt-1', subject: 'Meeting' },
        ]);

        const result = await repository.listEventsAsync(50);

        expect(result).toHaveLength(1);
      });
    });

    describe('listEventsByDateRange (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.listEventsByDateRange(0, 1000, 50)).toThrow('Use listEventsByDateRangeAsync()');
      });
    });

    describe('listEventsByDateRangeAsync', () => {
      it('returns events in date range', async () => {
        mockClient.listEvents.mockResolvedValue([
          { id: 'evt-1', subject: 'Meeting' },
        ]);

        const startDate = Math.floor(Date.now() / 1000);
        const endDate = startDate + 86400;

        const result = await repository.listEventsByDateRangeAsync(startDate, endDate, 50);

        expect(result).toHaveLength(1);
        expect(mockClient.listEvents).toHaveBeenCalledWith(
          50,
          undefined,
          expect.any(Date),
          expect.any(Date)
        );
      });
    });

    describe('getEvent (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.getEvent(123)).toThrow('Use getEventAsync()');
      });
    });

    describe('getEventAsync', () => {
      it('returns event by numeric ID when cached', async () => {
        // Populate cache via listEvents
        mockClient.listEvents.mockResolvedValue([
          { id: 'evt-1', subject: 'Meeting' },
        ]);
        await repository.listEventsAsync(50);

        mockClient.getEvent.mockResolvedValue({
          id: 'evt-1',
          subject: 'Team Meeting',
          start: { dateTime: '2024-01-15T10:00:00' },
        });

        const result = await repository.getEventAsync(hashStringToNumber('evt-1'));

        expect(mockClient.getEvent).toHaveBeenCalledWith('evt-1');
        expect(result).toBeDefined();
      });

      it('returns undefined when event ID not in cache', async () => {
        const result = await repository.getEventAsync(99999);
        expect(result).toBeUndefined();
      });

      it('returns undefined when event is not found', async () => {
        // Populate cache
        mockClient.listEvents.mockResolvedValue([
          { id: 'evt-1', subject: 'Meeting' },
        ]);
        await repository.listEventsAsync(50);

        mockClient.getEvent.mockResolvedValue(null);

        const result = await repository.getEventAsync(hashStringToNumber('evt-1'));
        expect(result).toBeUndefined();
      });
    });

    describe('listEventsByFolder (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.listEventsByFolder(1, 50)).toThrow('Use listEventsByFolderAsync()');
      });
    });

    describe('listEventsByFolderAsync', () => {
      it('returns events for cached calendar ID', async () => {
        // Populate calendar cache
        mockClient.listCalendars.mockResolvedValue([
          { id: 'cal-1', name: 'Work' },
        ]);
        await repository.listCalendarsAsync();

        mockClient.listEvents.mockResolvedValue([
          { id: 'evt-1', subject: 'Work Meeting' },
        ]);

        const result = await repository.listEventsByFolderAsync(hashStringToNumber('cal-1'), 50);

        expect(result).toHaveLength(1);
        expect(mockClient.listEvents).toHaveBeenCalledWith(50, 'cal-1');
      });

      it('falls back to all events when calendar not found', async () => {
        mockClient.listEvents.mockResolvedValue([
          { id: 'evt-1', subject: 'Meeting' },
        ]);

        const result = await repository.listEventsByFolderAsync(99999, 50);

        expect(result).toHaveLength(1);
      });
    });
  });

  describe('Contacts', () => {
    describe('listContacts (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.listContacts(50, 0)).toThrow('Use listContactsAsync()');
      });
    });

    describe('listContactsAsync', () => {
      it('returns mapped contact rows', async () => {
        mockClient.listContacts.mockResolvedValue([
          { id: 'contact-1', displayName: 'John Doe' },
        ]);

        const result = await repository.listContactsAsync(50, 0);

        expect(result).toHaveLength(1);
        expect(result[0].displayName).toBe('John Doe');
      });
    });

    describe('searchContacts (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.searchContacts('query', 50)).toThrow('Use searchContactsAsync()');
      });
    });

    describe('searchContactsAsync', () => {
      it('returns search results', async () => {
        mockClient.searchContacts.mockResolvedValue([
          { id: 'contact-1', displayName: 'John Doe' },
        ]);

        const result = await repository.searchContactsAsync('John', 50);

        expect(result).toHaveLength(1);
        expect(result[0].displayName).toBe('John Doe');
        expect(mockClient.searchContacts).toHaveBeenCalledWith('John', 50);
      });

      it('caches contact IDs', async () => {
        mockClient.searchContacts.mockResolvedValue([
          { id: 'contact-1', displayName: 'John Doe' },
        ]);

        await repository.searchContactsAsync('John', 50);

        const graphId = repository.getGraphId('contact', hashStringToNumber('contact-1'));
        expect(graphId).toBe('contact-1');
      });
    });

    describe('getContact (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.getContact(123)).toThrow('Use getContactAsync()');
      });
    });

    describe('getContactAsync', () => {
      it('returns contact by numeric ID when cached', async () => {
        // Populate cache
        mockClient.searchContacts.mockResolvedValue([
          { id: 'contact-1', displayName: 'John Doe' },
        ]);
        await repository.searchContactsAsync('John', 50);

        mockClient.getContact.mockResolvedValue({
          id: 'contact-1',
          displayName: 'John Doe',
          surname: 'Doe',
        });

        const result = await repository.getContactAsync(hashStringToNumber('contact-1'));

        expect(result?.displayName).toBe('John Doe');
        expect(mockClient.getContact).toHaveBeenCalledWith('contact-1');
      });

      it('returns undefined when contact ID not in cache', async () => {
        const result = await repository.getContactAsync(99999);
        expect(result).toBeUndefined();
      });

      it('returns undefined when contact is not found', async () => {
        // Populate cache
        mockClient.searchContacts.mockResolvedValue([
          { id: 'contact-1', displayName: 'John' },
        ]);
        await repository.searchContactsAsync('John', 50);

        mockClient.getContact.mockResolvedValue(null);

        const result = await repository.getContactAsync(hashStringToNumber('contact-1'));
        expect(result).toBeUndefined();
      });
    });
  });

  describe('Tasks', () => {
    describe('listTasks (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.listTasks(50, 0)).toThrow('Use listTasksAsync()');
      });
    });

    describe('listTasksAsync', () => {
      it('returns mapped task rows', async () => {
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'Task 1' },
        ]);

        const result = await repository.listTasksAsync(50, 0);

        expect(result).toHaveLength(1);
        expect(result[0].name).toBe('Task 1');
      });
    });

    describe('listIncompleteTasks (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.listIncompleteTasks(50, 0)).toThrow('Use listIncompleteTasksAsync()');
      });
    });

    describe('listIncompleteTasksAsync', () => {
      it('calls listAllTasks with includeCompleted=false', async () => {
        mockClient.listAllTasks.mockResolvedValue([]);

        await repository.listIncompleteTasksAsync(50, 0);

        expect(mockClient.listAllTasks).toHaveBeenCalledWith(50, 0, false);
      });
    });

    describe('searchTasks (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.searchTasks('query', 50)).toThrow('Use searchTasksAsync()');
      });
    });

    describe('searchTasksAsync', () => {
      it('returns search results', async () => {
        mockClient.searchTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'Matching Task' },
        ]);

        const result = await repository.searchTasksAsync('Matching', 50);

        expect(result).toHaveLength(1);
        expect(result[0].name).toBe('Matching Task');
        expect(mockClient.searchTasks).toHaveBeenCalledWith('Matching', 50);
      });

      it('caches task IDs for later retrieval', async () => {
        mockClient.searchTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'Task' },
        ]);

        await repository.searchTasksAsync('Task', 50);

        const taskInfo = repository.getTaskInfo(hashStringToNumber('task-1'));
        expect(taskInfo).toEqual({ taskListId: 'list-1', taskId: 'task-1' });
      });
    });

    describe('getTask (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.getTask(123)).toThrow('Use getTaskAsync()');
      });
    });

    describe('getTaskAsync', () => {
      it('returns task by numeric ID', async () => {
        // Populate task cache
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'Task 1' },
        ]);
        await repository.listTasksAsync(50, 0);

        mockClient.getTask.mockResolvedValue({
          id: 'task-1',
          title: 'Task 1',
          status: 'completed',
        });

        const result = await repository.getTaskAsync(hashStringToNumber('task-1'));

        expect(result?.name).toBe('Task 1');
        expect(mockClient.getTask).toHaveBeenCalledWith('list-1', 'task-1');
      });

      it('returns undefined when task ID not in cache', async () => {
        const result = await repository.getTaskAsync(99999);
        expect(result).toBeUndefined();
      });
    });
  });

  describe('Notes (NOT SUPPORTED)', () => {
    describe('listNotes', () => {
      it('returns empty array (sync)', () => {
        expect(repository.listNotes(50, 0)).toEqual([]);
      });
    });

    describe('listNotesAsync', () => {
      it('returns empty array', async () => {
        const result = await repository.listNotesAsync(50, 0);
        expect(result).toEqual([]);
      });
    });

    describe('getNote', () => {
      it('returns undefined (sync)', () => {
        expect(repository.getNote(123)).toBeUndefined();
      });
    });

    describe('getNoteAsync', () => {
      it('returns undefined', async () => {
        const result = await repository.getNoteAsync(123);
        expect(result).toBeUndefined();
      });
    });
  });

  describe('Write Operations (Async)', () => {
    describe('moveEmailAsync', () => {
      it('moves message using cached IDs', async () => {
        // Populate caches
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-dest', displayName: 'Archive' },
        ]);
        await repository.listFoldersAsync();

        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-1', subject: 'Test' },
        ]);
        await repository.searchEmailsAsync('Test', 50);

        mockClient.moveMessage.mockResolvedValue(undefined);

        await repository.moveEmailAsync(
          hashStringToNumber('msg-1'),
          hashStringToNumber('folder-dest')
        );

        expect(mockClient.moveMessage).toHaveBeenCalledWith('msg-1', 'folder-dest');
      });

      it('throws when message ID not in cache', async () => {
        await expect(repository.moveEmailAsync(99999, 88888)).rejects.toThrow(
          'Message ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.'
        );
      });

      it('throws when folder ID not in cache', async () => {
        // Populate message cache only
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-1', subject: 'Test' },
        ]);
        await repository.searchEmailsAsync('Test', 50);

        await expect(
          repository.moveEmailAsync(hashStringToNumber('msg-1'), 99999)
        ).rejects.toThrow('Folder ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('deleteEmailAsync', () => {
      it('deletes message using cached ID', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-1', subject: 'Test' },
        ]);
        await repository.searchEmailsAsync('Test', 50);

        mockClient.deleteMessage.mockResolvedValue(undefined);

        await repository.deleteEmailAsync(hashStringToNumber('msg-1'));

        expect(mockClient.deleteMessage).toHaveBeenCalledWith('msg-1');
      });

      it('throws when message ID not in cache', async () => {
        await expect(repository.deleteEmailAsync(99999)).rejects.toThrow(
          'Message ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.'
        );
      });
    });

    describe('createFolderAsync', () => {
      it('creates folder and caches the result', async () => {
        mockClient.createMailFolder.mockResolvedValue({
          id: 'new-folder-id',
          displayName: 'New Folder',
          totalItemCount: 0,
          unreadItemCount: 0,
        });

        const result = await repository.createFolderAsync('New Folder');

        expect(result.name).toBe('New Folder');
        expect(mockClient.createMailFolder).toHaveBeenCalledWith('New Folder', undefined);
        // Verify it was cached
        const graphId = repository.getGraphId('folder', hashStringToNumber('new-folder-id'));
        expect(graphId).toBe('new-folder-id');
      });
    });
  });

  describe('Draft & Send Operations (Async)', () => {
    describe('createDraftAsync', () => {
      it('creates a draft with to, cc, bcc recipients and caches the result', async () => {
        mockClient.createDraft.mockResolvedValue({
          id: 'draft-1',
          subject: 'Test Draft',
          isDraft: true,
        });

        const result = await repository.createDraftAsync({
          subject: 'Test Draft',
          body: 'Hello world',
          bodyType: 'text',
          to: ['alice@example.com'],
          cc: ['bob@example.com'],
          bcc: ['charlie@example.com'],
        });

        expect(mockClient.createDraft).toHaveBeenCalledWith({
          subject: 'Test Draft',
          body: { contentType: 'text', content: 'Hello world' },
          toRecipients: [{ emailAddress: { address: 'alice@example.com' } }],
          ccRecipients: [{ emailAddress: { address: 'bob@example.com' } }],
          bccRecipients: [{ emailAddress: { address: 'charlie@example.com' } }],
        });

        expect(result).toEqual({ numericId: hashStringToNumber('draft-1'), graphId: 'draft-1' });

        // Verify cached
        const graphId = repository.getGraphId('message', result.numericId);
        expect(graphId).toBe('draft-1');
      });

      it('creates a draft with no optional recipients', async () => {
        mockClient.createDraft.mockResolvedValue({
          id: 'draft-2',
          subject: 'No Recipients',
          isDraft: true,
        });

        const result = await repository.createDraftAsync({
          subject: 'No Recipients',
          body: '<p>Hello</p>',
          bodyType: 'html',
        });

        expect(mockClient.createDraft).toHaveBeenCalledWith({
          subject: 'No Recipients',
          body: { contentType: 'html', content: '<p>Hello</p>' },
          toRecipients: [],
          ccRecipients: [],
          bccRecipients: [],
        });

        expect(result).toEqual({ numericId: hashStringToNumber('draft-2'), graphId: 'draft-2' });
      });
    });

    describe('updateDraftAsync', () => {
      it('updates draft using cached ID', async () => {
        // Populate message cache
        mockClient.searchMessages.mockResolvedValue([
          { id: 'draft-1', subject: 'Old Subject' },
        ]);
        await repository.searchEmailsAsync('Old', 50);

        mockClient.updateDraft.mockResolvedValue({
          id: 'draft-1',
          subject: 'New Subject',
        });

        await repository.updateDraftAsync(hashStringToNumber('draft-1'), {
          subject: 'New Subject',
        });

        expect(mockClient.updateDraft).toHaveBeenCalledWith('draft-1', {
          subject: 'New Subject',
        });
      });

      it('throws when draft ID not in cache', async () => {
        await expect(
          repository.updateDraftAsync(99999, { subject: 'New' })
        ).rejects.toThrow('Message ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('listDraftsAsync', () => {
      it('lists drafts using the drafts well-known folder name', async () => {
        mockClient.listMessages.mockResolvedValue([
          { id: 'draft-1', subject: 'Draft 1', isDraft: true },
          { id: 'draft-2', subject: 'Draft 2', isDraft: true },
        ]);

        const result = await repository.listDraftsAsync(50, 0);

        expect(result).toHaveLength(2);
        expect(result[0].subject).toBe('Draft 1');
        expect(result[1].subject).toBe('Draft 2');
        expect(mockClient.listMessages).toHaveBeenCalledWith('drafts', 50, 0);
      });

      it('caches draft message IDs', async () => {
        mockClient.listMessages.mockResolvedValue([
          { id: 'draft-1', subject: 'Draft 1' },
        ]);

        await repository.listDraftsAsync(50, 0);

        const graphId = repository.getGraphId('message', hashStringToNumber('draft-1'));
        expect(graphId).toBe('draft-1');
      });
    });

    describe('sendDraftAsync', () => {
      it('sends draft using cached ID', async () => {
        // Populate message cache
        mockClient.searchMessages.mockResolvedValue([
          { id: 'draft-1', subject: 'Ready to Send' },
        ]);
        await repository.searchEmailsAsync('Ready', 50);

        mockClient.sendDraft.mockResolvedValue(undefined);

        await repository.sendDraftAsync(hashStringToNumber('draft-1'));

        expect(mockClient.sendDraft).toHaveBeenCalledWith('draft-1');
      });

      it('throws when draft ID not in cache', async () => {
        await expect(repository.sendDraftAsync(99999)).rejects.toThrow(
          'Message ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.'
        );
      });
    });

    describe('sendMailAsync', () => {
      it('sends mail with all recipient types', async () => {
        mockClient.sendMail.mockResolvedValue(undefined);

        await repository.sendMailAsync({
          subject: 'Direct Send',
          body: 'Hello',
          bodyType: 'text',
          to: ['alice@example.com', 'bob@example.com'],
          cc: ['carol@example.com'],
          bcc: ['dave@example.com'],
        });

        expect(mockClient.sendMail).toHaveBeenCalledWith({
          subject: 'Direct Send',
          body: { contentType: 'text', content: 'Hello' },
          toRecipients: [
            { emailAddress: { address: 'alice@example.com' } },
            { emailAddress: { address: 'bob@example.com' } },
          ],
          ccRecipients: [{ emailAddress: { address: 'carol@example.com' } }],
          bccRecipients: [{ emailAddress: { address: 'dave@example.com' } }],
        });
      });

      it('sends mail with only required fields', async () => {
        mockClient.sendMail.mockResolvedValue(undefined);

        await repository.sendMailAsync({
          subject: 'Simple',
          body: '<p>Hi</p>',
          bodyType: 'html',
          to: ['alice@example.com'],
        });

        expect(mockClient.sendMail).toHaveBeenCalledWith({
          subject: 'Simple',
          body: { contentType: 'html', content: '<p>Hi</p>' },
          toRecipients: [{ emailAddress: { address: 'alice@example.com' } }],
          ccRecipients: [],
          bccRecipients: [],
        });
      });
    });

    describe('replyMessageAsync', () => {
      it('replies to a message using cached ID', async () => {
        // Populate message cache
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-1', subject: 'Original' },
        ]);
        await repository.searchEmailsAsync('Original', 50);

        mockClient.replyMessage.mockResolvedValue(undefined);

        await repository.replyMessageAsync(
          hashStringToNumber('msg-1'),
          'Thanks for the info!',
          false
        );

        expect(mockClient.replyMessage).toHaveBeenCalledWith(
          'msg-1',
          'Thanks for the info!',
          false
        );
      });

      it('replies all to a message', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-1', subject: 'Team Email' },
        ]);
        await repository.searchEmailsAsync('Team', 50);

        mockClient.replyMessage.mockResolvedValue(undefined);

        await repository.replyMessageAsync(
          hashStringToNumber('msg-1'),
          'Sounds good!',
          true
        );

        expect(mockClient.replyMessage).toHaveBeenCalledWith(
          'msg-1',
          'Sounds good!',
          true
        );
      });

      it('throws when message ID not in cache', async () => {
        await expect(
          repository.replyMessageAsync(99999, 'Hello', false)
        ).rejects.toThrow('Message ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('forwardMessageAsync', () => {
      it('forwards a message with recipients and comment', async () => {
        // Populate message cache
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-1', subject: 'FYI' },
        ]);
        await repository.searchEmailsAsync('FYI', 50);

        mockClient.forwardMessage.mockResolvedValue(undefined);

        await repository.forwardMessageAsync(
          hashStringToNumber('msg-1'),
          ['recipient@example.com', 'other@example.com'],
          'Please review'
        );

        expect(mockClient.forwardMessage).toHaveBeenCalledWith(
          'msg-1',
          [
            { emailAddress: { address: 'recipient@example.com' } },
            { emailAddress: { address: 'other@example.com' } },
          ],
          'Please review'
        );
      });

      it('forwards a message without comment', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-1', subject: 'FYI' },
        ]);
        await repository.searchEmailsAsync('FYI', 50);

        mockClient.forwardMessage.mockResolvedValue(undefined);

        await repository.forwardMessageAsync(
          hashStringToNumber('msg-1'),
          ['recipient@example.com']
        );

        expect(mockClient.forwardMessage).toHaveBeenCalledWith(
          'msg-1',
          [{ emailAddress: { address: 'recipient@example.com' } }],
          undefined
        );
      });

      it('throws when message ID not in cache', async () => {
        await expect(
          repository.forwardMessageAsync(99999, ['a@b.com'])
        ).rejects.toThrow('Message ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });
  });

  describe('Sync method stubs throw errors', () => {
    it('moveEmail throws', () => {
      expect(() => repository.moveEmail(1, 2)).toThrow('Use moveEmailAsync()');
    });
    it('deleteEmail throws', () => {
      expect(() => repository.deleteEmail(1)).toThrow('Use deleteEmailAsync()');
    });
    it('archiveEmail throws', () => {
      expect(() => repository.archiveEmail(1)).toThrow('Use archiveEmailAsync()');
    });
    it('junkEmail throws', () => {
      expect(() => repository.junkEmail(1)).toThrow('Use junkEmailAsync()');
    });
    it('markEmailRead throws', () => {
      expect(() => repository.markEmailRead(1, true)).toThrow('Use markEmailReadAsync()');
    });
    it('setEmailFlag throws', () => {
      expect(() => repository.setEmailFlag(1, 0)).toThrow('Use setEmailFlagAsync()');
    });
    it('setEmailCategories throws', () => {
      expect(() => repository.setEmailCategories(1, ['cat'])).toThrow('Use setEmailCategoriesAsync()');
    });
    it('setEmailImportance throws', () => {
      expect(() => repository.setEmailImportance(1, 'high')).toThrow('Use setEmailImportanceAsync()');
    });
    it('createFolder throws', () => {
      expect(() => repository.createFolder('test')).toThrow('Use createFolderAsync()');
    });
    it('deleteFolder throws', () => {
      expect(() => repository.deleteFolder(1)).toThrow('Use deleteFolderAsync()');
    });
    it('renameFolder throws', () => {
      expect(() => repository.renameFolder(1, 'new')).toThrow('Use renameFolderAsync()');
    });
    it('moveFolder throws', () => {
      expect(() => repository.moveFolder(1, 2)).toThrow('Use moveFolderAsync()');
    });
    it('emptyFolder throws', () => {
      expect(() => repository.emptyFolder(1)).toThrow('Use emptyFolderAsync()');
    });
  });

  describe('Email Write Operations (Async)', () => {
    describe('archiveEmailAsync', () => {
      it('calls archiveMessage with the correct graph ID', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-arch', subject: 'Archive me' },
        ]);
        await repository.searchEmailsAsync('Archive me', 50);

        mockClient.archiveMessage.mockResolvedValue(undefined);
        await repository.archiveEmailAsync(hashStringToNumber('msg-arch'));

        expect(mockClient.archiveMessage).toHaveBeenCalledWith('msg-arch');
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.archiveEmailAsync(99999)
        ).rejects.toThrow('Message ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('junkEmailAsync', () => {
      it('calls junkMessage with the correct graph ID', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-junk', subject: 'Spam' },
        ]);
        await repository.searchEmailsAsync('Spam', 50);

        mockClient.junkMessage.mockResolvedValue(undefined);
        await repository.junkEmailAsync(hashStringToNumber('msg-junk'));

        expect(mockClient.junkMessage).toHaveBeenCalledWith('msg-junk');
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.junkEmailAsync(99999)
        ).rejects.toThrow('Message ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('markEmailReadAsync', () => {
      it('calls updateMessage with isRead flag', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-read', subject: 'Read me' },
        ]);
        await repository.searchEmailsAsync('Read me', 50);

        mockClient.updateMessage.mockResolvedValue(undefined);
        await repository.markEmailReadAsync(hashStringToNumber('msg-read'), true);

        expect(mockClient.updateMessage).toHaveBeenCalledWith('msg-read', { isRead: true });
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.markEmailReadAsync(99999, false)
        ).rejects.toThrow('Message ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('setEmailFlagAsync', () => {
      it('maps flag status 0 to notFlagged', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-flag', subject: 'Flag me' },
        ]);
        await repository.searchEmailsAsync('Flag me', 50);

        mockClient.updateMessage.mockResolvedValue(undefined);
        await repository.setEmailFlagAsync(hashStringToNumber('msg-flag'), 0);

        expect(mockClient.updateMessage).toHaveBeenCalledWith('msg-flag', {
          flag: { flagStatus: 'notFlagged' },
        });
      });

      it('maps flag status 1 to flagged', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-flag1', subject: 'Flag 1' },
        ]);
        await repository.searchEmailsAsync('Flag 1', 50);

        mockClient.updateMessage.mockResolvedValue(undefined);
        await repository.setEmailFlagAsync(hashStringToNumber('msg-flag1'), 1);

        expect(mockClient.updateMessage).toHaveBeenCalledWith('msg-flag1', {
          flag: { flagStatus: 'flagged' },
        });
      });

      it('maps flag status 2 to complete', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-flag2', subject: 'Flag 2' },
        ]);
        await repository.searchEmailsAsync('Flag 2', 50);

        mockClient.updateMessage.mockResolvedValue(undefined);
        await repository.setEmailFlagAsync(hashStringToNumber('msg-flag2'), 2);

        expect(mockClient.updateMessage).toHaveBeenCalledWith('msg-flag2', {
          flag: { flagStatus: 'complete' },
        });
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.setEmailFlagAsync(99999, 0)
        ).rejects.toThrow('Message ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('setEmailCategoriesAsync', () => {
      it('calls updateMessage with categories', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-cat', subject: 'Categorize me' },
        ]);
        await repository.searchEmailsAsync('Categorize me', 50);

        mockClient.updateMessage.mockResolvedValue(undefined);
        await repository.setEmailCategoriesAsync(
          hashStringToNumber('msg-cat'),
          ['Important', 'Work']
        );

        expect(mockClient.updateMessage).toHaveBeenCalledWith('msg-cat', {
          categories: ['Important', 'Work'],
        });
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.setEmailCategoriesAsync(99999, ['cat'])
        ).rejects.toThrow('Message ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('setEmailImportanceAsync', () => {
      it('updates message importance via updateMessage', async () => {
        mockClient.searchMessages.mockResolvedValue([{ id: 'msg-imp', subject: 'Test' }]);
        await repository.searchEmailsAsync('Test', 50);
        mockClient.updateMessage.mockResolvedValue(undefined);

        await repository.setEmailImportanceAsync(hashStringToNumber('msg-imp'), 'high');
        expect(mockClient.updateMessage).toHaveBeenCalledWith('msg-imp', { importance: 'high' });
      });

      it('throws when email not in cache', async () => {
        await expect(repository.setEmailImportanceAsync(99999, 'high'))
          .rejects.toThrow('Message ID 99999 not found in cache');
      });
    });
  });

  describe('Folder Write Operations (Async)', () => {
    describe('createFolderAsync', () => {
      it('calls createMailFolder and caches the new folder', async () => {
        mockClient.createMailFolder.mockResolvedValue({
          id: 'folder-new',
          displayName: 'Reports',
          parentFolderId: null,
          totalItemCount: 0,
          unreadItemCount: 0,
        });

        const result = await repository.createFolderAsync('Reports');

        expect(mockClient.createMailFolder).toHaveBeenCalledWith('Reports', undefined);
        expect(result.name).toBe('Reports');
        expect(result.id).toBe(hashStringToNumber('folder-new'));

        // Verify cache was updated
        const graphId = repository.getGraphId('folder', result.id);
        expect(graphId).toBe('folder-new');
      });

      it('passes parent folder graph ID when parentFolderId provided', async () => {
        // Populate folder cache
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'parent-folder', displayName: 'Parent', totalItemCount: 0, unreadItemCount: 0 },
        ]);
        await repository.listFoldersAsync();

        mockClient.createMailFolder.mockResolvedValue({
          id: 'sub-folder',
          displayName: 'SubFolder',
          parentFolderId: 'parent-folder',
          totalItemCount: 0,
          unreadItemCount: 0,
        });

        await repository.createFolderAsync('SubFolder', hashStringToNumber('parent-folder'));

        expect(mockClient.createMailFolder).toHaveBeenCalledWith('SubFolder', 'parent-folder');
      });
    });

    describe('deleteFolderAsync', () => {
      it('calls deleteMailFolder and removes from cache', async () => {
        // Populate folder cache
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-del', displayName: 'ToDelete', totalItemCount: 0, unreadItemCount: 0 },
        ]);
        await repository.listFoldersAsync();

        mockClient.deleteMailFolder.mockResolvedValue(undefined);

        const numericId = hashStringToNumber('folder-del');
        await repository.deleteFolderAsync(numericId);

        expect(mockClient.deleteMailFolder).toHaveBeenCalledWith('folder-del');
        expect(repository.getGraphId('folder', numericId)).toBeUndefined();
      });

      it('throws if folder not in cache', async () => {
        await expect(
          repository.deleteFolderAsync(99999)
        ).rejects.toThrow('Folder ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('renameFolderAsync', () => {
      it('calls renameMailFolder with the correct graph ID', async () => {
        // Populate folder cache
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-ren', displayName: 'OldName', totalItemCount: 0, unreadItemCount: 0 },
        ]);
        await repository.listFoldersAsync();

        mockClient.renameMailFolder.mockResolvedValue(undefined);

        await repository.renameFolderAsync(hashStringToNumber('folder-ren'), 'NewName');

        expect(mockClient.renameMailFolder).toHaveBeenCalledWith('folder-ren', 'NewName');
      });

      it('throws if folder not in cache', async () => {
        await expect(
          repository.renameFolderAsync(99999, 'NewName')
        ).rejects.toThrow('Folder ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('moveFolderAsync', () => {
      it('calls moveMailFolder with correct graph IDs', async () => {
        // Populate folder cache with both folders
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-src', displayName: 'Source', totalItemCount: 0, unreadItemCount: 0 },
          { id: 'folder-dest', displayName: 'Destination', totalItemCount: 0, unreadItemCount: 0 },
        ]);
        await repository.listFoldersAsync();

        mockClient.moveMailFolder.mockResolvedValue(undefined);

        await repository.moveFolderAsync(
          hashStringToNumber('folder-src'),
          hashStringToNumber('folder-dest')
        );

        expect(mockClient.moveMailFolder).toHaveBeenCalledWith('folder-src', 'folder-dest');
      });

      it('throws if source folder not in cache', async () => {
        await expect(
          repository.moveFolderAsync(99999, 88888)
        ).rejects.toThrow('Folder ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });

      it('throws if destination folder not in cache', async () => {
        // Populate only the source folder
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-only', displayName: 'Source', totalItemCount: 0, unreadItemCount: 0 },
        ]);
        await repository.listFoldersAsync();

        await expect(
          repository.moveFolderAsync(hashStringToNumber('folder-only'), 88888)
        ).rejects.toThrow('Parent folder ID 88888 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('emptyFolderAsync', () => {
      it('calls emptyMailFolder with the correct graph ID', async () => {
        // Populate folder cache
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-empty', displayName: 'Trash', totalItemCount: 5, unreadItemCount: 0 },
        ]);
        await repository.listFoldersAsync();

        mockClient.emptyMailFolder.mockResolvedValue(undefined);

        await repository.emptyFolderAsync(hashStringToNumber('folder-empty'));

        expect(mockClient.emptyMailFolder).toHaveBeenCalledWith('folder-empty');
      });

      it('throws if folder not in cache', async () => {
        await expect(
          repository.emptyFolderAsync(99999)
        ).rejects.toThrow('Folder ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });
  });

  describe('Reply/Forward as Draft (Async)', () => {
    describe('replyAsDraftAsync', () => {
      it('creates a reply draft and caches the result', async () => {
        // Populate message cache
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-orig', subject: 'Hello' },
        ]);
        await repository.searchEmailsAsync('Hello', 50);

        mockClient.createReplyDraft.mockResolvedValue({
          id: 'draft-reply-1',
          subject: 'Re: Hello',
          toRecipients: [{ emailAddress: { address: 'sender@example.com' } }],
        });

        const result = await repository.replyAsDraftAsync(hashStringToNumber('msg-orig'));

        expect(mockClient.createReplyDraft).toHaveBeenCalledWith('msg-orig');
        expect(result.numericId).toBe(hashStringToNumber('draft-reply-1'));
        expect(result.graphId).toBe('draft-reply-1');
        expect(repository.getGraphId('message', result.numericId)).toBe('draft-reply-1');
      });

      it('creates a reply-all draft when replyAll is true', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-orig2', subject: 'Team' },
        ]);
        await repository.searchEmailsAsync('Team', 50);

        mockClient.createReplyAllDraft.mockResolvedValue({
          id: 'draft-ra-1',
          subject: 'Re: Team',
          toRecipients: [{ emailAddress: { address: 'all@example.com' } }],
        });

        const result = await repository.replyAsDraftAsync(hashStringToNumber('msg-orig2'), true);

        expect(mockClient.createReplyAllDraft).toHaveBeenCalledWith('msg-orig2');
        expect(result.graphId).toBe('draft-ra-1');
      });

      it('updates draft body when comment is provided', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-comment', subject: 'FYI' },
        ]);
        await repository.searchEmailsAsync('FYI', 50);

        mockClient.createReplyDraft.mockResolvedValue({
          id: 'draft-comment-1',
          subject: 'Re: FYI',
          toRecipients: [],
        });
        mockClient.updateDraft.mockResolvedValue(undefined);

        await repository.replyAsDraftAsync(hashStringToNumber('msg-comment'), false, 'Thanks for sharing!');

        expect(mockClient.updateDraft).toHaveBeenCalledWith('draft-comment-1', {
          body: { contentType: 'text', content: 'Thanks for sharing!' },
        });
      });

      it('uses provided bodyType when updating comment', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-html', subject: 'HTML test' },
        ]);
        await repository.searchEmailsAsync('HTML test', 50);

        mockClient.createReplyDraft.mockResolvedValue({
          id: 'draft-html-1',
          subject: 'Re: HTML test',
          toRecipients: [],
        });
        mockClient.updateDraft.mockResolvedValue(undefined);

        await repository.replyAsDraftAsync(hashStringToNumber('msg-html'), false, '<p>HTML reply</p>', 'html');

        expect(mockClient.updateDraft).toHaveBeenCalledWith('draft-html-1', {
          body: { contentType: 'html', content: '<p>HTML reply</p>' },
        });
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.replyAsDraftAsync(99999)
        ).rejects.toThrow('Message ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('forwardAsDraftAsync', () => {
      it('creates a forward draft and caches the result', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-fwd', subject: 'Report' },
        ]);
        await repository.searchEmailsAsync('Report', 50);

        mockClient.createForwardDraft.mockResolvedValue({
          id: 'draft-fwd-1',
          subject: 'Fwd: Report',
          toRecipients: [],
        });

        const result = await repository.forwardAsDraftAsync(hashStringToNumber('msg-fwd'));

        expect(mockClient.createForwardDraft).toHaveBeenCalledWith('msg-fwd');
        expect(result.numericId).toBe(hashStringToNumber('draft-fwd-1'));
        expect(repository.getGraphId('message', result.numericId)).toBe('draft-fwd-1');
      });

      it('updates draft with recipients and comment when provided', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-fwd2', subject: 'Info' },
        ]);
        await repository.searchEmailsAsync('Info', 50);

        mockClient.createForwardDraft.mockResolvedValue({
          id: 'draft-fwd-2',
          subject: 'Fwd: Info',
          toRecipients: [],
        });
        mockClient.updateDraft.mockResolvedValue(undefined);

        await repository.forwardAsDraftAsync(
          hashStringToNumber('msg-fwd2'),
          ['alice@example.com', 'bob@example.com'],
          'Please review'
        );

        expect(mockClient.updateDraft).toHaveBeenCalledWith('draft-fwd-2', {
          toRecipients: [
            { emailAddress: { address: 'alice@example.com' } },
            { emailAddress: { address: 'bob@example.com' } },
          ],
          body: { contentType: 'text', content: 'Please review' },
        });
      });

      it('uses provided bodyType when updating comment', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-fwd-html', subject: 'HTML forward' },
        ]);
        await repository.searchEmailsAsync('HTML forward', 50);

        mockClient.createForwardDraft.mockResolvedValue({
          id: 'draft-fwd-html-1',
          subject: 'Fwd: HTML forward',
          toRecipients: [],
        });
        mockClient.updateDraft.mockResolvedValue(undefined);

        await repository.forwardAsDraftAsync(
          hashStringToNumber('msg-fwd-html'),
          undefined,
          '<p>HTML comment</p>',
          'html'
        );

        expect(mockClient.updateDraft).toHaveBeenCalledWith('draft-fwd-html-1', {
          body: { contentType: 'html', content: '<p>HTML comment</p>' },
        });
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.forwardAsDraftAsync(99999)
        ).rejects.toThrow('Message ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });
  });

  describe('Attachment Operations (Async)', () => {
    describe('listAttachmentsAsync', () => {
      it('lists attachments and caches them', async () => {
        // Populate message cache first
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-att-1', subject: 'With attachments' },
        ]);
        await repository.searchEmailsAsync('attachments', 50);

        mockClient.listAttachments.mockResolvedValue([
          { id: 'att-1', name: 'doc.pdf', size: 1024, contentType: 'application/pdf', isInline: false },
          { id: 'att-2', name: 'image.png', size: 2048, contentType: 'image/png', isInline: true },
        ]);

        const result = await repository.listAttachmentsAsync(hashStringToNumber('msg-att-1'));

        expect(mockClient.listAttachments).toHaveBeenCalledWith('msg-att-1');
        expect(result).toHaveLength(2);
        expect(result[0]).toEqual({
          id: hashStringToNumber('att-1'),
          name: 'doc.pdf',
          size: 1024,
          contentType: 'application/pdf',
          isInline: false,
        });
        expect(result[1]).toEqual({
          id: hashStringToNumber('att-2'),
          name: 'image.png',
          size: 2048,
          contentType: 'image/png',
          isInline: true,
        });
      });

      it('handles missing attachment fields with defaults', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-att-2', subject: 'Minimal' },
        ]);
        await repository.searchEmailsAsync('Minimal', 50);

        mockClient.listAttachments.mockResolvedValue([
          { id: null, name: null, size: null, contentType: null },
        ]);

        const result = await repository.listAttachmentsAsync(hashStringToNumber('msg-att-2'));

        expect(result).toHaveLength(1);
        expect(result[0].name).toBe('');
        expect(result[0].size).toBe(0);
        expect(result[0].contentType).toBe('application/octet-stream');
        expect(result[0].isInline).toBe(false);
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.listAttachmentsAsync(99999)
        ).rejects.toThrow('Message ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('downloadAttachmentAsync', () => {
      it('delegates to downloadAttachment helper with cached IDs', async () => {
        // Populate message cache
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-dl', subject: 'Download test' },
        ]);
        await repository.searchEmailsAsync('Download test', 50);

        // Populate attachment cache via listAttachmentsAsync
        mockClient.listAttachments.mockResolvedValue([
          { id: 'att-dl-1', name: 'file.zip', size: 5000, contentType: 'application/zip' },
        ]);
        await repository.listAttachmentsAsync(hashStringToNumber('msg-dl'));

        const mockResult = { filePath: '/tmp/file.zip', name: 'file.zip', size: 5000, contentType: 'application/zip' };
        vi.mocked(downloadAttachment).mockResolvedValue(mockResult);

        const result = await repository.downloadAttachmentAsync(hashStringToNumber('att-dl-1'));

        expect(downloadAttachment).toHaveBeenCalledWith(
          mockClient,
          'msg-dl',
          'att-dl-1'
        );
        expect(result).toEqual(mockResult);
      });

      it('throws if attachment not in cache', async () => {
        await expect(
          repository.downloadAttachmentAsync(99999)
        ).rejects.toThrow('Attachment ID 99999 not found in cache. Call list_attachments first.');
      });
    });
  });

  describe('Calendar Write Operations (Async)', () => {
    describe('createEventAsync', () => {
      it('creates an event with required fields and caches result', async () => {
        mockClient.createEvent.mockResolvedValue({
          id: 'event-new-1',
          subject: 'Team Meeting',
        });

        const numericId = await repository.createEventAsync({
          subject: 'Team Meeting',
          start: '2026-03-01T10:00:00',
          end: '2026-03-01T11:00:00',
          timezone: 'America/New_York',
        });

        expect(mockClient.createEvent).toHaveBeenCalledWith(
          {
            subject: 'Team Meeting',
            start: { dateTime: '2026-03-01T10:00:00', timeZone: 'America/New_York' },
            end: { dateTime: '2026-03-01T11:00:00', timeZone: 'America/New_York' },
          },
          undefined
        );

        expect(numericId).toBe(hashStringToNumber('event-new-1'));
        expect(repository.getGraphId('event', numericId)).toBe('event-new-1');
      });

      it('includes location when provided', async () => {
        mockClient.createEvent.mockResolvedValue({ id: 'event-loc' });

        await repository.createEventAsync({
          subject: 'Office Meeting',
          start: '2026-03-01T10:00:00',
          end: '2026-03-01T11:00:00',
          timezone: 'UTC',
          location: 'Conference Room B',
        });

        const callArgs = mockClient.createEvent.mock.calls[0][0];
        expect(callArgs.location).toEqual({ displayName: 'Conference Room B' });
      });

      it('includes body with type when provided', async () => {
        mockClient.createEvent.mockResolvedValue({ id: 'event-body' });

        await repository.createEventAsync({
          subject: 'Review',
          start: '2026-03-01T10:00:00',
          end: '2026-03-01T11:00:00',
          timezone: 'UTC',
          body: '<p>Agenda items</p>',
          bodyType: 'html',
        });

        const callArgs = mockClient.createEvent.mock.calls[0][0];
        expect(callArgs.body).toEqual({
          contentType: 'html',
          content: '<p>Agenda items</p>',
        });
      });

      it('defaults body type to text', async () => {
        mockClient.createEvent.mockResolvedValue({ id: 'event-body-text' });

        await repository.createEventAsync({
          subject: 'Review',
          start: '2026-03-01T10:00:00',
          end: '2026-03-01T11:00:00',
          timezone: 'UTC',
          body: 'Plain text agenda',
        });

        const callArgs = mockClient.createEvent.mock.calls[0][0];
        expect(callArgs.body.contentType).toBe('text');
      });

      it('includes attendees when provided', async () => {
        mockClient.createEvent.mockResolvedValue({ id: 'event-att' });

        await repository.createEventAsync({
          subject: 'Team Sync',
          start: '2026-03-01T10:00:00',
          end: '2026-03-01T11:00:00',
          timezone: 'UTC',
          attendees: [
            { email: 'alice@example.com', name: 'Alice', type: 'required' },
            { email: 'bob@example.com', type: 'optional' },
          ],
        });

        const callArgs = mockClient.createEvent.mock.calls[0][0];
        expect(callArgs.attendees).toEqual([
          { emailAddress: { address: 'alice@example.com', name: 'Alice' }, type: 'required' },
          { emailAddress: { address: 'bob@example.com', name: undefined }, type: 'optional' },
        ]);
      });

      it('sets isAllDay flag', async () => {
        mockClient.createEvent.mockResolvedValue({ id: 'event-allday' });

        await repository.createEventAsync({
          subject: 'Holiday',
          start: '2026-03-01',
          end: '2026-03-02',
          timezone: 'UTC',
          isAllDay: true,
        });

        const callArgs = mockClient.createEvent.mock.calls[0][0];
        expect(callArgs.isAllDay).toBe(true);
      });

      it('includes recurrence when provided', async () => {
        mockClient.createEvent.mockResolvedValue({ id: 'event-recur' });

        const recurrence = {
          pattern: { type: 'weekly' as const, interval: 1, daysOfWeek: ['monday'] },
          range: { type: 'noEnd' as const, startDate: '2026-03-01' },
        };

        await repository.createEventAsync({
          subject: 'Weekly Standup',
          start: '2026-03-01T09:00:00',
          end: '2026-03-01T09:15:00',
          timezone: 'UTC',
          recurrence,
        });

        const callArgs = mockClient.createEvent.mock.calls[0][0];
        expect(callArgs.recurrence).toEqual(recurrence);
      });

      it('passes calendarId when provided', async () => {
        // Populate folder cache with a calendar ID
        mockClient.listCalendars.mockResolvedValue([
          { id: 'cal-work', name: 'Work Calendar' },
        ]);
        await repository.listCalendarsAsync();

        mockClient.createEvent.mockResolvedValue({ id: 'event-cal' });

        await repository.createEventAsync({
          subject: 'Work Event',
          start: '2026-03-01T10:00:00',
          end: '2026-03-01T11:00:00',
          timezone: 'UTC',
          calendarId: hashStringToNumber('cal-work'),
        });

        expect(mockClient.createEvent).toHaveBeenCalledWith(
          expect.any(Object),
          'cal-work'
        );
      });
    });

    describe('updateEventAsync', () => {
      it('looks up graph ID and calls updateEvent', async () => {
        // Populate event cache
        mockClient.listEvents.mockResolvedValue([
          { id: 'event-upd', subject: 'Existing', start: {}, end: {} },
        ]);
        await repository.listEventsAsync(50, 0);

        mockClient.updateEvent.mockResolvedValue(undefined);

        await repository.updateEventAsync(hashStringToNumber('event-upd'), {
          subject: 'Updated Meeting',
        });

        expect(mockClient.updateEvent).toHaveBeenCalledWith('event-upd', {
          subject: 'Updated Meeting',
        });
      });

      it('throws if event not in cache', async () => {
        await expect(
          repository.updateEventAsync(99999, { subject: 'Nope' })
        ).rejects.toThrow('Event ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('deleteEventAsync', () => {
      it('deletes event and removes from cache', async () => {
        // Populate event cache
        mockClient.listEvents.mockResolvedValue([
          { id: 'event-del', subject: 'To Delete', start: {}, end: {} },
        ]);
        await repository.listEventsAsync(50, 0);

        mockClient.deleteEvent.mockResolvedValue(undefined);

        const numericId = hashStringToNumber('event-del');
        await repository.deleteEventAsync(numericId);

        expect(mockClient.deleteEvent).toHaveBeenCalledWith('event-del');
        expect(repository.getGraphId('event', numericId)).toBeUndefined();
      });

      it('throws if event not in cache', async () => {
        await expect(
          repository.deleteEventAsync(99999)
        ).rejects.toThrow('Event ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('respondToEventAsync', () => {
      it('responds to event with accept and comment', async () => {
        // Populate event cache
        mockClient.listEvents.mockResolvedValue([
          { id: 'event-resp', subject: 'Invitation', start: {}, end: {} },
        ]);
        await repository.listEventsAsync(50, 0);

        mockClient.respondToEvent.mockResolvedValue(undefined);

        await repository.respondToEventAsync(
          hashStringToNumber('event-resp'),
          'accept',
          true,
          'Looking forward to it!'
        );

        expect(mockClient.respondToEvent).toHaveBeenCalledWith(
          'event-resp',
          'accept',
          true,
          'Looking forward to it!'
        );
      });

      it('responds to event with decline without comment', async () => {
        // Populate event cache
        mockClient.listEvents.mockResolvedValue([
          { id: 'event-resp2', subject: 'Decline This', start: {}, end: {} },
        ]);
        await repository.listEventsAsync(50, 0);

        mockClient.respondToEvent.mockResolvedValue(undefined);

        await repository.respondToEventAsync(
          hashStringToNumber('event-resp2'),
          'decline',
          false
        );

        expect(mockClient.respondToEvent).toHaveBeenCalledWith(
          'event-resp2',
          'decline',
          false,
          undefined
        );
      });

      it('throws if event not in cache', async () => {
        await expect(
          repository.respondToEventAsync(99999, 'accept', true)
        ).rejects.toThrow('Event ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });
  });

  describe('Contact Write Operations (Async)', () => {
    describe('createContactAsync', () => {
      it('maps fields correctly and calls client.createContact', async () => {
        mockClient.createContact.mockResolvedValue({
          id: 'contact-new-1',
          displayName: 'John Doe',
        });

        const numericId = await repository.createContactAsync({
          given_name: 'John',
          surname: 'Doe',
          email: 'john@example.com',
          phone: '+1234567890',
          mobile_phone: '+0987654321',
          company: 'Acme Inc',
          job_title: 'Engineer',
          street_address: '123 Main St',
          city: 'Springfield',
          state: 'IL',
          postal_code: '62704',
          country: 'US',
        });

        expect(mockClient.createContact).toHaveBeenCalledWith({
          givenName: 'John',
          surname: 'Doe',
          emailAddresses: [{ address: 'john@example.com' }],
          businessPhones: ['+1234567890'],
          mobilePhone: '+0987654321',
          companyName: 'Acme Inc',
          jobTitle: 'Engineer',
          businessAddress: {
            street: '123 Main St',
            city: 'Springfield',
            state: 'IL',
            postalCode: '62704',
            countryOrRegion: 'US',
          },
        });

        expect(numericId).toBe(hashStringToNumber('contact-new-1'));
      });

      it('adds result to idCache', async () => {
        mockClient.createContact.mockResolvedValue({
          id: 'contact-new-2',
          displayName: 'Jane',
        });

        const numericId = await repository.createContactAsync({
          given_name: 'Jane',
        });

        const graphId = repository.getGraphId('contact', numericId);
        expect(graphId).toBe('contact-new-2');
      });

      it('handles minimal fields (only given_name)', async () => {
        mockClient.createContact.mockResolvedValue({
          id: 'contact-min',
          displayName: 'Min',
        });

        await repository.createContactAsync({ given_name: 'Min' });

        expect(mockClient.createContact).toHaveBeenCalledWith({
          givenName: 'Min',
        });
      });

      it('does not include businessAddress when no address fields provided', async () => {
        mockClient.createContact.mockResolvedValue({
          id: 'contact-no-addr',
          displayName: 'No Address',
        });

        await repository.createContactAsync({
          given_name: 'No',
          surname: 'Address',
        });

        const callArgs = mockClient.createContact.mock.calls[0][0];
        expect(callArgs).not.toHaveProperty('businessAddress');
      });
    });

    describe('updateContactAsync', () => {
      it('looks up graph ID and calls client.updateContact', async () => {
        // Populate contact cache
        mockClient.listContacts.mockResolvedValue([
          { id: 'contact-1', displayName: 'Existing Contact' },
        ]);
        await repository.listContactsAsync(50, 0);

        mockClient.updateContact.mockResolvedValue(undefined);

        await repository.updateContactAsync(hashStringToNumber('contact-1'), {
          givenName: 'Updated',
        });

        expect(mockClient.updateContact).toHaveBeenCalledWith('contact-1', {
          givenName: 'Updated',
        });
      });

      it('throws if contact not in cache', async () => {
        await expect(
          repository.updateContactAsync(99999, { givenName: 'Nope' })
        ).rejects.toThrow('Contact ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('deleteContactAsync', () => {
      it('calls client.deleteContact and removes from idCache', async () => {
        // Populate contact cache
        mockClient.listContacts.mockResolvedValue([
          { id: 'contact-del', displayName: 'To Delete' },
        ]);
        await repository.listContactsAsync(50, 0);

        mockClient.deleteContact.mockResolvedValue(undefined);

        const numericId = hashStringToNumber('contact-del');
        await repository.deleteContactAsync(numericId);

        expect(mockClient.deleteContact).toHaveBeenCalledWith('contact-del');

        // Verify it was removed from cache
        const graphId = repository.getGraphId('contact', numericId);
        expect(graphId).toBeUndefined();
      });

      it('throws if contact not in cache', async () => {
        await expect(
          repository.deleteContactAsync(99999)
        ).rejects.toThrow('Contact ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });
  });

  describe('Task Write Operations (Async)', () => {
    describe('taskLists cache population', () => {
      it('populates taskLists cache in listTasksAsync', async () => {
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'Task 1' },
          { id: 'task-2', taskListId: 'list-2', title: 'Task 2' },
        ]);

        await repository.listTasksAsync(50, 0);

        // Verify taskLists cache was populated
        const idCache = (repository as any).idCache;
        expect(idCache.taskLists.get(hashStringToNumber('list-1'))).toBe('list-1');
        expect(idCache.taskLists.get(hashStringToNumber('list-2'))).toBe('list-2');
      });

      it('populates taskLists cache in listIncompleteTasksAsync', async () => {
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-3', taskListId: 'list-3', title: 'Task 3' },
        ]);

        await repository.listIncompleteTasksAsync(50, 0);

        const idCache = (repository as any).idCache;
        expect(idCache.taskLists.get(hashStringToNumber('list-3'))).toBe('list-3');
      });
    });

    describe('createTaskAsync', () => {
      it('creates a task with all fields and caches the result', async () => {
        // Populate taskLists cache
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-existing', taskListId: 'list-1', title: 'Existing' },
        ]);
        await repository.listTasksAsync(50, 0);

        mockClient.createTask.mockResolvedValue({
          id: 'task-new-1',
          title: 'New Task',
        });

        const listNumericId = hashStringToNumber('list-1');
        const numericId = await repository.createTaskAsync({
          title: 'New Task',
          task_list_id: listNumericId,
          body: 'Some notes',
          body_type: 'text',
          due_date: '2026-03-01T00:00:00Z',
          importance: 'high',
          reminder_date: '2026-02-28T09:00:00Z',
        });

        expect(mockClient.createTask).toHaveBeenCalledWith('list-1', {
          title: 'New Task',
          body: { contentType: 'text', content: 'Some notes' },
          dueDateTime: { dateTime: '2026-03-01T00:00:00Z', timeZone: 'UTC' },
          importance: 'high',
          isReminderOn: true,
          reminderDateTime: { dateTime: '2026-02-28T09:00:00Z', timeZone: 'UTC' },
        });

        expect(numericId).toBe(hashStringToNumber('task-new-1'));

        // Verify cached in tasks
        const taskInfo = repository.getTaskInfo(numericId);
        expect(taskInfo).toEqual({ taskListId: 'list-1', taskId: 'task-new-1' });
      });

      it('creates a task with only required fields', async () => {
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-existing', taskListId: 'list-1', title: 'Existing' },
        ]);
        await repository.listTasksAsync(50, 0);

        mockClient.createTask.mockResolvedValue({
          id: 'task-min',
          title: 'Minimal Task',
        });

        const listNumericId = hashStringToNumber('list-1');
        await repository.createTaskAsync({
          title: 'Minimal Task',
          task_list_id: listNumericId,
        });

        expect(mockClient.createTask).toHaveBeenCalledWith('list-1', {
          title: 'Minimal Task',
        });
      });

      it('throws when task list ID not in cache', async () => {
        await expect(
          repository.createTaskAsync({
            title: 'Test',
            task_list_id: 99999,
          })
        ).rejects.toThrow('Task list ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('updateTaskAsync', () => {
      it('looks up task info and calls client.updateTask', async () => {
        // Populate task cache
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'Old Title' },
        ]);
        await repository.listTasksAsync(50, 0);

        mockClient.updateTask.mockResolvedValue(undefined);

        await repository.updateTaskAsync(hashStringToNumber('task-1'), {
          title: 'New Title',
        });

        expect(mockClient.updateTask).toHaveBeenCalledWith('list-1', 'task-1', {
          title: 'New Title',
        });
      });

      it('throws if task not in cache', async () => {
        await expect(
          repository.updateTaskAsync(99999, { title: 'Nope' })
        ).rejects.toThrow('Task ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('completeTaskAsync', () => {
      it('calls updateTaskAsync with completed status', async () => {
        // Populate task cache
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'To Complete' },
        ]);
        await repository.listTasksAsync(50, 0);

        mockClient.updateTask.mockResolvedValue(undefined);

        await repository.completeTaskAsync(hashStringToNumber('task-1'));

        expect(mockClient.updateTask).toHaveBeenCalledWith('list-1', 'task-1', {
          status: 'completed',
          completedDateTime: {
            dateTime: expect.any(String),
            timeZone: 'UTC',
          },
        });
      });

      it('throws if task not in cache', async () => {
        await expect(
          repository.completeTaskAsync(99999)
        ).rejects.toThrow('Task ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('deleteTaskAsync', () => {
      it('calls client.deleteTask and removes from idCache', async () => {
        // Populate task cache
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-del', taskListId: 'list-1', title: 'To Delete' },
        ]);
        await repository.listTasksAsync(50, 0);

        mockClient.deleteTask.mockResolvedValue(undefined);

        const numericId = hashStringToNumber('task-del');
        await repository.deleteTaskAsync(numericId);

        expect(mockClient.deleteTask).toHaveBeenCalledWith('list-1', 'task-del');

        // Verify it was removed from cache
        const taskInfo = repository.getTaskInfo(numericId);
        expect(taskInfo).toBeUndefined();
      });

      it('throws if task not in cache', async () => {
        await expect(
          repository.deleteTaskAsync(99999)
        ).rejects.toThrow('Task ID 99999 not found in cache. Try searching for or listing the item first to refresh the cache.');
      });
    });

    describe('createTaskListAsync', () => {
      it('creates task list and caches the result', async () => {
        mockClient.createTaskList.mockResolvedValue({
          id: 'new-list-1',
          displayName: 'My New List',
        });

        const numericId = await repository.createTaskListAsync('My New List');

        expect(mockClient.createTaskList).toHaveBeenCalledWith('My New List');
        expect(numericId).toBe(hashStringToNumber('new-list-1'));

        // Verify it was cached in taskLists
        const idCache = (repository as any).idCache;
        expect(idCache.taskLists.get(numericId)).toBe('new-list-1');
      });
    });
  });

  describe('Calendar Scheduling', () => {
    describe('getScheduleAsync', () => {
      it('calls client.getSchedule with formatted params and returns result', async () => {
        const mockSchedules = [
          { scheduleId: 'bob@example.com', availabilityView: '0120', scheduleItems: [] },
        ];
        mockClient.getSchedule.mockResolvedValue(mockSchedules);

        const result = await repository.getScheduleAsync({
          emailAddresses: ['bob@example.com'],
          startTime: '2026-02-24T08:00:00Z',
          endTime: '2026-02-24T18:00:00Z',
          availabilityViewInterval: 30,
        });

        expect(mockClient.getSchedule).toHaveBeenCalledWith({
          schedules: ['bob@example.com'],
          startTime: { dateTime: '2026-02-24T08:00:00Z', timeZone: 'UTC' },
          endTime: { dateTime: '2026-02-24T18:00:00Z', timeZone: 'UTC' },
          availabilityViewInterval: 30,
        });
        expect(result).toEqual(mockSchedules);
      });

      it('uses default interval of 30 when not specified', async () => {
        mockClient.getSchedule.mockResolvedValue([]);

        await repository.getScheduleAsync({
          emailAddresses: ['bob@example.com'],
          startTime: '2026-02-24T08:00:00Z',
          endTime: '2026-02-24T18:00:00Z',
        });

        expect(mockClient.getSchedule).toHaveBeenCalledWith(
          expect.objectContaining({ availabilityViewInterval: 30 })
        );
      });
    });

    describe('findMeetingTimesAsync', () => {
      it('calls client.findMeetingTimes with formatted attendees and ISO duration', async () => {
        const mockResult = {
          meetingTimeSuggestions: [{ confidence: 100 }],
          emptySuggestionsReason: '',
        };
        mockClient.findMeetingTimes.mockResolvedValue(mockResult);

        const result = await repository.findMeetingTimesAsync({
          attendees: ['bob@example.com', 'alice@example.com'],
          durationMinutes: 60,
          startTime: '2026-02-24T08:00:00Z',
          endTime: '2026-02-24T18:00:00Z',
          maxCandidates: 5,
        });

        expect(mockClient.findMeetingTimes).toHaveBeenCalledWith({
          attendees: [
            { emailAddress: { address: 'bob@example.com' }, type: 'required' },
            { emailAddress: { address: 'alice@example.com' }, type: 'required' },
          ],
          meetingDuration: 'PT1H0M',
          timeConstraint: {
            timeslots: [{
              start: { dateTime: '2026-02-24T08:00:00Z', timeZone: 'UTC' },
              end: { dateTime: '2026-02-24T18:00:00Z', timeZone: 'UTC' },
            }],
          },
          maxCandidates: 5,
        });
        expect(result).toEqual(mockResult);
      });

      it('omits timeConstraint when startTime/endTime not provided', async () => {
        mockClient.findMeetingTimes.mockResolvedValue({ meetingTimeSuggestions: [] });

        await repository.findMeetingTimesAsync({
          attendees: ['bob@example.com'],
          durationMinutes: 30,
        });

        expect(mockClient.findMeetingTimes).toHaveBeenCalledWith({
          attendees: [{ emailAddress: { address: 'bob@example.com' }, type: 'required' }],
          meetingDuration: 'PT0H30M',
          maxCandidates: 5,
        });
      });
    });
  });

  describe('Utility Methods', () => {
    describe('getClient', () => {
      it('returns the GraphClient instance', () => {
        const client = repository.getClient();
        expect(client).toBeDefined();
      });
    });

    describe('getGraphId', () => {
      it('returns undefined when ID not cached', () => {
        expect(repository.getGraphId('folder', 99999)).toBeUndefined();
        expect(repository.getGraphId('message', 99999)).toBeUndefined();
        expect(repository.getGraphId('event', 99999)).toBeUndefined();
        expect(repository.getGraphId('contact', 99999)).toBeUndefined();
      });
    });

    describe('getTaskInfo', () => {
      it('returns undefined when task ID not cached', () => {
        expect(repository.getTaskInfo(99999)).toBeUndefined();
      });

      it('returns task info when cached', async () => {
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'Task 1' },
        ]);
        await repository.listTasksAsync(50, 0);

        const info = repository.getTaskInfo(hashStringToNumber('task-1'));

        expect(info).toEqual({ taskListId: 'list-1', taskId: 'task-1' });
      });
    });
  });
});
