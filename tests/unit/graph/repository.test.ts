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
    };
  }),
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
          'Message ID 99999 not found in cache'
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
        ).rejects.toThrow('Folder ID 99999 not found in cache');
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
          'Message ID 99999 not found in cache'
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

        expect(result).toBe(hashStringToNumber('draft-1'));

        // Verify cached
        const graphId = repository.getGraphId('message', result);
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

        expect(result).toBe(hashStringToNumber('draft-2'));
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
        ).rejects.toThrow('Message ID 99999 not found in cache');
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
          'Message ID 99999 not found in cache'
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
        ).rejects.toThrow('Message ID 99999 not found in cache');
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
        ).rejects.toThrow('Message ID 99999 not found in cache');
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
