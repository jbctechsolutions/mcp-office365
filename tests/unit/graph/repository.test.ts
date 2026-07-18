/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Graph API repository.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { GraphRepository, createGraphRepository } from '../../../src/graph/repository.js';
import { mintSelfEncoded } from '../../../src/ids/token.js';
import { StateStore } from '../../../src/state/store.js';
import { downloadAttachment } from '../../../src/graph/attachments.js';
import * as fs from 'fs';
import * as path from 'path';

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
      searchMessagesFilter: vi.fn(),
      searchMessagesSearchValue: vi.fn(),
      searchMessagesQuery: vi.fn(),
      listConversationMessages: vi.fn(),
      getMessagesDelta: vi.fn(),
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
      // Calendar instance operations
      listEventInstances: vi.fn(),
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
      updateTaskList: vi.fn(),
      deleteTaskList: vi.fn(),
      // Calendar scheduling operations
      getSchedule: vi.fn(),
      findMeetingTimes: vi.fn(),
      // Automatic replies operations
      getAutomaticReplies: vi.fn(),
      setAutomaticReplies: vi.fn(),
      // Mail rules operations
      listMailRules: vi.fn(),
      createMailRule: vi.fn(),
      deleteMailRule: vi.fn(),
      // Master categories operations
      listMasterCategories: vi.fn(),
      createMasterCategory: vi.fn(),
      deleteMasterCategory: vi.fn(),
      // Focused inbox override operations
      listFocusedOverrides: vi.fn(),
      createFocusedOverride: vi.fn(),
      deleteFocusedOverride: vi.fn(),
      // Contact folder operations
      listContactFolders: vi.fn(),
      createContactFolder: vi.fn(),
      deleteContactFolder: vi.fn(),
      listContactsInFolder: vi.fn(),
      // Contact photo operations
      getContactPhoto: vi.fn(),
      setContactPhoto: vi.fn(),
      // Message headers & MIME operations
      getMessageHeaders: vi.fn(),
      getMessageMime: vi.fn(),
      // Mail tips operations
      getMailTips: vi.fn(),
      // Calendar group operations
      listCalendarGroups: vi.fn(),
      createCalendarGroup: vi.fn(),
      // Calendar permission operations
      listCalendarPermissions: vi.fn(),
      createCalendarPermission: vi.fn(),
      deleteCalendarPermission: vi.fn(),
      // Room lists & rooms operations
      listRoomLists: vi.fn(),
      listRooms: vi.fn(),
      // Teams operations
      listJoinedTeams: vi.fn(),
      listChannels: vi.fn(),
      getChannel: vi.fn(),
      createChannel: vi.fn(),
      updateChannel: vi.fn(),
      deleteChannel: vi.fn(),
      listTeamMembers: vi.fn(),
      // Teams channel message operations
      listChannelMessages: vi.fn(),
      getChannelMessage: vi.fn(),
      listChannelMessageReplies: vi.fn(),
      sendChannelMessage: vi.fn(),
      replyToChannelMessage: vi.fn(),
      // Teams chat operations
      listChats: vi.fn(),
      createChat: vi.fn(),
      getChat: vi.fn(),
      listChatMessages: vi.fn(),
      sendChatMessage: vi.fn(),
      listChatMembers: vi.fn(),
      getChatMessage: vi.fn(),
      // Teams message reaction operations
      setChannelMessageReaction: vi.fn(),
      unsetChannelMessageReaction: vi.fn(),
      setChatMessageReaction: vi.fn(),
      unsetChatMessageReaction: vi.fn(),
      // Planner operations
      listPlans: vi.fn(),
      getPlan: vi.fn(),
      createPlan: vi.fn(),
      updatePlan: vi.fn(),
      listBuckets: vi.fn(),
      createBucket: vi.fn(),
      updateBucket: vi.fn(),
      deleteBucket: vi.fn(),
      getBucket: vi.fn(),
      listPlannerTasks: vi.fn(),
      listMyPlannerTasks: vi.fn(),
      getPlannerTask: vi.fn(),
      createPlannerTask: vi.fn(),
      updatePlannerTask: vi.fn(),
      deletePlannerTask: vi.fn(),
      getPlannerTaskDetails: vi.fn(),
      updatePlannerTaskDetails: vi.fn(),
      // OneDrive operations
      listDriveItems: vi.fn(),
      searchDriveItems: vi.fn(),
      listRecentDriveItems: vi.fn(),
      listSharedWithMe: vi.fn(),
      // Online meeting operations
      listOnlineMeetings: vi.fn(),
      getOnlineMeeting: vi.fn(),
      listMeetingRecordings: vi.fn(),
      getMeetingRecordingContent: vi.fn(),
      listMeetingTranscripts: vi.fn(),
      getMeetingTranscriptContent: vi.fn(),
      // SharePoint sites & document library operations
      listFollowedSites: vi.fn(),
      searchSites: vi.fn(),
      getSite: vi.fn(),
      listDocumentLibraries: vi.fn(),
      listLibraryItems: vi.fn(),
      downloadLibraryFile: vi.fn(),
      listSharePointLists: vi.fn(),
      getSharePointList: vi.fn(),
      createSharePointList: vi.fn(),
      listSharePointListColumns: vi.fn(),
      listSharePointListItems: vi.fn(),
      getSharePointListItem: vi.fn(),
      createSharePointListItem: vi.fn(),
      updateSharePointListItem: vi.fn(),
      deleteSharePointListItem: vi.fn(),
    };
  }),
}));

// Mock the downloadAttachment helper and getDownloadDir
vi.mock('../../../src/graph/attachments.js', () => ({
  downloadAttachment: vi.fn(),
  getDownloadDir: vi.fn().mockReturnValue('/tmp/mcp-outlook-attachments'),
}));

// Mock fs and path for contact photo tests
vi.mock('fs', () => ({
  writeFileSync: vi.fn(),
  readFileSync: vi.fn().mockReturnValue(Buffer.from('fake-photo')),
  mkdirSync: vi.fn(),
  existsSync: vi.fn().mockReturnValue(false),
}));

vi.mock('path', () => ({
  join: vi.fn().mockImplementation((...args: string[]) => args.join('/')),
  extname: vi.fn().mockImplementation((p: string) => {
    const dot = p.lastIndexOf('.');
    return dot >= 0 ? p.substring(dot) : '';
  }),
  resolve: vi.fn().mockImplementation((p: string) => p),
  dirname: vi.fn().mockImplementation((p: string) => p.substring(0, p.lastIndexOf('/')) || '/'),
}));

describe('graph/repository', () => {
  let repository: GraphRepository;
  let mockClient: any;

  beforeEach(async () => {
    vi.clearAllMocks();
    // Alias-backed entities (teams, channels, …) need a store to resolve their
    // tokens. fs is mocked in this suite, so StateStore.open degrades to an
    // in-memory sqlite db — still a fully-functional alias table for the run.
    const store = StateStore.open({ dir: '/tmp/mcp-o365-repo-test', warn: () => {} });
    repository = createGraphRepository(undefined, store);
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
        // Rows carry self-encoding fd_ tokens (the mapper mints them, U5).
        expect(result[0].id).toBe(mintSelfEncoded('folder', 'folder-1'));
        expect(result[0].name).toBe('Inbox');
        expect(result[1].id).toBe(mintSelfEncoded('folder', 'folder-2'));
      });

      it('emits durable fd_ tokens for returned folders (no cache)', async () => {
        mockClient.listMailFolders.mockResolvedValue([
          { id: 'folder-1', displayName: 'Inbox' },
        ]);

        const result = await repository.listFoldersAsync();

        // Cold resolve via the token itself — no prior list/cache needed.
        expect(repository.getFolderGraphId(result[0].id)).toBe('folder-1');
      });
    });

    describe('getFolder (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.getFolder('folder-1')).toThrow('Use getFolderAsync()');
      });
    });

    describe('getFolderAsync', () => {
      it('returns folder by durable fd_ token', async () => {
        mockClient.getMailFolder.mockResolvedValue({
          id: 'folder-1',
          displayName: 'Inbox',
          totalItemCount: 100,
        });

        const result = await repository.getFolderAsync(mintSelfEncoded('folder', 'folder-1'));

        expect(result?.name).toBe('Inbox');
      });

      it('returns undefined when folder not found', async () => {
        mockClient.getMailFolder.mockResolvedValue(null);

        const result = await repository.getFolderAsync(mintSelfEncoded('folder', 'nonexistent'));

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
      it('resolves a folder fd_ token to the Graph id — no prior list/cache needed (cold state)', async () => {
        mockClient.listMessages.mockResolvedValue([
          { id: 'msg-1', subject: 'Test Email', isRead: true },
          { id: 'msg-2', subject: 'Another Email', isRead: false },
        ]);

        const result = await repository.listEmailsAsync(mintSelfEncoded('folder', 'folder-1'), 50, 0);

        expect(mockClient.listMessages).toHaveBeenCalledWith('folder-1', 50, 0);
        expect(result).toHaveLength(2);
        expect(result[0].subject).toBe('Test Email');
        expect(result[1].subject).toBe('Another Email');
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
      it('resolves an em_ token to the Graph id — no prior list/cache needed (cold state)', async () => {
        mockClient.getMessage.mockResolvedValue({
          id: 'msg-1',
          subject: 'Test Email',
          body: { content: 'Body content' },
        });

        const result = await repository.getEmailAsync(mintSelfEncoded('message', 'msg-1'));

        expect(mockClient.getMessage).toHaveBeenCalledWith('msg-1');
        expect(result?.subject).toBe('Test Email');
      });

      it('rejects a legacy numeric id on the Graph backend (NUMERIC_ID_UNSUPPORTED, D4)', async () => {
        await expect(repository.getEmailAsync(99999)).rejects.toMatchObject({
          code: 'NUMERIC_ID_UNSUPPORTED',
        });
      });

      it('rejects a token for a different entity kind (ID_ENTITY_MISMATCH)', async () => {
        await expect(
          repository.getEmailAsync(mintSelfEncoded('contact', 'contact-1'))
        ).rejects.toMatchObject({ code: 'ID_ENTITY_MISMATCH' });
        expect(mockClient.getMessage).not.toHaveBeenCalled();
      });

      it('returns undefined when the message is not found', async () => {
        mockClient.getMessage.mockResolvedValue(null);

        const result = await repository.getEmailAsync(mintSelfEncoded('message', 'msg-1'));
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
      it('resolves a folder fd_ token — no prior list/cache needed (cold state)', async () => {
        mockClient.getMailFolder.mockResolvedValue({
          id: 'folder-1',
          displayName: 'Inbox',
          unreadItemCount: 10,
        });

        const result = await repository.getUnreadCountByFolderAsync(mintSelfEncoded('folder', 'folder-1'));

        expect(result).toBe(10);
        expect(mockClient.getMailFolder).toHaveBeenCalledWith('folder-1');
      });

      it('returns 0 when getMailFolder returns null', async () => {
        mockClient.getMailFolder.mockResolvedValue(null);

        const result = await repository.getUnreadCountByFolderAsync(mintSelfEncoded('folder', 'folder-1'));

        expect(result).toBe(0);
      });
    });

    describe('searchEmailsInFolder (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.searchEmailsInFolder(1, 'query', 50)).toThrow('Use searchEmailsInFolderAsync()');
      });
    });

    describe('searchEmailsInFolderAsync', () => {
      it('resolves a folder fd_ token — no prior list/cache needed (cold state)', async () => {
        mockClient.searchMessagesInFolder.mockResolvedValue([
          { id: 'msg-1', subject: 'Match' },
        ]);

        const result = await repository.searchEmailsInFolderAsync(
          mintSelfEncoded('folder', 'folder-1'),
          'Match',
          50
        );

        expect(result).toHaveLength(1);
        expect(mockClient.searchMessagesInFolder).toHaveBeenCalledWith('folder-1', 'Match', 50);
      });
    });

    describe('searchEmailsStructuredAsync (U7)', () => {
      it('dispatches a filter mechanism to searchMessagesFilter and caches results', async () => {
        mockClient.searchMessagesFilter.mockResolvedValue([
          { id: 'msg-f-1', subject: 'Filter result', conversationId: 'conv-f' },
        ]);
        const results = await repository.searchEmailsStructuredAsync(
          { mechanism: 'filter', filter: "from/emailAddress/address eq 'a@b.com'" },
          50,
        );
        expect(results).toHaveLength(1);
        expect(mockClient.searchMessagesFilter).toHaveBeenCalledWith(
          "from/emailAddress/address eq 'a@b.com'",
          50,
        );
        expect(results[0].subject).toBe('Filter result');
      });

      it('dispatches a search mechanism to searchMessagesSearchValue', async () => {
        mockClient.searchMessagesSearchValue.mockResolvedValue([{ id: 'msg-s-1', subject: 'Search result' }]);
        await repository.searchEmailsStructuredAsync({ mechanism: 'search', search: '"budget"' }, 25);
        expect(mockClient.searchMessagesSearchValue).toHaveBeenCalledWith('"budget"', 25);
      });

      it('dispatches a searchQuery (mixed) mechanism to searchMessagesQuery', async () => {
        mockClient.searchMessagesQuery.mockResolvedValue([{ id: 'msg-q-1', subject: 'KQL result' }]);
        await repository.searchEmailsStructuredAsync(
          { mechanism: 'searchQuery', kql: 'from:"a@b.com" AND body:"x"' },
          10,
        );
        expect(mockClient.searchMessagesQuery).toHaveBeenCalledWith('from:"a@b.com" AND body:"x"', 10);
      });
    });
  });

  describe('Conversation / Thread', () => {
    describe('listConversationAsync', () => {
      it('lists messages in a conversation thread', async () => {
        // Resolve the em_ token → getMessage reads the raw conversationId → query.
        mockClient.getMessage.mockResolvedValue({
          id: 'msg-1', subject: 'Thread start', conversationId: 'conv-abc-123',
        });

        mockClient.listConversationMessages.mockResolvedValue([
          { id: 'msg-1', subject: 'Thread start', conversationId: 'conv-abc-123' },
          { id: 'msg-2', subject: 'Re: Thread start', conversationId: 'conv-abc-123' },
        ]);

        const result = await repository.listConversationAsync(mintSelfEncoded('message', 'msg-1'), 25);
        expect(mockClient.getMessage).toHaveBeenCalledWith('msg-1');
        expect(result).toHaveLength(2);
        expect(mockClient.listConversationMessages).toHaveBeenCalledWith('conv-abc-123', 25);
      });

      it('throws when message not found', async () => {
        mockClient.getMessage.mockResolvedValue(null);
        await expect(repository.listConversationAsync(mintSelfEncoded('message', 'missing'), 25))
          .rejects.toThrow('Message not found');
      });

      it('throws when message has no conversation ID', async () => {
        mockClient.getMessage.mockResolvedValue({
          id: 'msg-no-conv', subject: 'No conv', conversationId: undefined,
        });

        await expect(repository.listConversationAsync(mintSelfEncoded('message', 'msg-no-conv'), 25))
          .rejects.toThrow('no conversation ID');
      });
    });
  });

  describe('Delta Sync', () => {
    describe('checkNewEmailsAsync', () => {
      it('returns isInitialSync true on first call — resolves a folder fd_ token cold (no cache)', async () => {
        mockClient.getMessagesDelta = vi.fn().mockResolvedValue({
          messages: [
            { id: 'msg-delta-1', subject: 'New Email', conversationId: 'conv-d1' },
          ],
          deltaLink: 'https://graph.microsoft.com/v1.0/delta-token',
        });

        const folderId = mintSelfEncoded('folder', 'folder-delta');
        const result = await repository.checkNewEmailsAsync(folderId);

        expect(result.isInitialSync).toBe(true);
        expect(result.emails).toHaveLength(1);
        expect(result.emails[0].subject).toBe('New Email');
        expect(mockClient.getMessagesDelta).toHaveBeenCalledWith('folder-delta', undefined);
      });

      it('returns isInitialSync false on subsequent calls and passes delta link', async () => {
        const deltaToken = 'https://graph.microsoft.com/v1.0/delta-token-1';
        mockClient.getMessagesDelta = vi.fn().mockResolvedValue({
          messages: [{ id: 'msg-d1', subject: 'First' }],
          deltaLink: deltaToken,
        });

        const folderId = mintSelfEncoded('folder', 'folder-delta2');
        await repository.checkNewEmailsAsync(folderId);

        // Second call
        mockClient.getMessagesDelta.mockResolvedValue({
          messages: [{ id: 'msg-d2', subject: 'Second' }],
          deltaLink: 'https://graph.microsoft.com/v1.0/delta-token-2',
        });

        const result = await repository.checkNewEmailsAsync(folderId);
        expect(result.isInitialSync).toBe(false);
        expect(mockClient.getMessagesDelta).toHaveBeenCalledWith('folder-delta2', deltaToken);
      });

      it('shares one delta cursor for the same folder addressed by fd_ token OR raw Graph id', async () => {
        const deltaToken = 'https://graph.microsoft.com/v1.0/delta-token-shared';
        mockClient.getMessagesDelta = vi.fn().mockResolvedValue({
          messages: [{ id: 'msg-s1', subject: 'First' }],
          deltaLink: deltaToken,
        });

        // First call with the fd_ token establishes the cursor.
        await repository.checkNewEmailsAsync(mintSelfEncoded('folder', 'folder-shared'));

        // Second call with the RAW Graph id for the same folder must reuse that
        // cursor (keyed by the resolved Graph id) — not trigger a fresh sync.
        mockClient.getMessagesDelta.mockResolvedValue({ messages: [], deltaLink: deltaToken });
        const result = await repository.checkNewEmailsAsync('folder-shared');

        expect(result.isInitialSync).toBe(false);
        expect(mockClient.getMessagesDelta).toHaveBeenLastCalledWith('folder-shared', deltaToken);
      });

      it('rejects a legacy numeric folder id on the Graph backend (NUMERIC_ID_UNSUPPORTED, D4)', async () => {
        await expect(repository.checkNewEmailsAsync(99999 as unknown as string))
          .rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });

      it('filters out @removed messages', async () => {
        mockClient.getMessagesDelta = vi.fn().mockResolvedValue({
          messages: [
            { id: 'msg-active', subject: 'Active' },
            { id: 'msg-removed', subject: 'Removed', '@removed': { reason: 'deleted' } },
          ],
          deltaLink: 'https://graph.microsoft.com/v1.0/delta-rm',
        });

        const folderId = mintSelfEncoded('folder', 'folder-rm');
        const result = await repository.checkNewEmailsAsync(folderId);
        expect(result.emails).toHaveLength(1);
        expect(result.emails[0].subject).toBe('Active');
      });

      it('emits durable em_ tokens for returned messages (no cache)', async () => {
        mockClient.getMessagesDelta = vi.fn().mockResolvedValue({
          messages: [
            { id: 'msg-cache-1', subject: 'Cached', conversationId: 'conv-cache-1' },
          ],
          deltaLink: 'https://graph.microsoft.com/v1.0/delta-cache',
        });

        const folderId = mintSelfEncoded('folder', 'folder-cache');
        const result = await repository.checkNewEmailsAsync(folderId);

        expect(result.emails[0].id).toBe(mintSelfEncoded('message', 'msg-cache-1'));
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
      it('resolves an ev_ token to the Graph id — no prior list/cache needed (cold state)', async () => {
        mockClient.getEvent.mockResolvedValue({
          id: 'evt-1',
          subject: 'Team Meeting',
          start: { dateTime: '2024-01-15T10:00:00' },
        });

        const result = await repository.getEventAsync(mintSelfEncoded('event', 'evt-1'));

        expect(mockClient.getEvent).toHaveBeenCalledWith('evt-1');
        expect(result).toBeDefined();
      });

      it('rejects a legacy numeric id on the Graph backend (NUMERIC_ID_UNSUPPORTED, D4)', async () => {
        await expect(repository.getEventAsync(99999)).rejects.toMatchObject({
          code: 'NUMERIC_ID_UNSUPPORTED',
        });
      });

      it('rejects a token for a different entity kind (ID_ENTITY_MISMATCH)', async () => {
        await expect(
          repository.getEventAsync(mintSelfEncoded('contact', 'contact-1'))
        ).rejects.toMatchObject({ code: 'ID_ENTITY_MISMATCH' });
        expect(mockClient.getEvent).not.toHaveBeenCalled();
      });

      it('returns undefined when the event is not found', async () => {
        mockClient.getEvent.mockResolvedValue(null);

        const result = await repository.getEventAsync(mintSelfEncoded('event', 'evt-1'));
        expect(result).toBeUndefined();
      });
    });

    describe('listEventsByFolder (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.listEventsByFolder(1, 50)).toThrow('Use listEventsByFolderAsync()');
      });
    });

    describe('listEventsByFolderAsync', () => {
      it('resolves a calendar fd_ token — no prior list/cache needed (cold state)', async () => {
        mockClient.listEvents.mockResolvedValue([
          { id: 'evt-1', subject: 'Work Meeting' },
        ]);

        const result = await repository.listEventsByFolderAsync(mintSelfEncoded('folder', 'cal-1'), 50);

        expect(result).toHaveLength(1);
        expect(mockClient.listEvents).toHaveBeenCalledWith(50, 'cal-1');
      });

      it('rejects a legacy numeric calendar id on the Graph backend (NUMERIC_ID_UNSUPPORTED, D4)', async () => {
        await expect(repository.listEventsByFolderAsync(99999 as unknown as string, 50))
          .rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });

    describe('listEventInstancesAsync', () => {
      it('resolves the recurring-event token and returns mapped instances (cold)', async () => {
        mockClient.listEventInstances.mockResolvedValue([
          { id: 'inst-1', subject: 'Weekly Standup', start: { dateTime: '2024-01-08T10:00:00' } },
          { id: 'inst-2', subject: 'Weekly Standup', start: { dateTime: '2024-01-15T10:00:00' } },
        ]);

        const result = await repository.listEventInstancesAsync(
          mintSelfEncoded('event', 'evt-recurring'),
          '2024-01-01T00:00:00Z',
          '2024-01-31T23:59:59Z'
        );

        expect(result).toHaveLength(2);
        expect(mockClient.listEventInstances).toHaveBeenCalledWith(
          'evt-recurring',
          '2024-01-01T00:00:00Z',
          '2024-01-31T23:59:59Z'
        );
      });

      it('returns instances whose ids are durable ev_ tokens that resolve cold', async () => {
        mockClient.listEventInstances.mockResolvedValue([
          { id: 'inst-1', subject: 'Weekly Standup' },
        ]);

        const instances = await repository.listEventInstancesAsync(
          mintSelfEncoded('event', 'evt-recurring'),
          '2024-01-01T00:00:00Z',
          '2024-01-31T23:59:59Z'
        );

        expect(instances[0].id).toBe(mintSelfEncoded('event', 'inst-1'));
        // That token resolves on getEvent with no cache.
        mockClient.getEvent.mockResolvedValue({ id: 'inst-1', subject: 'Weekly Standup' });
        const event = await repository.getEventAsync(instances[0].id);
        expect(event).toBeDefined();
        expect(mockClient.getEvent).toHaveBeenCalledWith('inst-1');
      });

      it('rejects a legacy numeric id on Graph (NUMERIC_ID_UNSUPPORTED)', async () => {
        await expect(
          repository.listEventInstancesAsync(999999, '2024-01-01T00:00:00Z', '2024-01-31T23:59:59Z')
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
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

      it('rows carry a durable ct_ token that decodes to the Graph id (U5)', async () => {
        mockClient.searchContacts.mockResolvedValue([
          { id: 'contact-1', displayName: 'John Doe' },
        ]);

        const result = await repository.searchContactsAsync('John', 50);

        expect(result[0].id).toBe(mintSelfEncoded('contact', 'contact-1'));
      });
    });

    describe('getContact (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.getContact(123)).toThrow('Use getContactAsync()');
      });
    });

    describe('getContactAsync', () => {
      it('resolves a ct_ token to the Graph id — no prior list/cache needed (cold state)', async () => {
        mockClient.getContact.mockResolvedValue({
          id: 'contact-1',
          displayName: 'John Doe',
          surname: 'Doe',
        });

        // A fresh repository that never listed — the token alone resolves.
        const token = mintSelfEncoded('contact', 'contact-1');
        const result = await repository.getContactAsync(token);

        expect(result?.displayName).toBe('John Doe');
        expect(mockClient.getContact).toHaveBeenCalledWith('contact-1');
      });

      it('rejects a legacy numeric id on the Graph backend (NUMERIC_ID_UNSUPPORTED, D4)', async () => {
        await expect(repository.getContactAsync(99999)).rejects.toMatchObject({
          code: 'NUMERIC_ID_UNSUPPORTED',
        });
      });

      it('returns undefined when the contact is not found', async () => {
        mockClient.getContact.mockResolvedValue(null);

        const result = await repository.getContactAsync(mintSelfEncoded('contact', 'contact-1'));
        expect(result).toBeUndefined();
      });

      it('rejects a token for a different entity kind (ID_ENTITY_MISMATCH) before hitting Graph', async () => {
        const folderToken = mintSelfEncoded('folder', 'folder-1');
        await expect(repository.getContactAsync(folderToken)).rejects.toMatchObject({
          code: 'ID_ENTITY_MISMATCH',
        });
        expect(mockClient.getContact).not.toHaveBeenCalled();
      });
    });
  });

  describe('Tasks (durable tl_ / td_ tokens)', () => {
    describe('listTasks (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.listTasks(50, 0)).toThrow('Use listTasksAsync()');
      });
    });

    describe('listTasksAsync', () => {
      it('returns mapped task rows with durable td_/tl_ tokens', async () => {
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'Task 1' },
        ]);

        const result = await repository.listTasksAsync(50, 0);

        expect(result).toHaveLength(1);
        expect(result[0].name).toBe('Task 1');
        expect(result[0].id).toMatch(/^td_/);
        expect(result[0].folderId).toMatch(/^tl_/);
      });

      it('skips tasks missing an id or taskListId', async () => {
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-1', taskListId: undefined, title: 'No list' },
          { id: undefined, taskListId: 'list-1', title: 'No id' },
        ]);

        const result = await repository.listTasksAsync(50, 0);

        expect(result).toHaveLength(0);
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
      it('returns search results with a durable td_ token', async () => {
        mockClient.searchTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'Matching Task' },
        ]);

        const result = await repository.searchTasksAsync('Matching', 50);

        expect(result).toHaveLength(1);
        expect(result[0].name).toBe('Matching Task');
        expect(result[0].id).toMatch(/^td_/);
        expect(mockClient.searchTasks).toHaveBeenCalledWith('Matching', 50);
      });

      it('mints a td_ token resolvable via getTaskInfo', async () => {
        mockClient.searchTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'Task' },
        ]);

        const result = await repository.searchTasksAsync('Task', 50);

        const taskInfo = repository.getTaskInfo(result[0].id);
        expect(taskInfo).toEqual({ taskListId: 'list-1', taskId: 'task-1' });
      });
    });

    describe('getTask (sync)', () => {
      it('throws error (sync not supported)', () => {
        expect(() => repository.getTask(123)).toThrow('Use getTaskAsync()');
      });
    });

    describe('getTaskAsync', () => {
      it('returns task by td_ token', async () => {
        // Mint a td_ token via listTasksAsync
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'Task 1' },
        ]);
        const tasks = await repository.listTasksAsync(50, 0);
        const tok = tasks[0].id;

        mockClient.getTask.mockResolvedValue({
          id: 'task-1',
          title: 'Task 1',
          status: 'completed',
        });

        const result = await repository.getTaskAsync(tok);

        expect(result?.name).toBe('Task 1');
        expect(mockClient.getTask).toHaveBeenCalledWith('list-1', 'task-1');
      });

      it('returns undefined for an unknown td_ token (contract)', async () => {
        const result = await repository.getTaskAsync('td_bogus');
        expect(result).toBeUndefined();
      });

      it('returns undefined for a legacy numeric id (contract)', async () => {
        const result = await repository.getTaskAsync(99999);
        expect(result).toBeUndefined();
      });
    });
  });

  describe('listTaskListsAsync', () => {
    it('returns mapped task lists with durable tl_ tokens', async () => {
      mockClient.listTaskLists.mockResolvedValue([
        { id: 'list-1', displayName: 'My Tasks', wellknownListName: 'defaultList' },
        { id: 'list-2', displayName: 'Work', wellknownListName: 'none' },
      ]);

      const result = await repository.listTaskListsAsync();

      expect(result).toHaveLength(2);
      expect(result[0].id).toMatch(/^tl_/);
      expect(result[0].name).toBe('My Tasks');
      expect(result[0].isDefault).toBe(true);
      expect(result[1].id).toMatch(/^tl_/);
      expect(result[1].id).not.toBe(result[0].id);
      expect(result[1].name).toBe('Work');
      expect(result[1].isDefault).toBe(false);
    });

    it('mints a tl_ token resolvable via getTaskListGraphId', async () => {
      mockClient.listTaskLists.mockResolvedValue([
        { id: 'list-1', displayName: 'My Tasks', wellknownListName: 'defaultList' },
      ]);

      const result = await repository.listTaskListsAsync();

      const graphId = repository.getTaskListGraphId(result[0].id);
      expect(graphId).toBe('list-1');
    });

    it('detects default list via wellknownListName', async () => {
      mockClient.listTaskLists.mockResolvedValue([
        { id: 'list-1', displayName: 'Tasks', wellknownListName: 'defaultList' },
        { id: 'list-2', displayName: 'Custom', wellknownListName: 'none' },
        { id: 'list-3', displayName: 'Another' },
      ]);

      const result = await repository.listTaskListsAsync();

      expect(result[0].isDefault).toBe(true);
      expect(result[1].isDefault).toBe(false);
      expect(result[2].isDefault).toBe(false);
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
      it('resolves the message and folder tokens — no prior list/cache needed (cold state)', async () => {
        mockClient.moveMessage.mockResolvedValue(undefined);

        await repository.moveEmailAsync(
          mintSelfEncoded('message', 'msg-1'),
          mintSelfEncoded('folder', 'folder-dest')
        );

        expect(mockClient.moveMessage).toHaveBeenCalledWith('msg-1', 'folder-dest');
      });

      it('throws when message ID not in cache', async () => {
        await expect(repository.moveEmailAsync(99999, 'folder-dest')).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });

      it('rejects a legacy numeric destination folder id on the Graph backend (NUMERIC_ID_UNSUPPORTED, D4)', async () => {
        await expect(
          repository.moveEmailAsync(mintSelfEncoded('message', 'msg-1'), 99999 as unknown as string)
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });

    describe('deleteEmailAsync', () => {
      it('deletes message using cached ID', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-1', subject: 'Test' },
        ]);
        await repository.searchEmailsAsync('Test', 50);

        mockClient.deleteMessage.mockResolvedValue(undefined);

        await repository.deleteEmailAsync(mintSelfEncoded('message', 'msg-1'));

        expect(mockClient.deleteMessage).toHaveBeenCalledWith('msg-1');
      });

      it('throws when message ID not in cache', async () => {
        await expect(repository.deleteEmailAsync(99999)).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });

    describe('createFolderAsync', () => {
      it('creates a folder and returns a row carrying a durable fd_ token', async () => {
        mockClient.createMailFolder.mockResolvedValue({
          id: 'new-folder-id',
          displayName: 'New Folder',
          totalItemCount: 0,
          unreadItemCount: 0,
        });

        const result = await repository.createFolderAsync('New Folder');

        expect(result.name).toBe('New Folder');
        expect(mockClient.createMailFolder).toHaveBeenCalledWith('New Folder', undefined);
        // The mapper mints the fd_ token — no cache needed to resolve it (cold state).
        expect(result.id).toBe(mintSelfEncoded('folder', 'new-folder-id'));
        expect(repository.getFolderGraphId(result.id)).toBe('new-folder-id');
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

        expect(result).toEqual({ token: mintSelfEncoded('message', 'draft-1'), graphId: 'draft-1' });
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

        expect(result).toEqual({ token: mintSelfEncoded('message', 'draft-2'), graphId: 'draft-2' });
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

        await repository.updateDraftAsync(mintSelfEncoded('message', 'draft-1'), {
          subject: 'New Subject',
        });

        expect(mockClient.updateDraft).toHaveBeenCalledWith('draft-1', {
          subject: 'New Subject',
        });
      });

      it('throws when draft ID not in cache', async () => {
        await expect(
          repository.updateDraftAsync(99999, { subject: 'New' })
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
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

      it('emits durable em_ tokens for draft rows (no cache)', async () => {
        mockClient.listMessages.mockResolvedValue([
          { id: 'draft-1', subject: 'Draft 1' },
        ]);

        const result = await repository.listDraftsAsync(50, 0);

        expect(result[0].id).toBe(mintSelfEncoded('message', 'draft-1'));
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

        await repository.sendDraftAsync(mintSelfEncoded('message', 'draft-1'));

        expect(mockClient.sendDraft).toHaveBeenCalledWith('draft-1');
      });

      it('throws when draft ID not in cache', async () => {
        await expect(repository.sendDraftAsync(99999)).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
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
          mintSelfEncoded('message', 'msg-1'),
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
          mintSelfEncoded('message', 'msg-1'),
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
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
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
          mintSelfEncoded('message', 'msg-1'),
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
          mintSelfEncoded('message', 'msg-1'),
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
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
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
        await repository.archiveEmailAsync(mintSelfEncoded('message', 'msg-arch'));

        expect(mockClient.archiveMessage).toHaveBeenCalledWith('msg-arch');
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.archiveEmailAsync(99999)
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });

    describe('junkEmailAsync', () => {
      it('calls junkMessage with the correct graph ID', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-junk', subject: 'Spam' },
        ]);
        await repository.searchEmailsAsync('Spam', 50);

        mockClient.junkMessage.mockResolvedValue(undefined);
        await repository.junkEmailAsync(mintSelfEncoded('message', 'msg-junk'));

        expect(mockClient.junkMessage).toHaveBeenCalledWith('msg-junk');
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.junkEmailAsync(99999)
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });

    describe('markEmailReadAsync', () => {
      it('calls updateMessage with isRead flag', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-read', subject: 'Read me' },
        ]);
        await repository.searchEmailsAsync('Read me', 50);

        mockClient.updateMessage.mockResolvedValue(undefined);
        await repository.markEmailReadAsync(mintSelfEncoded('message', 'msg-read'), true);

        expect(mockClient.updateMessage).toHaveBeenCalledWith('msg-read', { isRead: true });
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.markEmailReadAsync(99999, false)
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });

    describe('setEmailFlagAsync', () => {
      it('maps flag status 0 to notFlagged', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-flag', subject: 'Flag me' },
        ]);
        await repository.searchEmailsAsync('Flag me', 50);

        mockClient.updateMessage.mockResolvedValue(undefined);
        await repository.setEmailFlagAsync(mintSelfEncoded('message', 'msg-flag'), 0);

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
        await repository.setEmailFlagAsync(mintSelfEncoded('message', 'msg-flag1'), 1);

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
        await repository.setEmailFlagAsync(mintSelfEncoded('message', 'msg-flag2'), 2);

        expect(mockClient.updateMessage).toHaveBeenCalledWith('msg-flag2', {
          flag: { flagStatus: 'complete' },
        });
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.setEmailFlagAsync(99999, 0)
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
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
          mintSelfEncoded('message', 'msg-cat'),
          ['Important', 'Work']
        );

        expect(mockClient.updateMessage).toHaveBeenCalledWith('msg-cat', {
          categories: ['Important', 'Work'],
        });
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.setEmailCategoriesAsync(99999, ['cat'])
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });

    describe('setEmailImportanceAsync', () => {
      it('updates message importance via updateMessage', async () => {
        mockClient.searchMessages.mockResolvedValue([{ id: 'msg-imp', subject: 'Test' }]);
        await repository.searchEmailsAsync('Test', 50);
        mockClient.updateMessage.mockResolvedValue(undefined);

        await repository.setEmailImportanceAsync(mintSelfEncoded('message', 'msg-imp'), 'high');
        expect(mockClient.updateMessage).toHaveBeenCalledWith('msg-imp', { importance: 'high' });
      });

      it('throws when email not in cache', async () => {
        await expect(repository.setEmailImportanceAsync(99999, 'high'))
          .rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });
  });

  describe('Folder Write Operations (Async)', () => {
    describe('createFolderAsync', () => {
      it('calls createMailFolder and returns a row carrying a durable fd_ token', async () => {
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
        expect(result.id).toBe(mintSelfEncoded('folder', 'folder-new'));
        expect(repository.getFolderGraphId(result.id)).toBe('folder-new');
      });

      it('passes parent folder graph ID when parentFolderId provided — cold resolve, no cache needed', async () => {
        mockClient.createMailFolder.mockResolvedValue({
          id: 'sub-folder',
          displayName: 'SubFolder',
          parentFolderId: 'parent-folder',
          totalItemCount: 0,
          unreadItemCount: 0,
        });

        await repository.createFolderAsync('SubFolder', mintSelfEncoded('folder', 'parent-folder'));

        expect(mockClient.createMailFolder).toHaveBeenCalledWith('SubFolder', 'parent-folder');
      });
    });

    describe('deleteFolderAsync', () => {
      it('calls deleteMailFolder — cold resolve, no cache needed', async () => {
        mockClient.deleteMailFolder.mockResolvedValue(undefined);

        await repository.deleteFolderAsync(mintSelfEncoded('folder', 'folder-del'));

        expect(mockClient.deleteMailFolder).toHaveBeenCalledWith('folder-del');
      });

      it('rejects a legacy numeric folder id on the Graph backend (NUMERIC_ID_UNSUPPORTED, D4)', async () => {
        await expect(
          repository.deleteFolderAsync(99999 as unknown as string)
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });

    describe('renameFolderAsync', () => {
      it('calls renameMailFolder with the correct graph ID — cold resolve, no cache needed', async () => {
        mockClient.renameMailFolder.mockResolvedValue(undefined);

        await repository.renameFolderAsync(mintSelfEncoded('folder', 'folder-ren'), 'NewName');

        expect(mockClient.renameMailFolder).toHaveBeenCalledWith('folder-ren', 'NewName');
      });

      it('rejects a legacy numeric folder id on the Graph backend (NUMERIC_ID_UNSUPPORTED, D4)', async () => {
        await expect(
          repository.renameFolderAsync(99999 as unknown as string, 'NewName')
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });

    describe('moveFolderAsync', () => {
      it('calls moveMailFolder with correct graph IDs — cold resolve, no cache needed', async () => {
        mockClient.moveMailFolder.mockResolvedValue(undefined);

        await repository.moveFolderAsync(
          mintSelfEncoded('folder', 'folder-src'),
          mintSelfEncoded('folder', 'folder-dest')
        );

        expect(mockClient.moveMailFolder).toHaveBeenCalledWith('folder-src', 'folder-dest');
      });

      it('rejects a legacy numeric source folder id on the Graph backend (NUMERIC_ID_UNSUPPORTED, D4)', async () => {
        await expect(
          repository.moveFolderAsync(99999 as unknown as string, mintSelfEncoded('folder', 'folder-dest'))
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });

      it('rejects a legacy numeric destination folder id on the Graph backend (NUMERIC_ID_UNSUPPORTED, D4)', async () => {
        await expect(
          repository.moveFolderAsync(mintSelfEncoded('folder', 'folder-only'), 88888 as unknown as string)
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });

    describe('emptyFolderAsync', () => {
      it('calls emptyMailFolder with the correct graph ID — cold resolve, no cache needed', async () => {
        mockClient.emptyMailFolder.mockResolvedValue(undefined);

        await repository.emptyFolderAsync(mintSelfEncoded('folder', 'folder-empty'));

        expect(mockClient.emptyMailFolder).toHaveBeenCalledWith('folder-empty');
      });

      it('rejects a legacy numeric folder id on the Graph backend (NUMERIC_ID_UNSUPPORTED, D4)', async () => {
        await expect(
          repository.emptyFolderAsync(99999 as unknown as string)
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
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

        const result = await repository.replyAsDraftAsync(mintSelfEncoded('message', 'msg-orig'));

        expect(mockClient.createReplyDraft).toHaveBeenCalledWith('msg-orig', undefined, undefined);
        expect(result.token).toBe(mintSelfEncoded('message', 'draft-reply-1'));
        expect(result.graphId).toBe('draft-reply-1');
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

        const result = await repository.replyAsDraftAsync(mintSelfEncoded('message', 'msg-orig2'), true);

        expect(mockClient.createReplyAllDraft).toHaveBeenCalledWith('msg-orig2', undefined, undefined);
        expect(result.graphId).toBe('draft-ra-1');
      });

      it('passes comment as body through createReply to preserve quoted thread', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-comment', subject: 'FYI' },
        ]);
        await repository.searchEmailsAsync('FYI', 50);

        mockClient.createReplyDraft.mockResolvedValue({
          id: 'draft-comment-1',
          subject: 'Re: FYI',
          toRecipients: [],
        });

        await repository.replyAsDraftAsync(mintSelfEncoded('message', 'msg-comment'), false, 'Thanks for sharing!');

        expect(mockClient.createReplyDraft).toHaveBeenCalledWith('msg-comment', undefined, {
          contentType: 'text', content: 'Thanks for sharing!',
        });
        expect(mockClient.updateDraft).not.toHaveBeenCalled();
      });

      it('uses provided bodyType when passing comment', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-html', subject: 'HTML test' },
        ]);
        await repository.searchEmailsAsync('HTML test', 50);

        mockClient.createReplyDraft.mockResolvedValue({
          id: 'draft-html-1',
          subject: 'Re: HTML test',
          toRecipients: [],
        });

        await repository.replyAsDraftAsync(mintSelfEncoded('message', 'msg-html'), false, '<p>HTML reply</p>', 'html');

        expect(mockClient.createReplyDraft).toHaveBeenCalledWith('msg-html', undefined, {
          contentType: 'html', content: '<p>HTML reply</p>',
        });
        expect(mockClient.updateDraft).not.toHaveBeenCalled();
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.replyAsDraftAsync(99999)
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
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

        const result = await repository.forwardAsDraftAsync(mintSelfEncoded('message', 'msg-fwd'));

        expect(mockClient.createForwardDraft).toHaveBeenCalledWith('msg-fwd');
        expect(result.token).toBe(mintSelfEncoded('message', 'draft-fwd-1'));
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
          mintSelfEncoded('message', 'msg-fwd2'),
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
          mintSelfEncoded('message', 'msg-fwd-html'),
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
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });
  });

  describe('Attachment Operations (Async)', () => {
    describe('listAttachmentsAsync', () => {
      it('lists attachments with durable at_ tokens', async () => {
        mockClient.listAttachments.mockResolvedValue([
          { id: 'att-1', name: 'doc.pdf', size: 1024, contentType: 'application/pdf', isInline: false },
          { id: 'att-2', name: 'image.png', size: 2048, contentType: 'image/png', isInline: true },
        ]);

        const result = await repository.listAttachmentsAsync('msg-att-1');

        expect(mockClient.listAttachments).toHaveBeenCalledWith('msg-att-1');
        expect(result).toHaveLength(2);
        expect(result[0].id).toMatch(/^at_/);
        expect(result[0]).toMatchObject({
          name: 'doc.pdf',
          size: 1024,
          contentType: 'application/pdf',
          isInline: false,
        });
        expect(result[1].id).toMatch(/^at_/);
        expect(result[1].id).not.toBe(result[0].id);
        expect(result[1]).toMatchObject({
          name: 'image.png',
          size: 2048,
          contentType: 'image/png',
          isInline: true,
        });
      });

      it('handles missing attachment fields with defaults', async () => {
        mockClient.listAttachments.mockResolvedValue([
          { id: null, name: null, size: null, contentType: null },
        ]);

        const result = await repository.listAttachmentsAsync('msg-att-2');

        expect(result).toHaveLength(1);
        expect(result[0].name).toBe('');
        expect(result[0].size).toBe(0);
        expect(result[0].contentType).toBe('application/octet-stream');
        expect(result[0].isInline).toBe(false);
      });

      it('rejects a legacy numeric email id', async () => {
        await expect(
          repository.listAttachmentsAsync(99999)
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });

    describe('downloadAttachmentAsync', () => {
      it('resolves the at_ token and delegates to downloadAttachment helper', async () => {
        mockClient.listAttachments.mockResolvedValue([
          { id: 'att-dl-1', name: 'file.zip', size: 5000, contentType: 'application/zip' },
        ]);
        const attachments = await repository.listAttachmentsAsync('msg-dl');
        const tok = attachments[0].id;

        const mockResult = { filePath: '/tmp/file.zip', name: 'file.zip', size: 5000, contentType: 'application/zip' };
        vi.mocked(downloadAttachment).mockResolvedValue(mockResult);

        const result = await repository.downloadAttachmentAsync(tok);

        expect(downloadAttachment).toHaveBeenCalledWith(
          mockClient,
          'msg-dl',
          'att-dl-1'
        );
        expect(result).toEqual(mockResult);
      });

      it('rejects an unknown at_ token', async () => {
        await expect(repository.downloadAttachmentAsync('at_bogus')).rejects.toThrow('Unknown or unresolvable');
      });

      it('rejects a legacy numeric id', async () => {
        await expect(repository.downloadAttachmentAsync(99999)).rejects.toThrow('not supported');
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

        // Returns a durable ev_ token that resolves back to the new event.
        expect(numericId).toBe(mintSelfEncoded('event', 'event-new-1'));
        mockClient.getEvent.mockResolvedValue({ id: 'event-new-1', subject: 'Team Meeting' });
        const row = await repository.getEventAsync(numericId);
        expect(row).toBeDefined();
        expect(mockClient.getEvent).toHaveBeenCalledWith('event-new-1');
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

      it('passes calendarId when provided — cold resolve, no cache needed', async () => {
        mockClient.createEvent.mockResolvedValue({ id: 'event-cal' });

        await repository.createEventAsync({
          subject: 'Work Event',
          start: '2026-03-01T10:00:00',
          end: '2026-03-01T11:00:00',
          timezone: 'UTC',
          calendarId: mintSelfEncoded('folder', 'cal-work'),
        });

        expect(mockClient.createEvent).toHaveBeenCalledWith(
          expect.any(Object),
          'cal-work'
        );
      });

      it('sets isOnlineMeeting and default provider when is_online_meeting is true', async () => {
        mockClient.createEvent.mockResolvedValue({ id: 'event-online' });

        await repository.createEventAsync({
          subject: 'Teams Meeting',
          start: '2026-03-01T10:00:00',
          end: '2026-03-01T11:00:00',
          timezone: 'UTC',
          is_online_meeting: true,
        });

        const callArgs = mockClient.createEvent.mock.calls[0][0];
        expect(callArgs.isOnlineMeeting).toBe(true);
        expect(callArgs.onlineMeetingProvider).toBe('teamsForBusiness');
      });

      it('uses specified online_meeting_provider', async () => {
        mockClient.createEvent.mockResolvedValue({ id: 'event-skype' });

        await repository.createEventAsync({
          subject: 'Skype Meeting',
          start: '2026-03-01T10:00:00',
          end: '2026-03-01T11:00:00',
          timezone: 'UTC',
          is_online_meeting: true,
          online_meeting_provider: 'skypeForBusiness',
        });

        const callArgs = mockClient.createEvent.mock.calls[0][0];
        expect(callArgs.isOnlineMeeting).toBe(true);
        expect(callArgs.onlineMeetingProvider).toBe('skypeForBusiness');
      });

      it('does not set online meeting fields when is_online_meeting is false', async () => {
        mockClient.createEvent.mockResolvedValue({ id: 'event-no-online' });

        await repository.createEventAsync({
          subject: 'Regular Meeting',
          start: '2026-03-01T10:00:00',
          end: '2026-03-01T11:00:00',
          timezone: 'UTC',
          is_online_meeting: false,
        });

        const callArgs = mockClient.createEvent.mock.calls[0][0];
        expect(callArgs.isOnlineMeeting).toBeUndefined();
        expect(callArgs.onlineMeetingProvider).toBeUndefined();
      });
    });

    describe('updateEventAsync', () => {
      it('resolves the ev_ token and calls updateEvent — no prior list needed', async () => {
        mockClient.updateEvent.mockResolvedValue(undefined);

        await repository.updateEventAsync(mintSelfEncoded('event', 'event-upd'), {
          subject: 'Updated Meeting',
        });

        expect(mockClient.updateEvent).toHaveBeenCalledWith('event-upd', {
          subject: 'Updated Meeting',
        });
      });

      it('rejects a legacy numeric id on Graph (NUMERIC_ID_UNSUPPORTED)', async () => {
        await expect(
          repository.updateEventAsync(99999, { subject: 'Nope' })
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });

      it('passes online meeting fields through to updateEvent', async () => {
        mockClient.updateEvent.mockResolvedValue(undefined);

        await repository.updateEventAsync(mintSelfEncoded('event', 'event-online-upd'), {
          isOnlineMeeting: true,
          onlineMeetingProvider: 'teamsForBusiness',
        });

        expect(mockClient.updateEvent).toHaveBeenCalledWith('event-online-upd', {
          isOnlineMeeting: true,
          onlineMeetingProvider: 'teamsForBusiness',
        });
      });
    });

    describe('deleteEventAsync', () => {
      it('resolves the ev_ token and calls deleteEvent — no prior list needed', async () => {
        mockClient.deleteEvent.mockResolvedValue(undefined);

        await repository.deleteEventAsync(mintSelfEncoded('event', 'event-del'));

        expect(mockClient.deleteEvent).toHaveBeenCalledWith('event-del');
      });

      it('rejects a legacy numeric id on Graph (NUMERIC_ID_UNSUPPORTED)', async () => {
        await expect(
          repository.deleteEventAsync(99999)
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });

    describe('respondToEventAsync', () => {
      it('responds to event with accept and comment', async () => {
        mockClient.respondToEvent.mockResolvedValue(undefined);

        await repository.respondToEventAsync(
          mintSelfEncoded('event', 'event-resp'),
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
        mockClient.respondToEvent.mockResolvedValue(undefined);

        await repository.respondToEventAsync(
          mintSelfEncoded('event', 'event-resp2'),
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

      it('rejects a legacy numeric id on Graph (NUMERIC_ID_UNSUPPORTED)', async () => {
        await expect(
          repository.respondToEventAsync(99999, 'accept', true)
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
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

        expect(numericId).toBe(mintSelfEncoded('contact', 'contact-new-1'));
      });

      it('returns a durable ct_ token that resolves back to the new contact', async () => {
        mockClient.createContact.mockResolvedValue({
          id: 'contact-new-2',
          displayName: 'Jane',
        });

        const token = await repository.createContactAsync({
          given_name: 'Jane',
        });

        expect(token).toBe(mintSelfEncoded('contact', 'contact-new-2'));
        // The token resolves with no cache — a follow-up get works cold.
        mockClient.getContact.mockResolvedValue({ id: 'contact-new-2', displayName: 'Jane' });
        const row = await repository.getContactAsync(token);
        expect(row?.displayName).toBe('Jane');
        expect(mockClient.getContact).toHaveBeenCalledWith('contact-new-2');
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
      it('resolves the ct_ token and calls client.updateContact — no prior list needed', async () => {
        mockClient.updateContact.mockResolvedValue(undefined);

        await repository.updateContactAsync(mintSelfEncoded('contact', 'contact-1'), {
          givenName: 'Updated',
        });

        expect(mockClient.updateContact).toHaveBeenCalledWith('contact-1', {
          givenName: 'Updated',
        });
      });

      it('rejects a legacy numeric id on Graph (NUMERIC_ID_UNSUPPORTED)', async () => {
        await expect(
          repository.updateContactAsync(99999, { givenName: 'Nope' })
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });

    describe('deleteContactAsync', () => {
      it('resolves the ct_ token and calls client.deleteContact — no prior list needed', async () => {
        mockClient.deleteContact.mockResolvedValue(undefined);

        await repository.deleteContactAsync(mintSelfEncoded('contact', 'contact-del'));

        expect(mockClient.deleteContact).toHaveBeenCalledWith('contact-del');
      });

      it('rejects a legacy numeric id on Graph (NUMERIC_ID_UNSUPPORTED)', async () => {
        await expect(
          repository.deleteContactAsync(99999)
        ).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });
  });

  describe('Task Write Operations (Async)', () => {
    describe('createTaskAsync', () => {
      it('creates a task with all fields and returns a resolvable td_ token', async () => {
        mockClient.listTaskLists.mockResolvedValue([
          { id: 'list-1', displayName: 'List', wellknownListName: 'none' },
        ]);
        const lists = await repository.listTaskListsAsync();
        const listTok = lists[0].id;

        mockClient.createTask.mockResolvedValue({
          id: 'task-new-1',
          title: 'New Task',
        });

        const taskTok = await repository.createTaskAsync({
          title: 'New Task',
          task_list_id: listTok,
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

        expect(taskTok).toMatch(/^td_/);

        // Verify the token resolves
        const taskInfo = repository.getTaskInfo(taskTok);
        expect(taskInfo).toEqual({ taskListId: 'list-1', taskId: 'task-new-1' });
      });

      it('creates a task with only required fields', async () => {
        mockClient.listTaskLists.mockResolvedValue([
          { id: 'list-1', displayName: 'List', wellknownListName: 'none' },
        ]);
        const lists = await repository.listTaskListsAsync();
        const listTok = lists[0].id;

        mockClient.createTask.mockResolvedValue({
          id: 'task-min',
          title: 'Minimal Task',
        });

        await repository.createTaskAsync({
          title: 'Minimal Task',
          task_list_id: listTok,
        });

        expect(mockClient.createTask).toHaveBeenCalledWith('list-1', {
          title: 'Minimal Task',
        });
      });

      it('creates a task with daily recurrence (noEnd)', async () => {
        mockClient.listTaskLists.mockResolvedValue([
          { id: 'list-1', displayName: 'List', wellknownListName: 'none' },
        ]);
        const lists = await repository.listTaskListsAsync();
        const listTok = lists[0].id;

        mockClient.createTask.mockResolvedValue({
          id: 'task-recur-1',
          title: 'Daily Task',
        });

        await repository.createTaskAsync({
          title: 'Daily Task',
          task_list_id: listTok,
          recurrence: {
            pattern: 'daily',
            interval: 1,
            range_type: 'noEnd',
            start_date: '2026-03-01',
          },
        });

        expect(mockClient.createTask).toHaveBeenCalledWith('list-1', {
          title: 'Daily Task',
          recurrence: {
            pattern: {
              type: 'daily',
              interval: 1,
            },
            range: {
              type: 'noEnd',
              startDate: '2026-03-01',
            },
          },
        });
      });

      it('creates a task with weekly recurrence and days_of_week', async () => {
        mockClient.listTaskLists.mockResolvedValue([
          { id: 'list-1', displayName: 'List', wellknownListName: 'none' },
        ]);
        const lists = await repository.listTaskListsAsync();
        const listTok = lists[0].id;

        mockClient.createTask.mockResolvedValue({
          id: 'task-recur-2',
          title: 'Weekly Task',
        });

        await repository.createTaskAsync({
          title: 'Weekly Task',
          task_list_id: listTok,
          recurrence: {
            pattern: 'weekly',
            interval: 2,
            days_of_week: ['monday', 'wednesday', 'friday'],
            range_type: 'endDate',
            start_date: '2026-03-01',
            end_date: '2026-06-01',
          },
        });

        expect(mockClient.createTask).toHaveBeenCalledWith('list-1', {
          title: 'Weekly Task',
          recurrence: {
            pattern: {
              type: 'weekly',
              interval: 2,
              daysOfWeek: ['monday', 'wednesday', 'friday'],
            },
            range: {
              type: 'endDate',
              startDate: '2026-03-01',
              endDate: '2026-06-01',
            },
          },
        });
      });

      it('creates a task with monthly recurrence and day_of_month', async () => {
        mockClient.listTaskLists.mockResolvedValue([
          { id: 'list-1', displayName: 'List', wellknownListName: 'none' },
        ]);
        const lists = await repository.listTaskListsAsync();
        const listTok = lists[0].id;

        mockClient.createTask.mockResolvedValue({
          id: 'task-recur-3',
          title: 'Monthly Task',
        });

        await repository.createTaskAsync({
          title: 'Monthly Task',
          task_list_id: listTok,
          recurrence: {
            pattern: 'monthly',
            day_of_month: 15,
            range_type: 'numbered',
            start_date: '2026-03-01',
            occurrences: 12,
          },
        });

        expect(mockClient.createTask).toHaveBeenCalledWith('list-1', {
          title: 'Monthly Task',
          recurrence: {
            pattern: {
              type: 'monthly',
              interval: 1,
              dayOfMonth: 15,
            },
            range: {
              type: 'numbered',
              startDate: '2026-03-01',
              numberOfOccurrences: 12,
            },
          },
        });
      });

      it('creates a task without recurrence — no recurrence field in graph object', async () => {
        mockClient.listTaskLists.mockResolvedValue([
          { id: 'list-1', displayName: 'List', wellknownListName: 'none' },
        ]);
        const lists = await repository.listTaskListsAsync();
        const listTok = lists[0].id;

        mockClient.createTask.mockResolvedValue({
          id: 'task-no-recur',
          title: 'No Recurrence',
        });

        await repository.createTaskAsync({
          title: 'No Recurrence',
          task_list_id: listTok,
        });

        const callArgs = mockClient.createTask.mock.calls[0][1];
        expect(callArgs).not.toHaveProperty('recurrence');
      });

      it('rejects a legacy numeric task list id', async () => {
        await expect(
          repository.createTaskAsync({
            title: 'Test',
            task_list_id: 99999,
          })
        ).rejects.toThrow('not supported');
      });

      it('rejects an unknown tl_ token', async () => {
        await expect(
          repository.createTaskAsync({
            title: 'Test',
            task_list_id: 'tl_bogus',
          })
        ).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('updateTaskAsync', () => {
      it('resolves the td_ token and calls client.updateTask', async () => {
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'Old Title' },
        ]);
        const tasks = await repository.listTasksAsync(50, 0);
        const tok = tasks[0].id;

        mockClient.updateTask.mockResolvedValue(undefined);

        await repository.updateTaskAsync(tok, {
          title: 'New Title',
        });

        expect(mockClient.updateTask).toHaveBeenCalledWith('list-1', 'task-1', {
          title: 'New Title',
        });
      });

      it('passes recurrence updates through to client.updateTask', async () => {
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'Old Title' },
        ]);
        const tasks = await repository.listTasksAsync(50, 0);
        const tok = tasks[0].id;

        mockClient.updateTask.mockResolvedValue(undefined);

        await repository.updateTaskAsync(tok, {
          recurrence: {
            pattern: {
              type: 'yearly',
              interval: 1,
            },
            range: {
              type: 'noEnd',
              startDate: '2026-01-01',
            },
          },
        });

        expect(mockClient.updateTask).toHaveBeenCalledWith('list-1', 'task-1', {
          recurrence: {
            pattern: {
              type: 'yearly',
              interval: 1,
            },
            range: {
              type: 'noEnd',
              startDate: '2026-01-01',
            },
          },
        });
      });

      it('rejects a legacy numeric task id', async () => {
        await expect(
          repository.updateTaskAsync(99999, { title: 'Nope' })
        ).rejects.toThrow('not supported');
      });

      it('rejects an unknown td_ token', async () => {
        await expect(
          repository.updateTaskAsync('td_bogus', { title: 'Nope' })
        ).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('completeTaskAsync', () => {
      it('calls updateTaskAsync with completed status', async () => {
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'To Complete' },
        ]);
        const tasks = await repository.listTasksAsync(50, 0);
        const tok = tasks[0].id;

        mockClient.updateTask.mockResolvedValue(undefined);

        await repository.completeTaskAsync(tok);

        expect(mockClient.updateTask).toHaveBeenCalledWith('list-1', 'task-1', {
          status: 'completed',
          completedDateTime: {
            dateTime: expect.any(String),
            timeZone: 'UTC',
          },
        });
      });

      it('rejects an unknown td_ token', async () => {
        await expect(
          repository.completeTaskAsync('td_bogus')
        ).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('deleteTaskAsync', () => {
      it('resolves the td_ token and calls client.deleteTask', async () => {
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-del', taskListId: 'list-1', title: 'To Delete' },
        ]);
        const tasks = await repository.listTasksAsync(50, 0);
        const tok = tasks[0].id;

        mockClient.deleteTask.mockResolvedValue(undefined);

        await repository.deleteTaskAsync(tok);

        expect(mockClient.deleteTask).toHaveBeenCalledWith('list-1', 'task-del');
      });

      it('rejects an unknown td_ token', async () => {
        await expect(
          repository.deleteTaskAsync('td_bogus')
        ).rejects.toThrow('Unknown or unresolvable');
      });

      it('rejects a legacy numeric task id', async () => {
        await expect(
          repository.deleteTaskAsync(99999)
        ).rejects.toThrow('not supported');
      });
    });

    describe('createTaskListAsync', () => {
      it('creates task list and returns a resolvable tl_ token', async () => {
        mockClient.createTaskList.mockResolvedValue({
          id: 'new-list-1',
          displayName: 'My New List',
        });

        const listTok = await repository.createTaskListAsync('My New List');

        expect(mockClient.createTaskList).toHaveBeenCalledWith('My New List');
        expect(listTok).toMatch(/^tl_/);

        const graphId = repository.getTaskListGraphId(listTok);
        expect(graphId).toBe('new-list-1');
      });
    });

    describe('renameTaskListAsync', () => {
      it('calls updateTaskList with correct args', async () => {
        mockClient.listTaskLists.mockResolvedValue([
          { id: 'list-abc', displayName: 'Old Name', isOwner: true, isShared: false, wellknownListName: 'none' },
        ]);
        const lists = await repository.listTaskListsAsync();
        const tok = lists[0].id;

        mockClient.updateTaskList.mockResolvedValue(undefined);
        await repository.renameTaskListAsync(tok, 'New Name');

        expect(mockClient.updateTaskList).toHaveBeenCalledWith('list-abc', { displayName: 'New Name' });
      });

      it('rejects an unknown tl_ token', async () => {
        await expect(repository.renameTaskListAsync('tl_bogus', 'Name')).rejects.toThrow('Unknown or unresolvable');
      });

      it('rejects a legacy numeric id', async () => {
        await expect(repository.renameTaskListAsync(999999, 'Name')).rejects.toThrow('not supported');
      });
    });

    describe('deleteTaskListAsync', () => {
      it('resolves the tl_ token and calls client.deleteTaskList', async () => {
        mockClient.listTaskLists.mockResolvedValue([
          { id: 'list-del', displayName: 'To Delete', isOwner: true, isShared: false, wellknownListName: 'none' },
        ]);
        const lists = await repository.listTaskListsAsync();
        const tok = lists[0].id;

        mockClient.deleteTaskList.mockResolvedValue(undefined);
        await repository.deleteTaskListAsync(tok);

        expect(mockClient.deleteTaskList).toHaveBeenCalledWith('list-del');
      });

      it('rejects an unknown tl_ token', async () => {
        await expect(repository.deleteTaskListAsync('tl_bogus')).rejects.toThrow('Unknown or unresolvable');
      });

      it('rejects a legacy numeric id', async () => {
        await expect(repository.deleteTaskListAsync(999999)).rejects.toThrow('not supported');
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

    describe('getFolderGraphId', () => {
      // getGraphId was removed entirely in U5b-2 — 'folder' (the last remaining
      // case) now resolves via the durable self-encoding fd_ token through
      // toGraphId, same as 'message', 'event', and 'contact' before it.
      it('resolves a folder fd_ token to its Graph id — cold state, no cache needed', () => {
        expect(repository.getFolderGraphId(mintSelfEncoded('folder', 'folder-1'))).toBe('folder-1');
      });

      it('rejects a legacy numeric folder id (NUMERIC_ID_UNSUPPORTED, D4)', () => {
        expect(() => repository.getFolderGraphId(99999 as unknown as string)).toThrowError(
          expect.objectContaining({ code: 'NUMERIC_ID_UNSUPPORTED' })
        );
      });
    });

    describe('getTaskInfo', () => {
      it('returns undefined for a legacy numeric id (contract)', () => {
        expect(repository.getTaskInfo(99999)).toBeUndefined();
      });

      it('returns undefined for an unknown td_ token (contract)', () => {
        expect(repository.getTaskInfo('td_bogus')).toBeUndefined();
      });

      it('returns task info for a resolvable td_ token', async () => {
        mockClient.listAllTasks.mockResolvedValue([
          { id: 'task-1', taskListId: 'list-1', title: 'Task 1' },
        ]);
        const tasks = await repository.listTasksAsync(50, 0);

        const info = repository.getTaskInfo(tasks[0].id);

        expect(info).toEqual({ taskListId: 'list-1', taskId: 'task-1' });
      });
    });
  });

  // ===========================================================================
  // Automatic Replies (Out of Office)
  // ===========================================================================

  describe('Automatic Replies', () => {
    describe('getAutomaticRepliesAsync', () => {
      it('returns mapped automatic replies settings', async () => {
        mockClient.getAutomaticReplies.mockResolvedValue({
          status: 'alwaysEnabled',
          externalAudience: 'all',
          internalReplyMessage: '<p>I am out</p>',
          externalReplyMessage: '<p>Away</p>',
          scheduledStartDateTime: { dateTime: '2026-03-01T00:00:00Z', timeZone: 'UTC' },
          scheduledEndDateTime: { dateTime: '2026-03-15T00:00:00Z', timeZone: 'UTC' },
        });

        const result = await repository.getAutomaticRepliesAsync();

        expect(result.status).toBe('alwaysEnabled');
        expect(result.externalAudience).toBe('all');
        expect(result.internalReplyMessage).toBe('<p>I am out</p>');
        expect(result.externalReplyMessage).toBe('<p>Away</p>');
        expect(result.scheduledStartDateTime).toBe('2026-03-01T00:00:00Z');
        expect(result.scheduledEndDateTime).toBe('2026-03-15T00:00:00Z');
      });

      it('returns defaults for missing fields', async () => {
        mockClient.getAutomaticReplies.mockResolvedValue({});

        const result = await repository.getAutomaticRepliesAsync();

        expect(result.status).toBe('disabled');
        expect(result.externalAudience).toBe('none');
        expect(result.internalReplyMessage).toBe('');
        expect(result.externalReplyMessage).toBe('');
        expect(result.scheduledStartDateTime).toBeNull();
        expect(result.scheduledEndDateTime).toBeNull();
      });
    });

    describe('setAutomaticRepliesAsync', () => {
      it('builds settings object with all fields', async () => {
        mockClient.setAutomaticReplies.mockResolvedValue(undefined);

        await repository.setAutomaticRepliesAsync({
          status: 'alwaysEnabled',
          externalAudience: 'all',
          internalReplyMessage: '<p>I am out</p>',
          externalReplyMessage: '<p>Away</p>',
        });

        expect(mockClient.setAutomaticReplies).toHaveBeenCalledWith({
          status: 'alwaysEnabled',
          externalAudience: 'all',
          internalReplyMessage: '<p>I am out</p>',
          externalReplyMessage: '<p>Away</p>',
        });
      });

      it('builds settings object with only status', async () => {
        mockClient.setAutomaticReplies.mockResolvedValue(undefined);

        await repository.setAutomaticRepliesAsync({ status: 'disabled' });

        expect(mockClient.setAutomaticReplies).toHaveBeenCalledWith({ status: 'disabled' });
      });

      it('handles scheduled dates', async () => {
        mockClient.setAutomaticReplies.mockResolvedValue(undefined);

        await repository.setAutomaticRepliesAsync({
          status: 'scheduled',
          scheduledStartDateTime: '2026-03-01T00:00:00Z',
          scheduledEndDateTime: '2026-03-15T00:00:00Z',
        });

        expect(mockClient.setAutomaticReplies).toHaveBeenCalledWith({
          status: 'scheduled',
          scheduledStartDateTime: { dateTime: '2026-03-01T00:00:00Z', timeZone: 'UTC' },
          scheduledEndDateTime: { dateTime: '2026-03-15T00:00:00Z', timeZone: 'UTC' },
        });
      });
    });
  });

  // ===========================================================================
  // Mail Rules
  // ===========================================================================

  describe('Mail Rules', () => {
    describe('listMailRulesAsync', () => {
      it('returns mapped rules with durable mr_ tokens', async () => {
        mockClient.listMailRules.mockResolvedValue([
          { id: 'rule-1', displayName: 'Rule 1', sequence: 1, isEnabled: true, conditions: { subjectContains: ['test'] }, actions: { markAsRead: true } },
          { id: 'rule-2', displayName: 'Rule 2', sequence: 2, isEnabled: false, conditions: {}, actions: {} },
        ]);

        const result = await repository.listMailRulesAsync();

        expect(result).toHaveLength(2);
        expect(result[0].id).toMatch(/^mr_/);
        expect(result[0].displayName).toBe('Rule 1');
        expect(result[0].sequence).toBe(1);
        expect(result[0].isEnabled).toBe(true);
        expect(result[0].conditions).toEqual({ subjectContains: ['test'] });
        expect(result[0].actions).toEqual({ markAsRead: true });
        expect(result[1].id).toMatch(/^mr_/);
        expect(result[1].id).not.toBe(result[0].id);
        expect(result[1].isEnabled).toBe(false);
      });

      it('mints an mr_ token resolvable for deletion', async () => {
        mockClient.listMailRules.mockResolvedValue([
          { id: 'rule-abc', displayName: 'Test' },
        ]);

        const result = await repository.listMailRulesAsync();
        const tok = result[0].id;

        mockClient.deleteMailRule.mockResolvedValue(undefined);
        await expect(repository.deleteMailRuleAsync(tok)).resolves.toBeUndefined();
        expect(mockClient.deleteMailRule).toHaveBeenCalledWith('rule-abc');
      });
    });

    describe('createMailRuleAsync', () => {
      it('creates a rule and returns a resolvable mr_ token', async () => {
        mockClient.createMailRule.mockResolvedValue({ id: 'rule-new', displayName: 'New Rule' });

        const result = await repository.createMailRuleAsync({ displayName: 'New Rule', isEnabled: true });

        expect(result).toMatch(/^mr_/);
        expect(mockClient.createMailRule).toHaveBeenCalledWith({ displayName: 'New Rule', isEnabled: true });

        mockClient.deleteMailRule.mockResolvedValue(undefined);
        await repository.deleteMailRuleAsync(result);
        expect(mockClient.deleteMailRule).toHaveBeenCalledWith('rule-new');
      });
    });

    describe('deleteMailRuleAsync', () => {
      it('resolves the mr_ token and calls client.deleteMailRule', async () => {
        mockClient.listMailRules.mockResolvedValue([{ id: 'rule-del', displayName: 'To Delete' }]);
        const rules = await repository.listMailRulesAsync();
        const tok = rules[0].id;

        mockClient.deleteMailRule.mockResolvedValue(undefined);
        await repository.deleteMailRuleAsync(tok);

        expect(mockClient.deleteMailRule).toHaveBeenCalledWith('rule-del');
      });

      it('rejects an unknown mr_ token', async () => {
        await expect(repository.deleteMailRuleAsync('mr_bogus')).rejects.toThrow('Unknown or unresolvable');
      });

      it('rejects a legacy numeric id', async () => {
        await expect(repository.deleteMailRuleAsync(999999)).rejects.toThrow('not supported');
      });
    });
  });

  // ===========================================================================
  // Master Categories
  // ===========================================================================

  describe('Master Categories', () => {
    describe('listCategoriesAsync', () => {
      it('returns mapped categories with durable cg_ tokens', async () => {
        mockClient.listMasterCategories.mockResolvedValue([
          { id: 'cat-1', displayName: 'Red Category', color: 'preset0' },
          { id: 'cat-2', displayName: 'Blue Category', color: 'preset1' },
        ]);

        const result = await repository.listCategoriesAsync();

        expect(result).toHaveLength(2);
        expect(result[0].id).toMatch(/^cg_/);
        expect(result[0].name).toBe('Red Category');
        expect(result[0].color).toBe('preset0');
        expect(result[1].id).toMatch(/^cg_/);
        expect(result[1].id).not.toBe(result[0].id);
        expect(result[1].name).toBe('Blue Category');
        expect(result[1].color).toBe('preset1');
      });

      it('mints a cg_ token resolvable for deletion', async () => {
        mockClient.listMasterCategories.mockResolvedValue([
          { id: 'cat-abc', displayName: 'Test' },
        ]);

        const result = await repository.listCategoriesAsync();
        const tok = result[0].id;

        mockClient.deleteMasterCategory.mockResolvedValue(undefined);
        await expect(repository.deleteCategoryAsync(tok)).resolves.toBeUndefined();
        expect(mockClient.deleteMasterCategory).toHaveBeenCalledWith('cat-abc');
      });
    });

    describe('createCategoryAsync', () => {
      it('creates a category and returns a resolvable cg_ token', async () => {
        mockClient.createMasterCategory.mockResolvedValue({ id: 'cat-new', displayName: 'Work', color: 'preset1' });

        const result = await repository.createCategoryAsync('Work', 'preset1');

        expect(result).toMatch(/^cg_/);
        expect(mockClient.createMasterCategory).toHaveBeenCalledWith('Work', 'preset1');

        mockClient.deleteMasterCategory.mockResolvedValue(undefined);
        await repository.deleteCategoryAsync(result);
        expect(mockClient.deleteMasterCategory).toHaveBeenCalledWith('cat-new');
      });
    });

    describe('deleteCategoryAsync', () => {
      it('resolves the cg_ token and calls client.deleteMasterCategory', async () => {
        mockClient.listMasterCategories.mockResolvedValue([{ id: 'cat-del', displayName: 'To Delete' }]);
        const categories = await repository.listCategoriesAsync();
        const tok = categories[0].id;

        mockClient.deleteMasterCategory.mockResolvedValue(undefined);
        await repository.deleteCategoryAsync(tok);

        expect(mockClient.deleteMasterCategory).toHaveBeenCalledWith('cat-del');
      });

      it('rejects an unknown cg_ token', async () => {
        await expect(repository.deleteCategoryAsync('cg_bogus')).rejects.toThrow('Unknown or unresolvable');
      });

      it('rejects a legacy numeric id', async () => {
        await expect(repository.deleteCategoryAsync(999999)).rejects.toThrow('not supported');
      });
    });
  });

  // ===========================================================================
  // Focused Inbox Overrides
  // ===========================================================================

  describe('Focused Inbox Overrides', () => {
    describe('listFocusedOverridesAsync', () => {
      it('returns mapped overrides with durable fo_ tokens', async () => {
        mockClient.listFocusedOverrides.mockResolvedValue([
          { id: 'ov-1', classifyAs: 'focused', senderEmailAddress: { address: 'a@b.com' } },
          { id: 'ov-2', classifyAs: 'other', senderEmailAddress: { address: 'c@d.com' } },
        ]);

        const result = await repository.listFocusedOverridesAsync();

        expect(result).toHaveLength(2);
        expect(result[0].id).toMatch(/^fo_/);
        expect(result[0].senderAddress).toBe('a@b.com');
        expect(result[0].classifyAs).toBe('focused');
        expect(result[1].id).toMatch(/^fo_/);
        expect(result[1].id).not.toBe(result[0].id);
        expect(result[1].senderAddress).toBe('c@d.com');
        expect(result[1].classifyAs).toBe('other');
      });

      it('mints an fo_ token resolvable for deletion', async () => {
        mockClient.listFocusedOverrides.mockResolvedValue([
          { id: 'ov-abc', classifyAs: 'focused', senderEmailAddress: { address: 'x@y.com' } },
        ]);

        const result = await repository.listFocusedOverridesAsync();
        const tok = result[0].id;

        mockClient.deleteFocusedOverride.mockResolvedValue(undefined);
        await expect(repository.deleteFocusedOverrideAsync(tok)).resolves.toBeUndefined();
        expect(mockClient.deleteFocusedOverride).toHaveBeenCalledWith('ov-abc');
      });
    });

    describe('createFocusedOverrideAsync', () => {
      it('creates an override and returns a resolvable fo_ token', async () => {
        mockClient.createFocusedOverride.mockResolvedValue({
          id: 'ov-new',
          classifyAs: 'focused',
          senderEmailAddress: { address: 'a@b.com' },
        });

        const result = await repository.createFocusedOverrideAsync('a@b.com', 'focused');

        expect(result).toMatch(/^fo_/);
        expect(mockClient.createFocusedOverride).toHaveBeenCalledWith('a@b.com', 'focused');

        mockClient.deleteFocusedOverride.mockResolvedValue(undefined);
        await repository.deleteFocusedOverrideAsync(result);
        expect(mockClient.deleteFocusedOverride).toHaveBeenCalledWith('ov-new');
      });
    });

    describe('deleteFocusedOverrideAsync', () => {
      it('resolves the fo_ token and calls client.deleteFocusedOverride', async () => {
        mockClient.listFocusedOverrides.mockResolvedValue([
          { id: 'ov-del', classifyAs: 'other', senderEmailAddress: { address: 'x@y.com' } },
        ]);
        const overrides = await repository.listFocusedOverridesAsync();
        const tok = overrides[0].id;

        mockClient.deleteFocusedOverride.mockResolvedValue(undefined);
        await repository.deleteFocusedOverrideAsync(tok);

        expect(mockClient.deleteFocusedOverride).toHaveBeenCalledWith('ov-del');
      });

      it('rejects an unknown fo_ token', async () => {
        await expect(repository.deleteFocusedOverrideAsync('fo_bogus')).rejects.toThrow('Unknown or unresolvable');
      });

      it('rejects a legacy numeric id', async () => {
        await expect(repository.deleteFocusedOverrideAsync(999999)).rejects.toThrow('not supported');
      });
    });
  });

  // ===========================================================================
  // Contact Folders
  // ===========================================================================

  describe('Contact Folders', () => {
    describe('listContactFoldersAsync', () => {
      it('returns mapped folders with durable cf_ tokens', async () => {
        mockClient.listContactFolders.mockResolvedValue([
          { id: 'cf-1', displayName: 'Work', parentFolderId: 'root-1' },
          { id: 'cf-2', displayName: 'Personal', parentFolderId: null },
        ]);

        const result = await repository.listContactFoldersAsync();

        expect(result).toHaveLength(2);
        expect(result[0].id).toMatch(/^cf_/);
        expect(result[0].name).toBe('Work');
        expect(result[0].parentFolderId).toBe('root-1');
        expect(result[1].id).toMatch(/^cf_/);
        expect(result[1].id).not.toBe(result[0].id);
        expect(result[1].name).toBe('Personal');
        expect(result[1].parentFolderId).toBeNull();
      });
    });

    describe('createContactFolderAsync', () => {
      it('creates a folder and returns a resolvable cf_ token', async () => {
        mockClient.createContactFolder.mockResolvedValue({ id: 'cf-new', displayName: 'Friends' });

        const result = await repository.createContactFolderAsync('Friends');

        expect(result).toMatch(/^cf_/);
        expect(mockClient.createContactFolder).toHaveBeenCalledWith('Friends');

        mockClient.deleteContactFolder.mockResolvedValue(undefined);
        await repository.deleteContactFolderAsync(result);
        expect(mockClient.deleteContactFolder).toHaveBeenCalledWith('cf-new');
      });
    });

    describe('deleteContactFolderAsync', () => {
      it('resolves the cf_ token and calls client.deleteContactFolder', async () => {
        mockClient.listContactFolders.mockResolvedValue([{ id: 'cf-del', displayName: 'To Delete' }]);
        const folders = await repository.listContactFoldersAsync();
        const tok = folders[0].id;

        mockClient.deleteContactFolder.mockResolvedValue(undefined);
        await repository.deleteContactFolderAsync(tok);

        expect(mockClient.deleteContactFolder).toHaveBeenCalledWith('cf-del');
      });

      it('rejects an unknown cf_ token', async () => {
        await expect(repository.deleteContactFolderAsync('cf_bogus')).rejects.toThrow('Unknown or unresolvable');
      });

      it('rejects a legacy numeric id', async () => {
        await expect(repository.deleteContactFolderAsync(999999)).rejects.toThrow('not supported');
      });
    });

    describe('listContactsInFolderAsync', () => {
      it('resolves the cf_ token and lists contacts in the folder', async () => {
        mockClient.listContactFolders.mockResolvedValue([{ id: 'cf-1', displayName: 'Work' }]);
        const folders = await repository.listContactFoldersAsync();
        const tok = folders[0].id;

        mockClient.listContactsInFolder.mockResolvedValue([
          { id: 'c-1', displayName: 'Alice', givenName: 'Alice', surname: 'Smith', emailAddresses: [], businessPhones: [] },
          { id: 'c-2', displayName: 'Bob', givenName: 'Bob', surname: 'Jones', emailAddresses: [], businessPhones: [] },
        ]);

        const result = await repository.listContactsInFolderAsync(tok, 50);

        expect(result).toHaveLength(2);
        expect(result[0].displayName).toBe('Alice');
        expect(result[1].displayName).toBe('Bob');
        expect(mockClient.listContactsInFolder).toHaveBeenCalledWith('cf-1', 50);
      });

      it('rejects an unknown cf_ token', async () => {
        await expect(repository.listContactsInFolderAsync('cf_bogus')).rejects.toThrow('Unknown or unresolvable');
      });

      it('rejects a legacy numeric id', async () => {
        await expect(repository.listContactsInFolderAsync(999999)).rejects.toThrow('not supported');
      });
    });
  });

  // ===========================================================================
  // Contact Photos
  // ===========================================================================

  describe('Contact Photos', () => {
    describe('getContactPhotoAsync', () => {
      it('downloads and saves photo, returns path', async () => {
        const mockPhotoData = new ArrayBuffer(8);
        mockClient.getContactPhoto.mockResolvedValue(mockPhotoData);

        // Cold: the ct_ token resolves without a prior list.
        const result = await repository.getContactPhotoAsync(mintSelfEncoded('contact', 'contact-1'));

        expect(mockClient.getContactPhoto).toHaveBeenCalledWith('contact-1');
        expect(fs.writeFileSync).toHaveBeenCalledWith(
          expect.stringMatching(/contact-[0-9a-f]{16}-photo\.jpg$/),
          expect.any(Buffer),
        );
        expect(result.filePath).toMatch(/contact-[0-9a-f]{16}-photo\.jpg$/);
        expect(result.contentType).toBe('image/jpeg');
      });

      it('rejects a legacy numeric id on Graph (NUMERIC_ID_UNSUPPORTED)', async () => {
        await expect(repository.getContactPhotoAsync(999999)).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });

    describe('setContactPhotoAsync', () => {
      it('reads file and uploads', async () => {
        mockClient.setContactPhoto.mockResolvedValue(undefined);

        await repository.setContactPhotoAsync(mintSelfEncoded('contact', 'contact-2'), '/tmp/photo.jpg');

        expect(fs.readFileSync).toHaveBeenCalledWith('/tmp/photo.jpg');
        expect(mockClient.setContactPhoto).toHaveBeenCalledWith(
          'contact-2',
          expect.any(Buffer),
          'image/jpeg',
        );
      });

      it('detects PNG content type', async () => {
        mockClient.setContactPhoto.mockResolvedValue(undefined);

        await repository.setContactPhotoAsync(mintSelfEncoded('contact', 'contact-3'), '/tmp/photo.png');

        expect(mockClient.setContactPhoto).toHaveBeenCalledWith(
          'contact-3',
          expect.any(Buffer),
          'image/png',
        );
      });

      it('rejects a legacy numeric id on Graph (NUMERIC_ID_UNSUPPORTED)', async () => {
        await expect(repository.setContactPhotoAsync(999999, '/tmp/photo.jpg')).rejects.toMatchObject({ code: 'NUMERIC_ID_UNSUPPORTED' });
      });
    });
  });

  // ===========================================================================
  // Message Headers & MIME
  // ===========================================================================

  describe('Message Headers & MIME', () => {
    const graphMsgId = 'AAMkAGTest123';
    const token = mintSelfEncoded('message', graphMsgId);

    describe('getMessageHeadersAsync', () => {
      it('returns internet message headers — resolves the em_ token cold, no cache needed', async () => {
        const mockHeaders = [
          { name: 'Received', value: 'from mx.example.com' },
          { name: 'Authentication-Results', value: 'spf=pass' },
        ];
        mockClient.getMessageHeaders.mockResolvedValue(mockHeaders);

        const result = await repository.getMessageHeadersAsync(token);

        expect(result).toEqual(mockHeaders);
        expect(mockClient.getMessageHeaders).toHaveBeenCalledWith(graphMsgId);
      });

      it('rejects a legacy numeric id on Graph (NUMERIC_ID_UNSUPPORTED)', async () => {
        await expect(repository.getMessageHeadersAsync(999999)).rejects.toMatchObject({
          code: 'NUMERIC_ID_UNSUPPORTED',
        });
      });
    });

    describe('getMessageMimeAsync', () => {
      it('downloads MIME content and returns file path — resolves the em_ token cold, no cache needed', async () => {
        const mimeContent = 'MIME-Version: 1.0\r\nContent-Type: text/plain\r\n\r\nHello';
        mockClient.getMessageMime.mockResolvedValue(mimeContent);

        const result = await repository.getMessageMimeAsync(token);

        expect(result.filePath).toMatch(/^\/tmp\/mcp-outlook-attachments\/email-[0-9a-f]{16}\.eml$/);
        expect(mockClient.getMessageMime).toHaveBeenCalledWith(graphMsgId);
        expect(fs.writeFileSync).toHaveBeenCalledWith(
          result.filePath,
          mimeContent,
          'utf-8'
        );
      });

      it('rejects a legacy numeric id on Graph (NUMERIC_ID_UNSUPPORTED)', async () => {
        await expect(repository.getMessageMimeAsync(999999)).rejects.toMatchObject({
          code: 'NUMERIC_ID_UNSUPPORTED',
        });
      });
    });
  });

  // ===========================================================================
  // Mail Tips
  // ===========================================================================

  describe('Mail Tips', () => {
    describe('getMailTipsAsync', () => {
      it('returns mapped mail tips for email addresses', async () => {
        mockClient.getMailTips.mockResolvedValue([
          {
            emailAddress: { address: 'alice@example.com' },
            automaticReplies: { message: 'I am on vacation' },
            mailboxFull: false,
            deliveryRestricted: false,
            externalMemberCount: 0,
            maxMessageSize: 37748736,
          },
          {
            emailAddress: { address: 'bob@example.com' },
            automaticReplies: null,
            mailboxFull: true,
            deliveryRestricted: true,
            externalMemberCount: 5,
            maxMessageSize: 10485760,
          },
        ]);

        const result = await repository.getMailTipsAsync(['alice@example.com', 'bob@example.com']);

        expect(result).toHaveLength(2);
        expect(result[0].emailAddress).toBe('alice@example.com');
        expect(result[0].automaticReplies).toEqual({ message: 'I am on vacation' });
        expect(result[0].mailboxFull).toBe(false);
        expect(result[0].deliveryRestricted).toBe(false);
        expect(result[0].externalMemberCount).toBe(0);
        expect(result[0].maxMessageSize).toBe(37748736);

        expect(result[1].emailAddress).toBe('bob@example.com');
        expect(result[1].automaticReplies).toBeNull();
        expect(result[1].mailboxFull).toBe(true);
        expect(result[1].deliveryRestricted).toBe(true);
        expect(result[1].externalMemberCount).toBe(5);
        expect(result[1].maxMessageSize).toBe(10485760);

        expect(mockClient.getMailTips).toHaveBeenCalledWith(['alice@example.com', 'bob@example.com']);
      });

      it('returns defaults for missing fields', async () => {
        mockClient.getMailTips.mockResolvedValue([
          {},
        ]);

        const result = await repository.getMailTipsAsync(['unknown@example.com']);

        expect(result).toHaveLength(1);
        expect(result[0].emailAddress).toBe('');
        expect(result[0].automaticReplies).toBeNull();
        expect(result[0].mailboxFull).toBe(false);
        expect(result[0].deliveryRestricted).toBe(false);
        expect(result[0].externalMemberCount).toBe(0);
        expect(result[0].maxMessageSize).toBe(0);
      });

      it('handles automaticReplies with empty message', async () => {
        mockClient.getMailTips.mockResolvedValue([
          {
            emailAddress: { address: 'test@example.com' },
            automaticReplies: { message: '' },
          },
        ]);

        const result = await repository.getMailTipsAsync(['test@example.com']);

        expect(result[0].automaticReplies).toBeNull();
      });
    });
  });

  describe('Calendar Groups', () => {
    // Orphan entity (U5b): no Graph URL takes a calendar-group id as a path
    // segment, so listCalendarGroupsAsync/createCalendarGroupAsync return the
    // raw Graph id string rather than minting a token — never resolved back.
    describe('listCalendarGroupsAsync', () => {
      it('returns the raw Graph id (no token minted)', async () => {
        mockClient.listCalendarGroups.mockResolvedValue([
          { id: 'cg-1', name: 'My Calendars', classId: '0006' },
          { id: 'cg-2', name: 'Other Calendars', classId: '0006' },
        ]);

        const result = await repository.listCalendarGroupsAsync();

        expect(result).toHaveLength(2);
        expect(result[0].id).toBe('cg-1');
        expect(result[0].name).toBe('My Calendars');
        expect(result[0].classId).toBe('0006');
        expect(result[1].id).toBe('cg-2');
        expect(result[1].name).toBe('Other Calendars');
      });

      it('handles empty results', async () => {
        mockClient.listCalendarGroups.mockResolvedValue([]);

        const result = await repository.listCalendarGroupsAsync();

        expect(result).toHaveLength(0);
      });

      it('handles missing name and classId', async () => {
        mockClient.listCalendarGroups.mockResolvedValue([
          { id: 'cg-1' },
        ]);

        const result = await repository.listCalendarGroupsAsync();

        expect(result[0].name).toBe('');
        expect(result[0].classId).toBe('');
      });
    });

    describe('createCalendarGroupAsync', () => {
      it('creates a calendar group and returns the raw Graph id (no token minted)', async () => {
        mockClient.createCalendarGroup.mockResolvedValue({
          id: 'cg-new',
          name: 'Work',
          classId: '0006',
        });

        const result = await repository.createCalendarGroupAsync('Work');

        expect(result).toBe('cg-new');
        expect(mockClient.createCalendarGroup).toHaveBeenCalledWith('Work');
      });
    });
  });

  // ===========================================================================
  // Calendar Permissions
  // ===========================================================================

  describe('Calendar Permissions', () => {
    describe('listCalendarPermissionsAsync', () => {
      it('mints durable cp_ tokens from the calendar fd_ token', async () => {
        mockClient.listCalendarPermissions.mockResolvedValue([
          { id: 'perm-1', emailAddress: { address: 'alice@example.com' }, role: 'read', isRemovable: true, isInsideOrganization: true },
          { id: 'perm-2', emailAddress: { address: 'bob@example.com' }, role: 'write', isRemovable: true, isInsideOrganization: false },
        ]);

        const result = await repository.listCalendarPermissionsAsync(mintSelfEncoded('folder', 'cal-1'));

        expect(result).toHaveLength(2);
        expect(result[0].id).toMatch(/^cp_/);
        expect(result[0].id).not.toBe(result[1].id);
        expect(result[0].emailAddress).toBe('alice@example.com');
        expect(result[0].role).toBe('read');
        expect(result[1].emailAddress).toBe('bob@example.com');
        expect(mockClient.listCalendarPermissions).toHaveBeenCalledWith('cal-1');
      });

      it('defaults role, isRemovable, isInsideOrganization when missing', async () => {
        mockClient.listCalendarPermissions.mockResolvedValue([
          { id: 'perm-1', emailAddress: { address: 'x@example.com' } },
        ]);

        const result = await repository.listCalendarPermissionsAsync(mintSelfEncoded('folder', 'cal-1'));

        expect(result[0].role).toBe('none');
        expect(result[0].isRemovable).toBe(false);
        expect(result[0].isInsideOrganization).toBe(false);
      });
    });

    describe('createCalendarPermissionAsync', () => {
      it('creates a permission and returns a resolvable cp_ token', async () => {
        const calTok = mintSelfEncoded('folder', 'cal-1');
        mockClient.createCalendarPermission.mockResolvedValue({ id: 'perm-new' });

        const permTok = await repository.createCalendarPermissionAsync(calTok, 'alice@example.com', 'read');

        expect(permTok).toMatch(/^cp_/);
        expect(mockClient.createCalendarPermission).toHaveBeenCalledWith('cal-1', {
          emailAddress: { address: 'alice@example.com', name: 'alice@example.com' },
          role: 'read',
        });

        mockClient.deleteCalendarPermission.mockResolvedValue(undefined);
        await repository.deleteCalendarPermissionAsync(permTok);
        expect(mockClient.deleteCalendarPermission).toHaveBeenCalledWith('cal-1', 'perm-new');
      });
    });

    describe('deleteCalendarPermissionAsync', () => {
      it('resolves the cp_ token to (calendarId, permissionId) and deletes', async () => {
        mockClient.listCalendarPermissions.mockResolvedValue([
          { id: 'perm-1', emailAddress: { address: 'alice@example.com' }, role: 'read' },
        ]);
        const permissions = await repository.listCalendarPermissionsAsync(mintSelfEncoded('folder', 'cal-1'));
        const permTok = permissions[0].id;

        mockClient.deleteCalendarPermission.mockResolvedValue(undefined);
        await repository.deleteCalendarPermissionAsync(permTok);

        expect(mockClient.deleteCalendarPermission).toHaveBeenCalledWith('cal-1', 'perm-1');
      });

      it('rejects a legacy numeric permission id', async () => {
        await expect(repository.deleteCalendarPermissionAsync(999999)).rejects.toThrow('not supported');
      });

      it('rejects an unknown cp_ token', async () => {
        await expect(repository.deleteCalendarPermissionAsync('cp_bogus')).rejects.toThrow('Unknown or unresolvable');
      });
    });
  });

  // ===========================================================================
  // Online Meetings
  // ===========================================================================

  describe('Online Meetings', () => {
    describe('listOnlineMeetingsAsync', () => {
      it('mints durable om_ tokens', async () => {
        mockClient.listOnlineMeetings.mockResolvedValue([
          { id: 'meet-1', subject: 'Sprint Planning', startDateTime: '2026-03-01T10:00:00Z', endDateTime: '2026-03-01T11:00:00Z', joinWebUrl: 'https://teams.microsoft.com/1' },
          { id: 'meet-2', subject: 'Standup', startDateTime: '2026-03-02T09:00:00Z', endDateTime: '2026-03-02T09:15:00Z', joinWebUrl: 'https://teams.microsoft.com/2' },
        ]);

        const result = await repository.listOnlineMeetingsAsync();

        expect(result).toHaveLength(2);
        expect(result[0].id).toMatch(/^om_/);
        expect(result[0].id).not.toBe(result[1].id);
        expect(result[0].subject).toBe('Sprint Planning');
        expect(result[1].subject).toBe('Standup');
        expect(mockClient.listOnlineMeetings).toHaveBeenCalledWith(20);
      });

      it('passes a custom limit', async () => {
        mockClient.listOnlineMeetings.mockResolvedValue([]);

        await repository.listOnlineMeetingsAsync(5);

        expect(mockClient.listOnlineMeetings).toHaveBeenCalledWith(5);
      });
    });

    describe('getOnlineMeetingAsync', () => {
      it('resolves the om_ token to the Graph id', async () => {
        mockClient.listOnlineMeetings.mockResolvedValue([
          { id: 'meet-1', subject: 'Sprint Planning', startDateTime: '2026-03-01T10:00:00Z', endDateTime: '2026-03-01T11:00:00Z', joinWebUrl: 'https://teams.microsoft.com/1' },
        ]);
        const meetings = await repository.listOnlineMeetingsAsync();
        const meetTok = meetings[0].id;

        mockClient.getOnlineMeeting.mockResolvedValue({
          subject: 'Sprint Planning',
          startDateTime: '2026-03-01T10:00:00Z',
          endDateTime: '2026-03-01T11:00:00Z',
          joinWebUrl: 'https://teams.microsoft.com/1',
          participants: { organizer: {} },
        });

        const result = await repository.getOnlineMeetingAsync(meetTok);

        expect(result?.id).toBe(meetTok);
        expect(result?.subject).toBe('Sprint Planning');
        expect(mockClient.getOnlineMeeting).toHaveBeenCalledWith('meet-1');
      });

      it('re-lists on a cold-miss om_ token then resolves', async () => {
        // An om_ token minted in a prior session isn't in this store's alias
        // table; getOnlineMeetingAsync re-lists (deterministic re-mint) and
        // retries the resolve, matching the resolveTeamId self-heal pattern.
        mockClient.listOnlineMeetings.mockResolvedValue([
          { id: 'meet-1', subject: 'Sprint Planning', startDateTime: '', endDateTime: '', joinWebUrl: '' },
        ]);
        const meetings = await repository.listOnlineMeetingsAsync();
        const meetTok = meetings[0].id;

        const fresh = StateStore.open({ dir: '/tmp/mcp-o365-repo-test-meetings-cold', warn: () => {} });
        const repo2 = createGraphRepository(undefined, fresh);
        const client2 = (repo2 as any).client;
        client2.listOnlineMeetings.mockResolvedValue([
          { id: 'meet-1', subject: 'Sprint Planning', startDateTime: '', endDateTime: '', joinWebUrl: '' },
        ]);
        client2.getOnlineMeeting.mockResolvedValue({ subject: 'Sprint Planning', participants: null });

        const result = await repo2.getOnlineMeetingAsync(meetTok);

        expect(result?.subject).toBe('Sprint Planning');
        expect(client2.listOnlineMeetings).toHaveBeenCalled();
        expect(client2.getOnlineMeeting).toHaveBeenCalledWith('meet-1');
      });

      it('returns undefined when the token is unresolvable even after re-list', async () => {
        mockClient.listOnlineMeetings.mockResolvedValue([]);

        const result = await repository.getOnlineMeetingAsync('om_bogus');

        expect(result).toBeUndefined();
      });
    });
  });

  // ===========================================================================
  // Meeting Recordings
  // ===========================================================================

  describe('Meeting Recordings', () => {
    describe('listMeetingRecordingsAsync', () => {
      it('resolves the parent om_ token and mints rc_ tokens', async () => {
        mockClient.listOnlineMeetings.mockResolvedValue([
          { id: 'meet-1', subject: 'Sprint Planning', startDateTime: '', endDateTime: '', joinWebUrl: '' },
        ]);
        const meetings = await repository.listOnlineMeetingsAsync();
        const meetTok = meetings[0].id;

        mockClient.listMeetingRecordings.mockResolvedValue([
          { id: 'rec-1', createdDateTime: '2026-03-01T11:05:00Z', recordingContentUrl: 'https://graph.microsoft.com/rec1' },
        ]);

        const result = await repository.listMeetingRecordingsAsync(meetTok);

        expect(result).toHaveLength(1);
        expect(result[0].id).toMatch(/^rc_/);
        expect(mockClient.listMeetingRecordings).toHaveBeenCalledWith('meet-1');
      });

      it('rejects an unknown parent om_ token', async () => {
        await expect(repository.listMeetingRecordingsAsync('om_bogus')).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('downloadMeetingRecordingAsync', () => {
      it('resolves the rc_ token to (meetingId, recordingId) and downloads', async () => {
        mockClient.listOnlineMeetings.mockResolvedValue([
          { id: 'meet-1', subject: '', startDateTime: '', endDateTime: '', joinWebUrl: '' },
        ]);
        const meetings = await repository.listOnlineMeetingsAsync();
        const meetTok = meetings[0].id;

        mockClient.listMeetingRecordings.mockResolvedValue([
          { id: 'rec-1', createdDateTime: '', recordingContentUrl: '' },
        ]);
        const recordings = await repository.listMeetingRecordingsAsync(meetTok);
        const recTok = recordings[0].id;

        mockClient.getMeetingRecordingContent.mockResolvedValue(new ArrayBuffer(8));

        const outputPath = await repository.downloadMeetingRecordingAsync(recTok, '/tmp/recording.mp4');

        expect(outputPath).toBe('/tmp/recording.mp4');
        expect(mockClient.getMeetingRecordingContent).toHaveBeenCalledWith('meet-1', 'rec-1');
        expect(fs.writeFileSync).toHaveBeenCalledWith('/tmp/recording.mp4', expect.any(Buffer));
      });

      it('rejects a legacy numeric recording id', async () => {
        await expect(repository.downloadMeetingRecordingAsync(999999, '/tmp/out.mp4')).rejects.toThrow('not supported');
      });

      it('rejects an unknown rc_ token', async () => {
        await expect(repository.downloadMeetingRecordingAsync('rc_bogus', '/tmp/out.mp4')).rejects.toThrow('Unknown or unresolvable');
      });
    });
  });

  // ===========================================================================
  // Meeting Transcripts
  // ===========================================================================

  describe('Meeting Transcripts', () => {
    describe('listMeetingTranscriptsAsync', () => {
      it('resolves the parent om_ token and mints tr_ tokens', async () => {
        mockClient.listOnlineMeetings.mockResolvedValue([
          { id: 'meet-1', subject: '', startDateTime: '', endDateTime: '', joinWebUrl: '' },
        ]);
        const meetings = await repository.listOnlineMeetingsAsync();
        const meetTok = meetings[0].id;

        mockClient.listMeetingTranscripts.mockResolvedValue([
          { id: 'tr-1', createdDateTime: '2026-03-01T11:05:00Z', contentUrl: 'https://graph.microsoft.com/tr1' },
        ]);

        const result = await repository.listMeetingTranscriptsAsync(meetTok);

        expect(result).toHaveLength(1);
        expect(result[0].id).toMatch(/^tr_/);
        expect(mockClient.listMeetingTranscripts).toHaveBeenCalledWith('meet-1');
      });

      it('rejects an unknown parent om_ token', async () => {
        await expect(repository.listMeetingTranscriptsAsync('om_bogus')).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('getMeetingTranscriptContentAsync', () => {
      it('resolves the tr_ token to (meetingId, transcriptId) and fetches content', async () => {
        mockClient.listOnlineMeetings.mockResolvedValue([
          { id: 'meet-1', subject: '', startDateTime: '', endDateTime: '', joinWebUrl: '' },
        ]);
        const meetings = await repository.listOnlineMeetingsAsync();
        const meetTok = meetings[0].id;

        mockClient.listMeetingTranscripts.mockResolvedValue([
          { id: 'tr-1', createdDateTime: '', contentUrl: '' },
        ]);
        const transcripts = await repository.listMeetingTranscriptsAsync(meetTok);
        const trTok = transcripts[0].id;

        mockClient.getMeetingTranscriptContent.mockResolvedValue('WEBVTT\n\ncontent');

        const content = await repository.getMeetingTranscriptContentAsync(trTok);

        expect(content).toBe('WEBVTT\n\ncontent');
        expect(mockClient.getMeetingTranscriptContent).toHaveBeenCalledWith('meet-1', 'tr-1', 'text/vtt');
      });

      it('passes a custom format', async () => {
        mockClient.listOnlineMeetings.mockResolvedValue([
          { id: 'meet-1', subject: '', startDateTime: '', endDateTime: '', joinWebUrl: '' },
        ]);
        const meetings = await repository.listOnlineMeetingsAsync();
        const meetTok = meetings[0].id;

        mockClient.listMeetingTranscripts.mockResolvedValue([
          { id: 'tr-1', createdDateTime: '', contentUrl: '' },
        ]);
        const transcripts = await repository.listMeetingTranscriptsAsync(meetTok);
        const trTok = transcripts[0].id;

        mockClient.getMeetingTranscriptContent.mockResolvedValue('plain text');

        await repository.getMeetingTranscriptContentAsync(trTok, 'text/plain');

        expect(mockClient.getMeetingTranscriptContent).toHaveBeenCalledWith('meet-1', 'tr-1', 'text/plain');
      });

      it('rejects a legacy numeric transcript id', async () => {
        await expect(repository.getMeetingTranscriptContentAsync(999999)).rejects.toThrow('not supported');
      });

      it('rejects an unknown tr_ token', async () => {
        await expect(repository.getMeetingTranscriptContentAsync('tr_bogus')).rejects.toThrow('Unknown or unresolvable');
      });
    });
  });

  // ===========================================================================
  // SharePoint Sites, Document Libraries & Library Items
  // ===========================================================================

  describe('SharePoint (durable si_ / dl_ / li_ tokens)', () => {
    describe('listSitesAsync / searchSitesAsync', () => {
      it('mints durable si_ tokens', async () => {
        mockClient.listFollowedSites.mockResolvedValue([
          { id: 'site-1', name: 'team', webUrl: 'https://contoso.sharepoint.com/sites/team', displayName: 'Team Site' },
          { id: 'site-2', name: 'hr', webUrl: 'https://contoso.sharepoint.com/sites/hr', displayName: 'HR Portal' },
        ]);

        const result = await repository.listSitesAsync();

        expect(result).toHaveLength(2);
        expect(result[0].id).toMatch(/^si_/);
        expect(result[0].id).not.toBe(result[1].id);
        expect(result[0].displayName).toBe('Team Site');
      });

      it('mints durable si_ tokens for search results', async () => {
        mockClient.searchSites.mockResolvedValue([
          { id: 'site-3', name: 'marketing', webUrl: 'https://contoso.sharepoint.com/sites/marketing', displayName: 'Marketing Hub' },
        ]);

        const result = await repository.searchSitesAsync('marketing');

        expect(mockClient.searchSites).toHaveBeenCalledWith('marketing');
        expect(result[0].id).toMatch(/^si_/);
      });
    });

    describe('getSiteAsync', () => {
      it('resolves the si_ token to the Graph id', async () => {
        mockClient.listFollowedSites.mockResolvedValue([
          { id: 'site-1', name: 'team', webUrl: 'https://contoso.sharepoint.com/sites/team', displayName: 'Team Site' },
        ]);
        const sites = await repository.listSitesAsync();
        const siteTok = sites[0].id;

        mockClient.getSite.mockResolvedValue({
          name: 'team', webUrl: 'https://contoso.sharepoint.com/sites/team',
          displayName: 'Team Site', description: 'Main collaboration site',
        });

        const result = await repository.getSiteAsync(siteTok);

        expect(mockClient.getSite).toHaveBeenCalledWith('site-1');
        expect(result.id).toBe(siteTok);
        expect(result.description).toBe('Main collaboration site');
      });

      it('rejects a legacy numeric site id', async () => {
        await expect(repository.getSiteAsync(999999)).rejects.toThrow('not supported');
      });

      it('rejects an unknown si_ token', async () => {
        await expect(repository.getSiteAsync('si_bogus')).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('listDocumentLibrariesAsync', () => {
      it('resolves the parent si_ token and mints dl_ tokens', async () => {
        mockClient.listFollowedSites.mockResolvedValue([
          { id: 'site-1', name: 'team', webUrl: '', displayName: 'Team Site' },
        ]);
        const sites = await repository.listSitesAsync();
        const siteTok = sites[0].id;

        mockClient.listDocumentLibraries.mockResolvedValue([
          { id: 'drive-1', name: 'Documents', webUrl: '', driveType: 'documentLibrary' },
        ]);

        const result = await repository.listDocumentLibrariesAsync(siteTok);

        expect(mockClient.listDocumentLibraries).toHaveBeenCalledWith('site-1');
        expect(result).toHaveLength(1);
        expect(result[0].id).toMatch(/^dl_/);
      });

      it('rejects an unknown parent si_ token', async () => {
        await expect(repository.listDocumentLibrariesAsync('si_bogus')).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('listLibraryItemsAsync', () => {
      it('resolves the parent dl_ token (driveId) and mints li_ tokens', async () => {
        mockClient.listFollowedSites.mockResolvedValue([
          { id: 'site-1', name: 'team', webUrl: '', displayName: 'Team Site' },
        ]);
        const sites = await repository.listSitesAsync();
        const siteTok = sites[0].id;

        mockClient.listDocumentLibraries.mockResolvedValue([
          { id: 'drive-1', name: 'Documents', webUrl: '', driveType: 'documentLibrary' },
        ]);
        const libraries = await repository.listDocumentLibrariesAsync(siteTok);
        const libTok = libraries[0].id;

        mockClient.listLibraryItems.mockResolvedValue([
          { id: 'item-1', name: 'Report.docx', size: 15000, webUrl: '', lastModifiedDateTime: '', folder: undefined },
          { id: 'item-2', name: 'Projects', size: 0, webUrl: '', lastModifiedDateTime: '', folder: {} },
        ]);

        const result = await repository.listLibraryItemsAsync(libTok);

        expect(mockClient.listLibraryItems).toHaveBeenCalledWith('drive-1', undefined);
        expect(result).toHaveLength(2);
        expect(result[0].id).toMatch(/^li_/);
        expect(result[0].isFolder).toBe(false);
        expect(result[1].isFolder).toBe(true);
      });

      it('resolves a folder li_ token to its itemId and browses into it', async () => {
        mockClient.listFollowedSites.mockResolvedValue([
          { id: 'site-1', name: 'team', webUrl: '', displayName: 'Team Site' },
        ]);
        const sites = await repository.listSitesAsync();
        const siteTok = sites[0].id;

        mockClient.listDocumentLibraries.mockResolvedValue([
          { id: 'drive-1', name: 'Documents', webUrl: '', driveType: 'documentLibrary' },
        ]);
        const libraries = await repository.listDocumentLibrariesAsync(siteTok);
        const libTok = libraries[0].id;

        mockClient.listLibraryItems.mockResolvedValue([
          { id: 'item-2', name: 'Projects', size: 0, webUrl: '', lastModifiedDateTime: '', folder: {} },
        ]);
        const items = await repository.listLibraryItemsAsync(libTok);
        const folderTok = items[0].id;

        mockClient.listLibraryItems.mockResolvedValue([
          { id: 'item-3', name: 'Proposal.pptx', size: 50000, webUrl: '', lastModifiedDateTime: '', folder: undefined },
        ]);

        const result = await repository.listLibraryItemsAsync(libTok, folderTok);

        expect(mockClient.listLibraryItems).toHaveBeenCalledWith('drive-1', 'item-2');
        expect(result[0].name).toBe('Proposal.pptx');
      });

      it('rejects an unknown parent dl_ token', async () => {
        await expect(repository.listLibraryItemsAsync('dl_bogus')).rejects.toThrow('Unknown or unresolvable');
      });

      it('rejects an unknown folder li_ token', async () => {
        mockClient.listFollowedSites.mockResolvedValue([
          { id: 'site-1', name: 'team', webUrl: '', displayName: 'Team Site' },
        ]);
        const sites = await repository.listSitesAsync();
        const siteTok = sites[0].id;

        mockClient.listDocumentLibraries.mockResolvedValue([
          { id: 'drive-1', name: 'Documents', webUrl: '', driveType: 'documentLibrary' },
        ]);
        const libraries = await repository.listDocumentLibrariesAsync(siteTok);
        const libTok = libraries[0].id;

        await expect(repository.listLibraryItemsAsync(libTok, 'li_bogus')).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('downloadLibraryFileAsync', () => {
      it('resolves the li_ token to (driveId, itemId) and downloads', async () => {
        mockClient.listFollowedSites.mockResolvedValue([
          { id: 'site-1', name: 'team', webUrl: '', displayName: 'Team Site' },
        ]);
        const sites = await repository.listSitesAsync();
        const siteTok = sites[0].id;

        mockClient.listDocumentLibraries.mockResolvedValue([
          { id: 'drive-1', name: 'Documents', webUrl: '', driveType: 'documentLibrary' },
        ]);
        const libraries = await repository.listDocumentLibrariesAsync(siteTok);
        const libTok = libraries[0].id;

        mockClient.listLibraryItems.mockResolvedValue([
          { id: 'item-1', name: 'Report.docx', size: 15000, webUrl: '', lastModifiedDateTime: '', folder: undefined },
        ]);
        const items = await repository.listLibraryItemsAsync(libTok);
        const itemTok = items[0].id;

        mockClient.downloadLibraryFile.mockResolvedValue(new ArrayBuffer(8));

        const outputPath = await repository.downloadLibraryFileAsync(itemTok, '/tmp/Report.docx');

        expect(mockClient.downloadLibraryFile).toHaveBeenCalledWith('drive-1', 'item-1');
        expect(outputPath).toBe('/tmp/Report.docx');
        expect(fs.writeFileSync).toHaveBeenCalledWith('/tmp/Report.docx', expect.any(Buffer));
      });

      it('rejects a legacy numeric item id', async () => {
        await expect(repository.downloadLibraryFileAsync(999999, '/tmp/out.docx')).rejects.toThrow('not supported');
      });

      it('rejects an unknown li_ token', async () => {
        await expect(repository.downloadLibraryFileAsync('li_bogus', '/tmp/out.docx')).rejects.toThrow('Unknown or unresolvable');
      });
    });
  });

  // ===========================================================================
  // SharePoint Lists (durable sl_ / sn_ tokens)
  // ===========================================================================

  describe('SharePoint Lists (durable sl_ / sn_ tokens)', () => {
    // Mints a site token the list chain hangs off of.
    async function siteToken(): Promise<string> {
      mockClient.listFollowedSites.mockResolvedValue([
        { id: 'site-1', name: 'team', webUrl: '', displayName: 'Team Site' },
      ]);
      const sites = await repository.listSitesAsync();
      return sites[0].id;
    }

    // Mints an sl_ list token under site-1 / list-1.
    async function listToken(): Promise<string> {
      const siteTok = await siteToken();
      mockClient.listSharePointLists.mockResolvedValue([
        { id: 'list-1', name: 'announcements', displayName: 'Announcements', description: '', webUrl: '' },
      ]);
      const lists = await repository.listSharePointListsAsync(siteTok);
      return lists[0].id;
    }

    describe('listSharePointListsAsync', () => {
      it('resolves the parent si_ token and mints sl_ tokens', async () => {
        const siteTok = await siteToken();
        mockClient.listSharePointLists.mockResolvedValue([
          { id: 'list-1', name: 'announcements', displayName: 'Announcements', description: 'News', webUrl: 'https://x/1' },
          { id: 'list-2', name: 'tasks', displayName: 'Tasks', description: '', webUrl: 'https://x/2' },
        ]);

        const result = await repository.listSharePointListsAsync(siteTok);

        expect(mockClient.listSharePointLists).toHaveBeenCalledWith('site-1');
        expect(result).toHaveLength(2);
        expect(result[0].id).toMatch(/^sl_/);
        expect(result[0].id).not.toBe(result[1].id);
        expect(result[0].displayName).toBe('Announcements');
      });

      it('rejects an unknown parent si_ token', async () => {
        await expect(repository.listSharePointListsAsync('si_bogus')).rejects.toThrow('Unknown or unresolvable');
      });

      it('skips lists with no id rather than minting an empty-tuple token', async () => {
        const siteTok = await siteToken();
        mockClient.listSharePointLists.mockResolvedValue([
          { id: 'list-1', displayName: 'Real', name: '', description: '', webUrl: '' },
          { displayName: 'No Id', name: '', description: '', webUrl: '' },
        ]);

        const result = await repository.listSharePointListsAsync(siteTok);

        expect(result).toHaveLength(1);
        expect(result[0].displayName).toBe('Real');
      });
    });

    describe('getSharePointListAsync', () => {
      it('resolves the sl_ token to (siteId, listId) and echoes the token', async () => {
        const listTok = await listToken();
        mockClient.getSharePointList.mockResolvedValue({
          name: 'announcements', displayName: 'Announcements', description: 'News', webUrl: 'https://x/1',
        });

        const result = await repository.getSharePointListAsync(listTok);

        expect(mockClient.getSharePointList).toHaveBeenCalledWith('site-1', 'list-1');
        expect(result.id).toBe(listTok);
        expect(result.displayName).toBe('Announcements');
      });

      it('rejects a legacy numeric list id', async () => {
        await expect(repository.getSharePointListAsync(999999)).rejects.toThrow('not supported');
      });

      it('rejects an unknown sl_ token', async () => {
        await expect(repository.getSharePointListAsync('sl_bogus')).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('createSharePointListAsync', () => {
      it('resolves the parent si_ token, creates, and mints an sl_ token', async () => {
        const siteTok = await siteToken();
        mockClient.createSharePointList.mockResolvedValue({ id: 'list-9' });

        const token = await repository.createSharePointListAsync(siteTok, 'My List', 'desc');

        expect(mockClient.createSharePointList).toHaveBeenCalledWith('site-1', {
          displayName: 'My List',
          list: { template: 'genericList' },
          description: 'desc',
        });
        expect(token).toMatch(/^sl_/);
        // The minted token round-trips to the same (siteId, listId) tuple.
        mockClient.getSharePointList.mockResolvedValue({ displayName: 'My List' });
        await repository.getSharePointListAsync(token);
        expect(mockClient.getSharePointList).toHaveBeenCalledWith('site-1', 'list-9');
      });

      it('throws rather than minting a poisoned token when Graph returns no id', async () => {
        const siteTok = await siteToken();
        mockClient.createSharePointList.mockResolvedValue({});

        await expect(repository.createSharePointListAsync(siteTok, 'My List')).rejects.toThrow('returned no id');
      });
    });

    describe('listSharePointListColumnsAsync', () => {
      it('resolves the sl_ token and derives the column type from its facet', async () => {
        const listTok = await listToken();
        mockClient.listSharePointListColumns.mockResolvedValue([
          { id: 'c1', name: 'Title', displayName: 'Title', text: {}, required: true, readOnly: false },
          { id: 'c2', name: 'Count', displayName: 'Count', number: {}, required: false, readOnly: false },
          { id: 'c3', name: 'Weird', displayName: 'Weird' },
        ]);

        const result = await repository.listSharePointListColumnsAsync(listTok);

        expect(mockClient.listSharePointListColumns).toHaveBeenCalledWith('site-1', 'list-1');
        expect(result[0].columnType).toBe('text');
        expect(result[1].columnType).toBe('number');
        expect(result[2].columnType).toBe('unknown');
        // Columns carry no durable token.
        expect(result[0].id).toBe('c1');
      });

      it('rejects an unknown sl_ token', async () => {
        await expect(repository.listSharePointListColumnsAsync('sl_bogus')).rejects.toThrow('Unknown or unresolvable');
      });

      it('pins facet precedence when a definition carries more than one facet', async () => {
        const listTok = await listToken();
        // Graph normally emits one type facet; pin the first-match order so a
        // future reorder of the facet list is caught. 'text' precedes 'number'.
        mockClient.listSharePointListColumns.mockResolvedValue([
          { id: 'c1', name: 'Both', displayName: 'Both', text: {}, number: {} },
        ]);

        const result = await repository.listSharePointListColumnsAsync(listTok);

        expect(result[0].columnType).toBe('text');
      });
    });

    describe('listSharePointListItemsAsync', () => {
      it('resolves the sl_ token and mints sn_ tokens carrying the field values', async () => {
        const listTok = await listToken();
        mockClient.listSharePointListItems.mockResolvedValue([
          { id: '1', fields: { Title: 'First' }, webUrl: 'https://x/i1', createdDateTime: 'c', lastModifiedDateTime: 'm' },
          { id: '2', fields: { Title: 'Second' }, webUrl: '', createdDateTime: '', lastModifiedDateTime: '' },
        ]);

        const result = await repository.listSharePointListItemsAsync(listTok, 25);

        expect(mockClient.listSharePointListItems).toHaveBeenCalledWith('site-1', 'list-1', 25);
        expect(result).toHaveLength(2);
        expect(result[0].id).toMatch(/^sn_/);
        expect(result[0].fields).toEqual({ Title: 'First' });
      });

      it('rejects an unknown sl_ token', async () => {
        await expect(repository.listSharePointListItemsAsync('sl_bogus')).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('getSharePointListItemAsync', () => {
      it('resolves the sn_ token to (siteId, listId, itemId) and echoes the token', async () => {
        const listTok = await listToken();
        mockClient.listSharePointListItems.mockResolvedValue([
          { id: '1', fields: { Title: 'First' }, webUrl: '', createdDateTime: '', lastModifiedDateTime: '' },
        ]);
        const items = await repository.listSharePointListItemsAsync(listTok);
        const itemTok = items[0].id;

        mockClient.getSharePointListItem.mockResolvedValue({ fields: { Title: 'First' }, webUrl: 'https://x/i1' });

        const result = await repository.getSharePointListItemAsync(itemTok);

        expect(mockClient.getSharePointListItem).toHaveBeenCalledWith('site-1', 'list-1', '1');
        expect(result.id).toBe(itemTok);
        expect(result.fields).toEqual({ Title: 'First' });
      });

      it('rejects an unknown sn_ token', async () => {
        await expect(repository.getSharePointListItemAsync('sn_bogus')).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('createSharePointListItemAsync', () => {
      it('resolves the parent sl_ token, creates, and mints an sn_ token', async () => {
        const listTok = await listToken();
        mockClient.createSharePointListItem.mockResolvedValue({ id: '99' });

        const token = await repository.createSharePointListItemAsync(listTok, { Title: 'New' });

        expect(mockClient.createSharePointListItem).toHaveBeenCalledWith('site-1', 'list-1', { Title: 'New' });
        expect(token).toMatch(/^sn_/);
        // The minted token round-trips to (site-1, list-1, 99).
        mockClient.getSharePointListItem.mockResolvedValue({ fields: {} });
        await repository.getSharePointListItemAsync(token);
        expect(mockClient.getSharePointListItem).toHaveBeenCalledWith('site-1', 'list-1', '99');
      });

      it('rejects an unknown parent sl_ token', async () => {
        await expect(repository.createSharePointListItemAsync('sl_bogus', { Title: 'x' })).rejects.toThrow('Unknown or unresolvable');
      });

      it('throws rather than minting a poisoned token when Graph returns no id', async () => {
        const listTok = await listToken();
        mockClient.createSharePointListItem.mockResolvedValue({});

        await expect(repository.createSharePointListItemAsync(listTok, { Title: 'x' })).rejects.toThrow('returned no id');
      });

      it('mints the same sn_ token a later list call would mint for the same item (determinism)', async () => {
        const listTok = await listToken();
        mockClient.createSharePointListItem.mockResolvedValue({ id: '7' });
        const createdTok = await repository.createSharePointListItemAsync(listTok, { Title: 'x' });

        mockClient.listSharePointListItems.mockResolvedValue([
          { id: '7', fields: {}, webUrl: '', createdDateTime: '', lastModifiedDateTime: '' },
        ]);
        const listed = await repository.listSharePointListItemsAsync(listTok);

        expect(listed[0].id).toBe(createdTok);
      });
    });

    describe('updateSharePointListItemAsync', () => {
      it('resolves the sn_ token and patches the fields', async () => {
        const listTok = await listToken();
        mockClient.listSharePointListItems.mockResolvedValue([
          { id: '1', fields: {}, webUrl: '', createdDateTime: '', lastModifiedDateTime: '' },
        ]);
        const items = await repository.listSharePointListItemsAsync(listTok);

        await repository.updateSharePointListItemAsync(items[0].id, { Title: 'Updated' });

        expect(mockClient.updateSharePointListItem).toHaveBeenCalledWith('site-1', 'list-1', '1', { Title: 'Updated' });
      });

      it('rejects an unknown sn_ token', async () => {
        await expect(repository.updateSharePointListItemAsync('sn_bogus', {})).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('deleteSharePointListItemAsync', () => {
      it('resolves the sn_ token and deletes', async () => {
        const listTok = await listToken();
        mockClient.listSharePointListItems.mockResolvedValue([
          { id: '1', fields: {}, webUrl: '', createdDateTime: '', lastModifiedDateTime: '' },
        ]);
        const items = await repository.listSharePointListItemsAsync(listTok);

        await repository.deleteSharePointListItemAsync(items[0].id);

        expect(mockClient.deleteSharePointListItem).toHaveBeenCalledWith('site-1', 'list-1', '1');
      });

      it('rejects an unknown sn_ token', async () => {
        await expect(repository.deleteSharePointListItemAsync('sn_bogus')).rejects.toThrow('Unknown or unresolvable');
      });
    });
  });

  // ===========================================================================
  // Room Lists & Rooms
  // ===========================================================================

  describe('room lists & rooms', () => {
    describe('listRoomListsAsync', () => {
      it('returns mapped room lists', async () => {
        mockClient.listRoomLists.mockResolvedValue([
          { name: 'Building A', address: 'buildinga@example.com' },
          { name: 'Building B', address: 'buildingb@example.com' },
        ]);

        const result = await repository.listRoomListsAsync();

        expect(result).toEqual([
          { name: 'Building A', address: 'buildinga@example.com' },
          { name: 'Building B', address: 'buildingb@example.com' },
        ]);
        expect(mockClient.listRoomLists).toHaveBeenCalled();
      });

      it('handles empty room lists', async () => {
        mockClient.listRoomLists.mockResolvedValue([]);

        const result = await repository.listRoomListsAsync();

        expect(result).toEqual([]);
      });

      it('defaults name and address to empty string when null', async () => {
        mockClient.listRoomLists.mockResolvedValue([
          { name: null, address: null },
        ]);

        const result = await repository.listRoomListsAsync();

        expect(result).toEqual([{ name: '', address: '' }]);
      });
    });

    describe('listRoomsAsync', () => {
      it('returns mapped rooms without filter', async () => {
        mockClient.listRooms.mockResolvedValue([
          { name: 'Room 101', address: 'room101@example.com' },
        ]);

        const result = await repository.listRoomsAsync();

        expect(result).toEqual([{ name: 'Room 101', address: 'room101@example.com' }]);
        expect(mockClient.listRooms).toHaveBeenCalledWith(undefined);
      });

      it('passes room list email filter', async () => {
        mockClient.listRooms.mockResolvedValue([
          { name: 'Room 201', address: 'room201@example.com' },
        ]);

        const result = await repository.listRoomsAsync('buildinga@example.com');

        expect(result).toEqual([{ name: 'Room 201', address: 'room201@example.com' }]);
        expect(mockClient.listRooms).toHaveBeenCalledWith('buildinga@example.com');
      });

      it('handles empty rooms', async () => {
        mockClient.listRooms.mockResolvedValue([]);

        const result = await repository.listRoomsAsync();

        expect(result).toEqual([]);
      });

      it('defaults name and address to empty string when null', async () => {
        mockClient.listRooms.mockResolvedValue([
          { name: null, address: null },
        ]);

        const result = await repository.listRoomsAsync();

        expect(result).toEqual([{ name: '', address: '' }]);
      });
    });
  });

  describe('Teams (durable tm_ / cn_ tokens)', () => {
    // Helper: list teams and return the durable tm_ token for the given Graph id.
    async function teamToken(graphId: string, displayName = 'Eng', description = ''): Promise<string> {
      mockClient.listJoinedTeams.mockResolvedValue([{ id: graphId, displayName, description }]);
      const teams = await repository.listTeamsAsync();
      return teams[0].id;
    }
    // Helper: list channels under a team token and return the first cn_ token.
    async function channelToken(teamTok: string, channel: Record<string, unknown>): Promise<string> {
      mockClient.listChannels.mockResolvedValue([channel]);
      const channels = await repository.listChannelsAsync(teamTok);
      return channels[0].id;
    }

    describe('listTeamsAsync', () => {
      it('mints durable tm_ tokens', async () => {
        mockClient.listJoinedTeams.mockResolvedValue([
          { id: 'team-abc', displayName: 'Engineering', description: 'Eng team' },
          { id: 'team-def', displayName: 'Marketing', description: 'Mktg team' },
        ]);

        const result = await repository.listTeamsAsync();

        expect(result).toHaveLength(2);
        expect(result[0].id).toMatch(/^tm_/);
        expect(result[0].id).not.toBe(result[1].id);
        expect(result[0].name).toBe('Engineering');
        expect(result[0].description).toBe('Eng team');
        expect(result[1].name).toBe('Marketing');
      });

      it('the tm_ token resolves for follow-up channel calls', async () => {
        const tok = await teamToken('team-abc');
        mockClient.listChannels.mockResolvedValue([]);
        await expect(repository.listChannelsAsync(tok)).resolves.toEqual([]);
        expect(mockClient.listChannels).toHaveBeenCalledWith('team-abc');
      });

      it('defaults displayName and description to empty string when null', async () => {
        mockClient.listJoinedTeams.mockResolvedValue([
          { id: 'team-null', displayName: null, description: null },
        ]);

        const result = await repository.listTeamsAsync();

        expect(result[0].name).toBe('');
        expect(result[0].description).toBe('');
      });
    });

    describe('listChannelsAsync', () => {
      it('mints durable cn_ tokens', async () => {
        const tok = await teamToken('team-abc');

        mockClient.listChannels.mockResolvedValue([
          { id: 'ch-1', displayName: 'General', description: 'Default', membershipType: 'standard' },
          { id: 'ch-2', displayName: 'Dev', description: '', membershipType: 'private' },
        ]);

        const result = await repository.listChannelsAsync(tok);

        expect(result).toHaveLength(2);
        expect(result[0].id).toMatch(/^cn_/);
        expect(result[0].id).not.toBe(result[1].id);
        expect(result[0].name).toBe('General');
        expect(result[0].membershipType).toBe('standard');
        expect(result[1].membershipType).toBe('private');
        expect(mockClient.listChannels).toHaveBeenCalledWith('team-abc');
      });

      it('rejects a legacy numeric team id', async () => {
        await expect(repository.listChannelsAsync(999999)).rejects.toThrow('not supported');
      });

      it('re-lists on a cold-miss tm_ token then resolves', async () => {
        // A tm_ token minted in a prior session isn't in this store; resolveTeamId
        // re-lists (deterministic re-mint) and resolves.
        const tok = await teamToken('team-abc');
        const fresh = StateStore.open({ dir: '/tmp/mcp-o365-repo-test-2', warn: () => {} });
        const repo2 = createGraphRepository(undefined, fresh);
        const client2 = (repo2 as any).client;
        client2.listJoinedTeams.mockResolvedValue([{ id: 'team-abc', displayName: 'Eng', description: '' }]);
        client2.listChannels.mockResolvedValue([]);
        await expect(repo2.listChannelsAsync(tok)).resolves.toEqual([]);
        expect(client2.listJoinedTeams).toHaveBeenCalled();
      });

      it('defaults fields to empty/standard when null', async () => {
        const tok = await teamToken('team-abc');

        mockClient.listChannels.mockResolvedValue([
          { id: 'ch-null', displayName: null, description: null, membershipType: null },
        ]);

        const result = await repository.listChannelsAsync(tok);

        expect(result[0].name).toBe('');
        expect(result[0].description).toBe('');
        expect(result[0].membershipType).toBe('standard');
      });
    });

    describe('getChannelAsync', () => {
      it('resolves the cn_ token to (teamId, channelId)', async () => {
        const tok = await teamToken('team-abc');
        const chTok = await channelToken(tok, { id: 'ch-1', displayName: 'General', description: '', membershipType: 'standard' });

        mockClient.getChannel.mockResolvedValue({
          id: 'ch-1',
          displayName: 'General',
          description: 'The general channel',
          membershipType: 'standard',
          webUrl: 'https://teams.microsoft.com/...',
        });

        const result = await repository.getChannelAsync(chTok);

        expect(result.id).toBe(chTok);
        expect(result.name).toBe('General');
        expect(result.webUrl).toBe('https://teams.microsoft.com/...');
        expect(mockClient.getChannel).toHaveBeenCalledWith('team-abc', 'ch-1');
      });

      it('rejects an unknown channel token', async () => {
        await expect(repository.getChannelAsync('cn_bogus')).rejects.toThrow('Unknown or unresolvable');
      });

      it('a cn_ token does NOT self-heal across a cold store (documented alias tradeoff)', async () => {
        // Unlike a tm_ token (which re-lists its parent), a composite cn_ token
        // can't self-heal on a cold store — it has no parent handle to re-list
        // from, so a fresh store yields ID_UNKNOWN. This locks the documented
        // machine-scoped contract for alias-backed composites.
        const tok = await teamToken('team-abc');
        const chTok = await channelToken(tok, { id: 'ch-1', displayName: 'General', description: '', membershipType: 'standard' });

        const fresh = StateStore.open({ dir: '/tmp/mcp-o365-repo-test-cold', warn: () => {} });
        const repo2 = createGraphRepository(undefined, fresh);
        await expect(repo2.getChannelAsync(chTok)).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('createChannelAsync', () => {
      it('creates a channel and returns a resolvable cn_ token', async () => {
        const tok = await teamToken('team-abc');

        mockClient.createChannel.mockResolvedValue({ id: 'ch-new', displayName: 'New Channel' });

        const chTok = await repository.createChannelAsync(tok, 'New Channel', 'Description');

        expect(chTok).toMatch(/^cn_/);
        expect(mockClient.createChannel).toHaveBeenCalledWith('team-abc', 'New Channel', 'Description');

        mockClient.getChannel.mockResolvedValue({
          id: 'ch-new', displayName: 'New Channel', description: 'Description',
          membershipType: 'standard', webUrl: '',
        });
        const channel = await repository.getChannelAsync(chTok);
        expect(channel.name).toBe('New Channel');
        expect(mockClient.getChannel).toHaveBeenCalledWith('team-abc', 'ch-new');
      });

      it('rejects a legacy numeric team id', async () => {
        await expect(repository.createChannelAsync(999999, 'Test')).rejects.toThrow('not supported');
      });
    });

    describe('updateChannelAsync', () => {
      it('updates channel with mapped properties', async () => {
        const tok = await teamToken('team-abc');
        const chTok = await channelToken(tok, { id: 'ch-1', displayName: 'General', description: '', membershipType: 'standard' });

        mockClient.updateChannel.mockResolvedValue(undefined);

        await repository.updateChannelAsync(chTok, { name: 'Renamed', description: 'New desc' });

        expect(mockClient.updateChannel).toHaveBeenCalledWith('team-abc', 'ch-1', {
          displayName: 'Renamed',
          description: 'New desc',
        });
      });

      it('only sends provided fields', async () => {
        const tok = await teamToken('team-abc');
        const chTok = await channelToken(tok, { id: 'ch-1', displayName: 'General', description: '', membershipType: 'standard' });

        mockClient.updateChannel.mockResolvedValue(undefined);

        await repository.updateChannelAsync(chTok, { name: 'Renamed' });

        expect(mockClient.updateChannel).toHaveBeenCalledWith('team-abc', 'ch-1', {
          displayName: 'Renamed',
        });
      });

      it('rejects an unknown channel token', async () => {
        await expect(repository.updateChannelAsync('cn_bogus', { name: 'Test' })).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('deleteChannelAsync', () => {
      it('resolves the token and deletes the channel', async () => {
        const tok = await teamToken('team-abc');
        const chTok = await channelToken(tok, { id: 'ch-1', displayName: 'General', description: '', membershipType: 'standard' });

        mockClient.deleteChannel.mockResolvedValue(undefined);

        await repository.deleteChannelAsync(chTok);

        expect(mockClient.deleteChannel).toHaveBeenCalledWith('team-abc', 'ch-1');
      });

      it('rejects an unknown channel token', async () => {
        await expect(repository.deleteChannelAsync('cn_bogus')).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('listTeamMembersAsync', () => {
      it('returns mapped team members', async () => {
        const tok = await teamToken('team-abc');

        mockClient.listTeamMembers.mockResolvedValue([
          { id: 'member-1', displayName: 'Alice', email: 'alice@example.com', roles: ['owner'] },
          { id: 'member-2', displayName: 'Bob', email: 'bob@example.com', roles: [] },
        ]);

        const result = await repository.listTeamMembersAsync(tok);

        expect(result).toHaveLength(2);
        expect(result[0]).toEqual({
          id: 'member-1',
          displayName: 'Alice',
          email: 'alice@example.com',
          roles: ['owner'],
        });
        expect(result[1].roles).toEqual([]);
        expect(mockClient.listTeamMembers).toHaveBeenCalledWith('team-abc');
      });

      it('rejects a legacy numeric team id', async () => {
        await expect(repository.listTeamMembersAsync(999999)).rejects.toThrow('not supported');
      });

      it('defaults fields to empty when null', async () => {
        const tok = await teamToken('team-abc');

        mockClient.listTeamMembers.mockResolvedValue([
          { id: null, displayName: null, email: null, roles: null },
        ]);

        const result = await repository.listTeamMembersAsync(tok);

        expect(result[0]).toEqual({
          id: '',
          displayName: '',
          email: '',
          roles: [],
        });
      });
    });
  });

  describe('Chats & Messages (durable ch_ / cm_ / xm_ tokens)', () => {
    // Helper: list teams and return the durable tm_ token for the given Graph id.
    async function teamToken(graphId: string, displayName = 'Eng', description = ''): Promise<string> {
      mockClient.listJoinedTeams.mockResolvedValue([{ id: graphId, displayName, description }]);
      const teams = await repository.listTeamsAsync();
      return teams[0].id;
    }
    // Helper: list channels under a team token and return the first cn_ token.
    async function channelToken(teamTok: string, channel: Record<string, unknown>): Promise<string> {
      mockClient.listChannels.mockResolvedValue([channel]);
      const channels = await repository.listChannelsAsync(teamTok);
      return channels[0].id;
    }
    // Helper: list chats and return the durable ch_ token for the given Graph id.
    async function chatToken(graphId: string, topic = 'Chat'): Promise<string> {
      mockClient.listChats.mockResolvedValue([{ id: graphId, topic, chatType: 'oneOnOne', createdDateTime: '' }]);
      const chats = await repository.listChatsAsync();
      return chats[0].id;
    }

    describe('listChatsAsync', () => {
      it('mints durable ch_ tokens', async () => {
        mockClient.listChats.mockResolvedValue([
          { id: 'chat-abc', topic: 'Project Chat', chatType: 'group', createdDateTime: '2026-01-01T00:00:00Z' },
          { id: 'chat-def', topic: '', chatType: 'oneOnOne', createdDateTime: '2026-01-02T00:00:00Z' },
        ]);

        const result = await repository.listChatsAsync();

        expect(result).toHaveLength(2);
        expect(result[0].id).toMatch(/^ch_/);
        expect(result[0].id).not.toBe(result[1].id);
        expect(result[0].topic).toBe('Project Chat');
        expect(result[1].chatType).toBe('oneOnOne');
        expect(mockClient.listChats).toHaveBeenCalledWith(25, { expandMembers: false });
      });

      it('includes members when expandMembers is true', async () => {
        mockClient.listChats.mockResolvedValue([
          {
            id: 'chat-abc',
            topic: '',
            chatType: 'oneOnOne',
            createdDateTime: '2026-01-01T00:00:00Z',
            members: [
              { displayName: 'Alice', email: 'alice@example.com', userId: 'u1', roles: ['owner'] },
            ],
          },
        ]);

        const result = await repository.listChatsAsync(10, true);

        expect(mockClient.listChats).toHaveBeenCalledWith(10, { expandMembers: true });
        expect(result[0].members).toEqual([
          { displayName: 'Alice', email: 'alice@example.com', userId: 'u1', roles: ['owner'] },
        ]);
      });
    });

    describe('findChatsAsync', () => {
      it('uses get-or-create for a single email 1:1 lookup', async () => {
        mockClient.createChat.mockResolvedValue({
          id: 'chat-1to1',
          topic: '',
          chatType: 'oneOnOne',
          createdDateTime: '2026-01-01T00:00:00Z',
          members: [
            { displayName: 'Me', email: 'me@example.com', userId: 'me', roles: ['owner'] },
            { displayName: 'Alice', email: 'alice@example.com', userId: 'u1', roles: ['owner'] },
          ],
        });

        const result = await repository.findChatsAsync({ participants: ['alice@example.com'] });

        expect(mockClient.createChat).toHaveBeenCalledWith('oneOnOne', ['alice@example.com']);
        expect(mockClient.listChats).not.toHaveBeenCalled();
        expect(result).toHaveLength(1);
        expect(result[0].id).toMatch(/^ch_/);
        expect(result[0].members[1].email).toBe('alice@example.com');
      });

      it('matches email exactly (case-insensitive) and returns all candidates', async () => {
        mockClient.listChats.mockResolvedValue([
          {
            id: 'chat-a',
            topic: 'A',
            chatType: 'group',
            createdDateTime: '',
            members: [
              { displayName: 'Alice', email: 'alice@example.com', userId: 'u1', roles: [] },
              { displayName: 'Bob', email: 'bob@example.com', userId: 'u2', roles: [] },
            ],
          },
          {
            id: 'chat-b',
            topic: 'B',
            chatType: 'group',
            createdDateTime: '',
            members: [
              { displayName: 'Alice', email: 'Alice@Example.com', userId: 'u1', roles: [] },
              { displayName: 'Carol', email: 'carol@example.com', userId: 'u3', roles: [] },
            ],
          },
          {
            id: 'chat-c',
            topic: 'C',
            chatType: 'group',
            createdDateTime: '',
            members: [
              { displayName: 'Dave', email: 'dave@example.com', userId: 'u4', roles: [] },
            ],
          },
        ]);

        const result = await repository.findChatsAsync({
          participants: ['alice@example.com'],
          chatType: 'group',
        });

        expect(mockClient.createChat).not.toHaveBeenCalled();
        expect(result.map((c) => c.topic).sort()).toEqual(['A', 'B']);
      });

      it('falls back to case-insensitive displayName and returns all matches', async () => {
        mockClient.listChats.mockResolvedValue([
          {
            id: 'chat-1',
            topic: 'One',
            chatType: 'group',
            createdDateTime: '',
            members: [{ displayName: 'Alice Smith', email: '', userId: 'u1', roles: [] }],
          },
          {
            id: 'chat-2',
            topic: 'Two',
            chatType: 'group',
            createdDateTime: '',
            members: [{ displayName: 'alice smith', email: '', userId: 'u2', roles: [] }],
          },
        ]);

        const result = await repository.findChatsAsync({ participants: ['Alice Smith'] });

        expect(result).toHaveLength(2);
      });
    });

    describe('resolveOrCreateChatAsync', () => {
      it('resolves a 1:1 chat from a single email', async () => {
        mockClient.createChat.mockResolvedValue({
          id: 'chat-1to1',
          topic: '',
          chatType: 'oneOnOne',
          createdDateTime: '',
          members: [{ displayName: 'Alice', email: 'alice@example.com', userId: 'u1', roles: [] }],
        });

        const result = await repository.resolveOrCreateChatAsync(['alice@example.com']);

        expect(result).toEqual({ chatId: expect.stringMatching(/^ch_/) });
      });

      it('returns an error with candidates when multiple group chats match', async () => {
        mockClient.listChats.mockResolvedValue([
          {
            id: 'chat-a',
            topic: 'A',
            chatType: 'group',
            createdDateTime: '',
            members: [
              { displayName: 'Alice', email: 'alice@example.com', userId: 'u1', roles: [] },
              { displayName: 'Bob', email: 'bob@example.com', userId: 'u2', roles: [] },
            ],
          },
          {
            id: 'chat-b',
            topic: 'B',
            chatType: 'group',
            createdDateTime: '',
            members: [
              { displayName: 'Alice', email: 'alice@example.com', userId: 'u1', roles: [] },
              { displayName: 'Bob', email: 'bob@example.com', userId: 'u2', roles: [] },
            ],
          },
        ]);

        const result = await repository.resolveOrCreateChatAsync(['alice@example.com', 'bob@example.com']);

        expect('error' in result).toBe(true);
        if ('error' in result) {
          expect(result.chats).toHaveLength(2);
        }
      });
    });

    describe('getChatAsync', () => {
      it('the ch_ token resolves for follow-up calls', async () => {
        const tok = await chatToken('chat-abc');

        mockClient.getChat.mockResolvedValue({
          id: 'chat-abc', topic: 'Project Chat', chatType: 'group', createdDateTime: '2026-01-01T00:00:00Z', webUrl: 'https://teams.microsoft.com/...',
        });

        const result = await repository.getChatAsync(tok);

        expect(result.id).toBe(tok);
        expect(result.topic).toBe('Project Chat');
        expect(mockClient.getChat).toHaveBeenCalledWith('chat-abc');
      });

      it('rejects a legacy numeric chat id', async () => {
        await expect(repository.getChatAsync(999999)).rejects.toThrow('not supported');
      });

      it('rejects an unknown chat token once re-list finds no match', async () => {
        mockClient.listChats.mockResolvedValue([]);
        await expect(repository.getChatAsync('ch_bogus')).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('listChatMessagesAsync', () => {
      it('mints durable cm_ tokens', async () => {
        const tok = await chatToken('chat-abc');

        mockClient.listChatMessages.mockResolvedValue([
          {
            id: 'msg-1', from: { user: { displayName: 'Alice' } },
            body: { content: 'Hello', contentType: 'text' }, createdDateTime: '2026-01-01T00:00:00Z',
          },
        ]);

        const result = await repository.listChatMessagesAsync(tok);

        expect(result[0].id).toMatch(/^cm_/);
        expect(result[0].senderName).toBe('Alice');
        expect(mockClient.listChatMessages).toHaveBeenCalledWith('chat-abc', 25);
      });

      it('rejects a legacy numeric chat id', async () => {
        await expect(repository.listChatMessagesAsync(999999)).rejects.toThrow('not supported');
      });
    });

    describe('sendChatMessageAsync', () => {
      it('sends and returns a resolvable cm_ token', async () => {
        const tok = await chatToken('chat-abc');
        mockClient.sendChatMessage.mockResolvedValue({ id: 'msg-new' });

        const msgTok = await repository.sendChatMessageAsync(tok, 'Hello!', 'text');

        expect(msgTok).toMatch(/^cm_/);
        expect(mockClient.sendChatMessage).toHaveBeenCalledWith('chat-abc', 'Hello!', 'text');
      });
    });

    describe('listChatMembersAsync', () => {
      it('resolves the ch_ token and lists members', async () => {
        const tok = await chatToken('chat-abc');
        mockClient.listChatMembers.mockResolvedValue([
          { displayName: 'Alice', email: 'alice@example.com', roles: ['owner'] },
        ]);

        const result = await repository.listChatMembersAsync(tok);

        expect(result[0].displayName).toBe('Alice');
        expect(mockClient.listChatMembers).toHaveBeenCalledWith('chat-abc');
      });
    });

    describe('listChannelMessagesAsync', () => {
      it('mints durable xm_ tokens', async () => {
        const teamTok = await teamToken('team-abc');
        const chanTok = await channelToken(teamTok, { id: 'ch-1', displayName: 'General', description: '', membershipType: 'standard' });

        mockClient.listChannelMessages.mockResolvedValue([
          {
            id: 'cmsg-1', from: { user: { displayName: 'Alice' } },
            body: { content: 'Hello', contentType: 'text' }, createdDateTime: '2026-01-01T00:00:00Z',
          },
        ]);

        const result = await repository.listChannelMessagesAsync(chanTok);

        expect(result[0].id).toMatch(/^xm_/);
        expect(mockClient.listChannelMessages).toHaveBeenCalledWith('team-abc', 'ch-1', 25);
      });
    });

    describe('getChannelMessageAsync', () => {
      it('resolves the xm_ token to (teamId, channelId, messageId) and mints reply tokens', async () => {
        const teamTok = await teamToken('team-abc');
        const chanTok = await channelToken(teamTok, { id: 'ch-1', displayName: 'General', description: '', membershipType: 'standard' });

        mockClient.listChannelMessages.mockResolvedValue([
          {
            id: 'cmsg-1', from: { user: { displayName: 'Alice' } },
            body: { content: 'Hello', contentType: 'text' }, createdDateTime: '2026-01-01T00:00:00Z',
          },
        ]);
        const msgTok = (await repository.listChannelMessagesAsync(chanTok))[0].id;

        mockClient.getChannelMessage.mockResolvedValue({
          id: 'cmsg-1', from: { user: { displayName: 'Alice' } },
          body: { content: 'Hello', contentType: 'text' }, createdDateTime: '2026-01-01T00:00:00Z',
        });
        mockClient.listChannelMessageReplies.mockResolvedValue([
          {
            id: 'reply-1', from: { user: { displayName: 'Bob' } },
            body: { content: 'Hi back', contentType: 'text' }, createdDateTime: '2026-01-01T01:00:00Z',
          },
        ]);

        const result = await repository.getChannelMessageAsync(msgTok);

        expect(result.id).toBe(msgTok);
        expect(result.replies[0].id).toMatch(/^xm_/);
        expect(mockClient.getChannelMessage).toHaveBeenCalledWith('team-abc', 'ch-1', 'cmsg-1');
        expect(mockClient.listChannelMessageReplies).toHaveBeenCalledWith('team-abc', 'ch-1', 'cmsg-1');
      });

      it('rejects an unknown message token (no self-heal for composites)', async () => {
        await expect(repository.getChannelMessageAsync('xm_bogus')).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('sendChannelMessageAsync', () => {
      it('sends and returns a resolvable xm_ token', async () => {
        const teamTok = await teamToken('team-abc');
        const chanTok = await channelToken(teamTok, { id: 'ch-1', displayName: 'General', description: '', membershipType: 'standard' });
        mockClient.sendChannelMessage.mockResolvedValue({ id: 'cmsg-new' });

        const msgTok = await repository.sendChannelMessageAsync(chanTok, 'Hello channel!', 'text');

        expect(msgTok).toMatch(/^xm_/);
        expect(mockClient.sendChannelMessage).toHaveBeenCalledWith('team-abc', 'ch-1', 'Hello channel!', 'text');
      });
    });

    describe('replyToChannelMessageAsync', () => {
      it('resolves the xm_ token and returns a new resolvable xm_ token', async () => {
        const teamTok = await teamToken('team-abc');
        const chanTok = await channelToken(teamTok, { id: 'ch-1', displayName: 'General', description: '', membershipType: 'standard' });
        mockClient.sendChannelMessage.mockResolvedValue({ id: 'cmsg-1' });
        const msgTok = await repository.sendChannelMessageAsync(chanTok, 'Hello!', 'text');

        mockClient.replyToChannelMessage.mockResolvedValue({ id: 'reply-1' });
        const replyTok = await repository.replyToChannelMessageAsync(msgTok, 'Great idea!', 'text');

        expect(replyTok).toMatch(/^xm_/);
        expect(mockClient.replyToChannelMessage).toHaveBeenCalledWith('team-abc', 'ch-1', 'cmsg-1', 'Great idea!', 'text');
      });

      it('rejects an unknown message token', async () => {
        await expect(repository.replyToChannelMessageAsync('xm_bogus', 'Reply')).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('Message Reactions', () => {
      it('lists / adds / removes reactions for a channel message by xm_ token', async () => {
        const teamTok = await teamToken('team-abc');
        const chanTok = await channelToken(teamTok, { id: 'ch-1', displayName: 'General', description: '', membershipType: 'standard' });
        mockClient.sendChannelMessage.mockResolvedValue({ id: 'cmsg-1' });
        const msgTok = await repository.sendChannelMessageAsync(chanTok, 'Hello!', 'text');

        mockClient.getChannelMessage.mockResolvedValue({
          reactions: [{ reactionType: 'like', user: { user: { displayName: 'Alice' } }, createdDateTime: '2026-01-01T00:00:00Z' }],
        });
        const reactions = await repository.listMessageReactionsAsync(msgTok, 'channel');
        expect(reactions[0].reactionType).toBe('like');
        expect(mockClient.getChannelMessage).toHaveBeenCalledWith('team-abc', 'ch-1', 'cmsg-1');

        mockClient.setChannelMessageReaction.mockResolvedValue(undefined);
        await repository.addMessageReactionAsync(msgTok, 'channel', 'heart');
        expect(mockClient.setChannelMessageReaction).toHaveBeenCalledWith('team-abc', 'ch-1', 'cmsg-1', 'heart');

        mockClient.unsetChannelMessageReaction.mockResolvedValue(undefined);
        await repository.removeMessageReactionAsync(msgTok, 'channel', 'heart');
        expect(mockClient.unsetChannelMessageReaction).toHaveBeenCalledWith('team-abc', 'ch-1', 'cmsg-1', 'heart');
      });

      it('lists / adds / removes reactions for a chat message by cm_ token', async () => {
        const chatTok = await chatToken('chat-abc');
        mockClient.sendChatMessage.mockResolvedValue({ id: 'msg-1' });
        const msgTok = await repository.sendChatMessageAsync(chatTok, 'Hello!', 'text');

        mockClient.getChatMessage.mockResolvedValue({
          reactions: [{ reactionType: 'laugh', user: { user: { displayName: 'Charlie' } }, createdDateTime: '2026-01-02T00:00:00Z' }],
        });
        const reactions = await repository.listMessageReactionsAsync(msgTok, 'chat');
        expect(reactions[0].reactionType).toBe('laugh');
        expect(mockClient.getChatMessage).toHaveBeenCalledWith('chat-abc', 'msg-1');

        mockClient.setChatMessageReaction.mockResolvedValue(undefined);
        await repository.addMessageReactionAsync(msgTok, 'chat', 'like');
        expect(mockClient.setChatMessageReaction).toHaveBeenCalledWith('chat-abc', 'msg-1', 'like');

        mockClient.unsetChatMessageReaction.mockResolvedValue(undefined);
        await repository.removeMessageReactionAsync(msgTok, 'chat', 'like');
        expect(mockClient.unsetChatMessageReaction).toHaveBeenCalledWith('chat-abc', 'msg-1', 'like');
      });

      it('rejects an unknown message token', async () => {
        await expect(repository.listMessageReactionsAsync('xm_bogus', 'channel')).rejects.toThrow('Unknown or unresolvable');
        await expect(repository.listMessageReactionsAsync('cm_bogus', 'chat')).rejects.toThrow('Unknown or unresolvable');
      });

      it('fails closed when message_type disagrees with the token kind', async () => {
        // A cm_ (chat) token asked to resolve as a channel message — and the
        // reverse — must ID_ENTITY_MISMATCH, never silently mis-resolve to the
        // wrong Graph collection. The token kind is authoritative.
        const chatTok = await chatToken('chat-abc');
        mockClient.sendChatMessage.mockResolvedValue({ id: 'msg-1' });
        const cmTok = await repository.sendChatMessageAsync(chatTok, 'Hi', 'text');
        await expect(repository.addMessageReactionAsync(cmTok, 'channel', 'like')).rejects.toThrow(/but a channelMessage ID was expected/);

        const teamTok = await teamToken('team-abc');
        const chTok = await channelToken(teamTok, { id: 'ch-1', displayName: 'General', description: '', membershipType: 'standard' });
        mockClient.sendChannelMessage.mockResolvedValue({ id: 'cmsg-1' });
        const xmTok = await repository.sendChannelMessageAsync(chTok, 'Yo', 'text');
        await expect(repository.addMessageReactionAsync(xmTok, 'chat', 'like')).rejects.toThrow(/but a chatMessage ID was expected/);
      });
    });
  });

  describe('Planner (durable pl_ / pb_ / pt_ tokens, U5b-5 fetch-before-update)', () => {
    // Helper: list plans and return the durable pl_ token for the given Graph id.
    async function planToken(graphId: string, title = 'Plan'): Promise<string> {
      mockClient.listPlans.mockResolvedValue([{ id: graphId, title, owner: 'group-1', createdDateTime: '' }]);
      const plans = await repository.listPlansAsync();
      return plans[0].id;
    }
    // Helper: list buckets under a plan token and return the durable pb_ token.
    async function bucketToken(planTok: string, bucketGraphId: string, name = 'Bucket'): Promise<string> {
      mockClient.listBuckets.mockResolvedValue([{ id: bucketGraphId, name, orderHint: '1' }]);
      const buckets = await repository.listBucketsAsync(planTok);
      return buckets[0].id;
    }
    // Helper: list tasks under a plan token and return the durable pt_ token.
    async function taskToken(planTok: string, taskGraphId: string, title = 'Task'): Promise<string> {
      mockClient.listPlannerTasks.mockResolvedValue([{
        id: taskGraphId, title, bucketId: null, assignments: null,
        percentComplete: 0, priority: 5, startDateTime: '', dueDateTime: '', createdDateTime: '',
      }]);
      const tasks = await repository.listPlannerTasksAsync(planTok);
      return tasks[0].id;
    }

    describe('Plans', () => {
      it('listPlansAsync mints durable pl_ tokens', async () => {
        mockClient.listPlans.mockResolvedValue([
          { id: 'graph-plan-1', title: 'Sprint Plan', owner: 'group-abc', createdDateTime: '2026-01-01T00:00:00Z' },
        ]);

        const plans = await repository.listPlansAsync();

        expect(plans[0].id).toMatch(/^pl_/);
        expect(plans[0].title).toBe('Sprint Plan');
      });

      it('getPlanAsync resolves the pl_ token and returns a freshly-fetched etag', async () => {
        const planTok = await planToken('graph-plan-1', 'Sprint Plan');
        mockClient.getPlan.mockResolvedValue({ title: 'Sprint Plan', owner: 'group-abc', createdDateTime: '', '@odata.etag': 'W/"plan-etag"' });

        const plan = await repository.getPlanAsync(planTok);

        expect(mockClient.getPlan).toHaveBeenCalledWith('graph-plan-1');
        expect(plan.etag).toBe('W/"plan-etag"');
      });

      it('createPlanAsync mints a resolvable pl_ token', async () => {
        mockClient.createPlan.mockResolvedValue({ id: 'graph-plan-new', '@odata.etag': 'W/"etag"' });

        const planTok = await repository.createPlanAsync('New Plan', 'group-xyz');

        expect(planTok).toMatch(/^pl_/);
        expect(mockClient.createPlan).toHaveBeenCalledWith('New Plan', 'group-xyz');
        // The minted token resolves cold — no separate list needed.
        mockClient.getPlan.mockResolvedValue({ title: 'New Plan', '@odata.etag': 'W/"etag"' });
        await repository.getPlanAsync(planTok);
        expect(mockClient.getPlan).toHaveBeenCalledWith('graph-plan-new');
      });

      it('updatePlanAsync fetches a fresh etag immediately before the write (U5b-5)', async () => {
        const planTok = await planToken('graph-plan-1');
        const callOrder: string[] = [];
        mockClient.getPlan.mockImplementation(async () => {
          callOrder.push('get');
          return { title: 'Sprint Plan', '@odata.etag': 'W/"etag-fresh"' };
        });
        mockClient.updatePlan.mockImplementation(async () => {
          callOrder.push('update');
          return { '@odata.etag': 'W/"etag-fresh"' };
        });

        await repository.updatePlanAsync(planTok, { title: 'Renamed' });

        expect(callOrder).toEqual(['get', 'update']);
        expect(mockClient.updatePlan).toHaveBeenCalledWith('graph-plan-1', { title: 'Renamed' }, 'W/"etag-fresh"');
      });

      it('updatePlanAsync retries once on a 412 with a re-fetched etag', async () => {
        const planTok = await planToken('graph-plan-1');
        mockClient.getPlan
          .mockResolvedValueOnce({ title: 'Sprint Plan', '@odata.etag': 'W/"etag-stale"' })
          .mockResolvedValueOnce({ title: 'Sprint Plan', '@odata.etag': 'W/"etag-refreshed"' });
        mockClient.updatePlan
          .mockRejectedValueOnce({ statusCode: 412 })
          .mockResolvedValueOnce({ title: 'Renamed', '@odata.etag': 'W/"etag-refreshed"' });

        await repository.updatePlanAsync(planTok, { title: 'Renamed' });

        expect(mockClient.getPlan).toHaveBeenCalledTimes(2);
        expect(mockClient.updatePlan).toHaveBeenCalledTimes(2);
        expect(mockClient.updatePlan).toHaveBeenNthCalledWith(1, 'graph-plan-1', { title: 'Renamed' }, 'W/"etag-stale"');
        expect(mockClient.updatePlan).toHaveBeenNthCalledWith(2, 'graph-plan-1', { title: 'Renamed' }, 'W/"etag-refreshed"');
      });

      it('propagates a second 412 without retrying a third time', async () => {
        const planTok = await planToken('graph-plan-1');
        mockClient.getPlan.mockResolvedValue({ title: 'P', '@odata.etag': 'W/"e"' });
        mockClient.updatePlan
          .mockRejectedValueOnce({ statusCode: 412 })
          .mockRejectedValueOnce({ statusCode: 412 });

        await expect(repository.updatePlanAsync(planTok, { title: 'X' })).rejects.toMatchObject({ statusCode: 412 });
        expect(mockClient.updatePlan).toHaveBeenCalledTimes(2);
        expect(mockClient.getPlan).toHaveBeenCalledTimes(2);
      });

      it('does NOT retry on a non-412 write error', async () => {
        const planTok = await planToken('graph-plan-1');
        mockClient.getPlan.mockResolvedValue({ title: 'P', '@odata.etag': 'W/"e"' });
        mockClient.updatePlan.mockRejectedValueOnce({ statusCode: 500 });

        await expect(repository.updatePlanAsync(planTok, { title: 'X' })).rejects.toMatchObject({ statusCode: 500 });
        expect(mockClient.updatePlan).toHaveBeenCalledTimes(1);
        expect(mockClient.getPlan).toHaveBeenCalledTimes(1);
      });

      it('fails loudly when the fetched entity has no @odata.etag (never sends an empty If-Match)', async () => {
        const planTok = await planToken('graph-plan-1');
        mockClient.getPlan.mockResolvedValue({ title: 'P' });

        await expect(repository.updatePlanAsync(planTok, { title: 'X' })).rejects.toThrow(/no @odata\.etag/);
        expect(mockClient.updatePlan).not.toHaveBeenCalled();
      });

      it('resolvePlanId re-lists on a cold-miss pl_ token then resolves', async () => {
        const planTok = await planToken('graph-plan-1');
        // Simulate a cold store: a fresh repository sharing no prior list.
        const freshStore = StateStore.open({ dir: '/tmp/mcp-o365-repo-test-planner-cold', warn: () => {} });
        const repo2 = createGraphRepository(undefined, freshStore);
        const client2 = (repo2 as any).client;
        client2.listPlans.mockResolvedValue([{ id: 'graph-plan-1', title: 'Sprint Plan', owner: '', createdDateTime: '' }]);
        client2.getPlan.mockResolvedValue({ title: 'Sprint Plan', '@odata.etag': 'W/"etag"' });

        const plan = await repo2.getPlanAsync(planTok);
        expect(client2.getPlan).toHaveBeenCalledWith('graph-plan-1');
        expect(plan.title).toBe('Sprint Plan');
      });

      it('rejects a legacy numeric plan id', async () => {
        await expect(repository.getPlanAsync(123456)).rejects.toThrow('not supported');
      });

      it('rejects an unknown pl_ token', async () => {
        mockClient.listPlans.mockResolvedValue([]);
        await expect(repository.getPlanAsync('pl_bogus00000000')).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('Buckets', () => {
      it('listBucketsAsync mints durable pb_ tokens', async () => {
        const planTok = await planToken('graph-plan-1');
        mockClient.listBuckets.mockResolvedValue([
          { id: 'graph-bucket-1', name: 'To Do', orderHint: '1' },
        ]);

        const buckets = await repository.listBucketsAsync(planTok);

        expect(buckets[0].id).toMatch(/^pb_/);
        expect(buckets[0].planId).toBe(planTok);
      });

      it('createBucketAsync mints a resolvable pb_ token', async () => {
        const planTok = await planToken('graph-plan-1');
        mockClient.createBucket.mockResolvedValue({ id: 'graph-bucket-new', '@odata.etag': 'W/"etag"' });

        const bucketTok = await repository.createBucketAsync(planTok, 'Done');

        expect(bucketTok).toMatch(/^pb_/);
        expect(mockClient.createBucket).toHaveBeenCalledWith('graph-plan-1', 'Done');
      });

      it('updateBucketAsync fetches a fresh etag immediately before the write (U5b-5)', async () => {
        const planTok = await planToken('graph-plan-1');
        const bucketTok = await bucketToken(planTok, 'graph-bucket-1');
        const callOrder: string[] = [];
        mockClient.getBucket.mockImplementation(async () => {
          callOrder.push('get');
          return { name: 'To Do', '@odata.etag': 'W/"bucket-etag-fresh"' };
        });
        mockClient.updateBucket.mockImplementation(async () => {
          callOrder.push('update');
          return { '@odata.etag': 'W/"bucket-etag-fresh"' };
        });

        await repository.updateBucketAsync(bucketTok, { name: 'Renamed' });

        expect(callOrder).toEqual(['get', 'update']);
        expect(mockClient.getBucket).toHaveBeenCalledWith('graph-bucket-1');
        expect(mockClient.updateBucket).toHaveBeenCalledWith('graph-bucket-1', { name: 'Renamed' }, 'W/"bucket-etag-fresh"');
      });

      it('updateBucketAsync retries once on a 412 with a re-fetched etag', async () => {
        const planTok = await planToken('graph-plan-1');
        const bucketTok = await bucketToken(planTok, 'graph-bucket-1');
        mockClient.getBucket
          .mockResolvedValueOnce({ name: 'To Do', '@odata.etag': 'W/"bucket-etag-stale"' })
          .mockResolvedValueOnce({ name: 'To Do', '@odata.etag': 'W/"bucket-etag-refreshed"' });
        mockClient.updateBucket
          .mockRejectedValueOnce({ statusCode: 412 })
          .mockResolvedValueOnce({ name: 'Renamed', '@odata.etag': 'W/"bucket-etag-refreshed"' });

        await repository.updateBucketAsync(bucketTok, { name: 'Renamed' });

        expect(mockClient.getBucket).toHaveBeenCalledTimes(2);
        expect(mockClient.updateBucket).toHaveBeenCalledTimes(2);
        expect(mockClient.updateBucket).toHaveBeenNthCalledWith(1, 'graph-bucket-1', { name: 'Renamed' }, 'W/"bucket-etag-stale"');
        expect(mockClient.updateBucket).toHaveBeenNthCalledWith(2, 'graph-bucket-1', { name: 'Renamed' }, 'W/"bucket-etag-refreshed"');
      });

      it('deleteBucketAsync fetches a fresh etag immediately before the write', async () => {
        const planTok = await planToken('graph-plan-1');
        const bucketTok = await bucketToken(planTok, 'graph-bucket-1');
        mockClient.getBucket.mockResolvedValue({ name: 'To Do', '@odata.etag': 'W/"bucket-etag"' });
        mockClient.deleteBucket.mockResolvedValue(undefined);

        await repository.deleteBucketAsync(bucketTok);

        expect(mockClient.getBucket).toHaveBeenCalledWith('graph-bucket-1');
        expect(mockClient.deleteBucket).toHaveBeenCalledWith('graph-bucket-1', 'W/"bucket-etag"');
      });

      it('rejects an unknown pb_ token', async () => {
        await expect(repository.updateBucketAsync('pb_bogus00000000', { name: 'x' })).rejects.toThrow('Unknown or unresolvable');
      });
    });

    describe('Tasks', () => {
      it('listPlannerTasksAsync mints durable pt_ tokens and pb_ bucket tokens', async () => {
        const planTok = await planToken('graph-plan-1');
        mockClient.listPlannerTasks.mockResolvedValue([
          {
            id: 'graph-task-1', title: 'Ship v3.1', bucketId: 'graph-bucket-1',
            assignments: { 'user-1': {} }, percentComplete: 40, priority: 3,
            startDateTime: '2026-07-01T00:00:00Z', dueDateTime: '2026-07-15T00:00:00Z',
            createdDateTime: '2026-06-01T00:00:00Z',
          },
          {
            id: 'graph-task-2', title: 'Review', bucketId: null,
            assignments: null, percentComplete: 0, priority: 5,
            startDateTime: '', dueDateTime: '', createdDateTime: '2026-06-02T00:00:00Z',
          },
        ]);

        const tasks = await repository.listPlannerTasksAsync(planTok);

        expect(tasks[0].id).toMatch(/^pt_/);
        expect(tasks[0].bucketId).toMatch(/^pb_/);
        expect(tasks[1].bucketId).toBeNull();
      });

      it('getPlannerTaskAsync resolves the pt_ token', async () => {
        const planTok = await planToken('graph-plan-1');
        const taskTok = await taskToken(planTok, 'graph-task-1', 'Ship v3.1');
        mockClient.getPlannerTask.mockResolvedValue({
          title: 'Ship v3.1', bucketId: null, assignments: null, percentComplete: 40, priority: 3,
          startDateTime: '', dueDateTime: '', createdDateTime: '', conversationThreadId: 'thread-1',
          orderHint: '1', '@odata.etag': 'W/"task-etag"',
        });

        const task = await repository.getPlannerTaskAsync(taskTok);

        expect(mockClient.getPlannerTask).toHaveBeenCalledWith('graph-task-1');
        expect(task.etag).toBe('W/"task-etag"');
      });

      it('createPlannerTaskAsync resolves an optional bucket token to its Graph id', async () => {
        const planTok = await planToken('graph-plan-1');
        const bucketTok = await bucketToken(planTok, 'graph-bucket-1');
        mockClient.createPlannerTask.mockResolvedValue({ id: 'graph-task-new', '@odata.etag': 'W/"etag"' });

        const taskTok = await repository.createPlannerTaskAsync(planTok, 'New Task', bucketTok);

        expect(taskTok).toMatch(/^pt_/);
        expect(mockClient.createPlannerTask).toHaveBeenCalledWith({
          planId: 'graph-plan-1', title: 'New Task', bucketId: 'graph-bucket-1',
        });
      });

      it('updatePlannerTaskAsync fetches a fresh etag immediately before the write (U5b-5)', async () => {
        const planTok = await planToken('graph-plan-1');
        const taskTok = await taskToken(planTok, 'graph-task-1');
        const callOrder: string[] = [];
        mockClient.getPlannerTask.mockImplementation(async () => {
          callOrder.push('get');
          return { title: 'Task', '@odata.etag': 'W/"task-etag-fresh"' };
        });
        mockClient.updatePlannerTask.mockImplementation(async () => {
          callOrder.push('update');
          return { '@odata.etag': 'W/"task-etag-fresh"' };
        });

        await repository.updatePlannerTaskAsync(taskTok, { title: 'Renamed Task' });

        expect(callOrder).toEqual(['get', 'update']);
        expect(mockClient.updatePlannerTask).toHaveBeenCalledWith('graph-task-1', { title: 'Renamed Task' }, 'W/"task-etag-fresh"');
      });

      it('updatePlannerTaskAsync retries once on a 412 with a re-fetched etag', async () => {
        const planTok = await planToken('graph-plan-1');
        const taskTok = await taskToken(planTok, 'graph-task-1');
        mockClient.getPlannerTask
          .mockResolvedValueOnce({ title: 'Task', '@odata.etag': 'W/"task-etag-stale"' })
          .mockResolvedValueOnce({ title: 'Task', '@odata.etag': 'W/"task-etag-refreshed"' });
        mockClient.updatePlannerTask
          .mockRejectedValueOnce({ statusCode: 412 })
          .mockResolvedValueOnce({ title: 'Renamed Task', '@odata.etag': 'W/"task-etag-refreshed"' });

        await repository.updatePlannerTaskAsync(taskTok, { title: 'Renamed Task' });

        expect(mockClient.getPlannerTask).toHaveBeenCalledTimes(2);
        expect(mockClient.updatePlannerTask).toHaveBeenCalledTimes(2);
        expect(mockClient.updatePlannerTask).toHaveBeenNthCalledWith(1, 'graph-task-1', { title: 'Renamed Task' }, 'W/"task-etag-stale"');
        expect(mockClient.updatePlannerTask).toHaveBeenNthCalledWith(2, 'graph-task-1', { title: 'Renamed Task' }, 'W/"task-etag-refreshed"');
      });

      it('deletePlannerTaskAsync fetches a fresh etag immediately before the write', async () => {
        const planTok = await planToken('graph-plan-1');
        const taskTok = await taskToken(planTok, 'graph-task-1');
        mockClient.getPlannerTask.mockResolvedValue({ title: 'Task', '@odata.etag': 'W/"task-etag"' });
        mockClient.deletePlannerTask.mockResolvedValue(undefined);

        await repository.deletePlannerTaskAsync(taskTok);

        expect(mockClient.getPlannerTask).toHaveBeenCalledWith('graph-task-1');
        expect(mockClient.deletePlannerTask).toHaveBeenCalledWith('graph-task-1', 'W/"task-etag"');
      });

      it('rejects an unknown pt_ token', async () => {
        await expect(repository.getPlannerTaskAsync('pt_bogus00000000')).rejects.toThrow('Unknown or unresolvable');
      });

      it('rejects a legacy numeric task id', async () => {
        await expect(repository.getPlannerTaskAsync(654321)).rejects.toThrow('not supported');
      });
    });

    describe('Task details (piggyback pt_)', () => {
      it('getPlannerTaskDetailsAsync resolves the pt_ token', async () => {
        const planTok = await planToken('graph-plan-1');
        const taskTok = await taskToken(planTok, 'graph-task-1');
        mockClient.getPlannerTaskDetails.mockResolvedValue({
          description: 'Notes', checklist: {}, references: {}, '@odata.etag': 'W/"details-etag"',
        });

        const details = await repository.getPlannerTaskDetailsAsync(taskTok);

        expect(mockClient.getPlannerTaskDetails).toHaveBeenCalledWith('graph-task-1');
        expect(details.etag).toBe('W/"details-etag"');
      });

      it('updatePlannerTaskDetailsAsync fetches a fresh etag immediately before the write (U5b-5)', async () => {
        const planTok = await planToken('graph-plan-1');
        const taskTok = await taskToken(planTok, 'graph-task-1');
        const callOrder: string[] = [];
        mockClient.getPlannerTaskDetails.mockImplementation(async () => {
          callOrder.push('get');
          return { description: 'Notes', checklist: {}, references: {}, '@odata.etag': 'W/"details-etag-fresh"' };
        });
        mockClient.updatePlannerTaskDetails.mockImplementation(async () => {
          callOrder.push('update');
          return { '@odata.etag': 'W/"details-etag-fresh"' };
        });

        await repository.updatePlannerTaskDetailsAsync(taskTok, { description: 'Updated' });

        expect(callOrder).toEqual(['get', 'update']);
        expect(mockClient.updatePlannerTaskDetails).toHaveBeenCalledWith('graph-task-1', { description: 'Updated' }, 'W/"details-etag-fresh"');
      });

      it('updatePlannerTaskDetailsAsync retries once on a 412 with a re-fetched etag', async () => {
        const planTok = await planToken('graph-plan-1');
        const taskTok = await taskToken(planTok, 'graph-task-1');
        mockClient.getPlannerTaskDetails
          .mockResolvedValueOnce({ description: 'Notes', checklist: {}, references: {}, '@odata.etag': 'W/"details-etag-stale"' })
          .mockResolvedValueOnce({ description: 'Notes', checklist: {}, references: {}, '@odata.etag': 'W/"details-etag-refreshed"' });
        mockClient.updatePlannerTaskDetails
          .mockRejectedValueOnce({ statusCode: 412 })
          .mockResolvedValueOnce({ description: 'Updated', checklist: {}, references: {}, '@odata.etag': 'W/"details-etag-refreshed"' });

        await repository.updatePlannerTaskDetailsAsync(taskTok, { description: 'Updated' });

        expect(mockClient.getPlannerTaskDetails).toHaveBeenCalledTimes(2);
        expect(mockClient.updatePlannerTaskDetails).toHaveBeenCalledTimes(2);
        expect(mockClient.updatePlannerTaskDetails).toHaveBeenNthCalledWith(1, 'graph-task-1', { description: 'Updated' }, 'W/"details-etag-stale"');
        expect(mockClient.updatePlannerTaskDetails).toHaveBeenNthCalledWith(2, 'graph-task-1', { description: 'Updated' }, 'W/"details-etag-refreshed"');
      });
    });

    describe('listMyPlannerTasksAsync', () => {
      it('maps cross-plan tasks, minting a durable pl_/pb_ token per task', async () => {
        mockClient.listMyPlannerTasks.mockResolvedValue([
          {
            id: 'graph-task-1', title: 'Ship v3.1', planId: 'graph-plan-A', bucketId: 'graph-bucket-1',
            assignments: { 'user-1': {} }, percentComplete: 40, priority: 3,
            startDateTime: '2026-07-01T00:00:00Z', dueDateTime: '2026-07-15T00:00:00Z',
            createdDateTime: '2026-06-01T00:00:00Z', '@odata.etag': 'W/"etag1"',
          },
          {
            id: 'graph-task-2', title: 'Review', planId: 'graph-plan-B', bucketId: null,
            assignments: null, percentComplete: 0, priority: 5,
            startDateTime: '', dueDateTime: '', createdDateTime: '2026-06-02T00:00:00Z',
          },
        ]);

        const tasks = await repository.listMyPlannerTasksAsync();

        expect(tasks[0].id).toMatch(/^pt_/);
        expect(tasks[0].planId).toMatch(/^pl_/);
        expect(tasks[0].bucketId).toMatch(/^pb_/);
        expect(tasks[0].assignees).toEqual(['user-1']);
        expect(tasks[1].planId).toMatch(/^pl_/);
        expect(tasks[1].bucketId).toBeNull();
      });

      it('mints a resolvable pl_ token per task so a follow-up get_plan resolves without a re-list', async () => {
        mockClient.listMyPlannerTasks.mockResolvedValue([
          { id: 'graph-task-1', title: 'T', planId: 'graph-plan-A', bucketId: null, assignments: null,
            percentComplete: 0, priority: 5, startDateTime: '', dueDateTime: '', createdDateTime: '' },
        ]);

        const tasks = await repository.listMyPlannerTasksAsync();

        mockClient.getPlan.mockResolvedValue({ title: 'Plan A', '@odata.etag': 'W/"etag"' });
        await repository.getPlanAsync(tasks[0].planId);
        expect(mockClient.getPlan).toHaveBeenCalledWith('graph-plan-A');
      });
    });
  });

  describe('OneDrive drive items (dr_ tokens)', () => {
    it('mints dr_ tokens for listed items and does not crash on an id-less row', async () => {
      // An id-less row must degrade to id:'' (parity with the folder mapper's
      // empty-id guard, #46) — mintSelfEncoded throws on an empty id, so an
      // unguarded mint would abort the whole list and hide every valid sibling.
      mockClient.listDriveItems.mockResolvedValue([
        { id: 'drive-item-1', name: 'report.pdf', size: 10, lastModifiedDateTime: '', folder: null, webUrl: '' },
        { id: '', name: 'ghost', size: 0, lastModifiedDateTime: '', folder: null, webUrl: '' },
      ]);

      const result = await repository.listDriveItemsAsync();

      expect(result[0].id).toBe(mintSelfEncoded('driveItem', 'drive-item-1'));
      expect(result[1].id).toBe('');
    });
  });
});
