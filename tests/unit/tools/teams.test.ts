/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Microsoft Teams tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { TeamsTools, type ITeamsRepository } from '../../../src/tools/teams.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';

describe('TeamsTools', () => {
  let repo: ITeamsRepository;
  let tokenManager: ApprovalTokenManager;
  let tools: TeamsTools;

  beforeEach(() => {
    repo = {
      listTeamsAsync: vi.fn(),
      listChannelsAsync: vi.fn(),
      getChannelAsync: vi.fn(),
      createChannelAsync: vi.fn(),
      updateChannelAsync: vi.fn(),
      deleteChannelAsync: vi.fn(),
      listTeamMembersAsync: vi.fn(),
      listChannelMessagesAsync: vi.fn(),
      getChannelMessageAsync: vi.fn(),
      sendChannelMessageAsync: vi.fn(),
      replyToChannelMessageAsync: vi.fn(),
      listChatsAsync: vi.fn(),
      getChatAsync: vi.fn(),
      listChatMessagesAsync: vi.fn(),
      sendChatMessageAsync: vi.fn(),
      listChatMembersAsync: vi.fn(),
      listMessageReactionsAsync: vi.fn(),
      addMessageReactionAsync: vi.fn(),
      removeMessageReactionAsync: vi.fn(),
    };
    tokenManager = new ApprovalTokenManager();
    tools = new TeamsTools(repo, tokenManager);
  });

  describe('listTeams', () => {
    it('returns teams from the repository', async () => {
      const mockTeams = [
        { id: 1, name: 'Engineering', description: 'Eng team' },
        { id: 2, name: 'Marketing', description: 'Mktg team' },
      ];
      vi.mocked(repo.listTeamsAsync).mockResolvedValue(mockTeams);

      const result = await tools.listTeams();

      expect(result.content).toHaveLength(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.teams).toEqual(mockTeams);
    });
  });

  describe('listChannels', () => {
    it('returns channels for a team', async () => {
      const mockChannels = [
        { id: 10, name: 'General', description: 'Default', membershipType: 'standard' },
      ];
      vi.mocked(repo.listChannelsAsync).mockResolvedValue(mockChannels);

      const result = await tools.listChannels({ team_id: 1 });

      expect(repo.listChannelsAsync).toHaveBeenCalledWith(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.channels).toEqual(mockChannels);
    });
  });

  describe('getChannel', () => {
    it('returns channel details', async () => {
      const mockChannel = {
        id: 10, name: 'General', description: 'Default', membershipType: 'standard', webUrl: 'https://...',
      };
      vi.mocked(repo.getChannelAsync).mockResolvedValue(mockChannel);

      const result = await tools.getChannel({ channel_id: 10 });

      expect(repo.getChannelAsync).toHaveBeenCalledWith(10);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.channel).toEqual(mockChannel);
    });
  });

  describe('createChannel', () => {
    it('creates a channel and returns the ID', async () => {
      vi.mocked(repo.createChannelAsync).mockResolvedValue(42);

      const result = await tools.createChannel({ team_id: 1, name: 'Dev', description: 'Dev channel' });

      expect(repo.createChannelAsync).toHaveBeenCalledWith(1, 'Dev', 'Dev channel');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.channel_id).toBe(42);
      expect(parsed.message).toBe('Channel created');
    });

    it('creates a channel without description', async () => {
      vi.mocked(repo.createChannelAsync).mockResolvedValue(42);

      await tools.createChannel({ team_id: 1, name: 'Dev' });

      expect(repo.createChannelAsync).toHaveBeenCalledWith(1, 'Dev', undefined);
    });
  });

  describe('updateChannel', () => {
    it('updates a channel', async () => {
      vi.mocked(repo.updateChannelAsync).mockResolvedValue(undefined);

      const result = await tools.updateChannel({ channel_id: 10, name: 'Renamed', description: 'New desc' });

      expect(repo.updateChannelAsync).toHaveBeenCalledWith(10, { name: 'Renamed', description: 'New desc' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Channel updated');
    });
  });

  describe('prepareDeleteChannel', () => {
    it('generates an approval token', () => {
      const result = tools.prepareDeleteChannel({ channel_id: 42 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.channel_id).toBe(42);
      expect(parsed.action).toContain('confirm_delete_channel');
    });
  });

  describe('confirmDeleteChannel', () => {
    it('deletes the channel with a valid token', async () => {
      vi.mocked(repo.deleteChannelAsync).mockResolvedValue(undefined);

      // Generate a token first
      const prepareResult = tools.prepareDeleteChannel({ channel_id: 42 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeleteChannel({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Channel deleted');
      expect(repo.deleteChannelAsync).toHaveBeenCalledWith(42);
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmDeleteChannel({ approval_token: 'invalid-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.deleteChannelAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeleteChannel({ channel_id: 42 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      // Consume the token
      await tools.confirmDeleteChannel({ approval_token });

      // Try to consume again
      const result = await tools.confirmDeleteChannel({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });

  describe('listTeamMembers', () => {
    it('returns team members', async () => {
      const mockMembers = [
        { id: 'm-1', displayName: 'Alice', email: 'alice@example.com', roles: ['owner'] },
        { id: 'm-2', displayName: 'Bob', email: 'bob@example.com', roles: [] },
      ];
      vi.mocked(repo.listTeamMembersAsync).mockResolvedValue(mockMembers);

      const result = await tools.listTeamMembers({ team_id: 1 });

      expect(repo.listTeamMembersAsync).toHaveBeenCalledWith(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.members).toEqual(mockMembers);
    });
  });

  // ===========================================================================
  // Channel Messages
  // ===========================================================================

  describe('listChannelMessages', () => {
    it('returns messages from the repository', async () => {
      const mockMessages = [
        {
          id: 100, senderName: 'Alice', senderEmail: 'alice@example.com',
          bodyPreview: 'Hello world', bodyContent: 'Hello world',
          contentType: 'text', createdDateTime: '2026-01-01T00:00:00Z',
        },
        {
          id: 200, senderName: 'Bob', senderEmail: 'bob@example.com',
          bodyPreview: 'Hi there', bodyContent: 'Hi there',
          contentType: 'text', createdDateTime: '2026-01-01T01:00:00Z',
        },
      ];
      vi.mocked(repo.listChannelMessagesAsync).mockResolvedValue(mockMessages);

      const result = await tools.listChannelMessages({ channel_id: 10 });

      expect(repo.listChannelMessagesAsync).toHaveBeenCalledWith(10, undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.messages).toEqual(mockMessages);
    });

    it('passes limit parameter', async () => {
      vi.mocked(repo.listChannelMessagesAsync).mockResolvedValue([]);

      await tools.listChannelMessages({ channel_id: 10, limit: 5 });

      expect(repo.listChannelMessagesAsync).toHaveBeenCalledWith(10, 5);
    });
  });

  describe('getChannelMessage', () => {
    it('returns message with replies', async () => {
      const mockMessage = {
        id: 100, senderName: 'Alice', senderEmail: 'alice@example.com',
        bodyContent: '<p>Hello</p>', contentType: 'html',
        createdDateTime: '2026-01-01T00:00:00Z',
        replies: [
          {
            id: 101, senderName: 'Bob', senderEmail: 'bob@example.com',
            bodyContent: '<p>Hi back</p>', contentType: 'html',
            createdDateTime: '2026-01-01T01:00:00Z',
          },
        ],
      };
      vi.mocked(repo.getChannelMessageAsync).mockResolvedValue(mockMessage);

      const result = await tools.getChannelMessage({ message_id: 100 });

      expect(repo.getChannelMessageAsync).toHaveBeenCalledWith(100);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.message).toEqual(mockMessage);
      expect(parsed.message.replies).toHaveLength(1);
    });
  });

  describe('prepareSendChannelMessage', () => {
    it('generates an approval token with preview', () => {
      const result = tools.prepareSendChannelMessage({
        channel_id: 10, body: 'Hello channel!', content_type: 'text',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.channel_id).toBe(10);
      expect(parsed.body_preview).toBe('Hello channel!');
      expect(parsed.content_type).toBe('text');
      expect(parsed.action).toContain('confirm_send_channel_message');
    });

    it('defaults content_type to html', () => {
      const result = tools.prepareSendChannelMessage({
        channel_id: 10, body: '<p>Hello</p>',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.content_type).toBe('html');
    });
  });

  describe('confirmSendChannelMessage', () => {
    it('sends message with valid token', async () => {
      vi.mocked(repo.sendChannelMessageAsync).mockResolvedValue(999);

      const prepareResult = tools.prepareSendChannelMessage({
        channel_id: 10, body: 'Hello!', content_type: 'text',
      });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmSendChannelMessage({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message_id).toBe(999);
      expect(parsed.message).toBe('Message sent');
      expect(repo.sendChannelMessageAsync).toHaveBeenCalledWith(10, 'Hello!', 'text');
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmSendChannelMessage({ approval_token: 'bad-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.sendChannelMessageAsync).mockResolvedValue(999);

      const prepareResult = tools.prepareSendChannelMessage({
        channel_id: 10, body: 'Hello!',
      });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      await tools.confirmSendChannelMessage({ approval_token });
      const result = await tools.confirmSendChannelMessage({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });

  describe('prepareReplyChannelMessage', () => {
    it('generates an approval token with preview', () => {
      const result = tools.prepareReplyChannelMessage({
        message_id: 100, body: 'Nice post!', content_type: 'text',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.message_id).toBe(100);
      expect(parsed.body_preview).toBe('Nice post!');
      expect(parsed.content_type).toBe('text');
      expect(parsed.action).toContain('confirm_reply_channel_message');
    });

    it('defaults content_type to html', () => {
      const result = tools.prepareReplyChannelMessage({
        message_id: 100, body: '<p>Reply</p>',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.content_type).toBe('html');
    });
  });

  describe('confirmReplyChannelMessage', () => {
    it('sends reply with valid token', async () => {
      vi.mocked(repo.replyToChannelMessageAsync).mockResolvedValue(888);

      const prepareResult = tools.prepareReplyChannelMessage({
        message_id: 100, body: 'Great idea!', content_type: 'text',
      });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmReplyChannelMessage({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.reply_id).toBe(888);
      expect(parsed.message).toBe('Reply sent');
      expect(repo.replyToChannelMessageAsync).toHaveBeenCalledWith(100, 'Great idea!', 'text');
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmReplyChannelMessage({ approval_token: 'bad-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.replyToChannelMessageAsync).mockResolvedValue(888);

      const prepareResult = tools.prepareReplyChannelMessage({
        message_id: 100, body: 'Reply!',
      });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      await tools.confirmReplyChannelMessage({ approval_token });
      const result = await tools.confirmReplyChannelMessage({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });

  // ===========================================================================
  // Chats
  // ===========================================================================

  describe('listChats', () => {
    it('returns chats from the repository', async () => {
      const mockChats = [
        { id: 1, topic: 'Project Chat', chatType: 'group', lastMessagePreview: 'Hello', createdDateTime: '2026-01-01T00:00:00Z' },
        { id: 2, topic: '', chatType: 'oneOnOne', lastMessagePreview: 'Hi there', createdDateTime: '2026-01-02T00:00:00Z' },
      ];
      vi.mocked(repo.listChatsAsync).mockResolvedValue(mockChats);

      const result = await tools.listChats({});

      expect(repo.listChatsAsync).toHaveBeenCalledWith(undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.chats).toEqual(mockChats);
    });

    it('passes limit parameter', async () => {
      vi.mocked(repo.listChatsAsync).mockResolvedValue([]);

      await tools.listChats({ limit: 10 });

      expect(repo.listChatsAsync).toHaveBeenCalledWith(10);
    });
  });

  describe('getChat', () => {
    it('returns chat details', async () => {
      const mockChat = {
        id: 1, topic: 'Project Chat', chatType: 'group', createdDateTime: '2026-01-01T00:00:00Z', webUrl: 'https://teams.microsoft.com/...',
      };
      vi.mocked(repo.getChatAsync).mockResolvedValue(mockChat);

      const result = await tools.getChat({ chat_id: 1 });

      expect(repo.getChatAsync).toHaveBeenCalledWith(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.chat).toEqual(mockChat);
    });
  });

  describe('listChatMessages', () => {
    it('returns messages from the repository', async () => {
      const mockMessages = [
        {
          id: 100, senderName: 'Alice', senderEmail: 'alice@example.com',
          bodyPreview: 'Hello', bodyContent: 'Hello',
          contentType: 'text', createdDateTime: '2026-01-01T00:00:00Z',
        },
      ];
      vi.mocked(repo.listChatMessagesAsync).mockResolvedValue(mockMessages);

      const result = await tools.listChatMessages({ chat_id: 1 });

      expect(repo.listChatMessagesAsync).toHaveBeenCalledWith(1, undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.messages).toEqual(mockMessages);
    });

    it('passes limit parameter', async () => {
      vi.mocked(repo.listChatMessagesAsync).mockResolvedValue([]);

      await tools.listChatMessages({ chat_id: 1, limit: 5 });

      expect(repo.listChatMessagesAsync).toHaveBeenCalledWith(1, 5);
    });
  });

  describe('prepareSendChatMessage', () => {
    it('generates an approval token with preview', () => {
      const result = tools.prepareSendChatMessage({
        chat_id: 1, body: 'Hello chat!', content_type: 'text',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.chat_id).toBe(1);
      expect(parsed.body_preview).toBe('Hello chat!');
      expect(parsed.content_type).toBe('text');
      expect(parsed.action).toContain('confirm_send_chat_message');
    });

    it('defaults content_type to html', () => {
      const result = tools.prepareSendChatMessage({
        chat_id: 1, body: '<p>Hello</p>',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.content_type).toBe('html');
    });
  });

  describe('confirmSendChatMessage', () => {
    it('sends message with valid token', async () => {
      vi.mocked(repo.sendChatMessageAsync).mockResolvedValue(999);

      const prepareResult = tools.prepareSendChatMessage({
        chat_id: 1, body: 'Hello!', content_type: 'text',
      });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmSendChatMessage({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message_id).toBe(999);
      expect(parsed.message).toBe('Message sent');
      expect(repo.sendChatMessageAsync).toHaveBeenCalledWith(1, 'Hello!', 'text');
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmSendChatMessage({ approval_token: 'bad-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.sendChatMessageAsync).mockResolvedValue(999);

      const prepareResult = tools.prepareSendChatMessage({
        chat_id: 1, body: 'Hello!',
      });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      await tools.confirmSendChatMessage({ approval_token });
      const result = await tools.confirmSendChatMessage({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });

  describe('listChatMembers', () => {
    it('returns chat members', async () => {
      const mockMembers = [
        { displayName: 'Alice', email: 'alice@example.com', roles: ['owner'] },
        { displayName: 'Bob', email: 'bob@example.com', roles: [] },
      ];
      vi.mocked(repo.listChatMembersAsync).mockResolvedValue(mockMembers);

      const result = await tools.listChatMembers({ chat_id: 1 });

      expect(repo.listChatMembersAsync).toHaveBeenCalledWith(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.members).toEqual(mockMembers);
    });
  });

  // ===========================================================================
  // Message Reactions
  // ===========================================================================

  describe('listMessageReactions', () => {
    it('returns reactions for a channel message', async () => {
      const mockReactions = [
        { reactionType: 'like', user: { displayName: 'Alice' }, createdDateTime: '2026-01-01T00:00:00Z' },
        { reactionType: 'heart', user: { displayName: 'Bob' }, createdDateTime: '2026-01-01T01:00:00Z' },
      ];
      vi.mocked(repo.listMessageReactionsAsync).mockResolvedValue(mockReactions);

      const result = await tools.listMessageReactions({ message_id: 100, message_type: 'channel' });

      expect(repo.listMessageReactionsAsync).toHaveBeenCalledWith(100, 'channel');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.reactions).toEqual(mockReactions);
    });

    it('returns reactions for a chat message', async () => {
      const mockReactions = [
        { reactionType: 'laugh', user: { displayName: 'Charlie' }, createdDateTime: '2026-01-02T00:00:00Z' },
      ];
      vi.mocked(repo.listMessageReactionsAsync).mockResolvedValue(mockReactions);

      const result = await tools.listMessageReactions({ message_id: 200, message_type: 'chat' });

      expect(repo.listMessageReactionsAsync).toHaveBeenCalledWith(200, 'chat');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.reactions).toEqual(mockReactions);
    });
  });

  describe('prepareAddMessageReaction', () => {
    it('generates an approval token', () => {
      const result = tools.prepareAddMessageReaction({
        message_id: 100, message_type: 'channel', reaction_type: 'like',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.message_id).toBe(100);
      expect(parsed.message_type).toBe('channel');
      expect(parsed.reaction_type).toBe('like');
      expect(parsed.action).toContain('confirm_add_message_reaction');
    });
  });

  describe('confirmAddMessageReaction', () => {
    it('adds reaction with valid token', async () => {
      vi.mocked(repo.addMessageReactionAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareAddMessageReaction({
        message_id: 100, message_type: 'channel', reaction_type: 'like',
      });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmAddMessageReaction({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Reaction added');
      expect(repo.addMessageReactionAsync).toHaveBeenCalledWith(100, 'channel', 'like');
    });

    it('rejects invalid token', async () => {
      const result = await tools.confirmAddMessageReaction({ approval_token: 'invalid-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });
  });

  describe('removeMessageReaction', () => {
    it('removes own reaction', async () => {
      vi.mocked(repo.removeMessageReactionAsync).mockResolvedValue(undefined);

      const result = await tools.removeMessageReaction({
        message_id: 100, message_type: 'channel', reaction_type: 'like',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Reaction removed');
      expect(repo.removeMessageReactionAsync).toHaveBeenCalledWith(100, 'channel', 'like');
    });
  });
});
