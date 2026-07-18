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
      findChatsAsync: vi.fn(),
      resolveOrCreateChatAsync: vi.fn(),
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
        { id: 'tm_eng', name: 'Engineering', description: 'Eng team' },
        { id: 'tm_mktg', name: 'Marketing', description: 'Mktg team' },
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
        { id: 'cn_ch', name: 'General', description: 'Default', membershipType: 'standard' },
      ];
      vi.mocked(repo.listChannelsAsync).mockResolvedValue(mockChannels);

      const result = await tools.listChannels({ team_id: 'tm_eng' });

      expect(repo.listChannelsAsync).toHaveBeenCalledWith('tm_eng');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.channels).toEqual(mockChannels);
    });
  });

  describe('getChannel', () => {
    it('returns channel details', async () => {
      const mockChannel = {
        id: 'cn_ch', name: 'General', description: 'Default', membershipType: 'standard', webUrl: 'https://...',
      };
      vi.mocked(repo.getChannelAsync).mockResolvedValue(mockChannel);

      const result = await tools.getChannel({ channel_id: 'cn_ch' });

      expect(repo.getChannelAsync).toHaveBeenCalledWith('cn_ch');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.channel).toEqual(mockChannel);
    });
  });

  describe('createChannel', () => {
    it('creates a channel and returns the ID', async () => {
      vi.mocked(repo.createChannelAsync).mockResolvedValue('cn_new');

      const result = await tools.createChannel({ team_id: 'tm_eng', name: 'Dev', description: 'Dev channel' });

      expect(repo.createChannelAsync).toHaveBeenCalledWith('tm_eng', 'Dev', 'Dev channel');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.channel_id).toBe('cn_new');
      expect(parsed.message).toBe('Channel created');
    });

    it('creates a channel without description', async () => {
      vi.mocked(repo.createChannelAsync).mockResolvedValue('cn_new');

      await tools.createChannel({ team_id: 'tm_eng', name: 'Dev' });

      expect(repo.createChannelAsync).toHaveBeenCalledWith('tm_eng', 'Dev', undefined);
    });
  });

  describe('updateChannel', () => {
    it('updates a channel', async () => {
      vi.mocked(repo.updateChannelAsync).mockResolvedValue(undefined);

      const result = await tools.updateChannel({ channel_id: 'cn_ch', name: 'Renamed', description: 'New desc' });

      expect(repo.updateChannelAsync).toHaveBeenCalledWith('cn_ch', { name: 'Renamed', description: 'New desc' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Channel updated');
    });
  });

  describe('prepareDeleteChannel', () => {
    it('generates an approval token', () => {
      const result = tools.prepareDeleteChannel({ channel_id: 'cn_del' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.channel_id).toBe('cn_del');
      expect(parsed.action).toContain('confirm_delete_channel');
    });
  });

  describe('confirmDeleteChannel', () => {
    it('deletes the channel with a valid token', async () => {
      vi.mocked(repo.deleteChannelAsync).mockResolvedValue(undefined);

      // Generate a token first
      const prepareResult = tools.prepareDeleteChannel({ channel_id: 'cn_del' });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeleteChannel({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Channel deleted');
      expect(repo.deleteChannelAsync).toHaveBeenCalledWith('cn_del');
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmDeleteChannel({ approval_token: 'invalid-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.deleteChannelAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeleteChannel({ channel_id: 'cn_del' });
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

      const result = await tools.listTeamMembers({ team_id: 'tm_eng' });

      expect(repo.listTeamMembersAsync).toHaveBeenCalledWith('tm_eng');
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
          id: 'xm_msg1', senderName: 'Alice', senderEmail: 'alice@example.com',
          bodyPreview: 'Hello world', bodyContent: 'Hello world',
          contentType: 'text', createdDateTime: '2026-01-01T00:00:00Z',
        },
        {
          id: 'xm_msg2', senderName: 'Bob', senderEmail: 'bob@example.com',
          bodyPreview: 'Hi there', bodyContent: 'Hi there',
          contentType: 'text', createdDateTime: '2026-01-01T01:00:00Z',
        },
      ];
      vi.mocked(repo.listChannelMessagesAsync).mockResolvedValue(mockMessages);

      const result = await tools.listChannelMessages({ channel_id: 'cn_ch' });

      expect(repo.listChannelMessagesAsync).toHaveBeenCalledWith('cn_ch', undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.messages).toEqual(mockMessages);
    });

    it('passes limit parameter', async () => {
      vi.mocked(repo.listChannelMessagesAsync).mockResolvedValue([]);

      await tools.listChannelMessages({ channel_id: 'cn_ch', limit: 5 });

      expect(repo.listChannelMessagesAsync).toHaveBeenCalledWith('cn_ch', 5);
    });
  });

  describe('getChannelMessage', () => {
    it('returns message with replies', async () => {
      const mockMessage = {
        id: 'xm_msg1', senderName: 'Alice', senderEmail: 'alice@example.com',
        bodyContent: '<p>Hello</p>', contentType: 'html',
        createdDateTime: '2026-01-01T00:00:00Z',
        replies: [
          {
            id: 'xm_reply1', senderName: 'Bob', senderEmail: 'bob@example.com',
            bodyContent: '<p>Hi back</p>', contentType: 'html',
            createdDateTime: '2026-01-01T01:00:00Z',
          },
        ],
      };
      vi.mocked(repo.getChannelMessageAsync).mockResolvedValue(mockMessage);

      const result = await tools.getChannelMessage({ message_id: 'xm_msg1' });

      expect(repo.getChannelMessageAsync).toHaveBeenCalledWith('xm_msg1');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.message).toEqual(mockMessage);
      expect(parsed.message.replies).toHaveLength(1);
    });
  });

  describe('prepareSendChannelMessage', () => {
    it('generates an approval token with preview', () => {
      const result = tools.prepareSendChannelMessage({
        channel_id: 'cn_ch', body: 'Hello channel!', content_type: 'text',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.channel_id).toBe('cn_ch');
      expect(parsed.body_preview).toBe('Hello channel!');
      expect(parsed.content_type).toBe('text');
      expect(parsed.action).toContain('confirm_send_channel_message');
    });

    it('defaults content_type to html', () => {
      const result = tools.prepareSendChannelMessage({
        channel_id: 'cn_ch', body: '<p>Hello</p>',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.content_type).toBe('html');
    });
  });

  describe('confirmSendChannelMessage', () => {
    it('sends message with valid token', async () => {
      vi.mocked(repo.sendChannelMessageAsync).mockResolvedValue('xm_999');

      const prepareResult = tools.prepareSendChannelMessage({
        channel_id: 'cn_ch', body: 'Hello!', content_type: 'text',
      });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmSendChannelMessage({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message_id).toBe('xm_999');
      expect(parsed.message).toBe('Message sent');
      expect(repo.sendChannelMessageAsync).toHaveBeenCalledWith('cn_ch', 'Hello!', 'text');
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmSendChannelMessage({ approval_token: 'bad-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.sendChannelMessageAsync).mockResolvedValue('xm_999');

      const prepareResult = tools.prepareSendChannelMessage({
        channel_id: 'cn_ch', body: 'Hello!',
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
        message_id: 'xm_msg1', body: 'Nice post!', content_type: 'text',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.message_id).toBe('xm_msg1');
      expect(parsed.body_preview).toBe('Nice post!');
      expect(parsed.content_type).toBe('text');
      expect(parsed.action).toContain('confirm_reply_channel_message');
    });

    it('defaults content_type to html', () => {
      const result = tools.prepareReplyChannelMessage({
        message_id: 'xm_msg1', body: '<p>Reply</p>',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.content_type).toBe('html');
    });
  });

  describe('confirmReplyChannelMessage', () => {
    it('sends reply with valid token', async () => {
      vi.mocked(repo.replyToChannelMessageAsync).mockResolvedValue('xm_888');

      const prepareResult = tools.prepareReplyChannelMessage({
        message_id: 'xm_msg1', body: 'Great idea!', content_type: 'text',
      });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmReplyChannelMessage({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.reply_id).toBe('xm_888');
      expect(parsed.message).toBe('Reply sent');
      expect(repo.replyToChannelMessageAsync).toHaveBeenCalledWith('xm_msg1', 'Great idea!', 'text');
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmReplyChannelMessage({ approval_token: 'bad-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.replyToChannelMessageAsync).mockResolvedValue('xm_888');

      const prepareResult = tools.prepareReplyChannelMessage({
        message_id: 'xm_msg1', body: 'Reply!',
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
        { id: 'ch_chat1', topic: 'Project Chat', chatType: 'group', lastMessagePreview: 'Hello', createdDateTime: '2026-01-01T00:00:00Z' },
        { id: 'ch_chat2', topic: '', chatType: 'oneOnOne', lastMessagePreview: 'Hi there', createdDateTime: '2026-01-02T00:00:00Z' },
      ];
      vi.mocked(repo.listChatsAsync).mockResolvedValue(mockChats);

      const result = await tools.listChats({});

      expect(repo.listChatsAsync).toHaveBeenCalledWith(undefined, false);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.chats).toEqual(mockChats);
    });

    it('passes limit parameter', async () => {
      vi.mocked(repo.listChatsAsync).mockResolvedValue([]);

      await tools.listChats({ limit: 10 });

      expect(repo.listChatsAsync).toHaveBeenCalledWith(10, false);
    });

    it('passes expand_members parameter', async () => {
      vi.mocked(repo.listChatsAsync).mockResolvedValue([]);

      await tools.listChats({ expand_members: true });

      expect(repo.listChatsAsync).toHaveBeenCalledWith(undefined, true);
    });
  });

  describe('findChat', () => {
    it('returns matching chats from the repository', async () => {
      const mockChats = [{
        id: 'ch_chat1',
        topic: '',
        chatType: 'oneOnOne',
        lastMessagePreview: '',
        createdDateTime: '2026-01-01T00:00:00Z',
        members: [
          { displayName: 'Alice', email: 'alice@example.com', userId: 'u1', roles: ['owner'] },
        ],
      }];
      vi.mocked(repo.findChatsAsync).mockResolvedValue(mockChats);

      const result = await tools.findChat({ participants: ['alice@example.com'] });

      expect(repo.findChatsAsync).toHaveBeenCalledWith({
        participants: ['alice@example.com'],
      });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.chats).toEqual(mockChats);
    });

    it('passes chat_type filter', async () => {
      vi.mocked(repo.findChatsAsync).mockResolvedValue([]);

      await tools.findChat({ participants: ['Alice'], chat_type: 'group' });

      expect(repo.findChatsAsync).toHaveBeenCalledWith({
        participants: ['Alice'],
        chatType: 'group',
      });
    });
  });

  describe('getChat', () => {
    it('returns chat details', async () => {
      const mockChat = {
        id: 'ch_chat1', topic: 'Project Chat', chatType: 'group', createdDateTime: '2026-01-01T00:00:00Z', webUrl: 'https://teams.microsoft.com/...',
      };
      vi.mocked(repo.getChatAsync).mockResolvedValue(mockChat);

      const result = await tools.getChat({ chat_id: 'ch_chat1' });

      expect(repo.getChatAsync).toHaveBeenCalledWith('ch_chat1');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.chat).toEqual(mockChat);
    });
  });

  describe('listChatMessages', () => {
    it('returns messages from the repository', async () => {
      const mockMessages = [
        {
          id: 'cm_msg1', senderName: 'Alice', senderEmail: 'alice@example.com',
          bodyPreview: 'Hello', bodyContent: 'Hello',
          contentType: 'text', createdDateTime: '2026-01-01T00:00:00Z',
        },
      ];
      vi.mocked(repo.listChatMessagesAsync).mockResolvedValue(mockMessages);

      const result = await tools.listChatMessages({ chat_id: 'ch_chat1' });

      expect(repo.listChatMessagesAsync).toHaveBeenCalledWith('ch_chat1', undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.messages).toEqual(mockMessages);
    });

    it('passes limit parameter', async () => {
      vi.mocked(repo.listChatMessagesAsync).mockResolvedValue([]);

      await tools.listChatMessages({ chat_id: 'ch_chat1', limit: 5 });

      expect(repo.listChatMessagesAsync).toHaveBeenCalledWith('ch_chat1', 5);
    });
  });

  describe('prepareSendChatMessage', () => {
    it('generates an approval token with preview', async () => {
      const result = await tools.prepareSendChatMessage({
        chat_id: 'ch_chat1', body: 'Hello chat!', content_type: 'text',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.chat_id).toBe('ch_chat1');
      expect(parsed.body_preview).toBe('Hello chat!');
      expect(parsed.content_type).toBe('text');
      expect(parsed.action).toContain('confirm_send_chat_message');
    });

    it('defaults content_type to html', async () => {
      const result = await tools.prepareSendChatMessage({
        chat_id: 'ch_chat1', body: '<p>Hello</p>',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.content_type).toBe('html');
    });

    it('resolves chat from to participants without prior chat_id', async () => {
      vi.mocked(repo.resolveOrCreateChatAsync).mockResolvedValue({ chatId: 'ch_resolved' });

      const result = await tools.prepareSendChatMessage({
        to: ['alice@example.com'], body: 'Hi Alice',
      });

      expect(repo.resolveOrCreateChatAsync).toHaveBeenCalledWith(['alice@example.com']);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.chat_id).toBe('ch_resolved');
      expect(parsed.approval_token).toBeDefined();
    });

    it('returns candidates when participant resolution is ambiguous', async () => {
      vi.mocked(repo.resolveOrCreateChatAsync).mockResolvedValue({
        error: 'Multiple chats match the given participants. Pass chat_id to disambiguate.',
        chats: [
          { id: 'ch_a', topic: 'A', chatType: 'group', members: [{ displayName: 'Alice', email: 'alice@example.com' }] },
          { id: 'ch_b', topic: 'B', chatType: 'group', members: [{ displayName: 'Alice', email: 'alice@example.com' }] },
        ],
      });

      const result = await tools.prepareSendChatMessage({
        to: ['alice@example.com', 'bob@example.com'], body: 'Hi',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.chats).toHaveLength(2);
      expect(parsed.approval_token).toBeUndefined();
    });
  });

  describe('confirmSendChatMessage', () => {
    it('sends message with valid token', async () => {
      vi.mocked(repo.sendChatMessageAsync).mockResolvedValue('cm_999');

      const prepareResult = await tools.prepareSendChatMessage({
        chat_id: 'ch_chat1', body: 'Hello!', content_type: 'text',
      });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmSendChatMessage({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message_id).toBe('cm_999');
      expect(parsed.message).toBe('Message sent');
      expect(repo.sendChatMessageAsync).toHaveBeenCalledWith('ch_chat1', 'Hello!', 'text');
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmSendChatMessage({ approval_token: 'bad-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.sendChatMessageAsync).mockResolvedValue('cm_999');

      const prepareResult = await tools.prepareSendChatMessage({
        chat_id: 'ch_chat1', body: 'Hello!',
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

      const result = await tools.listChatMembers({ chat_id: 'ch_chat1' });

      expect(repo.listChatMembersAsync).toHaveBeenCalledWith('ch_chat1');
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

      const result = await tools.listMessageReactions({ message_id: 'xm_msg1', message_type: 'channel' });

      expect(repo.listMessageReactionsAsync).toHaveBeenCalledWith('xm_msg1', 'channel');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.reactions).toEqual(mockReactions);
    });

    it('returns reactions for a chat message', async () => {
      const mockReactions = [
        { reactionType: 'laugh', user: { displayName: 'Charlie' }, createdDateTime: '2026-01-02T00:00:00Z' },
      ];
      vi.mocked(repo.listMessageReactionsAsync).mockResolvedValue(mockReactions);

      const result = await tools.listMessageReactions({ message_id: 'cm_msg2', message_type: 'chat' });

      expect(repo.listMessageReactionsAsync).toHaveBeenCalledWith('cm_msg2', 'chat');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.reactions).toEqual(mockReactions);
    });
  });

  describe('prepareAddMessageReaction', () => {
    it('generates an approval token', () => {
      const result = tools.prepareAddMessageReaction({
        message_id: 'xm_msg1', message_type: 'channel', reaction_type: 'like',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.message_id).toBe('xm_msg1');
      expect(parsed.message_type).toBe('channel');
      expect(parsed.reaction_type).toBe('like');
      expect(parsed.action).toContain('confirm_add_message_reaction');
    });
  });

  describe('confirmAddMessageReaction', () => {
    it('adds reaction with valid token', async () => {
      vi.mocked(repo.addMessageReactionAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareAddMessageReaction({
        message_id: 'xm_msg1', message_type: 'channel', reaction_type: 'like',
      });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmAddMessageReaction({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Reaction added');
      expect(repo.addMessageReactionAsync).toHaveBeenCalledWith('xm_msg1', 'channel', 'like');
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
        message_id: 'xm_msg1', message_type: 'channel', reaction_type: 'like',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Reaction removed');
      expect(repo.removeMessageReactionAsync).toHaveBeenCalledWith('xm_msg1', 'channel', 'like');
    });
  });
});
