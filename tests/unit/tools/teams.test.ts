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
});
