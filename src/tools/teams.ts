/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Microsoft Teams MCP tools.
 *
 * Provides tools for managing Teams, channels, and members with a two-phase
 * approval pattern for destructive delete operations.
 */

import { z } from 'zod';
import { Id } from '../ids/schema.js';
import { nextActionFor } from '../ids/next-action.js';
import type { ApprovalTokenManager } from '../approval/index.js';
import { defineTool } from '../registry/define-tool.js';
import { approvalTokenLink } from '../registry/elicit-links.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolContext, ToolDefinition } from '../registry/types.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    teams: TeamsTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const NoInput = z.strictObject({});

export const ListChannelsInput = z.strictObject({
  team_id: Id.team,
});

export const GetChannelInput = z.strictObject({
  channel_id: Id.channel,
});

export const CreateChannelInput = z.strictObject({
  team_id: Id.team,
  name: z.string().min(1).describe('Channel name'),
  description: z.string().optional().describe('Channel description'),
});

export const UpdateChannelInput = z.strictObject({
  channel_id: Id.channel,
  name: z.string().min(1).optional().describe('New name'),
  description: z.string().optional().describe('New description'),
});

export const PrepareDeleteChannelInput = z.strictObject({
  channel_id: Id.channel,
});

export const ConfirmDeleteChannelInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_channel'),
});

export const ListTeamMembersInput = z.strictObject({
  team_id: Id.team,
});

export const ListChannelMessagesInput = z.strictObject({
  channel_id: Id.channel,
  limit: z.number().int().min(1).max(50).optional().describe('Max messages to return (default 25, max 50)'),
});

export const GetChannelMessageInput = z.strictObject({
  message_id: Id.channelMessage,
});

export const PrepareSendChannelMessageInput = z.strictObject({
  channel_id: Id.channel.describe('Channel ID (cn_ token) to send the message to.'),
  body: z.string().min(1).describe('Message body'),
  content_type: z.enum(['text', 'html']).optional().describe('Content type (default: html)'),
});

export const ConfirmSendChannelMessageInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_send_channel_message'),
});

export const PrepareReplyChannelMessageInput = z.strictObject({
  message_id: Id.channelMessage.describe('Channel message ID (xm_ token) to reply to.'),
  body: z.string().min(1).describe('Reply body'),
  content_type: z.enum(['text', 'html']).optional().describe('Content type (default: html)'),
});

export const ConfirmReplyChannelMessageInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_reply_channel_message'),
});

export const ListChatsInput = z.strictObject({
  limit: z.number().int().min(1).max(50).optional().describe('Max chats to return (default 25, max 50)'),
});

export const GetChatInput = z.strictObject({
  chat_id: Id.chat,
});

export const ListChatMessagesInput = z.strictObject({
  chat_id: Id.chat,
  limit: z.number().int().min(1).max(50).optional().describe('Max messages to return (default 25, max 50)'),
});

export const PrepareSendChatMessageInput = z.strictObject({
  chat_id: Id.chat.describe('Chat ID (ch_ token) to send the message to.'),
  body: z.string().min(1).describe('Message body'),
  content_type: z.enum(['text', 'html']).optional().describe('Content type (default: html)'),
});

export const ConfirmSendChatMessageInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_send_chat_message'),
});

export const ListChatMembersInput = z.strictObject({
  chat_id: Id.chat,
});

export const ListMessageReactionsInput = z.strictObject({
  message_id: z.string().trim().min(1).describe('Message ID — a cm_ (chat) or xm_ (channel) token'),
  message_type: z.enum(['channel', 'chat']).describe('Whether this is a channel message or chat message'),
});

export const PrepareAddMessageReactionInput = z.strictObject({
  message_id: z.string().trim().min(1).describe('Message ID — a cm_ (chat) or xm_ (channel) token'),
  message_type: z.enum(['channel', 'chat']).describe('Whether this is a channel message or chat message'),
  reaction_type: z.string().describe('Reaction emoji name (e.g., "like", "heart", "laugh", "surprised", "sad", "angry")'),
});

export const ConfirmAddMessageReactionInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_add_message_reaction'),
});

export const RemoveMessageReactionInput = z.strictObject({
  message_id: z.string().trim().min(1).describe('Message ID — a cm_ (chat) or xm_ (channel) token'),
  message_type: z.enum(['channel', 'chat']).describe('Whether this is a channel message or chat message'),
  reaction_type: z.string().describe('Reaction emoji name to remove'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListChannelsParams = z.infer<typeof ListChannelsInput>;
export type GetChannelParams = z.infer<typeof GetChannelInput>;
export type CreateChannelParams = z.infer<typeof CreateChannelInput>;
export type UpdateChannelParams = z.infer<typeof UpdateChannelInput>;
export type PrepareDeleteChannelParams = z.infer<typeof PrepareDeleteChannelInput>;
export type ConfirmDeleteChannelParams = z.infer<typeof ConfirmDeleteChannelInput>;
export type ListTeamMembersParams = z.infer<typeof ListTeamMembersInput>;
export type ListChannelMessagesParams = z.infer<typeof ListChannelMessagesInput>;
export type GetChannelMessageParams = z.infer<typeof GetChannelMessageInput>;
export type PrepareSendChannelMessageParams = z.infer<typeof PrepareSendChannelMessageInput>;
export type ConfirmSendChannelMessageParams = z.infer<typeof ConfirmSendChannelMessageInput>;
export type PrepareReplyChannelMessageParams = z.infer<typeof PrepareReplyChannelMessageInput>;
export type ConfirmReplyChannelMessageParams = z.infer<typeof ConfirmReplyChannelMessageInput>;
export type ListChatsParams = z.infer<typeof ListChatsInput>;
export type GetChatParams = z.infer<typeof GetChatInput>;
export type ListChatMessagesParams = z.infer<typeof ListChatMessagesInput>;
export type PrepareSendChatMessageParams = z.infer<typeof PrepareSendChatMessageInput>;
export type ConfirmSendChatMessageParams = z.infer<typeof ConfirmSendChatMessageInput>;
export type ListChatMembersParams = z.infer<typeof ListChatMembersInput>;
export type ListMessageReactionsParams = z.infer<typeof ListMessageReactionsInput>;
export type PrepareAddMessageReactionParams = z.infer<typeof PrepareAddMessageReactionInput>;
export type ConfirmAddMessageReactionParams = z.infer<typeof ConfirmAddMessageReactionInput>;
export type RemoveMessageReactionParams = z.infer<typeof RemoveMessageReactionInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface ITeamsRepository {
  listTeamsAsync(): Promise<Array<{ id: string; name: string; description: string }>>;
  listChannelsAsync(teamId: string): Promise<Array<{ id: string; name: string; description: string; membershipType: string }>>;
  getChannelAsync(channelId: string): Promise<{ id: string; name: string; description: string; membershipType: string; webUrl: string }>;
  createChannelAsync(teamId: string, name: string, description?: string): Promise<string>;
  updateChannelAsync(channelId: string, updates: { name?: string; description?: string }): Promise<void>;
  deleteChannelAsync(channelId: string): Promise<void>;
  listTeamMembersAsync(teamId: string): Promise<Array<{ id: string; displayName: string; email: string; roles: string[] }>>;
  listChannelMessagesAsync(channelId: string, limit?: number): Promise<Array<{
    id: string; senderName: string; senderEmail: string; bodyPreview: string;
    bodyContent: string; contentType: string; createdDateTime: string;
  }>>;
  getChannelMessageAsync(messageId: string): Promise<{
    id: string; senderName: string; senderEmail: string; bodyContent: string;
    contentType: string; createdDateTime: string;
    replies: Array<{ id: string; senderName: string; senderEmail: string; bodyContent: string; contentType: string; createdDateTime: string }>;
  }>;
  sendChannelMessageAsync(channelId: string, body: string, contentType?: string): Promise<string>;
  replyToChannelMessageAsync(messageId: string, body: string, contentType?: string): Promise<string>;
  listChatsAsync(limit?: number): Promise<Array<{ id: string; topic: string; chatType: string; lastMessagePreview: string; createdDateTime: string }>>;
  getChatAsync(chatId: string): Promise<{ id: string; topic: string; chatType: string; createdDateTime: string; webUrl: string }>;
  listChatMessagesAsync(chatId: string, limit?: number): Promise<Array<{
    id: string; senderName: string; senderEmail: string; bodyPreview: string;
    bodyContent: string; contentType: string; createdDateTime: string;
  }>>;
  sendChatMessageAsync(chatId: string, body: string, contentType?: string): Promise<string>;
  listChatMembersAsync(chatId: string): Promise<Array<{ displayName: string; email: string; roles: string[] }>>;
  listMessageReactionsAsync(messageId: string, messageType: 'channel' | 'chat'): Promise<Array<{ reactionType: string; user: { displayName: string }; createdDateTime: string }>>;
  addMessageReactionAsync(messageId: string, messageType: 'channel' | 'chat', reactionType: string): Promise<void>;
  removeMessageReactionAsync(messageId: string, messageType: 'channel' | 'chat', reactionType: string): Promise<void>;
}

// =============================================================================
// Teams Tools
// =============================================================================

/**
 * Microsoft Teams tools with two-phase approval for delete operations.
 */
export class TeamsTools {
  constructor(
    private readonly repo: ITeamsRepository,
    private readonly tokenManager: ApprovalTokenManager,
  ) {}

  async listTeams(): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const teams = await this.repo.listTeamsAsync();
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ teams, next: nextActionFor('team') ?? undefined }, null, 2),
      }],
    };
  }

  async listChannels(params: ListChannelsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const channels = await this.repo.listChannelsAsync(params.team_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ channels, next: nextActionFor('channel') ?? undefined }, null, 2),
      }],
    };
  }

  async getChannel(params: GetChannelParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const channel = await this.repo.getChannelAsync(params.channel_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ channel }, null, 2),
      }],
    };
  }

  async createChannel(params: CreateChannelParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const channelId = await this.repo.createChannelAsync(params.team_id, params.name, params.description);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, channel_id: channelId, message: 'Channel created', next: nextActionFor('channel') ?? undefined }, null, 2),
      }],
    };
  }

  async updateChannel(params: UpdateChannelParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const updates: { name?: string; description?: string } = {};
    if (params.name != null) updates.name = params.name;
    if (params.description != null) updates.description = params.description;
    await this.repo.updateChannelAsync(params.channel_id, updates);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Channel updated' }, null, 2),
      }],
    };
  }

  prepareDeleteChannel(params: PrepareDeleteChannelParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'delete_channel',
      targetType: 'channel',
      targetId: params.channel_id,
      targetHash: String(params.channel_id),
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          channel_id: params.channel_id,
          action: `To confirm deleting channel ${params.channel_id}, call confirm_delete_channel with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmDeleteChannel(params: ConfirmDeleteChannelParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    // Look up the token to get the targetId, then consume it
    const token = this.tokenManager.lookupToken(params.approval_token);
    if (token == null) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({
            success: false,
            error: 'Token not found or already used',
          }, null, 2),
        }],
      };
    }

    const result = this.tokenManager.consumeToken(params.approval_token, 'delete_channel', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_channel again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_channel',
        TARGET_MISMATCH: 'Token was generated for a different channel',
        ALREADY_CONSUMED: 'Token has already been used',
      };
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({
            success: false,
            error: errorMessages[result.error ?? ''] ?? 'Invalid token',
          }, null, 2),
        }],
      };
    }

    await this.repo.deleteChannelAsync((result.token!.targetId));
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Channel deleted' }, null, 2),
      }],
    };
  }

  async listTeamMembers(params: ListTeamMembersParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const members = await this.repo.listTeamMembersAsync(params.team_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ members }, null, 2),
      }],
    };
  }

  // ===========================================================================
  // Channel Messages
  // ===========================================================================

  async listChannelMessages(params: ListChannelMessagesParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const messages = await this.repo.listChannelMessagesAsync(params.channel_id, params.limit);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ messages }, null, 2),
      }],
    };
  }

  async getChannelMessage(params: GetChannelMessageParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const message = await this.repo.getChannelMessageAsync(params.message_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ message }, null, 2),
      }],
    };
  }

  prepareSendChannelMessage(params: PrepareSendChannelMessageParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const contentType = params.content_type ?? 'html';
    const token = this.tokenManager.generateToken({
      operation: 'send_channel_message',
      targetType: 'channel_message',
      targetId: params.channel_id,
      targetHash: String(params.channel_id),
      metadata: { body: params.body, contentType },
    });
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          channel_id: params.channel_id,
          body_preview: params.body.substring(0, 200),
          content_type: contentType,
          action: 'Call confirm_send_channel_message with the approval_token to send.',
        }, null, 2),
      }],
    };
  }

  async confirmSendChannelMessage(params: ConfirmSendChannelMessageParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const token = this.tokenManager.lookupToken(params.approval_token);
    if (token == null) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ success: false, error: 'Token not found or already used' }, null, 2),
        }],
      };
    }
    const result = this.tokenManager.consumeToken(params.approval_token, 'send_channel_message', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_send_channel_message again.',
        OPERATION_MISMATCH: 'Token was not generated for send_channel_message',
        TARGET_MISMATCH: 'Token was generated for a different channel',
        ALREADY_CONSUMED: 'Token has already been used',
      };
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ success: false, error: errorMessages[result.error ?? ''] ?? 'Invalid token' }, null, 2),
        }],
      };
    }
    const { body, contentType } = result.token!.metadata as { body: string; contentType: string };
    const messageId = await this.repo.sendChannelMessageAsync((result.token!.targetId), body, contentType);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message_id: messageId, message: 'Message sent' }, null, 2),
      }],
    };
  }

  prepareReplyChannelMessage(params: PrepareReplyChannelMessageParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const contentType = params.content_type ?? 'html';
    const token = this.tokenManager.generateToken({
      operation: 'reply_channel_message',
      targetType: 'channel_message',
      targetId: params.message_id,
      targetHash: String(params.message_id),
      metadata: { body: params.body, contentType },
    });
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          message_id: params.message_id,
          body_preview: params.body.substring(0, 200),
          content_type: contentType,
          action: 'Call confirm_reply_channel_message with the approval_token to send the reply.',
        }, null, 2),
      }],
    };
  }

  async confirmReplyChannelMessage(params: ConfirmReplyChannelMessageParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const token = this.tokenManager.lookupToken(params.approval_token);
    if (token == null) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ success: false, error: 'Token not found or already used' }, null, 2),
        }],
      };
    }
    const result = this.tokenManager.consumeToken(params.approval_token, 'reply_channel_message', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_reply_channel_message again.',
        OPERATION_MISMATCH: 'Token was not generated for reply_channel_message',
        TARGET_MISMATCH: 'Token was generated for a different message',
        ALREADY_CONSUMED: 'Token has already been used',
      };
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ success: false, error: errorMessages[result.error ?? ''] ?? 'Invalid token' }, null, 2),
        }],
      };
    }
    const { body, contentType } = result.token!.metadata as { body: string; contentType: string };
    const replyId = await this.repo.replyToChannelMessageAsync((result.token!.targetId), body, contentType);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, reply_id: replyId, message: 'Reply sent' }, null, 2),
      }],
    };
  }

  // ===========================================================================
  // Chats
  // ===========================================================================

  async listChats(params: ListChatsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const chats = await this.repo.listChatsAsync(params.limit);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ chats, next: nextActionFor('chat') ?? undefined }, null, 2),
      }],
    };
  }

  async getChat(params: GetChatParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const chat = await this.repo.getChatAsync(params.chat_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ chat }, null, 2),
      }],
    };
  }

  async listChatMessages(params: ListChatMessagesParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const messages = await this.repo.listChatMessagesAsync(params.chat_id, params.limit);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ messages }, null, 2),
      }],
    };
  }

  prepareSendChatMessage(params: PrepareSendChatMessageParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const contentType = params.content_type ?? 'html';
    const token = this.tokenManager.generateToken({
      operation: 'send_chat_message',
      targetType: 'chat_message',
      targetId: params.chat_id,
      targetHash: String(params.chat_id),
      metadata: { body: params.body, contentType },
    });
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          chat_id: params.chat_id,
          body_preview: params.body.substring(0, 200),
          content_type: contentType,
          action: 'Call confirm_send_chat_message with the approval_token to send.',
        }, null, 2),
      }],
    };
  }

  async confirmSendChatMessage(params: ConfirmSendChatMessageParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const token = this.tokenManager.lookupToken(params.approval_token);
    if (token == null) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ success: false, error: 'Token not found or already used' }, null, 2),
        }],
      };
    }
    const result = this.tokenManager.consumeToken(params.approval_token, 'send_chat_message', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_send_chat_message again.',
        OPERATION_MISMATCH: 'Token was not generated for send_chat_message',
        TARGET_MISMATCH: 'Token was generated for a different chat',
        ALREADY_CONSUMED: 'Token has already been used',
      };
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ success: false, error: errorMessages[result.error ?? ''] ?? 'Invalid token' }, null, 2),
        }],
      };
    }
    const { body, contentType } = result.token!.metadata as { body: string; contentType: string };
    const messageId = await this.repo.sendChatMessageAsync((result.token!.targetId), body, contentType);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message_id: messageId, message: 'Message sent' }, null, 2),
      }],
    };
  }

  async listChatMembers(params: ListChatMembersParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const members = await this.repo.listChatMembersAsync(params.chat_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ members }, null, 2),
      }],
    };
  }

  // ===========================================================================
  // Message Reactions
  // ===========================================================================

  async listMessageReactions(params: ListMessageReactionsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const reactions = await this.repo.listMessageReactionsAsync(params.message_id, params.message_type);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ reactions }, null, 2),
      }],
    };
  }

  prepareAddMessageReaction(params: PrepareAddMessageReactionParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'add_message_reaction',
      targetType: 'message_reaction',
      targetId: params.message_id,
      targetHash: String(params.message_id),
      metadata: { reaction_type: params.reaction_type, message_type: params.message_type },
    });
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          message_id: params.message_id,
          message_type: params.message_type,
          reaction_type: params.reaction_type,
          action: 'Call confirm_add_message_reaction with the approval_token to add the reaction.',
        }, null, 2),
      }],
    };
  }

  async confirmAddMessageReaction(params: ConfirmAddMessageReactionParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const token = this.tokenManager.lookupToken(params.approval_token);
    if (token == null) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ success: false, error: 'Token not found or already used' }, null, 2),
        }],
      };
    }
    const result = this.tokenManager.consumeToken(params.approval_token, 'add_message_reaction', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_add_message_reaction again.',
        OPERATION_MISMATCH: 'Token was not generated for add_message_reaction',
        TARGET_MISMATCH: 'Token was generated for a different message',
        ALREADY_CONSUMED: 'Token has already been used',
      };
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ success: false, error: errorMessages[result.error ?? ''] ?? 'Invalid token' }, null, 2),
        }],
      };
    }
    const { reaction_type, message_type } = result.token!.metadata as { reaction_type: string; message_type: 'channel' | 'chat' };
    await this.repo.addMessageReactionAsync((result.token!.targetId), message_type, reaction_type);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Reaction added' }, null, 2),
      }],
    };
  }

  async removeMessageReaction(params: RemoveMessageReactionParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    await this.repo.removeMessageReactionAsync(params.message_id, params.message_type, params.reaction_type);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Reaction removed' }, null, 2),
      }],
    };
  }
}

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2)
// =============================================================================

/**
 * Registry tool definitions for the teams domain.
 */
export function teamsToolDefinitions(): ToolDefinition[] {
  const tools = (ctx: ToolContext): TeamsTools => requireGraphToolset(ctx, 'teams');

  return [
    defineTool({
      name: 'list_teams',
      description: 'List all Microsoft Teams the user has joined (Graph API)',
      input: NoInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx) => tools(ctx).listTeams(),
    }),
    defineTool({
      name: 'list_channels',
      description: 'List all channels in a team (Graph API)',
      input: ListChannelsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listChannels(params),
    }),
    defineTool({
      name: 'get_channel',
      description: 'Get details for a specific channel (Graph API)',
      input: GetChannelInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getChannel(params),
    }),
    defineTool({
      name: 'create_channel',
      description: 'Create a new channel in a team (Graph API)',
      input: CreateChannelInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).createChannel(params),
    }),
    defineTool({
      name: 'update_channel',
      description: 'Update a channel name or description (Graph API)',
      input: UpdateChannelInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).updateChannel(params),
    }),
    defineTool({
      name: 'prepare_delete_channel',
      description: 'Prepare to delete a channel. Returns an approval token. Call confirm_delete_channel to execute. (Graph API)',
      input: PrepareDeleteChannelInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareDeleteChannel(params),
      onElicit: approvalTokenLink('confirm_delete_channel'),
    }),
    defineTool({
      name: 'confirm_delete_channel',
      description: 'Confirm channel deletion with approval token (Graph API)',
      input: ConfirmDeleteChannelInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmDeleteChannel(params),
    }),
    defineTool({
      name: 'list_team_members',
      description: 'List all members of a team (Graph API)',
      input: ListTeamMembersInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listTeamMembers(params),
    }),
    defineTool({
      name: 'list_channel_messages',
      description: 'List recent messages in a channel (Graph API)',
      input: ListChannelMessagesInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listChannelMessages(params),
    }),
    defineTool({
      name: 'get_channel_message',
      description: 'Get a specific channel message with its replies (Graph API)',
      input: GetChannelMessageInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getChannelMessage(params),
    }),
    defineTool({
      name: 'prepare_send_channel_message',
      description: 'Prepare to send a message to a channel. Returns an approval token. Call confirm_send_channel_message to execute. (Graph API)',
      input: PrepareSendChannelMessageInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareSendChannelMessage(params),
      onElicit: approvalTokenLink('confirm_send_channel_message'),
    }),
    defineTool({
      name: 'confirm_send_channel_message',
      description: 'Confirm sending a channel message with approval token (Graph API)',
      input: ConfirmSendChannelMessageInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmSendChannelMessage(params),
    }),
    defineTool({
      name: 'prepare_reply_channel_message',
      description: 'Prepare to reply to a channel message. Returns an approval token. Call confirm_reply_channel_message to execute. (Graph API)',
      input: PrepareReplyChannelMessageInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareReplyChannelMessage(params),
      onElicit: approvalTokenLink('confirm_reply_channel_message'),
    }),
    defineTool({
      name: 'confirm_reply_channel_message',
      description: 'Confirm replying to a channel message with approval token (Graph API)',
      input: ConfirmReplyChannelMessageInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmReplyChannelMessage(params),
    }),
    defineTool({
      name: 'list_chats',
      description: 'List recent 1:1 and group chats (Graph API)',
      input: ListChatsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listChats(params),
    }),
    defineTool({
      name: 'get_chat',
      description: 'Get details of a specific chat (Graph API)',
      input: GetChatInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getChat(params),
    }),
    defineTool({
      name: 'list_chat_messages',
      description: 'List recent messages in a chat (Graph API)',
      input: ListChatMessagesInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listChatMessages(params),
    }),
    defineTool({
      name: 'prepare_send_chat_message',
      description: 'Prepare to send a message in a chat. Returns an approval token. Call confirm_send_chat_message to execute. (Graph API)',
      input: PrepareSendChatMessageInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareSendChatMessage(params),
      onElicit: approvalTokenLink('confirm_send_chat_message'),
    }),
    defineTool({
      name: 'confirm_send_chat_message',
      description: 'Confirm sending a chat message with approval token (Graph API)',
      input: ConfirmSendChatMessageInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmSendChatMessage(params),
    }),
    defineTool({
      name: 'list_chat_members',
      description: 'List members of a chat (Graph API)',
      input: ListChatMembersInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listChatMembers(params),
    }),
    defineTool({
      name: 'list_message_reactions',
      description: 'List reactions on a channel or chat message (Graph API)',
      input: ListMessageReactionsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listMessageReactions(params),
    }),
    defineTool({
      name: 'prepare_add_message_reaction',
      description: 'Prepare to add a reaction to a message. Returns an approval token. Call confirm_add_message_reaction to execute. (Graph API)',
      input: PrepareAddMessageReactionInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareAddMessageReaction(params),
      onElicit: approvalTokenLink('confirm_add_message_reaction'),
    }),
    defineTool({
      name: 'confirm_add_message_reaction',
      description: 'Confirm adding a reaction to a message with approval token (Graph API)',
      input: ConfirmAddMessageReactionInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmAddMessageReaction(params),
    }),
    defineTool({
      name: 'remove_message_reaction',
      description: 'Remove your own reaction from a channel or chat message (Graph API)',
      input: RemoveMessageReactionInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['teams'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).removeMessageReaction(params),
    }),
  ];
}
