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
import type { ApprovalTokenManager } from '../approval/index.js';

// =============================================================================
// Input Schemas
// =============================================================================

export const ListChannelsInput = z.strictObject({
  team_id: z.number().int().positive().describe('Team ID'),
});

export const GetChannelInput = z.strictObject({
  channel_id: z.number().int().positive().describe('Channel ID'),
});

export const CreateChannelInput = z.strictObject({
  team_id: z.number().int().positive().describe('Team ID'),
  name: z.string().min(1).describe('Channel name'),
  description: z.string().optional().describe('Channel description'),
});

export const UpdateChannelInput = z.strictObject({
  channel_id: z.number().int().positive().describe('Channel ID'),
  name: z.string().min(1).optional().describe('New name'),
  description: z.string().optional().describe('New description'),
});

export const PrepareDeleteChannelInput = z.strictObject({
  channel_id: z.number().int().positive().describe('Channel ID'),
});

export const ConfirmDeleteChannelInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_channel'),
});

export const ListTeamMembersInput = z.strictObject({
  team_id: z.number().int().positive().describe('Team ID'),
});

export const ListChannelMessagesInput = z.strictObject({
  channel_id: z.number().int().positive().describe('Channel ID from list_channels'),
  limit: z.number().int().min(1).max(50).optional().describe('Max messages to return (default 25, max 50)'),
});

export const GetChannelMessageInput = z.strictObject({
  message_id: z.number().int().positive().describe('Message ID from list_channel_messages'),
});

export const PrepareSendChannelMessageInput = z.strictObject({
  channel_id: z.number().int().positive().describe('Channel ID to send message to'),
  body: z.string().min(1).describe('Message body'),
  content_type: z.enum(['text', 'html']).optional().describe('Content type (default: html)'),
});

export const ConfirmSendChannelMessageInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_send_channel_message'),
});

export const PrepareReplyChannelMessageInput = z.strictObject({
  message_id: z.number().int().positive().describe('Message ID to reply to'),
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
  chat_id: z.number().int().positive().describe('Chat ID from list_chats'),
});

export const ListChatMessagesInput = z.strictObject({
  chat_id: z.number().int().positive().describe('Chat ID from list_chats'),
  limit: z.number().int().min(1).max(50).optional().describe('Max messages to return (default 25, max 50)'),
});

export const PrepareSendChatMessageInput = z.strictObject({
  chat_id: z.number().int().positive().describe('Chat ID to send message to'),
  body: z.string().min(1).describe('Message body'),
  content_type: z.enum(['text', 'html']).optional().describe('Content type (default: html)'),
});

export const ConfirmSendChatMessageInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_send_chat_message'),
});

export const ListChatMembersInput = z.strictObject({
  chat_id: z.number().int().positive().describe('Chat ID from list_chats'),
});

export const ListMessageReactionsInput = z.strictObject({
  message_id: z.number().describe('Numeric message ID'),
  message_type: z.enum(['channel', 'chat']).describe('Whether this is a channel message or chat message'),
});

export const PrepareAddMessageReactionInput = z.strictObject({
  message_id: z.number().describe('Numeric message ID'),
  message_type: z.enum(['channel', 'chat']).describe('Whether this is a channel message or chat message'),
  reaction_type: z.string().describe('Reaction emoji name (e.g., "like", "heart", "laugh", "surprised", "sad", "angry")'),
});

export const ConfirmAddMessageReactionInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_add_message_reaction'),
});

export const RemoveMessageReactionInput = z.strictObject({
  message_id: z.number().describe('Numeric message ID'),
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
  listTeamsAsync(): Promise<Array<{ id: number; name: string; description: string }>>;
  listChannelsAsync(teamId: number): Promise<Array<{ id: number; name: string; description: string; membershipType: string }>>;
  getChannelAsync(channelId: number): Promise<{ id: number; name: string; description: string; membershipType: string; webUrl: string }>;
  createChannelAsync(teamId: number, name: string, description?: string): Promise<number>;
  updateChannelAsync(channelId: number, updates: { name?: string; description?: string }): Promise<void>;
  deleteChannelAsync(channelId: number): Promise<void>;
  listTeamMembersAsync(teamId: number): Promise<Array<{ id: string; displayName: string; email: string; roles: string[] }>>;
  listChannelMessagesAsync(channelId: number, limit?: number): Promise<Array<{
    id: number; senderName: string; senderEmail: string; bodyPreview: string;
    bodyContent: string; contentType: string; createdDateTime: string;
  }>>;
  getChannelMessageAsync(messageId: number): Promise<{
    id: number; senderName: string; senderEmail: string; bodyContent: string;
    contentType: string; createdDateTime: string;
    replies: Array<{ id: number; senderName: string; senderEmail: string; bodyContent: string; contentType: string; createdDateTime: string }>;
  }>;
  sendChannelMessageAsync(channelId: number, body: string, contentType?: string): Promise<number>;
  replyToChannelMessageAsync(messageId: number, body: string, contentType?: string): Promise<number>;
  listChatsAsync(limit?: number): Promise<Array<{ id: number; topic: string; chatType: string; lastMessagePreview: string; createdDateTime: string }>>;
  getChatAsync(chatId: number): Promise<{ id: number; topic: string; chatType: string; createdDateTime: string; webUrl: string }>;
  listChatMessagesAsync(chatId: number, limit?: number): Promise<Array<{
    id: number; senderName: string; senderEmail: string; bodyPreview: string;
    bodyContent: string; contentType: string; createdDateTime: string;
  }>>;
  sendChatMessageAsync(chatId: number, body: string, contentType?: string): Promise<number>;
  listChatMembersAsync(chatId: number): Promise<Array<{ displayName: string; email: string; roles: string[] }>>;
  listMessageReactionsAsync(messageId: number, messageType: 'channel' | 'chat'): Promise<Array<{ reactionType: string; user: { displayName: string }; createdDateTime: string }>>;
  addMessageReactionAsync(messageId: number, messageType: 'channel' | 'chat', reactionType: string): Promise<void>;
  removeMessageReactionAsync(messageId: number, messageType: 'channel' | 'chat', reactionType: string): Promise<void>;
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
        text: JSON.stringify({ teams }, null, 2),
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
        text: JSON.stringify({ channels }, null, 2),
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
        text: JSON.stringify({ success: true, channel_id: channelId, message: 'Channel created' }, null, 2),
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

    await this.repo.deleteChannelAsync(result.token!.targetId);
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
    const messageId = await this.repo.sendChannelMessageAsync(result.token!.targetId, body, contentType);
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
    const replyId = await this.repo.replyToChannelMessageAsync(result.token!.targetId, body, contentType);
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
        text: JSON.stringify({ chats }, null, 2),
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
    const messageId = await this.repo.sendChatMessageAsync(result.token!.targetId, body, contentType);
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
    await this.repo.addMessageReactionAsync(result.token!.targetId, message_type, reaction_type);
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
