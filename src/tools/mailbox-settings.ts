/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Graph-backend mailbox settings tools (v3 registry-driven architecture, U2).
 * Holds the automatic-replies (out-of-office) and mailbox-settings logic that
 * previously lived inline in the `handleGraphToolCall` switch.
 */

import { z } from 'zod';
import type { GraphRepository } from '../graph/repository.js';
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolContext, ToolDefinition, ToolResult } from '../registry/types.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    mailboxSettings: GraphMailboxSettingsTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const GetAutomaticRepliesInput = z.strictObject({});

export const SetAutomaticRepliesInput = z.strictObject({
  status: z.enum(['disabled', 'alwaysEnabled', 'scheduled']).describe('OOF status'),
  external_audience: z.enum(['none', 'contactsOnly', 'all']).optional().describe('Who sees external reply'),
  internal_reply_message: z.string().optional().describe('Reply for internal senders (HTML)'),
  external_reply_message: z.string().optional().describe('Reply for external senders (HTML)'),
  scheduled_start: z.string().optional().describe('Schedule start (ISO 8601)'),
  scheduled_end: z.string().optional().describe('Schedule end (ISO 8601)'),
});

export const GetMailboxSettingsInput = z.strictObject({});

export const UpdateMailboxSettingsInput = z.strictObject({
  language: z.string().optional().describe('Locale code (e.g. en-US)'),
  time_zone: z.string().optional().describe('Time zone (e.g. America/New_York)'),
  date_format: z.string().optional().describe('Date format string'),
  time_format: z.string().optional().describe('Time format string'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type SetAutomaticRepliesParams = z.infer<typeof SetAutomaticRepliesInput>;
export type UpdateMailboxSettingsParams = z.infer<typeof UpdateMailboxSettingsInput>;

// =============================================================================
// Mailbox Settings Tools
// =============================================================================

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Mailbox settings tools (automatic replies + mailbox settings).
 */
export class GraphMailboxSettingsTools {
  constructor(private readonly repository: GraphRepository) {}

  async getAutomaticReplies(): Promise<ToolResult> {
    const result = await this.repository.getAutomaticRepliesAsync();
    return jsonResult(result);
  }

  async setAutomaticReplies(params: SetAutomaticRepliesParams): Promise<ToolResult> {
    const replyParams: Parameters<typeof this.repository.setAutomaticRepliesAsync>[0] = {
      status: params.status,
    };
    if (params.external_audience != null) replyParams.externalAudience = params.external_audience;
    if (params.internal_reply_message != null) replyParams.internalReplyMessage = params.internal_reply_message;
    if (params.external_reply_message != null) replyParams.externalReplyMessage = params.external_reply_message;
    if (params.scheduled_start != null) replyParams.scheduledStartDateTime = params.scheduled_start;
    if (params.scheduled_end != null) replyParams.scheduledEndDateTime = params.scheduled_end;
    await this.repository.setAutomaticRepliesAsync(replyParams);
    return jsonResult({ success: true, status: params.status });
  }

  async getMailboxSettings(): Promise<ToolResult> {
    const result = await this.repository.getMailboxSettingsAsync();
    return jsonResult(result);
  }

  async updateMailboxSettings(params: UpdateMailboxSettingsParams): Promise<ToolResult> {
    const settingsParams: Parameters<typeof this.repository.updateMailboxSettingsAsync>[0] = {};
    if (params.language != null) settingsParams.language = params.language;
    if (params.time_zone != null) settingsParams.timeZone = params.time_zone;
    if (params.date_format != null) settingsParams.dateFormat = params.date_format;
    if (params.time_format != null) settingsParams.timeFormat = params.time_format;
    await this.repository.updateMailboxSettingsAsync(settingsParams);
    return jsonResult({ success: true });
  }
}

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2)
// =============================================================================

/**
 * Registry tool definitions for the mailbox-settings domain.
 */
export function mailboxSettingsToolDefinitions(): ToolDefinition[] {
  const tools = (ctx: ToolContext): GraphMailboxSettingsTools => requireGraphToolset(ctx, 'mailboxSettings');

  return [
    defineTool({
      name: 'get_automatic_replies',
      description: 'Get the current automatic replies (out-of-office) settings',
      input: GetAutomaticRepliesInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx) => tools(ctx).getAutomaticReplies(),
    }),
    defineTool({
      name: 'set_automatic_replies',
      description: 'Set automatic replies (out-of-office) settings',
      input: SetAutomaticRepliesInput,
      annotations: { readOnlyHint: false, destructiveHint: false },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).setAutomaticReplies(params),
    }),
    defineTool({
      name: 'get_mailbox_settings',
      description: 'Get the current mailbox settings (language, time zone, date/time formats, working hours)',
      input: GetMailboxSettingsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx) => tools(ctx).getMailboxSettings(),
    }),
    defineTool({
      name: 'update_mailbox_settings',
      description: 'Update mailbox settings (language, time zone, date/time formats)',
      input: UpdateMailboxSettingsInput,
      annotations: { readOnlyHint: false, destructiveHint: false },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).updateMailboxSettings(params),
    }),
  ];
}
