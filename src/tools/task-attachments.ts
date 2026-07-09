/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Task Attachment MCP tools for Microsoft To Do.
 *
 * Provides tools for managing attachments on To Do tasks with a two-phase
 * approval pattern for destructive delete operations.
 */

import { z } from 'zod';
import { Id } from '../ids/schema.js';
import type { ApprovalTokenManager } from '../approval/index.js';
import { defineTool } from '../registry/define-tool.js';
import { approvalTokenLink } from '../registry/elicit-links.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolContext, ToolDefinition } from '../registry/types.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    taskAttachments: TaskAttachmentsTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const ListTaskAttachmentsInput = z.strictObject({
  task_id: Id.task,
});

export const CreateTaskAttachmentInput = z.strictObject({
  task_id: Id.task,
  name: z.string().min(1).describe('File name of the attachment'),
  content_bytes: z.string().min(1).describe('Base64-encoded file content'),
  content_type: z.string().optional().describe('MIME type (default: application/octet-stream)'),
});

export const PrepareDeleteTaskAttachmentInput = z.strictObject({
  task_attachment_id: Id.taskAttachment,
});

export const ConfirmDeleteTaskAttachmentInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_task_attachment'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListTaskAttachmentsParams = z.infer<typeof ListTaskAttachmentsInput>;
export type CreateTaskAttachmentParams = z.infer<typeof CreateTaskAttachmentInput>;
export type PrepareDeleteTaskAttachmentParams = z.infer<typeof PrepareDeleteTaskAttachmentInput>;
export type ConfirmDeleteTaskAttachmentParams = z.infer<typeof ConfirmDeleteTaskAttachmentInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface ITaskAttachmentsRepository {
  listTaskAttachmentsAsync(taskId: string): Promise<Array<{ id: string; name: string; size: number; contentType: string }>>;
  createTaskAttachmentAsync(taskId: string, name: string, contentBytes: string, contentType?: string): Promise<string>;
  deleteTaskAttachmentAsync(taskAttachmentId: string): Promise<void>;
}

// =============================================================================
// Task Attachments Tools
// =============================================================================

/**
 * Task attachment tools with two-phase approval for delete operations.
 */
export class TaskAttachmentsTools {
  constructor(
    private readonly repo: ITaskAttachmentsRepository,
    private readonly tokenManager: ApprovalTokenManager,
  ) {}

  async listTaskAttachments(params: ListTaskAttachmentsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const items = await this.repo.listTaskAttachmentsAsync(params.task_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ task_attachments: items }, null, 2),
      }],
    };
  }

  async createTaskAttachment(params: CreateTaskAttachmentParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const itemId = await this.repo.createTaskAttachmentAsync(params.task_id, params.name, params.content_bytes, params.content_type);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, task_attachment_id: itemId, message: 'Task attachment created' }, null, 2),
      }],
    };
  }

  prepareDeleteTaskAttachment(params: PrepareDeleteTaskAttachmentParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'delete_task_attachment',
      targetType: 'task_attachment',
      targetId: params.task_attachment_id,
      targetHash: String(params.task_attachment_id),
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          task_attachment_id: params.task_attachment_id,
          action: `To confirm deleting task attachment ${params.task_attachment_id}, call confirm_delete_task_attachment with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmDeleteTaskAttachment(params: ConfirmDeleteTaskAttachmentParams): Promise<{
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

    const result = this.tokenManager.consumeToken(params.approval_token, 'delete_task_attachment', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_task_attachment again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_task_attachment',
        TARGET_MISMATCH: 'Token was generated for a different task attachment',
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

    await this.repo.deleteTaskAttachmentAsync((result.token!.targetId));
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Task attachment deleted' }, null, 2),
      }],
    };
  }
}

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2)
// =============================================================================

/**
 * Registry tool definitions for the task-attachments domain.
 */
export function taskAttachmentsToolDefinitions(): ToolDefinition[] {
  const tools = (ctx: ToolContext): TaskAttachmentsTools => requireGraphToolset(ctx, 'taskAttachments');

  return [
    defineTool({
      name: 'list_task_attachments',
      description: 'List attachments on a To Do task (Graph API)',
      input: ListTaskAttachmentsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listTaskAttachments(params),
    }),
    defineTool({
      name: 'create_task_attachment',
      description: 'Attach a file to a To Do task (small files only, base64 encoded) (Graph API)',
      input: CreateTaskAttachmentInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).createTaskAttachment(params),
    }),
    defineTool({
      name: 'prepare_delete_task_attachment',
      description: 'Prepare to delete a task attachment. Returns an approval token. Call confirm_delete_task_attachment to execute. (Graph API)',
      input: PrepareDeleteTaskAttachmentInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareDeleteTaskAttachment(params),
      onElicit: approvalTokenLink('confirm_delete_task_attachment'),
    }),
    defineTool({
      name: 'confirm_delete_task_attachment',
      description: 'Confirm deletion of a task attachment with approval token (Graph API)',
      input: ConfirmDeleteTaskAttachmentInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmDeleteTaskAttachment(params),
    }),
  ];
}
