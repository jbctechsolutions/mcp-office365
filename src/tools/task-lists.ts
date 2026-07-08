/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Graph-backend task list tools (v3 registry-driven architecture, U2). Holds
 * the Microsoft To Do task list logic that previously lived inline in the
 * `handleGraphToolCall` switch, with a two-phase approval pattern for the
 * destructive delete operation.
 */

import { z } from 'zod';
import type { GraphRepository } from '../graph/repository.js';
import type { ApprovalTokenManager } from '../approval/index.js';
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolContext, ToolDefinition, ToolResult } from '../registry/types.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    taskLists: GraphTaskListsTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const ListTaskListsInput = z.strictObject({});

export const CreateTaskListInput = z.strictObject({
  display_name: z.string().min(1).describe('Name for the new task list'),
});

export const RenameTaskListInput = z.strictObject({
  task_list_id: z.number().int().positive().describe('Task list ID'),
  name: z.string().min(1).describe('New name for the task list'),
});

export const PrepareDeleteTaskListInput = z.strictObject({
  task_list_id: z.number().int().positive().describe('Task list ID to delete'),
});

export const ConfirmDeleteTaskListInput = z.strictObject({
  token_id: z.string().uuid().describe('Approval token from prepare_delete_task_list'),
  task_list_id: z.number().int().positive().describe('The task list ID to delete'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type CreateTaskListParams = z.infer<typeof CreateTaskListInput>;
export type RenameTaskListParams = z.infer<typeof RenameTaskListInput>;
export type PrepareDeleteTaskListParams = z.infer<typeof PrepareDeleteTaskListInput>;
export type ConfirmDeleteTaskListParams = z.infer<typeof ConfirmDeleteTaskListInput>;

// =============================================================================
// Task Lists Tools
// =============================================================================

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Task list tools with two-phase approval for delete operations.
 */
export class GraphTaskListsTools {
  constructor(
    private readonly repository: GraphRepository,
    private readonly tokenManager: ApprovalTokenManager
  ) {}

  async listTaskLists(): Promise<ToolResult> {
    const lists = await this.repository.listTaskListsAsync();
    return jsonResult({ task_lists: lists });
  }

  async createTaskList(params: CreateTaskListParams): Promise<ToolResult> {
    const numericId = await this.repository.createTaskListAsync(params.display_name);
    return jsonResult({ id: numericId, display_name: params.display_name, status: 'created' });
  }

  async renameTaskList(params: RenameTaskListParams): Promise<ToolResult> {
    await this.repository.renameTaskListAsync(params.task_list_id, params.name);
    return jsonResult({ success: true, message: 'Task list renamed' });
  }

  prepareDeleteTaskList(params: PrepareDeleteTaskListParams): ToolResult {
    const token = this.tokenManager.generateToken({
      operation: 'delete_task_list',
      targetType: 'task_list',
      targetId: params.task_list_id,
      targetHash: String(params.task_list_id),
    });

    return jsonResult({
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      task_list_id: params.task_list_id,
      action: `To confirm deleting task list ${params.task_list_id}, call confirm_delete_task_list with the token_id and task_list_id.`,
    });
  }

  async confirmDeleteTaskList(params: ConfirmDeleteTaskListParams): Promise<ToolResult> {
    const validation = this.tokenManager.consumeToken(params.token_id, 'delete_task_list', params.task_list_id);
    if (!validation.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_task_list again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_task_list',
        TARGET_MISMATCH: 'Token was generated for a different task list',
        ALREADY_CONSUMED: 'Token has already been used',
      };
      return jsonResult({ success: false, error: errorMessages[validation.error ?? ''] ?? 'Invalid token' });
    }

    await this.repository.deleteTaskListAsync(params.task_list_id);
    return jsonResult({ success: true, message: 'Task list deleted' });
  }
}

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2)
// =============================================================================

/**
 * Registry tool definitions for the task-lists domain.
 */
export function taskListsToolDefinitions(): ToolDefinition[] {
  const tools = (ctx: ToolContext): GraphTaskListsTools => requireGraphToolset(ctx, 'taskLists');

  return [
    defineTool({
      name: 'list_task_lists',
      description: 'List all task lists (Microsoft To Do) (Graph API)',
      input: ListTaskListsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx) => tools(ctx).listTaskLists(),
    }),
    defineTool({
      name: 'create_task_list',
      description: 'Create a new task list',
      input: CreateTaskListInput,
      annotations: { readOnlyHint: false, destructiveHint: false },
      destructive: false,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).createTaskList(params),
    }),
    defineTool({
      name: 'rename_task_list',
      description: 'Rename a task list (Graph API)',
      input: RenameTaskListInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).renameTaskList(params),
    }),
    defineTool({
      name: 'prepare_delete_task_list',
      description: 'Prepare to delete a task list. Returns an approval token. Call confirm_delete_task_list to execute. (Graph API)',
      input: PrepareDeleteTaskListInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareDeleteTaskList(params),
    }),
    defineTool({
      name: 'confirm_delete_task_list',
      description: 'Confirm task list deletion with approval token (Graph API)',
      input: ConfirmDeleteTaskListInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmDeleteTaskList(params),
    }),
  ];
}
