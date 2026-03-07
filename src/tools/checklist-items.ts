/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Checklist Item (subtask) MCP tools for Microsoft To Do.
 *
 * Provides tools for managing checklist items on To Do tasks with a two-phase
 * approval pattern for destructive delete operations.
 */

import { z } from 'zod';
import type { ApprovalTokenManager } from '../approval/index.js';

// =============================================================================
// Input Schemas
// =============================================================================

export const ListChecklistItemsInput = z.strictObject({
  task_id: z.number().int().positive().describe('Task ID from list_tasks or search_tasks'),
});

export const CreateChecklistItemInput = z.strictObject({
  task_id: z.number().int().positive().describe('Task ID'),
  display_name: z.string().min(1).describe('Checklist item text'),
  is_checked: z.boolean().optional().describe('Whether the item is checked (default: false)'),
});

export const UpdateChecklistItemInput = z.strictObject({
  checklist_item_id: z.number().int().positive().describe('Checklist item ID'),
  display_name: z.string().min(1).optional().describe('New text'),
  is_checked: z.boolean().optional().describe('Toggle checked state'),
});

export const PrepareDeleteChecklistItemInput = z.strictObject({
  checklist_item_id: z.number().int().positive().describe('Checklist item ID to delete'),
});

export const ConfirmDeleteChecklistItemInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_checklist_item'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListChecklistItemsParams = z.infer<typeof ListChecklistItemsInput>;
export type CreateChecklistItemParams = z.infer<typeof CreateChecklistItemInput>;
export type UpdateChecklistItemParams = z.infer<typeof UpdateChecklistItemInput>;
export type PrepareDeleteChecklistItemParams = z.infer<typeof PrepareDeleteChecklistItemInput>;
export type ConfirmDeleteChecklistItemParams = z.infer<typeof ConfirmDeleteChecklistItemInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface IChecklistItemsRepository {
  listChecklistItemsAsync(taskId: number): Promise<Array<{ id: number; displayName: string; isChecked: boolean; createdDateTime: string }>>;
  createChecklistItemAsync(taskId: number, displayName: string, isChecked?: boolean): Promise<number>;
  updateChecklistItemAsync(checklistItemId: number, updates: { displayName?: string; isChecked?: boolean }): Promise<void>;
  deleteChecklistItemAsync(checklistItemId: number): Promise<void>;
}

// =============================================================================
// Checklist Items Tools
// =============================================================================

/**
 * Checklist item tools with two-phase approval for delete operations.
 */
export class ChecklistItemsTools {
  constructor(
    private readonly repo: IChecklistItemsRepository,
    private readonly tokenManager: ApprovalTokenManager,
  ) {}

  async listChecklistItems(params: ListChecklistItemsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const items = await this.repo.listChecklistItemsAsync(params.task_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ checklist_items: items }, null, 2),
      }],
    };
  }

  async createChecklistItem(params: CreateChecklistItemParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const itemId = await this.repo.createChecklistItemAsync(params.task_id, params.display_name, params.is_checked);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, checklist_item_id: itemId, message: 'Checklist item created' }, null, 2),
      }],
    };
  }

  async updateChecklistItem(params: UpdateChecklistItemParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const updates: { displayName?: string; isChecked?: boolean } = {};
    if (params.display_name != null) updates.displayName = params.display_name;
    if (params.is_checked != null) updates.isChecked = params.is_checked;
    await this.repo.updateChecklistItemAsync(params.checklist_item_id, updates);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Checklist item updated' }, null, 2),
      }],
    };
  }

  prepareDeleteChecklistItem(params: PrepareDeleteChecklistItemParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'delete_checklist_item',
      targetType: 'checklist_item',
      targetId: params.checklist_item_id,
      targetHash: String(params.checklist_item_id),
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          checklist_item_id: params.checklist_item_id,
          action: `To confirm deleting checklist item ${params.checklist_item_id}, call confirm_delete_checklist_item with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmDeleteChecklistItem(params: ConfirmDeleteChecklistItemParams): Promise<{
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

    const result = this.tokenManager.consumeToken(params.approval_token, 'delete_checklist_item', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_checklist_item again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_checklist_item',
        TARGET_MISMATCH: 'Token was generated for a different checklist item',
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

    await this.repo.deleteChecklistItemAsync(result.token!.targetId);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Checklist item deleted' }, null, 2),
      }],
    };
  }
}
