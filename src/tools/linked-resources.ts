/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Linked Resource MCP tools for Microsoft To Do.
 *
 * Provides tools for managing linked resources on To Do tasks with a two-phase
 * approval pattern for destructive delete operations.
 */

import { z } from 'zod';
import type { ApprovalTokenManager } from '../approval/index.js';

// =============================================================================
// Input Schemas
// =============================================================================

export const ListLinkedResourcesInput = z.strictObject({
  task_id: z.number().int().positive().describe('Task ID from list_tasks or search_tasks'),
});

export const CreateLinkedResourceInput = z.strictObject({
  task_id: z.number().int().positive().describe('Task ID'),
  web_url: z.string().min(1).describe('URL of the linked resource'),
  application_name: z.string().min(1).describe('Name of the application the resource is associated with'),
  display_name: z.string().min(1).optional().describe('Display name of the linked resource'),
});

export const PrepareDeleteLinkedResourceInput = z.strictObject({
  linked_resource_id: z.number().int().positive().describe('Linked resource ID to delete'),
});

export const ConfirmDeleteLinkedResourceInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_linked_resource'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListLinkedResourcesParams = z.infer<typeof ListLinkedResourcesInput>;
export type CreateLinkedResourceParams = z.infer<typeof CreateLinkedResourceInput>;
export type PrepareDeleteLinkedResourceParams = z.infer<typeof PrepareDeleteLinkedResourceInput>;
export type ConfirmDeleteLinkedResourceParams = z.infer<typeof ConfirmDeleteLinkedResourceInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface ILinkedResourcesRepository {
  listLinkedResourcesAsync(taskId: number): Promise<Array<{ id: number; webUrl: string; applicationName: string; displayName: string }>>;
  createLinkedResourceAsync(taskId: number, webUrl: string, applicationName: string, displayName?: string): Promise<number>;
  deleteLinkedResourceAsync(linkedResourceId: number): Promise<void>;
}

// =============================================================================
// Linked Resources Tools
// =============================================================================

/**
 * Linked resource tools with two-phase approval for delete operations.
 */
export class LinkedResourcesTools {
  constructor(
    private readonly repo: ILinkedResourcesRepository,
    private readonly tokenManager: ApprovalTokenManager,
  ) {}

  async listLinkedResources(params: ListLinkedResourcesParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const items = await this.repo.listLinkedResourcesAsync(params.task_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ linked_resources: items }, null, 2),
      }],
    };
  }

  async createLinkedResource(params: CreateLinkedResourceParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const itemId = await this.repo.createLinkedResourceAsync(params.task_id, params.web_url, params.application_name, params.display_name);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, linked_resource_id: itemId, message: 'Linked resource created' }, null, 2),
      }],
    };
  }

  prepareDeleteLinkedResource(params: PrepareDeleteLinkedResourceParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'delete_linked_resource',
      targetType: 'linked_resource',
      targetId: params.linked_resource_id,
      targetHash: String(params.linked_resource_id),
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          linked_resource_id: params.linked_resource_id,
          action: `To confirm deleting linked resource ${params.linked_resource_id}, call confirm_delete_linked_resource with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmDeleteLinkedResource(params: ConfirmDeleteLinkedResourceParams): Promise<{
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

    const result = this.tokenManager.consumeToken(params.approval_token, 'delete_linked_resource', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_linked_resource again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_linked_resource',
        TARGET_MISMATCH: 'Token was generated for a different linked resource',
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

    await this.repo.deleteLinkedResourceAsync(result.token!.targetId);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Linked resource deleted' }, null, 2),
      }],
    };
  }
}
