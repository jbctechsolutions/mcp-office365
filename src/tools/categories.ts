/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Master category management MCP tools.
 *
 * Provides tools for managing Outlook master categories with a two-phase
 * approval pattern for destructive delete operations.
 */

import { z } from 'zod';
import type { ApprovalTokenManager } from '../approval/index.js';

// =============================================================================
// Input Schemas
// =============================================================================

export const CreateCategoryInput = z.strictObject({
  name: z.string().min(1).describe('Category name'),
  color: z.enum(['preset0','preset1','preset2','preset3','preset4','preset5','preset6','preset7','preset8','preset9','preset10','preset11','preset12','preset13','preset14','preset15','preset16','preset17','preset18','preset19','preset20','preset21','preset22','preset23','preset24','none']).describe('Category color preset'),
});

export const PrepareDeleteCategoryInput = z.strictObject({
  category_id: z.number().int().positive().describe('Category ID to delete'),
});

export const ConfirmDeleteCategoryInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_category'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type CreateCategoryParams = z.infer<typeof CreateCategoryInput>;
export type PrepareDeleteCategoryParams = z.infer<typeof PrepareDeleteCategoryInput>;
export type ConfirmDeleteCategoryParams = z.infer<typeof ConfirmDeleteCategoryInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface ICategoriesRepository {
  listCategoriesAsync(): Promise<Array<{ id: number; name: string; color: string }>>;
  createCategoryAsync(name: string, color: string): Promise<number>;
  deleteCategoryAsync(categoryId: number): Promise<void>;
}

// =============================================================================
// Categories Tools
// =============================================================================

/**
 * Master category tools with two-phase approval for delete operations.
 */
export class CategoriesTools {
  constructor(
    private readonly repo: ICategoriesRepository,
    private readonly tokenManager: ApprovalTokenManager,
  ) {}

  async listCategories(): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const categories = await this.repo.listCategoriesAsync();
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ categories }, null, 2),
      }],
    };
  }

  async createCategory(params: CreateCategoryParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const categoryId = await this.repo.createCategoryAsync(params.name, params.color);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, category_id: categoryId, message: 'Category created' }, null, 2),
      }],
    };
  }

  prepareDeleteCategory(params: PrepareDeleteCategoryParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'delete_category',
      targetType: 'category',
      targetId: params.category_id,
      targetHash: String(params.category_id),
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          category_id: params.category_id,
          action: `To confirm deleting category ${params.category_id}, call confirm_delete_category with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmDeleteCategory(params: ConfirmDeleteCategoryParams): Promise<{
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

    const result = this.tokenManager.consumeToken(params.approval_token, 'delete_category', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_category again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_category',
        TARGET_MISMATCH: 'Token was generated for a different category',
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

    await this.repo.deleteCategoryAsync(result.token!.targetId);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Category deleted' }, null, 2),
      }],
    };
  }
}
