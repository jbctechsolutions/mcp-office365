/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Focused inbox override management MCP tools.
 *
 * Provides tools for managing Outlook focused inbox overrides with a two-phase
 * approval pattern for destructive delete operations.
 */

import { z } from 'zod';
import type { ApprovalTokenManager } from '../approval/index.js';

// =============================================================================
// Input Schemas
// =============================================================================

export const CreateFocusedOverrideInput = z.strictObject({
  sender_address: z.string().email().describe('Sender email address'),
  classify_as: z.enum(['focused', 'other']).describe('Classification'),
});

export const PrepareDeleteFocusedOverrideInput = z.strictObject({
  override_id: z.number().int().positive().describe('Override ID to delete'),
});

export const ConfirmDeleteFocusedOverrideInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_focused_override'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type CreateFocusedOverrideParams = z.infer<typeof CreateFocusedOverrideInput>;
export type PrepareDeleteFocusedOverrideParams = z.infer<typeof PrepareDeleteFocusedOverrideInput>;
export type ConfirmDeleteFocusedOverrideParams = z.infer<typeof ConfirmDeleteFocusedOverrideInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface IFocusedOverridesRepository {
  listFocusedOverridesAsync(): Promise<Array<{ id: number; senderAddress: string; classifyAs: string }>>;
  createFocusedOverrideAsync(senderAddress: string, classifyAs: 'focused' | 'other'): Promise<number>;
  deleteFocusedOverrideAsync(overrideId: number): Promise<void>;
}

// =============================================================================
// Focused Overrides Tools
// =============================================================================

/**
 * Focused inbox override tools with two-phase approval for delete operations.
 */
export class FocusedOverridesTools {
  constructor(
    private readonly repo: IFocusedOverridesRepository,
    private readonly tokenManager: ApprovalTokenManager,
  ) {}

  async listFocusedOverrides(): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const overrides = await this.repo.listFocusedOverridesAsync();
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ overrides }, null, 2),
      }],
    };
  }

  async createFocusedOverride(params: CreateFocusedOverrideParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const overrideId = await this.repo.createFocusedOverrideAsync(params.sender_address, params.classify_as);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, override_id: overrideId, message: 'Focused override created' }, null, 2),
      }],
    };
  }

  prepareDeleteFocusedOverride(params: PrepareDeleteFocusedOverrideParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'delete_focused_override',
      targetType: 'focused_override',
      targetId: params.override_id,
      targetHash: String(params.override_id),
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          override_id: params.override_id,
          action: `To confirm deleting focused override ${params.override_id}, call confirm_delete_focused_override with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmDeleteFocusedOverride(params: ConfirmDeleteFocusedOverrideParams): Promise<{
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

    const result = this.tokenManager.consumeToken(params.approval_token, 'delete_focused_override', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_focused_override again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_focused_override',
        TARGET_MISMATCH: 'Token was generated for a different focused override',
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

    await this.repo.deleteFocusedOverrideAsync(result.token!.targetId);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Focused override deleted' }, null, 2),
      }],
    };
  }
}
