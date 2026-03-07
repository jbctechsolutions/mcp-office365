/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Calendar permission management MCP tools.
 *
 * Provides tools for managing calendar sharing permissions with a two-phase
 * approval pattern for destructive delete operations.
 */

import { z } from 'zod';
import type { ApprovalTokenManager } from '../approval/index.js';

// =============================================================================
// Input Schemas
// =============================================================================

export const ListCalendarPermissionsInput = z.strictObject({
  calendar_id: z.number().int().positive().describe('Calendar ID'),
});

export const CreateCalendarPermissionInput = z.strictObject({
  calendar_id: z.number().int().positive().describe('Calendar ID'),
  email_address: z.string().email().describe('Email of person to share with'),
  role: z.enum(['read', 'write', 'delegateWithoutPrivateEventAccess', 'delegateWithPrivateEventAccess']).describe('Permission level'),
});

export const PrepareDeleteCalendarPermissionInput = z.strictObject({
  permission_id: z.number().int().positive().describe('Calendar permission ID'),
});

export const ConfirmDeleteCalendarPermissionInput = z.strictObject({
  approval_token: z.string().describe('Approval token'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListCalendarPermissionsParams = z.infer<typeof ListCalendarPermissionsInput>;
export type CreateCalendarPermissionParams = z.infer<typeof CreateCalendarPermissionInput>;
export type PrepareDeleteCalendarPermissionParams = z.infer<typeof PrepareDeleteCalendarPermissionInput>;
export type ConfirmDeleteCalendarPermissionParams = z.infer<typeof ConfirmDeleteCalendarPermissionInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface ICalendarPermissionsRepository {
  listCalendarPermissionsAsync(calendarId: number): Promise<Array<{ id: number; emailAddress: string; role: string; isRemovable: boolean; isInsideOrganization: boolean }>>;
  createCalendarPermissionAsync(calendarId: number, email: string, role: string): Promise<number>;
  deleteCalendarPermissionAsync(permissionId: number): Promise<void>;
}

// =============================================================================
// Calendar Permissions Tools
// =============================================================================

/**
 * Calendar permission tools with two-phase approval for delete operations.
 */
export class CalendarPermissionsTools {
  constructor(
    private readonly repo: ICalendarPermissionsRepository,
    private readonly tokenManager: ApprovalTokenManager,
  ) {}

  async listCalendarPermissions(params: ListCalendarPermissionsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const permissions = await this.repo.listCalendarPermissionsAsync(params.calendar_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ permissions }, null, 2),
      }],
    };
  }

  async createCalendarPermission(params: CreateCalendarPermissionParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const permissionId = await this.repo.createCalendarPermissionAsync(params.calendar_id, params.email_address, params.role);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, permission_id: permissionId, message: 'Calendar permission created' }, null, 2),
      }],
    };
  }

  prepareDeleteCalendarPermission(params: PrepareDeleteCalendarPermissionParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'delete_calendar_permission',
      targetType: 'calendar_permission',
      targetId: params.permission_id,
      targetHash: String(params.permission_id),
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          permission_id: params.permission_id,
          action: `To confirm deleting calendar permission ${params.permission_id}, call confirm_delete_calendar_permission with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmDeleteCalendarPermission(params: ConfirmDeleteCalendarPermissionParams): Promise<{
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

    const result = this.tokenManager.consumeToken(params.approval_token, 'delete_calendar_permission', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_calendar_permission again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_calendar_permission',
        TARGET_MISMATCH: 'Token was generated for a different calendar permission',
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

    await this.repo.deleteCalendarPermissionAsync(result.token!.targetId);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Calendar permission deleted' }, null, 2),
      }],
    };
  }
}
