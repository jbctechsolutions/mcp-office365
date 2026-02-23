/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Types for the approval token system.
 *
 * Provides safety for destructive mailbox operations by requiring
 * explicit user approval per object via time-limited tokens.
 */

// =============================================================================
// Operation Types
// =============================================================================

/**
 * Operations that require approval tokens.
 */
export type OperationType =
  | 'delete_email'
  | 'move_email'
  | 'archive_email'
  | 'junk_email'
  | 'delete_folder'
  | 'empty_folder'
  | 'batch_delete_emails'
  | 'batch_move_emails'
  | 'send_draft'
  | 'send_email'
  | 'reply_email'
  | 'forward_email'
  | 'delete_event'
  | 'delete_contact'
  | 'delete_task';

/**
 * Target resource types.
 */
export type TargetType = 'email' | 'folder' | 'event' | 'contact' | 'task';

// =============================================================================
// Token Types
// =============================================================================

/**
 * An approval token authorizing a single destructive operation.
 */
export interface ApprovalToken {
  readonly tokenId: string;
  readonly operation: OperationType;
  readonly targetType: TargetType;
  readonly targetId: number;
  readonly targetHash: string;
  readonly createdAt: number;
  readonly expiresAt: number;
  readonly metadata: Readonly<Record<string, unknown>>;
}

/**
 * Reasons a token validation can fail.
 */
export type ValidationErrorReason =
  | 'EXPIRED'
  | 'NOT_FOUND'
  | 'OPERATION_MISMATCH'
  | 'TARGET_MISMATCH'
  | 'TARGET_CHANGED'
  | 'ALREADY_CONSUMED';

/**
 * Result of token validation.
 */
export interface ValidationResult {
  readonly valid: boolean;
  readonly error?: ValidationErrorReason;
  readonly token?: ApprovalToken;
}
