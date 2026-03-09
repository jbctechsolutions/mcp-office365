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
  | 'delete_task'
  | 'delete_mail_rule'
  | 'delete_task_list'
  | 'delete_contact_folder'
  | 'delete_category'
  | 'delete_focused_override'
  | 'delete_calendar_permission'
  | 'delete_channel'
  | 'send_channel_message'
  | 'reply_channel_message'
  | 'send_chat_message'
  | 'delete_checklist_item'
  | 'delete_linked_resource'
  | 'delete_task_attachment'
  | 'delete_bucket'
  | 'delete_planner_task'
  | 'update_excel_range'
  | 'upload_file'
  | 'delete_drive_item';

/**
 * Target resource types.
 */
export type TargetType = 'email' | 'folder' | 'event' | 'contact' | 'task' | 'task_list' | 'rule' | 'contact_folder' | 'category' | 'focused_override' | 'calendar_permission' | 'channel' | 'channel_message' | 'chat_message' | 'checklist_item' | 'linked_resource' | 'task_attachment' | 'bucket' | 'planner_task' | 'excel_range' | 'drive_item';

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
