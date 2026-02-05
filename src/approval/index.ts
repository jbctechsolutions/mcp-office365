/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Approval token system for safe destructive operations.
 *
 * Re-exports all approval types, hashing, and token management.
 */

export {
  type OperationType,
  type TargetType,
  type ApprovalToken,
  type ValidationErrorReason,
  type ValidationResult,
} from './types.js';

export { hashEmailForApproval, hashFolderForApproval } from './hash.js';

export { ApprovalTokenManager } from './token-manager.js';
