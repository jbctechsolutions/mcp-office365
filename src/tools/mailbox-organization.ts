/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Mailbox organization MCP tools.
 *
 * Provides tools for organizing emails and folders with a two-phase
 * approval pattern for destructive operations (prepare → confirm).
 */

import { z } from 'zod';
import type { IMailboxRepository, EmailRow, FolderRow } from '../database/repository.js';
import {
  ApprovalTokenManager,
  hashEmailForApproval,
  hashFolderForApproval,
  type ApprovalToken,
  type OperationType,
  type ValidationErrorReason,
} from '../approval/index.js';
import {
  ApprovalExpiredError,
  ApprovalInvalidError,
  TargetChangedError,
  NotFoundError,
} from '../utils/errors.js';
import { appleTimestampToIso } from '../utils/dates.js';

// =============================================================================
// Input Schemas — Destructive Operations (Two-Phase)
// =============================================================================

export const PrepareDeleteEmailInput = z
  .object({
    email_id: z.number().int().positive().describe('The email ID to delete'),
  })
  .strict();

export const ConfirmDeleteEmailInput = z
  .object({
    token_id: z.string().uuid().describe('The approval token from prepare_delete_email'),
    email_id: z.number().int().positive().describe('The email ID to delete'),
  })
  .strict();

export const PrepareMoveEmailInput = z
  .object({
    email_id: z.number().int().positive().describe('The email ID to move'),
    destination_folder_id: z.number().int().positive().describe('The destination folder ID'),
  })
  .strict();

export const ConfirmMoveEmailInput = z
  .object({
    token_id: z.string().uuid().describe('The approval token from prepare_move_email'),
    email_id: z.number().int().positive().describe('The email ID to move'),
  })
  .strict();

export const PrepareArchiveEmailInput = z
  .object({
    email_id: z.number().int().positive().describe('The email ID to archive'),
  })
  .strict();

export const ConfirmArchiveEmailInput = z
  .object({
    token_id: z.string().uuid().describe('The approval token from prepare_archive_email'),
    email_id: z.number().int().positive().describe('The email ID to archive'),
  })
  .strict();

export const PrepareJunkEmailInput = z
  .object({
    email_id: z.number().int().positive().describe('The email ID to mark as junk'),
  })
  .strict();

export const ConfirmJunkEmailInput = z
  .object({
    token_id: z.string().uuid().describe('The approval token from prepare_junk_email'),
    email_id: z.number().int().positive().describe('The email ID to mark as junk'),
  })
  .strict();

export const PrepareDeleteFolderInput = z
  .object({
    folder_id: z.number().int().positive().describe('The folder ID to delete'),
  })
  .strict();

export const ConfirmDeleteFolderInput = z
  .object({
    token_id: z.string().uuid().describe('The approval token from prepare_delete_folder'),
    folder_id: z.number().int().positive().describe('The folder ID to delete'),
  })
  .strict();

export const PrepareEmptyFolderInput = z
  .object({
    folder_id: z.number().int().positive().describe('The folder ID to empty'),
  })
  .strict();

export const ConfirmEmptyFolderInput = z
  .object({
    token_id: z.string().uuid().describe('The approval token from prepare_empty_folder'),
    folder_id: z.number().int().positive().describe('The folder ID to empty'),
  })
  .strict();

export const PrepareBatchDeleteEmailsInput = z
  .object({
    email_ids: z
      .array(z.number().int().positive())
      .min(1)
      .max(50)
      .describe('The email IDs to delete (max 50)'),
  })
  .strict();

export const PrepareBatchMoveEmailsInput = z
  .object({
    email_ids: z
      .array(z.number().int().positive())
      .min(1)
      .max(50)
      .describe('The email IDs to move (max 50)'),
    destination_folder_id: z.number().int().positive().describe('The destination folder ID'),
  })
  .strict();

export const ConfirmBatchOperationInput = z
  .object({
    tokens: z
      .array(
        z.object({
          token_id: z.string().uuid().describe('The approval token'),
          email_id: z.number().int().positive().describe('The email ID'),
        })
      )
      .min(1)
      .max(50)
      .describe('Array of token/email pairs to confirm'),
  })
  .strict();

// =============================================================================
// Input Schemas — Low-Risk Modifications (Single Tool)
// =============================================================================

export const MarkEmailReadInput = z
  .object({
    email_id: z.number().int().positive().describe('The email ID to mark as read'),
  })
  .strict();

export const MarkEmailUnreadInput = z
  .object({
    email_id: z.number().int().positive().describe('The email ID to mark as unread'),
  })
  .strict();

export const SetEmailFlagInput = z
  .object({
    email_id: z.number().int().positive().describe('The email ID to flag'),
    flag_status: z
      .number()
      .int()
      .min(0)
      .max(2)
      .describe('Flag status: 0=not flagged, 1=flagged, 2=completed'),
  })
  .strict();

export const ClearEmailFlagInput = z
  .object({
    email_id: z.number().int().positive().describe('The email ID to clear the flag from'),
  })
  .strict();

export const SetEmailCategoriesInput = z
  .object({
    email_id: z.number().int().positive().describe('The email ID'),
    categories: z
      .array(z.string().min(1))
      .describe('Categories to set (replaces existing). Use empty array to clear.'),
  })
  .strict();

// =============================================================================
// Input Schemas — Non-Destructive Operations
// =============================================================================

export const CreateFolderInput = z
  .object({
    name: z.string().min(1).max(255).describe('Name for the new folder'),
    parent_folder_id: z
      .number()
      .int()
      .positive()
      .optional()
      .describe('Optional parent folder ID (creates top-level if omitted)'),
  })
  .strict();

export const RenameFolderInput = z
  .object({
    folder_id: z.number().int().positive().describe('The folder ID to rename'),
    new_name: z.string().min(1).max(255).describe('The new folder name'),
  })
  .strict();

export const MoveFolderInput = z
  .object({
    folder_id: z.number().int().positive().describe('The folder ID to move'),
    destination_parent_id: z
      .number()
      .int()
      .positive()
      .describe('The destination parent folder ID'),
  })
  .strict();

// =============================================================================
// Type Exports
// =============================================================================

export type PrepareDeleteEmailParams = z.infer<typeof PrepareDeleteEmailInput>;
export type ConfirmDeleteEmailParams = z.infer<typeof ConfirmDeleteEmailInput>;
export type PrepareMoveEmailParams = z.infer<typeof PrepareMoveEmailInput>;
export type ConfirmMoveEmailParams = z.infer<typeof ConfirmMoveEmailInput>;
export type PrepareArchiveEmailParams = z.infer<typeof PrepareArchiveEmailInput>;
export type ConfirmArchiveEmailParams = z.infer<typeof ConfirmArchiveEmailInput>;
export type PrepareJunkEmailParams = z.infer<typeof PrepareJunkEmailInput>;
export type ConfirmJunkEmailParams = z.infer<typeof ConfirmJunkEmailInput>;
export type PrepareDeleteFolderParams = z.infer<typeof PrepareDeleteFolderInput>;
export type ConfirmDeleteFolderParams = z.infer<typeof ConfirmDeleteFolderInput>;
export type PrepareEmptyFolderParams = z.infer<typeof PrepareEmptyFolderInput>;
export type ConfirmEmptyFolderParams = z.infer<typeof ConfirmEmptyFolderInput>;
export type PrepareBatchDeleteEmailsParams = z.infer<typeof PrepareBatchDeleteEmailsInput>;
export type PrepareBatchMoveEmailsParams = z.infer<typeof PrepareBatchMoveEmailsInput>;
export type ConfirmBatchOperationParams = z.infer<typeof ConfirmBatchOperationInput>;
export type MarkEmailReadParams = z.infer<typeof MarkEmailReadInput>;
export type MarkEmailUnreadParams = z.infer<typeof MarkEmailUnreadInput>;
export type SetEmailFlagParams = z.infer<typeof SetEmailFlagInput>;
export type ClearEmailFlagParams = z.infer<typeof ClearEmailFlagInput>;
export type SetEmailCategoriesParams = z.infer<typeof SetEmailCategoriesInput>;
export type CreateFolderParams = z.infer<typeof CreateFolderInput>;
export type RenameFolderParams = z.infer<typeof RenameFolderInput>;
export type MoveFolderParams = z.infer<typeof MoveFolderInput>;

// =============================================================================
// Preview Helpers
// =============================================================================

function emailPreview(row: EmailRow): {
  id: number;
  subject: string | null;
  sender: string | null;
  senderAddress: string | null;
  folderId: number | null;
  timeReceived: string | null;
} {
  return {
    id: row.id,
    subject: row.subject,
    sender: row.sender,
    senderAddress: row.senderAddress,
    folderId: row.folderId,
    timeReceived: row.timeReceived != null ? appleTimestampToIso(row.timeReceived) : null,
  };
}

function folderPreview(row: FolderRow): {
  id: number;
  name: string | null;
  messageCount: number;
  unreadCount: number;
} {
  return {
    id: row.id,
    name: row.name,
    messageCount: row.messageCount,
    unreadCount: row.unreadCount,
  };
}

// =============================================================================
// Validation Helpers
// =============================================================================

function throwValidationError(error: ValidationErrorReason): never {
  switch (error) {
    case 'EXPIRED':
      throw new ApprovalExpiredError();
    case 'NOT_FOUND':
      throw new ApprovalInvalidError('Token not found or already used');
    case 'OPERATION_MISMATCH':
      throw new ApprovalInvalidError('Token was issued for a different operation');
    case 'TARGET_MISMATCH':
      throw new ApprovalInvalidError('Token was issued for a different target');
    case 'ALREADY_CONSUMED':
      throw new ApprovalInvalidError('Token has already been used');
    case 'TARGET_CHANGED':
      throw new TargetChangedError();
  }
}

// =============================================================================
// Mailbox Organization Tools
// =============================================================================

/**
 * Mailbox organization tools with two-phase approval for destructive ops.
 *
 * Works with both sync (AppleScript) and async (Graph) backends via
 * the IMailboxRepository interface and MaybePromise return types.
 */
export class MailboxOrganizationTools {
  constructor(
    private readonly repository: IMailboxRepository,
    private readonly tokenManager: ApprovalTokenManager
  ) {}

  // ---------------------------------------------------------------------------
  // Prepare Methods (Destructive — Phase 1)
  // ---------------------------------------------------------------------------

  async prepareDeleteEmail(params: PrepareDeleteEmailParams): Promise<{
    token_id: string;
    expires_at: string;
    email: ReturnType<typeof emailPreview>;
    action: string;
  }> {
    const email = await this.requireEmail(params.email_id);
    const hash = hashEmailForApproval(email);
    const token = this.tokenManager.generateToken({
      operation: 'delete_email',
      targetType: 'email',
      targetId: email.id,
      targetHash: hash,
    });

    return {
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      email: emailPreview(email),
      action: 'This email will be moved to the Deleted Items folder.',
    };
  }

  async prepareMoveEmail(params: PrepareMoveEmailParams): Promise<{
    token_id: string;
    expires_at: string;
    email: ReturnType<typeof emailPreview>;
    destination_folder: ReturnType<typeof folderPreview>;
    action: string;
  }> {
    const email = await this.requireEmail(params.email_id);
    const destFolder = await this.requireFolder(params.destination_folder_id);
    const hash = hashEmailForApproval(email);
    const token = this.tokenManager.generateToken({
      operation: 'move_email',
      targetType: 'email',
      targetId: email.id,
      targetHash: hash,
      metadata: { destinationFolderId: destFolder.id },
    });

    return {
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      email: emailPreview(email),
      destination_folder: folderPreview(destFolder),
      action: `This email will be moved to "${destFolder.name ?? 'Unnamed'}".`,
    };
  }

  async prepareArchiveEmail(params: PrepareArchiveEmailParams): Promise<{
    token_id: string;
    expires_at: string;
    email: ReturnType<typeof emailPreview>;
    action: string;
  }> {
    const email = await this.requireEmail(params.email_id);
    const hash = hashEmailForApproval(email);
    const token = this.tokenManager.generateToken({
      operation: 'archive_email',
      targetType: 'email',
      targetId: email.id,
      targetHash: hash,
    });

    return {
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      email: emailPreview(email),
      action: 'This email will be moved to the Archive folder.',
    };
  }

  async prepareJunkEmail(params: PrepareJunkEmailParams): Promise<{
    token_id: string;
    expires_at: string;
    email: ReturnType<typeof emailPreview>;
    action: string;
  }> {
    const email = await this.requireEmail(params.email_id);
    const hash = hashEmailForApproval(email);
    const token = this.tokenManager.generateToken({
      operation: 'junk_email',
      targetType: 'email',
      targetId: email.id,
      targetHash: hash,
    });

    return {
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      email: emailPreview(email),
      action: 'This email will be moved to the Junk folder.',
    };
  }

  async prepareDeleteFolder(params: PrepareDeleteFolderParams): Promise<{
    token_id: string;
    expires_at: string;
    folder: ReturnType<typeof folderPreview>;
    action: string;
  }> {
    const folder = await this.requireFolder(params.folder_id);
    const hash = hashFolderForApproval(folder);
    const token = this.tokenManager.generateToken({
      operation: 'delete_folder',
      targetType: 'folder',
      targetId: folder.id,
      targetHash: hash,
    });

    return {
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      folder: folderPreview(folder),
      action: `This folder and its ${folder.messageCount} messages will be deleted.`,
    };
  }

  async prepareEmptyFolder(params: PrepareEmptyFolderParams): Promise<{
    token_id: string;
    expires_at: string;
    folder: ReturnType<typeof folderPreview>;
    action: string;
  }> {
    const folder = await this.requireFolder(params.folder_id);
    const hash = hashFolderForApproval(folder);
    const token = this.tokenManager.generateToken({
      operation: 'empty_folder',
      targetType: 'folder',
      targetId: folder.id,
      targetHash: hash,
    });

    return {
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      folder: folderPreview(folder),
      action: `All ${folder.messageCount} messages in this folder will be deleted.`,
    };
  }

  async prepareBatchDeleteEmails(params: PrepareBatchDeleteEmailsParams): Promise<{
    tokens: Array<{ token_id: string; email: ReturnType<typeof emailPreview> }>;
    expires_at: string | null;
    action: string;
  }> {
    const tokens = [];
    for (const emailId of params.email_ids) {
      const email = await this.requireEmail(emailId);
      const hash = hashEmailForApproval(email);
      const token = this.tokenManager.generateToken({
        operation: 'batch_delete_emails',
        targetType: 'email',
        targetId: email.id,
        targetHash: hash,
      });
      tokens.push({
        token_id: token.tokenId,
        email: emailPreview(email),
      });
    }

    const firstToken = this.tokenManager.validateToken(
      tokens[0]!.token_id,
      'batch_delete_emails',
      params.email_ids[0]!
    );

    return {
      tokens,
      expires_at: firstToken.token != null
        ? new Date(firstToken.token.expiresAt).toISOString()
        : null,
      action: `${tokens.length} emails will be moved to the Deleted Items folder. You may selectively confirm by omitting tokens.`,
    };
  }

  async prepareBatchMoveEmails(params: PrepareBatchMoveEmailsParams): Promise<{
    tokens: Array<{ token_id: string; email: ReturnType<typeof emailPreview> }>;
    destination_folder: ReturnType<typeof folderPreview>;
    expires_at: string | null;
    action: string;
  }> {
    const destFolder = await this.requireFolder(params.destination_folder_id);

    const tokens = [];
    for (const emailId of params.email_ids) {
      const email = await this.requireEmail(emailId);
      const hash = hashEmailForApproval(email);
      const token = this.tokenManager.generateToken({
        operation: 'batch_move_emails',
        targetType: 'email',
        targetId: email.id,
        targetHash: hash,
        metadata: { destinationFolderId: destFolder.id },
      });
      tokens.push({
        token_id: token.tokenId,
        email: emailPreview(email),
      });
    }

    const firstToken = this.tokenManager.validateToken(
      tokens[0]!.token_id,
      'batch_move_emails',
      params.email_ids[0]!
    );

    return {
      tokens,
      destination_folder: folderPreview(destFolder),
      expires_at: firstToken.token != null
        ? new Date(firstToken.token.expiresAt).toISOString()
        : null,
      action: `${tokens.length} emails will be moved to "${destFolder.name ?? 'Unnamed'}". You may selectively confirm by omitting tokens.`,
    };
  }

  // ---------------------------------------------------------------------------
  // Confirm Methods (Destructive — Phase 2)
  // ---------------------------------------------------------------------------

  async confirmDeleteEmail(params: ConfirmDeleteEmailParams): Promise<{ success: boolean; message: string }> {
    await this.consumeAndVerifyEmail(params.token_id, 'delete_email', params.email_id);
    await this.repository.deleteEmail(params.email_id);
    return { success: true, message: 'Email moved to Deleted Items.' };
  }

  async confirmMoveEmail(params: ConfirmMoveEmailParams): Promise<{ success: boolean; message: string }> {
    // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
    const token = await this.consumeAndVerifyEmail(params.token_id, 'move_email', params.email_id);
    // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
    const destFolderId = token.metadata['destinationFolderId'] as number;
    await this.repository.moveEmail(params.email_id, destFolderId);
    return { success: true, message: 'Email moved successfully.' };
  }

  async confirmArchiveEmail(params: ConfirmArchiveEmailParams): Promise<{ success: boolean; message: string }> {
    await this.consumeAndVerifyEmail(params.token_id, 'archive_email', params.email_id);
    await this.repository.archiveEmail(params.email_id);
    return { success: true, message: 'Email moved to Archive.' };
  }

  async confirmJunkEmail(params: ConfirmJunkEmailParams): Promise<{ success: boolean; message: string }> {
    await this.consumeAndVerifyEmail(params.token_id, 'junk_email', params.email_id);
    await this.repository.junkEmail(params.email_id);
    return { success: true, message: 'Email moved to Junk.' };
  }

  async confirmDeleteFolder(params: ConfirmDeleteFolderParams): Promise<{ success: boolean; message: string }> {
    await this.consumeAndVerifyFolder(params.token_id, 'delete_folder', params.folder_id);
    await this.repository.deleteFolder(params.folder_id);
    return { success: true, message: 'Folder deleted.' };
  }

  async confirmEmptyFolder(params: ConfirmEmptyFolderParams): Promise<{ success: boolean; message: string }> {
    await this.consumeAndVerifyFolder(params.token_id, 'empty_folder', params.folder_id);
    await this.repository.emptyFolder(params.folder_id);
    return { success: true, message: 'Folder emptied.' };
  }

  async confirmBatchOperation(params: ConfirmBatchOperationParams): Promise<{
    results: Array<{ email_id: number; success: true } | { email_id: number; success: false; error: string }>;
    summary: { total: number; succeeded: number; failed: number };
  }> {
    const results = [];
    for (const { token_id, email_id } of params.tokens) {
      try {
        // Peek at the token to determine the operation type
        const peekResult = this.tokenManager.validateToken(token_id, 'batch_delete_emails', email_id);

        let operation: OperationType;
        if (peekResult.valid) {
          operation = 'batch_delete_emails';
        } else if (peekResult.error === 'OPERATION_MISMATCH') {
          // Try the other batch operation
          const moveResult = this.tokenManager.validateToken(token_id, 'batch_move_emails', email_id);
          if (!moveResult.valid) {
            throwValidationError(moveResult.error!);
          }
          operation = 'batch_move_emails';
        } else {
          throwValidationError(peekResult.error!);
          // unreachable, but satisfies typescript
          operation = 'batch_delete_emails';
        }

        // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
        const token = await this.consumeAndVerifyEmail(token_id, operation, email_id);

        if (operation === 'batch_delete_emails') {
          await this.repository.deleteEmail(email_id);
        } else {
          // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
          const destFolderId = token.metadata['destinationFolderId'] as number;
          await this.repository.moveEmail(email_id, destFolderId);
        }

        results.push({ email_id, success: true as const });
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Unknown error';
        results.push({ email_id, success: false as const, error: message });
      }
    }

    const succeeded = results.filter((r) => r.success).length;
    const failed = results.filter((r) => !r.success).length;

    return {
      results,
      summary: { total: results.length, succeeded, failed },
    };
  }

  // ---------------------------------------------------------------------------
  // Low-Risk Modifications (Single Tool)
  // ---------------------------------------------------------------------------

  async markEmailRead(params: MarkEmailReadParams): Promise<{ success: boolean; message: string }> {
    await this.requireEmail(params.email_id);
    await this.repository.markEmailRead(params.email_id, true);
    return { success: true, message: 'Email marked as read.' };
  }

  async markEmailUnread(params: MarkEmailUnreadParams): Promise<{ success: boolean; message: string }> {
    await this.requireEmail(params.email_id);
    await this.repository.markEmailRead(params.email_id, false);
    return { success: true, message: 'Email marked as unread.' };
  }

  async setEmailFlag(params: SetEmailFlagParams): Promise<{ success: boolean; message: string }> {
    await this.requireEmail(params.email_id);
    await this.repository.setEmailFlag(params.email_id, params.flag_status);
    return { success: true, message: 'Email flag updated.' };
  }

  async clearEmailFlag(params: ClearEmailFlagParams): Promise<{ success: boolean; message: string }> {
    await this.requireEmail(params.email_id);
    await this.repository.setEmailFlag(params.email_id, 0);
    return { success: true, message: 'Email flag cleared.' };
  }

  async setEmailCategories(params: SetEmailCategoriesParams): Promise<{ success: boolean; message: string }> {
    await this.requireEmail(params.email_id);
    await this.repository.setEmailCategories(params.email_id, params.categories);
    return { success: true, message: 'Email categories updated.' };
  }

  // ---------------------------------------------------------------------------
  // Non-Destructive Operations
  // ---------------------------------------------------------------------------

  async createFolder(params: CreateFolderParams): Promise<{ success: boolean; folder: ReturnType<typeof folderPreview> }> {
    const newFolder = await this.repository.createFolder(params.name, params.parent_folder_id);
    return { success: true, folder: folderPreview(newFolder) };
  }

  async renameFolder(params: RenameFolderParams): Promise<{ success: boolean; message: string }> {
    await this.requireFolder(params.folder_id);
    await this.repository.renameFolder(params.folder_id, params.new_name);
    return { success: true, message: `Folder renamed to "${params.new_name}".` };
  }

  async moveFolder(params: MoveFolderParams): Promise<{ success: boolean; message: string }> {
    await this.requireFolder(params.folder_id);
    await this.requireFolder(params.destination_parent_id);
    await this.repository.moveFolder(params.folder_id, params.destination_parent_id);
    return { success: true, message: 'Folder moved.' };
  }

  // ---------------------------------------------------------------------------
  // Private Helpers
  // ---------------------------------------------------------------------------

  private async requireEmail(emailId: number): Promise<EmailRow> {
    const email = await this.repository.getEmail(emailId);
    if (email == null) {
      throw new NotFoundError('Email', emailId);
    }
    return email;
  }

  private async requireFolder(folderId: number): Promise<FolderRow> {
    const folder = await this.repository.getFolder(folderId);
    if (folder == null) {
      throw new NotFoundError('Folder', folderId);
    }
    return folder;
  }

  /**
   * Consumes a token and verifies the email hasn't changed.
   * Returns the consumed token for metadata access.
   */
  private async consumeAndVerifyEmail(
    tokenId: string,
    operation: OperationType,
    emailId: number
  ): Promise<ApprovalToken> {
    const result = this.tokenManager.consumeToken(tokenId, operation, emailId);
    if (!result.valid) {
      throwValidationError(result.error!);
    }

    const token = result.token!;
    const email = await this.requireEmail(emailId);
    const currentHash = hashEmailForApproval(email);

    if (currentHash !== token.targetHash) {
      throw new TargetChangedError();
    }

    return token;
  }

  /**
   * Consumes a token and verifies the folder hasn't changed.
   */
  private async consumeAndVerifyFolder(
    tokenId: string,
    operation: OperationType,
    folderId: number
  ): Promise<void> {
    const result = this.tokenManager.consumeToken(tokenId, operation, folderId);
    if (!result.valid) {
      throwValidationError(result.error!);
    }

    const token = result.token!;
    const folder = await this.requireFolder(folderId);
    const currentHash = hashFolderForApproval(folder);

    if (currentHash !== token.targetHash) {
      throw new TargetChangedError();
    }
  }
}

/**
 * Creates mailbox organization tools with the given repository and token manager.
 */
export function createMailboxOrganizationTools(
  repository: IMailboxRepository,
  tokenManager: ApprovalTokenManager
): MailboxOrganizationTools {
  return new MailboxOrganizationTools(repository, tokenManager);
}
