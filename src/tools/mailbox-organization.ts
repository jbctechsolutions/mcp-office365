/**
 * Mailbox organization MCP tools.
 *
 * Provides tools for organizing emails and folders with a two-phase
 * approval pattern for destructive operations (prepare → confirm).
 */

import { z } from 'zod';
import type { IWriteableRepository, EmailRow, FolderRow } from '../database/repository.js';
import {
  ApprovalTokenManager,
  hashEmailForApproval,
  hashFolderForApproval,
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

function emailPreview(row: EmailRow) {
  return {
    id: row.id,
    subject: row.subject,
    sender: row.sender,
    senderAddress: row.senderAddress,
    folderId: row.folderId,
    timeReceived: row.timeReceived != null ? appleTimestampToIso(row.timeReceived) : null,
  };
}

function folderPreview(row: FolderRow) {
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
 * Used in AppleScript (sync) mode. Graph API mode handles these
 * operations directly in the tool call handler with async methods.
 */
export class MailboxOrganizationTools {
  constructor(
    private readonly repository: IWriteableRepository,
    private readonly tokenManager: ApprovalTokenManager
  ) {}

  // ---------------------------------------------------------------------------
  // Prepare Methods (Destructive — Phase 1)
  // ---------------------------------------------------------------------------

  prepareDeleteEmail(params: PrepareDeleteEmailParams) {
    const email = this.requireEmail(params.email_id);
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

  prepareMoveEmail(params: PrepareMoveEmailParams) {
    const email = this.requireEmail(params.email_id);
    const destFolder = this.requireFolder(params.destination_folder_id);
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

  prepareArchiveEmail(params: PrepareArchiveEmailParams) {
    const email = this.requireEmail(params.email_id);
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

  prepareJunkEmail(params: PrepareJunkEmailParams) {
    const email = this.requireEmail(params.email_id);
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

  prepareDeleteFolder(params: PrepareDeleteFolderParams) {
    const folder = this.requireFolder(params.folder_id);
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

  prepareEmptyFolder(params: PrepareEmptyFolderParams) {
    const folder = this.requireFolder(params.folder_id);
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

  prepareBatchDeleteEmails(params: PrepareBatchDeleteEmailsParams) {
    const tokens = params.email_ids.map((emailId) => {
      const email = this.requireEmail(emailId);
      const hash = hashEmailForApproval(email);
      const token = this.tokenManager.generateToken({
        operation: 'batch_delete_emails',
        targetType: 'email',
        targetId: email.id,
        targetHash: hash,
      });
      return {
        token_id: token.tokenId,
        email: emailPreview(email),
      };
    });

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

  prepareBatchMoveEmails(params: PrepareBatchMoveEmailsParams) {
    const destFolder = this.requireFolder(params.destination_folder_id);

    const tokens = params.email_ids.map((emailId) => {
      const email = this.requireEmail(emailId);
      const hash = hashEmailForApproval(email);
      const token = this.tokenManager.generateToken({
        operation: 'batch_move_emails',
        targetType: 'email',
        targetId: email.id,
        targetHash: hash,
        metadata: { destinationFolderId: destFolder.id },
      });
      return {
        token_id: token.tokenId,
        email: emailPreview(email),
      };
    });

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

  confirmDeleteEmail(params: ConfirmDeleteEmailParams) {
    this.consumeAndVerifyEmail(params.token_id, 'delete_email', params.email_id);
    this.repository.deleteEmail(params.email_id);
    return { success: true, message: 'Email moved to Deleted Items.' };
  }

  confirmMoveEmail(params: ConfirmMoveEmailParams) {
    const token = this.consumeAndVerifyEmail(params.token_id, 'move_email', params.email_id);
    const destFolderId = token.metadata['destinationFolderId'] as number;
    this.repository.moveEmail(params.email_id, destFolderId);
    return { success: true, message: 'Email moved successfully.' };
  }

  confirmArchiveEmail(params: ConfirmArchiveEmailParams) {
    this.consumeAndVerifyEmail(params.token_id, 'archive_email', params.email_id);
    this.repository.archiveEmail(params.email_id);
    return { success: true, message: 'Email moved to Archive.' };
  }

  confirmJunkEmail(params: ConfirmJunkEmailParams) {
    this.consumeAndVerifyEmail(params.token_id, 'junk_email', params.email_id);
    this.repository.junkEmail(params.email_id);
    return { success: true, message: 'Email moved to Junk.' };
  }

  confirmDeleteFolder(params: ConfirmDeleteFolderParams) {
    this.consumeAndVerifyFolder(params.token_id, 'delete_folder', params.folder_id);
    this.repository.deleteFolder(params.folder_id);
    return { success: true, message: 'Folder deleted.' };
  }

  confirmEmptyFolder(params: ConfirmEmptyFolderParams) {
    this.consumeAndVerifyFolder(params.token_id, 'empty_folder', params.folder_id);
    this.repository.emptyFolder(params.folder_id);
    return { success: true, message: 'Folder emptied.' };
  }

  confirmBatchOperation(params: ConfirmBatchOperationParams) {
    const results = params.tokens.map(({ token_id, email_id }) => {
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

        const token = this.consumeAndVerifyEmail(token_id, operation, email_id);

        if (operation === 'batch_delete_emails') {
          this.repository.deleteEmail(email_id);
        } else {
          const destFolderId = token.metadata['destinationFolderId'] as number;
          this.repository.moveEmail(email_id, destFolderId);
        }

        return { email_id, success: true as const };
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Unknown error';
        return { email_id, success: false as const, error: message };
      }
    });

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

  markEmailRead(params: MarkEmailReadParams) {
    this.requireEmail(params.email_id);
    this.repository.markEmailRead(params.email_id, true);
    return { success: true, message: 'Email marked as read.' };
  }

  markEmailUnread(params: MarkEmailUnreadParams) {
    this.requireEmail(params.email_id);
    this.repository.markEmailRead(params.email_id, false);
    return { success: true, message: 'Email marked as unread.' };
  }

  setEmailFlag(params: SetEmailFlagParams) {
    this.requireEmail(params.email_id);
    this.repository.setEmailFlag(params.email_id, params.flag_status);
    return { success: true, message: 'Email flag updated.' };
  }

  clearEmailFlag(params: ClearEmailFlagParams) {
    this.requireEmail(params.email_id);
    this.repository.setEmailFlag(params.email_id, 0);
    return { success: true, message: 'Email flag cleared.' };
  }

  setEmailCategories(params: SetEmailCategoriesParams) {
    this.requireEmail(params.email_id);
    this.repository.setEmailCategories(params.email_id, params.categories);
    return { success: true, message: 'Email categories updated.' };
  }

  // ---------------------------------------------------------------------------
  // Non-Destructive Operations
  // ---------------------------------------------------------------------------

  createFolder(params: CreateFolderParams) {
    const newFolder = this.repository.createFolder(params.name, params.parent_folder_id);
    return { success: true, folder: folderPreview(newFolder) };
  }

  renameFolder(params: RenameFolderParams) {
    this.requireFolder(params.folder_id);
    this.repository.renameFolder(params.folder_id, params.new_name);
    return { success: true, message: `Folder renamed to "${params.new_name}".` };
  }

  moveFolder(params: MoveFolderParams) {
    this.requireFolder(params.folder_id);
    this.requireFolder(params.destination_parent_id);
    this.repository.moveFolder(params.folder_id, params.destination_parent_id);
    return { success: true, message: 'Folder moved.' };
  }

  // ---------------------------------------------------------------------------
  // Private Helpers
  // ---------------------------------------------------------------------------

  private requireEmail(emailId: number): EmailRow {
    const email = this.repository.getEmail(emailId);
    if (email == null) {
      throw new NotFoundError('Email', emailId);
    }
    return email;
  }

  private requireFolder(folderId: number): FolderRow {
    const folder = this.repository.getFolder(folderId);
    if (folder == null) {
      throw new NotFoundError('Folder', folderId);
    }
    return folder;
  }

  /**
   * Consumes a token and verifies the email hasn't changed.
   * Returns the consumed token for metadata access.
   */
  private consumeAndVerifyEmail(
    tokenId: string,
    operation: OperationType,
    emailId: number
  ) {
    const result = this.tokenManager.consumeToken(tokenId, operation, emailId);
    if (!result.valid) {
      throwValidationError(result.error!);
    }

    const token = result.token!;
    const email = this.requireEmail(emailId);
    const currentHash = hashEmailForApproval(email);

    if (currentHash !== token.targetHash) {
      throw new TargetChangedError();
    }

    return token;
  }

  /**
   * Consumes a token and verifies the folder hasn't changed.
   */
  private consumeAndVerifyFolder(
    tokenId: string,
    operation: OperationType,
    folderId: number
  ): void {
    const result = this.tokenManager.consumeToken(tokenId, operation, folderId);
    if (!result.valid) {
      throwValidationError(result.error!);
    }

    const token = result.token!;
    const folder = this.requireFolder(folderId);
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
  repository: IWriteableRepository,
  tokenManager: ApprovalTokenManager
): MailboxOrganizationTools {
  return new MailboxOrganizationTools(repository, tokenManager);
}
