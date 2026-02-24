/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Mail send MCP tools.
 *
 * Provides tools for email draft management and two-phase approval
 * for all send operations (send draft, send email, reply, forward).
 */

import { z } from 'zod';
import type { EmailRow } from '../database/repository.js';
import {
  ApprovalTokenManager,
  hashDraftForSend,
  hashDirectSendForApproval,
  hashReplyForApproval,
  hashForwardForApproval,
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

// =============================================================================
// Repository Interface
// =============================================================================

/**
 * Async-compatible repository interface for mail send tools.
 *
 * Matches the async methods on GraphRepository.
 */
export interface IMailSendRepository {
  getEmailAsync(id: number): Promise<EmailRow | undefined>;
  createDraftAsync(params: {
    subject: string;
    body: string;
    bodyType: string;
    to?: string[];
    cc?: string[];
    bcc?: string[];
  }): Promise<number>;
  updateDraftAsync(draftId: number, updates: Record<string, unknown>): Promise<void>;
  listDraftsAsync(limit: number, offset: number): Promise<EmailRow[]>;
  sendDraftAsync(draftId: number): Promise<void>;
  sendMailAsync(params: {
    subject: string;
    body: string;
    bodyType: string;
    to: string[];
    cc?: string[];
    bcc?: string[];
  }): Promise<void>;
  replyMessageAsync(messageId: number, comment: string, replyAll: boolean): Promise<void>;
  forwardMessageAsync(messageId: number, toRecipients: string[], comment?: string): Promise<void>;
}

// =============================================================================
// Input Schemas -- Non-Destructive Operations
// =============================================================================

export const CreateDraftInput = z.strictObject({
  to: z.array(z.string().email()).optional().describe('To recipients'),
  cc: z.array(z.string().email()).optional().describe('CC recipients'),
  bcc: z.array(z.string().email()).optional().describe('BCC recipients'),
  subject: z.string().describe('Email subject'),
  body: z.string().describe('Email body'),
  body_type: z.enum(['text', 'html']).default('text').describe('Body content type'),
});

export const UpdateDraftInput = z.strictObject({
  draft_id: z.number().int().positive().describe('The draft ID to update'),
  to: z.array(z.string().email()).optional().describe('To recipients'),
  cc: z.array(z.string().email()).optional().describe('CC recipients'),
  bcc: z.array(z.string().email()).optional().describe('BCC recipients'),
  subject: z.string().optional().describe('Email subject'),
  body: z.string().optional().describe('Email body'),
  body_type: z.enum(['text', 'html']).optional().describe('Body content type'),
});

export const ListDraftsInput = z.strictObject({
  limit: z.number().int().min(1).max(100).default(50).describe('Maximum drafts to return'),
  offset: z.number().int().min(0).default(0).describe('Number to skip'),
});

// =============================================================================
// Input Schemas -- Destructive Operations (Two-Phase)
// =============================================================================

// Send Draft
export const PrepareSendDraftInput = z.strictObject({
  draft_id: z.number().int().positive().describe('The draft ID to send'),
});

export const ConfirmSendDraftInput = z.strictObject({
  token_id: z.uuid().describe('Approval token from prepare_send_draft'),
  draft_id: z.number().int().positive().describe('The draft ID to send'),
});

// Send Email (direct)
export const PrepareSendEmailInput = z.strictObject({
  to: z.array(z.string().email()).min(1).describe('To recipients'),
  cc: z.array(z.string().email()).optional().describe('CC recipients'),
  bcc: z.array(z.string().email()).optional().describe('BCC recipients'),
  subject: z.string().describe('Email subject'),
  body: z.string().describe('Email body'),
  body_type: z.enum(['text', 'html']).default('text').describe('Body content type'),
});

export const ConfirmSendEmailInput = z.strictObject({
  token_id: z.uuid().describe('Approval token from prepare_send_email'),
});

// Reply Email
export const PrepareReplyEmailInput = z.strictObject({
  message_id: z.number().int().positive().describe('The message ID to reply to'),
  comment: z.string().describe('Reply body'),
  reply_all: z.boolean().default(true).describe('Reply to all recipients (default true)'),
});

export const ConfirmReplyEmailInput = z.strictObject({
  token_id: z.uuid().describe('Approval token from prepare_reply_email'),
  message_id: z.number().int().positive().describe('The message ID being replied to'),
});

// Forward Email
export const PrepareForwardEmailInput = z.strictObject({
  message_id: z.number().int().positive().describe('The message ID to forward'),
  to_recipients: z.array(z.string().email()).min(1).describe('Forward to recipients'),
  comment: z.string().optional().describe('Optional comment to include'),
});

export const ConfirmForwardEmailInput = z.strictObject({
  token_id: z.uuid().describe('Approval token from prepare_forward_email'),
  message_id: z.number().int().positive().describe('The message ID being forwarded'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type CreateDraftParams = z.infer<typeof CreateDraftInput>;
export type UpdateDraftParams = z.infer<typeof UpdateDraftInput>;
export type ListDraftsParams = z.infer<typeof ListDraftsInput>;
export type PrepareSendDraftParams = z.infer<typeof PrepareSendDraftInput>;
export type ConfirmSendDraftParams = z.infer<typeof ConfirmSendDraftInput>;
export type PrepareSendEmailParams = z.infer<typeof PrepareSendEmailInput>;
export type ConfirmSendEmailParams = z.infer<typeof ConfirmSendEmailInput>;
export type PrepareReplyEmailParams = z.infer<typeof PrepareReplyEmailInput>;
export type ConfirmReplyEmailParams = z.infer<typeof ConfirmReplyEmailInput>;
export type PrepareForwardEmailParams = z.infer<typeof PrepareForwardEmailInput>;
export type ConfirmForwardEmailParams = z.infer<typeof ConfirmForwardEmailInput>;

// =============================================================================
// Preview Helpers
// =============================================================================

function draftPreview(row: EmailRow): {
  id: number;
  subject: string | null;
  to: string | null;
  cc: string | null;
} {
  return {
    id: row.id,
    subject: row.subject,
    to: row.toAddresses ?? row.displayTo,
    cc: row.ccAddresses,
  };
}

function messagePreview(row: EmailRow): {
  id: number;
  subject: string | null;
  sender: string | null;
  senderAddress: string | null;
} {
  return {
    id: row.id,
    subject: row.subject,
    sender: row.sender,
    senderAddress: row.senderAddress,
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

/**
 * Counts recipients from a semicolon-separated address string.
 */
function countRecipients(addresses: string | null | undefined): number {
  if (addresses == null || addresses.trim() === '') return 0;
  return addresses.split(';').filter((a) => a.trim() !== '').length;
}

// =============================================================================
// Mail Send Tools
// =============================================================================

/**
 * Mail send tools with two-phase approval for send operations.
 *
 * Provides draft management (create, update, list) as non-destructive ops,
 * and send/reply/forward as two-phase approval operations.
 */
export class MailSendTools {
  constructor(
    private readonly repository: IMailSendRepository,
    private readonly tokenManager: ApprovalTokenManager
  ) {}

  // ---------------------------------------------------------------------------
  // Non-Destructive Operations
  // ---------------------------------------------------------------------------

  async createDraft(params: CreateDraftParams): Promise<{ success: boolean; draft_id: number }> {
    const draftId = await this.repository.createDraftAsync({
      subject: params.subject,
      body: params.body,
      bodyType: params.body_type,
      ...(params.to != null ? { to: params.to } : {}),
      ...(params.cc != null ? { cc: params.cc } : {}),
      ...(params.bcc != null ? { bcc: params.bcc } : {}),
    });
    return { success: true, draft_id: draftId };
  }

  async updateDraft(params: UpdateDraftParams): Promise<{ success: boolean; message: string }> {
    const updates: Record<string, unknown> = {};
    if (params.to !== undefined) updates['to'] = params.to;
    if (params.cc !== undefined) updates['cc'] = params.cc;
    if (params.bcc !== undefined) updates['bcc'] = params.bcc;
    if (params.subject !== undefined) updates['subject'] = params.subject;
    if (params.body !== undefined) updates['body'] = params.body;
    if (params.body_type !== undefined) updates['bodyType'] = params.body_type;

    await this.repository.updateDraftAsync(params.draft_id, updates);
    return { success: true, message: 'Draft updated.' };
  }

  async listDrafts(params: ListDraftsParams): Promise<EmailRow[]> {
    return this.repository.listDraftsAsync(params.limit, params.offset);
  }

  // ---------------------------------------------------------------------------
  // Prepare Methods (Destructive -- Phase 1)
  // ---------------------------------------------------------------------------

  async prepareSendDraft(params: PrepareSendDraftParams): Promise<{
    token_id: string;
    expires_at: string;
    draft: ReturnType<typeof draftPreview>;
    action: string;
  }> {
    const draft = await this.requireEmail(params.draft_id);
    const recipientCount =
      countRecipients(draft.toAddresses) +
      countRecipients(draft.ccAddresses);

    const hash = hashDraftForSend({
      id: draft.id,
      subject: draft.subject,
      recipientCount,
    });

    const token = this.tokenManager.generateToken({
      operation: 'send_draft',
      targetType: 'email',
      targetId: draft.id,
      targetHash: hash,
    });

    return {
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      draft: draftPreview(draft),
      action: `This draft will be sent: "${draft.subject ?? '(no subject)'}".`,
    };
  }

  async prepareSendEmail(params: PrepareSendEmailParams): Promise<{
    token_id: string;
    expires_at: string;
    preview: {
      subject: string;
      to: string[];
      cc?: string[];
      bcc?: string[];
    };
    action: string;
  }> {
    const hash = hashDirectSendForApproval({
      subject: params.subject,
      toCount: params.to.length,
      ccCount: params.cc?.length ?? 0,
      bccCount: params.bcc?.length ?? 0,
    });

    // Use 0 as targetId since there's no pre-existing entity
    const token = this.tokenManager.generateToken({
      operation: 'send_email',
      targetType: 'email',
      targetId: 0,
      targetHash: hash,
      metadata: {
        subject: params.subject,
        body: params.body,
        bodyType: params.body_type,
        to: params.to,
        cc: params.cc,
        bcc: params.bcc,
      },
    });

    return {
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      preview: {
        subject: params.subject,
        to: params.to,
        ...(params.cc != null ? { cc: params.cc } : {}),
        ...(params.bcc != null ? { bcc: params.bcc } : {}),
      },
      action: `A new email will be sent: "${params.subject}" to ${params.to.length} recipient(s).`,
    };
  }

  async prepareReplyEmail(params: PrepareReplyEmailParams): Promise<{
    token_id: string;
    expires_at: string;
    original_message: ReturnType<typeof messagePreview>;
    action: string;
  }> {
    const original = await this.requireEmail(params.message_id);

    const hash = hashReplyForApproval({
      originalId: original.id,
      commentLength: params.comment.length,
      replyAll: params.reply_all,
    });

    const token = this.tokenManager.generateToken({
      operation: 'reply_email',
      targetType: 'email',
      targetId: original.id,
      targetHash: hash,
      metadata: {
        comment: params.comment,
        replyAll: params.reply_all,
      },
    });

    const replyType = params.reply_all ? 'reply all' : 'reply';
    return {
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      original_message: messagePreview(original),
      action: `A ${replyType} will be sent to "${original.subject ?? '(no subject)'}".`,
    };
  }

  async prepareForwardEmail(params: PrepareForwardEmailParams): Promise<{
    token_id: string;
    expires_at: string;
    original_message: ReturnType<typeof messagePreview>;
    action: string;
  }> {
    const original = await this.requireEmail(params.message_id);

    const hash = hashForwardForApproval({
      originalId: original.id,
      recipientCount: params.to_recipients.length,
    });

    const token = this.tokenManager.generateToken({
      operation: 'forward_email',
      targetType: 'email',
      targetId: original.id,
      targetHash: hash,
      metadata: {
        toRecipients: params.to_recipients,
        comment: params.comment,
      },
    });

    return {
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      original_message: messagePreview(original),
      action: `This email will be forwarded to ${params.to_recipients.length} recipient(s): "${original.subject ?? '(no subject)'}".`,
    };
  }

  // ---------------------------------------------------------------------------
  // Confirm Methods (Destructive -- Phase 2)
  // ---------------------------------------------------------------------------

  async confirmSendDraft(params: ConfirmSendDraftParams): Promise<{ success: boolean; message: string }> {
    await this.consumeAndVerifyDraft(params.token_id, 'send_draft', params.draft_id);
    await this.repository.sendDraftAsync(params.draft_id);
    return { success: true, message: 'Draft sent successfully.' };
  }

  async confirmSendEmail(params: ConfirmSendEmailParams): Promise<{ success: boolean; message: string }> {
    const token = this.consumeTokenForSendEmail(params.token_id);

    // Read the full send params from token metadata
    const metadata = token.metadata as {
      subject: string;
      body: string;
      bodyType: string;
      to: string[];
      cc?: string[];
      bcc?: string[];
    };

    await this.repository.sendMailAsync({
      subject: metadata.subject,
      body: metadata.body,
      bodyType: metadata.bodyType,
      to: metadata.to,
      ...(metadata.cc != null ? { cc: metadata.cc } : {}),
      ...(metadata.bcc != null ? { bcc: metadata.bcc } : {}),
    });

    return { success: true, message: 'Email sent successfully.' };
  }

  async confirmReplyEmail(params: ConfirmReplyEmailParams): Promise<{ success: boolean; message: string }> {
    const token = this.consumeTokenForMessage(params.token_id, 'reply_email', params.message_id);

    const metadata = token.metadata as {
      comment: string;
      replyAll: boolean;
    };

    await this.repository.replyMessageAsync(params.message_id, metadata.comment, metadata.replyAll);
    return { success: true, message: 'Reply sent successfully.' };
  }

  async confirmForwardEmail(params: ConfirmForwardEmailParams): Promise<{ success: boolean; message: string }> {
    const token = this.consumeTokenForMessage(params.token_id, 'forward_email', params.message_id);

    const metadata = token.metadata as {
      toRecipients: string[];
      comment?: string;
    };

    await this.repository.forwardMessageAsync(params.message_id, metadata.toRecipients, metadata.comment);
    return { success: true, message: 'Email forwarded successfully.' };
  }

  // ---------------------------------------------------------------------------
  // Private Helpers
  // ---------------------------------------------------------------------------

  private async requireEmail(emailId: number): Promise<EmailRow> {
    const email = await this.repository.getEmailAsync(emailId);
    if (email == null) {
      throw new NotFoundError('Email', emailId);
    }
    return email;
  }

  /**
   * Consumes a token and verifies the draft hasn't changed.
   * Returns the consumed token for metadata access.
   */
  private async consumeAndVerifyDraft(
    tokenId: string,
    operation: OperationType,
    draftId: number
  ): Promise<ApprovalToken> {
    const result = this.tokenManager.consumeToken(tokenId, operation, draftId);
    if (!result.valid) {
      throwValidationError(result.error!);
    }

    const token = result.token!;
    const draft = await this.requireEmail(draftId);

    const recipientCount =
      countRecipients(draft.toAddresses) +
      countRecipients(draft.ccAddresses);

    const currentHash = hashDraftForSend({
      id: draft.id,
      subject: draft.subject,
      recipientCount,
    });

    if (currentHash !== token.targetHash) {
      throw new TargetChangedError();
    }

    return token;
  }

  /**
   * Consumes a token for send_email (no pre-existing entity to verify).
   * The targetId is 0, so we use that for validation.
   */
  private consumeTokenForSendEmail(tokenId: string): ApprovalToken {
    const result = this.tokenManager.consumeToken(tokenId, 'send_email', 0);
    if (!result.valid) {
      throwValidationError(result.error!);
    }
    return result.token!;
  }

  /**
   * Consumes a token for reply/forward operations.
   * Validates token but does not re-hash the original message
   * (the original message content is not expected to change).
   */
  private consumeTokenForMessage(
    tokenId: string,
    operation: OperationType,
    messageId: number
  ): ApprovalToken {
    const result = this.tokenManager.consumeToken(tokenId, operation, messageId);
    if (!result.valid) {
      throwValidationError(result.error!);
    }
    return result.token!;
  }
}

/**
 * Creates mail send tools with the given repository and token manager.
 */
export function createMailSendTools(
  repository: IMailSendRepository,
  tokenManager: ApprovalTokenManager
): MailSendTools {
  return new MailSendTools(repository, tokenManager);
}
