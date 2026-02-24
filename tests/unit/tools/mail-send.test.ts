/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import {
  MailSendTools,
  createMailSendTools,
  type IMailSendRepository,
} from '../../../src/tools/mail-send.js';
import type { EmailRow } from '../../../src/database/repository.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';
import {
  NotFoundError,
  ApprovalExpiredError,
  ApprovalInvalidError,
  TargetChangedError,
} from '../../../src/utils/errors.js';

// Mock the attachments module
vi.mock('../../../src/graph/attachments.js', () => ({
  uploadAttachment: vi.fn().mockResolvedValue(undefined),
}));

import { uploadAttachment } from '../../../src/graph/attachments.js';

// Mock the signature module
const { mockReadSignature, mockWriteSignature, mockAppendSignature } = vi.hoisted(() => ({
  mockReadSignature: vi.fn().mockReturnValue(null),
  mockWriteSignature: vi.fn(),
  mockAppendSignature: vi.fn(),
}));

vi.mock('../../../src/signature.js', () => ({
  readSignature: mockReadSignature,
  writeSignature: mockWriteSignature,
  appendSignature: mockAppendSignature,
}));

import { readSignature, writeSignature, appendSignature } from '../../../src/signature.js';

// =============================================================================
// Test Fixtures
// =============================================================================

function makeEmailRow(overrides: Partial<EmailRow> = {}): EmailRow {
  return {
    id: 1,
    folderId: 10,
    subject: 'Test Email Subject',
    sender: 'Alice Sender',
    senderAddress: 'alice@example.com',
    preview: 'This is a preview of the email body...',
    isRead: 0,
    timeReceived: 700000000,
    timeSent: 699999000,
    hasAttachment: 0,
    size: 4096,
    priority: 0,
    flagStatus: 0,
    categories: null,
    messageId: '<msg-001@example.com>',
    conversationId: 100,
    dataFilePath: '/path/to/data.olk15',
    recipients: 'Bob Receiver',
    displayTo: 'Bob Receiver',
    toAddresses: 'bob@example.com',
    ccAddresses: null,
    ...overrides,
  };
}

// =============================================================================
// Mock Repository
// =============================================================================

const mockGraphClient = {
  sendDraft: vi.fn().mockResolvedValue(undefined),
  addAttachment: vi.fn().mockResolvedValue(undefined),
  createUploadSession: vi.fn(),
} as unknown as ReturnType<IMailSendRepository['getGraphClient']>;

function createMockRepository(): IMailSendRepository {
  return {
    getEmailAsync: vi.fn(),
    createDraftAsync: vi.fn(),
    updateDraftAsync: vi.fn(),
    listDraftsAsync: vi.fn(),
    sendDraftAsync: vi.fn(),
    sendMailAsync: vi.fn(),
    replyMessageAsync: vi.fn(),
    forwardMessageAsync: vi.fn(),
    replyAsDraftAsync: vi.fn(),
    forwardAsDraftAsync: vi.fn(),
    getGraphClient: vi.fn().mockReturnValue(mockGraphClient),
  };
}

// =============================================================================
// Tests
// =============================================================================

describe('MailSendTools', () => {
  let repo: ReturnType<typeof createMockRepository>;
  let tokenManager: ApprovalTokenManager;
  let tools: MailSendTools;

  const testEmail = makeEmailRow({ id: 1 });
  const testDraft = makeEmailRow({
    id: 5,
    folderId: 99,
    subject: 'Draft Subject',
    sender: null,
    senderAddress: null,
    recipients: 'bob@example.com; carol@example.com',
    displayTo: 'Bob; Carol',
    toAddresses: 'bob@example.com; carol@example.com',
  });

  beforeEach(() => {
    repo = createMockRepository();
    tokenManager = new ApprovalTokenManager();
    tools = new MailSendTools(repo, tokenManager);

    // Default mocks
    (repo.getEmailAsync as ReturnType<typeof vi.fn>).mockImplementation(async (id: number) => {
      if (id === 1) return testEmail;
      if (id === 5) return testDraft;
      return undefined;
    });

    // Reset attachment mocks
    vi.mocked(uploadAttachment).mockClear();
    vi.mocked(mockGraphClient.sendDraft).mockClear();
  });

  // ===========================================================================
  // Non-Destructive Operations
  // ===========================================================================

  describe('createDraft', () => {
    it('creates a draft and returns the draft_id', async () => {
      (repo.createDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({ numericId: 42, graphId: 'AAA-graph-42' });

      const result = await tools.createDraft({
        subject: 'Hello',
        body: 'World',
        body_type: 'text',
        to: ['bob@example.com'],
      });

      expect(result).toEqual({ success: true, draft_id: 42 });
      expect(repo.createDraftAsync).toHaveBeenCalledWith({
        subject: 'Hello',
        body: 'World',
        bodyType: 'text',
        to: ['bob@example.com'],
        cc: undefined,
        bcc: undefined,
      });
    });

    it('passes cc and bcc when provided', async () => {
      (repo.createDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({ numericId: 43, graphId: 'AAA-graph-43' });

      await tools.createDraft({
        subject: 'Hello',
        body: 'World',
        body_type: 'html',
        to: ['bob@example.com'],
        cc: ['carol@example.com'],
        bcc: ['dave@example.com'],
      });

      expect(repo.createDraftAsync).toHaveBeenCalledWith({
        subject: 'Hello',
        body: 'World',
        bodyType: 'html',
        to: ['bob@example.com'],
        cc: ['carol@example.com'],
        bcc: ['dave@example.com'],
      });
    });

    it('uploads attachments after creating draft when provided', async () => {
      (repo.createDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({ numericId: 44, graphId: 'AAA-graph-44' });
      vi.mocked(uploadAttachment).mockResolvedValue(undefined);

      const result = await tools.createDraft({
        subject: 'With Attachment',
        body: 'See attached',
        body_type: 'text',
        to: ['bob@example.com'],
        attachments: [
          { file_path: '/tmp/report.pdf' },
          { file_path: '/tmp/photo.jpg', name: 'vacation.jpg', content_type: 'image/jpeg' },
        ],
      });

      expect(result).toEqual({ success: true, draft_id: 44 });
      expect(uploadAttachment).toHaveBeenCalledTimes(2);
      expect(uploadAttachment).toHaveBeenCalledWith(
        mockGraphClient, 'AAA-graph-44', '/tmp/report.pdf', undefined, undefined
      );
      expect(uploadAttachment).toHaveBeenCalledWith(
        mockGraphClient, 'AAA-graph-44', '/tmp/photo.jpg', 'vacation.jpg', 'image/jpeg'
      );
    });

    it('does not call uploadAttachment when no attachments provided', async () => {
      (repo.createDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({ numericId: 45, graphId: 'AAA-graph-45' });
      vi.mocked(uploadAttachment).mockClear();

      await tools.createDraft({
        subject: 'No Attachment',
        body: 'Just text',
        body_type: 'text',
      });

      expect(uploadAttachment).not.toHaveBeenCalled();
    });
  });

  describe('updateDraft', () => {
    it('builds the correct updates object from non-undefined params', async () => {
      (repo.updateDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

      const result = await tools.updateDraft({
        draft_id: 5,
        subject: 'Updated Subject',
        body: 'Updated body',
      });

      expect(result).toEqual({ success: true, message: 'Draft updated.' });
      expect(repo.updateDraftAsync).toHaveBeenCalledWith(5, {
        subject: 'Updated Subject',
        body: 'Updated body',
      });
    });

    it('only includes defined fields in updates', async () => {
      (repo.updateDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

      await tools.updateDraft({
        draft_id: 5,
        to: ['new-recipient@example.com'],
      });

      expect(repo.updateDraftAsync).toHaveBeenCalledWith(5, {
        to: ['new-recipient@example.com'],
      });
    });

    it('includes body_type as bodyType in updates', async () => {
      (repo.updateDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

      await tools.updateDraft({
        draft_id: 5,
        body_type: 'html',
      });

      expect(repo.updateDraftAsync).toHaveBeenCalledWith(5, {
        bodyType: 'html',
      });
    });
  });

  describe('listDrafts', () => {
    it('delegates to repository with limit and offset', async () => {
      const drafts = [testDraft];
      (repo.listDraftsAsync as ReturnType<typeof vi.fn>).mockResolvedValue(drafts);

      const result = await tools.listDrafts({ limit: 50, offset: 0 });

      expect(result).toBe(drafts);
      expect(repo.listDraftsAsync).toHaveBeenCalledWith(50, 0);
    });
  });

  // ===========================================================================
  // Send Draft (Two-Phase)
  // ===========================================================================

  describe('prepareSendDraft / confirmSendDraft', () => {
    it('prepareSendDraft returns token and draft preview', async () => {
      const result = await tools.prepareSendDraft({ draft_id: 5 });

      expect(result.token_id).toBeDefined();
      expect(typeof result.token_id).toBe('string');
      expect(result.expires_at).toBeDefined();
      expect(result.draft.id).toBe(5);
      expect(result.draft.subject).toBe('Draft Subject');
      expect(result.action).toContain('sent');
    });

    it('confirmSendDraft validates token and calls sendDraftAsync', async () => {
      const prepared = await tools.prepareSendDraft({ draft_id: 5 });
      const result = await tools.confirmSendDraft({
        token_id: prepared.token_id,
        draft_id: 5,
      });

      expect(result.success).toBe(true);
      expect(result.message).toContain('sent');
      expect(repo.sendDraftAsync).toHaveBeenCalledWith(5);
    });

    it('prepareSendDraft throws NotFoundError for missing draft', async () => {
      await expect(tools.prepareSendDraft({ draft_id: 999 })).rejects.toThrow(NotFoundError);
    });

    it('confirmSendDraft throws if draft changed between prepare and confirm', async () => {
      const prepared = await tools.prepareSendDraft({ draft_id: 5 });

      // Simulate the draft changing after prepare
      const modifiedDraft = makeEmailRow({
        id: 5,
        subject: 'Modified Draft Subject',
        recipients: 'bob@example.com; carol@example.com; dave@example.com',
      });
      (repo.getEmailAsync as ReturnType<typeof vi.fn>).mockImplementation(async (id: number) => {
        if (id === 5) return modifiedDraft;
        return undefined;
      });

      await expect(
        tools.confirmSendDraft({ token_id: prepared.token_id, draft_id: 5 })
      ).rejects.toThrow(TargetChangedError);

      expect(repo.sendDraftAsync).not.toHaveBeenCalled();
    });
  });

  // ===========================================================================
  // Send Email (Two-Phase)
  // ===========================================================================

  describe('prepareSendEmail / confirmSendEmail', () => {
    const sendParams = {
      to: ['bob@example.com'] as [string, ...string[]],
      subject: 'Direct Send',
      body: 'Hello from direct send',
      body_type: 'text' as const,
    };

    it('prepareSendEmail returns token and preview of params', async () => {
      const result = await tools.prepareSendEmail(sendParams);

      expect(result.token_id).toBeDefined();
      expect(typeof result.token_id).toBe('string');
      expect(result.expires_at).toBeDefined();
      expect(result.preview.subject).toBe('Direct Send');
      expect(result.preview.to).toEqual(['bob@example.com']);
      expect(result.action).toContain('sent');
    });

    it('confirmSendEmail reads params from token metadata and calls sendMailAsync', async () => {
      const prepared = await tools.prepareSendEmail(sendParams);

      const result = await tools.confirmSendEmail({
        token_id: prepared.token_id,
      });

      expect(result.success).toBe(true);
      expect(result.message).toContain('sent');
      expect(repo.sendMailAsync).toHaveBeenCalledWith({
        subject: 'Direct Send',
        body: 'Hello from direct send',
        bodyType: 'text',
        to: ['bob@example.com'],
        cc: undefined,
        bcc: undefined,
      });
    });

    it('prepareSendEmail stores cc and bcc in metadata', async () => {
      const paramsWithCcBcc = {
        ...sendParams,
        cc: ['carol@example.com'],
        bcc: ['dave@example.com'],
      };

      const prepared = await tools.prepareSendEmail(paramsWithCcBcc);

      const result = await tools.confirmSendEmail({
        token_id: prepared.token_id,
      });

      expect(result.success).toBe(true);
      expect(repo.sendMailAsync).toHaveBeenCalledWith({
        subject: 'Direct Send',
        body: 'Hello from direct send',
        bodyType: 'text',
        to: ['bob@example.com'],
        cc: ['carol@example.com'],
        bcc: ['dave@example.com'],
      });
    });

    it('confirmSendEmail with attachments creates draft, uploads, and sends draft', async () => {
      (repo.createDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({ numericId: 50, graphId: 'AAA-graph-50' });
      vi.mocked(uploadAttachment).mockResolvedValue(undefined);

      const prepared = await tools.prepareSendEmail({
        ...sendParams,
        attachments: [
          { file_path: '/tmp/doc.pdf', name: 'document.pdf' },
        ],
      });

      const result = await tools.confirmSendEmail({
        token_id: prepared.token_id,
      });

      expect(result.success).toBe(true);
      expect(result.message).toContain('sent');

      // Should NOT call sendMailAsync (used the draft path instead)
      expect(repo.sendMailAsync).not.toHaveBeenCalled();

      // Should create a draft
      expect(repo.createDraftAsync).toHaveBeenCalledWith({
        subject: 'Direct Send',
        body: 'Hello from direct send',
        bodyType: 'text',
        to: ['bob@example.com'],
        cc: undefined,
        bcc: undefined,
      });

      // Should upload the attachment
      expect(uploadAttachment).toHaveBeenCalledWith(
        mockGraphClient, 'AAA-graph-50', '/tmp/doc.pdf', 'document.pdf', undefined
      );

      // Should send the draft via GraphClient
      expect(mockGraphClient.sendDraft).toHaveBeenCalledWith('AAA-graph-50');
    });

    it('confirmSendEmail without attachments calls sendMailAsync directly', async () => {
      const prepared = await tools.prepareSendEmail(sendParams);

      const result = await tools.confirmSendEmail({
        token_id: prepared.token_id,
      });

      expect(result.success).toBe(true);
      expect(repo.sendMailAsync).toHaveBeenCalled();
      expect(repo.createDraftAsync).not.toHaveBeenCalled();
      expect(uploadAttachment).not.toHaveBeenCalled();
      expect(mockGraphClient.sendDraft).not.toHaveBeenCalled();
    });
  });

  // ===========================================================================
  // Reply Email (Two-Phase)
  // ===========================================================================

  describe('prepareReplyEmail / confirmReplyEmail', () => {
    it('prepareReplyEmail fetches original message and returns preview', async () => {
      const result = await tools.prepareReplyEmail({
        message_id: 1,
        comment: 'Thanks for your email!',
        reply_all: true,
      });

      expect(result.token_id).toBeDefined();
      expect(result.expires_at).toBeDefined();
      expect(result.original_message.id).toBe(1);
      expect(result.original_message.subject).toBe('Test Email Subject');
      expect(result.original_message.sender).toBe('Alice Sender');
      expect(result.action).toContain('reply');
    });

    it('confirmReplyEmail validates token and calls replyMessageAsync', async () => {
      const prepared = await tools.prepareReplyEmail({
        message_id: 1,
        comment: 'Thanks for your email!',
        reply_all: true,
      });

      const result = await tools.confirmReplyEmail({
        token_id: prepared.token_id,
        message_id: 1,
      });

      expect(result.success).toBe(true);
      expect(result.message).toContain('Reply sent');
      expect(repo.replyMessageAsync).toHaveBeenCalledWith(1, 'Thanks for your email!', true);
    });

    it('prepareReplyEmail throws NotFoundError for missing message', async () => {
      await expect(
        tools.prepareReplyEmail({ message_id: 999, comment: 'Reply', reply_all: true })
      ).rejects.toThrow(NotFoundError);
    });

    it('confirmReplyEmail with reply_all=false calls replyMessageAsync correctly', async () => {
      const prepared = await tools.prepareReplyEmail({
        message_id: 1,
        comment: 'Just to you',
        reply_all: false,
      });

      await tools.confirmReplyEmail({
        token_id: prepared.token_id,
        message_id: 1,
      });

      expect(repo.replyMessageAsync).toHaveBeenCalledWith(1, 'Just to you', false);
    });
  });

  // ===========================================================================
  // Forward Email (Two-Phase)
  // ===========================================================================

  describe('prepareForwardEmail / confirmForwardEmail', () => {
    it('prepareForwardEmail fetches original message and returns preview', async () => {
      const result = await tools.prepareForwardEmail({
        message_id: 1,
        to_recipients: ['dave@example.com'],
        comment: 'FYI',
      });

      expect(result.token_id).toBeDefined();
      expect(result.expires_at).toBeDefined();
      expect(result.original_message.id).toBe(1);
      expect(result.original_message.subject).toBe('Test Email Subject');
      expect(result.action).toContain('forward');
    });

    it('confirmForwardEmail validates token and calls forwardMessageAsync', async () => {
      const prepared = await tools.prepareForwardEmail({
        message_id: 1,
        to_recipients: ['dave@example.com'],
        comment: 'FYI',
      });

      const result = await tools.confirmForwardEmail({
        token_id: prepared.token_id,
        message_id: 1,
      });

      expect(result.success).toBe(true);
      expect(result.message).toContain('forwarded');
      expect(repo.forwardMessageAsync).toHaveBeenCalledWith(1, ['dave@example.com'], 'FYI');
    });

    it('prepareForwardEmail throws NotFoundError for missing message', async () => {
      await expect(
        tools.prepareForwardEmail({
          message_id: 999,
          to_recipients: ['dave@example.com'],
        })
      ).rejects.toThrow(NotFoundError);
    });

    it('confirmForwardEmail works without a comment', async () => {
      const prepared = await tools.prepareForwardEmail({
        message_id: 1,
        to_recipients: ['dave@example.com'],
      });

      await tools.confirmForwardEmail({
        token_id: prepared.token_id,
        message_id: 1,
      });

      expect(repo.forwardMessageAsync).toHaveBeenCalledWith(1, ['dave@example.com'], undefined);
    });
  });

  // ===========================================================================
  // Draft Reply/Forward (Non-Destructive)
  // ===========================================================================

  describe('replyAsDraft', () => {
    it('creates a reply draft and returns draft info', async () => {
      (repo.replyAsDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({
        numericId: 101,
        graphId: 'draft-reply-graph-101',
      });

      const result = await tools.replyAsDraft({
        message_id: 42,
        reply_all: false,
      });

      expect(repo.replyAsDraftAsync).toHaveBeenCalledWith(42, false, undefined);
      expect(result).toEqual({
        success: true,
        draft_id: 101,
        message: 'Reply draft created. Use update_draft to edit, then prepare_send_draft or prepare_send_email to send.',
      });
    });

    it('passes comment and reply_all to repository', async () => {
      (repo.replyAsDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({
        numericId: 102,
        graphId: 'draft-ra-graph-102',
      });

      await tools.replyAsDraft({
        message_id: 42,
        comment: 'Thanks!',
        reply_all: true,
      });

      expect(repo.replyAsDraftAsync).toHaveBeenCalledWith(42, true, 'Thanks!');
    });
  });

  describe('forwardAsDraft', () => {
    it('creates a forward draft and returns draft info', async () => {
      (repo.forwardAsDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({
        numericId: 201,
        graphId: 'draft-fwd-graph-201',
      });

      const result = await tools.forwardAsDraft({
        message_id: 42,
      });

      expect(repo.forwardAsDraftAsync).toHaveBeenCalledWith(42, undefined, undefined);
      expect(result).toEqual({
        success: true,
        draft_id: 201,
        message: 'Forward draft created. Use update_draft to edit, then prepare_send_draft or prepare_send_email to send.',
      });
    });

    it('passes recipients and comment to repository', async () => {
      (repo.forwardAsDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({
        numericId: 202,
        graphId: 'draft-fwd-graph-202',
      });

      await tools.forwardAsDraft({
        message_id: 42,
        to_recipients: ['alice@example.com'],
        comment: 'FYI',
      });

      expect(repo.forwardAsDraftAsync).toHaveBeenCalledWith(
        42,
        ['alice@example.com'],
        'FYI'
      );
    });
  });

  // ===========================================================================
  // Token Error Handling
  // ===========================================================================

  describe('token error handling', () => {
    it('confirmSendDraft throws ApprovalExpiredError for expired token', async () => {
      vi.useFakeTimers();
      try {
        const prepared = await tools.prepareSendDraft({ draft_id: 5 });

        // Advance past the 5-minute TTL
        vi.advanceTimersByTime(6 * 60 * 1000);

        await expect(
          tools.confirmSendDraft({ token_id: prepared.token_id, draft_id: 5 })
        ).rejects.toThrow(ApprovalExpiredError);

        expect(repo.sendDraftAsync).not.toHaveBeenCalled();
      } finally {
        vi.useRealTimers();
      }
    });

    it('confirmSendEmail throws ApprovalExpiredError for expired token', async () => {
      vi.useFakeTimers();
      try {
        const prepared = await tools.prepareSendEmail({
          to: ['bob@example.com'],
          subject: 'Test',
          body: 'Body',
          body_type: 'text',
        });

        vi.advanceTimersByTime(6 * 60 * 1000);

        await expect(
          tools.confirmSendEmail({ token_id: prepared.token_id })
        ).rejects.toThrow(ApprovalExpiredError);

        expect(repo.sendMailAsync).not.toHaveBeenCalled();
      } finally {
        vi.useRealTimers();
      }
    });

    it('confirmSendDraft throws ApprovalInvalidError for invalid token', async () => {
      await expect(
        tools.confirmSendDraft({
          token_id: '00000000-0000-0000-0000-000000000000',
          draft_id: 5,
        })
      ).rejects.toThrow(ApprovalInvalidError);
    });

    it('confirmSendEmail throws ApprovalInvalidError for invalid token', async () => {
      await expect(
        tools.confirmSendEmail({
          token_id: '00000000-0000-0000-0000-000000000000',
        })
      ).rejects.toThrow(ApprovalInvalidError);
    });

    it('confirmReplyEmail throws ApprovalInvalidError for invalid token', async () => {
      await expect(
        tools.confirmReplyEmail({
          token_id: '00000000-0000-0000-0000-000000000000',
          message_id: 1,
        })
      ).rejects.toThrow(ApprovalInvalidError);
    });

    it('confirmForwardEmail throws ApprovalInvalidError for invalid token', async () => {
      await expect(
        tools.confirmForwardEmail({
          token_id: '00000000-0000-0000-0000-000000000000',
          message_id: 1,
        })
      ).rejects.toThrow(ApprovalInvalidError);
    });

    it('token cannot be reused (one-time use)', async () => {
      const prepared = await tools.prepareSendDraft({ draft_id: 5 });

      // First confirm succeeds
      const result = await tools.confirmSendDraft({
        token_id: prepared.token_id,
        draft_id: 5,
      });
      expect(result.success).toBe(true);

      // Second confirm with the same token fails
      await expect(
        tools.confirmSendDraft({
          token_id: prepared.token_id,
          draft_id: 5,
        })
      ).rejects.toThrow(ApprovalInvalidError);
    });
  });

  // ===========================================================================
  // Signature Management
  // ===========================================================================

  describe('setSignature', () => {
    it('writes HTML signature and returns success', async () => {
      const result = await tools.setSignature({ content: '<p>Joel</p>', content_type: 'html' });

      expect(result).toEqual({ success: true, message: 'Signature saved successfully.' });
      expect(writeSignature).toHaveBeenCalledWith('<p>Joel</p>', 'html');
    });

    it('writes text signature and returns success', async () => {
      const result = await tools.setSignature({ content: '-- Joel', content_type: 'text' });

      expect(result).toEqual({ success: true, message: 'Signature saved successfully.' });
      expect(writeSignature).toHaveBeenCalledWith('-- Joel', 'text');
    });
  });

  describe('getSignature', () => {
    it('returns signature content when set', async () => {
      vi.mocked(readSignature).mockReturnValue('<p>-- Joel</p>');

      const result = await tools.getSignature();

      expect(result).toEqual({ has_signature: true, content: '<p>-- Joel</p>' });
    });

    it('returns no-signature message when not set', async () => {
      vi.mocked(readSignature).mockReturnValue(null);

      const result = await tools.getSignature();

      expect(result).toEqual({ has_signature: false, message: 'No signature is set. Use set_signature to create one.' });
    });
  });

  // ===========================================================================
  // Factory Function
  // ===========================================================================

  describe('createMailSendTools', () => {
    it('creates a MailSendTools instance', () => {
      const instance = createMailSendTools(repo, tokenManager);
      expect(instance).toBeInstanceOf(MailSendTools);
    });
  });
});
