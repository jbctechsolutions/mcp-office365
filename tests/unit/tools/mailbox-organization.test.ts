/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  MailboxOrganizationTools,
  createMailboxOrganizationTools,
} from '../../../src/tools/mailbox-organization.js';
import type { IMailboxRepository, EmailRow, FolderRow } from '../../../src/database/repository.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';
import {
  NotFoundError,
  ApprovalExpiredError,
  ApprovalInvalidError,
  TargetChangedError,
} from '../../../src/utils/errors.js';

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
    timeReceived: 700000000, // Apple epoch timestamp
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

function makeFolderRow(overrides: Partial<FolderRow> = {}): FolderRow {
  return {
    id: 10,
    name: 'Inbox',
    parentId: null,
    specialType: 1,
    folderType: 0,
    accountId: 1,
    messageCount: 25,
    unreadCount: 5,
    ...overrides,
  };
}

// =============================================================================
// Mock Repository
// =============================================================================

function createMockRepository(): IMailboxRepository {
  return {
    getEmail: vi.fn(),
    getFolder: vi.fn(),
    moveEmail: vi.fn(),
    deleteEmail: vi.fn(),
    archiveEmail: vi.fn(),
    junkEmail: vi.fn(),
    markEmailRead: vi.fn(),
    setEmailFlag: vi.fn(),
    setEmailCategories: vi.fn(),
    createFolder: vi.fn(),
    deleteFolder: vi.fn(),
    renameFolder: vi.fn(),
    moveFolder: vi.fn(),
    emptyFolder: vi.fn(),
  };
}

// =============================================================================
// Tests
// =============================================================================

describe('MailboxOrganizationTools', () => {
  let repo: ReturnType<typeof createMockRepository>;
  let tokenManager: ApprovalTokenManager;
  let tools: MailboxOrganizationTools;

  const testEmail = makeEmailRow({ id: 1, folderId: 10 });
  const testEmail2 = makeEmailRow({ id: 2, folderId: 10, subject: 'Second Email' });
  const testEmail3 = makeEmailRow({ id: 3, folderId: 10, subject: 'Third Email' });
  const testFolder = makeFolderRow({ id: 10, name: 'Inbox', messageCount: 25 });
  const destFolder = makeFolderRow({ id: 20, name: 'Archive', messageCount: 0, specialType: 0 });

  beforeEach(() => {
    repo = createMockRepository();
    tokenManager = new ApprovalTokenManager();
    tools = new MailboxOrganizationTools(repo, tokenManager);

    // Default mock setup: emails and folders exist
    (repo.getEmail as ReturnType<typeof vi.fn>).mockImplementation((id: number) => {
      if (id === 1) return testEmail;
      if (id === 2) return testEmail2;
      if (id === 3) return testEmail3;
      return undefined;
    });

    (repo.getFolder as ReturnType<typeof vi.fn>).mockImplementation((id: number) => {
      if (id === 10) return testFolder;
      if (id === 20) return destFolder;
      return undefined;
    });
  });

  // ===========================================================================
  // Prepare / Confirm Flow Tests
  // ===========================================================================

  describe('prepareDeleteEmail / confirmDeleteEmail', () => {
    it('prepareDeleteEmail returns token and email preview', async () => {
      const result = await tools.prepareDeleteEmail({ email_id: 1 });

      expect(result.token_id).toBeDefined();
      expect(typeof result.token_id).toBe('string');
      expect(result.expires_at).toBeDefined();
      expect(result.email).toBeDefined();
      expect(result.email.id).toBe(1);
      expect(result.email.subject).toBe('Test Email Subject');
      expect(result.email.sender).toBe('Alice Sender');
      expect(result.email.senderAddress).toBe('alice@example.com');
      expect(result.email.folderId).toBe(10);
      expect(result.action).toContain('Deleted Items');
    });

    it('confirmDeleteEmail validates token and calls deleteEmail', async () => {
      const prepared = await tools.prepareDeleteEmail({ email_id: 1 });
      const result = await tools.confirmDeleteEmail({
        token_id: prepared.token_id,
        email_id: 1,
      });

      expect(result.success).toBe(true);
      expect(result.message).toContain('Deleted Items');
      expect(repo.deleteEmail).toHaveBeenCalledWith(1);
    });

    it('confirmDeleteEmail throws if token expired', async () => {
      vi.useFakeTimers();

      try {
        const prepared = await tools.prepareDeleteEmail({ email_id: 1 });

        // Advance past the 5-minute TTL
        vi.advanceTimersByTime(6 * 60 * 1000);

        await expect(
          tools.confirmDeleteEmail({
            token_id: prepared.token_id,
            email_id: 1,
          })
        ).rejects.toThrow(ApprovalExpiredError);

        expect(repo.deleteEmail).not.toHaveBeenCalled();
      } finally {
        vi.useRealTimers();
      }
    });

    it('confirmDeleteEmail throws if email changed between prepare and confirm', async () => {
      const prepared = await tools.prepareDeleteEmail({ email_id: 1 });

      // Simulate the email changing after prepare (different subject changes the hash)
      const modifiedEmail = makeEmailRow({
        id: 1,
        folderId: 10,
        subject: 'Modified Subject After Prepare',
      });
      (repo.getEmail as ReturnType<typeof vi.fn>).mockImplementation((id: number) => {
        if (id === 1) return modifiedEmail;
        return undefined;
      });

      await expect(
        tools.confirmDeleteEmail({
          token_id: prepared.token_id,
          email_id: 1,
        })
      ).rejects.toThrow(TargetChangedError);

      expect(repo.deleteEmail).not.toHaveBeenCalled();
    });
  });

  describe('prepareMoveEmail / confirmMoveEmail', () => {
    it('prepareMoveEmail includes destination folder in response', async () => {
      const result = await tools.prepareMoveEmail({
        email_id: 1,
        destination_folder_id: 20,
      });

      expect(result.token_id).toBeDefined();
      expect(result.email.id).toBe(1);
      expect(result.destination_folder).toBeDefined();
      expect(result.destination_folder.id).toBe(20);
      expect(result.destination_folder.name).toBe('Archive');
      expect(result.action).toContain('Archive');
    });

    it('confirmMoveEmail calls moveEmail with correct destination', async () => {
      const prepared = await tools.prepareMoveEmail({
        email_id: 1,
        destination_folder_id: 20,
      });

      const result = await tools.confirmMoveEmail({
        token_id: prepared.token_id,
        email_id: 1,
      });

      expect(result.success).toBe(true);
      expect(repo.moveEmail).toHaveBeenCalledWith(1, 20);
    });
  });

  describe('prepareArchiveEmail / confirmArchiveEmail', () => {
    it('prepareArchiveEmail returns token and email preview', async () => {
      const result = await tools.prepareArchiveEmail({ email_id: 1 });

      expect(result.token_id).toBeDefined();
      expect(result.email.id).toBe(1);
      expect(result.action).toContain('Archive');
    });

    it('confirmArchiveEmail calls archiveEmail', async () => {
      const prepared = await tools.prepareArchiveEmail({ email_id: 1 });

      const result = await tools.confirmArchiveEmail({
        token_id: prepared.token_id,
        email_id: 1,
      });

      expect(result.success).toBe(true);
      expect(repo.archiveEmail).toHaveBeenCalledWith(1);
    });
  });

  describe('prepareJunkEmail / confirmJunkEmail', () => {
    it('prepareJunkEmail returns token and email preview', async () => {
      const result = await tools.prepareJunkEmail({ email_id: 1 });

      expect(result.token_id).toBeDefined();
      expect(result.email.id).toBe(1);
      expect(result.action).toContain('Junk');
    });

    it('confirmJunkEmail calls junkEmail', async () => {
      const prepared = await tools.prepareJunkEmail({ email_id: 1 });

      const result = await tools.confirmJunkEmail({
        token_id: prepared.token_id,
        email_id: 1,
      });

      expect(result.success).toBe(true);
      expect(repo.junkEmail).toHaveBeenCalledWith(1);
    });
  });

  describe('prepareDeleteFolder / confirmDeleteFolder', () => {
    it('prepareDeleteFolder returns token and folder preview', async () => {
      const result = await tools.prepareDeleteFolder({ folder_id: 10 });

      expect(result.token_id).toBeDefined();
      expect(result.folder.id).toBe(10);
      expect(result.folder.name).toBe('Inbox');
      expect(result.folder.messageCount).toBe(25);
      expect(result.action).toContain('25');
    });

    it('confirmDeleteFolder calls deleteFolder', async () => {
      const prepared = await tools.prepareDeleteFolder({ folder_id: 10 });

      const result = await tools.confirmDeleteFolder({
        token_id: prepared.token_id,
        folder_id: 10,
      });

      expect(result.success).toBe(true);
      expect(repo.deleteFolder).toHaveBeenCalledWith(10);
    });
  });

  describe('prepareEmptyFolder / confirmEmptyFolder', () => {
    it('prepareEmptyFolder returns token and folder preview', async () => {
      const result = await tools.prepareEmptyFolder({ folder_id: 10 });

      expect(result.token_id).toBeDefined();
      expect(result.folder.id).toBe(10);
      expect(result.action).toContain('25');
    });

    it('confirmEmptyFolder calls emptyFolder', async () => {
      const prepared = await tools.prepareEmptyFolder({ folder_id: 10 });

      const result = await tools.confirmEmptyFolder({
        token_id: prepared.token_id,
        folder_id: 10,
      });

      expect(result.success).toBe(true);
      expect(repo.emptyFolder).toHaveBeenCalledWith(10);
    });
  });

  // ===========================================================================
  // Batch Operation Tests
  // ===========================================================================

  describe('prepareBatchDeleteEmails / confirmBatchOperation', () => {
    it('prepareBatchDeleteEmails returns individual tokens per email', async () => {
      const result = await tools.prepareBatchDeleteEmails({
        email_ids: [1, 2, 3],
      });

      expect(result.tokens).toHaveLength(3);
      expect(result.tokens[0]!.email.id).toBe(1);
      expect(result.tokens[1]!.email.id).toBe(2);
      expect(result.tokens[2]!.email.id).toBe(3);

      // Each token should be unique
      const tokenIds = result.tokens.map((t) => t.token_id);
      expect(new Set(tokenIds).size).toBe(3);

      expect(result.action).toContain('3');
      expect(result.expires_at).toBeDefined();
    });

    it('confirmBatchOperation processes each token independently', async () => {
      const prepared = await tools.prepareBatchDeleteEmails({
        email_ids: [1, 2, 3],
      });

      const result = await tools.confirmBatchOperation({
        tokens: prepared.tokens.map((t) => ({
          token_id: t.token_id,
          email_id: t.email.id,
        })),
      });

      expect(result.summary.total).toBe(3);
      expect(result.summary.succeeded).toBe(3);
      expect(result.summary.failed).toBe(0);

      expect(repo.deleteEmail).toHaveBeenCalledWith(1);
      expect(repo.deleteEmail).toHaveBeenCalledWith(2);
      expect(repo.deleteEmail).toHaveBeenCalledWith(3);
    });

    it('confirmBatchOperation reports partial success (some succeed, some fail)', async () => {
      const prepared = await tools.prepareBatchDeleteEmails({
        email_ids: [1, 2, 3],
      });

      // Make email 2 disappear before confirm to cause a failure on hash check
      (repo.getEmail as ReturnType<typeof vi.fn>).mockImplementation((id: number) => {
        if (id === 1) return testEmail;
        if (id === 2) return undefined; // Email 2 is gone
        if (id === 3) return testEmail3;
        return undefined;
      });

      const result = await tools.confirmBatchOperation({
        tokens: prepared.tokens.map((t) => ({
          token_id: t.token_id,
          email_id: t.email.id,
        })),
      });

      expect(result.summary.total).toBe(3);
      expect(result.summary.succeeded).toBe(2);
      expect(result.summary.failed).toBe(1);

      // Verify the failed entry
      const failedResult = result.results.find((r) => r.email_id === 2);
      expect(failedResult).toBeDefined();
      expect(failedResult!.success).toBe(false);
      expect(failedResult!.success === false && failedResult!.error).toBeDefined();

      // Verify successful entries
      expect(repo.deleteEmail).toHaveBeenCalledWith(1);
      expect(repo.deleteEmail).toHaveBeenCalledWith(3);
      expect(repo.deleteEmail).not.toHaveBeenCalledWith(2);
    });
  });

  // ===========================================================================
  // Low-Risk Operation Tests
  // ===========================================================================

  describe('markEmailRead', () => {
    it('calls markEmailRead(id, true)', async () => {
      const result = await tools.markEmailRead({ email_id: 1 });

      expect(result.success).toBe(true);
      expect(result.message).toContain('read');
      expect(repo.markEmailRead).toHaveBeenCalledWith(1, true);
    });
  });

  describe('markEmailUnread', () => {
    it('calls markEmailRead(id, false)', async () => {
      const result = await tools.markEmailUnread({ email_id: 1 });

      expect(result.success).toBe(true);
      expect(result.message).toContain('unread');
      expect(repo.markEmailRead).toHaveBeenCalledWith(1, false);
    });
  });

  describe('setEmailFlag', () => {
    it('calls setEmailFlag with correct status', async () => {
      const result = await tools.setEmailFlag({ email_id: 1, flag_status: 1 });

      expect(result.success).toBe(true);
      expect(repo.setEmailFlag).toHaveBeenCalledWith(1, 1);
    });
  });

  describe('clearEmailFlag', () => {
    it('calls setEmailFlag with 0', async () => {
      const result = await tools.clearEmailFlag({ email_id: 1 });

      expect(result.success).toBe(true);
      expect(repo.setEmailFlag).toHaveBeenCalledWith(1, 0);
    });
  });

  describe('setEmailCategories', () => {
    it('calls setEmailCategories with categories', async () => {
      const categories = ['Important', 'Work'];
      const result = await tools.setEmailCategories({ email_id: 1, categories });

      expect(result.success).toBe(true);
      expect(repo.setEmailCategories).toHaveBeenCalledWith(1, ['Important', 'Work']);
    });
  });

  // ===========================================================================
  // Non-Destructive Operation Tests
  // ===========================================================================

  describe('createFolder', () => {
    it('calls createFolder and returns preview', async () => {
      const newFolder = makeFolderRow({
        id: 30,
        name: 'My New Folder',
        messageCount: 0,
        unreadCount: 0,
      });
      (repo.createFolder as ReturnType<typeof vi.fn>).mockReturnValue(newFolder);

      const result = await tools.createFolder({ name: 'My New Folder' });

      expect(result.success).toBe(true);
      expect(result.folder.id).toBe(30);
      expect(result.folder.name).toBe('My New Folder');
      expect(result.folder.messageCount).toBe(0);
      expect(repo.createFolder).toHaveBeenCalledWith('My New Folder', undefined);
    });

    it('passes parent_folder_id when provided', async () => {
      const newFolder = makeFolderRow({ id: 31, name: 'Subfolder', parentId: 10 });
      (repo.createFolder as ReturnType<typeof vi.fn>).mockReturnValue(newFolder);

      await tools.createFolder({ name: 'Subfolder', parent_folder_id: 10 });

      expect(repo.createFolder).toHaveBeenCalledWith('Subfolder', 10);
    });
  });

  describe('renameFolder', () => {
    it('calls renameFolder', async () => {
      const result = await tools.renameFolder({ folder_id: 10, new_name: 'Renamed Inbox' });

      expect(result.success).toBe(true);
      expect(result.message).toContain('Renamed Inbox');
      expect(repo.renameFolder).toHaveBeenCalledWith(10, 'Renamed Inbox');
    });
  });

  describe('moveFolder', () => {
    it('calls moveFolder', async () => {
      const result = await tools.moveFolder({ folder_id: 10, destination_parent_id: 20 });

      expect(result.success).toBe(true);
      expect(repo.moveFolder).toHaveBeenCalledWith(10, 20);
    });
  });

  // ===========================================================================
  // Error Handling Tests
  // ===========================================================================

  describe('error handling', () => {
    it('prepareDeleteEmail throws NotFoundError for missing email', async () => {
      await expect(tools.prepareDeleteEmail({ email_id: 999 })).rejects.toThrow(NotFoundError);
    });

    it('prepareMoveEmail throws NotFoundError for missing email', async () => {
      await expect(
        tools.prepareMoveEmail({ email_id: 999, destination_folder_id: 20 })
      ).rejects.toThrow(NotFoundError);
    });

    it('prepareMoveEmail throws NotFoundError for missing destination folder', async () => {
      await expect(
        tools.prepareMoveEmail({ email_id: 1, destination_folder_id: 999 })
      ).rejects.toThrow(NotFoundError);
    });

    it('prepareArchiveEmail throws NotFoundError for missing email', async () => {
      await expect(tools.prepareArchiveEmail({ email_id: 999 })).rejects.toThrow(NotFoundError);
    });

    it('prepareJunkEmail throws NotFoundError for missing email', async () => {
      await expect(tools.prepareJunkEmail({ email_id: 999 })).rejects.toThrow(NotFoundError);
    });

    it('prepareDeleteFolder throws NotFoundError for missing folder', async () => {
      await expect(tools.prepareDeleteFolder({ folder_id: 999 })).rejects.toThrow(NotFoundError);
    });

    it('prepareEmptyFolder throws NotFoundError for missing folder', async () => {
      await expect(tools.prepareEmptyFolder({ folder_id: 999 })).rejects.toThrow(NotFoundError);
    });

    it('markEmailRead throws NotFoundError for missing email', async () => {
      await expect(tools.markEmailRead({ email_id: 999 })).rejects.toThrow(NotFoundError);
    });

    it('markEmailUnread throws NotFoundError for missing email', async () => {
      await expect(tools.markEmailUnread({ email_id: 999 })).rejects.toThrow(NotFoundError);
    });

    it('setEmailFlag throws NotFoundError for missing email', async () => {
      await expect(tools.setEmailFlag({ email_id: 999, flag_status: 1 })).rejects.toThrow(NotFoundError);
    });

    it('clearEmailFlag throws NotFoundError for missing email', async () => {
      await expect(tools.clearEmailFlag({ email_id: 999 })).rejects.toThrow(NotFoundError);
    });

    it('setEmailCategories throws NotFoundError for missing email', async () => {
      await expect(
        tools.setEmailCategories({ email_id: 999, categories: ['Work'] })
      ).rejects.toThrow(NotFoundError);
    });

    it('renameFolder throws NotFoundError for missing folder', async () => {
      await expect(
        tools.renameFolder({ folder_id: 999, new_name: 'New Name' })
      ).rejects.toThrow(NotFoundError);
    });

    it('moveFolder throws NotFoundError for missing source folder', async () => {
      await expect(
        tools.moveFolder({ folder_id: 999, destination_parent_id: 20 })
      ).rejects.toThrow(NotFoundError);
    });

    it('moveFolder throws NotFoundError for missing destination folder', async () => {
      await expect(
        tools.moveFolder({ folder_id: 10, destination_parent_id: 999 })
      ).rejects.toThrow(NotFoundError);
    });

    it('token cannot be reused (one-time use)', async () => {
      const prepared = await tools.prepareDeleteEmail({ email_id: 1 });

      // First confirm succeeds
      const result = await tools.confirmDeleteEmail({
        token_id: prepared.token_id,
        email_id: 1,
      });
      expect(result.success).toBe(true);

      // Second confirm with the same token fails
      await expect(
        tools.confirmDeleteEmail({
          token_id: prepared.token_id,
          email_id: 1,
        })
      ).rejects.toThrow(ApprovalInvalidError);
    });
  });

  // ===========================================================================
  // Factory Function Tests
  // ===========================================================================

  describe('createMailboxOrganizationTools', () => {
    it('creates a MailboxOrganizationTools instance', () => {
      const instance = createMailboxOrganizationTools(repo, tokenManager);
      expect(instance).toBeInstanceOf(MailboxOrganizationTools);
    });
  });
});
