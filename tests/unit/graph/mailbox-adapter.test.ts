/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for GraphMailboxAdapter.
 *
 * Verifies that each IMailboxRepository method delegates
 * to the corresponding GraphRepository xxxAsync() method.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { GraphMailboxAdapter } from '../../../src/graph/mailbox-adapter.js';
import type { GraphRepository } from '../../../src/graph/repository.js';
import type { FolderRow, EmailRow } from '../../../src/database/repository.js';

// =============================================================================
// Helpers
// =============================================================================

function makeFakeEmail(overrides: Partial<EmailRow> = {}): EmailRow {
  return {
    id: 1,
    folderId: 10,
    subject: 'Test',
    sender: 'Alice',
    senderAddress: 'alice@example.com',
    recipients: 'bob@example.com',
    displayTo: 'Bob',
    toAddresses: 'bob@example.com',
    ccAddresses: null,
    preview: 'Hello',
    isRead: 0,
    timeReceived: 1000,
    timeSent: 900,
    hasAttachment: 0,
    size: 1024,
    priority: 0,
    flagStatus: 0,
    categories: null,
    messageId: 'msg-1',
    conversationId: null,
    dataFilePath: null,
    ...overrides,
  };
}

function makeFakeFolder(overrides: Partial<FolderRow> = {}): FolderRow {
  return {
    id: 10,
    name: 'Inbox',
    parentId: null,
    specialType: 0,
    folderType: 0,
    accountId: 1,
    messageCount: 5,
    unreadCount: 2,
    ...overrides,
  };
}

// =============================================================================
// Mock GraphRepository
// =============================================================================

function createMockGraphRepo() {
  return {
    getEmailAsync: vi.fn<[number], Promise<EmailRow | undefined>>(),
    getFolderAsync: vi.fn<[number], Promise<FolderRow | undefined>>(),
    moveEmailAsync: vi.fn<[number, number], Promise<void>>(),
    deleteEmailAsync: vi.fn<[number], Promise<void>>(),
    archiveEmailAsync: vi.fn<[number], Promise<void>>(),
    junkEmailAsync: vi.fn<[number], Promise<void>>(),
    markEmailReadAsync: vi.fn<[number, boolean], Promise<void>>(),
    setEmailFlagAsync: vi.fn<[number, number], Promise<void>>(),
    setEmailCategoriesAsync: vi.fn<[number, string[]], Promise<void>>(),
    createFolderAsync: vi.fn<[string, number?], Promise<FolderRow>>(),
    deleteFolderAsync: vi.fn<[number], Promise<void>>(),
    renameFolderAsync: vi.fn<[number, string], Promise<void>>(),
    moveFolderAsync: vi.fn<[number, number], Promise<void>>(),
    emptyFolderAsync: vi.fn<[number], Promise<void>>(),
  } as unknown as GraphRepository &
    Record<string, ReturnType<typeof vi.fn>>;
}

// =============================================================================
// Tests
// =============================================================================

describe('GraphMailboxAdapter', () => {
  let graphRepo: ReturnType<typeof createMockGraphRepo>;
  let adapter: GraphMailboxAdapter;

  beforeEach(() => {
    graphRepo = createMockGraphRepo();
    adapter = new GraphMailboxAdapter(graphRepo as unknown as GraphRepository);
  });

  // ---------------------------------------------------------------------------
  // Read
  // ---------------------------------------------------------------------------

  it('getEmail delegates to getEmailAsync', async () => {
    const email = makeFakeEmail();
    (graphRepo.getEmailAsync as ReturnType<typeof vi.fn>).mockResolvedValue(email);

    const result = await adapter.getEmail(1);
    expect(result).toBe(email);
    expect(graphRepo.getEmailAsync).toHaveBeenCalledWith(1);
  });

  it('getFolder delegates to getFolderAsync', async () => {
    const folder = makeFakeFolder();
    (graphRepo.getFolderAsync as ReturnType<typeof vi.fn>).mockResolvedValue(folder);

    const result = await adapter.getFolder(10);
    expect(result).toBe(folder);
    expect(graphRepo.getFolderAsync).toHaveBeenCalledWith(10);
  });

  // ---------------------------------------------------------------------------
  // Email organization
  // ---------------------------------------------------------------------------

  it('moveEmail delegates to moveEmailAsync', async () => {
    (graphRepo.moveEmailAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    await adapter.moveEmail(1, 20);
    expect(graphRepo.moveEmailAsync).toHaveBeenCalledWith(1, 20);
  });

  it('deleteEmail delegates to deleteEmailAsync', async () => {
    (graphRepo.deleteEmailAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    await adapter.deleteEmail(1);
    expect(graphRepo.deleteEmailAsync).toHaveBeenCalledWith(1);
  });

  it('archiveEmail delegates to archiveEmailAsync', async () => {
    (graphRepo.archiveEmailAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    await adapter.archiveEmail(1);
    expect(graphRepo.archiveEmailAsync).toHaveBeenCalledWith(1);
  });

  it('junkEmail delegates to junkEmailAsync', async () => {
    (graphRepo.junkEmailAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    await adapter.junkEmail(1);
    expect(graphRepo.junkEmailAsync).toHaveBeenCalledWith(1);
  });

  it('markEmailRead delegates to markEmailReadAsync', async () => {
    (graphRepo.markEmailReadAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    await adapter.markEmailRead(1, true);
    expect(graphRepo.markEmailReadAsync).toHaveBeenCalledWith(1, true);
  });

  it('setEmailFlag delegates to setEmailFlagAsync', async () => {
    (graphRepo.setEmailFlagAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    await adapter.setEmailFlag(1, 2);
    expect(graphRepo.setEmailFlagAsync).toHaveBeenCalledWith(1, 2);
  });

  it('setEmailCategories delegates to setEmailCategoriesAsync', async () => {
    (graphRepo.setEmailCategoriesAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    await adapter.setEmailCategories(1, ['Important']);
    expect(graphRepo.setEmailCategoriesAsync).toHaveBeenCalledWith(1, ['Important']);
  });

  // ---------------------------------------------------------------------------
  // Folder management
  // ---------------------------------------------------------------------------

  it('createFolder delegates to createFolderAsync', async () => {
    const folder = makeFakeFolder({ id: 99, name: 'New Folder' });
    (graphRepo.createFolderAsync as ReturnType<typeof vi.fn>).mockResolvedValue(folder);

    const result = await adapter.createFolder('New Folder', 10);
    expect(result).toBe(folder);
    expect(graphRepo.createFolderAsync).toHaveBeenCalledWith('New Folder', 10);
  });

  it('deleteFolder delegates to deleteFolderAsync', async () => {
    (graphRepo.deleteFolderAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    await adapter.deleteFolder(10);
    expect(graphRepo.deleteFolderAsync).toHaveBeenCalledWith(10);
  });

  it('renameFolder delegates to renameFolderAsync', async () => {
    (graphRepo.renameFolderAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    await adapter.renameFolder(10, 'Renamed');
    expect(graphRepo.renameFolderAsync).toHaveBeenCalledWith(10, 'Renamed');
  });

  it('moveFolder delegates to moveFolderAsync', async () => {
    (graphRepo.moveFolderAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    await adapter.moveFolder(10, 20);
    expect(graphRepo.moveFolderAsync).toHaveBeenCalledWith(10, 20);
  });

  it('emptyFolder delegates to emptyFolderAsync', async () => {
    (graphRepo.emptyFolderAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

    await adapter.emptyFolder(10);
    expect(graphRepo.emptyFolderAsync).toHaveBeenCalledWith(10);
  });
});
