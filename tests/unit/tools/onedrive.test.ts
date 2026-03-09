/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for OneDrive personal file tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { OneDriveTools, type IOneDriveRepository } from '../../../src/tools/onedrive.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';

describe('OneDriveTools', () => {
  let repo: IOneDriveRepository;
  let tokenManager: ApprovalTokenManager;
  let tools: OneDriveTools;

  beforeEach(() => {
    repo = {
      listDriveItemsAsync: vi.fn(),
      searchDriveItemsAsync: vi.fn(),
      getDriveItemAsync: vi.fn(),
      downloadFileAsync: vi.fn(),
      uploadFileAsync: vi.fn(),
      listRecentFilesAsync: vi.fn(),
      listSharedWithMeAsync: vi.fn(),
      createSharingLinkAsync: vi.fn(),
      deleteDriveItemAsync: vi.fn(),
    };
    tokenManager = new ApprovalTokenManager();
    tools = new OneDriveTools(repo, tokenManager);
  });

  // ===========================================================================
  // List Drive Items
  // ===========================================================================

  describe('listDriveItems', () => {
    it('returns items from the root when no folder_id provided', async () => {
      const mockItems = [
        { id: 1, name: 'Documents', size: 0, lastModified: '2026-01-01T00:00:00Z', isFolder: true, webUrl: 'https://example.com/Documents' },
        { id: 2, name: 'report.pdf', size: 1024, lastModified: '2026-01-02T00:00:00Z', isFolder: false, webUrl: 'https://example.com/report.pdf' },
      ];
      vi.mocked(repo.listDriveItemsAsync).mockResolvedValue(mockItems);

      const result = await tools.listDriveItems({});

      expect(repo.listDriveItemsAsync).toHaveBeenCalledWith(undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.items).toEqual(mockItems);
    });

    it('returns items from a specific folder', async () => {
      const mockItems = [
        { id: 3, name: 'notes.txt', size: 256, lastModified: '2026-01-03T00:00:00Z', isFolder: false, webUrl: 'https://example.com/notes.txt' },
      ];
      vi.mocked(repo.listDriveItemsAsync).mockResolvedValue(mockItems);

      const result = await tools.listDriveItems({ folder_id: 1 });

      expect(repo.listDriveItemsAsync).toHaveBeenCalledWith(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.items).toHaveLength(1);
    });
  });

  // ===========================================================================
  // Search Drive Items
  // ===========================================================================

  describe('searchDriveItems', () => {
    it('searches and returns matching items', async () => {
      const mockItems = [
        { id: 10, name: 'budget.xlsx', size: 2048, lastModified: '2026-02-01T00:00:00Z', isFolder: false, webUrl: 'https://example.com/budget.xlsx' },
      ];
      vi.mocked(repo.searchDriveItemsAsync).mockResolvedValue(mockItems);

      const result = await tools.searchDriveItems({ query: 'budget' });

      expect(repo.searchDriveItemsAsync).toHaveBeenCalledWith('budget', undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.items).toEqual(mockItems);
    });

    it('passes limit parameter', async () => {
      vi.mocked(repo.searchDriveItemsAsync).mockResolvedValue([]);

      await tools.searchDriveItems({ query: 'test', limit: 5 });

      expect(repo.searchDriveItemsAsync).toHaveBeenCalledWith('test', 5);
    });
  });

  // ===========================================================================
  // Get Drive Item
  // ===========================================================================

  describe('getDriveItem', () => {
    it('returns item details', async () => {
      const mockItem = {
        id: 10, name: 'budget.xlsx', size: 2048, lastModified: '2026-02-01T00:00:00Z',
        isFolder: false, webUrl: 'https://example.com/budget.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        createdBy: 'John Doe',
      };
      vi.mocked(repo.getDriveItemAsync).mockResolvedValue(mockItem);

      const result = await tools.getDriveItem({ item_id: 10 });

      expect(repo.getDriveItemAsync).toHaveBeenCalledWith(10);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.item).toEqual(mockItem);
    });
  });

  // ===========================================================================
  // Download File
  // ===========================================================================

  describe('downloadFile', () => {
    it('downloads and returns the saved path and size', async () => {
      vi.mocked(repo.downloadFileAsync).mockResolvedValue({ savedPath: '/tmp/report.pdf', size: 1024 });

      const result = await tools.downloadFile({ item_id: 2, output_path: '/tmp/report.pdf' });

      expect(repo.downloadFileAsync).toHaveBeenCalledWith(2, '/tmp/report.pdf');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.savedPath).toBe('/tmp/report.pdf');
      expect(parsed.size).toBe(1024);
    });
  });

  // ===========================================================================
  // Upload File (Two-phase)
  // ===========================================================================

  describe('prepareUploadFile', () => {
    it('returns an approval token with upload metadata', () => {
      const result = tools.prepareUploadFile({
        parent_path: 'Documents/Reports',
        file_name: 'quarterly.pdf',
        local_file_path: '/tmp/quarterly.pdf',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.parent_path).toBe('Documents/Reports');
      expect(parsed.file_name).toBe('quarterly.pdf');
      expect(parsed.local_file_path).toBe('/tmp/quarterly.pdf');
      expect(parsed.action).toContain('confirm_upload_file');
    });
  });

  describe('confirmUploadFile', () => {
    it('uploads the file using stored metadata', async () => {
      vi.mocked(repo.uploadFileAsync).mockResolvedValue(42);

      const prepareResult = tools.prepareUploadFile({
        parent_path: 'Documents',
        file_name: 'test.txt',
        local_file_path: '/tmp/test.txt',
      });
      const token = JSON.parse(prepareResult.content[0].text).approval_token;

      const result = await tools.confirmUploadFile({ approval_token: token });

      expect(repo.uploadFileAsync).toHaveBeenCalledWith('Documents', 'test.txt', '/tmp/test.txt');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.item_id).toBe(42);
    });

    it('rejects invalid token', async () => {
      const result = await tools.confirmUploadFile({ approval_token: 'invalid-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toContain('not found');
    });

    it('rejects already consumed token', async () => {
      vi.mocked(repo.uploadFileAsync).mockResolvedValue(42);

      const prepareResult = tools.prepareUploadFile({
        parent_path: 'Documents',
        file_name: 'test.txt',
        local_file_path: '/tmp/test.txt',
      });
      const token = JSON.parse(prepareResult.content[0].text).approval_token;

      // First use should succeed
      await tools.confirmUploadFile({ approval_token: token });

      // Second use should fail
      const result = await tools.confirmUploadFile({ approval_token: token });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });

  // ===========================================================================
  // Recent Files
  // ===========================================================================

  describe('listRecentFiles', () => {
    it('returns recent files', async () => {
      const mockItems = [
        { id: 20, name: 'recent.docx', size: 512, lastModified: '2026-03-01T00:00:00Z', isFolder: false, webUrl: 'https://example.com/recent.docx' },
      ];
      vi.mocked(repo.listRecentFilesAsync).mockResolvedValue(mockItems);

      const result = await tools.listRecentFiles();

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.items).toEqual(mockItems);
    });
  });

  // ===========================================================================
  // Shared With Me
  // ===========================================================================

  describe('listSharedWithMe', () => {
    it('returns shared files', async () => {
      const mockItems = [
        { id: 30, name: 'shared.pptx', size: 4096, lastModified: '2026-03-02T00:00:00Z', isFolder: false, webUrl: 'https://example.com/shared.pptx' },
      ];
      vi.mocked(repo.listSharedWithMeAsync).mockResolvedValue(mockItems);

      const result = await tools.listSharedWithMe();

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.items).toEqual(mockItems);
    });
  });

  // ===========================================================================
  // Create Sharing Link
  // ===========================================================================

  describe('createSharingLink', () => {
    it('creates a view link with anonymous scope', async () => {
      vi.mocked(repo.createSharingLinkAsync).mockResolvedValue({
        webUrl: 'https://example.com/share/abc',
        type: 'view',
        scope: 'anonymous',
      });

      const result = await tools.createSharingLink({ item_id: 10, type: 'view', scope: 'anonymous' });

      expect(repo.createSharingLinkAsync).toHaveBeenCalledWith(10, 'view', 'anonymous');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.link.webUrl).toBe('https://example.com/share/abc');
      expect(parsed.link.type).toBe('view');
      expect(parsed.link.scope).toBe('anonymous');
    });

    it('creates an edit link with organization scope', async () => {
      vi.mocked(repo.createSharingLinkAsync).mockResolvedValue({
        webUrl: 'https://example.com/share/def',
        type: 'edit',
        scope: 'organization',
      });

      const result = await tools.createSharingLink({ item_id: 10, type: 'edit', scope: 'organization' });

      expect(repo.createSharingLinkAsync).toHaveBeenCalledWith(10, 'edit', 'organization');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.link.type).toBe('edit');
      expect(parsed.link.scope).toBe('organization');
    });
  });

  // ===========================================================================
  // Delete Drive Item (Two-phase)
  // ===========================================================================

  describe('prepareDeleteDriveItem', () => {
    it('returns an approval token', () => {
      const result = tools.prepareDeleteDriveItem({ item_id: 10 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.item_id).toBe(10);
      expect(parsed.action).toContain('confirm_delete_drive_item');
    });
  });

  describe('confirmDeleteDriveItem', () => {
    it('deletes the drive item with valid token', async () => {
      vi.mocked(repo.deleteDriveItemAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeleteDriveItem({ item_id: 10 });
      const token = JSON.parse(prepareResult.content[0].text).approval_token;

      const result = await tools.confirmDeleteDriveItem({ approval_token: token });

      expect(repo.deleteDriveItemAsync).toHaveBeenCalledWith(10);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Drive item deleted');
    });

    it('rejects invalid token', async () => {
      const result = await tools.confirmDeleteDriveItem({ approval_token: 'bad-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toContain('not found');
    });

    it('rejects already consumed token', async () => {
      vi.mocked(repo.deleteDriveItemAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeleteDriveItem({ item_id: 10 });
      const token = JSON.parse(prepareResult.content[0].text).approval_token;

      await tools.confirmDeleteDriveItem({ approval_token: token });

      const result = await tools.confirmDeleteDriveItem({ approval_token: token });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });
});
