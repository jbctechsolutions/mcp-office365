/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for SharePoint Document Library tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { SharePointTools, type ISharePointRepository } from '../../../src/tools/sharepoint.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';

describe('SharePointTools', () => {
  let repo: ISharePointRepository;
  let tokenManager: ApprovalTokenManager;
  let tools: SharePointTools;

  beforeEach(() => {
    repo = {
      listSitesAsync: vi.fn(),
      searchSitesAsync: vi.fn(),
      getSiteAsync: vi.fn(),
      listDocumentLibrariesAsync: vi.fn(),
      listLibraryItemsAsync: vi.fn(),
      downloadLibraryFileAsync: vi.fn(),
      createLibraryFolderAsync: vi.fn(),
      uploadLibraryFileAsync: vi.fn(),
    };
    tokenManager = new ApprovalTokenManager();
    tools = new SharePointTools(repo, tokenManager);
  });

  // ===========================================================================
  // Sites
  // ===========================================================================

  describe('listSites', () => {
    it('returns followed sites from the repository', async () => {
      const mockSites = [
        { id: 'si_aaaaaaaaaaaaaa', name: 'Team Site', webUrl: 'https://contoso.sharepoint.com/sites/team', displayName: 'Team Site' },
        { id: 'si_bbbbbbbbbbbbbb', name: 'HR Portal', webUrl: 'https://contoso.sharepoint.com/sites/hr', displayName: 'HR Portal' },
      ];
      vi.mocked(repo.listSitesAsync).mockResolvedValue(mockSites);

      const result = await tools.listSites();

      expect(result.content).toHaveLength(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.sites).toEqual(mockSites);
    });

    it('returns empty array when no sites are followed', async () => {
      vi.mocked(repo.listSitesAsync).mockResolvedValue([]);

      const result = await tools.listSites();

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.sites).toEqual([]);
    });
  });

  describe('searchSites', () => {
    it('searches sites with a query', async () => {
      const mockSites = [
        { id: 'si_cccccccccccccc', name: 'Marketing', webUrl: 'https://contoso.sharepoint.com/sites/marketing', displayName: 'Marketing Hub' },
      ];
      vi.mocked(repo.searchSitesAsync).mockResolvedValue(mockSites);

      const result = await tools.searchSites({ query: 'marketing' });

      expect(repo.searchSitesAsync).toHaveBeenCalledWith('marketing');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.sites).toEqual(mockSites);
    });

    it('returns empty results for no matches', async () => {
      vi.mocked(repo.searchSitesAsync).mockResolvedValue([]);

      const result = await tools.searchSites({ query: 'nonexistent' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.sites).toEqual([]);
    });
  });

  describe('getSite', () => {
    it('returns site details', async () => {
      const mockSite = {
        id: 'si_aaaaaaaaaaaaaa', name: 'Team Site', webUrl: 'https://contoso.sharepoint.com/sites/team',
        displayName: 'Team Site', description: 'Main collaboration site',
      };
      vi.mocked(repo.getSiteAsync).mockResolvedValue(mockSite);

      const result = await tools.getSite({ site_id: 'si_aaaaaaaaaaaaaa' });

      expect(repo.getSiteAsync).toHaveBeenCalledWith('si_aaaaaaaaaaaaaa');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.site).toEqual(mockSite);
      expect(parsed.site.description).toBe('Main collaboration site');
    });
  });

  // ===========================================================================
  // Document Libraries
  // ===========================================================================

  describe('listDocumentLibraries', () => {
    it('returns document libraries for a site', async () => {
      const mockLibraries = [
        { id: 'dl_dddddddddddddd', name: 'Documents', webUrl: 'https://contoso.sharepoint.com/sites/team/Shared%20Documents', driveType: 'documentLibrary' },
        { id: 'dl_eeeeeeeeeeeeee', name: 'Site Assets', webUrl: 'https://contoso.sharepoint.com/sites/team/SiteAssets', driveType: 'documentLibrary' },
      ];
      vi.mocked(repo.listDocumentLibrariesAsync).mockResolvedValue(mockLibraries);

      const result = await tools.listDocumentLibraries({ site_id: 'si_aaaaaaaaaaaaaa' });

      expect(repo.listDocumentLibrariesAsync).toHaveBeenCalledWith('si_aaaaaaaaaaaaaa');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.libraries).toEqual(mockLibraries);
    });
  });

  // ===========================================================================
  // Library Items
  // ===========================================================================

  describe('listLibraryItems', () => {
    it('returns items from root of library', async () => {
      const mockItems = [
        { id: 'li_ffffffffffffff', name: 'Report.docx', size: 15000, webUrl: 'https://contoso.sharepoint.com/sites/team/Shared%20Documents/Report.docx', lastModifiedDateTime: '2026-03-01T10:00:00Z', isFolder: false },
        { id: 'li_gggggggggggggg', name: 'Projects', size: 0, webUrl: 'https://contoso.sharepoint.com/sites/team/Shared%20Documents/Projects', lastModifiedDateTime: '2026-02-15T08:00:00Z', isFolder: true },
      ];
      vi.mocked(repo.listLibraryItemsAsync).mockResolvedValue(mockItems);

      const result = await tools.listLibraryItems({ library_id: 'dl_dddddddddddddd' });

      expect(repo.listLibraryItemsAsync).toHaveBeenCalledWith('dl_dddddddddddddd', undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.items).toEqual(mockItems);
    });

    it('returns items from a subfolder when folder_id is provided', async () => {
      const mockItems = [
        { id: 'li_hhhhhhhhhhhhhh', name: 'Proposal.pptx', size: 50000, webUrl: 'https://contoso.sharepoint.com/sites/team/Shared%20Documents/Projects/Proposal.pptx', lastModifiedDateTime: '2026-03-05T14:00:00Z', isFolder: false },
      ];
      vi.mocked(repo.listLibraryItemsAsync).mockResolvedValue(mockItems);

      const result = await tools.listLibraryItems({ library_id: 'dl_dddddddddddddd', folder_id: 'li_gggggggggggggg' });

      expect(repo.listLibraryItemsAsync).toHaveBeenCalledWith('dl_dddddddddddddd', 'li_gggggggggggggg');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.items).toEqual(mockItems);
    });

    it('returns empty array for empty folder', async () => {
      vi.mocked(repo.listLibraryItemsAsync).mockResolvedValue([]);

      const result = await tools.listLibraryItems({ library_id: 'dl_dddddddddddddd' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.items).toEqual([]);
    });
  });

  // ===========================================================================
  // Download
  // ===========================================================================

  describe('downloadLibraryFile', () => {
    it('downloads file and returns saved path', async () => {
      vi.mocked(repo.downloadLibraryFileAsync).mockResolvedValue('/tmp/downloads/Report.docx');

      const result = await tools.downloadLibraryFile({ item_id: 'li_ffffffffffffff', output_path: '/tmp/downloads/Report.docx' });

      expect(repo.downloadLibraryFileAsync).toHaveBeenCalledWith('li_ffffffffffffff', '/tmp/downloads/Report.docx');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.path).toBe('/tmp/downloads/Report.docx');
      expect(parsed.message).toBe('File downloaded');
    });
  });

  // ===========================================================================
  // Create Folder
  // ===========================================================================

  describe('createLibraryFolder', () => {
    it('creates a folder at the library root with default conflict behavior', async () => {
      const created = { id: 'li_newfolder00001', name: 'Candidates', webUrl: 'https://contoso.sharepoint.com/sites/hr/Shared%20Documents/Candidates', isFolder: true };
      vi.mocked(repo.createLibraryFolderAsync).mockResolvedValue(created);

      const result = await tools.createLibraryFolder({ library_id: 'dl_aaaaaaaaaaaaaa', folder_name: 'Candidates', conflict_behavior: 'fail' });

      expect(repo.createLibraryFolderAsync).toHaveBeenCalledWith('dl_aaaaaaaaaaaaaa', undefined, 'Candidates', 'fail');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.folder).toEqual(created);
      expect(parsed.message).toContain('Candidates');
    });

    it('creates a folder inside a parent folder with rename conflict behavior', async () => {
      const created = { id: 'li_newfolder00002', name: 'Jane Doe', webUrl: 'https://contoso.sharepoint.com/x', isFolder: true };
      vi.mocked(repo.createLibraryFolderAsync).mockResolvedValue(created);

      const result = await tools.createLibraryFolder({ library_id: 'dl_aaaaaaaaaaaaaa', parent_folder_id: 'li_newfolder00001', folder_name: 'Jane Doe', conflict_behavior: 'rename' });

      expect(repo.createLibraryFolderAsync).toHaveBeenCalledWith('dl_aaaaaaaaaaaaaa', 'li_newfolder00001', 'Jane Doe', 'rename');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.folder.id).toBe('li_newfolder00002');
    });
  });

  // ===========================================================================
  // Upload File (prepare / confirm)
  // ===========================================================================

  describe('prepareUploadLibraryFile', () => {
    it('returns an approval token with upload metadata', () => {
      const result = tools.prepareUploadLibraryFile({
        library_id: 'dl_aaaaaaaaaaaaaa',
        parent_folder_id: 'li_newfolder00001',
        file_name: 'resume.pdf',
        local_file_path: '/tmp/resume.pdf',
        conflict_behavior: 'fail',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.library_id).toBe('dl_aaaaaaaaaaaaaa');
      expect(parsed.parent_folder_id).toBe('li_newfolder00001');
      expect(parsed.file_name).toBe('resume.pdf');
      expect(parsed.action).toContain('confirm_upload_library_file');
    });
  });

  describe('confirmUploadLibraryFile', () => {
    it('uploads the file using stored metadata', async () => {
      const uploaded = { id: 'li_uploaded000001', name: 'resume.pdf', webUrl: 'https://contoso.sharepoint.com/y', size: 2048 };
      vi.mocked(repo.uploadLibraryFileAsync).mockResolvedValue(uploaded);

      const prepareResult = tools.prepareUploadLibraryFile({
        library_id: 'dl_aaaaaaaaaaaaaa',
        parent_folder_id: 'li_newfolder00001',
        file_name: 'resume.pdf',
        local_file_path: '/tmp/resume.pdf',
        conflict_behavior: 'fail',
      });
      const token = JSON.parse(prepareResult.content[0].text).approval_token;

      const result = await tools.confirmUploadLibraryFile({ approval_token: token });

      expect(repo.uploadLibraryFileAsync).toHaveBeenCalledWith('dl_aaaaaaaaaaaaaa', 'li_newfolder00001', 'resume.pdf', '/tmp/resume.pdf', 'fail');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.item).toEqual(uploaded);
    });

    it('uploads at the library root when no parent folder is given', async () => {
      const uploaded = { id: 'li_uploaded000002', name: 'notes.txt', webUrl: 'https://contoso.sharepoint.com/z', size: 12 };
      vi.mocked(repo.uploadLibraryFileAsync).mockResolvedValue(uploaded);

      const prepareResult = tools.prepareUploadLibraryFile({
        library_id: 'dl_aaaaaaaaaaaaaa',
        file_name: 'notes.txt',
        local_file_path: '/tmp/notes.txt',
        conflict_behavior: 'replace',
      });
      const token = JSON.parse(prepareResult.content[0].text).approval_token;

      await tools.confirmUploadLibraryFile({ approval_token: token });

      expect(repo.uploadLibraryFileAsync).toHaveBeenCalledWith('dl_aaaaaaaaaaaaaa', undefined, 'notes.txt', '/tmp/notes.txt', 'replace');
    });

    it('rejects an invalid token', async () => {
      const result = await tools.confirmUploadLibraryFile({ approval_token: 'invalid-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toContain('not found');
      expect(repo.uploadLibraryFileAsync).not.toHaveBeenCalled();
    });

    it('rejects an already consumed token', async () => {
      const uploaded = { id: 'li_uploaded000001', name: 'resume.pdf', webUrl: 'https://contoso.sharepoint.com/y', size: 2048 };
      vi.mocked(repo.uploadLibraryFileAsync).mockResolvedValue(uploaded);

      const prepareResult = tools.prepareUploadLibraryFile({
        library_id: 'dl_aaaaaaaaaaaaaa',
        file_name: 'resume.pdf',
        local_file_path: '/tmp/resume.pdf',
        conflict_behavior: 'fail',
      });
      const token = JSON.parse(prepareResult.content[0].text).approval_token;

      await tools.confirmUploadLibraryFile({ approval_token: token });
      const result = await tools.confirmUploadLibraryFile({ approval_token: token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });
});
