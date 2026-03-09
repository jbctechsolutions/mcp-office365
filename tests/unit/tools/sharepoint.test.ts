/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for SharePoint Document Library tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { SharePointTools, type ISharePointRepository } from '../../../src/tools/sharepoint.js';

describe('SharePointTools', () => {
  let repo: ISharePointRepository;
  let tools: SharePointTools;

  beforeEach(() => {
    repo = {
      listSitesAsync: vi.fn(),
      searchSitesAsync: vi.fn(),
      getSiteAsync: vi.fn(),
      listDocumentLibrariesAsync: vi.fn(),
      listLibraryItemsAsync: vi.fn(),
      downloadLibraryFileAsync: vi.fn(),
    };
    tools = new SharePointTools(repo);
  });

  // ===========================================================================
  // Sites
  // ===========================================================================

  describe('listSites', () => {
    it('returns followed sites from the repository', async () => {
      const mockSites = [
        { id: 1, name: 'Team Site', webUrl: 'https://contoso.sharepoint.com/sites/team', displayName: 'Team Site' },
        { id: 2, name: 'HR Portal', webUrl: 'https://contoso.sharepoint.com/sites/hr', displayName: 'HR Portal' },
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
        { id: 3, name: 'Marketing', webUrl: 'https://contoso.sharepoint.com/sites/marketing', displayName: 'Marketing Hub' },
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
        id: 1, name: 'Team Site', webUrl: 'https://contoso.sharepoint.com/sites/team',
        displayName: 'Team Site', description: 'Main collaboration site',
      };
      vi.mocked(repo.getSiteAsync).mockResolvedValue(mockSite);

      const result = await tools.getSite({ site_id: 1 });

      expect(repo.getSiteAsync).toHaveBeenCalledWith(1);
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
        { id: 10, name: 'Documents', webUrl: 'https://contoso.sharepoint.com/sites/team/Shared%20Documents', driveType: 'documentLibrary' },
        { id: 11, name: 'Site Assets', webUrl: 'https://contoso.sharepoint.com/sites/team/SiteAssets', driveType: 'documentLibrary' },
      ];
      vi.mocked(repo.listDocumentLibrariesAsync).mockResolvedValue(mockLibraries);

      const result = await tools.listDocumentLibraries({ site_id: 1 });

      expect(repo.listDocumentLibrariesAsync).toHaveBeenCalledWith(1);
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
        { id: 100, name: 'Report.docx', size: 15000, webUrl: 'https://contoso.sharepoint.com/sites/team/Shared%20Documents/Report.docx', lastModifiedDateTime: '2026-03-01T10:00:00Z', isFolder: false },
        { id: 101, name: 'Projects', size: 0, webUrl: 'https://contoso.sharepoint.com/sites/team/Shared%20Documents/Projects', lastModifiedDateTime: '2026-02-15T08:00:00Z', isFolder: true },
      ];
      vi.mocked(repo.listLibraryItemsAsync).mockResolvedValue(mockItems);

      const result = await tools.listLibraryItems({ library_id: 10 });

      expect(repo.listLibraryItemsAsync).toHaveBeenCalledWith(10, undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.items).toEqual(mockItems);
    });

    it('returns items from a subfolder when folder_id is provided', async () => {
      const mockItems = [
        { id: 200, name: 'Proposal.pptx', size: 50000, webUrl: 'https://contoso.sharepoint.com/sites/team/Shared%20Documents/Projects/Proposal.pptx', lastModifiedDateTime: '2026-03-05T14:00:00Z', isFolder: false },
      ];
      vi.mocked(repo.listLibraryItemsAsync).mockResolvedValue(mockItems);

      const result = await tools.listLibraryItems({ library_id: 10, folder_id: 101 });

      expect(repo.listLibraryItemsAsync).toHaveBeenCalledWith(10, 101);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.items).toEqual(mockItems);
    });

    it('returns empty array for empty folder', async () => {
      vi.mocked(repo.listLibraryItemsAsync).mockResolvedValue([]);

      const result = await tools.listLibraryItems({ library_id: 10 });

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

      const result = await tools.downloadLibraryFile({ item_id: 100, output_path: '/tmp/downloads/Report.docx' });

      expect(repo.downloadLibraryFileAsync).toHaveBeenCalledWith(100, '/tmp/downloads/Report.docx');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.path).toBe('/tmp/downloads/Report.docx');
      expect(parsed.message).toBe('File downloaded');
    });
  });
});
