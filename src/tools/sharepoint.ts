/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * SharePoint Document Library MCP tools.
 *
 * Provides read-only tools for browsing SharePoint sites, document libraries,
 * and downloading files from team document libraries.
 */

import { z } from 'zod';

// =============================================================================
// Input Schemas
// =============================================================================

export const ListSitesInput = z.strictObject({});

export const SearchSitesInput = z.strictObject({
  query: z.string().min(1).describe('Search keyword for SharePoint sites'),
});

export const GetSiteInput = z.strictObject({
  site_id: z.number().int().positive().describe('Site ID from list_sites or search_sites'),
});

export const ListDocumentLibrariesInput = z.strictObject({
  site_id: z.number().int().positive().describe('Site ID from list_sites or search_sites'),
});

export const ListLibraryItemsInput = z.strictObject({
  library_id: z.number().int().positive().describe('Library ID from list_document_libraries'),
  folder_id: z.number().int().positive().optional().describe('Folder ID to browse into (from a previous list_library_items call)'),
});

export const DownloadLibraryFileInput = z.strictObject({
  item_id: z.number().int().positive().describe('Item ID from list_library_items'),
  output_path: z.string().min(1).describe('Local file path to save the downloaded file'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListSitesParams = z.infer<typeof ListSitesInput>;
export type SearchSitesParams = z.infer<typeof SearchSitesInput>;
export type GetSiteParams = z.infer<typeof GetSiteInput>;
export type ListDocumentLibrariesParams = z.infer<typeof ListDocumentLibrariesInput>;
export type ListLibraryItemsParams = z.infer<typeof ListLibraryItemsInput>;
export type DownloadLibraryFileParams = z.infer<typeof DownloadLibraryFileInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface ISharePointRepository {
  listSitesAsync(): Promise<Array<{ id: number; name: string; webUrl: string; displayName: string }>>;
  searchSitesAsync(query: string): Promise<Array<{ id: number; name: string; webUrl: string; displayName: string }>>;
  getSiteAsync(siteId: number): Promise<{ id: number; name: string; webUrl: string; displayName: string; description: string }>;
  listDocumentLibrariesAsync(siteId: number): Promise<Array<{ id: number; name: string; webUrl: string; driveType: string }>>;
  listLibraryItemsAsync(libraryId: number, folderId?: number): Promise<Array<{
    id: number; name: string; size: number; webUrl: string;
    lastModifiedDateTime: string; isFolder: boolean;
  }>>;
  downloadLibraryFileAsync(itemId: number, outputPath: string): Promise<string>;
}

// =============================================================================
// SharePoint Tools
// =============================================================================

/**
 * SharePoint document library tools for browsing sites, libraries, and files.
 */
export class SharePointTools {
  constructor(
    private readonly repo: ISharePointRepository,
  ) {}

  async listSites(): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const sites = await this.repo.listSitesAsync();
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ sites }, null, 2),
      }],
    };
  }

  async searchSites(params: SearchSitesParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const sites = await this.repo.searchSitesAsync(params.query);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ sites }, null, 2),
      }],
    };
  }

  async getSite(params: GetSiteParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const site = await this.repo.getSiteAsync(params.site_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ site }, null, 2),
      }],
    };
  }

  async listDocumentLibraries(params: ListDocumentLibrariesParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const libraries = await this.repo.listDocumentLibrariesAsync(params.site_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ libraries }, null, 2),
      }],
    };
  }

  async listLibraryItems(params: ListLibraryItemsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const items = await this.repo.listLibraryItemsAsync(params.library_id, params.folder_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ items }, null, 2),
      }],
    };
  }

  async downloadLibraryFile(params: DownloadLibraryFileParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const savedPath = await this.repo.downloadLibraryFileAsync(params.item_id, params.output_path);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, path: savedPath, message: 'File downloaded' }, null, 2),
      }],
    };
  }
}
