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
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolContext, ToolDefinition } from '../registry/types.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    sharePoint: SharePointTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const ListSitesInput = z.strictObject({});

export const SearchSitesInput = z.strictObject({
  query: z.string().min(1).describe('Search keyword for SharePoint sites'),
});

export const GetSiteInput = z.strictObject({
  site_id: z.string().min(1).describe('Site ID (si_ token from list_sites or search_sites)'),
});

export const ListDocumentLibrariesInput = z.strictObject({
  site_id: z.string().min(1).describe('Site ID (si_ token from list_sites or search_sites)'),
});

export const ListLibraryItemsInput = z.strictObject({
  library_id: z.string().min(1).describe('Library ID (dl_ token from list_document_libraries)'),
  folder_id: z.string().min(1).optional().describe('Folder ID to browse into (li_ token from a previous list_library_items call)'),
});

export const DownloadLibraryFileInput = z.strictObject({
  item_id: z.string().min(1).describe('Item ID (li_ token from list_library_items)'),
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
  listSitesAsync(): Promise<Array<{ id: string; name: string; webUrl: string; displayName: string }>>;
  searchSitesAsync(query: string): Promise<Array<{ id: string; name: string; webUrl: string; displayName: string }>>;
  getSiteAsync(siteId: string | number): Promise<{ id: string; name: string; webUrl: string; displayName: string; description: string }>;
  listDocumentLibrariesAsync(siteId: string | number): Promise<Array<{ id: string; name: string; webUrl: string; driveType: string }>>;
  listLibraryItemsAsync(libraryId: string | number, folderId?: string | number): Promise<Array<{
    id: string; name: string; size: number; webUrl: string;
    lastModifiedDateTime: string; isFolder: boolean;
  }>>;
  downloadLibraryFileAsync(itemId: string | number, outputPath: string): Promise<string>;
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

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2)
// =============================================================================

/**
 * Registry tool definitions for the sharepoint domain.
 */
export function sharePointToolDefinitions(): ToolDefinition[] {
  const tools = (ctx: ToolContext): SharePointTools => requireGraphToolset(ctx, 'sharePoint');

  return [
    defineTool({
      name: 'list_sites',
      description: 'List SharePoint sites the user follows (Graph API)',
      input: ListSitesInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx) => tools(ctx).listSites(),
    }),
    defineTool({
      name: 'search_sites',
      description: 'Search for SharePoint sites by keyword (Graph API)',
      input: SearchSitesInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).searchSites(params),
    }),
    defineTool({
      name: 'get_site',
      description: 'Get details for a specific SharePoint site (Graph API)',
      input: GetSiteInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getSite(params),
    }),
    defineTool({
      name: 'list_document_libraries',
      description: 'List document libraries (drives) for a SharePoint site (Graph API)',
      input: ListDocumentLibrariesInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listDocumentLibraries(params),
    }),
    defineTool({
      name: 'list_library_items',
      description: 'List files and folders in a document library or subfolder (Graph API)',
      input: ListLibraryItemsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listLibraryItems(params),
    }),
    defineTool({
      name: 'download_library_file',
      description: 'Download a file from a SharePoint document library to a local path (Graph API)',
      input: DownloadLibraryFileInput,
      // Writes the file to output_path on local disk — not read-only.
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).downloadLibraryFile(params),
    }),
  ];
}
