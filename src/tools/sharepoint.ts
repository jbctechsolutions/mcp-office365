/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * SharePoint Document Library MCP tools.
 *
 * Provides tools for browsing SharePoint sites and document libraries,
 * downloading files, and writing into team document libraries (creating
 * folders and uploading files) with two-phase approval for uploads.
 */

import { z } from 'zod';
import type { ApprovalTokenManager } from '../approval/index.js';
import { Id } from '../ids/schema.js';
import { nextActionFor } from '../ids/next-action.js';
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
  site_id: Id.site,
});

export const ListDocumentLibrariesInput = z.strictObject({
  site_id: Id.site,
});

export const ListLibraryItemsInput = z.strictObject({
  library_id: Id.documentLibrary,
  folder_id: Id.libraryDriveItem.optional().describe('Folder ID to browse into — a li_ token from a previous list_library_items call.'),
});

export const DownloadLibraryFileInput = z.strictObject({
  item_id: Id.libraryDriveItem,
  output_path: z.string().min(1).describe('Local file path to save the downloaded file'),
});

export const CreateLibraryFolderInput = z.strictObject({
  library_id: Id.documentLibrary,
  parent_folder_id: Id.libraryDriveItem.optional().describe('Folder to create inside — a li_ token from list_library_items. Omit to create at the library root.'),
  folder_name: z.string().min(1).describe('Name for the new folder'),
  conflict_behavior: z.enum(['fail', 'rename']).default('fail').describe('How to handle a name collision: fail (default) or rename'),
});

export const PrepareUploadLibraryFileInput = z.strictObject({
  library_id: Id.documentLibrary,
  parent_folder_id: Id.libraryDriveItem.optional().describe('Folder to upload into — a li_ token from list_library_items. Omit to upload at the library root.'),
  file_name: z.string().min(1).describe('Name for the file in the document library'),
  local_file_path: z.string().min(1).describe('Absolute path to the local file to upload (simple upload, 4 MB limit)'),
  conflict_behavior: z.enum(['fail', 'replace', 'rename']).default('fail').describe('How to handle a name collision: fail (default), replace, or rename'),
});

export const ConfirmUploadLibraryFileInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_upload_library_file'),
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
export type CreateLibraryFolderParams = z.infer<typeof CreateLibraryFolderInput>;
export type PrepareUploadLibraryFileParams = z.infer<typeof PrepareUploadLibraryFileInput>;
export type ConfirmUploadLibraryFileParams = z.infer<typeof ConfirmUploadLibraryFileInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface ISharePointRepository {
  listSitesAsync(): Promise<Array<{ id: string; name: string; webUrl: string; displayName: string }>>;
  searchSitesAsync(query: string): Promise<Array<{ id: string; name: string; webUrl: string; displayName: string }>>;
  getSiteAsync(siteId: string): Promise<{ id: string; name: string; webUrl: string; displayName: string; description: string }>;
  listDocumentLibrariesAsync(siteId: string): Promise<Array<{ id: string; name: string; webUrl: string; driveType: string }>>;
  listLibraryItemsAsync(libraryId: string, folderId?: string): Promise<Array<{
    id: string; name: string; size: number; webUrl: string;
    lastModifiedDateTime: string; isFolder: boolean;
  }>>;
  downloadLibraryFileAsync(itemId: string, outputPath: string): Promise<string>;
  createLibraryFolderAsync(libraryId: string, parentFolderId: string | undefined, folderName: string, conflictBehavior: string): Promise<{
    id: string; name: string; webUrl: string; isFolder: boolean;
  }>;
  uploadLibraryFileAsync(libraryId: string, parentFolderId: string | undefined, fileName: string, localFilePath: string, conflictBehavior: string): Promise<{
    id: string; name: string; webUrl: string; size: number;
  }>;
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
    private readonly tokenManager: ApprovalTokenManager,
  ) {}

  async listSites(): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const sites = await this.repo.listSitesAsync();
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ sites, next: nextActionFor('site') ?? undefined }, null, 2),
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
        text: JSON.stringify({ sites, next: nextActionFor('site') ?? undefined }, null, 2),
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
        text: JSON.stringify({ libraries, next: nextActionFor('documentLibrary') ?? undefined }, null, 2),
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
        text: JSON.stringify({ items, next: nextActionFor('libraryDriveItem') ?? undefined }, null, 2),
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

  async createLibraryFolder(params: CreateLibraryFolderParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const folder = await this.repo.createLibraryFolderAsync(
      params.library_id,
      params.parent_folder_id,
      params.folder_name,
      params.conflict_behavior,
    );
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          success: true,
          folder,
          message: `Folder "${folder.name}" created`,
          next: nextActionFor('libraryDriveItem') ?? undefined,
        }, null, 2),
      }],
    };
  }

  prepareUploadLibraryFile(params: PrepareUploadLibraryFileParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'upload_library_file',
      targetType: 'library_item',
      targetId: params.library_id,
      targetHash: `${params.library_id}/${params.parent_folder_id ?? 'root'}/${params.file_name}`,
      metadata: {
        library_id: params.library_id,
        parent_folder_id: params.parent_folder_id,
        file_name: params.file_name,
        local_file_path: params.local_file_path,
        conflict_behavior: params.conflict_behavior,
      },
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          library_id: params.library_id,
          parent_folder_id: params.parent_folder_id,
          file_name: params.file_name,
          local_file_path: params.local_file_path,
          conflict_behavior: params.conflict_behavior,
          action: `To confirm uploading "${params.file_name}" to library ${params.library_id}, call confirm_upload_library_file with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmUploadLibraryFile(params: ConfirmUploadLibraryFileParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const token = this.tokenManager.lookupToken(params.approval_token);
    if (token == null) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({
            success: false,
            error: 'Token not found or already used',
          }, null, 2),
        }],
      };
    }

    const result = this.tokenManager.consumeToken(params.approval_token, 'upload_library_file', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_upload_library_file again.',
        OPERATION_MISMATCH: 'Token was not generated for upload_library_file',
        TARGET_MISMATCH: 'Token was generated for a different upload',
        ALREADY_CONSUMED: 'Token has already been used',
      };
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({
            success: false,
            error: errorMessages[result.error ?? ''] ?? 'Invalid token',
          }, null, 2),
        }],
      };
    }

    const metadata = result.token!.metadata as {
      library_id: string; parent_folder_id?: string; file_name: string;
      local_file_path: string; conflict_behavior: string;
    };
    const uploaded = await this.repo.uploadLibraryFileAsync(
      metadata.library_id,
      metadata.parent_folder_id,
      metadata.file_name,
      metadata.local_file_path,
      metadata.conflict_behavior,
    );
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          success: true,
          item: uploaded,
          message: `File "${uploaded.name}" uploaded to library ${metadata.library_id}`,
        }, null, 2),
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
    defineTool({
      name: 'create_library_folder',
      description: 'Create a folder in a SharePoint document library or subfolder. conflict_behavior defaults to "fail". (Graph API)',
      input: CreateLibraryFolderInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).createLibraryFolder(params),
    }),
    defineTool({
      name: 'prepare_upload_library_file',
      description: 'Prepare to upload a local file into a SharePoint document library. Returns an approval token. Simple upload (4 MB limit). (Graph API)',
      input: PrepareUploadLibraryFileInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareUploadLibraryFile(params),
    }),
    defineTool({
      name: 'confirm_upload_library_file',
      description: 'Confirm uploading a file into a SharePoint document library using the approval token from prepare_upload_library_file. (Graph API)',
      input: ConfirmUploadLibraryFileInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmUploadLibraryFile(params),
    }),
  ];
}
