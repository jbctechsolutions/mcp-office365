/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * OneDrive personal file MCP tools.
 *
 * Provides tools for managing OneDrive files and folders with two-phase
 * approval for destructive upload and delete operations.
 */

import { z } from 'zod';
import type { ApprovalTokenManager } from '../approval/index.js';

// =============================================================================
// Input Schemas
// =============================================================================

export const ListDriveItemsInput = z.strictObject({
  folder_id: z.number().int().positive().optional().describe('Folder ID from a previous list_drive_items call. Omit to list root.'),
});

export const SearchDriveItemsInput = z.strictObject({
  query: z.string().min(1).describe('Search query string'),
  limit: z.number().int().positive().optional().describe('Maximum results to return (default 25)'),
});

export const GetDriveItemInput = z.strictObject({
  item_id: z.number().int().positive().describe('Drive item ID from list_drive_items or search_drive_items'),
});

export const DownloadFileInput = z.strictObject({
  item_id: z.number().int().positive().describe('Drive item ID from list_drive_items or search_drive_items'),
  output_path: z.string().min(1).describe('Absolute file path where the file should be saved'),
});

export const PrepareUploadFileInput = z.strictObject({
  parent_path: z.string().min(1).describe('Parent folder path in OneDrive (e.g., "Documents/Reports")'),
  file_name: z.string().min(1).describe('Name for the file in OneDrive'),
  local_file_path: z.string().min(1).describe('Absolute path to the local file to upload'),
});

export const ConfirmUploadFileInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_upload_file'),
});

export const ListRecentFilesInput = z.strictObject({});

export const ListSharedWithMeInput = z.strictObject({});

export const CreateSharingLinkInput = z.strictObject({
  item_id: z.number().int().positive().describe('Drive item ID from list_drive_items or search_drive_items'),
  type: z.enum(['view', 'edit']).describe('Permission type: view (read-only) or edit (read-write)'),
  scope: z.enum(['anonymous', 'organization']).describe('Link scope: anonymous (anyone with link) or organization (org members only)'),
});

export const PrepareDeleteDriveItemInput = z.strictObject({
  item_id: z.number().int().positive().describe('Drive item ID from list_drive_items or search_drive_items'),
});

export const ConfirmDeleteDriveItemInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_drive_item'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListDriveItemsParams = z.infer<typeof ListDriveItemsInput>;
export type SearchDriveItemsParams = z.infer<typeof SearchDriveItemsInput>;
export type GetDriveItemParams = z.infer<typeof GetDriveItemInput>;
export type DownloadFileParams = z.infer<typeof DownloadFileInput>;
export type PrepareUploadFileParams = z.infer<typeof PrepareUploadFileInput>;
export type ConfirmUploadFileParams = z.infer<typeof ConfirmUploadFileInput>;
export type ListRecentFilesParams = z.infer<typeof ListRecentFilesInput>;
export type ListSharedWithMeParams = z.infer<typeof ListSharedWithMeInput>;
export type CreateSharingLinkParams = z.infer<typeof CreateSharingLinkInput>;
export type PrepareDeleteDriveItemParams = z.infer<typeof PrepareDeleteDriveItemInput>;
export type ConfirmDeleteDriveItemParams = z.infer<typeof ConfirmDeleteDriveItemInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface IOneDriveRepository {
  listDriveItemsAsync(folderId?: number): Promise<Array<{
    id: number; name: string; size: number; lastModified: string;
    isFolder: boolean; webUrl: string;
  }>>;
  searchDriveItemsAsync(query: string, limit?: number): Promise<Array<{
    id: number; name: string; size: number; lastModified: string;
    isFolder: boolean; webUrl: string;
  }>>;
  getDriveItemAsync(itemId: number): Promise<{
    id: number; name: string; size: number; lastModified: string;
    isFolder: boolean; webUrl: string; mimeType: string; createdBy: string;
  }>;
  downloadFileAsync(itemId: number, outputPath: string): Promise<{ savedPath: string; size: number }>;
  uploadFileAsync(parentPath: string, fileName: string, localFilePath: string): Promise<number>;
  listRecentFilesAsync(): Promise<Array<{
    id: number; name: string; size: number; lastModified: string;
    isFolder: boolean; webUrl: string;
  }>>;
  listSharedWithMeAsync(): Promise<Array<{
    id: number; name: string; size: number; lastModified: string;
    isFolder: boolean; webUrl: string;
  }>>;
  createSharingLinkAsync(itemId: number, type: string, scope: string): Promise<{
    webUrl: string; type: string; scope: string;
  }>;
  deleteDriveItemAsync(itemId: number): Promise<void>;
}

// =============================================================================
// OneDrive Tools
// =============================================================================

/**
 * OneDrive personal file tools with two-phase approval for upload and delete.
 */
export class OneDriveTools {
  constructor(
    private readonly repo: IOneDriveRepository,
    private readonly tokenManager: ApprovalTokenManager,
  ) {}

  async listDriveItems(params: ListDriveItemsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const items = await this.repo.listDriveItemsAsync(params.folder_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ items }, null, 2),
      }],
    };
  }

  async searchDriveItems(params: SearchDriveItemsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const items = await this.repo.searchDriveItemsAsync(params.query, params.limit);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ items }, null, 2),
      }],
    };
  }

  async getDriveItem(params: GetDriveItemParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const item = await this.repo.getDriveItemAsync(params.item_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ item }, null, 2),
      }],
    };
  }

  async downloadFile(params: DownloadFileParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const result = await this.repo.downloadFileAsync(params.item_id, params.output_path);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, ...result }, null, 2),
      }],
    };
  }

  prepareUploadFile(params: PrepareUploadFileParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'upload_file',
      targetType: 'drive_item',
      targetId: 0,
      targetHash: `${params.parent_path}/${params.file_name}`,
      metadata: {
        parent_path: params.parent_path,
        file_name: params.file_name,
        local_file_path: params.local_file_path,
      },
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          parent_path: params.parent_path,
          file_name: params.file_name,
          local_file_path: params.local_file_path,
          action: `To confirm uploading "${params.file_name}" to "${params.parent_path}", call confirm_upload_file with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmUploadFile(params: ConfirmUploadFileParams): Promise<{
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

    const result = this.tokenManager.consumeToken(params.approval_token, 'upload_file', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_upload_file again.',
        OPERATION_MISMATCH: 'Token was not generated for upload_file',
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

    const metadata = result.token!.metadata as { parent_path: string; file_name: string; local_file_path: string };
    const itemId = await this.repo.uploadFileAsync(metadata.parent_path, metadata.file_name, metadata.local_file_path);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          success: true,
          item_id: itemId,
          message: `File "${metadata.file_name}" uploaded to "${metadata.parent_path}"`,
        }, null, 2),
      }],
    };
  }

  async listRecentFiles(): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const items = await this.repo.listRecentFilesAsync();
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ items }, null, 2),
      }],
    };
  }

  async listSharedWithMe(): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const items = await this.repo.listSharedWithMeAsync();
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ items }, null, 2),
      }],
    };
  }

  async createSharingLink(params: CreateSharingLinkParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const link = await this.repo.createSharingLinkAsync(params.item_id, params.type, params.scope);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, link }, null, 2),
      }],
    };
  }

  prepareDeleteDriveItem(params: PrepareDeleteDriveItemParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'delete_drive_item',
      targetType: 'drive_item',
      targetId: params.item_id,
      targetHash: String(params.item_id),
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          item_id: params.item_id,
          action: `To confirm deleting drive item ${params.item_id}, call confirm_delete_drive_item with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmDeleteDriveItem(params: ConfirmDeleteDriveItemParams): Promise<{
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

    const result = this.tokenManager.consumeToken(params.approval_token, 'delete_drive_item', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_drive_item again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_drive_item',
        TARGET_MISMATCH: 'Token was generated for a different drive item',
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

    await this.repo.deleteDriveItemAsync(result.token!.targetId);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Drive item deleted' }, null, 2),
      }],
    };
  }
}
