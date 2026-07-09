/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Graph-backend contact folder tools (v3 registry-driven architecture, U2).
 * Holds the contact folder logic that previously lived inline in the
 * `handleGraphToolCall` switch, with a two-phase approval pattern for the
 * destructive delete operation.
 */

import { z } from 'zod';
import { Id } from '../ids/schema.js';
import type { GraphRepository } from '../graph/repository.js';
import type { ApprovalTokenManager } from '../approval/index.js';
import { defineTool } from '../registry/define-tool.js';
import { tokenIdLink } from '../registry/elicit-links.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolContext, ToolDefinition, ToolResult } from '../registry/types.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    contactFolders: GraphContactFoldersTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const ListContactFoldersInput = z.strictObject({});

export const CreateContactFolderInput = z.strictObject({
  name: z.string().min(1).describe('Contact folder name'),
});

export const PrepareDeleteContactFolderInput = z.strictObject({
  folder_id: Id.contactFolder,
});

export const ConfirmDeleteContactFolderInput = z.strictObject({
  token_id: z.string().uuid().describe('Approval token from prepare_delete_contact_folder'),
  folder_id: Id.contactFolder,
});

// =============================================================================
// Type Exports
// =============================================================================

export type CreateContactFolderParams = z.infer<typeof CreateContactFolderInput>;
export type PrepareDeleteContactFolderParams = z.infer<typeof PrepareDeleteContactFolderInput>;
export type ConfirmDeleteContactFolderParams = z.infer<typeof ConfirmDeleteContactFolderInput>;

// =============================================================================
// Contact Folders Tools
// =============================================================================

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Contact folder tools with two-phase approval for delete operations.
 */
export class GraphContactFoldersTools {
  constructor(
    private readonly repository: GraphRepository,
    private readonly tokenManager: ApprovalTokenManager
  ) {}

  async listContactFolders(): Promise<ToolResult> {
    const folders = await this.repository.listContactFoldersAsync();
    return jsonResult({ contact_folders: folders });
  }

  async createContactFolder(params: CreateContactFolderParams): Promise<ToolResult> {
    const folderId = await this.repository.createContactFolderAsync(params.name);
    return jsonResult({ id: folderId, name: params.name, status: 'created' });
  }

  prepareDeleteContactFolder(params: PrepareDeleteContactFolderParams): ToolResult {
    const token = this.tokenManager.generateToken({
      operation: 'delete_contact_folder',
      targetType: 'contact_folder',
      targetId: params.folder_id,
      targetHash: String(params.folder_id),
    });

    return jsonResult({
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      folder_id: params.folder_id,
      action: `To confirm deleting contact folder ${params.folder_id}, call confirm_delete_contact_folder with the token_id and folder_id.`,
    });
  }

  async confirmDeleteContactFolder(params: ConfirmDeleteContactFolderParams): Promise<ToolResult> {
    const validation = this.tokenManager.consumeToken(params.token_id, 'delete_contact_folder', params.folder_id);
    if (!validation.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_contact_folder again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_contact_folder',
        TARGET_MISMATCH: 'Token was generated for a different contact folder',
        ALREADY_CONSUMED: 'Token has already been used',
      };
      return jsonResult({ success: false, error: errorMessages[validation.error ?? ''] ?? 'Invalid token' });
    }

    await this.repository.deleteContactFolderAsync(params.folder_id);
    return jsonResult({ success: true, message: 'Contact folder deleted' });
  }
}

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2)
// =============================================================================

/**
 * Registry tool definitions for the contact-folders domain.
 */
export function contactFoldersToolDefinitions(): ToolDefinition[] {
  const tools = (ctx: ToolContext): GraphContactFoldersTools => requireGraphToolset(ctx, 'contactFolders');

  return [
    defineTool({
      name: 'list_contact_folders',
      description: 'List all contact folders (Graph API)',
      input: ListContactFoldersInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['contacts'],
      backends: ['graph'],
      handler: (ctx) => tools(ctx).listContactFolders(),
    }),
    defineTool({
      name: 'create_contact_folder',
      description: 'Create a contact folder (Graph API)',
      input: CreateContactFolderInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['contacts'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).createContactFolder(params),
    }),
    defineTool({
      name: 'prepare_delete_contact_folder',
      description: 'Prepare to delete a contact folder. Returns an approval token. Call confirm_delete_contact_folder to execute. (Graph API)',
      input: PrepareDeleteContactFolderInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['contacts'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareDeleteContactFolder(params),
      onElicit: tokenIdLink('confirm_delete_contact_folder', ['folder_id']),
    }),
    defineTool({
      name: 'confirm_delete_contact_folder',
      description: 'Confirm contact folder deletion with approval token (Graph API)',
      input: ConfirmDeleteContactFolderInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['contacts'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmDeleteContactFolder(params),
    }),
  ];
}
