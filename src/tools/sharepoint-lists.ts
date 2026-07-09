/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * SharePoint Lists MCP tools (issue #38).
 *
 * Tools for browsing and editing SharePoint lists, their columns, and their
 * items. Lists live under a site (durable sl_ token carries {siteId, listId});
 * items live under a list (durable sn_ token carries {siteId, listId, itemId}).
 * Item deletion uses the two-phase approval pattern.
 */

import { z } from 'zod';
import { Id } from '../ids/schema.js';
import { nextActionFor } from '../ids/next-action.js';
import type { ApprovalTokenManager } from '../approval/index.js';
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolContext, ToolDefinition, ToolResult } from '../registry/types.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    sharePointLists: SharePointListsTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const ListListsInput = z.strictObject({
  site_id: Id.site,
});

export const GetListInput = z.strictObject({
  list_id: Id.sharePointList,
});

export const CreateListInput = z.strictObject({
  site_id: Id.site,
  display_name: z.string().min(1).describe('Display name for the new list'),
  description: z.string().optional().describe('Optional description for the list'),
});

export const ListListColumnsInput = z.strictObject({
  list_id: Id.sharePointList,
});

export const ListListItemsInput = z.strictObject({
  list_id: Id.sharePointList,
  limit: z.number().int().min(1).max(200).default(50).describe('Maximum items to return (1-200)'),
});

export const GetListItemInput = z.strictObject({
  item_id: Id.sharePointListItem,
});

export const CreateListItemInput = z.strictObject({
  list_id: Id.sharePointList,
  fields: z.record(z.string(), z.unknown()).describe('Column name → value map for the new item (e.g. { "Title": "New row", "Status": "Open" })'),
});

export const UpdateListItemInput = z.strictObject({
  item_id: Id.sharePointListItem,
  fields: z.record(z.string(), z.unknown()).describe('Column name → value map of fields to update'),
});

export const PrepareDeleteListItemInput = z.strictObject({
  item_id: Id.sharePointListItem,
});

export const ConfirmDeleteListItemInput = z.strictObject({
  token_id: z.string().uuid().describe('Approval token from prepare_delete_list_item'),
  item_id: Id.sharePointListItem,
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListListsParams = z.infer<typeof ListListsInput>;
export type GetListParams = z.infer<typeof GetListInput>;
export type CreateListParams = z.infer<typeof CreateListInput>;
export type ListListColumnsParams = z.infer<typeof ListListColumnsInput>;
export type ListListItemsParams = z.infer<typeof ListListItemsInput>;
export type GetListItemParams = z.infer<typeof GetListItemInput>;
export type CreateListItemParams = z.infer<typeof CreateListItemInput>;
export type UpdateListItemParams = z.infer<typeof UpdateListItemInput>;
export type PrepareDeleteListItemParams = z.infer<typeof PrepareDeleteListItemInput>;
export type ConfirmDeleteListItemParams = z.infer<typeof ConfirmDeleteListItemInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface ISharePointListsRepository {
  listSharePointListsAsync(siteId: string | number): Promise<Array<{
    id: string; name: string; displayName: string; description: string; webUrl: string;
  }>>;
  getSharePointListAsync(listId: string | number): Promise<{
    id: string; name: string; displayName: string; description: string; webUrl: string;
  }>;
  createSharePointListAsync(siteId: string | number, displayName: string, description?: string): Promise<string>;
  listSharePointListColumnsAsync(listId: string | number): Promise<Array<{
    id: string; name: string; displayName: string; columnType: string; required: boolean; readOnly: boolean;
  }>>;
  listSharePointListItemsAsync(listId: string | number, limit?: number): Promise<Array<{
    id: string; fields: Record<string, unknown>; webUrl: string;
    createdDateTime: string; lastModifiedDateTime: string;
  }>>;
  getSharePointListItemAsync(itemId: string | number): Promise<{
    id: string; fields: Record<string, unknown>; webUrl: string;
    createdDateTime: string; lastModifiedDateTime: string;
  }>;
  createSharePointListItemAsync(listId: string | number, fields: Record<string, unknown>): Promise<string>;
  updateSharePointListItemAsync(itemId: string | number, fields: Record<string, unknown>): Promise<void>;
  deleteSharePointListItemAsync(itemId: string | number): Promise<void>;
}

// =============================================================================
// SharePoint Lists Tools
// =============================================================================

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * SharePoint list tools with two-phase approval for item deletion.
 */
export class SharePointListsTools {
  constructor(
    private readonly repository: ISharePointListsRepository,
    private readonly tokenManager: ApprovalTokenManager,
  ) {}

  async listLists(params: ListListsParams): Promise<ToolResult> {
    const lists = await this.repository.listSharePointListsAsync(params.site_id);
    return jsonResult({ lists, next: nextActionFor('sharePointList') ?? undefined });
  }

  async getList(params: GetListParams): Promise<ToolResult> {
    const list = await this.repository.getSharePointListAsync(params.list_id);
    return jsonResult({ list });
  }

  async createList(params: CreateListParams): Promise<ToolResult> {
    const listId = await this.repository.createSharePointListAsync(params.site_id, params.display_name, params.description);
    return jsonResult({ id: listId, display_name: params.display_name, status: 'created', next: nextActionFor('sharePointList') ?? undefined });
  }

  async listListColumns(params: ListListColumnsParams): Promise<ToolResult> {
    const columns = await this.repository.listSharePointListColumnsAsync(params.list_id);
    return jsonResult({ columns });
  }

  async listListItems(params: ListListItemsParams): Promise<ToolResult> {
    const items = await this.repository.listSharePointListItemsAsync(params.list_id, params.limit);
    return jsonResult({ items, next: nextActionFor('sharePointListItem') ?? undefined });
  }

  async getListItem(params: GetListItemParams): Promise<ToolResult> {
    const item = await this.repository.getSharePointListItemAsync(params.item_id);
    return jsonResult({ item });
  }

  async createListItem(params: CreateListItemParams): Promise<ToolResult> {
    const itemId = await this.repository.createSharePointListItemAsync(params.list_id, params.fields);
    return jsonResult({ id: itemId, status: 'created', next: nextActionFor('sharePointListItem') ?? undefined });
  }

  async updateListItem(params: UpdateListItemParams): Promise<ToolResult> {
    await this.repository.updateSharePointListItemAsync(params.item_id, params.fields);
    return jsonResult({ id: params.item_id, status: 'updated' });
  }

  prepareDeleteListItem(params: PrepareDeleteListItemParams): ToolResult {
    const token = this.tokenManager.generateToken({
      operation: 'delete_list_item',
      targetType: 'list_item',
      targetId: params.item_id,
      targetHash: String(params.item_id),
    });

    return jsonResult({
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      item_id: params.item_id,
      action: `To confirm deleting list item ${params.item_id}, call confirm_delete_list_item with the token_id and item_id.`,
    });
  }

  async confirmDeleteListItem(params: ConfirmDeleteListItemParams): Promise<ToolResult> {
    const validation = this.tokenManager.consumeToken(params.token_id, 'delete_list_item', params.item_id);
    if (!validation.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_list_item again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_list_item',
        TARGET_MISMATCH: 'Token was generated for a different list item',
        ALREADY_CONSUMED: 'Token has already been used',
      };
      return jsonResult({ success: false, error: errorMessages[validation.error ?? ''] ?? 'Invalid token' });
    }

    await this.repository.deleteSharePointListItemAsync(params.item_id);
    return jsonResult({ success: true, message: 'List item deleted' });
  }
}

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2)
// =============================================================================

/**
 * Registry tool definitions for the sharepoint-lists domain.
 */
export function sharePointListsToolDefinitions(): ToolDefinition[] {
  const tools = (ctx: ToolContext): SharePointListsTools => requireGraphToolset(ctx, 'sharePointLists');

  return [
    defineTool({
      name: 'list_lists',
      description: 'List the SharePoint lists in a site (Graph API)',
      input: ListListsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listLists(params),
    }),
    defineTool({
      name: 'get_list',
      description: 'Get details for a specific SharePoint list (Graph API)',
      input: GetListInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getList(params),
    }),
    defineTool({
      name: 'create_list',
      description: 'Create a new SharePoint list in a site (Graph API)',
      input: CreateListInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).createList(params),
    }),
    defineTool({
      name: 'list_list_columns',
      description: 'List the column definitions for a SharePoint list (Graph API)',
      input: ListListColumnsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listListColumns(params),
    }),
    defineTool({
      name: 'list_list_items',
      description: 'List the items in a SharePoint list, with their field values (Graph API)',
      input: ListListItemsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listListItems(params),
    }),
    defineTool({
      name: 'get_list_item',
      description: 'Get a specific SharePoint list item with its field values (Graph API)',
      input: GetListItemInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getListItem(params),
    }),
    defineTool({
      name: 'create_list_item',
      description: 'Create an item in a SharePoint list from a column name → value map (Graph API)',
      input: CreateListItemInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).createListItem(params),
    }),
    defineTool({
      name: 'update_list_item',
      description: 'Update the field values of a SharePoint list item (Graph API)',
      input: UpdateListItemInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).updateListItem(params),
    }),
    defineTool({
      name: 'prepare_delete_list_item',
      description: 'Prepare to delete a SharePoint list item. Returns an approval token. Call confirm_delete_list_item to execute. (Graph API)',
      input: PrepareDeleteListItemInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareDeleteListItem(params),
    }),
    defineTool({
      name: 'confirm_delete_list_item',
      description: 'Confirm SharePoint list item deletion with an approval token (Graph API)',
      input: ConfirmDeleteListItemInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['sharepoint'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmDeleteListItem(params),
    }),
  ];
}
