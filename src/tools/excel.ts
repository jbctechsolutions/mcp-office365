/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Excel Online (Workbook) MCP tools.
 *
 * Provides tools for reading and updating Excel workbooks stored in OneDrive
 * or SharePoint via the Microsoft Graph API, with two-phase approval for
 * range update operations.
 */

import { z } from 'zod';
import type { ApprovalTokenManager } from '../approval/index.js';
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolContext, ToolDefinition } from '../registry/types.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    excel: ExcelTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const ListWorksheetsInput = z.strictObject({
  file_id: z.string().min(1).describe('Durable ID of the Excel file (dr_ token from list_drive_items/search_drive_items, or a raw Graph item id)'),
});

export const GetWorksheetRangeInput = z.strictObject({
  file_id: z.string().min(1).describe('Durable ID of the Excel file (dr_ token from list_drive_items/search_drive_items, or a raw Graph item id)'),
  worksheet_name: z.string().describe('Name of the worksheet'),
  range: z.string().describe('Cell range e.g. "A1:D10"'),
});

export const GetUsedRangeInput = z.strictObject({
  file_id: z.string().min(1).describe('Durable ID of the Excel file (dr_ token from list_drive_items/search_drive_items, or a raw Graph item id)'),
  worksheet_name: z.string().describe('Name of the worksheet'),
});

export const PrepareUpdateRangeInput = z.strictObject({
  file_id: z.string().min(1).describe('Durable ID of the Excel file (dr_ token from list_drive_items/search_drive_items, or a raw Graph item id)'),
  worksheet_name: z.string().describe('Name of the worksheet'),
  range: z.string().describe('Cell range e.g. "A1:D10"'),
  values: z.array(z.array(z.unknown())).describe('2D array of cell values'),
});

export const ConfirmUpdateRangeInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_update_range'),
});

export const GetTableDataInput = z.strictObject({
  file_id: z.string().min(1).describe('Durable ID of the Excel file (dr_ token from list_drive_items/search_drive_items, or a raw Graph item id)'),
  table_name: z.string().describe('Name of the Excel table'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListWorksheetsParams = z.infer<typeof ListWorksheetsInput>;
export type GetWorksheetRangeParams = z.infer<typeof GetWorksheetRangeInput>;
export type GetUsedRangeParams = z.infer<typeof GetUsedRangeInput>;
export type PrepareUpdateRangeParams = z.infer<typeof PrepareUpdateRangeInput>;
export type ConfirmUpdateRangeParams = z.infer<typeof ConfirmUpdateRangeInput>;
export type GetTableDataParams = z.infer<typeof GetTableDataInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface IExcelRepository {
  listWorksheetsAsync(fileId: string): Promise<Record<string, unknown>[]>;
  getWorksheetRangeAsync(fileId: string, worksheetName: string, range: string): Promise<Record<string, unknown>>;
  getUsedRangeAsync(fileId: string, worksheetName: string): Promise<Record<string, unknown>>;
  updateWorksheetRangeAsync(fileId: string, worksheetName: string, range: string, values: unknown[][]): Promise<Record<string, unknown>>;
  getTableDataAsync(fileId: string, tableName: string): Promise<Record<string, unknown>[]>;
}

// =============================================================================
// Excel Tools
// =============================================================================

/**
 * Excel Online tools with two-phase approval for range updates.
 */
export class ExcelTools {
  constructor(
    private readonly repo: IExcelRepository,
    private readonly tokenManager: ApprovalTokenManager,
  ) {}

  async listWorksheets(params: ListWorksheetsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const worksheets = await this.repo.listWorksheetsAsync(params.file_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ worksheets }, null, 2),
      }],
    };
  }

  async getWorksheetRange(params: GetWorksheetRangeParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const range = await this.repo.getWorksheetRangeAsync(params.file_id, params.worksheet_name, params.range);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ range }, null, 2),
      }],
    };
  }

  async getUsedRange(params: GetUsedRangeParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const range = await this.repo.getUsedRangeAsync(params.file_id, params.worksheet_name);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ range }, null, 2),
      }],
    };
  }

  async getTableData(params: GetTableDataParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const rows = await this.repo.getTableDataAsync(params.file_id, params.table_name);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ rows }, null, 2),
      }],
    };
  }

  prepareUpdateRange(params: PrepareUpdateRangeParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'update_excel_range',
      targetType: 'excel_range',
      targetId: params.file_id,
      targetHash: `${params.file_id}:${params.worksheet_name}:${params.range}`,
      metadata: {
        worksheet_name: params.worksheet_name,
        range: params.range,
        values: params.values,
      },
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          file_id: params.file_id,
          worksheet_name: params.worksheet_name,
          range: params.range,
          cell_count: params.values.reduce((sum, row) => sum + row.length, 0),
          action: `To confirm updating range ${params.range} in worksheet "${params.worksheet_name}", call confirm_update_range with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmUpdateRange(params: ConfirmUpdateRangeParams): Promise<{
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

    const result = this.tokenManager.consumeToken(params.approval_token, 'update_excel_range', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_update_range again.',
        OPERATION_MISMATCH: 'Token was not generated for update_excel_range',
        TARGET_MISMATCH: 'Token was generated for a different file',
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

    const metadata = result.token!.metadata;
    const worksheetName = metadata['worksheet_name'] as string;
    const range = metadata['range'] as string;
    const values = metadata['values'] as unknown[][];

    await this.repo.updateWorksheetRangeAsync((result.token!.targetId as string), worksheetName, range, values);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          success: true,
          message: `Range ${range} in worksheet "${worksheetName}" updated successfully`,
        }, null, 2),
      }],
    };
  }
}

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2)
// =============================================================================

/**
 * Registry tool definitions for the excel domain.
 */
export function excelToolDefinitions(): ToolDefinition[] {
  const tools = (ctx: ToolContext): ExcelTools => requireGraphToolset(ctx, 'excel');

  return [
    defineTool({
      name: 'list_worksheets',
      description: 'List all worksheets in an Excel workbook stored in OneDrive or SharePoint. (Graph API)',
      input: ListWorksheetsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['excel'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listWorksheets(params),
    }),
    defineTool({
      name: 'get_worksheet_range',
      description: 'Get cell values for a specific range in an Excel worksheet. (Graph API)',
      input: GetWorksheetRangeInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['excel'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getWorksheetRange(params),
    }),
    defineTool({
      name: 'get_used_range',
      description: 'Get all used data in an Excel worksheet (automatically detects the data region). (Graph API)',
      input: GetUsedRangeInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['excel'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getUsedRange(params),
    }),
    defineTool({
      name: 'prepare_update_range',
      description: 'Prepare to update cell values in an Excel worksheet range. Returns an approval token. (Graph API)',
      input: PrepareUpdateRangeInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['excel'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareUpdateRange(params),
    }),
    defineTool({
      name: 'confirm_update_range',
      description: 'Confirm updating cell values in an Excel worksheet range using the approval token from prepare_update_range. (Graph API)',
      input: ConfirmUpdateRangeInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['excel'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmUpdateRange(params),
    }),
    defineTool({
      name: 'get_table_data',
      description: 'Get rows from a named table in an Excel workbook. (Graph API)',
      input: GetTableDataInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['excel'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getTableData(params),
    }),
  ];
}
