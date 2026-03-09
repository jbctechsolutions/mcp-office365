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

// =============================================================================
// Input Schemas
// =============================================================================

export const ListWorksheetsInput = z.strictObject({
  file_id: z.number().describe('Numeric ID of the Excel file (from OneDrive or SharePoint)'),
});

export const GetWorksheetRangeInput = z.strictObject({
  file_id: z.number().describe('Numeric ID of the Excel file (from OneDrive or SharePoint)'),
  worksheet_name: z.string().describe('Name of the worksheet'),
  range: z.string().describe('Cell range e.g. "A1:D10"'),
});

export const GetUsedRangeInput = z.strictObject({
  file_id: z.number().describe('Numeric ID of the Excel file (from OneDrive or SharePoint)'),
  worksheet_name: z.string().describe('Name of the worksheet'),
});

export const PrepareUpdateRangeInput = z.strictObject({
  file_id: z.number().describe('Numeric ID of the Excel file (from OneDrive or SharePoint)'),
  worksheet_name: z.string().describe('Name of the worksheet'),
  range: z.string().describe('Cell range e.g. "A1:D10"'),
  values: z.array(z.array(z.unknown())).describe('2D array of cell values'),
});

export const ConfirmUpdateRangeInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_update_range'),
});

export const GetTableDataInput = z.strictObject({
  file_id: z.number().describe('Numeric ID of the Excel file (from OneDrive or SharePoint)'),
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
  listWorksheetsAsync(fileId: number): Promise<any[]>;
  getWorksheetRangeAsync(fileId: number, worksheetName: string, range: string): Promise<any>;
  getUsedRangeAsync(fileId: number, worksheetName: string): Promise<any>;
  updateWorksheetRangeAsync(fileId: number, worksheetName: string, range: string, values: unknown[][]): Promise<any>;
  getTableDataAsync(fileId: number, tableName: string): Promise<any>;
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

    await this.repo.updateWorksheetRangeAsync(result.token!.targetId, worksheetName, range, values);
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
