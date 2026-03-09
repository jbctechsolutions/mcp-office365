/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Excel Online (Workbook) tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ExcelTools, type IExcelRepository } from '../../../src/tools/excel.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';

describe('ExcelTools', () => {
  let repo: IExcelRepository;
  let tokenManager: ApprovalTokenManager;
  let tools: ExcelTools;

  beforeEach(() => {
    repo = {
      listWorksheetsAsync: vi.fn(),
      getWorksheetRangeAsync: vi.fn(),
      getUsedRangeAsync: vi.fn(),
      updateWorksheetRangeAsync: vi.fn(),
      getTableDataAsync: vi.fn(),
    };
    tokenManager = new ApprovalTokenManager();
    tools = new ExcelTools(repo, tokenManager);
  });

  // ===========================================================================
  // listWorksheets
  // ===========================================================================

  describe('listWorksheets', () => {
    it('returns worksheet list from the repository', async () => {
      const mockWorksheets = [
        { id: 'sheet1', name: 'Sheet1', position: 0 },
        { id: 'sheet2', name: 'Data', position: 1 },
      ];
      vi.mocked(repo.listWorksheetsAsync).mockResolvedValue(mockWorksheets);

      const result = await tools.listWorksheets({ file_id: 100 });

      expect(repo.listWorksheetsAsync).toHaveBeenCalledWith(100);
      expect(result.content).toHaveLength(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.worksheets).toEqual(mockWorksheets);
    });
  });

  // ===========================================================================
  // getWorksheetRange
  // ===========================================================================

  describe('getWorksheetRange', () => {
    it('returns cell data for the specified range', async () => {
      const mockRange = {
        address: 'Sheet1!A1:B2',
        values: [['Name', 'Age'], ['Alice', 30]],
      };
      vi.mocked(repo.getWorksheetRangeAsync).mockResolvedValue(mockRange);

      const result = await tools.getWorksheetRange({
        file_id: 100,
        worksheet_name: 'Sheet1',
        range: 'A1:B2',
      });

      expect(repo.getWorksheetRangeAsync).toHaveBeenCalledWith(100, 'Sheet1', 'A1:B2');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.range.values).toEqual([['Name', 'Age'], ['Alice', 30]]);
    });
  });

  // ===========================================================================
  // getUsedRange
  // ===========================================================================

  describe('getUsedRange', () => {
    it('returns all used data for the worksheet', async () => {
      const mockRange = {
        address: 'Sheet1!A1:C3',
        values: [['A', 'B', 'C'], [1, 2, 3], [4, 5, 6]],
      };
      vi.mocked(repo.getUsedRangeAsync).mockResolvedValue(mockRange);

      const result = await tools.getUsedRange({
        file_id: 100,
        worksheet_name: 'Sheet1',
      });

      expect(repo.getUsedRangeAsync).toHaveBeenCalledWith(100, 'Sheet1');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.range.values).toEqual([['A', 'B', 'C'], [1, 2, 3], [4, 5, 6]]);
    });
  });

  // ===========================================================================
  // getTableData
  // ===========================================================================

  describe('getTableData', () => {
    it('returns table rows from the repository', async () => {
      const mockRows = [
        { values: [['Alice', 30]] },
        { values: [['Bob', 25]] },
      ];
      vi.mocked(repo.getTableDataAsync).mockResolvedValue(mockRows);

      const result = await tools.getTableData({
        file_id: 100,
        table_name: 'EmployeeTable',
      });

      expect(repo.getTableDataAsync).toHaveBeenCalledWith(100, 'EmployeeTable');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.rows).toEqual(mockRows);
    });
  });

  // ===========================================================================
  // prepareUpdateRange
  // ===========================================================================

  describe('prepareUpdateRange', () => {
    it('generates an approval token with range info', () => {
      const result = tools.prepareUpdateRange({
        file_id: 100,
        worksheet_name: 'Sheet1',
        range: 'A1:B2',
        values: [['X', 'Y'], [1, 2]],
      });

      expect(result.content).toHaveLength(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.file_id).toBe(100);
      expect(parsed.worksheet_name).toBe('Sheet1');
      expect(parsed.range).toBe('A1:B2');
      expect(parsed.cell_count).toBe(4);
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.action).toContain('confirm_update_range');
    });

    it('correctly counts cells in values array', () => {
      const result = tools.prepareUpdateRange({
        file_id: 100,
        worksheet_name: 'Data',
        range: 'A1:C1',
        values: [['a', 'b', 'c']],
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.cell_count).toBe(3);
    });
  });

  // ===========================================================================
  // confirmUpdateRange
  // ===========================================================================

  describe('confirmUpdateRange', () => {
    it('updates cells with a valid token', async () => {
      vi.mocked(repo.updateWorksheetRangeAsync).mockResolvedValue({});

      // Generate a token first
      const prepareResult = tools.prepareUpdateRange({
        file_id: 100,
        worksheet_name: 'Sheet1',
        range: 'A1:B2',
        values: [['X', 'Y'], [1, 2]],
      });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      // Now confirm
      const result = await tools.confirmUpdateRange({ approval_token });

      expect(repo.updateWorksheetRangeAsync).toHaveBeenCalledWith(
        100,
        'Sheet1',
        'A1:B2',
        [['X', 'Y'], [1, 2]],
      );
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toContain('A1:B2');
      expect(parsed.message).toContain('Sheet1');
    });

    it('rejects an invalid token', async () => {
      const result = await tools.confirmUpdateRange({
        approval_token: 'invalid-token-abc',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
      expect(repo.updateWorksheetRangeAsync).not.toHaveBeenCalled();
    });

    it('rejects a token that has already been consumed', async () => {
      vi.mocked(repo.updateWorksheetRangeAsync).mockResolvedValue({});

      // Generate and consume a token
      const prepareResult = tools.prepareUpdateRange({
        file_id: 100,
        worksheet_name: 'Sheet1',
        range: 'A1:A1',
        values: [['done']],
      });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);
      await tools.confirmUpdateRange({ approval_token });

      // Try to use it again
      const result = await tools.confirmUpdateRange({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBeDefined();
    });

    it('rejects an expired token', async () => {
      // Create a token manager with very short TTL for testing
      const shortTtlManager = new ApprovalTokenManager(1); // 1ms TTL
      const shortTools = new ExcelTools(repo, shortTtlManager);

      const prepareResult = shortTools.prepareUpdateRange({
        file_id: 100,
        worksheet_name: 'Sheet1',
        range: 'A1:A1',
        values: [['expired']],
      });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      // Wait for token to expire
      await new Promise(resolve => setTimeout(resolve, 10));

      const result = await shortTools.confirmUpdateRange({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toContain('expired');
    });
  });
});
