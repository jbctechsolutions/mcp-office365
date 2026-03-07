/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Checklist Items tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ChecklistItemsTools, type IChecklistItemsRepository } from '../../../src/tools/checklist-items.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';

describe('ChecklistItemsTools', () => {
  let repo: IChecklistItemsRepository;
  let tokenManager: ApprovalTokenManager;
  let tools: ChecklistItemsTools;

  beforeEach(() => {
    repo = {
      listChecklistItemsAsync: vi.fn(),
      createChecklistItemAsync: vi.fn(),
      updateChecklistItemAsync: vi.fn(),
      deleteChecklistItemAsync: vi.fn(),
    };
    tokenManager = new ApprovalTokenManager();
    tools = new ChecklistItemsTools(repo, tokenManager);
  });

  describe('listChecklistItems', () => {
    it('returns checklist items from the repository', async () => {
      const mockItems = [
        { id: 1, displayName: 'Buy milk', isChecked: false, createdDateTime: '2026-01-01T00:00:00Z' },
        { id: 2, displayName: 'Buy eggs', isChecked: true, createdDateTime: '2026-01-01T01:00:00Z' },
      ];
      vi.mocked(repo.listChecklistItemsAsync).mockResolvedValue(mockItems);

      const result = await tools.listChecklistItems({ task_id: 42 });

      expect(repo.listChecklistItemsAsync).toHaveBeenCalledWith(42);
      expect(result.content).toHaveLength(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.checklist_items).toEqual(mockItems);
    });
  });

  describe('createChecklistItem', () => {
    it('creates a checklist item and returns the ID', async () => {
      vi.mocked(repo.createChecklistItemAsync).mockResolvedValue(100);

      const result = await tools.createChecklistItem({ task_id: 42, display_name: 'Buy milk' });

      expect(repo.createChecklistItemAsync).toHaveBeenCalledWith(42, 'Buy milk', undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.checklist_item_id).toBe(100);
      expect(parsed.message).toBe('Checklist item created');
    });

    it('handles is_checked default (false)', async () => {
      vi.mocked(repo.createChecklistItemAsync).mockResolvedValue(101);

      await tools.createChecklistItem({ task_id: 42, display_name: 'Task A' });

      expect(repo.createChecklistItemAsync).toHaveBeenCalledWith(42, 'Task A', undefined);
    });

    it('passes is_checked when provided', async () => {
      vi.mocked(repo.createChecklistItemAsync).mockResolvedValue(102);

      await tools.createChecklistItem({ task_id: 42, display_name: 'Task B', is_checked: true });

      expect(repo.createChecklistItemAsync).toHaveBeenCalledWith(42, 'Task B', true);
    });
  });

  describe('updateChecklistItem', () => {
    it('updates a checklist item with display_name', async () => {
      vi.mocked(repo.updateChecklistItemAsync).mockResolvedValue(undefined);

      const result = await tools.updateChecklistItem({ checklist_item_id: 100, display_name: 'Updated text' });

      expect(repo.updateChecklistItemAsync).toHaveBeenCalledWith(100, { displayName: 'Updated text' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Checklist item updated');
    });

    it('updates a checklist item with is_checked', async () => {
      vi.mocked(repo.updateChecklistItemAsync).mockResolvedValue(undefined);

      await tools.updateChecklistItem({ checklist_item_id: 100, is_checked: true });

      expect(repo.updateChecklistItemAsync).toHaveBeenCalledWith(100, { isChecked: true });
    });

    it('passes both fields when provided', async () => {
      vi.mocked(repo.updateChecklistItemAsync).mockResolvedValue(undefined);

      await tools.updateChecklistItem({ checklist_item_id: 100, display_name: 'New name', is_checked: false });

      expect(repo.updateChecklistItemAsync).toHaveBeenCalledWith(100, { displayName: 'New name', isChecked: false });
    });
  });

  describe('prepareDeleteChecklistItem', () => {
    it('generates an approval token', () => {
      const result = tools.prepareDeleteChecklistItem({ checklist_item_id: 100 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.checklist_item_id).toBe(100);
      expect(parsed.action).toContain('confirm_delete_checklist_item');
    });
  });

  describe('confirmDeleteChecklistItem', () => {
    it('deletes the checklist item with a valid token', async () => {
      vi.mocked(repo.deleteChecklistItemAsync).mockResolvedValue(undefined);

      // Generate a token first
      const prepareResult = tools.prepareDeleteChecklistItem({ checklist_item_id: 100 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeleteChecklistItem({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Checklist item deleted');
      expect(repo.deleteChecklistItemAsync).toHaveBeenCalledWith(100);
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmDeleteChecklistItem({ approval_token: 'invalid-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.deleteChecklistItemAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeleteChecklistItem({ checklist_item_id: 100 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      // Consume the token
      await tools.confirmDeleteChecklistItem({ approval_token });

      // Try to consume again
      const result = await tools.confirmDeleteChecklistItem({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });
});
