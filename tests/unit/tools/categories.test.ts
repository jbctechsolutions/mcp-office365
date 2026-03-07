/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for master categories tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { CategoriesTools, type ICategoriesRepository } from '../../../src/tools/categories.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';

describe('CategoriesTools', () => {
  let repo: ICategoriesRepository;
  let tokenManager: ApprovalTokenManager;
  let tools: CategoriesTools;

  beforeEach(() => {
    repo = {
      listCategoriesAsync: vi.fn(),
      createCategoryAsync: vi.fn(),
      deleteCategoryAsync: vi.fn(),
    };
    tokenManager = new ApprovalTokenManager();
    tools = new CategoriesTools(repo, tokenManager);
  });

  describe('listCategories', () => {
    it('returns categories from the repository', async () => {
      const mockCategories = [
        { id: 1, name: 'Red Category', color: 'preset0' },
        { id: 2, name: 'Blue Category', color: 'preset1' },
      ];
      vi.mocked(repo.listCategoriesAsync).mockResolvedValue(mockCategories);

      const result = await tools.listCategories();

      expect(result.content).toHaveLength(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.categories).toEqual(mockCategories);
    });
  });

  describe('createCategory', () => {
    it('creates a category and returns the ID', async () => {
      vi.mocked(repo.createCategoryAsync).mockResolvedValue(42);

      const result = await tools.createCategory({ name: 'Work', color: 'preset1' });

      expect(repo.createCategoryAsync).toHaveBeenCalledWith('Work', 'preset1');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.category_id).toBe(42);
      expect(parsed.message).toBe('Category created');
    });
  });

  describe('prepareDeleteCategory', () => {
    it('generates an approval token', () => {
      const result = tools.prepareDeleteCategory({ category_id: 42 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.category_id).toBe(42);
      expect(parsed.action).toContain('confirm_delete_category');
    });
  });

  describe('confirmDeleteCategory', () => {
    it('deletes the category with a valid token', async () => {
      vi.mocked(repo.deleteCategoryAsync).mockResolvedValue(undefined);

      // Generate a token first
      const prepareResult = tools.prepareDeleteCategory({ category_id: 42 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeleteCategory({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Category deleted');
      expect(repo.deleteCategoryAsync).toHaveBeenCalledWith(42);
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmDeleteCategory({
        approval_token: '00000000-0000-0000-0000-000000000000',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBeDefined();
      expect(repo.deleteCategoryAsync).not.toHaveBeenCalled();
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.deleteCategoryAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeleteCategory({ category_id: 42 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      // Consume the token
      await tools.confirmDeleteCategory({ approval_token });

      // Try to use it again
      const result = await tools.confirmDeleteCategory({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });
});
