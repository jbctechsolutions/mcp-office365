/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Linked Resources tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { LinkedResourcesTools, type ILinkedResourcesRepository } from '../../../src/tools/linked-resources.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';

describe('LinkedResourcesTools', () => {
  let repo: ILinkedResourcesRepository;
  let tokenManager: ApprovalTokenManager;
  let tools: LinkedResourcesTools;

  beforeEach(() => {
    repo = {
      listLinkedResourcesAsync: vi.fn(),
      createLinkedResourceAsync: vi.fn(),
      deleteLinkedResourceAsync: vi.fn(),
    };
    tokenManager = new ApprovalTokenManager();
    tools = new LinkedResourcesTools(repo, tokenManager);
  });

  describe('listLinkedResources', () => {
    it('returns linked resources from the repository', async () => {
      const mockItems = [
        { id: 1, webUrl: 'https://example.com/1', applicationName: 'TestApp', displayName: 'Resource 1' },
        { id: 2, webUrl: 'https://example.com/2', applicationName: 'TestApp', displayName: 'Resource 2' },
      ];
      vi.mocked(repo.listLinkedResourcesAsync).mockResolvedValue(mockItems);

      const result = await tools.listLinkedResources({ task_id: 42 });

      expect(repo.listLinkedResourcesAsync).toHaveBeenCalledWith(42);
      expect(result.content).toHaveLength(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.linked_resources).toEqual(mockItems);
    });
  });

  describe('createLinkedResource', () => {
    it('creates a linked resource and returns the ID', async () => {
      vi.mocked(repo.createLinkedResourceAsync).mockResolvedValue(100);

      const result = await tools.createLinkedResource({ task_id: 42, web_url: 'https://example.com', application_name: 'TestApp' });

      expect(repo.createLinkedResourceAsync).toHaveBeenCalledWith(42, 'https://example.com', 'TestApp', undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.linked_resource_id).toBe(100);
      expect(parsed.message).toBe('Linked resource created');
    });

    it('passes display_name when provided', async () => {
      vi.mocked(repo.createLinkedResourceAsync).mockResolvedValue(101);

      await tools.createLinkedResource({ task_id: 42, web_url: 'https://example.com', application_name: 'TestApp', display_name: 'My Link' });

      expect(repo.createLinkedResourceAsync).toHaveBeenCalledWith(42, 'https://example.com', 'TestApp', 'My Link');
    });
  });

  describe('prepareDeleteLinkedResource', () => {
    it('generates an approval token', () => {
      const result = tools.prepareDeleteLinkedResource({ linked_resource_id: 100 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.linked_resource_id).toBe(100);
      expect(parsed.action).toContain('confirm_delete_linked_resource');
    });
  });

  describe('confirmDeleteLinkedResource', () => {
    it('deletes the linked resource with a valid token', async () => {
      vi.mocked(repo.deleteLinkedResourceAsync).mockResolvedValue(undefined);

      // Generate a token first
      const prepareResult = tools.prepareDeleteLinkedResource({ linked_resource_id: 100 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeleteLinkedResource({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Linked resource deleted');
      expect(repo.deleteLinkedResourceAsync).toHaveBeenCalledWith(100);
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmDeleteLinkedResource({ approval_token: 'invalid-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.deleteLinkedResourceAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeleteLinkedResource({ linked_resource_id: 100 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      // Consume the token
      await tools.confirmDeleteLinkedResource({ approval_token });

      // Try to consume again
      const result = await tools.confirmDeleteLinkedResource({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });
});
