/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for focused inbox override tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { FocusedOverridesTools, type IFocusedOverridesRepository } from '../../../src/tools/focused-overrides.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';

describe('FocusedOverridesTools', () => {
  let repo: IFocusedOverridesRepository;
  let tokenManager: ApprovalTokenManager;
  let tools: FocusedOverridesTools;

  beforeEach(() => {
    repo = {
      listFocusedOverridesAsync: vi.fn(),
      createFocusedOverrideAsync: vi.fn(),
      deleteFocusedOverrideAsync: vi.fn(),
    };
    tokenManager = new ApprovalTokenManager();
    tools = new FocusedOverridesTools(repo, tokenManager);
  });

  describe('listFocusedOverrides', () => {
    it('returns overrides from the repository', async () => {
      const mockOverrides = [
        { id: 1, senderAddress: 'a@b.com', classifyAs: 'focused' },
        { id: 2, senderAddress: 'c@d.com', classifyAs: 'other' },
      ];
      vi.mocked(repo.listFocusedOverridesAsync).mockResolvedValue(mockOverrides);

      const result = await tools.listFocusedOverrides();

      expect(result.content).toHaveLength(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.overrides).toEqual(mockOverrides);
    });
  });

  describe('createFocusedOverride', () => {
    it('creates an override and returns the ID', async () => {
      vi.mocked(repo.createFocusedOverrideAsync).mockResolvedValue(42);

      const result = await tools.createFocusedOverride({ sender_address: 'a@b.com', classify_as: 'focused' });

      expect(repo.createFocusedOverrideAsync).toHaveBeenCalledWith('a@b.com', 'focused');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.override_id).toBe(42);
      expect(parsed.message).toBe('Focused override created');
    });
  });

  describe('prepareDeleteFocusedOverride', () => {
    it('generates an approval token', () => {
      const result = tools.prepareDeleteFocusedOverride({ override_id: 42 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.override_id).toBe(42);
      expect(parsed.action).toContain('confirm_delete_focused_override');
    });
  });

  describe('confirmDeleteFocusedOverride', () => {
    it('deletes the override with a valid token', async () => {
      vi.mocked(repo.deleteFocusedOverrideAsync).mockResolvedValue(undefined);

      // Generate a token first
      const prepareResult = tools.prepareDeleteFocusedOverride({ override_id: 42 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeleteFocusedOverride({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Focused override deleted');
      expect(repo.deleteFocusedOverrideAsync).toHaveBeenCalledWith(42);
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmDeleteFocusedOverride({
        approval_token: '00000000-0000-0000-0000-000000000000',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBeDefined();
      expect(repo.deleteFocusedOverrideAsync).not.toHaveBeenCalled();
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.deleteFocusedOverrideAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeleteFocusedOverride({ override_id: 42 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      // Consume the token
      await tools.confirmDeleteFocusedOverride({ approval_token });

      // Try to use it again
      const result = await tools.confirmDeleteFocusedOverride({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });
});
