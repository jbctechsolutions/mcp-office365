/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for mail rules tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { MailRulesTools, type IMailRulesRepository } from '../../../src/tools/mail-rules.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';

describe('MailRulesTools', () => {
  let repo: IMailRulesRepository;
  let tokenManager: ApprovalTokenManager;
  let tools: MailRulesTools;

  beforeEach(() => {
    repo = {
      listMailRulesAsync: vi.fn(),
      createMailRuleAsync: vi.fn(),
      deleteMailRuleAsync: vi.fn(),
      getFolderGraphId: vi.fn(),
    };
    tokenManager = new ApprovalTokenManager();
    tools = new MailRulesTools(repo, tokenManager);
  });

  describe('listMailRules', () => {
    it('returns rules from the repository', async () => {
      const mockRules = [
        { id: 1, displayName: 'Rule 1', sequence: 1, isEnabled: true, conditions: {}, actions: {} },
        { id: 2, displayName: 'Rule 2', sequence: 2, isEnabled: false, conditions: {}, actions: {} },
      ];
      vi.mocked(repo.listMailRulesAsync).mockResolvedValue(mockRules);

      const result = await tools.listMailRules();

      expect(result.content).toHaveLength(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.rules).toEqual(mockRules);
    });
  });

  describe('createMailRule', () => {
    it('builds correct Graph rule object and creates', async () => {
      vi.mocked(repo.createMailRuleAsync).mockResolvedValue(42);

      const result = await tools.createMailRule({
        display_name: 'Test Rule',
        is_enabled: true,
        conditions: {
          subject_contains: ['urgent'],
          from_addresses: ['alice@example.com'],
        },
        actions: {
          mark_as_read: true,
          stop_processing_rules: true,
        },
      });

      expect(repo.createMailRuleAsync).toHaveBeenCalledWith({
        displayName: 'Test Rule',
        isEnabled: true,
        conditions: {
          subjectContains: ['urgent'],
          fromAddresses: [{ emailAddress: { address: 'alice@example.com' } }],
        },
        actions: {
          markAsRead: true,
          stopProcessingRules: true,
        },
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.rule_id).toBe(42);
    });

    it('resolves move_to_folder to Graph folder ID', async () => {
      vi.mocked(repo.getFolderGraphId).mockReturnValue('graph-folder-id-abc');
      vi.mocked(repo.createMailRuleAsync).mockResolvedValue(99);

      await tools.createMailRule({
        display_name: 'Move Rule',
        is_enabled: true,
        conditions: {},
        actions: {
          move_to_folder: 123,
        },
      });

      expect(repo.getFolderGraphId).toHaveBeenCalledWith(123);
      expect(repo.createMailRuleAsync).toHaveBeenCalledWith(
        expect.objectContaining({
          actions: expect.objectContaining({
            moveToFolder: 'graph-folder-id-abc',
          }),
        })
      );
    });

    it('throws when move_to_folder ID is not in cache', async () => {
      vi.mocked(repo.getFolderGraphId).mockReturnValue(undefined);

      await expect(
        tools.createMailRule({
          display_name: 'Move Rule',
          is_enabled: true,
          conditions: {},
          actions: { move_to_folder: 999 },
        })
      ).rejects.toThrow('Folder ID 999 not found in cache');
    });

    it('handles forward_to addresses', async () => {
      vi.mocked(repo.createMailRuleAsync).mockResolvedValue(10);

      await tools.createMailRule({
        display_name: 'Forward Rule',
        is_enabled: true,
        conditions: {},
        actions: {
          forward_to: ['bob@example.com', 'carol@example.com'],
        },
      });

      expect(repo.createMailRuleAsync).toHaveBeenCalledWith(
        expect.objectContaining({
          actions: expect.objectContaining({
            forwardTo: [
              { emailAddress: { address: 'bob@example.com' } },
              { emailAddress: { address: 'carol@example.com' } },
            ],
          }),
        })
      );
    });
  });

  describe('prepareDeleteMailRule', () => {
    it('generates an approval token', () => {
      const result = tools.prepareDeleteMailRule({ rule_id: 42 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.token_id).toBeDefined();
      expect(typeof parsed.token_id).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.rule_id).toBe(42);
      expect(parsed.action).toContain('confirm_delete_mail_rule');
    });
  });

  describe('confirmDeleteMailRule', () => {
    it('deletes the rule with a valid token', async () => {
      vi.mocked(repo.deleteMailRuleAsync).mockResolvedValue(undefined);

      // Generate a token first
      const prepareResult = tools.prepareDeleteMailRule({ rule_id: 42 });
      const { token_id } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeleteMailRule({ token_id, rule_id: 42 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Mail rule deleted');
      expect(repo.deleteMailRuleAsync).toHaveBeenCalledWith(42);
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmDeleteMailRule({
        token_id: '00000000-0000-0000-0000-000000000000',
        rule_id: 42,
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBeDefined();
      expect(repo.deleteMailRuleAsync).not.toHaveBeenCalled();
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.deleteMailRuleAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeleteMailRule({ rule_id: 42 });
      const { token_id } = JSON.parse(prepareResult.content[0].text);

      // Consume the token
      await tools.confirmDeleteMailRule({ token_id, rule_id: 42 });

      // Try to use it again
      const result = await tools.confirmDeleteMailRule({ token_id, rule_id: 42 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });

    it('returns error for wrong rule ID', async () => {
      const prepareResult = tools.prepareDeleteMailRule({ rule_id: 42 });
      const { token_id } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeleteMailRule({ token_id, rule_id: 99 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(repo.deleteMailRuleAsync).not.toHaveBeenCalled();
    });
  });
});
