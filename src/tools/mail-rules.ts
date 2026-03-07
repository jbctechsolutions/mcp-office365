/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Mail rules MCP tools.
 *
 * Provides tools for managing inbox mail rules with a two-phase
 * approval pattern for destructive delete operations.
 */

import { z } from 'zod';
import type { ApprovalTokenManager } from '../approval/index.js';

// =============================================================================
// Input Schemas
// =============================================================================

export const CreateMailRuleInput = z.strictObject({
  display_name: z.string().describe('Rule name'),
  sequence: z.number().int().min(1).optional().describe('Rule priority order'),
  is_enabled: z.boolean().default(true).describe('Whether rule is active'),
  conditions: z.strictObject({
    from_addresses: z.array(z.string().email()).optional().describe('Match sender addresses'),
    subject_contains: z.array(z.string()).optional().describe('Subject contains any of these strings'),
    body_contains: z.array(z.string()).optional().describe('Body contains any of these strings'),
    sender_contains: z.array(z.string()).optional().describe('Sender field contains these strings'),
    has_attachments: z.boolean().optional().describe('Has attachments'),
    importance: z.enum(['low', 'normal', 'high']).optional().describe('Match importance level'),
  }).describe('Conditions that trigger the rule'),
  actions: z.strictObject({
    move_to_folder: z.number().int().positive().optional().describe('Folder ID to move to'),
    mark_as_read: z.boolean().optional().describe('Mark as read'),
    mark_importance: z.enum(['low', 'normal', 'high']).optional().describe('Set importance'),
    forward_to: z.array(z.string().email()).optional().describe('Forward to these addresses'),
    delete: z.boolean().optional().describe('Delete the message'),
    stop_processing_rules: z.boolean().optional().describe('Stop processing more rules'),
  }).describe('Actions to perform'),
});

export const PrepareDeleteMailRuleInput = z.strictObject({
  rule_id: z.number().int().positive().describe('The rule ID to delete'),
});

export const ConfirmDeleteMailRuleInput = z.strictObject({
  token_id: z.string().uuid().describe('Approval token from prepare_delete_mail_rule'),
  rule_id: z.number().int().positive().describe('The rule ID to delete'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type CreateMailRuleParams = z.infer<typeof CreateMailRuleInput>;
export type PrepareDeleteMailRuleParams = z.infer<typeof PrepareDeleteMailRuleInput>;
export type ConfirmDeleteMailRuleParams = z.infer<typeof ConfirmDeleteMailRuleInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface IMailRulesRepository {
  listMailRulesAsync(): Promise<Array<{ id: number; displayName: string; sequence: number; isEnabled: boolean; conditions: unknown; actions: unknown }>>;
  createMailRuleAsync(rule: Record<string, unknown>): Promise<number>;
  deleteMailRuleAsync(ruleId: number): Promise<void>;
  /** Resolve a numeric folder ID to the Graph string ID for move_to_folder action. */
  getFolderGraphId(folderId: number): string | undefined;
}

// =============================================================================
// Mail Rules Tools
// =============================================================================

/**
 * Mail rules tools with two-phase approval for delete operations.
 */
export class MailRulesTools {
  constructor(
    private readonly repo: IMailRulesRepository,
    private readonly tokenManager: ApprovalTokenManager,
  ) {}

  async listMailRules(): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const rules = await this.repo.listMailRulesAsync();
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ rules }, null, 2),
      }],
    };
  }

  async createMailRule(params: CreateMailRuleParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    // Build Graph API rule object
    const graphRule: Record<string, unknown> = {
      displayName: params.display_name,
      isEnabled: params.is_enabled,
    };
    if (params.sequence != null) graphRule['sequence'] = params.sequence;

    // Build conditions
    const conditions: Record<string, unknown> = {};
    if (params.conditions.from_addresses != null) {
      conditions['fromAddresses'] = params.conditions.from_addresses.map((addr) => ({
        emailAddress: { address: addr },
      }));
    }
    if (params.conditions.subject_contains != null) conditions['subjectContains'] = params.conditions.subject_contains;
    if (params.conditions.body_contains != null) conditions['bodyContains'] = params.conditions.body_contains;
    if (params.conditions.sender_contains != null) conditions['senderContains'] = params.conditions.sender_contains;
    if (params.conditions.has_attachments != null) conditions['hasAttachments'] = params.conditions.has_attachments;
    if (params.conditions.importance != null) conditions['importance'] = params.conditions.importance;
    graphRule['conditions'] = conditions;

    // Build actions
    const actions: Record<string, unknown> = {};
    if (params.actions.move_to_folder != null) {
      const folderId = this.repo.getFolderGraphId(params.actions.move_to_folder);
      if (folderId == null) throw new Error(`Folder ID ${params.actions.move_to_folder} not found in cache. Try listing folders first.`);
      actions['moveToFolder'] = folderId;
    }
    if (params.actions.mark_as_read != null) actions['markAsRead'] = params.actions.mark_as_read;
    if (params.actions.mark_importance != null) actions['markImportance'] = params.actions.mark_importance;
    if (params.actions.forward_to != null) {
      actions['forwardTo'] = params.actions.forward_to.map((addr) => ({
        emailAddress: { address: addr },
      }));
    }
    if (params.actions.delete != null) actions['delete'] = params.actions.delete;
    if (params.actions.stop_processing_rules != null) actions['stopProcessingRules'] = params.actions.stop_processing_rules;
    graphRule['actions'] = actions;

    const ruleId = await this.repo.createMailRuleAsync(graphRule);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, rule_id: ruleId, message: 'Mail rule created' }, null, 2),
      }],
    };
  }

  prepareDeleteMailRule(params: PrepareDeleteMailRuleParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'delete_mail_rule',
      targetType: 'rule',
      targetId: params.rule_id,
      targetHash: String(params.rule_id),
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          token_id: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          rule_id: params.rule_id,
          action: `To confirm deleting mail rule ${params.rule_id}, call confirm_delete_mail_rule with the token_id and rule_id.`,
        }, null, 2),
      }],
    };
  }

  async confirmDeleteMailRule(params: ConfirmDeleteMailRuleParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const result = this.tokenManager.consumeToken(params.token_id, 'delete_mail_rule', params.rule_id);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_mail_rule again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_mail_rule',
        TARGET_MISMATCH: 'Token was generated for a different rule',
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

    await this.repo.deleteMailRuleAsync(params.rule_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Mail rule deleted' }, null, 2),
      }],
    };
  }
}
