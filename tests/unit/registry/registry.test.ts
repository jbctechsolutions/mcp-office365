/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, vi } from 'vitest';
import { z } from 'zod';
import { ToolRegistry, toInputSchema } from '../../../src/registry/registry.js';
import { defineTool } from '../../../src/registry/define-tool.js';
import { requireGraphToolset } from '../../../src/registry/context.js';
import type { ToolContext, ToolDefinition } from '../../../src/registry/types.js';
import { mailRulesToolDefinitions } from '../../../src/tools/mail-rules.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';

function readContext(rules: unknown): ToolContext {
  return {
    backend: 'graph',
    tokenManager: new ApprovalTokenManager(),
    graph: { rules: rules as never },
    applescript: null,
  };
}

const graphSurface = { backend: 'graph' as const };

describe('requireGraphToolset', () => {
  it('returns the toolset when the graph bag is present', () => {
    const marker = { listMailRules: () => undefined };
    const ctx = readContext(marker);
    expect(requireGraphToolset(ctx, 'rules')).toBe(marker);
  });

  it('throws a clear error when the graph backend is unavailable', () => {
    const ctx: ToolContext = {
      backend: 'applescript',
      tokenManager: new ApprovalTokenManager(),
      graph: null,
      applescript: {},
    };
    expect(() => requireGraphToolset(ctx, 'rules')).toThrow(/Microsoft Graph API backend/);
  });
});

describe('toInputSchema', () => {
  it('produces an MCP-compatible object schema with no $schema key', () => {
    const schema = z.strictObject({
      rule_id: z.number().int().positive().describe('The rule ID'),
      name: z.string().optional(),
    });
    const json = toInputSchema(schema) as Record<string, unknown>;

    expect(json['$schema']).toBeUndefined();
    expect(json['type']).toBe('object');
    const props = json['properties'] as Record<string, unknown>;
    expect(props['rule_id']).toMatchObject({ type: 'integer', description: 'The rule ID' });
    expect(json['required']).toEqual(['rule_id']);
  });

  it('defaults type to object when the Zod schema produces no type', () => {
    const json = toInputSchema(z.any()) as Record<string, unknown>;
    expect(json['type']).toBe('object');
  });

  it('treats a field with a default as optional in the advertised schema (io:input)', () => {
    // A .default() field is optional for the client — the server fills it.
    const schema = z.strictObject({
      plan_id: z.number(),
      format: z.enum(['png', 'svg']).default('png'),
    });
    const json = toInputSchema(schema) as Record<string, unknown>;
    expect(json['required']).toEqual(['plan_id']);
    expect(json['required']).not.toContain('format');
  });

  it('round-trips enum and required/optional fields from the Zod schema', () => {
    // create_mail_rule is the plan's spot-check target.
    const def = mailRulesToolDefinitions().find((d) => d.name === 'create_mail_rule');
    const json = toInputSchema(def!.input) as Record<string, unknown>;
    const props = json['properties'] as Record<string, Record<string, unknown>>;
    const conditions = props['conditions'] as Record<string, unknown>;
    const condProps = conditions['properties'] as Record<string, Record<string, unknown>>;

    expect(json['required']).toEqual(expect.arrayContaining(['display_name', 'conditions', 'actions']));
    expect(condProps['importance']['enum']).toEqual(['low', 'normal', 'high']);
  });
});

describe('ToolRegistry', () => {
  it('throws on a duplicate tool name', () => {
    const registry = new ToolRegistry();
    const dup: ToolDefinition = defineTool({
      name: 'list_mail_rules',
      description: 'dup',
      input: z.strictObject({}),
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: [],
      backends: ['graph'],
      handler: () => ({ content: [{ type: 'text', text: '' }] }),
    });
    registry.register(mailRulesToolDefinitions());
    expect(() => registry.register([dup])).toThrow(/Duplicate tool registration/);
  });

  it('lists MCP-shaped tools carrying annotations for every entry', () => {
    const registry = new ToolRegistry();
    registry.register(mailRulesToolDefinitions());
    const tools = registry.listTools(graphSurface);

    expect(tools.map((t) => t.name)).toEqual([
      'list_mail_rules',
      'create_mail_rule',
      'prepare_delete_mail_rule',
      'confirm_delete_mail_rule',
    ]);
    for (const tool of tools) {
      expect(tool.inputSchema.type).toBe('object');
      expect(tool.annotations).toBeDefined();
    }
    const list = tools.find((t) => t.name === 'list_mail_rules');
    expect(list!.annotations!.readOnlyHint).toBe(true);
  });

  it('dispatches a registered tool to its handler with parsed params', async () => {
    const registry = new ToolRegistry();
    registry.register(mailRulesToolDefinitions());
    const listMailRules = vi.fn().mockResolvedValue({ content: [{ type: 'text', text: '{"rules":[]}' }] });

    const result = await registry.dispatch('list_mail_rules', {}, readContext({ listMailRules }), graphSurface);

    expect(listMailRules).toHaveBeenCalledOnce();
    expect(result).toEqual({ content: [{ type: 'text', text: '{"rules":[]}' }] });
  });

  it('dispatches a no-input tool when arguments is omitted (undefined)', async () => {
    // MCP marks `arguments` optional; a no-input tool must not throw on undefined.
    const registry = new ToolRegistry();
    registry.register(mailRulesToolDefinitions());
    const listMailRules = vi.fn().mockResolvedValue({ content: [{ type: 'text', text: '{}' }] });
    const result = await registry.dispatch('list_mail_rules', undefined, readContext({ listMailRules }), graphSurface);
    expect(listMailRules).toHaveBeenCalledOnce();
    expect(result).toBeDefined();
  });

  it('throws (does not fall through) when a tool is filtered out of the current mode', async () => {
    const registry = new ToolRegistry();
    registry.register(mailRulesToolDefinitions());
    // confirm_delete_mail_rule is destructive; read-only mode filters it out,
    // so dispatch must reject rather than return undefined into legacy fallback.
    await expect(
      registry.dispatch(
        'confirm_delete_mail_rule',
        { token_id: '00000000-0000-0000-0000-000000000000', rule_id: 1 },
        readContext({}),
        { backend: 'graph', readOnly: true },
      ),
    ).rejects.toThrow(/not available in the current mode/);
  });

  it('returns undefined for an unregistered tool so legacy dispatch can take over', async () => {
    const registry = new ToolRegistry();
    registry.register(mailRulesToolDefinitions());
    const result = await registry.dispatch('list_emails', {}, readContext({}), graphSurface);
    expect(result).toBeUndefined();
  });

  it('validates arguments against the schema before dispatch', async () => {
    const registry = new ToolRegistry();
    registry.register(mailRulesToolDefinitions());
    await expect(
      registry.dispatch('confirm_delete_mail_rule', { rule_id: 'not-a-number' }, readContext({}), graphSurface),
    ).rejects.toThrow();
  });

  it('excludes graph-only tools in AppleScript mode', () => {
    const registry = new ToolRegistry();
    registry.register(mailRulesToolDefinitions());
    const tools = registry.listTools({ backend: 'applescript' });
    expect(tools).toHaveLength(0); // all mail-rules tools are graph-only
  });

  it('read-only mode excludes destructive tools (prepare/confirm delete)', () => {
    const registry = new ToolRegistry();
    registry.register(mailRulesToolDefinitions());
    const names = registry.listTools({ backend: 'graph', readOnly: true }).map((t) => t.name);
    expect(names).toContain('list_mail_rules');
    expect(names).toContain('create_mail_rule');
    expect(names).not.toContain('prepare_delete_mail_rule');
    expect(names).not.toContain('confirm_delete_mail_rule');
  });

  it('preset filtering keeps mail tools and drops others', () => {
    const registry = new ToolRegistry();
    registry.register(mailRulesToolDefinitions());
    expect(registry.listTools({ backend: 'graph', presets: ['mail'] })).toHaveLength(4);
    expect(registry.listTools({ backend: 'graph', presets: ['calendar'] })).toHaveLength(0);
  });

  it('exposes has/get/names/isExposed accessors', () => {
    const registry = new ToolRegistry();
    registry.register(mailRulesToolDefinitions());
    expect(registry.has('list_mail_rules')).toBe(true);
    expect(registry.has('nope')).toBe(false);
    expect(registry.get('list_mail_rules')?.name).toBe('list_mail_rules');
    expect(registry.get('nope')).toBeUndefined();
    expect(registry.names()).toContain('create_mail_rule');
    expect(registry.isExposed('confirm_delete_mail_rule', graphSurface)).toBe(true);
    expect(registry.isExposed('confirm_delete_mail_rule', { backend: 'graph', readOnly: true })).toBe(false);
    expect(registry.isExposed('nope', graphSurface)).toBe(false);
  });

  it('omits the annotations key for a tool with no annotations', () => {
    const registry = new ToolRegistry();
    registry.register([
      defineTool({
        name: 'bare_tool',
        description: 'no annotations',
        input: z.strictObject({}),
        annotations: {},
        destructive: false,
        presets: [],
        backends: ['graph'],
        handler: () => ({ content: [{ type: 'text', text: '' }] }),
      }),
    ]);
    const tool = registry.listTools(graphSurface).find((t) => t.name === 'bare_tool');
    expect(tool).toBeDefined();
    expect(tool).not.toHaveProperty('annotations');
  });
});
