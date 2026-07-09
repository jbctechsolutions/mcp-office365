/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, beforeEach, vi } from 'vitest';
import { z } from 'zod';
import { ToolRegistry } from '../../../src/registry/registry.js';
import { defineTool } from '../../../src/registry/define-tool.js';
import { approvalTokenLink, tokenIdLink, batchLink } from '../../../src/registry/elicit-links.js';
import { createServerElicitor } from '../../../src/registry/elicitor.js';
import { allToolDefinitions } from '../../../src/registry/all-tools.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';
import type { ToolContext, ToolResult, ElicitOutcome, Elicitor } from '../../../src/registry/types.js';

const SURFACE = { backend: 'graph' as const };

function json(obj: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(obj) }] };
}

function parse(result: ToolResult | undefined): Record<string, unknown> {
  return JSON.parse(result!.content[0]!.text) as Record<string, unknown>;
}

/** Records which confirm handler ran and with what params. */
let confirmCalls: Array<{ tool: string; params: unknown }>;
let tm: ApprovalTokenManager;

/** Builds a registry with one prepare/confirm pair per confirm shape. */
function buildRegistry(): ToolRegistry {
  const reg = new ToolRegistry();
  reg.register([
    // Shape A — confirm takes { approval_token }.
    defineTool({
      name: 'prepare_a', description: '', input: z.strictObject({ override_id: z.string() }),
      annotations: {}, destructive: true, presets: [], backends: ['graph'],
      handler: (_ctx, p) => {
        const t = tm.generateToken({ operation: 'delete_focused_override', targetType: 'focused_override', targetId: p.override_id, targetHash: p.override_id });
        return json({ approval_token: t.tokenId, override_id: p.override_id });
      },
      onElicit: approvalTokenLink('confirm_a'),
    }),
    defineTool({
      name: 'confirm_a', description: '', input: z.strictObject({ approval_token: z.string() }),
      annotations: {}, destructive: true, presets: [], backends: ['graph'],
      handler: (_ctx, p) => { confirmCalls.push({ tool: 'confirm_a', params: p }); return json({ success: true, ran: 'confirm_a' }); },
    }),
    // Shape B — confirm takes { token_id, task_list_id }.
    defineTool({
      name: 'prepare_b', description: '', input: z.strictObject({ task_list_id: z.string() }),
      annotations: {}, destructive: true, presets: [], backends: ['graph'],
      handler: (_ctx, p) => {
        const t = tm.generateToken({ operation: 'delete_task_list', targetType: 'task_list', targetId: p.task_list_id, targetHash: p.task_list_id });
        return json({ token_id: t.tokenId, task_list_id: p.task_list_id });
      },
      onElicit: tokenIdLink('confirm_b', ['task_list_id']),
    }),
    defineTool({
      name: 'confirm_b', description: '', input: z.strictObject({ token_id: z.string(), task_list_id: z.string() }),
      annotations: {}, destructive: true, presets: [], backends: ['graph'],
      handler: (_ctx, p) => { confirmCalls.push({ tool: 'confirm_b', params: p }); return json({ success: true, ran: 'confirm_b' }); },
    }),
    // Shape C — batch confirm takes { tokens: [{ token_id, email_id }] }.
    defineTool({
      name: 'prepare_c', description: '', input: z.strictObject({ email_ids: z.array(z.string()) }),
      annotations: {}, destructive: true, presets: [], backends: ['graph'],
      handler: (_ctx, p) => {
        const tokens = p.email_ids.map((id) => ({
          token_id: tm.generateToken({ operation: 'batch_delete_emails', targetType: 'email', targetId: id, targetHash: id }).tokenId,
          email_id: id,
        }));
        return json({ tokens });
      },
      onElicit: batchLink('confirm_c'),
    }),
    defineTool({
      name: 'confirm_c', description: '', input: z.strictObject({ tokens: z.array(z.strictObject({ token_id: z.string(), email_id: z.union([z.string(), z.number()]) })) }),
      annotations: {}, destructive: true, presets: [], backends: ['graph'],
      handler: (_ctx, p) => { confirmCalls.push({ tool: 'confirm_c', params: p }); return json({ success: true, ran: 'confirm_c' }); },
    }),
  ]);
  return reg;
}

function ctxWith(elicit: Elicitor | undefined, confirmMode: 'token' | 'elicit' = 'elicit'): ToolContext {
  return { backend: 'graph', tokenManager: tm, graph: null, confirmMode, ...(elicit != null ? { elicit } : {}) };
}

const elicitReturning = (outcome: ElicitOutcome): Elicitor => vi.fn(() => Promise.resolve(outcome));

beforeEach(() => {
  confirmCalls = [];
  tm = new ApprovalTokenManager();
});

describe('dispatch elicitation interceptor', () => {
  it('shape A accept: runs the confirm tool with the minted token', async () => {
    const reg = buildRegistry();
    const res = await reg.dispatch('prepare_a', { override_id: 'ov1' }, ctxWith(elicitReturning('accept')), SURFACE);

    expect(confirmCalls).toEqual([{ tool: 'confirm_a', params: { approval_token: expect.any(String) } }]);
    expect(parse(res).ran).toBe('confirm_a');
    // The token handed to confirm is the one the prepare minted.
    const prepared = tm.lookupToken((confirmCalls[0]!.params as { approval_token: string }).approval_token);
    expect(prepared?.operation).toBe('delete_focused_override');
  });

  it('shape A decline: does not run confirm, invalidates the token, reports declined', async () => {
    const reg = buildRegistry();
    const res = await reg.dispatch('prepare_a', { override_id: 'ov1' }, ctxWith(elicitReturning('decline')), SURFACE);

    expect(confirmCalls).toEqual([]);
    expect(parse(res)).toMatchObject({ success: false, declined: true });
    // Exactly one token minted, and it was burned.
    expect(tm.size).toBe(0);
  });

  it('shape A degrade: returns the prepare token result unchanged, token still valid', async () => {
    const reg = buildRegistry();
    const elicit = elicitReturning('degrade');
    const res = await reg.dispatch('prepare_a', { override_id: 'ov1' }, ctxWith(elicit), SURFACE);

    expect(confirmCalls).toEqual([]);
    const body = parse(res);
    expect(body).toHaveProperty('approval_token');
    // Degrade leaves the token redeemable via the normal confirm flow.
    expect(tm.lookupToken(body.approval_token as string)).toBeDefined();
    expect(elicit).toHaveBeenCalledOnce();
  });

  it('shape B accept: builds { token_id, task_list_id } for the confirm tool', async () => {
    const reg = buildRegistry();
    await reg.dispatch('prepare_b', { task_list_id: 'tl_99' }, ctxWith(elicitReturning('accept')), SURFACE);

    expect(confirmCalls).toHaveLength(1);
    expect(confirmCalls[0]!.tool).toBe('confirm_b');
    expect(confirmCalls[0]!.params).toMatchObject({ task_list_id: 'tl_99', token_id: expect.any(String) });
  });

  it('shape C (batch) accept: passes all { token_id, email_id } pairs and consumes each', async () => {
    const reg = buildRegistry();
    await reg.dispatch('prepare_c', { email_ids: ['e1', 'e2', 'e3'] }, ctxWith(elicitReturning('accept')), SURFACE);

    expect(confirmCalls).toHaveLength(1);
    const tokens = (confirmCalls[0]!.params as { tokens: Array<{ email_id: string }> }).tokens;
    expect(tokens.map((t) => t.email_id)).toEqual(['e1', 'e2', 'e3']);
  });

  it('shape C (batch) decline: invalidates every minted token, runs no confirm', async () => {
    const reg = buildRegistry();
    await reg.dispatch('prepare_c', { email_ids: ['e1', 'e2', 'e3'] }, ctxWith(elicitReturning('decline')), SURFACE);

    expect(confirmCalls).toEqual([]);
    expect(tm.size).toBe(0); // all three burned
  });

  it('token mode bypasses elicitation entirely', async () => {
    const reg = buildRegistry();
    const elicit = elicitReturning('accept');
    const res = await reg.dispatch('prepare_a', { override_id: 'ov1' }, ctxWith(elicit, 'token'), SURFACE);

    expect(elicit).not.toHaveBeenCalled();
    expect(confirmCalls).toEqual([]);
    expect(parse(res)).toHaveProperty('approval_token');
  });

  it('no elicitor present → normal two-phase (no interception)', async () => {
    const reg = buildRegistry();
    const res = await reg.dispatch('prepare_a', { override_id: 'ov1' }, ctxWith(undefined, 'elicit'), SURFACE);
    expect(confirmCalls).toEqual([]);
    expect(parse(res)).toHaveProperty('approval_token');
  });
});

describe('onElicit wiring across the real registry', () => {
  it('every prepare tool with onElicit points at a registered confirm tool', () => {
    const reg = new ToolRegistry();
    reg.register(allToolDefinitions());
    const linked = allToolDefinitions().filter((d) => d.onElicit != null);
    // Sanity: the whole two-phase surface is wired (34 pairs).
    expect(linked.length).toBeGreaterThanOrEqual(34);
    for (const def of linked) {
      expect(reg.get(def.onElicit!.confirmTool), `${def.name} -> ${def.onElicit!.confirmTool}`).toBeDefined();
    }
  });
});

describe('createServerElicitor', () => {
  function fakeServer(caps: unknown, elicitImpl?: () => Promise<{ action: string }>) {
    return {
      getClientCapabilities: () => caps,
      elicitInput: vi.fn(elicitImpl ?? (() => Promise.resolve({ action: 'accept' }))),
    };
  }

  it('degrades without calling elicitInput when the client lacks the capability', async () => {
    const server = fakeServer(undefined);
    const elicit = createServerElicitor(server as never);
    expect(await elicit({ message: 'x' })).toBe('degrade');
    expect(server.elicitInput).not.toHaveBeenCalled();
  });

  it('maps accept/decline/cancel and passes a 60s timeout', async () => {
    const accept = createServerElicitor(fakeServer({ elicitation: {} }, () => Promise.resolve({ action: 'accept' })) as never);
    expect(await accept({ message: 'x' })).toBe('accept');

    const server = fakeServer({ elicitation: {} }, () => Promise.resolve({ action: 'accept' }));
    const el = createServerElicitor(server as never);
    await el({ message: 'x' });
    expect(server.elicitInput).toHaveBeenCalledWith(expect.anything(), { timeout: 60000 });

    const decline = createServerElicitor(fakeServer({ elicitation: {} }, () => Promise.resolve({ action: 'decline' })) as never);
    expect(await decline({ message: 'x' })).toBe('decline');

    const cancel = createServerElicitor(fakeServer({ elicitation: {} }, () => Promise.resolve({ action: 'cancel' })) as never);
    expect(await cancel({ message: 'x' })).toBe('degrade');
  });

  it('degrades (never throws) when elicitInput rejects (timeout / transport)', async () => {
    const server = fakeServer({ elicitation: {} }, () => Promise.reject(new Error('timeout')));
    const elicit = createServerElicitor(server as never);
    expect(await elicit({ message: 'x' })).toBe('degrade');
  });
});
