/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * U8 audit log — the write/destructive trail (R16). Exercises the recorder's
 * classification, prepare↔confirm linkage, the fail-closed confirm path, and the
 * fail-open non-two-phase write path, plus the store's query + upgrade-boundary
 * readability.
 */

import { afterEach, describe, expect, it, vi } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { StateStore, type AuditStore } from '../../../src/state/store.js';
import { createAuditRecorder, type AuditToolInfo } from '../../../src/remote/audit.js';
import { AuditUnavailableError } from '../../../src/utils/errors.js';
import type { ToolResult } from '../../../src/registry/types.js';

const IDENTITY = { oid: 'oid-1', tid: 'tid-1' };

const dirs: string[] = [];
const stores: StateStore[] = [];
function store(): StateStore {
  const dir = mkdtempSync(join(tmpdir(), 'mcp-u8-'));
  dirs.push(dir);
  const s = StateStore.open({ dir });
  stores.push(s);
  return s;
}
afterEach(() => {
  while (stores.length > 0) stores.pop()?.close();
  while (dirs.length > 0) {
    const d = dirs.pop();
    if (d != null) rmSync(d, { recursive: true, force: true });
  }
});

const readTool: AuditToolInfo = { name: 'list_emails', readOnly: true };
const writeTool: AuditToolInfo = { name: 'mark_email_read', readOnly: false };
const confirmSend: AuditToolInfo = { name: 'confirm_send_email', readOnly: false };
const confirmDelete: AuditToolInfo = { name: 'confirm_delete_email', readOnly: false };
const prepareSend: AuditToolInfo = {
  name: 'prepare_send_email',
  readOnly: false,
  collectTokenIds: (r: ToolResult): string[] => {
    const parsed = JSON.parse(r.content[0]?.text ?? '{}') as { approval_token?: string };
    return parsed.approval_token != null ? [parsed.approval_token] : [];
  },
};

function tokenResult(token: string): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify({ approval_token: token }) }] };
}

describe('audit recorder — classification (U8)', () => {
  it('does not audit a read-only tool', () => {
    const s = store();
    const rec = createAuditRecorder(s, IDENTITY);
    expect(rec.begin(readTool, { folder: 'inbox' })).toBeNull();
    expect(s.listAudit()).toHaveLength(0);
  });

  it('records a non-two-phase write once, with identity + target + outcome', () => {
    const s = store();
    const rec = createAuditRecorder(s, IDENTITY);
    const pending = rec.begin(writeTool, { email_id: 'em_42' });
    expect(pending).not.toBeNull();
    // Fail-open writes are recorded AFTER the call — nothing yet.
    expect(s.listAudit()).toHaveLength(0);
    pending!.finish({ ok: true });

    const rows = s.listAudit();
    expect(rows).toHaveLength(1);
    expect(rows[0]).toMatchObject({
      oid: 'oid-1',
      tid: 'tid-1',
      tool: 'mark_email_read',
      phase: 'write',
      outcome: 'ok',
    });
    expect(rows[0]?.target).toContain('em_42');
  });
});

describe('audit recorder — prepare/confirm linkage (U8)', () => {
  it('links a prepare row to its confirm row by the approval-token id', () => {
    const s = store();
    const rec = createAuditRecorder(s, IDENTITY);

    // prepare_send_email mints a token in its result.
    const prep = rec.begin(prepareSend, { to: 'a@tester.jbc.dev' });
    prep!.finish({ ok: true, result: tokenResult('ap_link_9') });

    // confirm_send_email carries that token in its args.
    const conf = rec.begin(confirmSend, { approval_token: 'ap_link_9' });
    conf!.finish({ ok: true });

    const rows = s.listAudit({ oid: 'oid-1' });
    expect(rows).toHaveLength(2);
    const prepareRow = rows.find((r) => r.phase === 'prepare');
    const confirmRow = rows.find((r) => r.phase === 'confirm');
    expect(prepareRow?.linkKey).toBe('ap_link_9');
    expect(confirmRow?.linkKey).toBe('ap_link_9');
    expect(confirmRow?.outcome).toBe('ok');
  });

  it('reserves the confirm row BEFORE the mutation, then finalizes it', () => {
    const s = store();
    const rec = createAuditRecorder(s, IDENTITY);
    const conf = rec.begin(confirmDelete, { email_id: 'em_7', approval_token: 'ap_1' });
    // Reserved immediately (pending), before finish().
    let rows = s.listAudit();
    expect(rows).toHaveLength(1);
    expect(rows[0]?.outcome).toBe('pending');
    conf!.finish({ ok: false, errorCode: 'GRAPH_ERROR' });
    rows = s.listAudit();
    expect(rows).toHaveLength(1); // same row, updated in place
    expect(rows[0]?.outcome).toBe('error');
    expect(rows[0]?.errorCode).toBe('GRAPH_ERROR');
  });
});

describe('audit recorder — fail-closed vs fail-open (U8)', () => {
  /** A store whose audit writes always throw (audit table unavailable). */
  const brokenStore: AuditStore = {
    recordAudit: () => {
      throw new Error('audit table unavailable');
    },
    updateAuditOutcome: () => {
      throw new Error('audit table unavailable');
    },
  };

  it('aborts confirm_send_email when the audit write fails (fail-closed)', () => {
    const rec = createAuditRecorder(brokenStore, IDENTITY, { warn: () => {} });
    expect(() => rec.begin(confirmSend, { approval_token: 'ap_1' })).toThrow(AuditUnavailableError);
  });

  it('aborts confirm_delete_email when the audit write fails (fail-closed)', () => {
    const rec = createAuditRecorder(brokenStore, IDENTITY, { warn: () => {} });
    expect(() => rec.begin(confirmDelete, { email_id: 'em_7', approval_token: 'ap_1' })).toThrow(
      AuditUnavailableError,
    );
  });

  it('lets a non-two-phase write proceed with a warning when the audit write fails (fail-open)', () => {
    const warn = vi.fn();
    const rec = createAuditRecorder(brokenStore, IDENTITY, { warn });
    // begin() must not throw for a non-confirm write.
    const pending = rec.begin(writeTool, { email_id: 'em_9' });
    expect(pending).not.toBeNull();
    // The record write happens (and fails) at finish → warns, does not throw.
    expect(() => pending!.finish({ ok: true })).not.toThrow();
    expect(warn).toHaveBeenCalledOnce();
  });
});

describe('audit recorder — batch shapes (U8)', () => {
  const confirmBatch: AuditToolInfo = { name: 'confirm_batch_operation', readOnly: false };
  const prepareBatchDelete: AuditToolInfo = { name: 'prepare_batch_delete_emails', readOnly: false };

  it('records every token id and email id in a batch confirm (not just the last)', () => {
    const s = store();
    const rec = createAuditRecorder(s, IDENTITY);
    // Real confirm_batch_operation shape: tokens: [{ token_id, email_id }, …].
    rec.begin(confirmBatch, {
      tokens: [
        { token_id: 't1', email_id: 'em_1' },
        { token_id: 't2', email_id: 'em_2' },
        { token_id: 't3', email_id: 'em_3' },
      ],
    })!.finish({ ok: true });

    const row = s.listAudit()[0];
    expect(row?.phase).toBe('confirm');
    // All three token ids linked (deduped, comma-joined).
    expect(row?.linkKey).toBe('t1,t2,t3');
    // All three email ids present in the target — no overwrite collapse.
    const target = row?.target ?? '';
    for (const id of ['em_1', 'em_2', 'em_3']) expect(target).toContain(id);
  });

  it('records every email id in a batch prepare (email_ids: string[])', () => {
    const s = store();
    const rec = createAuditRecorder(s, IDENTITY);
    // Real prepare_batch_delete_emails shape: email_ids: [str, …].
    rec.begin(prepareBatchDelete, { email_ids: ['em_a', 'em_b', 'em_c'] })!.finish({ ok: true });
    const target = s.listAudit()[0]?.target ?? '';
    for (const id of ['em_a', 'em_b', 'em_c']) expect(target).toContain(id);
  });

  it('never stores batch token ids in the target (they are the link key)', () => {
    const s = store();
    const rec = createAuditRecorder(s, IDENTITY);
    rec.begin(confirmBatch, { tokens: [{ token_id: 'tok_secret', email_id: 'em_1' }] })!.finish({
      ok: true,
    });
    expect(s.listAudit()[0]?.target ?? '').not.toContain('tok_secret');
  });
});

describe('audit recorder — no content material stored (U8)', () => {
  it('records only id-shaped params, never subject/body content', () => {
    const s = store();
    const rec = createAuditRecorder(s, IDENTITY);
    rec
      .begin(writeTool, {
        email_id: 'em_1',
        subject: 'Secret merger terms',
        body: 'Do not disclose',
        comment: 'private',
      })!
      .finish({ ok: true });
    const target = s.listAudit()[0]?.target ?? '';
    expect(target).toContain('em_1');
    expect(target).not.toContain('Secret');
    expect(target).not.toContain('disclose');
    expect(target).not.toContain('private');
  });
});

describe('audit store — query + upgrade-boundary readability (U8)', () => {
  it('filters by oid and returns newest first', () => {
    const s = store();
    s.recordAudit({ oid: 'a', tool: 't1', phase: 'write', outcome: 'ok', createdAt: 100 });
    s.recordAudit({ oid: 'b', tool: 't2', phase: 'write', outcome: 'ok', createdAt: 200 });
    s.recordAudit({ oid: 'a', tool: 't3', phase: 'write', outcome: 'ok', createdAt: 300 });
    const rowsA = s.listAudit({ oid: 'a' });
    expect(rowsA.map((r) => r.tool)).toEqual(['t3', 't1']);
    expect(s.listAudit({ since: 250 })).toHaveLength(1);
    expect(s.listAudit({ limit: 1 })[0]?.tool).toBe('t3');
  });

  it('audit rows remain readable after reopening the store (additive upgrade)', () => {
    const dir = mkdtempSync(join(tmpdir(), 'mcp-u8-reopen-'));
    dirs.push(dir);
    const first = StateStore.open({ dir });
    first.recordAudit({ oid: 'oid-1', tool: 'confirm_send_email', phase: 'confirm', outcome: 'ok' });
    first.close();

    const second = StateStore.open({ dir });
    stores.push(second);
    const rows = second.listAudit({ oid: 'oid-1' });
    expect(rows).toHaveLength(1);
    expect(rows[0]).toMatchObject({ tool: 'confirm_send_email', phase: 'confirm', outcome: 'ok' });
  });
});
