/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * U8 audit wiring at the CallTool chokepoint. Drives a real per-request server
 * (createServer) over an in-memory transport in remote mode. The Graph backend
 * is unprovisioned (remoteMode + no OBO), so tool calls fail closed at init —
 * but the audit chokepoint records the attempt regardless, which is exactly the
 * R16 behavior under test. The fail-closed confirm case asserts the mutation is
 * short-circuited before the backend is even touched.
 */

import { afterEach, describe, expect, it } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createServer, type ServerOptions } from '../../../src/index.js';
import { StateStore, type AuditStore } from '../../../src/state/store.js';

const IDENTITY = { oid: 'oid-abc', tid: 'tid-xyz' };

const dirs: string[] = [];
const stores: StateStore[] = [];
const clients: Client[] = [];

function store(): StateStore {
  const dir = mkdtempSync(join(tmpdir(), 'mcp-u8c-'));
  dirs.push(dir);
  const s = StateStore.open({ dir });
  stores.push(s);
  return s;
}

async function connect(options: ServerOptions): Promise<Client> {
  const server = createServer(options);
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);
  const client = new Client({ name: 'test', version: '1.0.0' }, { capabilities: {} });
  await client.connect(clientTransport);
  clients.push(client);
  return client;
}

/** Parses the error envelope carried in a tool result's text content. */
function envelopeOf(result: unknown): { code?: string; retriable?: boolean } {
  const content = (result as { content?: Array<{ type: string; text?: string }> }).content ?? [];
  const text = content.find((b) => b.type === 'text')?.text ?? '{}';
  return JSON.parse(text) as { code?: string; retriable?: boolean };
}

afterEach(async () => {
  while (clients.length > 0) await clients.pop()?.close();
  while (stores.length > 0) stores.pop()?.close();
  while (dirs.length > 0) {
    const d = dirs.pop();
    if (d != null) rmSync(d, { recursive: true, force: true });
  }
});

describe('audit chokepoint wiring (U8)', () => {
  it('records a write tool call under the request identity', async () => {
    const s = store();
    const client = await connect({
      stateStore: s,
      remoteMode: true,
      audit: { store: s, oid: IDENTITY.oid, tid: IDENTITY.tid },
    });
    // Fails closed at backend init (no OBO), but the attempt is still audited.
    await client.callTool({ name: 'mark_email_read', arguments: { email_id: 'em_1' } });

    const rows = s.listAudit({ oid: 'oid-abc' });
    expect(rows).toHaveLength(1);
    expect(rows[0]).toMatchObject({ tool: 'mark_email_read', phase: 'write', tid: 'tid-xyz' });
    expect(rows[0]?.target).toContain('em_1');
    expect(rows[0]?.outcome).toBe('error'); // init failed, but recorded
  });

  it('does not record a read-only tool call', async () => {
    const s = store();
    const client = await connect({
      stateStore: s,
      remoteMode: true,
      audit: { store: s, oid: IDENTITY.oid, tid: IDENTITY.tid },
    });
    await client.callTool({ name: 'list_folders', arguments: {} });
    expect(s.listAudit()).toHaveLength(0);
  });

  it('fails closed on confirm_* when the audit store cannot record (AUDIT_UNAVAILABLE)', async () => {
    const s = store();
    const brokenAudit: AuditStore = {
      recordAudit: () => {
        throw new Error('audit table unavailable');
      },
      updateAuditOutcome: () => {},
    };
    const client = await connect({
      stateStore: s,
      remoteMode: true,
      audit: { store: brokenAudit, oid: IDENTITY.oid, tid: IDENTITY.tid },
    });
    const result = await client.callTool({
      name: 'confirm_delete_email',
      arguments: { email_id: 'em_9', approval_token: 'ap_1' },
    });
    const env = envelopeOf(result);
    // Short-circuited before the backend: the code is the audit abort, NOT a
    // Graph/init error — proving the mutation never ran.
    expect(env.code).toBe('AUDIT_UNAVAILABLE');
    expect(env.retriable).toBe(true);
  });

  it('reserves then finalizes a confirm_* row through the chokepoint', async () => {
    const s = store();
    const client = await connect({
      stateStore: s,
      remoteMode: true,
      audit: { store: s, oid: IDENTITY.oid, tid: IDENTITY.tid },
    });
    await client.callTool({
      name: 'confirm_delete_email',
      arguments: { email_id: 'em_5', approval_token: 'ap_2' },
    });
    const rows = s.listAudit({ oid: 'oid-abc' });
    expect(rows).toHaveLength(1);
    expect(rows[0]).toMatchObject({ tool: 'confirm_delete_email', phase: 'confirm', outcome: 'error' });
    expect(rows[0]?.linkKey).toBe('ap_2');
  });
});
