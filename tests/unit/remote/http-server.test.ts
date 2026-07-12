/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * U3: stateless Streamable HTTP transport for remote connector mode.
 *
 * These exercise the transport and routing layer only — no authentication (U4)
 * and no tool calls (which would trigger the Graph backend). `tools/list` is
 * static, so the tool surface (including the two-phase pairs, AE3) is verified
 * over HTTP without credentials. The behavioral prepare→confirm gate lands with
 * U5, where the Graph backend is injectable/mockable.
 */

import { afterEach, describe, expect, it } from 'vitest';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { AddressInfo } from 'node:net';
import type { Server as HttpServer } from 'node:http';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import { startHttpServer } from '../../../src/remote/http-server.js';
import { StateStore } from '../../../src/state/store.js';

function tempStore(): StateStore {
  // A throwaway temp-dir store is sufficient for transport-level tests.
  return StateStore.open({ dir: mkdtempSync(join(tmpdir(), 'mcp-u3-')) });
}

async function startTestServer(): Promise<{ server: HttpServer; url: string }> {
  const server = await startHttpServer({
    host: '127.0.0.1',
    port: 0, // ephemeral
    serverOptions: { confirmMode: 'token' },
    stateStore: tempStore(),
  });
  const { port } = server.address() as AddressInfo;
  return { server, url: `http://127.0.0.1:${port}/mcp` };
}

describe('remote HTTP server (U3)', () => {
  let running: HttpServer | undefined;

  afterEach(async () => {
    if (running != null) {
      await new Promise<void>((resolve) => running?.close(() => resolve()));
      running = undefined;
    }
  });

  it('serves an MCP session over stateless Streamable HTTP and lists the full tool surface', async () => {
    const { server, url } = await startTestServer();
    running = server;

    const client = new Client({ name: 'test', version: '1.0.0' }, { capabilities: {} });
    const transport = new StreamableHTTPClientTransport(new URL(url));
    await client.connect(transport);

    const result = await client.listTools();
    const names = result.tools.map((t) => t.name);

    expect(result.tools.length).toBeGreaterThan(200);
    expect(names).toContain('list_folders');

    // Covers AE3 (surface): the two-phase destructive pair survives the HTTP
    // transport. The behavioral gate (confirm required before delete) is tested
    // in U5 where the backend is mockable.
    expect(names).toContain('prepare_delete_email');
    expect(names).toContain('confirm_delete_email');

    await client.close();
  });

  it('answers /healthz without leaking version or config', async () => {
    const { server } = await startTestServer();
    running = server;
    const { port } = server.address() as AddressInfo;

    const res = await fetch(`http://127.0.0.1:${port}/healthz`);
    expect(res.status).toBe(200);
    const body = (await res.json()) as Record<string, unknown>;
    expect(body).toEqual({ status: 'ok' });
  });

  it('rejects GET and DELETE on /mcp with 405 in stateless mode', async () => {
    const { server } = await startTestServer();
    running = server;
    const { port } = server.address() as AddressInfo;

    for (const method of ['GET', 'DELETE']) {
      const res = await fetch(`http://127.0.0.1:${port}/mcp`, { method });
      expect(res.status).toBe(405);
    }
  });

  it('rejects an over-limit request body', async () => {
    const { server } = await startTestServer();
    running = server;
    const { port } = server.address() as AddressInfo;

    const huge = JSON.stringify({ blob: 'x'.repeat(5 * 1024 * 1024) });
    const res = await fetch(`http://127.0.0.1:${port}/mcp`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: huge,
    });
    expect(res.status).toBe(413);
  });

  it('refuses a non-loopback bind until the auth layer is present', () => {
    expect(() =>
      startHttpServer({
        host: '0.0.0.0',
        port: 0,
        serverOptions: {},
        stateStore: tempStore(),
      }),
    ).toThrow(/require the authentication layer/i);
  });
});
