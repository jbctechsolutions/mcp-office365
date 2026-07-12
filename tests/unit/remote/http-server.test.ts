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
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { AddressInfo } from 'node:net';
import http, { type Server as HttpServer } from 'node:http';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import { startHttpServer } from '../../../src/remote/http-server.js';
import { serveServerOptions } from '../../../src/index.js';
import { StateStore } from '../../../src/state/store.js';

const tempDirs: string[] = [];

function tempStore(): StateStore {
  // A throwaway temp-dir store is sufficient for transport-level tests.
  const dir = mkdtempSync(join(tmpdir(), 'mcp-u3-'));
  tempDirs.push(dir);
  return StateStore.open({ dir });
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
    while (tempDirs.length > 0) {
      const dir = tempDirs.pop();
      if (dir != null) rmSync(dir, { recursive: true, force: true });
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

  it('rejects a POST /mcp with a spoofed Host header (DNS-rebinding protection)', async () => {
    const { server } = await startTestServer();
    running = server;
    const { port } = server.address() as AddressInfo;

    // Raw http.request — undici's fetch forbids overriding the Host header, so a
    // real spoof needs the low-level client.
    const status = await new Promise<number>((resolve, reject) => {
      const req = http.request(
        {
          host: '127.0.0.1',
          port,
          path: '/mcp',
          method: 'POST',
          headers: {
            'content-type': 'application/json',
            accept: 'application/json, text/event-stream',
            host: 'evil.example.com',
          },
        },
        (res) => {
          res.resume();
          resolve(res.statusCode ?? 0);
        },
      );
      req.on('error', reject);
      req.end(JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: {} }));
    });
    // The SDK's Host validation rejects an unlisted Host before dispatching.
    expect(status).toBeGreaterThanOrEqual(400);
    expect(status).toBeLessThan(500);
  });

  it('returns a JSON-RPC error (not HTML) for a malformed JSON body', async () => {
    const { server } = await startTestServer();
    running = server;
    const { port } = server.address() as AddressInfo;

    const res = await fetch(`http://127.0.0.1:${port}/mcp`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: '{ not valid json',
    });
    expect(res.status).toBe(400);
    expect(res.headers.get('content-type')).toMatch(/application\/json/);
    const body = (await res.json()) as { jsonrpc?: string; error?: { code?: number } };
    expect(body.jsonrpc).toBe('2.0');
    expect(body.error?.code).toBe(-32700);
  });
});

describe('serveServerOptions (U3 remote overrides)', () => {
  it('forces token confirm and non-interactive auth regardless of input', () => {
    // The load-bearing guarantee: remote mode can never fall back to elicitation
    // (no channel) or interactive device-code (would hang the HTTP request).
    expect(serveServerOptions({ confirmMode: 'elicit', readOnly: true })).toMatchObject({
      confirmMode: 'token',
      interactiveAuth: false,
      readOnly: true,
    });
    expect(serveServerOptions({})).toMatchObject({
      confirmMode: 'token',
      interactiveAuth: false,
    });
  });
});
