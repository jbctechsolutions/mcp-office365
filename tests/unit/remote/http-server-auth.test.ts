/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * U4 auth wiring on the HTTP server: PRM discovery, the 401 challenge, valid
 * tokens reaching the handler, and deny-list 403. Uses a stub verifier (the JWT
 * validation itself is covered in verify.test.ts).
 */

import { afterEach, describe, expect, it } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { AddressInfo } from 'node:net';
import type { Server as HttpServer } from 'node:http';
import {
  startHttpServer,
  type RemoteAuthBundle,
} from '../../../src/remote/http-server.js';
import { loadRemoteAuthConfig } from '../../../src/remote/config.js';
import { AuthChallengeError } from '../../../src/remote/auth/verify.js';
import { StateStore } from '../../../src/state/store.js';

const config = loadRemoteAuthConfig({
  OUTLOOK_MCP_TENANT_ID: '761e2c5f-0000-4000-8000-000000000001',
  OUTLOOK_MCP_CONNECTOR_API_ID: 'api-guid',
  OUTLOOK_MCP_CONNECTOR_URL: 'https://mcp.example.com/mcp',
} as NodeJS.ProcessEnv);

const tempDirs: string[] = [];
const tempStores: StateStore[] = [];
function tempStore(): StateStore {
  const dir = mkdtempSync(join(tmpdir(), 'mcp-u4-'));
  tempDirs.push(dir);
  const store = StateStore.open({ dir });
  tempStores.push(store);
  return store;
}

// Stub verifier: token 'good' → a member; 'denied' → a deny-listed oid; else invalid.
const authBundle: RemoteAuthBundle = {
  config,
  verify: async (token: string) => {
    if (token === 'good') {
      return { oid: 'member-oid', tid: config.tenantId, homeAccountId: `member-oid.${config.tenantId}`, scopes: ['access_as_user'] };
    }
    if (token === 'denied') {
      return { oid: 'revoked-oid', tid: config.tenantId, homeAccountId: `revoked-oid.${config.tenantId}`, scopes: ['access_as_user'] };
    }
    throw new AuthChallengeError('bad_token');
  },
  denyList: { isDenied: (oid: string) => oid === 'revoked-oid' },
};

async function startAuthServer(): Promise<{ server: HttpServer; port: number }> {
  const server = await startHttpServer({
    host: '127.0.0.1',
    port: 0,
    serverOptions: { confirmMode: 'token' },
    stateStore: tempStore(),
    auth: authBundle,
  });
  return { server, port: (server.address() as AddressInfo).port };
}

function post(port: number, headers: Record<string, string>): Promise<Response> {
  return fetch(`http://127.0.0.1:${port}/mcp`, {
    method: 'POST',
    headers: { 'content-type': 'application/json', accept: 'application/json, text/event-stream', ...headers },
    body: JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 't', version: '1' } } }),
  });
}

describe('remote HTTP server auth (U4)', () => {
  let running: HttpServer | undefined;
  afterEach(async () => {
    if (running != null) {
      await new Promise<void>((resolve) => running?.close(() => resolve()));
      running = undefined;
    }
    while (tempStores.length > 0) tempStores.pop()?.close();
    while (tempDirs.length > 0) {
      const d = tempDirs.pop();
      if (d != null) { try { rmSync(d, { recursive: true, force: true }); } catch { /* noop */ } }
    }
  });

  it('serves the PRM document unauthenticated', async () => {
    const { server, port } = await startAuthServer();
    running = server;
    const res = await fetch(`http://127.0.0.1:${port}/.well-known/oauth-protected-resource`);
    expect(res.status).toBe(200);
    const body = (await res.json()) as Record<string, unknown>;
    expect(body.resource).toBe('https://mcp.example.com/mcp');
    expect(body.authorization_servers).toEqual([config.issuer]);
  });

  it('serves the PRM at the /mcp path-suffixed variant too', async () => {
    const { server, port } = await startAuthServer();
    running = server;
    const res = await fetch(
      `http://127.0.0.1:${port}/.well-known/oauth-protected-resource/mcp`,
    );
    expect(res.status).toBe(200);
    expect(((await res.json()) as Record<string, unknown>).resource).toBe(
      'https://mcp.example.com/mcp',
    );
  });

  it('challenges an unauthenticated /mcp with 401 + WWW-Authenticate, leaking no token material', async () => {
    const { server, port } = await startAuthServer();
    running = server;
    const res = await post(port, { authorization: 'Bearer supersecrettokenvalue' });
    expect(res.status).toBe(401);
    expect(res.headers.get('www-authenticate')).toMatch(/resource_metadata=/);
    const body = await res.text();
    expect(body).not.toContain('supersecrettokenvalue');
  });

  it('accepts a case-insensitive bearer scheme', async () => {
    const { server, port } = await startAuthServer();
    running = server;
    const res = await post(port, { authorization: 'bearer good' });
    expect(res.status).toBe(200);
  });

  it('rejects an invalid token with 401', async () => {
    const { server, port } = await startAuthServer();
    running = server;
    const res = await post(port, { authorization: 'Bearer nonsense' });
    expect(res.status).toBe(401);
  });

  it('rejects a deny-listed identity with 403 (no re-auth loop)', async () => {
    const { server, port } = await startAuthServer();
    running = server;
    const res = await post(port, { authorization: 'Bearer denied' });
    expect(res.status).toBe(403);
  });

  it('lets a valid member token reach the MCP handler', async () => {
    const { server, port } = await startAuthServer();
    running = server;
    const res = await post(port, { authorization: 'Bearer good' });
    // Reaches the transport and completes the MCP initialize handshake (not 401/403).
    expect(res.status).toBe(200);
    const text = await res.text();
    expect(text).toContain('"result"');
    expect(text).toContain('serverInfo');
  });
});
