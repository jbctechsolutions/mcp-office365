/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Remote connector mode (U3): serves the MCP tool surface over stateless
 * Streamable HTTP so the server can be added as a claude.ai custom connector.
 *
 * Stateless by design: a fresh MCP Server + transport is built per request over
 * a shared, process-scoped StateStore (injected into `createServer`), so there
 * is no session map and no per-request SQLite open. Authentication is NOT part
 * of this unit — the endpoint binds to loopback only until the resource-server
 * auth layer (U4) lands; a non-localhost bind is refused here as a fail-closed
 * invariant so the endpoint can never be exposed unauthenticated.
 */

import type { Server as McpServer } from '@modelcontextprotocol/sdk/server/index.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import type { Transport } from '@modelcontextprotocol/sdk/shared/transport.js';
import express, { type Express, type Request, type Response } from 'express';
import type { Server as HttpServer } from 'node:http';
import { createServer, type ServerOptions } from '../index.js';
import type { StateStore } from '../state/store.js';

/** Maximum accepted request body size — a coarse DoS guard on the public endpoint. */
const MAX_BODY_SIZE = '4mb';

/** Options for {@link startHttpServer}. */
export interface HttpServerOptions {
  /** Interface to bind (loopback only until U4 auth lands). */
  readonly host: string;
  /** TCP port to listen on. */
  readonly port: number;
  /** Tool-surface options threaded into each per-request MCP server. */
  readonly serverOptions: ServerOptions;
  /** Shared, process-scoped durable store backing approvals and aliases. */
  readonly stateStore: StateStore;
  /**
   * Whether the resource-server auth layer (U4) is mounted. Until it is, only a
   * loopback bind is permitted. Defaults to false.
   */
  readonly authConfigured?: boolean;
}

/** True for loopback / link-local-loopback addresses. */
function isLoopbackHost(host: string): boolean {
  return host === '127.0.0.1' || host === '::1' || host === 'localhost';
}

/**
 * Builds the Express app that fronts the stateless Streamable HTTP transport.
 * Exported for tests (drive it with supertest / a real listen).
 */
export function buildHttpApp(options: HttpServerOptions): Express {
  const app = express();
  app.use(express.json({ limit: MAX_BODY_SIZE }));

  // Health check — intentionally leaks no version or configuration.
  app.get('/healthz', (_req: Request, res: Response): void => {
    res.status(200).json({ status: 'ok' });
  });

  // Stateless Streamable HTTP: one MCP server + transport per POST, sharing the
  // process-scoped store. GET/DELETE carry no stateless semantics → 405.
  app.post('/mcp', async (req: Request, res: Response): Promise<void> => {
    const server: McpServer = createServer({
      ...options.serverOptions,
      stateStore: options.stateStore,
    });
    // Stateless mode: the SDK signals it by an absent sessionIdGenerator (an
    // explicit `undefined` is rejected under exactOptionalPropertyTypes).
    const transport = new StreamableHTTPServerTransport({});

    // Tear down per-request instances when the response closes so neither the
    // transport nor the server leaks across requests.
    res.on('close', () => {
      void transport.close();
      void server.close();
    });

    try {
      // Cast: the transport's getter/setter `onclose` types clash with the
      // Transport interface under exactOptionalPropertyTypes; the instance is a
      // valid Transport at runtime.
      await server.connect(transport as Transport);
      await transport.handleRequest(req, res, req.body);
    } catch {
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: '2.0',
          error: { code: -32603, message: 'Internal server error.' },
          id: null,
        });
      }
    }
  });

  const methodNotAllowed = (_req: Request, res: Response): void => {
    res.status(405).json({
      jsonrpc: '2.0',
      error: { code: -32000, message: 'Method not allowed.' },
      id: null,
    });
  };
  app.get('/mcp', methodNotAllowed);
  app.delete('/mcp', methodNotAllowed);

  return app;
}

/**
 * Starts the remote HTTP server. Resolves once the socket is listening.
 *
 * @throws if asked to bind a non-loopback interface before U4 auth is mounted —
 *   the endpoint must never be reachable off-host without authentication.
 */
export function startHttpServer(options: HttpServerOptions): Promise<HttpServer> {
  if (!isLoopbackHost(options.host) && options.authConfigured !== true) {
    throw new Error(
      `Refusing to bind ${options.host}: remote (non-loopback) binds require the ` +
        `authentication layer, which is not yet available. Bind 127.0.0.1 for local use.`,
    );
  }

  const app = buildHttpApp(options);

  return new Promise<HttpServer>((resolve, reject) => {
    const httpServer = app.listen(options.port, options.host, () => {
      resolve(httpServer);
    });
    httpServer.once('error', reject);
  });
}
