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
import express, { type Express, type NextFunction, type Request, type Response } from 'express';
import type { AddressInfo } from 'node:net';
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
  /**
   * Returns the actually-bound port. Set by {@link startHttpServer} after listen
   * so DNS-rebinding `allowedHosts` are correct even when binding port 0
   * (ephemeral). Falls back to the configured port when absent.
   */
  readonly getBoundPort?: () => number;
}

/** True for loopback / link-local-loopback addresses. */
function isLoopbackHost(host: string): boolean {
  return host === '127.0.0.1' || host === '::1' || host === 'localhost';
}

/**
 * DNS-rebinding `allowedHosts` for the bound interface. Since there is no auth
 * yet, the loopback bind alone does NOT stop a browser page from rebinding its
 * domain to 127.0.0.1 and POSTing same-origin to `/mcp`; the SDK's Host-header
 * validation is what closes that vector. Uses the actually-bound port so
 * ephemeral (port 0) binds still validate.
 */
function allowedHostsFor(options: HttpServerOptions): string[] {
  const port = options.getBoundPort?.() ?? options.port;
  return [
    `${options.host}:${port}`,
    `127.0.0.1:${port}`,
    `localhost:${port}`,
    `[::1]:${port}`,
  ];
}

/**
 * Builds the Express app that fronts the stateless Streamable HTTP transport.
 * Exported for tests; production callers use {@link startHttpServer}, which
 * enforces the loopback-bind guard this function does not.
 */
export function buildHttpApp(options: HttpServerOptions): Express {
  const app = express();
  app.use(express.json({ limit: MAX_BODY_SIZE }));

  // Merged per-request server options are invariant for the server's lifetime —
  // build them once, not per request.
  const serverOptions: ServerOptions = {
    ...options.serverOptions,
    stateStore: options.stateStore,
  };

  // Health check — intentionally leaks no version or configuration.
  app.get('/healthz', (_req: Request, res: Response): void => {
    res.status(200).json({ status: 'ok' });
  });

  // Stateless Streamable HTTP: one MCP server + transport per POST, sharing the
  // process-scoped store. GET/DELETE carry no stateless semantics → 405.
  app.post('/mcp', async (req: Request, res: Response): Promise<void> => {
    const server: McpServer = createServer(serverOptions);
    // Stateless mode: the SDK signals it by an absent sessionIdGenerator (an
    // explicit `undefined` is rejected under exactOptionalPropertyTypes).
    // DNS-rebinding protection validates the Host header (no auth yet).
    const transport = new StreamableHTTPServerTransport({
      enableDnsRebindingProtection: true,
      allowedHosts: allowedHostsFor(options),
    });

    // Tear down the per-request server when the response closes. `server.close()`
    // also closes its transport, so one call suffices; the `.catch` prevents a
    // teardown rejection from becoming an unhandledRejection that crashes the
    // single shared process (taking every concurrent request with it).
    res.on('close', () => {
      void server.close().catch(() => {});
    });

    try {
      // Cast: the transport's getter/setter `onclose` types clash with the
      // Transport interface under exactOptionalPropertyTypes; the instance is a
      // valid Transport at runtime.
      await server.connect(transport as Transport);
      await transport.handleRequest(req, res, req.body);
    } catch (error) {
      process.stderr.write(
        `[mcp-office365] serve request error: ${error instanceof Error ? error.message : String(error)}\n`,
      );
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: '2.0',
          error: { code: -32603, message: 'Internal server error.' },
          id: null,
        });
      } else {
        // Headers already sent (mid-stream failure): end the response so the
        // client isn't left hanging until a proxy/idle timeout.
        res.end();
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

  // Body-parser failures (malformed JSON, over-limit) reach here; without this,
  // Express's default handler returns an HTML error page, which an MCP client
  // can't parse. Emit a JSON-RPC-shaped error instead.
  app.use((err: unknown, _req: Request, res: Response, next: NextFunction): void => {
    if (res.headersSent) {
      next(err);
      return;
    }
    const status = (err as { status?: number; statusCode?: number }).status
      ?? (err as { statusCode?: number }).statusCode
      ?? 400;
    res.status(status).json({
      jsonrpc: '2.0',
      error: {
        code: -32700,
        message: status === 413 ? 'Request body too large.' : 'Parse error.',
      },
      id: null,
    });
  });

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

  // Filled in after listen so allowedHostsFor() sees the real (possibly
  // ephemeral) port.
  let boundPort = options.port;
  const app = buildHttpApp({ ...options, getBoundPort: () => boundPort });

  return new Promise<HttpServer>((resolve, reject) => {
    const httpServer = app.listen(options.port, options.host, () => {
      boundPort = (httpServer.address() as AddressInfo).port;
      // Swap the startup reject-listener for a runtime one: a post-listen socket
      // error would otherwise reject an already-settled promise (a silent no-op).
      httpServer.removeListener('error', reject);
      httpServer.on('error', (err) => {
        process.stderr.write(`[mcp-office365] serve socket error: ${err.message}\n`);
      });
      resolve(httpServer);
    });
    httpServer.once('error', reject);
  });
}
