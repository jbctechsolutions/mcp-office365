---
module: remote-transport
date: 2026-07-12
problem_type: architecture_pattern
component: tooling
severity: high
applies_when:
  - "Adding an HTTP/Streamable-HTTP transport to a previously stdio-only MCP server"
  - "Running per-request MCP Server instances that must share durable state"
  - "Standing up an unauthenticated MCP endpoint on loopback ahead of an auth layer"
tags:
  - mcp
  - streamable-http
  - stateless-transport
  - statestore
  - dns-rebinding
  - express
  - better-sqlite3
  - windows
  - graceful-shutdown
related_components:
  - src/remote/http-server.ts
  - src/index.ts
  - src/state/store.ts
---

# Stateless Streamable HTTP transport for a stdio MCP server

## Context

`mcp-office365` shipped as a stdio-only MCP server. Remote connector mode (U3)
adds a `serve` subcommand exposing the same tool surface over **stateless
Streamable HTTP** (the transport a claude.ai custom connector fronts). This is
the repo's first HTTP transport, first per-request server instances, and first
internet-adjacent endpoint — none of which the existing stdio patterns covered.
A 10-persona code review surfaced a cluster of gotchas that are non-obvious the
first time and will bite the next remote-mode change if not written down.

## Guidance

### 1. Inject one process-scoped `StateStore`; don't open one per request

`createServer()` opened its own SQLite `StateStore` internally. In stateless
HTTP mode a fresh `createServer()` runs **per request**, so an un-refactored
factory re-opens SQLite and re-runs migrations on every POST and leaks a file
handle per request. Add an injection seam and open the store **once** at process
scope, passing the same instance into every per-request server:

```ts
// ServerOptions gains an optional injected store; stdio path still opens its own.
const stateStore = options.stateStore ?? StateStore.open();

// http-server.ts — one store for the process, threaded into each request's server:
const server = createServer({ ...serverOptions, stateStore });
```

This is also *required* for correctness, not just efficiency: durable-ID alias
tokens and two-phase approval tokens resolve **only where the store lives** (see
[[alias-backed-composite-durable-id-pattern]]). A `prepare_*` token minted on one
request must be redeemable on a later request — which only holds if every
per-request server shares the same store instance, on a single replica with a
persistent volume. Multi-replica or per-request stores silently break token
resolution.

### 2. An unauthenticated loopback bind is NOT safe — enable DNS-rebinding protection

Binding `127.0.0.1` stops off-host TCP but does **not** stop a web page the user
visits from rebinding its domain to `127.0.0.1` and POSTing same-origin to
`/mcp` — driving the signed-in mailbox while there is no auth. The SDK's
Host-header validation is what closes this:

```ts
new StreamableHTTPServerTransport({
  enableDnsRebindingProtection: true,
  allowedHosts: [`127.0.0.1:${port}`, `localhost:${port}`, `[::1]:${port}`],
});
```

Use the **actually-bound** port (read `server.address().port` after `listen`),
not the configured one, or ephemeral (port 0) binds fail their own Host check.

### 3. Guard per-request teardown promises or one rejection crashes everything

Tearing down the per-request server on `res` close with a bare `void
server.close()` floats the promise. Under Node's default unhandled-rejection
behavior a single rejected teardown crashes the **whole shared process**, taking
every concurrent request with it. Always `.catch()`:

```ts
res.on('close', () => { void server.close().catch(() => {}); });
```

(`server.close()` also closes its transport, so one guarded call suffices — a
separate `transport.close()` is a redundant double-close.)

### 4. Fail fast instead of triggering interactive auth over HTTP

An unauthenticated tool call in stdio mode triggers the interactive device-code
flow. Over HTTP there is no device-code channel, so the request **hangs** (~15
min until MSAL times out) and, with per-request servers, concurrent calls spawn
multiple flows. Thread a non-interactive flag so serve mode throws a typed
"authenticate first" error instead:

```ts
if (!authenticated) {
  if (options.interactiveAuth === false) throw new GraphAuthRequiredError('not_authenticated');
  await getAccessToken(); // stdio path unchanged (defaults to interactive)
}
```

### 5. Return protocol-shaped errors and drain on shutdown

- Add Express error middleware so malformed/over-limit bodies return a **JSON-RPC**
  error, not Express's HTML page (an MCP client can't parse HTML).
- Log caught request errors to stderr (don't swallow them) and `res.end()` on a
  mid-stream (`headersSent`) failure so clients don't hang.
- Capture the returned `HttpServer` and close it on `SIGTERM`/`SIGINT` — Container
  Apps sends SIGTERM, and an uncaptured server never drains.

### 6. Close SQLite handles before removing temp dirs in tests (Windows)

Test cleanup that `rmSync`'s a temp dir while the `StateStore`'s `state.db`
handle is still open passes on macOS/Linux (POSIX allows unlinking open files)
and **fails on Windows with `EBUSY`**. Close the store first:

```ts
afterEach(() => {
  tempStores.forEach((s) => s.close()); // release the SQLite handle first
  tempDirs.forEach((d) => { try { rmSync(d, { recursive: true, force: true }); } catch {} });
});
```

This only surfaced in CI's Windows matrix — a reminder that the cross-OS matrix
earns its keep for anything touching native file handles.

## Why This Matters

Each of these is silent or platform-specific: the per-request store leak degrades
slowly, the DNS-rebinding hole is invisible until exploited, the teardown crash
needs a rejection to fire, the auth hang only hits the first unauthenticated
user, and the Windows `EBUSY` never appears on a Mac dev machine. They cluster
around the same shift — **one shared process now serves many requests instead of
one stdio client** — so per-client isolation assumptions from stdio no longer
hold, and blast radius is process-wide.

## When to Apply

- Any future change to `src/remote/` or a second transport.
- Before assuming loopback == safe for an unauthenticated dev endpoint.
- When multi-replica hosting is ever considered — revisit guidance #1 first
  (the store/token-resolution constraint is the blocker, not the transport).

## Examples

The full implementation is `src/remote/http-server.ts` + the `serve` branch in
`src/index.ts` (mcp-office365 PR #78). The gotchas above were each caught by
`ce-code-review` before merge (findings #1–#7) or by CI's Windows matrix
(#6) — not by the initial happy-path implementation, which passed local tests on
macOS.
