---
title: "MCP server crash on unloadable better-sqlite3 + stale npm-link build misreporting its version"
date: 2026-07-12
category: runtime-errors
module: state/store
problem_type: runtime_error
component: tooling
symptoms:
  - "MCP host shows the server as not connected despite valid Graph device-code auth; transport logs show it connected"
  - "Server process exits before the MCP handshake with a raw better-sqlite3 ERR_DLOPEN_FAILED stack; reconnect deterministically re-crashes"
  - "list_accounts returns an empty array regardless of auth state, giving a false not-connected signal"
  - "serverInfo.version reports the current release from a dist/ built three releases earlier"
root_cause: config_error
resolution_type: code_fix
severity: critical
related_components:
  - authentication
  - database
tags: [better-sqlite3, native-module, node-abi, npm-link, build-stamp, version-drift, fallback-degradation, dlopen]
---

# MCP server crash on unloadable better-sqlite3 + stale npm-link build misreporting its version

## Problem

A stale, npm-linked build of the MCP server (shadowing the published npx package) misreported its version and answered tool calls with a false `list_accounts: []`, making a fully-authenticated server look disconnected — while a separate latent bug crashed the server outright on any Node ABI mismatch, because its "in-memory fallback" was backed by the same native module that had just failed to load. Fixed in v4.2.1 (PR #77).

## Symptoms

- Claude Code showed the `office365` MCP server as "not connected," yet the MCP transport logs (`~/Library/Caches/claude-cli-nodejs/<project>/mcp-logs-<server>/*.jsonl`) showed the transport **successfully connected in every session**. The "not connected" indicator was not reporting transport state.
- Calling `list_accounts` through the live server returned `{"accounts": []}` — looked like no auth — but `tokens.json` held a valid MSAL cache for the correct `client_id`/`tenant`, and `get_user_profile` through the *same* server **succeeded**. Auth was fine; `list_accounts` was lying.
- The installed npx package (4.2.0) contained **no `list_accounts` tool at all**, yet the running server answered it — a contradiction that only resolved once the running binary turned out not to be the published one.
- The server self-reported version `4.2.0` while running a `dist/` built three days (three releases) earlier.
- Under a mismatched Node runtime, the published build crashed pre-handshake with a raw `better-sqlite3 ERR_DLOPEN_FAILED` stack out of `StateStore.open`; the host showed a generic "failed to connect," and every reconnect re-crashed identically.

## What Didn't Work

- **Re-authenticating (twice), via device-code flow.** Auth was never broken. `tokens.json` updated correctly each time, but the connection indicator (`list_accounts`) in the stale build didn't read Graph auth state at all — it was AppleScript-backed against local Outlook.app and returned `[]` on this machine regardless of Graph tokens. Re-authing couldn't fix a signal that wasn't reading the thing it claimed to report.
- **Trusting `serverVersion` in the MCP logs.** The stale build read `serverInfo.version` from `package.json` at runtime, so it reported the *current* version (4.2.0) while executing three-release-old compiled code. The one field that should have exposed "you're running an old build" actively concealed it.
- **Assuming the running code was the published npx package.** It wasn't. `ps` on the live processes revealed `/opt/homebrew/bin/mcp-office365` — an `npm link` pointing back at the main repo checkout, which npx resolved **ahead of** the published package. The tool-set contradiction (`list_accounts` present live, absent in 4.2.0) was the tell.
- **First classifier draft for the native-load fix.** It matched `ERR_DLOPEN_FAILED` and the `NODE_MODULE_VERSION` message but missed the **code-less** error the `bindings` package throws when the compiled `.node` artifact is missing entirely ("Could not locate the bindings file…", no `error.code`). That variant would have fallen straight through to the raw `:memory:` crash the fix exists to prevent. Caught in CodeRabbit review, closed with a regression test.

## Solution

Three fixes shipped in v4.2.1 (PR #77).

**1. Classify native-module load failures and fail fast with remediation** instead of letting the degradation fallback rethrow a raw dlopen stack (`src/state/store.ts`, the `open` catch block):

```ts
} catch (error) {
  if (fileDb !== undefined) {
    try { fileDb.close(); } catch { /* best-effort */ }
  }
  const reason = error instanceof Error ? error.message : String(error);
  // The in-memory fallback is backed by the same native module, so a module
  // that cannot load (Node ABI mismatch, missing build) would just rethrow
  // the raw dlopen stack from the fallback. Surface remediation instead.
  if (isNativeLoadFailure(error)) {
    throw new Error(
      `better-sqlite3 native module failed to load — ABI mismatch or missing compiled ` +
        `binding (running Node.js ${process.version}), so neither the on-disk state ` +
        `store nor its in-memory fallback can start.\n` +
        `Fix one of:\n` +
        `  - npm rebuild better-sqlite3   # in the directory the server is installed in\n` +
        `  - rm -rf ~/.npm/_npx           # clear the npx cache so it recompiles on next run\n` +
        `  - run the server under the Node.js version that installed it\n` +
        `Original error: ${reason}`,
      { cause: error },
    );
  }
  warn(`[mcp-office365] state store unavailable (${reason}); running in-memory (durability degraded).`);
  const mem = new Database(':memory:');
  configurePragmas(mem);
  runMigrations(mem);
  return new StateStore(mem, ':memory:', true, now);
}
```

The classifier covers all three shapes of "binding can't load," including the code-less `bindings` error the first draft missed (`src/state/store.ts`):

```ts
function isNativeLoadFailure(error: unknown): boolean {
  if (!(error instanceof Error)) return false;
  const code = (error as NodeJS.ErrnoException).code;
  if (code === 'ERR_DLOPEN_FAILED') return true;
  if (/NODE_MODULE_VERSION|was compiled against a different Node\.js version/.test(error.message)) {
    return true;
  }
  // The `bindings` package throws a code-less plain Error when the compiled
  // artifact is missing entirely (never built / pruned).
  if (/Could not locate the bindings file/.test(error.message)) return true;
  return code === 'MODULE_NOT_FOUND' && error.message.includes('better_sqlite3.node');
}
```

**2. Build-stamp the version so the server reports the build it's running, not the source tree beside it.** The build writes `dist/build-info.json` (`scripts/write-build-info.mjs`):

```js
writeFileSync(
  join(root, 'dist', 'build-info.json'),
  JSON.stringify({ version: pkg.version, builtAt: new Date().toISOString() }, null, 2) + '\n',
);
```

Startup prefers the stamp (which sits next to the compiled `index.js`), validates it as a non-empty string, and falls back to `package.json` only when no stamp exists — i.e. running from source in dev/tests (`src/index.ts`):

```ts
const requireFromHere = createRequire(import.meta.url);
function loadVersion(): string {
  try {
    const stamped = (requireFromHere('./build-info.json') as { version?: unknown }).version;
    if (typeof stamped === 'string' && stamped !== '') return stamped;
  } catch {
    /* no stamp — running from src (dev/tests) */
  }
  return (requireFromHere('../package.json') as { version: string }).version;
}
```

**3. Pass auth config through from the environment in checked-in configs.** The repo `.mcp.json` and plugin `.mcp.json` forward `OUTLOOK_MCP_CLIENT_ID` / `OUTLOOK_MCP_TENANT_ID` from the environment (required since #75), so the checked-in server entries work under 4.2.0+ instead of failing auth with no client ID.

**Regression test** (`tests/unit/state/native-load-failure.test.ts`) mocks `better-sqlite3` to throw at construction and asserts `open` throws one actionable error (never the raw stack), preserves the original as `cause`, names `process.version`, and — critically — does **not** emit the "running in-memory" degrade warning, since the fallback is impossible here. A dedicated case drives the code-less "Could not locate the bindings file" variant through the same path.

## Why This Works

Root cause was two independent failures, both hidden behind a misleading signal:

- The degradation fallback depended on the very component whose failure it was meant to handle. `StateStore.open` degrades to an in-memory SQLite db when the on-disk file is unusable — but in-memory SQLite is still `better-sqlite3`, the same native binding. When the *binding* is what failed (ABI mismatch, missing `.node`), the fallback re-invokes it and rethrows the same dlopen error, so a fallback meant to keep the server alive instead guaranteed a pre-handshake crash. Splitting "bad file" (degrade) from "bad binding" (fail fast + remediation) restores the fallback to only the cases it can handle.
- A version string read at runtime from `package.json` describes the *source tree*, not the *running build*. With `npm link` (and npx resolving the link ahead of the published package), the executing `dist/` and the adjacent `package.json` drift apart, and a `package.json`-sourced version silently reports the newer number. Stamping the version at build time into an artifact that ships *inside* `dist/` binds the reported version to the code that was actually compiled.

Two generalizable principles:

1. **A degradation fallback must not depend on the component whose failure it handles.** If the fallback shares the failing dependency, it isn't a fallback — it's a second attempt at the same crash. Classify the failure and fail fast with remediation when no genuine fallback exists.
2. **A version string read at runtime from `package.json` describes the source tree, not the running build.** Any signal that identifies "what is running" must be bound to the build, not to a sibling file that can drift from it.

## Prevention

- **Classify native-module load failures explicitly, and include the code-less variant.** Match `ERR_DLOPEN_FAILED`, the `NODE_MODULE_VERSION` / "compiled against a different Node.js version" message, a `MODULE_NOT_FOUND` naming the `.node` file, **and** the `bindings` package's code-less "Could not locate the bindings file" plain `Error`. The last one has no `error.code`, so a code-only classifier silently misses the "never built / pruned" case.
- **A degradation fallback that shares the failed dependency is not a fallback.** Before writing a catch-block fallback, ask what it depends on; if it re-uses the thing that just failed, fail fast with an actionable message naming the running runtime and the fix, rather than pretending to degrade.
- **Build-stamp the running version.** Emit a build-time artifact (`{version, builtAt}`) inside the compiled output and prefer it at startup; fall back to `package.json` only when no stamp exists (dev/source). Validate the stamp is a non-empty string before trusting it.
- **A connection/health indicator must read the state it claims to report.** `list_accounts` returned `[]` from a local-Outlook backend while reporting on a Graph-authenticated server — indicator and state were decoupled. Health/status signals must query the actual subsystem they represent, or they lie under exactly the conditions you most need them.
- **Watch for `npm link` shadowing npx.** When a server "answers tools it shouldn't have" or reports a version that contradicts its behavior, run `ps` to find the actual binary path; a Homebrew/`npm link` symlink can resolve ahead of the published package. The tool-set is a faster tell than the version string.
- **Regression-test the fallback path with the module mocked broken.** Mock the native dependency to throw at construction and assert the surfaced error is actionable (names remediation + runtime, preserves `cause`) and that the false-degrade warning is *not* emitted. The fallback path only runs under failure, so it's exactly the path that rots untested.

## Related Issues

- [#77](https://github.com/jbctechsolutions/mcp-office365/pull/77) — the fix PR (v4.2.1)
- [#76](https://github.com/jbctechsolutions/mcp-office365/issues/76) — follow-up: true graceful no-sqlite degrade (JS Map-backed store), the step beyond fail-fast
- [#75](https://github.com/jbctechsolutions/mcp-office365/pull/75) — the fail-fast-with-remediation precedent this fix follows
- [#70](https://github.com/jbctechsolutions/mcp-office365/pull/70) — the **install-time** cousin of this bug: `better-sqlite3@12.6.2` capped `engines.node` at 25.x, so `npx` install aborted with `EBADENGINE` on Node 26 (fixed by bumping to `^12.11.1`). Together with this doc: better-sqlite3 can break at install time (engine cap) or at load time (ABI mismatch / missing binding) — both present as "the server never starts."
- [device-code-auth-undefined-invalid-grant](../integration-issues/device-code-auth-undefined-invalid-grant.md) — same fail-fast-with-remediation discipline, for the auth path
- [test-external-api-assumptions-before-building-defenses](../best-practices/test-external-api-assumptions-before-building-defenses.md) — verify the real failure before defending against it
