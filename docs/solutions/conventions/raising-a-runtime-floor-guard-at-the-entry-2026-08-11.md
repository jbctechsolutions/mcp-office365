---
title: Raising a runtime floor needs an entry guard and a test below the floor
date: 2026-08-11
category: conventions
module: packaging/entry-point
problem_type: convention
component: tooling
severity: high
applies_when:
  - Raising `engines.node` (or any runtime floor) in a released package
  - Adopting a stdlib or platform API that only exists above a certain runtime version
  - Dropping support for a runtime major in a breaking release
symptoms:
  - "A supported-looking install dies at launch with a raw `ERR_UNKNOWN_BUILTIN_MODULE` stack"
  - "An MCP host reports only \"failed to connect\" with no usable cause"
  - "`npm install` succeeds on a runtime the package does not support"
related_components:
  - development_workflow
tags:
  - engines
  - node-floor
  - breaking-change
  - startup-failure
  - esm-module-graph
  - release-checklist
---

# Raising a runtime floor needs an entry guard and a test below the floor

## Context

v5.0.0 replaced `better-sqlite3` with the stdlib `node:sqlite` (#108). The entire point was to end a class of illegible startup failures: a compiled binding that could not load under the running Node killed the server before the MCP handshake, and the host showed nothing but "failed to connect" (#77, #88, #95).

The new driver required Node 24, so `engines.node` moved from `>=20.0.0` to `>=24.0.0`.

That re-created the exact failure class the release existed to eliminate — in a new form, and nobody noticed for four merged changes. It surfaced only while re-reading an old issue (#76) to decide whether to close it, not from testing the boundary that had just been created.

## Guidance

**1. `engines` is documentation, not enforcement.** npm only *warns* on an unsatisfied `engines.node`. It refuses the install only when the consumer sets `engine-strict=true`, which almost nobody does. So declaring a floor does not stop anyone below it from installing — it just changes where they find out.

**2. Guard the floor at runtime, and say what to do.** Check the version at startup and fail with remedies, not a stack trace:

```text
mcp-office365 requires Node.js 24 or newer — this is Node.js 20.20.2.

Fix one of:
  - Upgrade to Node.js 24+ and relaunch
  - Point your MCP client at a Node.js 24+ binary (an absolute path
    avoids version-manager shims the client cannot see)
  - Pin the previous major, which supports Node.js 20+:
      npm install @jbctechsolutions/mcp-office365@4
```

Naming the fallback version matters most. It is the one remedy that works *today* for someone who cannot upgrade.

**3. In ESM, the guard needs its own entry module.** This is the non-obvious part. ESM resolves and links an entire module graph *before executing any of it*, so a version check placed anywhere inside the application never runs — Node rejects the graph at the unavailable import first. The check must live in an entry file that statically imports **nothing** reaching the new API, then pulls the real entry in dynamically:

```ts
// index.ts — the executable entry
import { nodeFloorError } from './node-floor.js';   // imports nothing itself

const floorError = nodeFloorError(process.versions.node);
if (floorError !== null) {
  process.stderr.write(floorError);
  process.exitCode = 1;          // NOT process.exit() — see below
} else {
  const { main } = await import('./server.js');     // dynamic on purpose
  main().catch((error: unknown) => {
    console.error('Fatal error:', error);
    process.exitCode = 1;                           // same reason as above
  });
}
```

**4. Do not `process.exit()` right after writing the message.** Node writes to `stderr` asynchronously when it is a pipe or socket, and synchronously to a file or TTY. A host that captures output gives you the asynchronous case, which is exactly when the guidance matters most. Calling `process.exit()` can terminate before the buffer drains and truncate the very guidance the guard exists to deliver. Set `process.exitCode` and let the process end naturally.

**5. Test *below* the floor.** A matrix that only runs supported versions proves nothing about what an unsupported user sees. Assert on the message, not just the exit code, and pin the constant to the manifest so the two cannot drift:

```ts
it.each([`${MIN_NODE_MAJOR - 4}.20.2`, `${MIN_NODE_MAJOR - 2}.11.0`, `${MIN_NODE_MAJOR - 1}.0.0`])(
  'refuses to start on Node %s',
  (v) => {
    expect(nodeFloorError(v)).toContain(`requires Node.js ${MIN_NODE_MAJOR}`);
  },
);

it('matches the floor declared in engines.node', () => {
  expect(MIN_NODE_MAJOR).toBe(Number(/(\d+)/.exec(pkg.engines.node)?.[1]));
});
```

**6. Check what the entry split did to your package entries.** Splitting an entry point is easy to do halfway. `main`, `types`, and `exports` pointed at the file that had just become the executable shim, so importing the package as a library would have started a server as a side effect and exported nothing. Library entries go to the module; `bin` goes to the guard.

**7. Verify the boundary claim, do not infer it from release notes.** The floor was chosen as 24 rather than 22.5 because measurement disagreed with the documentation:

| Node | `DatabaseSync` + `prepare`/`exec`/`run`, unflagged |
|---|---|
| 22.11.0 | fails — `No such built-in module` |
| 22.13.0 / 22.15.0 | works, emits `ExperimentalWarning` on every launch |
| **24.11.0** | **works, still emits `ExperimentalWarning`** |
| 24.18.0 | works, silent |
| 26.2.0 | works, silent |

Three lessons in one table. The commonly cited "available since 22.5" is true only behind `--experimental-sqlite`; unflagged availability starts at **22.13**. And "warning-free from 24" — the reason this floor was set at 24 rather than 22.13 — is **wrong at major granularity**: 24.11 still warns, because the module only reached Release Candidate mid-24.x. A major-only floor of `>=24.0.0` therefore still admits versions carrying the exact noise 22.13 was rejected for.

That is the trap: a floor expressed in majors cannot express a boundary that moved at a patch. Measure the *specific* APIs you call, at more than one patch level inside the major you are about to require, and say which points you actually tested.

## Why This Matters

The failure mode is worse than an ordinary bug because it is invisible to the person who caused it. Everyone shipping the change is above the floor by definition; CI runs only supported versions; every local check passes. The only people who see it are users, and what they see is a stack trace their host renders as "failed to connect."

A release that removes an illegible-startup failure can install a new one through a different door. Removing the compiled dependency made the old crash impossible, and raising the floor to do so made a new crash possible — same symptom, same user experience, different cause.

## When to Apply

- Any change to `engines.node` (or the equivalent floor in another ecosystem)
- Adopting a stdlib API gated on runtime version — check unflagged *and* warning-free availability separately
- Dropping a runtime major in a breaking release, especially for a CLI or server launched by a host that captures stderr

The *entry guard* is the only part a pure library can skip — there is no entry to guard. Everything else still applies: a consumer importing your ESM graph hits the unavailable module in their runtime, not at build time, so keep the manifest-alignment test and add a below-floor import test.

## Examples

Before — what a Node 20 user got from v5.0.0:

```text
node:internal/modules/esm/translators:391
    throw new ERR_UNKNOWN_BUILTIN_MODULE(url);
          ^
```

After — v5.0.1, verified by installing the published tarball and launching it under Node 20.20.2: the full 571-byte message arrives intact through a pipe, exit code 1.

## Related

- [`mcp-server-crash-on-unloadable-better-sqlite3-2026-07-12.md`](../runtime-errors/mcp-server-crash-on-unloadable-better-sqlite3-2026-07-12.md) — the earlier form of the same illegible-startup class, from the native binding. Shares the symptom and the module, differs on cause and remedy; worth reading together.
- jbctechsolutions/mcp-office365#76 — the issue whose re-reading surfaced this, narrowed from "pure-JS store fallback" to this guard
- jbctechsolutions/mcp-office365#108 — the driver swap that raised the floor
- jbctechsolutions/mcp-office365#113 — the fix
