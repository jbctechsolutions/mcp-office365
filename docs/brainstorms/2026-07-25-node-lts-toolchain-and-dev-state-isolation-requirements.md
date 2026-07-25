# Requirements: LTS toolchain + dev state isolation (issue #99)

**Date:** 2026-07-25
**Issue:** #99
**Scope:** Standard
**Status:** Ready for planning

## Problem

Running the dev build (Node-26-pinned toolchain from #88) against the same shared
state dir and npx cache that installed/published builds use leaves those installed
builds broken in two independent ways:

1. **Native-module ABI mismatch (hard crash → MCP `-32000`).** `better-sqlite3`
   built under Node 26 ships ABI `147`; an LTS-Node consumer (22 → `127`, 24 → `137`)
   cannot load it. The native-load failure is fatal (the in-memory fallback uses the
   same module), so the connection hard-crashes.
2. **State schema forward-incompatibility (silent degrade).** A dev-only 4th migration
   bumped the **shared** `~/.mcp-office365/state.db` to v4. Published builds cap at v3
   (`4.3.0`) and can't open it, silently dropping to in-memory (auth/state stop
   persisting).

Root cause is one design flaw with two faces: **#88 pinned the *release/publish* build
to a non-LTS Current Node**, so every published/prebuilt ABI is unloadable by LTS
consumers — and the dev build shares consumers' state dir + npx cache, so a dev
migration corrupts the shared `state.db`.

## Goal

Installed/published builds keep working regardless of what the dev toolchain does.
The dev build can never poison a consumer's ABI or migrate a consumer's `state.db`.

## Users / who hits this

Any machine that runs the dev build (or an npx launch under a newer Node) and later
falls back to the installed/published build. On this repo that's effectively the
maintainer's own machines — the direct-path install (`~/.mcp-servers/office365/...`,
launched under Node 22.15.0) is the supported consumer path, and it broke.

## Decisions

| # | Decision | Choice |
|---|----------|--------|
| 1 | Toolchain pin | **Node 24 LTS everywhere** — `.tool-versions`, `release.yml`, `integration.yml` → 24. Test matrix keeps 20/22/24/26 for coverage. Reverses #88's Current-Node pin. |
| 2 | Dev state isolation | **First-class `--state-dir` CLI flag** (alongside the existing `OUTLOOK_MCP_STATE_DIR` env var), and wire the dev launch to an isolated dir (e.g. `~/.mcp-office365-dev`). |
| 3 | Forward-schema / native errors | **Harden incrementally.** The forward-schema guard (`runMigrations` throws) and native-load remediation message already exist. Add: offending binding path in the rebuild hint; degrade message points at `--state-dir`. |
| 4 | npx path | **Document, don't prebuild.** Document the direct-path install as the supported launch; note the npx stale-cache-ABI trap. No multi-ABI prebuild work (ruled out by decision 1). |
| 5 | Regression guard | **Add one.** A test/startup assertion catching the exact #99 modes (build-ABI vs runtime-Node mismatch; schema newer than supported). |

## Requirements

### Functional
- The published/release artifact is built under Node 24 LTS; its native `better-sqlite3`
  binding loads under Node 20/22/24 without a rebuild.
- A `--state-dir <path>` CLI flag sets the state directory; precedence resolves cleanly
  against the existing `OUTLOOK_MCP_STATE_DIR` env var (flag beats env beats default —
  to confirm in planning).
- The dev launch (`npm run dev` / documented dev invocation) uses an isolated state dir
  distinct from the default `~/.mcp-office365`, so a dev migration never mutates the
  shared/installed `state.db`.
- When the native module can't load, the fatal error names the exact offending binding
  path plus the rebuild/cache-clear remedy.
- When `state.db` is at a schema newer than the build supports, the degrade message is
  actionable and mentions `--state-dir` as the isolation remedy.

### Non-functional
- No change to the default consumer state path (`~/.mcp-office365`) — installed builds
  keep reading their existing db.
- `tokens.json` handling is unchanged (schemaless; not the source of either failure).
- CI must exercise the LTS build path that actually ships.

## Success criteria
- A clean install launched under Node 22/24 boots, loads the native module, and reads a
  v3 `state.db` — no `-32000`, no in-memory degrade. (Reproduces the #99 workaround as a
  guaranteed steady state.)
- Running the dev build creates/migrates only its isolated dev state dir; the shared
  `~/.mcp-office365/state.db` is untouched (verify schema version unchanged after a dev run).
- The regression guard fails loudly if a future change reintroduces a Current-Node
  release pin or a shared-dir dev migration.

## Non-goals
- Multi-ABI prebuilt binaries (`prebuildify`) for arbitrary consumer Node versions.
- Reworking the state-store in-memory fallback architecture.
- Migrating or relocating `tokens.json`.
- Fixing/authoring the unreleased v4 migration itself — it lives on another branch and is
  out of scope here (this work only ensures it can't harm shared state).

## Dependencies / assumptions
- **Assumption:** reverting to Node 24 LTS carries no loss vs #88 — #88's stated aim
  ("avoid npx-cache ABI mismatches") is better served by an LTS pin, since the mismatch
  came from publishing a Current-Node ABI. To confirm nothing else in the repo requires
  Node 26 APIs.
- The `OUTLOOK_MCP_STATE_DIR` override already exists and is honored by `StateStore.open`
  (`src/index.ts`, `src/state/store.ts`) — the flag builds on it rather than replacing it.
- The forward-schema guard (`src/state/migrate.ts`) and native-load remediation
  (`src/state/store.ts`) already exist — decision 3 is polish, not net-new machinery.

## Open questions for planning
- `--state-dir` vs `OUTLOOK_MCP_STATE_DIR` precedence order and whether the flag applies
  to all subcommands (`serve`, `auth`, `revoke`) or just the server.
- Where the regression guard lives: a Vitest unit test, a CI assertion on the built
  artifact's ABI, or a runtime startup check — or a combination.
- Exact isolated dev dir name/location and how it's set (npm script env vs documented flag).
