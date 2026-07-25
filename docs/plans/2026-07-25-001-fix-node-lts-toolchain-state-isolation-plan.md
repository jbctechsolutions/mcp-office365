---
title: "fix: LTS toolchain + dev state isolation (issue #99)"
type: fix
status: active
date: 2026-07-25
origin: docs/brainstorms/2026-07-25-node-lts-toolchain-and-dev-state-isolation-requirements.md
---

# fix: LTS toolchain + dev state isolation (issue #99)

## Summary

Stop the dev toolchain from breaking installed builds: repin the local/CI Node
toolchain to Node 24 LTS, add a first-class `--state-dir` flag (resolving flag > env >
default) and wire it into every state-store open site so the dev build can run against
an isolated dir, harden the existing native-load and forward-schema error messages, and
add a regression guard that fails if a non-LTS Node is ever pinned again.

---

## Problem Frame

Running the dev build (Node-26 toolchain from #88) against the shared `~/.mcp-office365`
state dir and shared npx cache breaks installed builds two ways: a `better-sqlite3` ABI
mismatch (`NODE_MODULE_VERSION 147` vs an LTS consumer's `127`/`137`) that hard-crashes
the MCP connection with `-32000`, and a dev-only migration that bumped the shared
`state.db` to a schema no released build understands (silent in-memory degrade). See
origin for full symptom capture and issue #99.

A nuance that shapes the fix: the published package's `dist/` is pure JS, and
`better-sqlite3` resolves its ABI at the *consumer's* install time — so the primary ABI
lever is the **local `.tool-versions`** (which governs what the shared npx cache compiles
under) plus **dev state isolation**, not the release workflow's Node version. The
`release.yml`/`integration.yml` repins are consistency/coverage, not the root fix.

---

## Requirements

- R1. The local and CI toolchain pin to Node 24 LTS; no build/release step pins a
  non-LTS (odd-major/Current) Node.
- R2. A `--state-dir <path>` CLI flag sets the state directory, resolving with precedence
  flag > `OUTLOOK_MCP_STATE_DIR` env > default `~/.mcp-office365`, honored at every site
  that opens the state store.
- R3. The dev launch uses an isolated state dir distinct from `~/.mcp-office365`, so a dev
  migration can never mutate an installed build's `state.db`.
- R4. When the native module can't load, the error names the offending binding path plus
  the rebuild/cache-clear remedy; when `state.db` is newer than supported, the degrade
  message is actionable and mentions `--state-dir`.
- R5. The direct-path install is documented as the supported launch, with the npx
  stale-cache-ABI trap called out.
- R6. A regression guard fails loudly if a future change reintroduces a non-LTS Node pin
  or a forward-schema open regresses to a crash.

**Origin acceptance examples:** success criteria in origin — clean boot under Node 22/24
with a v3 `state.db`; a dev run leaves the shared `state.db` untouched; the guard trips on
a Current-Node pin.

---

## Scope Boundaries

- No multi-ABI prebuilt binaries (`prebuildify`) — ruled out by the Node-24 pin decision.
- No rework of the state-store in-memory fallback architecture (the bad-file-degrade vs
  bad-binding-fail-fast split from PR #77 stays intact).
- No migration or relocation of `tokens.json` (schemaless; not a source of either failure).
- Not authoring/fixing the unreleased v4 migration itself — it lives on another branch;
  this work only ensures it cannot harm shared state.

### Deferred to Follow-Up Work

- True no-SQLite graceful degrade (JS Map-backed store): tracked in #76 — the step beyond
  fail-fast, out of scope here.

---

## Context & Research

### Relevant Code and Patterns

- **CLI flag parsing:** mirror the value-flag idiom in `parseServeOptions` (`src/cli.ts`,
  the `--host`/`--host=` indexed-loop with throw-on-missing-value). Subcommand dispatch is
  `parseCliCommand` (`src/cli.ts`) consumed in `main()` (`src/index.ts`). No shared flag
  helper exists — parsing is per-subcommand.
- **State store open sites (three):** default stdio path `src/index.ts` (`options.stateStore ?? StateStore.open()` — currently ignores even the env var), `revoke` path `src/index.ts` (`handleRevokeCommand`), and the `serve` block `src/index.ts`. Only the
  latter two read `OUTLOOK_MCP_STATE_DIR` today.
- **StateStore.open signature:** `StateStore.open(options: StateStoreOptions = {})` in
  `src/state/store.ts`; `StateStoreOptions.dir?` overrides `DEFAULT_DIR = ~/.mcp-office365`.
  Passing `{ dir }` is the only wiring needed — dir creation, pragmas, migrations, 0600
  perms all happen inside `open`. Preserve the process-scoped injection seam (used by both
  `src/index.ts` and `src/remote/http-server.ts`).
- **Native-load classification:** `isNativeLoadFailure()` in `src/state/store.ts` matches
  four shapes (`ERR_DLOPEN_FAILED`, `NODE_MODULE_VERSION` message, `MODULE_NOT_FOUND` for
  `better_sqlite3.node`, and the code-less `bindings` "Could not locate the bindings file").
  Regression test: `tests/unit/state/native-load-failure.test.ts`.
- **Forward-schema guard:** `runMigrations` in `src/state/migrate.ts` already throws when
  the on-disk schema exceeds `MIGRATIONS.length`; the caller degrades to in-memory.
- **Node pin sites:** `.tool-versions`, `.github/workflows/test.yml` (matrix + the `'26.x'`
  primary-version guards at the lint/coverage steps), `.github/workflows/release.yml`,
  `.github/workflows/integration.yml`. `package.json` `engines` is a consumer floor
  (`>=20.0.0`). Cosmetic: stale `v20.10.0` placeholder in `.github/ISSUE_TEMPLATE/bug_report.yml`.
- **Test conventions:** Vitest; tests under `tests/unit/**/*.test.ts` mirroring `src/`,
  `.js` ESM imports; state tests use `mkdtempSync` temp dirs and **close the store before
  `rmSync`** (Windows `EBUSY`); CLI parser tests are pure-function `toEqual`/`toThrow`.

### Institutional Learnings

- `docs/solutions/runtime-errors/mcp-server-crash-on-unloadable-better-sqlite3-2026-07-12.md`
  — the direct ancestor (shipped v4.2.1 / PR #77). Keep the bad-file-degrade vs
  bad-binding-fail-fast split; re-verify all four `isNativeLoadFailure` shapes fire under
  Node 24; do not regress the "no false degrade warning on binding failure" assertion.
- `docs/solutions/integration-issues/device-code-auth-undefined-invalid-grant.md` — house
  convention is `OUTLOOK_MCP_*` env vars + fail-fast-with-setup-guidance. `--state-dir` and
  its resolution should follow this shape.
- `docs/solutions/conventions/adversarial-review-as-primary-gate.md` — upgrade-boundary
  coercion (old `state.db` rows stay readable) and "confirm which CI leg enforces the
  ABI/coverage gate" both apply to this toolchain/schema change.
- `better-sqlite3` is already `^12.11.1`, so the #70 `EBADENGINE` engine-cap (capped at
  25.x) is already resolved — Node 24 is safe with no dependency bump.

---

## Key Technical Decisions

- **Node 24 LTS everywhere; keep the test matrix.** `.tool-versions`, `release.yml`,
  `integration.yml`, and the `test.yml` primary-version guards move to 24; the test matrix
  keeps 20/22/24/26 so forward-compat is still exercised. Rationale: consumers run LTS; the
  shared npx cache must compile `better-sqlite3` under the same ABI class installed builds
  load.
- **Single state-dir resolution helper, global wiring.** Add one resolver (flag > env >
  default) and call it at all three store-open sites, closing the latent gap where the
  default stdio path ignores even `OUTLOOK_MCP_STATE_DIR`. Rationale: a flag that silently
  no-ops on the most common launch is worse than no flag.
- **Regression guard = LTS-pin assertion test.** A Vitest test reads the pin sites and
  fails if the pinned Node major is odd (Current). Rationale: catches the exact #99
  reintroduction cheaply; cross-Node ABI is already covered by the CI matrix.
- **`engines.node` stays `>=20.0.0`.** It's the consumer floor, not a toolchain pin; no
  reason to raise it.

---

## Open Questions

### Resolved During Planning

- **`--state-dir` precedence?** flag > `OUTLOOK_MCP_STATE_DIR` > default (conventional
  flag-over-env).
- **Which subcommands honor the flag?** All store-open sites (default stdio, `serve`,
  `revoke`); `auth` doesn't touch the state store, so no wiring there.
- **Does Node 24 need a `better-sqlite3` bump?** No — already `^12.11.1` (> the #70 cap).
- **Guard shape?** LTS-pin assertion test over the pin files, not a runtime ABI probe.

### Deferred to Implementation

- Exact resolver name/location (likely a small exported function in `src/cli.ts` or a
  `src/state` helper) — decide when wiring the three call sites.
- Exact isolated dev dir name (e.g. `~/.mcp-office365-dev`) and whether dev wiring is an
  npm script, a `.mcp.json` dev entry, or a documented `--state-dir` invocation.

---

## Implementation Units

### U1. Repin toolchain to Node 24 LTS

**Goal:** Every toolchain pin targets Node 24 LTS; nothing pins a non-LTS Node.

**Requirements:** R1

**Dependencies:** None

**Files:**
- Modify: `.tool-versions` (→ `nodejs 24.x` latest patch)
- Modify: `.github/workflows/release.yml` (node-version → 24)
- Modify: `.github/workflows/integration.yml` (node-version → 24)
- Modify: `.github/workflows/test.yml` (move the `'26.x'` primary-version lint/coverage
  guards to `'24.x'`; keep the matrix `[20.x, 22.x, 24.x, 26.x]`)
- Modify: `.github/ISSUE_TEMPLATE/bug_report.yml` (refresh stale `v20.10.0` placeholder)

**Approach:**
- Pick the current 24.x patch for `.tool-versions` and the workflow pins.
- Keep the matrix intact so 26.x stays exercised; only the *primary/coverage* leg and the
  *release/integration* pins move to 24.
- Leave `package.json` `engines` at `>=20.0.0`.

**Patterns to follow:** commit b9a11c2 (#88) made the inverse change — mirror the same
touch surface in reverse.

**Test scenarios:** Test expectation: none for the config edit itself — behavior is
asserted by U6's LTS-pin guard and by CI going green on the 24 legs.

**Verification:** CI runs green with 24 as the primary leg; `asdf` resolves Node 24
locally; no pin references an odd-major Node.

---

### U2. `--state-dir` flag + global state-dir resolution

**Goal:** A `--state-dir <path>` flag resolves flag > env > default and is honored at every
state-store open site.

**Requirements:** R2

**Dependencies:** None

**Files:**
- Modify: `src/cli.ts` (add value-flag parsing for `--state-dir`/`--state-dir=`; add a
  resolver applying flag > `OUTLOOK_MCP_STATE_DIR` > default)
- Modify: `src/index.ts` (call the resolver at all three `StateStore.open` sites: default
  stdio path, `handleRevokeCommand`, `serve` block)
- Test: `tests/unit/cli.test.ts`

**Approach:**
- Copy the `parseServeOptions` `--host` idiom: support both `--state-dir <path>` and
  `--state-dir=<path>`, throw on missing/`--`-prefixed value, ignore unknown args.
- Centralize precedence in one resolver so all three sites agree; keep the existing
  `StateStore.open({ dir })` seam and the process-scoped injection.
- Replace the two inline `process.env.OUTLOOK_MCP_STATE_DIR` reads with the resolver.

**Execution note:** Implement the parser/resolver test-first — it's a pure function, the
cheapest place to pin precedence.

**Patterns to follow:** `parseServeOptions` value-flag loop (`src/cli.ts`); existing
`OUTLOOK_MCP_STATE_DIR` idiom in `handleRevokeCommand`/`serve`.

**Test scenarios:**
- Happy path: `--state-dir /tmp/x` → resolver returns `/tmp/x`.
- Happy path: `--state-dir=/tmp/x` → resolver returns `/tmp/x`.
- Edge: no flag, `OUTLOOK_MCP_STATE_DIR` set → returns the env value.
- Edge: no flag, no env → returns the default `~/.mcp-office365`.
- Edge: both flag and env set → flag wins (precedence).
- Error path: `--state-dir` with no following value (or next token starts with `--`) →
  throws a clear error.
- Edge: unknown sibling args present → ignored, flag still parsed.

**Verification:** All three open sites route through the resolver; a set env var now also
affects the default stdio launch (previously ignored); parser suite green.

---

### U3. Isolate the dev launch's state dir

**Goal:** The dev launch runs against an isolated state dir, never `~/.mcp-office365`.

**Requirements:** R3

**Dependencies:** U2

**Files:**
- Modify: `package.json` (add a dev serve script that passes `--state-dir` at an isolated
  dir) and/or `.mcp.json` (dev entry) — exact vehicle decided at implementation
- Modify: `README.md` or `docs/` (document the isolated dev invocation)

**Approach:**
- Use the U2 flag to point dev at e.g. `~/.mcp-office365-dev`. `dev` today is `tsc --watch`
  (build only); add a distinct serve-with-isolated-dir path rather than overloading it.
- Belt-and-suspenders with U1: aligning `.tool-versions` to 24 already removes the ABI
  poisoning; the isolated dir removes the schema-poisoning.

**Test scenarios:** Test expectation: none — config/script wiring; behavior covered by U2's
resolver tests and U6's isolation assertion.

**Verification:** Running the dev launch creates/migrates only the isolated dir; the shared
`~/.mcp-office365/state.db` schema version is unchanged after a dev run.

---

### U4. Harden native-load and forward-schema messages

**Goal:** Failure messages name the offending path and point at the isolation remedy.

**Requirements:** R4

**Dependencies:** U2 (so the degrade message can reference `--state-dir`)

**Files:**
- Modify: `src/state/store.ts` (native-load remediation: include the offending binding
  path; keep all four `isNativeLoadFailure` shapes)
- Modify: `src/state/store.ts` and/or `src/state/migrate.ts` (forward-schema degrade
  message mentions `--state-dir`/`OUTLOOK_MCP_STATE_DIR` as the isolation remedy)
- Test: `tests/unit/state/native-load-failure.test.ts` (extend), `tests/unit/state/store.test.ts`

**Approach:**
- Additive to PR #77's messaging — do not re-couple bad-file-degrade with
  bad-binding-fail-fast.
- Preserve the "no false degrade warning on binding failure" assertion.

**Test scenarios:**
- Error path: native-load failure → error string includes the offending binding path and
  the rebuild/cache-clear remedy; `cause` preserved; `process.version` named.
- Error path: forward-schema open (on-disk version > `MIGRATIONS.length`) → degrade message
  is emitted, is actionable, and mentions the state-dir isolation remedy.
- Regression: all four `isNativeLoadFailure` shapes still classify under Node 24.
- Edge: bad-file (corrupt db, loadable binding) still degrades to `:memory:` without the
  native-load remediation text.

**Verification:** Extended tests green; messages readable and path-specific.

---

### U5. Document supported launch + npx trap

**Goal:** Direct-path install is documented as supported; the npx stale-cache-ABI trap is
called out; CHANGELOG updated.

**Requirements:** R5

**Dependencies:** U2, U3

**Files:**
- Modify: `README.md` (supported direct-path install; `--state-dir` usage; npx caveat)
- Modify: `CHANGELOG.md` (`[Unreleased]` entry for the toolchain repin, `--state-dir`,
  and message hardening)
- Optionally add: `docs/solutions/` note if a durable gotcha writeup fits (defer to
  ce-compound rather than forcing here)

**Approach:**
- Explain that npx inherits whatever ABI the cache last compiled under; recommend the
  direct-path install (`~/.mcp-servers/...`) or matching the runtime Node.

**Test scenarios:** Test expectation: none — docs only.

**Verification:** README states the supported launch and the npx caveat; `[Unreleased]`
reflects the change set.

---

### U6. Regression guards

**Goal:** Fail loudly if a non-LTS Node pin returns or a forward-schema open regresses.

**Requirements:** R6

**Dependencies:** U1, U2

**Files:**
- Create: `tests/unit/toolchain-pins.test.ts` (assert pinned Node major is even/LTS across
  `.tool-versions`, `release.yml`, `integration.yml`)
- Test: `tests/unit/state/store.test.ts` (forward-schema open degrades, does not crash)

**Approach:**
- The pin test reads the files as text and extracts the major version; odd major → fail
  with a message pointing back to #99.
- The forward-schema test writes a `state.db` with `schema_version` above
  `MIGRATIONS.length`, opens via `StateStore.open`, and asserts a degraded store plus an
  actionable warning (not a throw to the caller).

**Test scenarios:**
- Happy path: all pin files at an even/LTS major → test passes.
- Error path: a pin file set to an odd major → test fails with the #99-referencing message.
- Edge: forward-schema `state.db` → `open` returns a degraded store, emits the warning,
  does not throw to the caller. (close before `rmSync`.)

**Verification:** Guards pass on the Node-24 tree; flipping any pin to an odd major or
regressing the forward-schema handling turns them red.

---

## System-Wide Impact

- **Interaction graph:** state-dir resolver feeds three `StateStore.open` call sites in
  `src/index.ts`; the process-scoped store injection (`src/index.ts`, `src/remote/http-server.ts`) is preserved, not rerouted.
- **Error propagation:** native-load stays fatal-with-remediation (can't degrade — same
  module backs the fallback); forward-schema stays a caught degrade-to-in-memory, not a
  crash.
- **State lifecycle risks:** the whole point — dev migrations must not touch the shared
  `state.db`. Isolation (U3) plus the LTS pin (U1) remove both poisoning vectors.
- **API surface parity:** `--state-dir` complements, does not replace, `OUTLOOK_MCP_STATE_DIR`; both remain honored.
- **Unchanged invariants:** default state path stays `~/.mcp-office365`; `tokens.json`
  handling, the PR #77 degrade/fail-fast split, and `engines.node` floor are unchanged.

---

## Risks & Dependencies

| Risk | Mitigation |
|------|------------|
| Repinning to 24 loses something #88 wanted from 26 | #88's stated aim ("avoid npx-cache ABI mismatches") is better served by an LTS pin; matrix keeps 26 coverage. Verify no repo code needs Node-26-only APIs (none found). |
| A store-open site is missed, so `--state-dir` silently no-ops there | Centralize resolution in one helper and wire all three sites in U2; U6 + manual boot check confirm the default stdio path now honors it. |
| Message hardening accidentally re-couples degrade paths (regresses #77) | Additive edits only; keep the four-shape classifier and the "no false degrade warning" assertion (U4 tests). |
| LTS-pin guard is brittle to workflow YAML formatting | Parse tolerantly (extract major only); assert the invariant (even major), not exact strings. |

---

## Documentation / Operational Notes

- README: supported direct-path install, `--state-dir`, npx caveat (U5).
- CHANGELOG `[Unreleased]`: toolchain repin, `--state-dir`, message hardening (U5).
- No runtime migration or rollout step for consumers — default path and schema are
  unchanged; existing installs keep reading their v3 `state.db`.

---

## Sources & References

- **Origin document:** docs/brainstorms/2026-07-25-node-lts-toolchain-and-dev-state-isolation-requirements.md
- Related code: `src/cli.ts`, `src/index.ts`, `src/state/store.ts`, `src/state/migrate.ts`
- Related learnings: `docs/solutions/runtime-errors/mcp-server-crash-on-unloadable-better-sqlite3-2026-07-12.md`
- Related PRs/issues: #99 (this), #88 (Node-26 pin being reversed), #77 (native-load
  fail-fast baseline), #76 (no-SQLite degrade follow-up), #70 (`EBADENGINE` engine cap)
