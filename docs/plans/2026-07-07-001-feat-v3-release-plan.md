---
title: "feat: mcp-office365 v3.0.0 — registry architecture, durable IDs, reliability core, MCP-native surface"
type: feat
status: active
created: 2026-07-07
origin: docs/ideation/2026-07-07-v3-release-ideation.md
---

# feat: mcp-office365 v3.0.0

## Summary

Major release combining the fixes for every failure class observed in 4 months of real usage (621 calls, ~26% error rate) with new capabilities. Seven workstreams: (1) registry-driven tool architecture + contract harness, (2) self-resolving durable IDs, (3) canonical entity vocabulary, (4) structured search, (5) reliability core (transport + durable state), (6) MCP-native surface (presets/annotations/elicitation), (7) delta-sync local mirror. Workstream 7 may slip to v3.1; everything else must land.

**Origin:** `docs/ideation/2026-07-07-v3-release-ideation.md` (7 survivors, all adopted). User decisions already made: all 7 workstreams in scope; AppleScript backend frozen + deprecated (functional, no new v3 features, removal targeted v4); confirmation stays durable two-phase by default with elicitation behind `--confirm elicit`.

---

## Problem Frame

The server's 219 tools are defined across four independently hand-maintained lists that drift (JSON-schema `TOOLS` array, per-module Zod schemas, dispatch switches, `GRAPH_ONLY_TOOL_NAMES`), IDs are session-scoped lossy hashes that strand agents after restarts, there is no retry/transport layer, approval tokens die in memory, and the static 219-tool surface taxes every session's context. The 10 observed failure classes (KQL rejections, `id`/`email_id` fumbles, field drift, cache misses, auth expiry, broken downloads, `$top` rejection, expired approval tokens, raw 5xx, inconsistent taxonomy) all trace to these structural causes.

## Scope Boundaries

**In scope:** everything in the seven workstreams below, release engineering (version 3.0.0, CHANGELOG migration matrix, README), CI dual-publish fix.

### Deferred to Follow-Up Work
- Workstream 7 (delta mirror) slips to v3.1 if it endangers the release — explicit user decision.
- Elicitation as *default* confirm mode (v3.1, after real-world trial behind the flag).
- Webhook change notifications (requires reachable endpoint; delta polling captures the value).
- `docs/solutions/` knowledge base seeding from v3 learnings.
- AppleScript backend removal (v4).

### Outside this product's identity
- Generic Graph passthrough tools (lokka-style) — the curated tool surface is the product.
- Multi-tenant server deployment.

---

## Key Technical Decisions

These resolve the design questions raised by flow analysis. Numbered for traceability from units.

**D1 — Token format: hybrid — self-encoding tokens for simple entities, alias-backed tokens for composite/mutable.** *(Revised after doc review: a one-way hash cannot be reversed to a Graph ID on a cold `state.db`, which would reintroduce the exact cache-miss class this release kills. Self-encoding is required wherever cold-state resolution must work without storage.)*
- **Simple single-key entities** (`em_` message, `ev_` event, `ct_` contact, `fd_` folder, `td_` task, `dr_` driveItem): the token *self-encodes* the immutable Graph ID — `<prefix>_<base64url(immutableGraphId)>`. Resolution is `base64url-decode`, **zero storage**, so a token pasted on a fresh machine or after `state.db` loss resolves immediately. This delivers the ideation's "resolves forever, no cache lookups" promise for the common case. Tokens are long (~200–350 chars); the context-efficiency goal is served by Workstream 6 (presets/fewer tools), not by ID length — durability wins for Workstream 2's core promise.
- **Composite and mutable entities** (`at_` attachment `{messageId, attachmentId}`, channels/chat-messages/checklist-items tuples, and any `$search`-minted mutable-ID row) require a stored tuple regardless, so they use a short deterministic token `<prefix>_<base32(sha256(canonicalKey))[0..13]>` backed by the alias table (D3). These are **machine-scoped**: a composite token does not resolve on a cold `state.db` and returns `ID_UNKNOWN` with a re-list hint — documented in the migration matrix as an accepted, narrow limitation (you rarely paste an attachment ID across machines).
- The alias table (D3) is append-only, **no eviction**, and functions as an optimization/reverse-index for self-encoded entities and the source of truth for composite/mutable ones. For the short composite tokens, a 70-bit hash prefix makes collisions negligible; on insert-time collision (same token, different key) the store raises a typed integrity error rather than mis-resolving (no runtime variable-length handling — see D1a).

**D1a — Collision policy.** Composite tokens use a fixed length. An insert-time collision (astronomically unlikely at this scale) surfaces `ID_COLLISION` and is logged; there is no dynamic token-length extension (removed as unjustified complexity — it is not one of the observed failure classes).

**D2 — Immutable ID preference with upgrade-on-mint.** Wherever Graph supports `Prefer: IdType="ImmutableId"` (list/get on messages/events), mint self-encoding tokens (D1) from the immutable ID — these are the durable, cold-state-resolvable tokens. `$search` paths return mutable IDs — the minting layer translates to immutable via `translateExchangeIds` in batch at mint time so search results also yield durable self-encoded tokens; **if translation fails or is rate-limited for part of a listing**, mint a short alias-backed token marked `mutable=1` (D1, machine-scoped), and the resolver re-resolves (re-list or translate) on 404, returning typed `ID_STALE` with a recovery hint if that fails. A partially-translated listing is not silently mixed: the response envelope flags `degraded_ids: <count>` when any row fell back to a mutable alias token, so the agent knows a subset is machine-scoped.

**D3 — Alias table ships in Unit 4's state.db from day one; the delta mirror (U12) later *joins* it.** A v3.1 slip of the mirror cannot strand the ID workstream.

**D4 — AppleScript ID compatibility.** ID params are typed `string | number` unions. Numeric IDs pass through unchanged on the AppleScript backend (its native rowid space). On the Graph backend, a numeric ID yields typed `NUMERIC_ID_UNSUPPORTED` with a hint to re-list (v2 hash IDs are lossy and unrecoverable by design).

**D5 — Retry-exclusion table.** Retry with jittered exponential backoff honoring `Retry-After` applies to: 429 (all *retriable* verbs — see exclusion), 502/503/504 + connection-reset (GET/idempotent only, and POST `/$batch` whose sub-requests are all reads). Never auto-retry: `sendMail`/`confirm_send_*` (even on 429, after the body is sent → `retriable: false` + receipt-check suggestion), 401 (route to auth refresh, never loop under the auth mutex), 412 (Planner ETag — one fetch-then-retry, see D6), any streamed download mid-body (restart from byte 0). **Implementation note (from doc review):** the SDK's default middleware chain (`Client.init`, used today) already includes a `RetryHandler` that auto-retries 429/503/504 on *all* buffered requests including `sendMail` — directly violating this rule. U8 must therefore use `Client.initWithMiddleware` with a middleware array whose custom RetryHandler enforces this table and **replaces** the default RetryHandler; the "wrap the auth-provider fetch" fallback is only acceptable if it also disables the SDK default retry (`RetryHandlerOptions maxRetries=0`).

**D6 — Planner ETags are never persisted.** Updates do fetch-then-patch: GET for a fresh ETag immediately before PATCH, one retry on 412. Removes the ETag/ID-cache coupling.

**D7 — state.db concurrency.** `~/.mcp-office365/state.db`, WAL mode, `busy_timeout=5000`, atomic token consume (`UPDATE … RETURNING`-style guarded write). Cross-process redemption is *allowed by design* (a token prepared in one Claude Code window redeemable in another is a feature); every row is stamped with the MSAL home-account-id, and rows from a foreign account yield `ID_FOREIGN_ACCOUNT`.

**D8 — Approval token semantics.** Durable (SQLite), TTL 24h (up from 5 min), idempotent redemption keyed on the *operation* via an **atomic claim** (a unique-constraint INSERT / guarded `UPDATE … RETURNING` on the receipts journal that must win *before* the Graph call fires; a losing claim returns the existing receipt and never re-executes). Content-hash sealed over target-type-specific critical fields — for `send` this includes **recipients + subject + body-hash + attachment-list hash (IDs + byte-hashes)** so a mutated attachment set cannot pass the seal; `delete`: target id + subject; `upload`: path + bytes-hash. Seal mismatch at confirm returns `APPROVAL_TARGET_CHANGED` with a fresh preview and re-prepare hint. Boot-time purge drops `approval_tokens`/`receipts` rows past a 90-day retention window (distinct from the 24h redemption TTL) so recipient/subject content does not accumulate indefinitely.

**D9 — Structured search compilation.** Pure-property queries (sender, date range, flags) compile to `$filter`; pure free-text compiles to quoted `$search`; mixed free-text + property queries compile to server-built KQL against `POST /search/query` (Graph's `$filter`/`$search` exclusivity on messages makes this the only single-request path). Execution note: spike the `/search/query` path against a real mailbox before freezing the schema (integration workflow exists).

**D10 — Error envelope.** Every tool failure returns `{ code, message, retriable, suggestion }` with a stable code vocabulary (`AUTH_EXPIRED`, `ID_STALE`, `ID_UNKNOWN`, `ID_FOREIGN_ACCOUNT`, `NUMERIC_ID_UNSUPPORTED`, `APPROVAL_TARGET_CHANGED`, `THROTTLED`, `GRAPH_UNAVAILABLE`, `VALIDATION`, `READ_ONLY_MODE`, …). Ends the GRAPH_ERROR/DATABASE_ERROR inconsistency.

**D11 — Alias coercion rule.** Input normalization coerces `id` → the tool's ID param **only when the tool has exactly one ID-typed param**; multi-ID tools (e.g. move: email + folder) reject bare `id` with a hint naming both params. String↔number coercion applies everywhere per D4.

**D12 — Presets default to `all`.** Existing users' surface must not shrink on upgrade; presets are opt-in via `--preset`. Next-action hints are filtered against the live registry surface so a hint never names an unexposed tool.

**D13 — Read-only mode** (`--read-only`): registry filters out every tool whose annotation is not read-only — including `prepare_*` (they mint credentials) — and token redemption is refused at runtime (`READ_ONLY_MODE`) as defense in depth.

**D14 — Elicitation** (`--confirm elicit`): capability-detected at initialize; silently behaves as durable two-phase when the client lacks elicitation. On the 60s timeout the pending elicitation is cancelled *before* the durable fallback token is returned (never both live); a decline revokes any pre-minted token. **The elicit-accept execution path routes through the identical atomic operation_key claim as token redemption (D8)** — both must win the atomic claim before the Graph call fires, so a late accept arriving after the timeout fallback token has already been redeemed loses the claim and executes zero additional times. Idempotency is therefore a hard guarantee, not a timing assumption.

**D15 — Migration & degraded mode.** `~/.mcp-office365` is already the current state dir (token cache moved there in v2.x); v3 boot creates `state.db` there with **`0700` dir / `0600` file permissions** (D18), and if legacy `~/.outlook-mcp` exists, copies tokens.json only when the new dir has none (never merge/overwrite), leaves the old dir untouched, and writes a migration marker for idempotence. Corrupt/locked state.db at boot degrades to in-memory state with a stderr warning (server stays usable; durability features degrade). **Because degraded-mode composite/mutable tokens are non-durable but otherwise byte-indistinguishable, minted tokens in this mode carry a non-durable marker and every response/resolve envelope while degraded includes `degraded: true` + a "state store unavailable; IDs are session-scoped this run" suggestion** — so an agent or human never saves a token expecting durability it silently won't get. (Self-encoded simple-entity tokens still resolve in degraded mode — they need no store — so only composite/mutable tokens are affected.)

**D18 — At-rest protection of state.db.** state.db holds 24h send/delete-authorizing approval tokens and sits beside the MSAL token cache. OS file permissions are the at-rest control: create `~/.mcp-office365` `0700` and `state.db` `0600` on creation, verify/repair on boot. The trust boundary is explicit and documented: **redemption trusts any process running as the same OS user / MSAL identity** — a malicious local process with home-directory read access is out of scope for cryptographic defense (same posture as the existing plaintext MSAL cache). SQLCipher/OS-keychain encryption is deferred (noted in risk table); permissions + the documented boundary are the v3 bar.

**D16 — Contract harness scope.** Registry invariants run over both backends' registered surfaces; VCR-style record/replay fixtures are Graph-only. AppleScript is frozen — characterization only, no new fixtures.

**D17 — Entrypoint/packaging stability.** `dist/index.js` remains the bin and sole export; `createServer(options)` gains a config parameter. The `isMainModule` heuristic is preserved (tests import `createServer` without auto-start). Server version string is read from package.json (fixes hardcoded `'0.1.0'`).

---

## High-Level Technical Design

*This illustrates the intended approach and is directional guidance for review, not implementation specification.*

```
                       ┌──────────────────────────────────────────────┐
  MCP client ──stdio──▶│ server.ts (entry: CLI flags → createServer)  │
                       │  ListTools / CallTool ◀── ToolRegistry       │
                       └──────────────┬───────────────────────────────┘
                                      │ registry entries: { name, description,
                                      │   inputSchema (zod → z.toJSONSchema),
                                      │   annotations, destructive, presets[],
                                      │   backends[], handler(ctx, args) }
                       ┌──────────────▼───────────────┐
                       │ input pipeline: alias coerce  │  D11
                       │ → zod parse → typed errors    │  D10
                       └──────────────┬───────────────┘
              ┌───────────────────────┼──────────────────────┐
      Graph handlers          approval manager         AppleScript handlers
              │              (durable tokens, D7/D8)        (frozen)
      ┌───────▼────────┐              │
      │ GraphTransport │ retry/backoff/Retry-After (D5), download
      │ (one chokepoint)│ normalization, error mapping (D10)
      └───────┬────────┘
              │                ┌────────────────────────────┐
      Microsoft Graph          │ StateStore (state.db, D7)  │
                               │  aliases (D1–D3) │ tokens  │
                               │  delta links │ receipts    │
                               └────────────────────────────┘
```

Registry entry sketch (directional):

```
defineTool({
  name: 'get_email',
  description: '…',
  input: GetEmailInput,            // zod; JSON Schema derived at registration
  annotations: { readOnlyHint: true },
  presets: ['mail'],
  backends: ['graph', 'applescript'],
  handler: async (ctx, params) => { … },   // ctx: { repo, ids, store, config }
})
```

---

## Implementation Units

Phases: **A** foundation (U1–U3) → **B** state & IDs (U4–U6) → **C** search & reliability (U7–U9) → **D** surface & release (U10–U13). Dependencies are explicit per unit.

### U1. Tool registry core + pilot domain migration

**Goal:** A `ToolRegistry` that is the single source of truth for name, description, input schema (Zod → JSON Schema via zod 4 `z.toJSONSchema`), MCP annotations, destructive flag, preset membership, backend availability, and handler — with ListTools/CallTool served from it. One pilot domain (mail-rules: modern module style, small) fully migrated; hybrid dispatch (registry first, legacy switch fallback) keeps everything else working.

**Requirements:** Workstream 1 (origin ideation idea 7).
**Dependencies:** none.
**Files:** `src/registry/types.ts`, `src/registry/registry.ts`, `src/registry/define-tool.ts` (new); `src/index.ts` (ListTools/CallTool handlers route through registry with fallback); `src/tools/mail-rules.ts` (pilot: export registry entries); `tests/unit/registry/registry.test.ts` (new).
**Approach:** Registry entries carry everything the four drifting lists carry today (`TOOLS` array `src/index.ts:313-3628`, Zod schemas in modules, dispatch switches, `GRAPH_ONLY_TOOL_NAMES` `src/index.ts:3766-3907`). Handler context object carries repo/token-manager/config so the 23-positional-parameter `handleGraphToolCall` signature dies. JSON Schema output must match MCP `Tool.inputSchema` shape. **`z.toJSONSchema` note:** it emits `additionalProperties: false` by default, which would advertise the D11/U6 transitional alias keys (`id`, `start`, `body`) as forbidden even though server-side coercion accepts them — a strict client would reject alias calls before coercion runs. U1 decides the toJSONSchema options and documents them: emit the transitional aliases as optional properties in the advertised schema so coercion is discoverable and validation-safe. Add a test asserting the emitted schema's `additionalProperties`/`$schema` handling matches what the MCP SDK accepts. Preserve lazy backend init + auth mutex behavior.
**Patterns to follow:** modern tool-module style (`src/tools/mail-rules.ts` narrow structural repo interface + token manager); ESLint strictness (explicit return types).
**Test scenarios:**
- Registering two tools with the same name throws at startup.
- `listTools()` returns MCP-shaped tools whose JSON Schema round-trips the Zod schema (spot-check `create_mail_rule` required/optional/enum fields).
- CallTool for a registry-migrated tool dispatches to its handler with parsed params; CallTool for an unmigrated tool falls through to the legacy switch unchanged.
- Registry filters by backend: AppleScript mode excludes graph-only entries (parity with `GRAPH_ONLY_TOOL_NAMES` for the pilot domain).
- Annotations present on every registered entry (readOnlyHint for list/get; destructiveHint for confirm-deletes).
**Verification:** server starts in both backends; pilot-domain tools behave identically to v2 through an InMemoryTransport client; typecheck/lint/tests green.

### U2. Full registry migration — dismantle the monolith

**Goal:** All 219 tools registered via domain modules; `TOOLS` array, all five dispatch switch functions, and `GRAPH_ONLY_TOOL_NAMES` deleted; `src/index.ts` reduced to entrypoint + `createServer(options)` wiring; `src/graph/repository.ts.bak` deleted and `*.bak` gitignored; server version read from package.json (D17); CI dual-publish fixed.
**Requirements:** Workstream 1.
**Dependencies:** U1.
**Files:** all `src/tools/*.ts` (each exports its registry entries); `src/index.ts` (shrinks drastically); `src/server.ts` (new: createServer with options); delete `src/graph/repository.ts.bak`; `.gitignore`; `.github/workflows/publish.yml` (drop the duplicate npm-publish path — release.yml owns tag-driven publishing); `tests/integration/server.test.ts` (expected-tool-list assertions move to registry-derived).
**Approach:** Mechanical but large. Legacy factory-style domains (mail, calendar, contacts, tasks) have Graph logic living in `handleGraphToolCall` switch cases, not in the classes — extract those case bodies into registry handlers per domain. Keep output-shaping transforms (`transformEmailRow` etc.) as shared helpers in a mappers module for now (U6 revisits shapes). Preserve `isMainModule` semantics (D17). **Behavioral gate (from doc review):** name-set parity proves the *surface* matches but nothing about extracted handler *behavior*, and this unit deletes the legacy fallback — so before deleting each domain's switch, record v2 output snapshots for a representative call **per tool** (not just per entity) and assert the v3 registry handler is byte-identical modulo ID fields. This promotes the U5 characterization technique into U2 as the blocking per-domain merge gate. **Coverage (from doc review):** `test.yml` enforces the global 75% threshold with no continue-on-error, so a PR that migrates a domain faster than its tests land reds CI. Each domain's tests (including the branch coverage of its handler error paths) MUST land in the **same PR** as its migration; do not split migration and tests across PRs behind the global gate.
**Execution note:** migrate domain-by-domain with the U3 harness (built in parallel) run after each domain; one PR per domain carrying migration + tests together.
**Test scenarios:**
- Registry-derived tool list matches v2's tool list exactly (name set parity in Graph mode and AppleScript mode) — surface no-regression gate.
- Per-tool behavioral snapshot: v3 handler output byte-identical to v2 (captured pre-deletion) modulo ID fields, for a representative call of every migrated tool — behavioral no-regression gate.
- Every tool name in the registry has a handler that Zod-parses its own advertised schema (no orphan schemas, no orphan handlers).
- `createServer({})` defaults preserve v2 behavior.
- CI: publish.yml no longer publishes on release events (workflow lint / manual inspection).
**Verification:** `wc -l src/index.ts` drops below ~300; full suite green; e2e InMemoryTransport test lists 219 tools and calls representative ones per domain.

### U3. Contract harness + failure-corpus regression fixtures

**Goal:** Registry-iterating invariant tests that make the observed bug classes structurally impossible to reintroduce, plus record/replay Graph fixtures encoding the 10 real failure classes.
**Requirements:** Workstream 1; flow-analysis recommendations (harness invariants pre-date the features they guard).
**Dependencies:** U1 (registry to iterate); grows with U2/U5/U6/U10.
**Files:** `tests/contract/invariants.test.ts`, `tests/contract/failure-corpus.test.ts`, `tests/fixtures/graph/` (canned Graph JSON fixtures), `tests/fixtures/fake-graph-client.ts` (new in-memory fake with canned entity data) (all new).
**Approach:** Invariants iterate the live registry (D16): (a) every ID field name a list/search tool returns is accepted by the paired get/update/delete tool; (b) create/get/update for one entity share one field vocabulary; (c) every download tool round-trips a mocked binary response to bytes on disk; (d) every next-action hint target is a registered tool under every preset (D12); (e) prepare/confirm pairs absent under `--read-only` (D13); (f) durable-ID round-trip mint→resolve for every entity type (added in U5). Failure corpus: one named regression test per observed class (KQL rejection inputs, `id` vs `email_id` calls, `start`/`end` update payloads, restart-mid-approval, ReadableStream download, `$top` on onlineMeetings, expired-token redemption, device_code mid-session, 502 retry, taxonomy consistency). Fixtures via the fake Graph client — no live HTTP in CI.
**Test scenarios:** the unit IS test scenarios; the meta-scenario is: introduce a deliberate field-name drift in a scratch branch → invariant (b) fails.
**Verification:** harness runs in `npm test`; each corpus test passes against v3. **v2 baseline (from doc review):** the "fails against v2" leg of each corpus test cannot be run after U2/U5/U7/U8 delete the legacy paths in-branch, so capture the baseline first — tag the pre-migration commit (`v2-baseline`) and record v2's failing outputs as committed golden files under `tests/fixtures/v2-baseline/`; each corpus test asserts against the archived v2 failure, proving the test is not tautologically green. Reword verification to reference the tagged baseline, not "legacy code paths where feasible."

### U4. Durable state store (state.db)

**Goal:** One SQLite-backed `StateStore` at `~/.mcp-office365/state.db` (WAL, busy_timeout, schema migrations table) with account-stamped tables for aliases, approval tokens, receipts, delta links — plus boot-time migration/degradation semantics (D7, D15).
**Requirements:** Workstream 5 (state half); D3 (alias table lives here).
**Dependencies:** none (parallel with U1–U3).
**Files:** `src/state/store.ts`, `src/state/schema.ts`, `src/state/migrate.ts` (new); `tests/unit/state/store.test.ts` (new).
**Approach:** better-sqlite3 (already a dependency). Directory/file created with `0700`/`0600` permissions, verified/repaired on boot (D18). Tables: `aliases(token TEXT PK, graph_id TEXT, entity_type TEXT, account_id TEXT, mutable INT, created_at)`, `approval_tokens(token TEXT PK, operation_key TEXT UNIQUE, action TEXT, target_json TEXT, content_hash TEXT, account_id TEXT, expires_at, redeemed_at, receipt_json)`, `meta(key, value)` for schema version + migration marker. (The `delta_links` table ships with U12, not here — D3's day-one justification covers only the alias table; keeping U4 free of Workstream-7-only schema honors the U12 slip rule.) Atomic consume via the `operation_key` unique constraint / guarded `UPDATE … RETURNING` (D8). Corrupt/locked db → in-memory fallback + stderr warning + degraded marker on tokens (D15). Legacy `~/.outlook-mcp` tokens.json copied only if new dir lacks one. 90-day purge of expired approval_tokens/receipts on boot (D8).
**Test scenarios:**
- WAL + busy_timeout set on open (pragma assertions).
- Two store instances on one db file: token consumed in A cannot be consumed in B (atomicity).
- Account stamping: rows written under account X are invisible to queries scoped to account Y.
- Corrupt db file at open → in-memory fallback, warning emitted, all operations still succeed.
- Migration idempotence: running boot migration twice is a no-op (marker respected); tokens.json never overwritten when present.
- Schema-version bump path applies migrations in order.
- Permissions: dir created `0700`, `state.db` `0600` (D18); a pre-existing loose-permission file is repaired to `0600` on boot.
- 90-day purge removes expired approval_tokens/receipts rows on boot; unexpired rows untouched.
**Verification:** unit suite green; manual boot creates the db with expected schema and `0600` permissions.

### U5. Self-resolving durable IDs

**Goal:** Replace `hashStringToNumber` + in-memory `idCache` with hybrid typed tokens (D1: self-encoding for simple entities, alias-backed for composite/mutable), immutable-ID preference (D2), a universal resolver, and Planner fetch-before-update (D6). String|number unions for AppleScript compatibility (D4).
**Requirements:** Workstream 2 — the defining v3 breaking change.
**Dependencies:** U2 (registry handlers to thread through), U4 (alias table).
**Files:** `src/ids/token.ts` (mint/parse/prefix map), `src/ids/resolver.ts` (new); `src/graph/repository.ts` (idCache removal, resolver injection — largest diff in the release); `src/graph/mappers/*.ts` (mint tokens instead of hashing); `src/graph/client/graph-client.ts` (Prefer: IdType=ImmutableId headers; translateExchangeIds batch call); `tests/unit/ids/*.test.ts` (new); `tests/contract/invariants.test.ts` (add round-trip invariant).
**Approach:** Prefix registry — **self-encoding** (D1): `em_` message, `ev_` event, `ct_` contact, `fd_` folder, `dr_` driveItem, `td_` task (token = `<prefix>_<base64url(immutableGraphId)>`, resolved by decode, zero storage, cold-state durable). **Alias-backed** (D1, machine-scoped): `pl_` plan, `pt_` plannerTask, `ch_` chat, `tm_` team, `at_` attachment (`{messageId, attachmentId}`), and composite-keyed entities (channels, chat messages, checklist items) storing their tuple as JSON in the alias row; also any `$search`-minted row where immutable translation failed (`mutable=1`). Resolver: self-encoded token → decode directly; alias token → table lookup → on miss (composite) `ID_UNKNOWN` + re-list hint, or (mutable) re-list/translate → `ID_STALE` (D2). Delete `hashStringToNumber` and all 34 idCache maps. Planner: drop ETag caching, fetch-then-patch with one 412 retry.
**Execution note:** characterization coverage first — snapshot v2 tool outputs for representative tools per entity, then assert v3 outputs differ only in ID fields.
**Test scenarios:**
- Self-encoded cold resolve: mint `em_…` from an immutable ID, discard state.db entirely, resolve on a fresh store instance → decodes to the same Graph ID (proves cold-state durability — the core fix).
- Determinism: self-encoded token for one Graph ID is byte-identical across mints; composite alias token for one canonical key is byte-identical across mints.
- Round-trip: mint→resolve for every prefix type (contract invariant f).
- Composite cold miss: `at_…` token with empty alias table → `ID_UNKNOWN` + re-list hint (documented machine-scoped limitation, not silent mis-resolve).
- Collision (composite): forced same-token/different-key insert → `ID_COLLISION` integrity error, never mis-resolves (D1a; no length extension).
- Mutable-ID lifecycle: `$search`-minted token (`mutable=1`), item moved → resolver re-list/translate; failure → `ID_STALE`; listing with partial translation failure sets `degraded_ids` count (D2).
- Numeric ID on Graph backend → `NUMERIC_ID_UNSUPPORTED`; numeric on AppleScript backend passes through (D4).
- Foreign account token → `ID_FOREIGN_ACCOUNT` (D7).
- Planner update: stale ETag first PATCH → 412 → refetch → second PATCH succeeds; second 412 surfaces typed error.
**Verification:** corpus tests for cache-miss classes pass; no `hashStringToNumber` references remain; e2e list→get→update chains work across a simulated server restart (fresh process) — self-encoded IDs resolve even with the state.db deleted between processes.

### U6. Canonical entity vocabulary + alias coercion + next-action hints

**Goal:** One canonical Zod schema per entity (Email, Event, Contact, Task, Folder, Chat, Message, PlannerTask, DriveItem) with create/get/update views derived from it; input alias coercion (D11); compact `next` hints in responses naming the follow-up tool + param (filtered per D12).
**Requirements:** Workstream 3.
**Dependencies:** U2 (registry), U5 (ID field is the token type).
**Files:** `src/entities/*.ts` (new canonical schemas); `src/registry/input-pipeline.ts` (coercion layer, new); `src/tools/*.ts` (derive views, breaking field renames: `update_event` adopts `start_date`/`end_date`/`description`); shared output shaping replaces `transformXRow` helpers; `tests/unit/entities/*.test.ts`; contract invariants (a)/(b) now enforce this.
**Approach:** Canonical field vocabulary decision: adopt the create/get names (`start_date`, `end_date`, `description`) as canonical — they're the majority and match observed model behavior. Derived views: create = required subset, update = `.partial()` + ID, get output = full shape. Transitional aliases (`start`→`start_date`, `body`→`description` on events; `id`→single-ID param) accepted with coercion. Replace the legacy fake-SQLite-row pipeline (Graph → EmailRow → transform) with Graph → canonical entity for Graph mode; AppleScript keeps row shapes internally but maps to canonical output (frozen backend, minimal touch). Next-action hints: small static map per tool (`get_email` → reply_as_draft/forward_as_draft/…), rendered only when target is in the live surface.
**Test scenarios:**
- Derived-view parity: every field in `get_event` output is accepted by `update_event` (invariant b, per entity).
- Alias coercion: `{id: 'em_x'}` accepted by `get_email`; `{id: …}` on `move_email` rejected with hint naming `email_id` and `folder_id` (D11).
- String/number coercion both directions per D4.
- Legacy field names: `update_event` with `start`/`end`/`body` still works via aliases AND emits the canonical names in output.
- Hint filtering: with `--preset mail`, no hint names a calendar tool.
- Unknown key with no alias → `VALIDATION` envelope listing accepted keys (not bare zod dump).
**Verification:** corpus tests for `id`/`email_id` and `update_event` drift classes pass; e2e get→update round-trips on events and drafts.

### U7. Structured search compilation

**Goal:** Replace raw-KQL `search_emails_advanced` with structured params compiled server-side per D9; align `search_events`/`search_drive_items` params to the same vocabulary where applicable.
**Requirements:** Workstream 4 — kills the #1 failure class (20 errors).
**Dependencies:** U2; benefits from U6 vocabulary.
**Files:** `src/search/compiler.ts` (new); `src/tools/mail.ts` (schema replacement); `src/graph/client/graph-client.ts` (`/search/query` method, new); `tests/unit/search/compiler.test.ts` (new).
**Approach:** Params: `from`, `to`, `subject_contains`, `body_contains` (free-text), `received_after`, `received_before`, `has_attachments`, `is_unread`, `folder_id`, `importance`. Compiler picks the mechanism: property-only → `$filter`; free-text-only → quoted `$search`; mixed → server-built KQL via `POST /search/query` (D9). KQL building is server-owned — correct operator syntax, quoting, and date formatting are code, not model guesswork. The old `query: string` param is removed (breaking, listed in migration matrix).
**Execution note:** spike `/search/query` against a real mailbox via the manual integration workflow before freezing which mixed-mode fields are offered (D9 caveat).
**Test scenarios:**
- Each param alone compiles to the documented mechanism (snapshot the compiled `$filter`/`$search`/KQL strings).
- `from` + `received_after` (two properties) → single `$filter` with correct ISO dates.
- `body_contains` + `received_after` (mixed) → `/search/query` KQL `received>=YYYY-MM-DD` + quoted term.
- Date validation: `received_after: 'yesterday'` → `VALIDATION` envelope suggesting ISO format.
- Empty params → `VALIDATION` requiring at least one criterion.
- Corpus: each of the 6 logged KQL-rejection inputs, expressed as structured params, compiles and executes against fixtures.
**Verification:** corpus KQL class passes; live spike documented in the plan-adjacent notes before merge.

### U8. Resilient Graph transport

**Goal:** One transport chokepoint: retry/backoff honoring `Retry-After` per the D5 exclusion table; download normalization (fixes `download_file`, `download_library_file`, recording downloads); `list_online_meetings` `$top` fix; typed error envelope (D10) mapped at the single place errors are born.
**Requirements:** Workstream 5 (transport half) — kills classes 6, 7, 9, 10.
**Dependencies:** U2 (handlers route through GraphClient); independent of U4–U7.
**Files:** `src/graph/client/transport.ts` (new middleware/wrapper); `src/graph/client/graph-client.ts` (init via middleware chain; `.responseType(ResponseType.ARRAYBUFFER)` on the three binary methods at `:2047`, `:2133`, `:1961`; drop `.top()` on onlineMeetings, slice client-side); `src/utils/errors.ts` (envelope + code vocabulary); `src/graph/client/batch.ts` (per-item failure surfacing; retry failed read sub-requests with the longest Retry-After); `tests/unit/graph/client/transport.test.ts` (new).
**Approach:** Use `Client.initWithMiddleware` with an explicit middleware array `[AuthenticationHandler, custom RetryHandler(D5), TelemetryHandler, HTTPMessageHandler]` whose custom RetryHandler **replaces** the SDK default RetryHandler (the default retries 429/503/504 on all buffered requests including `sendMail`, violating D5 — see D5 implementation note). Only fall back to wrapping the auth-provider fetch if `initWithMiddleware` proves unworkable, and then set the default `RetryHandlerOptions maxRetries=0` so the default retry cannot fire on writes. Map every Graph/network/MSAL failure to the D10 envelope in one function; delete ad-hoc `GRAPH_ERROR:`/`DATABASE_ERROR:` string prefixes.
**Test scenarios:**
- 429 with `Retry-After: 2` → one retry after ≥2s (fake timers), success second attempt.
- 502 on GET → retried; 502 on `sendMail` POST → NOT retried, `retriable: false` + receipt-check suggestion (D5).
- Exhausted retries → `THROTTLED`/`GRAPH_UNAVAILABLE` envelope with `retriable: true`.
- Download: fixture returns a ReadableStream → file on disk byte-equal to fixture content (corpus class 6); ArrayBuffer and Blob response shapes also normalized.
- `list_online_meetings`: request contains no `$top`; limit applied client-side (corpus class 7).
- Batch: mixed-status batch response → failed read items retried; failed write items surfaced individually.
- 401 mid-session routes to token refresh once, then `AUTH_EXPIRED` envelope (no retry loop, no mutex deadlock).
**Verification:** corpus classes 6/7/9/10 pass; no raw `GRAPH_ERROR:` strings remain in src.

### U9. Durable approvals + auth resilience

**Goal:** Approval tokens move to state.db with D8 semantics (24h TTL, idempotent redemption, content-hash sealing, receipts + `list_recent_operations` tool); proactive token refresh under the existing auth mutex; `AUTH_EXPIRED` guidance replaces silent `device_code_expired` mid-session failures.
**Requirements:** Workstream 5 (state half) — kills classes 5, 8.
**Dependencies:** U4 (store), U2 (registry for the new tool).
**Files:** `src/approval/token-manager.ts` (store-backed rewrite, same public surface); `src/approval/hash.ts` (per-target-type seal field maps); `src/approval/receipts.ts` + `list_recent_operations` registry entry (new); `src/graph/auth/device-code-flow.ts` (proactive refresh: refresh when expiry < 10 min at call time; typed failure); `tests/unit/approval/*.test.ts` (extend).
**Approach:** `ApprovalToken.targetId: number` becomes the opaque token string (WS2→WS5 type handoff). Redemption: atomic consume (U4); re-redemption returns stored receipt (D8). Seal fields per action type (send: recipients+subject+body-hash; delete: target id+subject; upload: path+bytes-hash). Receipts table doubles as the send idempotency journal (D5).
**Test scenarios:**
- Token survives simulated restart (new manager on same db) and redeems.
- Double redemption → original receipt returned, operation executed once (spy on handler).
- Seal mismatch (draft body changed between prepare and confirm) → `APPROVAL_TARGET_CHANGED` + fresh preview.
- Expired (>24h) → typed expiry error with re-prepare hint.
- Read-only mode redemption refused (D13).
- Proactive refresh: token expiring in 5 min → refresh fires before the Graph call; refresh failure → `AUTH_EXPIRED` envelope naming the CLI escape hatch (`mcp-office365 auth`).
- `list_recent_operations` returns receipts newest-first with outcome + target summary.
**Verification:** corpus classes 5/8 pass; manual: prepare in one process, confirm in a second process sharing state.db.

### U10. Presets, read-only mode, annotations, CLI flags

**Goal:** `--preset <names>` (default `all`, D12), `--read-only` (D13), annotations already carried by registry entries surfaced in ListTools; CLI parsing extended; `createServer(options)` wired.
**Requirements:** Workstream 6 (non-elicitation half).
**Dependencies:** U2 (registry metadata).
**Files:** `src/cli.ts` (flag parsing beyond the `auth` subcommand); `src/server.ts` (options plumbed to registry filters); `README.md` (flags section); `tests/unit/cli.test.ts`, `tests/integration/server.test.ts` (extend).
**Approach:** Preset names mirror domain modules (`mail`, `calendar`, `contacts`, `tasks`, `teams`, `planner`, `files`, `sharepoint`, `excel`, `people`, `meetings`, `all`). Filtering is a registry query. Read-only = annotation filter + runtime redemption guard (D13). Unknown preset → startup error listing valid names.
**Test scenarios:**
- Default (no flags) exposes the full v2-parity surface (D12 — no shrink on upgrade).
- `--preset mail,calendar` exposes exactly the union; a Teams tool call returns MCP unknown-tool.
- `--read-only` surface contains zero tools with destructive/write annotations and zero `prepare_*`/`confirm_*`; redemption attempt → `READ_ONLY_MODE`.
- Every listed tool carries annotations; read-only tools have `readOnlyHint: true` (contract invariant).
- Hints filtered per active preset (with U6).
**Verification:** e2e ListTools under three flag combinations matches expectations.

### U11. Elicitation confirm mode (flag-gated)

**Goal:** `--confirm elicit`: destructive registry entries execute as single tools that elicit human confirmation in-protocol per D14; default mode unchanged.
**Requirements:** Workstream 6 (user decision: flag-gated in v3.0.0, default candidate for v3.1).
**Dependencies:** U9 (durable tokens as the degradation path), U10 (flag plumbing).
**Files:** `src/approval/elicitation.ts` (new); `src/registry/registry.ts` (destructive entries get an elicit-wrapping execution mode); `tests/unit/approval/elicitation.test.ts` (new).
**Approach:** Capability check at initialize; absent → silent two-phase (D14). Elicit payload renders the same preview the prepare tool builds today. 60s timeout → cancel elicitation, mint durable token, return it with instructions (never both live); decline → revoke pre-minted token; operation-keyed idempotency (D8) backstops races.
**Test scenarios:**
- Client without elicitation capability + `--confirm elicit` → behaves exactly as two-phase (surface + flow parity).
- Accept path: single `send_email` call elicits, user accepts → executes once, receipt written.
- Decline → nothing executed, no live token remains.
- Timeout at 60s → elicitation cancelled, durable token returned; late accept after timeout is ignored (operation executes zero times until token redemption).
- Race: timeout fallback token redeemed AND late elicitation response arrives concurrently → **spy on the Graph handler** (not the receipt count) and assert it fires exactly once, because both paths contend for the same atomic `operation_key` claim (D8/D14) and the loser returns the existing receipt without executing.
**Verification:** e2e with a stub client advertising elicitation capability; concurrent-race test proves single Graph execution via handler spy.

### U12. Delta-sync local mirror + `what_changed` (slippable to v3.1)

**Goal:** Persisted delta links (U4 table) for mail folders, calendar, contacts; incremental local metadata mirror in state.db serving `list_emails`/`check_new_emails`/`search_emails` hot paths with a `freshness` marker and `fresh: true` escape hatch; new `what_changed` tool.
**Requirements:** Workstream 7. **Slip rule:** if this unit endangers the release date, it ships in v3.1 — U1–U11 and U13 do not depend on it.
**Dependencies:** U4, U5 (mirror rows keyed by durable tokens), U8 (transport).
**Files:** `src/mirror/sync.ts`, `src/mirror/read-path.ts`, `what_changed` registry entry (new); `src/state/schema.ts` (mirror tables); `tests/unit/mirror/*.test.ts` (new).
**Approach:** Pull-based delta on demand (no background daemon in v3.0): each hot-path read first advances the delta link (cheap when nothing changed), then serves locally. `what_changed` cold start returns `first_sync: true` with no items (D-flow gap 17). Local search sidesteps `$search` semantics for mirrored fields. Account-stamped rows; sign-out leaves rows inert (foreign-account rule).
**Test scenarios:**
- Delta advance persists across restart (link from run 1 used in run 2).
- Item created in fixture delta feed appears in local list without a full re-fetch.
- `fresh: true` bypasses the mirror and hits Graph.
- `what_changed` first call → `first_sync: true`; second call after fixture changes → exactly the delta items.
- Deleted-item tombstone in delta feed removes the mirror row and invalidates its alias gracefully (`ID_STALE` on later resolve).
- Freshness marker present on mirror-served responses.
**Verification:** corpus-style fixture flows; hot-path latency sanity check.

### U13. Release engineering — v3.0.0

**Goal:** Version 3.0.0 everywhere; CHANGELOG with a breaking-changes migration matrix (old→new field names per entity, numeric→token ID migration, removed `query` param, flags); README rewrite of affected sections; design-doc pairing satisfied (this plan + the ideation doc); tag + PR.
**Requirements:** Release conventions (SemVer, Keep-a-Changelog, `vX.Y.Z` tag; awk-extracted CHANGELOG section format in release.yml).
**Dependencies:** U1–U11 (U12 if it made the cut).
**Files:** `package.json` (3.0.0), `CHANGELOG.md`, `README.md`, `docs/plans/` (this doc updated to `status: done` at completion), `.claude-plugin`/`plugin.json` version fields if applicable.
**Approach:** Migration matrix table: every renamed field (per U6), ID scheme change with the D4 numeric-rejection behavior, removed params, new flags, new error codes (D10). CHANGELOG heading format must match release.yml's awk extraction (`## [3.0.0]`).
**Test scenarios:** Test expectation: none — documentation/versioning unit; correctness gated by release.yml's changelog-extraction step and U2's CI checks.
**Verification:** `npm run build` produces executable `dist/index.js`; release workflow dry-run logic (awk extraction) validated against the new CHANGELOG section locally.

---

## System-Wide Impact

- **Every consumer of numeric IDs breaks** — intended, documented in the migration matrix. Agents adapt instantly (IDs are opaque either way); humans with saved numeric IDs cannot migrate them (lossy hash) — documented.
- **AppleScript mode** keeps numeric rowids (D4); its surface is registry-listed but annotation/preset metadata is best-effort (frozen backend).
- **Coverage thresholds** now apply to code leaving `src/index.ts` — test budget included per unit.
- **npm consumers** of the package export see `createServer` signature gain an optional options param (non-breaking for zero-arg callers).
- **CI**: publish.yml loses its duplicate npm-publish job; release.yml is the single publisher.

## Risk Analysis & Mitigation

| Risk | Mitigation |
|---|---|
| U2 migration regressions across 219 tools | Name-set parity test is the gate; domain-by-domain commits; U3 harness runs per domain |
| `/search/query` semantics differ from `$search` (result ranking, folder scoping) | D9 spike via manual integration workflow before schema freeze; fall back to two-pass `$filter`+client-side filter if unacceptable |
| SDK middleware fights the D5 exclusion table | Fallback: wrap fetch at the auth-provider seam instead |
| ID resolver re-list fallback causes hidden latency storms | Resolver re-lists at most once per (entity-type, call); miss after re-list is a fast typed error |
| state.db contention across concurrent Claude Code windows | WAL + busy_timeout + atomic consume (D7); degradation path tested (U4) |
| Elicitation protocol immaturity | Flag-gated (user decision); zero impact when flag absent |
| Scope: 13 units is a large release | Phase gates A→D; U12 has an explicit slip rule; U3 harness de-risks everything after it |
| state.db at-rest exposure (24h send/delete tokens + MSAL cache) | `0700`/`0600` permissions (D18) + documented trust boundary (redemption trusts the OS user/MSAL identity). SQLCipher/OS-keychain encryption **deferred to a follow-up** — permissions match the existing plaintext MSAL-cache posture; revisit if a stronger bar is required |
| Cross-process token redemption abused by a co-resident malicious process | Accepted risk under the D18 trust boundary (any process as the same OS user can already read the MSAL cache). Optional session-binding mode noted as future work if isolation between agents on one machine becomes a requirement |

## Deferred Implementation Notes

- Exact SDK middleware API shape (`Client.initWithMiddleware` vs fetch wrapper) — decided in U8 against the real SDK version.
- Final preset↔tool membership table — derived mechanically from module boundaries during U2.
- Whether `search_events`/`search_drive_items` adopt the full structured vocabulary in v3.0 or v3.1 — decided in U7 by effort remaining.
- Mirror table column set — decided in U12 from actual hot-path field usage.
