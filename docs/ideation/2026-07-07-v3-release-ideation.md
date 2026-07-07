---
date: 2026-07-07
topic: v3-release
focus: fixes from 4-month usage corpus + new capabilities for v3.0.0
mode: repo-grounded
---

# Ideation: mcp-office365 v3.0.0

## Grounding Context

**Codebase Context:** TypeScript MCP server, 219 tools, dual backend (Graph default / AppleScript legacy). 6,747-line `src/index.ts` holds all schemas + dispatch. Two-phase `prepare_`/`confirm_` approval. Graph ID cache is an in-memory `Map` keyed by lossy 32-bit `hashStringToNumber`, populated only by list/search side-effects; auto-resolvers exist for only teams/plans/chats. Approval tokens in-memory, single-use, 5-min TTL. No retry/backoff. `$batch` infra exists but underused. Delta links in-memory, mail-only.

**Usage corpus (4 months, 621 calls, ~26% error rate):** top failure classes — KQL rejections in `search_emails_advanced` (20), `id` vs `email_id` schema fumbles (11), `update_event` field drift (10), ID-cache misses (~40), `device_code_expired` mid-session, `download_library_file`/`download_file` 100% broken (ReadableStream→Buffer.from), `list_online_meetings` `$top` rejection, expired approval tokens, raw 502s, inconsistent error taxonomy.

**External context (2026):** softeria/ms-365-mcp-server (presets, `--read-only`, scope gating, `--discovery`); merill/lokka (generic wrapper); Microsoft MCP Server for Enterprise (read-only, RAG suggest_queries, honors throttling). Static >200-tool exposure considered a failing pattern (85–98% token reduction via progressive disclosure). MCP tool annotations used by clients; no competitor annotates. No competitor uses MCP resources/prompts/elicitation. Graph underused: delta queries, webhooks, `$batch`, immutable IDs, Retry-After.

## Topic Axes
1. tool-schema-and-id-ergonomics
2. reliability-and-auth
3. context-efficiency-and-tool-surface
4. graph-api-depth
5. architecture-and-release

## Ranked Ideas

### 1. Self-Resolving Durable IDs
**Description:** Replace the lossy numeric `hashStringToNumber` cache with opaque typed tokens (`em_…`, `evt_…`) that encode the Graph ID (base64url), backed by a SQLite alias table and a universal resolver so any tool accepts any ID from any session. Opt into Graph immutable IDs (`Prefer: IdType="ImmutableId"`) on non-`$search` paths.
**Axis:** tool-schema-and-id-ergonomics
**Basis:** direct: "ID-cache misses (~40 total) — 'ID X not found in cache'" — a quarter of all observed errors; all six ideation frames independently converged here.
**Rationale:** IDs stop being session-scoped ephemera; the entire "re-list to refresh the cache" recovery dance disappears structurally.
**Downsides:** The defining v3 breaking change; Planner ETag coupling must be redesigned; migration story needed.
**Confidence:** 95% | **Complexity:** High | **Status:** Unexplored

### 2. Canonical Entity Vocabulary + Alias Coercion + Next-Action Hints
**Description:** One canonical Zod schema per entity; create/get/update views derived from it (field drift becomes unbuildable). Input layer coerces obvious aliases (`id`→`email_id`, string↔number). Responses carry a compact next-action envelope naming the exact follow-up tool and parameter key.
**Axis:** tool-schema-and-id-ergonomics
**Basis:** direct: "get_email schema fumbles (11)… update_event unrecognized_keys (10) — update wants start/end/body but create/get use start_date/end_date/description."
**Rationale:** 21+ errors were the model doing the reasonable thing and being punished; agents learn each entity's shape once.
**Downsides:** Breaking field renames; requires the registry (idea 7) to be economical.
**Confidence:** 95% | **Complexity:** Medium | **Status:** Unexplored

### 3. Structured Search, Compiled Server-Side
**Description:** Delete the raw-KQL surface. `search_emails_advanced` takes structured fields (`from`, `received_after`, `has_attachments`, …) and the server compiles per-field to the right Graph mechanism (`$filter` vs quoted `$search` vs `/search/query`). Invalid searches become unrepresentable.
**Axis:** tool-schema-and-id-ergonomics
**Basis:** direct: "search_emails_advanced KQL rejections (20) — tool description advertises from:/received>= syntax that Graph $search rejects" — the #1 failure class on the #1 tool family.
**Rationale:** The server, not the model, owns Graph's query-language quirks.
**Downsides:** Removes an advertised (broken) capability; per-field compile logic to maintain.
**Confidence:** 90% | **Complexity:** Medium | **Status:** Unexplored

### 4. Reliability Core (Transport + Durable State)
**Description:** Single Graph transport: jittered exponential backoff on 429/5xx/connection-reset honoring `Retry-After`; stream-safe download normalization (fixes the 100%-broken download paths); `$top` endpoint quirks handled centrally; typed error envelope (stable code + `retriable` + machine-actionable `suggestion`). One SQLite state store (`~/.mcp-office365/state.db`): approval tokens (durable, idempotent redemption, content-hash-sealed), delta links, auth state with proactive token refresh. Idempotency receipts for sends.
**Axis:** reliability-and-auth
**Basis:** direct: failure classes 5–10 (downloads broken, no retry, expired tokens, device_code_expired, taxonomy inconsistency); external: Microsoft's own MCP server headline-features throttling compliance.
**Rationale:** One chokepoint retires five failure classes for all 219 tools at once; greenfield per past learnings.
**Downsides:** Retry-on-write idempotency policy needs care (double-send risk); at-rest token storage security review.
**Confidence:** 95% | **Complexity:** Medium-High | **Status:** Unexplored

### 5. MCP-Native Surface (Presets, Annotations, Elicitation)
**Description:** Presets (`--preset mail,calendar`), `--read-only` mode, and MCP tool annotations (`readOnlyHint`/`destructiveHint`/`idempotentHint`) on every tool. Where the client supports elicitation, collapse `prepare_/confirm_` pairs into single tools that elicit human confirmation in-protocol (two-phase kept as fallback), cutting ~60–90 tools from the surface.
**Axis:** context-efficiency-and-tool-surface
**Basis:** external: softeria presets/read-only/discovery; "static >200-tool exposure considered failing pattern"; "none of the competitors annotate"; MCP elicitation whitespace. direct: usage concentrates in ~6 of 27 domains.
**Rationale:** Context is the primary user's scarce resource; annotations + elicitation make the existing safety differentiator spec-native.
**Downsides:** Elicitation client support uneven — needs capability detection; preset defaults are a positioning decision.
**Confidence:** 85% | **Complexity:** Medium | **Status:** Unexplored

### 6. Delta-Sync Local Mirror + `what_changed`
**Description:** Persist Graph delta links (SQLite) for mail/calendar/contacts; maintain an incremental local metadata mirror that serves hot reads (`search_emails`, `list_emails`, `check_new_emails`) with a freshness marker and `fresh: true` escape hatch. New `what_changed` tool: everything new since last session in one call. Mirror doubles as the ID alias store for idea 1.
**Axis:** graph-api-depth
**Basis:** external: "Graph underused: delta queries + change-notification webhooks (CDC pattern)"; direct: reads dominate usage (search_emails 70, get_email 44); delta mechanism already proven in-repo but volatile and mail-only.
**Rationale:** The genuinely-new flagship capability: faster, cheaper, throttle-immune reads, and local search that sidesteps `$search` quirks entirely.
**Downsides:** Biggest new machinery (storage growth, staleness bounds, multi-account); should land after 1/4.
**Confidence:** 75% | **Complexity:** High | **Status:** Unexplored

### 7. Registry-Driven Architecture + Failure-Corpus Contract Harness
**Description:** Each domain module exports tool definitions as data (name, schema, handler, annotations, destructive flag); the server assembles the MCP surface from the registry — the 6,747-line `index.ts` dispatch switch dies. Contract harness iterates the registry asserting invariants (get accepts what list returns; create/update/get share vocabulary; downloads round-trip streams), plus VCR-style record/replay fixtures seeded from the 10 real failure classes.
**Axis:** architecture-and-release
**Basis:** direct: "src/index.ts is 6747 lines"; "70 test files but none caught the failure classes"; reasoned: field drift exists precisely because schema and behavior live in two hand-synced places.
**Rationale:** The enabling refactor — presets, annotations, derived schemas, and invariant tests all become registry queries instead of 219 hand-edits. Sequencing: this lands first.
**Downsides:** Big refactor with no user-visible feature; must not become a rewrite-everything trap.
**Confidence:** 90% | **Complexity:** High | **Status:** Unexplored

## Rejection Summary

| # | Idea | Reason Rejected |
|---|------|-----------------|
| 1 | Five-tool meta facade (discover/read/act/prepare/confirm) | Duplicates idea 5's progressive disclosure in a more radical form; better as a brainstorm variant after presets prove out |
| 2 | Go Graph-only (freeze AppleScript backend into legacy package) | Real option but a product decision, not failure-data-driven; loses Notes with no Graph equivalent — surfaced to user for an explicit call |
| 3 | Standalone HATEOAS response envelope | Merged into idea 2 |
| 4 | Standalone idempotency keys + receipts journal | Merged into idea 4 |
| 5 | Standalone typed error taxonomy | Merged into idea 4 |
| 6 | TOON compact output format | Marginal token win vs idea 5's presets/discovery; experimental format, weak basis |
| 7 | Webhook change notifications (full CDC push) | Requires reachable endpoint — out of scope for a local stdio server; delta polling (idea 6) captures the value |
