---
title: "feat: Remote connector mode (claude.ai custom connector for JP)"
type: feat
status: active
date: 2026-07-11
deepened: 2026-07-11
origin: docs/brainstorms/2026-07-11-remote-connector-mode-requirements.md
---

# feat: Remote connector mode (claude.ai custom connector for JP)

## Summary

Add a `serve` mode exposing the existing tool surface over stateless Streamable HTTP as a claude.ai custom connector. Microsoft Entra ID is the OAuth authorization server directly (pre-registered client, no custom AS); this server is a pure resource server that validates Entra-issued JWTs per request and exchanges them via On-Behalf-Of for per-user delegated Graph tokens. Per-user state rides the existing account-scoped SQLite store on a single replica with a persistent volume. Entitlements extend the preset machinery with exclusion support and per-user config; revocation is purge + deny-list; all write/destructive calls are audit-logged.

---

## Problem Frame

The server is stdio-only, so claude.ai (which JP staff use) cannot reach it; full context, actors, flows, and requirements live in the origin doc (see Sources & References). Plan-specific framing: research resolved the origin's highest-risk unknown — Anthropic officially supports Entra ID as a custom-connector authorization server with a pre-registered client ID, and Microsoft's recommended shape for Entra-backed MCP servers is resource-server + On-Behalf-Of. That collapses the feared "build our own OAuth server" scope into JWT validation plus an MSAL exchange.

---

## Requirements

This plan implements origin requirements R1–R16 (R12 reserved) and honors origin flows F1–F4 plus a new pre-provisioning flow (F0, added by this plan — see U1). Key traceability:

- R1 (Streamable HTTP, stdio unchanged) → U2, U3
- R2 (platform-agnostic core) → U3, U9
- R3/R4 (Entra sign-in, JP members only, guests rejected) → U1, U4
- R5 (per-user tokens, isolated, persistent) → U5
- R6/R7/R14 (entitlements, pinned default surface, Joel parity) → U6
- R8 (two-phase guards, model-mediated caveat) → U6, U8
- R9/R10 (Azure via jp-infrastructure, cost estimate) → U9
- R11 (staged pilot with exit criteria) → U9
- R13 (end-user docs) → U9
- R15 (offboarding/revocation) → U7
- R16 (audit log) → U8

**Origin actors:** A1 (Joel — operator + user), A2 (JP staff member), A3 (claude.ai as MCP client)
**Origin flows:** F1 (onboarding), F2 (document write-back), F3 (entitlement change), F4 (offboarding) — plus plan-added F0 (admin pre-provisioning)
**Origin acceptance examples:** AE1 (token isolation), AE2 (non-JP rejection), AE3 (two-phase delete), AE4 (entitlement exclusion), AE5 (stdio no-regression), AE6 (guest rejection), AE7 (revoke)

---

## Scope Boundaries

- No custom OAuth authorization server (no DCR, no CIMD, no proxy AS). If the client-ID-entry UX ever becomes unacceptable, that is a separate follow-up decision.
- No multi-replica support: single replica is a stated v1 constraint. Redis/Postgres store swap and MCP session rehydration are out until replica count > 1 is actually needed.
- No elicitation-based confirmation in remote mode — claude.ai does not support MCP elicitation; token-mode two-phase is the only confirmation path.
- App-level encryption of stored token blobs is not built at v1; Azure storage-service encryption plus volume/file permissions is the accepted at-rest posture (origin's encryption question, resolved here).
- Origin deferrals stand: entitlement dashboard, multi-platform self-host docs, other tenants, hosted SaaS.

### Deferred to Follow-Up Work

- Terraform/Terramate for the Azure deployment: `jp-infrastructure` repo, driven by U9's deployment-requirements deliverable.
- Multi-platform self-host documentation (Vercel/GCP/AWS): future phase per origin; U3's platform-agnostic design enables it.
- `--preset` doc debt and the missing CLI `shared` preset entry (pre-existing, unrelated to this plan): separate standalone chore PR — not bundled into any unit's diff here.

---

## Context & Research

### Relevant Code and Patterns

- `src/index.ts` — `createServer(options)` factory (per-session friendly); auth mutex + lazy Graph init; `toErrorEnvelope` chokepoint; `main()` CLI routing where `serve` slots in.
- `src/registry/registry.ts` — `matches()` filter: backend → read-only → presets. Presets are include-only; **empty `presets` means always-exposed** (bypass U6 must close). `SurfaceOptions` passed per call — per-user surfaces need no registry rework.
- `src/graph/auth/` — `device-code-flow.ts` module-level `msalInstance`/`cachedConfig`; `token-cache.ts` `FileTokenCachePlugin` (ICachePlugin seam); `account-id.ts` `cachedAccountId` memo, identity = MSAL `homeAccountId` (`<oid>.<tid>` — embeds tenant, aligns with Entra oid+tid claims).
- `src/state/store.ts` — SQLite `aliases` + `approval_tokens` tables already `account_id`-scoped and fail-closed on foreign accounts; atomic single-redeemer consume in one SQL statement (must preserve). `migrate.ts` for DDL.
- `src/approval/token-manager.ts` — `accountId` accepted as a function resolved per op: the designed seam for request-scoped identity.
- `src/graph/client/graph-client.ts` — AuthenticationProvider imports global `getAccessToken`; must become injected per user.
- `tests/contract/invariants.test.ts` — registry-wide invariant assertions; new surface rules land here. `src/registry/**` has a 90% coverage gate.
- CLI subcommand pattern: `parseCliCommand` + injectable `print` + exit codes (`auth` subcommand) — template for `serve`, `revoke`, `audit`.

### Institutional Learnings

- `docs/solutions/architecture-patterns/alias-backed-composite-durable-id-pattern.md` — alias-backed IDs resolve only where the store lives; single replica + persistent volume keeps them working; store rows already per-account.
- `docs/solutions/integration-issues/device-code-auth-undefined-invalid-grant.md` — fail fast on missing config with actionable diagnostics; surface real OAuth endpoint errors (MSAL can swallow 4xx); don't start auth for invalid requests.
- `docs/solutions/best-practices/test-external-api-assumptions-before-building-defenses.md` — U1 spike exists because of this: probe Entra↔claude.ai live before building.
- `docs/solutions/conventions/adversarial-review-as-primary-gate.md` — this feature (auth + routes + persisted state) is squarely in the mandatory adversarial-review zone; schema changes must keep v4.x-written rows readable.
- `docs/solutions/design-patterns/fetch-before-update-for-mutable-etags.md` — approval tokens must never embed volatile preconditions; remote latency makes this more important.

### External References

- claude.ai connector auth requirements: https://claude.com/docs/connectors/building/authentication (+ /building, /troubleshooting — includes the documented Entra ID path and `resource` fix)
- Entra-based MCP servers with OBO (Microsoft): https://techcommunity.microsoft.com/blog/azuredevcommunityblog/using-on-behalf-of-flow-for-entra-based-mcp-servers/4486760
- MCP security best practices (no token passthrough; sessions not for auth): https://modelcontextprotocol.io/specification/2025-11-25/basic/security_best_practices
- MSAL Node multi-user caching (partition by homeAccountId; DistributedCachePlugin): https://learn.microsoft.com/en-us/entra/msal/javascript/node/caching
- Optional claims (`acct` 0=member/1=guest, `idtyp`): https://learn.microsoft.com/en-us/entra/identity-platform/optional-claims-reference
- Restrict app to assigned users: https://learn.microsoft.com/en-us/entra/identity-platform/howto-restrict-your-app-to-a-set-of-users
- SDK 1.29.0 Streamable HTTP + auth middleware (verified from source); session-destruction regression to test: https://github.com/modelcontextprotocol/typescript-sdk/issues/1852
- SQLite on Azure Files SMB is unsafe (WAL/locking); NFS or `nobrl` single-writer mitigations: https://learn.microsoft.com/en-us/answers/questions/318948/ , https://azureossd.github.io/2024/05/16/Preventing-File-Locks-Azure-Container-Apps/

---

## Key Technical Decisions

- **Entra ID as the authorization server directly (Path A)**: claude.ai accepts a pre-registered client ID; Entra lacks DCR but that is irrelevant on this path. Two app registrations (public client + API) avoid the token-for-itself error; admin enters the client ID once when adding the org connector. Rejected: building our own AS/proxy — Microsoft recommends against it for production and it creates the confused-deputy surface.
- **Resource-server-only + On-Behalf-Of**: inbound tokens must have `aud` = our API (token passthrough is spec-forbidden); OBO mints per-user delegated Graph tokens. Consequence: no hand-rolled Graph refresh-token vault — MSAL's per-user cache is the token store (satisfies origin R5 with less machinery than anticipated).
- **Stateless Streamable HTTP**: bearer token verified on every request; identity from validated claims (`oid`+`tid` → existing `homeAccountId` format), never from session IDs (spec MUST). Two-phase state is durable-store-backed, so statelessness doesn't break prepare→confirm across conversations.
- **Member-only enforcement is fail-closed app-side claim validation**: grant only on a positive delegated-member signal — `tid` == JP tenant AND `acct` === 0 present AND a delegated-user signal present (`scp` claim exists / not app-only); reject when `acct` is absent, when a user signal is absent, or when `idtyp === 'app'` / a `roles` claim is present. Absent optional claims are treated as rejection, never as member (the reviewers' P1: `acct`/`idtyp` are Entra *optional* claims, so a reject-on-match test admits a claim-absent token as a member). Single-tenant registration + enterprise-app "assignment required" are control-plane defense in depth. Never authorize on email/UPN (mutable).
- **JWT verification pins RS256** (reject `none`/HS*), validates `nbf`/`iat` with ±5-min skew tolerance, and fetches JWKS and issuer **only from the configured tenant authority — never from the token's own `iss`** (the SSRF vector). JWKS caching refetches on unknown-`kid` with a rate cap and fails closed on JWKS-endpoint outage (prefer `jose` `createRemoteJWKSet`, which supplies cooldown/backoff — this is why it's the U4 default).
- **`serve` refuses to operate on a degraded store**: `StateStore.open` degrades to in-memory on any file error, which in remote mode silently empties the deny-list (revoked users re-admitted) and makes the audit log non-durable. In `serve` mode a degraded store must fail startup / return 503 rather than serve — the deny-list and audit are security controls, not best-effort local state.
- **State stays in the existing SQLite store, single replica, persistent volume**: new tables (MSAL cache blobs keyed by homeAccountId, deny-list, audit log) via `migrate.ts`. The SQLite file IS credential material (MSAL blobs hold refresh tokens) — volume access controls, and any backup/snapshot policy, inherit that sensitivity. Volume must support SQLite locking (Azure Files NFS, or SMB `nobrl` + non-WAL journal — final call in U9's cost/deployment deliverable). Rejected for v1: Redis/Postgres (only needed multi-replica).
- **Secrets are never logged or audited**: Authorization headers, raw/parsed JWTs, delegated Graph tokens, MSAL blobs, and the OBO client secret never reach stdout, the audit table, or error envelopes. This is called out explicitly because U5 deliberately *increases* diagnostic verbosity (raw AADSTS codes) — diagnostics surface only the error code plus non-sensitive claim ids (oid/tid).
- **Ingress abuse controls**: a per-IP rate limit on unauthenticated requests, a per-oid request/OBO quota, and an `express.json` body-size cap — a flood or one misbehaving user must not exhaust the single replica or the shared app registration. Consider an inbound allowlist to Anthropic egress (160.79.104.0/21) at ingress as a day-one option. No permissive CORS on `/mcp` (claude.ai connects server-to-server).
- **Two-phase remote TTL shortened**: the approval-token 24h TTL was sized for a local single-user CLI; remote mode uses a shorter TTL (≈1h via the existing `ttlMs` option) so a prepared destructive action against client data doesn't stay live for a day.
- **Remote mode forces `--confirm token`**: claude.ai has no elicitation; the elicit path stays capability-gated for stdio clients.
- **Revocation = purge + deny-list**: purging tokens alone silently re-onboards (claude.ai still holds a valid Entra token and OBO would re-mint). Deny-list checked in auth middleware; entry removed to re-admit.
- **Default surface excludes `download_*`/`get_*_photo` tools** in the pinned list: they write files to the server's disk, useless and confusing for a claude.ai user (same rationale the read-only flag already applies).
- **SDK pinned to 1.x (`^1.29.0`)**: v2 is beta with renamed packages; 1.x is the supported production line. zod 4.3.6 is compatible (SDK 1.29 supports zod ^3.25 || ^4).

---

## Open Questions

### Resolved During Planning

- OAuth pattern for claude.ai custom connectors: resolved — pre-registered client against Entra directly; PRM (RFC 9728) + 401 handshake required; PKCE S256 always; `resource` (RFC 8707) audience validation against the canonical MCP URL.
- Server-side storage backend: existing SQLite store extended (single replica, persistent volume); encryption-at-rest = Azure SSE + file permissions at v1.
- Azure service choice: recommend Azure Container Apps, single always-on replica; finalized with numbers in U9.
- Per-user preset enforcement: per-request `SurfaceOptions` (already supported by registry API) + exclusion capability + per-user config; no per-user server pools needed in stateless mode.
- Elicitation on claude.ai: not supported — token-mode confirmed as the only remote confirmation path.

### Deferred to Implementation

- Entitlement config file format and loading/reload mechanics (per-request read vs mtime cache): settle in U6 against real config shape.
- JWKS/JWT library choice (`jose` is bundled with the SDK; direct dependency vs SDK re-export): settle in U4.
- Exact DDL for new tables and the ICacheClient adapter surface: settle in U5 against MSAL's DistributedCachePlugin contract.
- Whether Application ID URI can be the full https MCP URL in the JP tenant (tenant policy varies): U1 spike answers empirically.

---

## High-Level Technical Design

> *This illustrates the intended approach and is directional guidance for review, not implementation specification. The implementing agent should treat it as context, not code to reproduce.*

```mermaid
sequenceDiagram
    participant C as claude.ai (MCP client)
    participant S as mcp-office365 serve (resource server)
    participant E as Entra ID (JP tenant)
    participant G as Microsoft Graph

    Note over C,S: First contact (no token)
    C->>S: POST /mcp (no Authorization)
    S-->>C: 401 + WWW-Authenticate resource_metadata=…/.well-known/oauth-protected-resource
    C->>S: GET PRM document
    C->>E: OIDC discovery → authorize (PKCE S256, resource=MCP URL)
    Note over E: User signs in with JP account
    E-->>C: code → token (aud = our API)

    Note over C,S: Every request thereafter
    C->>S: POST /mcp + Bearer token
    S->>S: Validate JWT (iss, aud, exp, tid=JP, acct=0, not app-only, not deny-listed)
    S->>E: OBO exchange (per-user, MSAL cache partition = homeAccountId)
    E-->>S: Delegated Graph token for this user
    S->>G: Graph call as the user
    G-->>S: Result
    S-->>C: Tool result (surface filtered by user entitlements; writes audit-logged)
```

---

## Implementation Units

### U1. Entra ↔ claude.ai handshake spike + provisioning runbook (F0)

**Goal:** Empirically validate the load-bearing unknowns before building, and produce the one-time JP-tenant provisioning runbook that corrects origin F1's "no admin action needed" claim.

**Requirements:** R3, R4 (origin Dependencies); flow-gap fixes F0/F1

**Dependencies:** None (requires Joel's access to the JP tenant / a test tenant)

**Files:**
- Create: `docs/plans/2026-07-11-remote-spike-findings.md` (spike results)
- Create: `docs/remote/provisioning.md` (admin runbook: app registrations, consent, claims, assignment)

**Approach:**
- Register the two-app shape in a test or JP tenant: API app (expose `access_as_user`; Application ID URIs = `api://{guid}` AND the exact MCP URL; `requestedAccessTokenVersion: 2`; optional claims `acct`, `idtyp`) + public client app (redirect `https://claude.ai/api/mcp/auth_callback`, pre-authorized for the API scope).
- Stand up a throwaway Streamable HTTP endpoint (can be a minimal script, not product code) returning 401 + PRM; add it as a claude.ai custom connector with the client ID; verify the full auth handshake completes against Entra (specifically: missing `code_challenge_methods_supported` advertisement tolerated; Application ID URI accepts the https URL form in this tenant; `resource` validation passes).
- Verify one OBO exchange mints a Graph token and a `/me` call succeeds; verify a B2B guest and a foreign-tenant account are distinguishable via claims.
- Verify session longevity through Azure ingress (SDK keepalive regression #1852) if the throwaway endpoint is hosted; otherwise defer that check to U9's pilot smoke test.
- Document tenant-wide admin-consent steps and Conditional Access implications (test one sign-in from claude.ai under JP's CA policies).

**Execution note:** Probe-first by design — this unit exists to test external assumptions before defenses are built (per `docs/solutions/best-practices/`).

**Test scenarios:**
- Happy path: claude.ai connector add → Entra sign-in (member) → handshake completes → authenticated MCP request reaches the throwaway endpoint.
- Error path: guest account sign-in → token claims show `acct=1` / foreign home tenant (rejection evidence recorded for U4).
- Error path: sign-in under JP Conditional Access → document outcome (pass or the AADSTS error users would see).

**Verification:**
- Spike findings doc answers the four residual unknowns (PKCE advertisement tolerance, App ID URI format, CA behavior, OBO viability) with observed evidence; provisioning runbook is executable by a JP admin without Joel improvising.

---

### U2. MCP SDK upgrade with stdio regression gate

**Goal:** Move `@modelcontextprotocol/sdk` from lockfile 1.26.0 to `^1.29.0` (declared `^1.0.0` today) with zero stdio behavior change, unlocking current Streamable HTTP + auth middleware.

**Requirements:** R1 (AE5)

**Dependencies:** None

**Files:**
- Modify: `package.json`, `package-lock.json`
- Test: existing suites (`tests/unit`, `tests/contract`, `tests/e2e`)

**Approach:**
- Bump the declared range to `^1.29.0` so the floor is explicit rather than lockfile luck; run the full suite as the regression gate.
- Watch zod interop (SDK 1.29 imports `zod/v4`; repo pins zod 4.3.6 — compatible, but schema-conversion call sites are the blast radius).

**Test scenarios:**
- Integration: full existing vitest suite green on Node 20/22/24 (CI matrix) — this IS the AE5 gate at build time.

**Verification:**
- `npx @jbctechsolutions/mcp-office365` with no new flags behaves exactly as today (AE5); lockfile shows a single SDK version.

---

### U3. `serve` subcommand + stateless Streamable HTTP transport

**Goal:** The server runs as an HTTP service exposing `/mcp` (stateless Streamable HTTP) alongside unchanged stdio, with platform-agnostic config (env/flags only) and no auth yet (binds localhost by default until U4 lands).

**Requirements:** R1, R2

**Dependencies:** U2

**Files:**
- Create: `src/remote/http-server.ts` (Express app assembly, `/mcp` + `/healthz` routes)
- Modify: `src/cli.ts` (parse `serve` subcommand + `--port`/`--host` flags), `src/index.ts` (hoist `StateStore`/registry to process scope; split backend construction from the device-code flow so a per-request Server can be built with an injected identity)
- Test: `tests/unit/remote/http-server.test.ts`

**Approach:**
- **Refactor prerequisite (feasibility correction):** `createServer()` today opens a `StateStore` and hardwires `initializeGraphBackend` → `getAccessToken` (device-code) internally, so it is NOT cheap to call per request (re-runs migration/purge, leaks a SQLite/WAL handle per request) and cannot reach the OBO path. Hoist the `StateStore` (opened once with `{ dir }` pointing at the persistent volume) and the static registry to process scope, and separate backend construction from the device-code flow so a per-request Server can be assembled with an injected token provider (device-code for stdio, OBO for remote — the remote provider lands in U5).
- Stateless mode: `sessionIdGenerator: undefined`, new per-request Server over the shared store + registry. No session map, no EventStore.
- Bind fail-closed: `serve` refuses a non-localhost bind unless auth config (U4) is present; `express.json` carries a body-size cap; no permissive CORS on `/mcp`.
- Reuse the CLI subcommand pattern from `auth` (injectable `print`, exit codes). Server flags (`--preset`, `--read-only`) remain valid for `serve` as the process-wide outer bound; per-user surfaces layer inside it in U6.
- `serve` forces `confirmMode: 'token'` (claude.ai has no elicitation).
- Keep everything cloud-agnostic: configuration via env/flags, logs to stdout — Azure specifics stay in jp-infrastructure (R2).

**Patterns to follow:** `parseServerOptions` / `parseCliCommand` in `src/cli.ts`; error envelope chokepoint in `src/index.ts`.

**Test scenarios:**
- Happy path: initialize + tools/list + tools/call round-trip over HTTP against a mocked backend.
- Edge case: request with `Mcp-Session-Id` header still works in stateless mode (header ignored, not erroring).
- Error path: malformed JSON-RPC body → protocol-level error response, process stays up.
- Edge case: over-limit request body → rejected by the body cap, process stays up.
- Covers AE3: `prepare_delete`/`confirm_delete` driven over the stateless HTTP transport with `confirmMode: 'token'` forced still requires the `confirm_*` call before deletion executes (re-validates the two-phase gate survives the transport change).
- Integration: stdio entry path untouched when `serve` is not passed (AE5 guard at unit level).

**Verification:**
- `mcp-office365 serve --port N` serves MCP over HTTP locally; `mcp-office365` alone still starts stdio; `/healthz` returns OK without auth.

---

### U4. Resource-server auth layer (PRM, 401 handshake, JWT validation, deny-list)

**Goal:** Every `/mcp` request is authenticated: RFC 9728 metadata + 401 challenge for discovery, then per-request Entra JWT validation enforcing JP-tenant members only.

**Requirements:** R3, R4 (AE2, AE6); deny-list check for R15

**Dependencies:** U1 (validated registration shape), U3

**Files:**
- Create: `src/remote/auth/verify.ts` (JWKS fetch/cache + claim validation), `src/remote/auth/metadata.ts` (PRM document + WWW-Authenticate), `src/remote/config.ts` (env: tenant ID, API audience/client IDs, public MCP URL — fail-fast diagnostics)
- Modify: `src/remote/http-server.ts` (mount middleware), `src/utils/errors.ts` (auth error codes)
- Test: `tests/unit/remote/auth/verify.test.ts`, `tests/unit/remote/auth/metadata.test.ts`

**Approach:**
- Serve `/.well-known/oauth-protected-resource` (and the path-suffixed variant claude.ai probes); unauthenticated requests get 401 + `WWW-Authenticate: Bearer resource_metadata="…"` — mandatory for claude.ai discovery. PRM values (`resource`, `authorization_servers`) come only from `OUTLOOK_MCP_PUBLIC_URL` config and the pinned tenant issuer — never computed from `Host`/`X-Forwarded-*` (Host-header injection would poison discovery).
- Validate, fail-closed (membership granted only on a positive delegated-member signal — see Key Technical Decisions): signature via tenant JWKS pinned to **RS256**, `iss` == configured v2 issuer (reject v1 `sts.windows.net`), `aud` (accept `api://{guid}` and bare-GUID forms; canonical RFC 8707 resource form, reject the public-client app ID and Graph as audiences), `exp` + `nbf`/`iat` with ±5-min skew, `tid` == JP tenant, `acct` === 0 **present**, `scp` contains `access_as_user`, reject app-only/`roles` tokens, reject when expected claims are absent, reject deny-listed `oid`s.
- JWKS: fetch only from the configured tenant authority (never the token's `iss`); refetch on unknown-`kid` with a rate cap; fail closed on JWKS outage. `jose` `createRemoteJWKSet` supplies cooldown/backoff.
- Ingress abuse controls (per Key Technical Decisions): per-IP unauthenticated rate limit + per-oid quota; `/healthz` leaks no version/config.
- Derive identity `oid.tid` (== MSAL homeAccountId format) and attach to request context for downstream units. Deny-list is read per-request from the store (never process-cached — a later in-memory cache would reintroduce a staleness window).
- Log authorization DENIALS as security events (guest/foreign-tenant/app-only rejections, deny-list hits, later OBO consent/interaction failures) with oid-where-present, reason code, timestamp — the only detection surface for probing during the pilot; the R11 exit review reads them.
- Config loading mirrors the v4.2.0 fail-fast pattern: missing env → startup error with setup guidance pointing at `docs/remote/provisioning.md`.
- Rejection responses carry a machine-readable body (guests hit this AFTER a successful Microsoft sign-in — the error text is the only UX; R13 documents the symptom). Bodies and logs never include token material.

**Patterns to follow:** typed `OutlookMcpError` subclasses + envelope; `loadGraphConfig` fail-fast diagnostics.

**Test scenarios (mocked JWKS/tokens):**
- Happy path: valid member token (`acct=0`, `scp` has `access_as_user`) → request passes, identity attached.
- Covers AE2: foreign-tenant token (`tid` mismatch) → 401/403, nothing stored.
- Covers AE6: JP-directory guest token (`acct=1`) → rejected with descriptive body, nothing stored.
- Error path: token with `acct`/`idtyp` entirely absent → rejected (fail-closed, not admitted as member); app-only/`roles` token → rejected; deny-listed oid → rejected.
- Error path: expired token → 401 + WWW-Authenticate (claude.ai refresh trigger); not-yet-valid (`nbf` future beyond skew) → rejected; within-skew clock drift → accepted.
- Edge case: `aud` bare-GUID and `api://` both accepted; `aud` = public-client app ID rejected (token-for-itself); `aud` = Graph rejected (passthrough); `alg: none`/HS* rejected.
- Edge case: rotated signing key (unknown `kid`) → one JWKS refetch → passes; unknown-`kid` flood → refetch rate-capped; JWKS endpoint down → fail closed.
- Happy path: unauthenticated GET of PRM document succeeds; PRM `resource` matches configured public URL and ignores a spoofed `Host` header.
- Security-event: each rejection above writes a denial log row with reason code and no token material.

**Verification:**
- With U1's registrations, adding the connector against a deployed/tunneled instance completes discovery + sign-in; a guest account's request is rejected with the documented error.

---

### U5. Per-user Graph access: partitioned MSAL cache + On-Behalf-Of

**Goal:** Replace process-global auth identity with request-scoped identity: per-user Graph tokens via OBO, cached per user in the durable store, with the stdio device-code path fully preserved.

**Requirements:** R5 (AE1)

**Dependencies:** U4

**Files:**
- Create: `src/remote/auth/obo.ts` (ConfidentialClientApplication + OBO acquire), `src/remote/auth/user-cache.ts` (ICacheClient/partition manager over the state store; one MSAL blob per homeAccountId)
- Modify: `src/state/store.ts` + `src/state/migrate.ts` (msal_cache table, account-scoped), `src/graph/client/graph-client.ts` (inject token provider instead of importing global `getAccessToken`), `src/graph/repository.ts` (thread a token-provider param through the `GraphRepository` constructor and the `createGraphRepository` factory — `GraphClient` is instantiated here, not at the index.ts call site), `src/index.ts` (`buildToolContext`/toolset construction takes request identity; `StateStore.open({ dir })` from env; approval-manager accountId thunk becomes request-scoped in remote), `src/graph/auth/account-id.ts` (allow request-scoped override; memo stays for stdio)
- Test: `tests/unit/remote/auth/obo.test.ts`, `tests/unit/remote/auth/user-cache.test.ts`, `tests/unit/state/store.test.ts` (new table)

**Approach:**
- MSAL `DistributedCachePlugin` pattern: partition key = homeAccountId; CCA (or partition manager) instantiated per request so each cache access loads only that user's blob. API app's confidential credential (required for OBO) comes from env — never in the repo; prefer a certificate over a shared secret (record the decision either way).
- Map OBO failures explicitly (flow-gap fix): `invalid_grant`/`interaction_required` → 401 + WWW-Authenticate (claude.ai re-auths); consent-missing (AADSTS65001) → descriptive tool error naming the admin-consent runbook; transient → retriable envelope. Surface the raw AADSTS code in diagnostics — but scrub token material first: OBO assertions, Graph tokens, and MSAL error objects are never logged (U5 raises verbosity, so redaction is explicit here).
- Revocation-race ordering (with U7): the deny-list check runs before OBO, but a request already past middleware can write a fresh cache blob after a purge — so U7 inserts the deny-list row *before* purging, and the ICacheClient re-checks the deny-list at cache-write time (or the one-straddling-request window is documented as accepted).
- Graph client/toolsets constructed per request with the user's token provider; stdio path keeps the existing lazy singleton flow (`serve` never calls device-code; stdio never calls OBO).
- Migration discipline: new tables additive-only; rows written by v4.x local mode must remain readable (upgrade-boundary learning).

**Execution note:** Test-first on the claim/identity plumbing — AE1 (cross-user isolation) is the highest-stakes invariant in the feature.

**Test scenarios:**
- Covers AE1: two identities issue interleaved requests → each Graph call uses its own token; cache partitions never cross (assert partition-key derivation and store isolation).
- Happy path: first request per user runs OBO; second within TTL hits cache silently.
- Error path: OBO `invalid_grant` → 401 challenge; consent error → descriptive tool error; store unavailable → fail closed (no cross-user fallback to a shared cache).
- Edge case: same user concurrent requests (two conversations) → both succeed; blob writes don't corrupt (last-writer-wins on the serialized cache is acceptable, assert no exception).
- Integration: stdio mode end-to-end still authenticates via device-code path untouched (AE5).

**Verification:**
- A signed-in pilot user's mail/files calls return their own data across server restarts without re-consenting; a second user cannot observe the first's state through any code path.

---

### U6. Per-user entitlements: exclusions, pinned default surface, per-user config

**Goal:** Tool surface is computed per user from config: a pinned, versioned default tool list; per-user overrides (including Joel's full-parity entry); exclusion capability closing the include-only and empty-presets gaps; re-checked at dispatch and confirm time.

**Requirements:** R6, R7, R14 (AE4); confirm-time re-check supports R8/F3/F4

**Dependencies:** U4 (identity), U3

**Files:**
- Create: `src/remote/entitlements.ts` (config load/parse/reload + per-user SurfaceOptions resolution)
- Modify: `src/registry/registry.ts` + `src/registry/types.ts` (exclusion support in `matches()`; explicit handling for empty-presets tools), `src/remote/http-server.ts` (resolve entitlements per request)
- Test: `tests/unit/remote/entitlements.test.ts`, `tests/unit/registry/registry.test.ts`, `tests/contract/invariants.test.ts` (new invariants)

**Approach:**
- Entitlement config is a file (mounted from jp-infrastructure config, path via env): pinned default tool list (explicit names, versioned — a server upgrade cannot widen anyone's surface without a config change), per-user entries keyed by oid with preset/tool grants or restrictions; Joel's entry grants the full surface including shared-mailbox/mail-rules (R14). `download_*`/`get_*_photo` excluded from the default list (server-local writes are meaningless to remote users).
- Exclusion semantics decided at the registry: surface = (explicit tool list ∪ preset expansion) − exclusions; empty-presets tools are only exposed when explicitly listed or when no surface constraint is active (stdio default) — the always-exposed bypass must not leak into remote mode.
- Reload on next tool invocation (F3's propagation promise): read-through with mtime check — no restart, no re-auth. The config file is the privilege-granting boundary (Joel's full-parity entry) — mount it read-only and isolated from the SQLite state volume (see U9) so a store-path-write bug can't become privilege escalation.
- Dispatch re-resolves entitlements per call; `confirm_*` handlers re-check entitlement + deny-list at execution time (prepare-time checks are not trusted — flow-gap fix).

**Note:** The pre-existing `--preset` doc debt and the missing CLI `shared` preset entry are unrelated to this plan and are NOT bundled here — see Deferred to Follow-Up Work.

**Patterns to follow:** `SurfaceOptions` flow through `listTools`/`dispatch`; contract invariants file for registry-wide rules.

**Test scenarios:**
- Covers AE4: user config excluding Planner → Planner tools absent from tools/list (not erroring on call).
- Happy path: unconfigured JP user gets exactly the pinned default list; Joel's entry exposes shared-mailbox + mail-rules tools.
- Edge case: empty-presets (always-exposed) tool NOT in the pinned list → hidden in remote mode, still exposed in stdio.
- Edge case: config edit → next invocation reflects new surface without restart.
- Error path: malformed config → fail closed to the pinned default (or refuse startup if the default itself is invalid), with diagnostics.
- Error path: entitlement narrowed between prepare and confirm → confirm rejected with model-readable guidance.
- Contract invariant: every tool name in the shipped default list exists in the registry (catches upgrade drift both directions).

**Verification:**
- tools/list for a test user matches the pinned list exactly; a config-only change flips a user's surface on their next call; registry coverage gate (90%) still met.

---

### U7. Offboarding and revocation (purge + deny-list)

**Goal:** A departed or revoked user is actually out: their stored tokens are purged AND their oid is deny-listed so a still-valid Entra token (or re-add) cannot silently re-onboard them.

**Requirements:** R15 (AE7), F4

**Dependencies:** U5 (token store), U4 (middleware reads deny-list)

**Files:**
- Create: `src/remote/revocation.ts` (purge + deny-list operations), CLI wiring for `revoke <user>` / `revoke --list` / `revoke --readmit <user>` in `src/cli.ts`
- Modify: `src/state/store.ts` + `migrate.ts` (deny-list table), `docs/remote/provisioning.md` (offboarding runbook section)
- Test: `tests/unit/remote/revocation.test.ts`

**Approach:**
- `revoke` **inserts the deny-list row first, then** purges the user's MSAL cache blob + approval tokens; middleware (U4) rejects deny-listed identities before OBO. Ordering matters: a request already past middleware could otherwise write a fresh cache blob after a purge — inserting the deny-list first (and re-checking it at cache-write time, per U5) closes that straddle. Re-admittance is an explicit removal.
- Entra account disablement is the independent backstop (OBO fails) — documented, not implemented.
- Runbook documents that connector removal in claude.ai does NOT clear server state (origin R15 wording).

**Test scenarios:**
- Covers AE7: revoke → tokens gone; next request from a still-valid Entra bearer → rejected (deny-list); after readmit + fresh sign-in → works.
- Edge case: revoking a user with no stored tokens succeeds idempotently.
- Error path: revoked user's in-flight confirm token → consume rejected (identity check), not replayable.
- Race: revoke fires while a request for that user is mid-flight → no cache blob for that user survives after revocation completes (deny-list-first ordering holds).

**Verification:**
- End-to-end on a test user: revoke while their claude.ai conversation is open → their next tool call fails with the documented error; nothing of theirs remains in the store except the deny-list row and audit history.

---

### U8. Audit log for write/destructive operations

**Goal:** Every write/destructive tool invocation in remote mode is durably logged (who, what tool, what target, when, prepare/confirm outcome), fail-closed for destructive confirms, readable for the pilot-exit review.

**Requirements:** R16; supports R11 exit criteria

**Dependencies:** U5 (identity), U6 (dispatch path)

**Files:**
- Create: `src/remote/audit.ts` (writer + query), CLI `audit` subcommand in `src/cli.ts` (filter by user/time, human-readable output)
- Modify: `src/state/store.ts` + `migrate.ts` (audit table), dispatch path in `src/remote/http-server.ts` or `src/index.ts` chokepoint (log at the CallTool boundary — single place, no per-tool changes)
- Test: `tests/unit/remote/audit.test.ts`

**Approach:**
- Log at the dispatch chokepoint using tool annotations already present (`destructive`, `readOnlyHint`): record non-read tools only (identity oid, tool name, target resource extracted from the durable-ID param when present, timestamp, outcome incl. prepare/confirm linkage and errors).
- Fail-closed rule: if the audit insert fails, **all `confirm_*` operations abort** with a retriable error — `confirm_send_email` sends mail from a client tenant, which is exactly what R16 exists to make non-repudiable, so it must not proceed unaudited. Fail-open (proceed with a logged warning) is reserved for non-two-phase writes only.
- Never write token material to the audit table (identity is oid/tid only) — consistent with the redaction rule in Key Technical Decisions.
- Retention: keep everything through the pilot (no pruning at v1); the R11 pilot-exit review reads `audit` output, including the U4 authorization-denial security events.

**Test scenarios:**
- Happy path: send_email via prepare/confirm → two linked audit rows with identity + outcome.
- Happy path: read-only tool → no audit row.
- Error path: audit table unavailable → confirm_delete AND confirm_send_email abort (fail-closed); a non-two-phase write proceeds with warning.
- Edge case: audit rows written by this version remain readable after a schema-additive upgrade (upgrade-boundary invariant).

**Verification:**
- After a scripted session of mixed reads/writes, `mcp-office365 audit --user X` reconstructs exactly the writes with correct attribution.

---

### U9. Documentation, deployment hand-off, and cost estimate

**Goal:** Everything the pilot needs that isn't server code: JP end-user setup doc, deployment-requirements spec for jp-infrastructure, the R10 cost estimate, and the pilot runbook with exit criteria.

**Requirements:** R9, R10, R11, R13; R2 (deployment boundary)

**Dependencies:** U1 (provisioning facts); content finalized after U3–U8 stabilize

**Files:**
- Create: `docs/remote/user-guide.md` (R13: add connector, sign in, what errors mean — guest rejection, reconnect on expiry, CA symptoms, keep per-tool approval prompts ON for confirm tools), `docs/remote/deployment.md` (requirements hand-off: container image, env contract, volume requirements incl. SQLite locking constraint — NFS or nobrl+non-WAL, single replica, health endpoint, log expectations, Anthropic egress allowlist 160.79.104.0/21 if WAF'd, IPv4 A-record requirement; SQLite file + volume snapshots are credential material; entitlement config mounted read-only and isolated from the state volume; TLS-terminating ingress assumed), `docs/remote/pilot-runbook.md` (R11 exit criteria + observation checklist: throttling under shared app registration, 300s-timeout-prone tools, keepalive regression symptoms, auth-failure-rate / 401-spike watch, security-denial review, audit review step)
- Also in provisioning.md (U1): OBO confidential-credential lifecycle — expiry (Entra secrets ≤24mo, often 6), the create-new-then-swap-then-delete rotation procedure (MSAL CCA takes one credential), a calendar expiry reminder, and the AADSTS7000222 total-outage symptom
- Modify: `README.md` (remote mode section; also close the `--preset` doc debt if cheap)
- Test: none (docs) — Test expectation: none — documentation-only unit.

**Approach:**
- Cost estimate (R10, delivered in `docs/remote/deployment.md`): Azure Container Apps single always-on replica (~0.25–0.5 vCPU) + premium file share (NFS) + log analytics; produce pilot (~3 users) and full-JP projections with the sensitivity note that the estimate precedes real usage data (treat as provisional for the R11 gate). Terraform itself is jp-infrastructure follow-up work consuming this doc.
- User guide is written for MCP-unfamiliar staff (training goal) and covers the failure symptoms flow analysis identified (guest rejection surface, token-expiry reconnect, CA blocks, "service busy" under throttling).

**Verification:**
- A JP pilot user onboards end-to-end using only the user guide (origin success criterion); jp-infrastructure work can start from `deployment.md` without asking this repo questions; a written cost estimate with a number Joel can set a ceiling against exists before any deployment.

---

## System-Wide Impact

- **Interaction graph:** New HTTP entry path converges on the existing CallTool chokepoint — error envelope, elicitation gating, and approval flows all still route through `src/index.ts` handlers; audit logging attaches at that single point.
- **Error propagation:** New failure classes (401 challenge, OBO failures, entitlement-denied, deny-listed, audit-fail-closed) become typed `OutlookMcpError` codes so the envelope contract (`code`, `retriable`, `suggestion`) holds for remote clients; model-readable suggestions matter more remotely (Claude self-recovers from expired approval tokens, narrowed entitlements). A single redaction rule governs all of these plus U5 diagnostics: no token material in envelopes, logs, or audit rows.
- **Audit chokepoint location:** The CallTool chokepoint lives in `createServer`'s handler in `src/index.ts` (shared with stdio), not `src/remote/http-server.ts`. U8's audit sink threads through `buildToolContext`/`ToolContext` and is null for stdio — remote-only auditing keys off the injected sink, not a separate code path.
- **State lifecycle risks:** New tables (msal_cache, deny_list, audit) join account-scoped SQLite; single-writer single-replica assumption is load-bearing — documented in deployment.md; approval-token atomic-consume semantics unchanged.
- **API surface parity:** stdio remains the npm default and must not regress (AE5 gate in U2/U3/U5); read-only and preset flags mean the same thing in both modes; remote adds per-user layering inside the process-wide bounds.
- **Integration coverage:** Cross-layer scenarios unit tests won't prove — real Entra handshake (U1 spike + pilot), claude.ai discovery quirks, keepalive behavior behind Azure ingress (pilot smoke test), Graph throttling under one shared app registration (pilot observation).
- **Unchanged invariants:** Tool definitions, zod schemas, durable-ID system, two-phase token semantics, Graph repository behavior — untouched. The registry gains exclusion support but existing include-only semantics for stdio flags are preserved.

---

## Risk Analysis & Mitigation

| Risk | Likelihood | Impact | Mitigation |
|------|-----------|--------|------------|
| claude.ai rejects Entra metadata quirks (missing PKCE advertisement) | Low | High | U1 spike proves the handshake before any product code depends on it |
| Application ID URI can't be the https MCP URL in JP tenant | Low | Med | U1 spike; fallback is api://GUID + Anthropic's documented `resource` guidance |
| SDK keepalive regression (#1852) kills sessions behind Azure ingress | Med | Med | Stateless mode reduces exposure; pilot smoke test; pin SDK version consciously |
| Conditional Access blocks claude.ai browser auth for JP staff | Med | High | U1 tests a real CA sign-in; provisioning runbook documents required CA exemptions before pilot |
| Graph throttling shared across all users (one app registration) | Med | Med | Existing retry middleware + honest "service busy" errors; pilot-runbook observation item; revisit at broader rollout |
| SQLite on wrong volume type corrupts under load | Low | High | deployment.md hard requirement (NFS or nobrl+non-WAL); single replica |
| Store degrades to in-memory (locked/bad volume) → empty deny-list + non-durable audit (fail-open security controls) | Med | High | `serve` refuses to start / returns 503 on `store.degraded` — never serves with an empty deny-list (U4) |
| OBO confidential credential expires → total outage (AADSTS7000222) for all users at once | Med | High | Rotation runbook + calendar expiry reminder (U1/U9); prefer certificate; pilot-runbook monitors for the outage shape |
| Unauthenticated flood or single-user OBO amplification exhausts the 0.25-vCPU replica / shared app registration | Med | Med | Per-IP + per-oid rate limits, body-size cap, optional ingress IP allowlist (U4/deployment.md); pilot observation item |
| Prompt-injected content (e.g. a read email) drives prepare→confirm end-to-end — both are model-mediated | Med | Med | claude.ai per-tool approval prompts (client-side, user-disableable — server can't verify), curated default surface, audit trail; R11 exit review evaluates it consciously |
| 300s claude.ai tool timeout on large uploads/downloads | Med | Low | download tools excluded from default surface; upload size guidance in user guide; observe in pilot |
| Tool count still too large for good claude.ai UX | Med | Med | Pinned default list is the curation lever; pilot exit review tunes it — config change, not code |

---

## Phased Delivery

- **Phase 1 — Validate (U1):** spike + provisioning runbook. Nothing else starts until the handshake is proven.
- **Phase 2 — Core server (U2 → U3 → U4 → U5):** SDK bump, transport, auth, per-user Graph. After U5 a single pilot user (Joel, signing in with his **JP-tenant** account per R4 — not jbc.dev) can use it end-to-end via a tunnel. His oid is recorded by U1's provisioning runbook for U6's entitlement config, and his account is added to the enterprise-app assignment list.
- **Phase 3 — Governance (U6 → U7 → U8):** entitlements, revocation, audit. Gate for inviting non-Joel pilot users.
- **Phase 4 — Ship to pilot (U9):** docs, deployment hand-off, cost estimate; jp-infrastructure Terraform lands; pilot begins under the R11 exit criteria.

---

## Documentation / Operational Notes

- Adversarial code review is the merge gate for U4–U7 (auth, identity, persisted state — per `docs/solutions/conventions/adversarial-review-as-primary-gate.md`).
- New env contract (documented in deployment.md): tenant ID, client/API app IDs, API client secret (OBO), public MCP URL, entitlement config path, state dir. Secret handling via the deployment platform, never in repo or image.
- Release posture unchanged: npm package keeps working for stdio users; container image publication (if added) follows the OIDC/no-long-lived-secret pattern used for npm.

---

## Sources & References

- **Origin document:** [docs/brainstorms/2026-07-11-remote-connector-mode-requirements.md](../brainstorms/2026-07-11-remote-connector-mode-requirements.md)
- Related code: `src/index.ts`, `src/registry/registry.ts`, `src/graph/auth/`, `src/state/store.ts`, `src/approval/token-manager.ts`
- Related learnings: `docs/solutions/` (five docs cited in Context & Research)
- External docs: claude.ai connector auth, MS OBO-for-MCP guidance, MCP security best practices, MSAL Node caching (URLs in Context & Research)
