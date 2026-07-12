# Remote Connector — Continuation Handoff (for Cursor / next agent)

**Status as of 2026-07-12:** 7 of 9 units merged to `main`. The entire
security-critical core (auth, per-user Graph, entitlements, revocation) is done
and covered by ~2023 passing tests. **U8 (audit log)** and **U9 (deployment +
docs + cost estimate)** remain. This doc is the cold-start brief: read it, read
the plan, then continue.

> This file is a durable pointer for whoever picks the work up next. It is not
> itself part of the feature. If you are that agent — start at "Start here".

---

## Start here (first 5 minutes)

1. Read the plan: `docs/plans/2026-07-11-001-feat-remote-connector-mode-plan.md`
   — units U1–U9, requirements, per-unit files/approach/verification. **U8 is
   §"U8. Audit log"; U9 is §"U9. Documentation, deployment hand-off".**
2. Read the requirements: `docs/brainstorms/2026-07-11-remote-connector-mode-requirements.md`
   (R1–R16 — U8 implements **R16**, U9 implements **R9/R10/R11/R13**).
3. Read the provisioning runbook: `docs/remote/provisioning.md` — the Entra apps,
   the concrete IDs, admin-consent commands, and the "add to claude.ai" steps.
4. `npm ci && npm run build && npm test` — confirm green before you touch anything
   (~2023 tests). Lint/typecheck: `npm run lint && npm run typecheck`.
5. This is a JBC repo (**not** a client-deliverable org), so **keep** the
   `Co-Authored-By` trailer on commits.

---

## What is already done (merged to `main`)

| Unit | What | Key files |
|------|------|-----------|
| U2 | SDK bump to `@modelcontextprotocol/sdk` ^1.29.0 (Streamable HTTP) | `package.json` |
| U3 | Stateless Streamable HTTP transport, `serve` subcommand | `src/remote/http-server.ts`, `src/cli.ts` |
| U4 | Entra JWT validation + RFC 9728 PRM handshake (401 + WWW-Authenticate) | `src/remote/auth/{verify,metadata,middleware,deny-list}.ts`, `src/remote/config.ts` |
| U5 | On-Behalf-Of per-user Graph (MSAL CCA, `homeAccountId` state key) | `src/remote/auth/obo.ts`, `src/graph/client/graph-client.ts`, `src/graph/repository.ts` |
| U6 | Per-user entitlements (pinned tool surface, mtime hot-reload) | `src/remote/entitlements.ts`, `src/registry/registry.ts` |
| U7 | Revocation deny-list + `revoke` CLI subcommand | `src/remote/revocation.ts`, `src/state/{schema,store}.ts`, `src/cli.ts` |

Infra (separate repo `jbctechsolutions`/`jp-infrastructure`, applied to the JP
tenant): `stacks/azure/entra/mcp-office365-connector/` — two Entra app
registrations, admin consent granted.

Docs merged: `docs/remote/provisioning.md`,
`docs/solutions/architecture-patterns/stateless-http-transport-for-stdio-mcp-server.md`.

### Concrete values (JP tenant — already provisioned)

| Thing | Value |
|-------|-------|
| Tenant ID | `761e2c5f-34bd-4872-b86c-3a9f3b29d63a` (single-tenant) |
| Connector **Client** app (goes in claude.ai) | `340f710a-af99-4887-b4de-361b47cdd938` |
| Connector **API** app (OBO client → Graph) | `484c0657-6a05-4aad-a175-dabac48acb05` |
| Application ID URI | `api://mcp-office365-connector` |

### Environment contract (the `serve` runtime reads these)

```
OUTLOOK_MCP_TENANT_ID            = 761e2c5f-34bd-4872-b86c-3a9f3b29d63a
OUTLOOK_MCP_CONNECTOR_API_ID     = 484c0657-6a05-4aad-a175-dabac48acb05
OUTLOOK_MCP_CONNECTOR_URL        = https://<public-mcp-host>/mcp   (full URL incl. /mcp)
OUTLOOK_MCP_CONNECTOR_APP_ID_URI = api://mcp-office365-connector   (optional; derived if unset)
OUTLOOK_MCP_ENTITLEMENTS         = /path/to/entitlements.json      (read-only mount)
OUTLOOK_MCP_STATE_DIR            = /path/to/state                  (SQLite; see volume rules)
# OBO credential — certificate PREFERRED over secret:
OUTLOOK_MCP_CONNECTOR_CERT_KEY        + OUTLOOK_MCP_CONNECTOR_CERT_THUMBPRINT   (cert path)
OUTLOOK_MCP_CONNECTOR_CLIENT_SECRET                                            (fallback)
```

Run it locally: `node dist/index.js serve --host 127.0.0.1 --port 3000`
(endpoints: `POST /mcp`, `GET /healthz`, plus the PRM metadata routes).
`node dist/index.js revoke --oid <oid> --reason "..."` for revocation.

---

## Non-negotiable invariants (these caught real bugs — do not regress)

Adversarial review is the **merge gate** for auth/identity/state units
(`docs/solutions/conventions/adversarial-review-as-primary-gate.md`). Every one
of these was a live defect found and fixed:

1. **Fail-closed everywhere.** A token is accepted only on a *positive*
   delegated-member signal; any absent optional claim → reject (see
   `src/remote/auth/verify.ts`: `acct===0` required, app-only rejected, `scp`
   must contain `access_as_user`). Never treat "claim missing" as pass.
2. **Degraded StateStore = refuse to serve.** If SQLite degrades to in-memory,
   the deny-list would be empty (fail-open security control). `serve` refuses to
   start / returns 503, and `revoke` aborts with non-zero exit. Keep this.
3. **Per-user isolation via `homeAccountId` (`<oid>.<tid>`).** Never bind
   multiple users to one cached account. When OBO isn't provisioned, the
   `remoteMode` flag makes `initializeGraphBackend` throw fail-closed rather than
   fall through to a shared/device-code account.
4. **oid/tid are lowercased** on the read path (`verify.ts`) and the `revoke`
   write path so deny-list keys and account keys match byte-for-byte.
5. **DNS-rebinding protection** is required even on the loopback endpoint
   (`allowedHostsFor` in `http-server.ts`).
6. **`res.on('close')` teardown must `.catch()`** — an unhandled rejection crashes
   the shared process.
7. **No token material** in envelopes, logs, deny-list, or (U8) audit rows —
   identity is oid/tid only.
8. **Entitlement preset composition:** `--preset` outer bound composes with the
   per-user allow-list by **intersection** (`registry.ts` `matches()`); the
   elicit path re-checks `matches()` too.
9. **stdio must not regress** — npm stdio mode is the default; remote is additive.

---

## Remaining work

### U8 — Audit log (fully autonomous; do this first)

**Goal (R16):** every write/destructive tool call in remote mode is durably
logged — oid, tool name, target, timestamp, prepare/confirm outcome — readable
for the pilot-exit review.

- **Create** `src/remote/audit.ts` (writer + query); add an `audit` subcommand to
  `src/cli.ts` (filter by user/time, human-readable output).
- **Modify** `src/state/store.ts` + `src/state/schema.ts` (new `audit` table,
  additive **v3→v4** migration — mirror the U7 `deny_list` migration exactly).
  Log at the **CallTool chokepoint** — it lives in `createServer`'s handler in
  `src/index.ts` (shared with stdio), *not* in `http-server.ts`. Thread an audit
  sink through `buildToolContext`/`ToolContext`; it is **null for stdio** so
  auditing keys off the injected sink, not a separate code path.
- **Record non-read tools only** — key off the existing `destructive` /
  `readOnlyHint` tool annotations. Extract the target from the durable-ID param
  when present. Link prepare→confirm rows.
- **Fail-closed rule:** if the audit insert fails, **all `confirm_*` operations
  abort** with a retriable error (a `confirm_send_email` sends mail from a client
  tenant — it must not proceed unaudited). Fail-open-with-warning is reserved for
  non-two-phase writes only.
- **Retention:** keep everything through the pilot (no pruning at v1).
- **Test** `tests/unit/remote/audit.test.ts`: (a) prepare/confirm send_email →
  two linked rows w/ identity+outcome; (b) read-only tool → no row; (c) audit
  table unavailable → confirm_delete AND confirm_send_email abort, non-two-phase
  write proceeds with warning; (d) rows stay readable after an additive upgrade.
- **Verify:** after a scripted mixed read/write session,
  `audit --user <oid>` reconstructs exactly the writes with correct attribution.

**New error classes** (401 challenge, OBO failure, entitlement-denied,
deny-listed, audit-fail-closed) should be typed `OutlookMcpError` codes so the
`{code, retriable, suggestion}` envelope holds for remote clients. Note
`OutlookMcpError` is **abstract** — extend a concrete subclass (`GraphError`,
`GraphAuthRequiredError`).

### U9 — Deployment hand-off, docs, cost estimate (needs Joel's infra)

- **Create** `docs/remote/user-guide.md` (R13, for MCP-unfamiliar JP staff: add
  connector, sign in, what errors mean — guest rejection, reconnect on expiry, CA
  symptoms, keep per-tool approval prompts ON for confirm tools).
- **Create** `docs/remote/deployment.md` (requirements hand-off for
  jp-infrastructure): container image, env contract (above), **SQLite volume
  constraint — NFS or `nobrl`+non-WAL, single replica**, health endpoint, log
  expectations, Anthropic egress allowlist `160.79.104.0/21` if WAF'd, IPv4
  A-record requirement, entitlement config mounted read-only and isolated from the
  state volume, TLS-terminating ingress assumed. **SQLite file + volume snapshots
  are credential material.**
- **Create** `docs/remote/pilot-runbook.md` (R11 exit criteria + observation
  checklist: throttling under one shared app registration, 300s-timeout-prone
  tools, keepalive regression, 401-spike watch, security-denial review, audit
  review step).
- **Add to `provisioning.md`:** OBO credential lifecycle — Entra secrets ≤24mo
  (often 6), the create-new→swap→delete rotation (MSAL CCA takes one credential),
  a calendar expiry reminder, and the AADSTS7000222 total-outage symptom.
- **Cost estimate (R10, in deployment.md):** Azure Container Apps single
  always-on replica (~0.25–0.5 vCPU) + premium file share (NFS) + Log Analytics;
  pilot (~3 users) and full-JP projections, flagged provisional (precedes real
  usage data).
- **Modify `README.md`:** remote mode section; close the `--preset` doc debt.

---

## Joel's manual items (not code — need his hands/tenant)

1. **U1 spike / first live test.** The connector is technically addable now.
   Point claude.ai at a tunnel to a local `serve` (OBO client secret set) and
   confirm the handshake + a real tool call. In claude.ai: Settings → Connectors →
   Add custom connector → URL `https://<tunnel>/mcp`, Advanced → Client ID
   `340f710a-af99-4887-b4de-361b47cdd938` (no client secret — public client).
   **Critical check:** a real JP member's token must carry `acct: 0` (member); a
   guest gets `acct: 1` → 403 by design. Record Joel's JP-tenant oid for the U6
   entitlement config, and add his account to the enterprise-app assignment list.
   NB: sign in with his **JP-tenant** account per R4 — **not** jbc.dev.
2. **U9 deployment** — the Azure Container App + Key Vault + the **OBO
   certificate** (the last thing needed for a production end-to-end Graph call).
   Codify in `jp-infrastructure` (Terraform/Terramate) — never ad-hoc portal.
3. **Admin consent** (if not already done): `az ad app permission admin-consent
   --id 484c0657-6a05-4aad-a175-dabac48acb05` (JP tenant).

---

## Housekeeping / known issues

- **windows-latest 20.x CI is a flaky native build** (`better-sqlite3` prebuild).
  It has been admin-bypassed on each merge. Real fix (a chore, not part of this
  feature): pin/cache the `better-sqlite3` prebuild for Windows 20.x.
- **A few fix commits are unsigned** — the 1Password signing agent was down during
  the session. Re-sign if branch protection requires signed commits.
- Repo specifics: ESM, strict TS with `exactOptionalPropertyTypes: true`, eslint
  `explicit-function-return-type`. `jose` v6 dropped `KeyLike` — use
  `Parameters<typeof jwtVerify>[1]`. For `vi.mock` of `@azure/msal-node`, prefer
  extracting+exporting the unit under test (as was done with `mapOboError`).

---

## Pipeline / workflow expectations (from CLAUDE.md)

Nontrivial feature work runs the compound-engineering pipeline:
`ce-brainstorm → ce-plan → ce-work → ce-code-review → ce-commit-push-pr →
ce-compound`. The plan and brainstorm already exist, so for U8/U9 the loop is
`ce-work` (with TDD + verification-before-completion as sub-steps) →
`ce-code-review` (**adversarial review is the merge gate for U8** — it persists
state and touches the write path) → commit/push/PR → `ce-compound`. Feature
branches cut from `main`, PRs target `main`, squash-merge + delete branch.
Conventional commits. Keep the `Co-Authored-By` trailer (JBC repo).

At session end, append a dated 1–3 line entry under `## Log` in
`~/vaults/cairn/20-projects/JBC-MCP-Office365.md`. Never write secrets there.
