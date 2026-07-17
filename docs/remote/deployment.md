# Remote connector — deployment requirements & cost estimate (JP)

The requirements hand-off for standing up the remote connector in Azure. It is
written so the `jp-infrastructure` Terraform work can proceed **without asking
this repo questions**: everything the runtime needs — the image, the environment
contract, the storage constraints, the network posture, and the credential
lifecycle — is here, plus a provisional cost estimate (R10) Joel can set a
ceiling against.

Platform target: **Azure Container Apps** (the JP standard). The server code is
platform-agnostic Node.js, so the same contract self-hosts on Vercel / Cloud Run
/ ECS; only the ingress + volume mechanics differ. See "Self-hosting elsewhere".

Related: [`provisioning.md`](./provisioning.md) (Entra apps + admin consent, done),
[`user-guide.md`](./user-guide.md) (end-user setup), [`pilot-runbook.md`](./pilot-runbook.md)
(exit criteria).

---

## 1. What runs

One long-lived process: `node dist/index.js serve --host 0.0.0.0 --port 8080`.

The container image is built from the repo's [`Dockerfile`](../../Dockerfile)
(multi-stage, Node 26, non-root `node` user, the native `better-sqlite3` addon
compiled into the image, `/healthz` HEALTHCHECK, `serve` as the entrypoint). CI
([`.github/workflows/deploy-connector.yml`](../../.github/workflows/deploy-connector.yml))
builds it, pushes a SHA-tagged image to the JP registry `jpcontainerregistry`
over Azure OIDC (no stored credentials), then rolls the staging Container App and
health-gates the new revision.

- **Transport:** stateless Streamable HTTP. A fresh MCP server is built per
  request over one process-scoped SQLite store — there is no session affinity to
  preserve, but see the single-replica constraint below (the store is the shared
  state, not the transport).
- **Endpoints:**
  - `POST /mcp` — the MCP endpoint (Entra JWT required).
  - `GET /healthz` — liveness/readiness. Returns `{"status":"ok"}`, leaks no
    version or config. Point the Container App health probes here.
  - `GET /.well-known/oauth-protected-resource` (+ `/.well-known/oauth-protected-resource/mcp`)
    — RFC 9728 Protected Resource Metadata, unauthenticated (the discovery
    handshake claude.ai performs before auth).
- **Auth is enabled** whenever `OUTLOOK_MCP_CONNECTOR_URL` is set. With it set, a
  non-loopback bind is permitted and every `/mcp` call requires a valid member
  JWT. Without it, the server refuses any non-loopback bind (fail-safe: it can
  never be exposed unauthenticated).

---

## 2. Environment contract

Set via the deployment platform (Container Apps secrets / Key Vault refs), never
in the image or the repo.

| Variable | Required | Purpose |
|----------|----------|---------|
| `OUTLOOK_MCP_TENANT_ID` | yes | JP tenant id (`761e2c5f-34bd-4872-b86c-3a9f3b29d63a`). Single-tenant. |
| `OUTLOOK_MCP_CONNECTOR_API_ID` | yes | API app id (`484c0657-6a05-4aad-a175-dabac48acb05`) — the JWT audience + OBO client. |
| `OUTLOOK_MCP_CONNECTOR_URL` | yes | The full public MCP URL incl. path, e.g. `https://mcp-o365.jp.example.org/mcp`. This is the RFC 8707 resource + the PRM `resource`; it also enables auth. |
| `OUTLOOK_MCP_CONNECTOR_APP_ID_URI` | no | Defaults to `api://mcp-office365-connector`. Set only if the app's Application ID URI differs. |
| `OUTLOOK_MCP_STATE_DIR` | yes | Path to the SQLite state dir (the mounted volume — see §3). |
| `OUTLOOK_MCP_ENTITLEMENTS` | recommended | Path to the read-only entitlement config (see §5). Absent → every user gets the pinned default surface. |
| **OBO credential — one of:** | yes¹ | |
| `OUTLOOK_MCP_CONNECTOR_CERT_KEY` + `OUTLOOK_MCP_CONNECTOR_CERT_THUMBPRINT` | preferred | Path to the PEM private key + its cert thumbprint (certificate credential). |
| `OUTLOOK_MCP_CONNECTOR_CLIENT_SECRET` | fallback | Client secret (simpler, shorter-lived — prefer the cert). |

¹ The discovery handshake and sign-in work without an OBO credential, but **tool
calls fail closed** (`GRAPH_*`) until one is set. Ship with the certificate.

The server fails fast at startup on a partial/invalid auth config (a set
`OUTLOOK_MCP_CONNECTOR_URL` with missing tenant/API id), so a misconfigured
deploy never serves in a half-authenticated state.

---

## 3. Storage — the load-bearing constraint

The durable store is a single **better-sqlite3** database at
`$OUTLOOK_MCP_STATE_DIR/state.db` (WAL mode). It holds approval tokens, durable-ID
aliases, delta cursors, the **revocation deny-list** (U7), and the **audit log**
(U8). Two of those are security controls, so durability is not optional.

**Hard requirements:**

1. **A real POSIX-locking filesystem.** SQLite's WAL needs working file locks.
   - **Azure Files NFS v4.1** (premium tier) works and is the recommended mount.
   - **Azure Files SMB does NOT** honor the locks WAL needs — if SMB is the only
     option, mount with `nobrl` and run the DB in a non-WAL journal mode. NFS is
     strongly preferred; treat SMB as a last resort.
   - An ephemeral container-local disk works for a smoke test but loses all
     durable state (deny-list + audit) on every restart — never for the pilot.
2. **Exactly one replica.** `minReplicas = maxReplicas = 1`. The single-writer
   assumption is load-bearing: two replicas writing one SQLite file over a shared
   mount will corrupt it. Scale vertically (more vCPU/memory), never horizontally,
   until the store is replaced with a networked DB (out of scope for v1).
3. **If the store can't be opened durably, the server refuses to start** in
   authenticated remote mode (a degraded in-memory store would fail *open* —
   empty deny-list, non-durable audit). So a broken mount surfaces as a failed
   deploy / failing `/healthz`, never as a silently-insecure server.

**The SQLite file + any volume snapshot is credential-adjacent material.** It
contains approval tokens and per-user identifiers. Restrict access to the share,
encrypt at rest (Azure Files does by default), and treat snapshots with the same
care as a secret store. Keep the entitlement config on a **separate** read-only
mount from the state volume.

---

## 4. Network posture

- **TLS-terminating ingress** is assumed (Container Apps ingress provides it).
  The Node process speaks plain HTTP behind it.
- **IPv4 A-record required.** claude.ai resolves the connector over IPv4; a
  purely-AAAA endpoint won't be reachable. Ensure the public hostname has an A
  record.
- **DNS-rebinding protection** is on in the server; it validates the `Host`
  header against the configured public host. Ingress must forward the real
  `Host` (Container Apps does).
- **Egress to Anthropic:** if egress is filtered/WAF'd, allow the Anthropic range
  `160.79.104.0/21`. (Inbound is claude.ai → connector; outbound is the
  connector → Microsoft Graph, `graph.microsoft.com` / `login.microsoftonline.com`.)
- **Coarse DoS guards** are built in (4 MB body cap). Consider an ingress-level
  per-IP rate limit and, if the pilot is closed, an IP allowlist — the app
  registration is shared across all users, so one abusive caller can throttle
  everyone (see the pilot runbook).

---

## 5. Entitlement config (optional but recommended)

A JSON file mounted **read-only**, path in `OUTLOOK_MCP_ENTITLEMENTS`. Hot-reloaded
on mtime change (edits take effect on the next call, no restart). Keyed by Entra
`oid`. Shape:

```json
{
  "version": 1,
  "users": {
    "<joel-jp-oid>": { "fullAccess": true },
    "<staffer-oid>": { "allow": ["list_emails", "search_emails", "send_email"] },
    "<other-oid>":   { "exclude": ["confirm_delete_email"] }
  }
}
```

- No entry → the **pinned default surface** (~50 curated tools), never the full
  surface. A server upgrade that adds tools cannot widen a user's surface unless
  they're added here — that's what makes the entitlement pinnable.
- `fullAccess` gives stdio parity (Joel). `allow` replaces the default with an
  explicit list. `exclude` removes from whatever the base is.
- A malformed file fails **safe** (the last good config stays in effect, warned
  once per edit) — a bad edit never widens access.

---

## 6. OBO certificate — lifecycle & rotation runbook

The API app calls Graph On-Behalf-Of using a **confidential credential**. Prefer a
**certificate** over a client secret: the private key lives in the connector's
Azure Key Vault and is mounted/referenced into the container; the public cert is
attached to the API app registration (the `obo_certificate_pem` variable in the
`jp-infrastructure` Entra stack).

**Expiry is a total-outage event.** MSAL's confidential client holds exactly one
credential; when it expires, **every** user's tool calls fail at once with
`AADSTS7000222` (invalid client credential). There is no partial degradation.

**Rotation (do this before expiry, zero-downtime):**

1. Create the **new** cert/secret and attach it to the API app **alongside** the
   current one (Entra allows multiple credentials on an app).
2. Deploy the connector pointed at the new credential (new `CERT_KEY` /
   `CERT_THUMBPRINT`, or new secret) — Container Apps rolls a new revision.
3. Verify a real tool call succeeds on the new revision.
4. **Then** remove the old credential from the API app.

**Guardrails:**
- Entra certificate/secret lifetimes are capped (secrets ≤ 24 months, often
  policy-limited to 6). Set a **calendar reminder ~30 days before expiry** — this
  is the single most likely cause of a full pilot outage.
- Store the expiry date in the infra repo alongside the stack so it's visible.
- The pilot runbook's watch list includes the `AADSTS7000222` outage shape.

---

## 7. Cost estimate (R10) — provisional

**Treat as provisional: this precedes real usage data and is for setting a
ceiling, not billing.** Revisit after the pilot (the R11 gate). USD/month,
East US-ish list prices, one always-on replica.

| Component | Pilot (~3 users) | Full JP (~30 users) | Notes |
|-----------|------------------|---------------------|-------|
| Container Apps (1 replica, always-on) | ~$12–20 | ~$20–35 | 0.25 vCPU / 0.5 GiB pilot; bump to 0.5 vCPU / 1 GiB for full JP. Min=max=1 replica (no scale-out). |
| Azure Files **premium NFS** share | ~$16 | ~$16 | Premium is required for NFS; billed on provisioned capacity (100 GiB min ≈ $16). The DB is tiny; you're paying for the floor. |
| Log Analytics | ~$5 | ~$5–15 | ~5 GB/mo free, then ~$2.30/GB. Audit/security logs are small text. |
| Key Vault | ~$1 | ~$1 | Per-operation; negligible for one cert + occasional reads. |
| **Total** | **~$35–45/mo** | **~$45–70/mo** | Dominated by the storage floor + always-on compute, not per-user load. |

**Why cost barely scales with users:** it's one small always-on container plus a
fixed storage floor. The real scaling limit is **not cost** — it's Graph
throttling under the single shared app registration and the single-replica
throughput ceiling. Both are pilot observation items, not budget items. If full
JP outgrows one replica, the next step is replacing SQLite with a networked DB
(Postgres) to allow horizontal scale — a larger change to cost against then, not
now.

---

## 8. Deploy checklist

- [ ] Image built from this repo, `serve` as the entrypoint, non-root user.
- [ ] All §2 env vars set; OBO **certificate** wired from Key Vault.
- [ ] `OUTLOOK_MCP_STATE_DIR` on a **premium NFS** Azure Files mount; `minReplicas = maxReplicas = 1`.
- [ ] Entitlement config on a separate **read-only** mount; `OUTLOOK_MCP_ENTITLEMENTS` set.
- [ ] Public hostname has an **A record**; TLS ingress forwards the real `Host`.
- [ ] Health probes → `GET /healthz`.
- [ ] Egress allows `160.79.104.0/21` (if filtered) + Graph/login endpoints.
- [ ] Admin consent granted (provisioning Step 1); pilot users assigned (Step 2).
- [ ] Cert expiry reminder set (§6).
- [ ] Smoke test: PRM resolves, a member signs in, one read + one write tool call
      succeed, `node dist/index.js audit` shows the write.

---

## Self-hosting elsewhere

The runtime is portable Node.js; only ingress + the SQLite volume change:

- **Vercel:** the stateless transport fits serverless, but the single SQLite file
  does not — you'd need a networked store (or Vercel's storage) and to drop the
  single-replica assumption. Not recommended for v1 without that swap.
- **Google Cloud Run / AWS ECS/Fargate:** same as Container Apps — one always-on
  task, a POSIX-locking network volume (Filestore / EFS) for the DB, TLS ingress,
  the same env contract. EFS/Filestore honor NFS locks; the single-replica rule
  still holds.
