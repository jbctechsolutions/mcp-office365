# Remote connector — provisioning runbook (JP)

One-time admin setup for the mcp-office365 remote **claude.ai custom connector**
in the Joshua Project tenant. This is the U1 deliverable from
[docs/plans/2026-07-11-001-feat-remote-connector-mode-plan.md](../plans/2026-07-11-001-feat-remote-connector-mode-plan.md);
it covers the manual steps Terraform can't do (admin consent, connector add,
Conditional Access validation) and doubles as the U1 handshake-spike checklist.

The connector is **separate** from the existing device-code app used by
stdio/Claude Code — that keeps working untouched.

---

## Identities (already provisioned)

The two Entra app registrations are codified in `jp-infrastructure`
(`stacks/azure/entra/mcp-office365-connector/`) and applied to the JP tenant.

| Registration | Client ID | Role |
| --- | --- | --- |
| **MCP Office365 Connector (Client)** | `340f710a-af99-4887-b4de-361b47cdd938` | Public client entered in claude.ai; carries the redirect URI. |
| **MCP Office365 Connector (API)** | `484c0657-6a05-4aad-a175-dabac48acb05` | Resource server (`api://mcp-office365-connector`); OBO client to Graph. |

- **Tenant:** Joshua Project (`761e2c5f-34bd-4872-b86c-3a9f3b29d63a`), single-tenant.
- **Redirect URI** (on the Client app): `https://claude.ai/api/mcp/auth_callback`.
- **Member-only:** the Client service principal requires assignment; only assigned
  users (or a group) can sign in. Joel is seeded as an assigned member.

Changing any of the above is a `jp-infrastructure` change (Terraform), never a
portal edit.

---

## Prerequisites

- Entra admin in the JP tenant (Global Admin or Privileged Role Admin for the
  consent step).
- claude.ai plan that allows custom connectors (Team/Enterprise with connectors
  enabled by an Owner, or an individual Pro/Max account). Confirmed available
  2026-07-11.
- `az` CLI signed in to the JP tenant for the CLI variants below.

---

## Step 1 — Grant admin consent for the API app's Graph scopes

The API app requests delegated Microsoft Graph scopes (mail, calendar,
files/SharePoint, Planner-read) but they are **not yet consented**. Grant
tenant-wide admin consent so pilot users are not each prompted (and so
`Sites.ReadWrite.All`, which requires admin consent, works).

**Portal:** Entra ID → App registrations → **MCP Office365 Connector (API)** →
API permissions → **Grant admin consent for Joshua Project** → confirm all rows
show "Granted for Joshua Project".

**CLI:**

```bash
az ad app permission admin-consent --id 484c0657-6a05-4aad-a175-dabac48acb05
```

Verify:

```bash
az ad app permission list-grants --id 484c0657-6a05-4aad-a175-dabac48acb05 -o table
```

> Planner **write** is intentionally not requested yet — only Planner-read.
> Enabling Planner writes needs a broader group scope; validate the exact scope
> and re-consent (a `jp-infrastructure` change) before turning it on.

---

## Step 2 — Assign pilot users

Only assigned members can sign in (member-only enforcement). Add pilot users (or
a security group) to the **Client** enterprise app.

**Portal:** Entra ID → Enterprise applications → **MCP Office365 Connector
(Client)** → Users and groups → Add user/group.

Prefer a security group as the rollout widens; per-user assignment can also be
codified in `jp-infrastructure` alongside the seeded operator assignment.

---

## Step 3 — Add the connector in claude.ai

1. claude.ai → Settings → Connectors → **Add custom connector** (Team/Enterprise:
   an Owner adds it under Admin settings; individual plans: your own settings).
2. **URL:** the deployed MCP endpoint, e.g. `https://<host>/mcp`. Until the server
   is deployed (U9), use a tunnel to a local `mcp-office365 serve` for the spike.
3. **Advanced → Client ID:** `340f710a-af99-4887-b4de-361b47cdd938` (client secret
   left blank — public client, PKCE).
4. Sign in with a **JP** account. A member completes Microsoft sign-in and the
   connector's tools appear.

---

## Step 4 — Validate the handshake (U1 spike)

Confirm the two residual unknowns the plan flagged, before U4/U5 build on them:

1. **PKCE advertisement.** Entra's OIDC metadata does not advertise
   `code_challenge_methods_supported` even though S256 works. Confirm claude.ai
   completes the flow anyway (evidence strongly suggests it tolerates this).
2. **https identifier URI.** Once a deployment hostname exists, add the exact MCP
   URL as an Application ID URI on the API app (via `mcp_public_url` in the
   Terraform stack) and confirm the JP tenant accepts it — some tenants restrict
   https identifier URIs to verified domains. claude.ai sends the MCP URL as the
   OAuth `resource` (RFC 8707), so Entra must recognise it or it rejects the
   token request.
3. **Conditional Access.** Sign in once from claude.ai under JP's CA policies and
   record the outcome. If a policy blocks the browser-based client, note the
   `AADSTS` code and add the required exemption before inviting pilot users.

---

## Step 5 — OBO certificate (deferred to deployment / U9)

The API app calls Graph via On-Behalf-Of, which needs a **confidential
certificate credential** on the API app. The private key lives in the connector's
Azure Key Vault (deployment infra, not yet built); the public cert is attached to
the API app via the `obo_certificate_pem` variable in the Terraform stack.

The certificate is **not** required for the Step 4 auth-code/PKCE handshake test —
OBO to Graph is exercised only once the server is deployed and serving requests
(U5/U9). Rotation runbook and the total-outage symptom on expiry
(`AADSTS7000222`) are captured with the deployment work.

---

## Offboarding

- **Revoke access now:** unassign the user from the **Client** enterprise app
  (Step 2, in reverse), or disable their Entra account. Either stops new
  sign-ins.
- **Server-side token purge + deny-list** (so a still-valid claude.ai token can't
  silently re-onboard) is the `revoke` subcommand landing in U7 — see the plan.
- Removing the connector in claude.ai does **not** clear any server-side state.

---

## Troubleshooting (common AADSTS codes)

| Code | Meaning | Fix |
| --- | --- | --- |
| `AADSTS65001` | Consent not granted for a scope | Run Step 1 (admin consent). |
| `AADSTS50105` | User not assigned to the app | Assign the user (Step 2). |
| `AADSTS7000222` | API app credential expired | Rotate the OBO certificate (U9 runbook). |
| `AADSTS53003` / sign-in loop | Conditional Access blocked the client | Add a CA exemption for the connector (Step 4.3). |
| Connector "failed" after a successful Microsoft sign-in | Guest/non-member reached the server and was rejected | Expected — only assigned JP members are authorized. |
