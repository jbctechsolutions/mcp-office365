---
date: 2026-07-11
topic: remote-connector-mode
---

# Remote Connector Mode (claude.ai custom connector for JP)

## Summary

Add a remote mode to mcp-office365: one shared instance deployed on Azure (via jp-infrastructure) that JP-tenant users add as a claude.ai custom connector, signing in with their Microsoft identity. Per-user tool entitlements live in config reusing the existing preset machinery; everyone gets write access with two-phase destructive guards kept on. The remote-mode code stays platform-agnostic so self-hosting docs for other platforms can follow later.

---

## Problem Frame

The server is stdio-only today (`src/index.ts` uses `StdioServerTransport`), so it works in Claude Code and Claude Desktop but cannot be reached from the claude.ai web app, which only accepts remote connectors. Joel is the only person using Claude Code; JP staff use claude.ai. When Claude helps produce a document that belongs in SharePoint, the current workaround is manual: send it over Teams or paste it into project files, losing provenance and adding friction. Token storage is a local single-machine MSAL file cache (`src/graph/auth/token-cache.ts` → `~/.mcp-office365/tokens.json`), so there is no way for multiple people to use one instance today.

This also serves a training goal: getting JP staff comfortable with agentic workflows against their real M365 environment, with Joel wanting the same capabilities he already has locally.

---

## Actors

- A1. Joel: operates the deployment; also a user, connecting with his JP-tenant account.
- A2. JP staff member: adds the connector in claude.ai, signs in with their JP M365 account, reads and writes SharePoint documents and Planner items.
- A3. claude.ai (web app): the MCP client; connects over Streamable HTTP and drives the OAuth sign-in.

---

## Key Flows

- F1. Onboarding a user
  - **Trigger:** A JP staff member adds the connector URL in claude.ai settings.
  - **Actors:** A2, A3
  - **Steps:** Add custom connector → claude.ai initiates OAuth → user signs in with their JP Microsoft account → server stores that user's Graph tokens keyed to their identity → connector tools appear, scoped to the user's entitlements.
  - **Outcome:** The user can invoke tools against their own mailbox/files; no admin action needed beyond the user being in the JP tenant.
  - **Covered by:** R3, R4, R5, R6, R13

- F2. Document write-back (the core value flow)
  - **Trigger:** A user drafts a document with Claude in claude.ai and wants it in SharePoint.
  - **Actors:** A1 or A2, A3
  - **Steps:** User asks Claude to save the document → Claude calls the upload/write tool → two-phase confirmation fires for destructive/overwrite cases → file lands in the SharePoint library under the user's own identity.
  - **Outcome:** Document is in the right library with correct authorship; no Teams-forwarding step.
  - **Covered by:** R5, R7, R8

- F3. Entitlement change
  - **Trigger:** A user needs a broader or narrower tool surface.
  - **Actors:** A1
  - **Steps:** Joel edits the entitlement config → change is committed/reviewed like other infra changes → server picks up the new entitlements.
  - **Outcome:** The user's visible tool set changes; no dashboard, no redeploy of user state.
  - **Covered by:** R6

- F4. Offboarding a user (inverse of F1)
  - **Trigger:** A JP staff member departs, asks to disconnect, or Joel needs to cut off access.
  - **Actors:** A1
  - **Steps:** Joel triggers the revoke action (or removes the user's config entry with revoke) → the user's stored Graph tokens are deleted server-side → any disabled Entra account loses access on next token use regardless.
  - **Outcome:** No credentials for that user remain at rest; documentation states that removing the connector in claude.ai does not by itself clear server-side tokens.
  - **Covered by:** R15

---

## Requirements

**Remote transport**
- R1. The server supports a remote mode usable as a claude.ai custom connector (Streamable HTTP), in addition to the existing stdio mode. Stdio remains the default npm/local experience; remote mode is additive and must not regress local usage.
- R2. Remote-mode code is platform-agnostic: no Azure-specific coupling in the server; Azure specifics live only in the deployment layer.

**Identity and auth**
- R3. Users authenticate by signing in with their Microsoft account; Microsoft identity is the only user identity (no separate account/password system).
- R4. Sign-in is restricted to member accounts homed in the JP tenant at v1. Non-JP accounts (including Joel's jbc.dev identity) and Entra B2B guest/external identities in the JP directory are rejected.
- R5. Graph tokens are stored server-side per user identity, isolated so one user's session can never read or act through another user's tokens, and persist across server restarts.
- R15. Offboarding: Joel (or a JP admin) can trigger server-side revocation/purge of a specific user's stored tokens on demand; a user whose JP Entra account is disabled loses all access on next token use; and the operational docs state that removing the connector in claude.ai does not by itself clear server-side tokens.

**Entitlements and tool surface**
- R6. Per-user tool entitlements are defined in config (extending the existing preset/read-only machinery to apply per-user rather than per-process) and enforced server-side. Entitlement changes take effect on the user's next tool invocation, without requiring a server restart or user re-authentication.
- R7. The default v1 tool surface is SharePoint/files, Planner, mail, and calendar. Shared-mailbox and mail-rules tools are excluded from the remote surface at v1 — note this requires an exclusion capability, since those tools are members of presets the default surface includes and the current preset filter is include-only. Any JP-tenant user without an explicit config entry gets this default surface. The default surface is pinned to an explicit, versioned tool list in the jp-infrastructure config, so a server upgrade cannot change any user's effective toolset without a reviewed config change in JP infra.
- R8. All v1 users are write-enabled; the existing two-phase (`prepare_*`/`confirm_*`) guards for destructive operations remain active in remote mode. In remote mode the prepare/confirm pair is model-mediated (Claude issues both calls) and is not by itself a human confirmation: R13's documentation must instruct users to keep claude.ai's per-tool approval prompts enabled for `confirm_*` tools (that prompt is the human confirmation layer), and the R11 pilot explicitly validates that destructive flows pause for the user.
- R14. Joel's config entry grants the full local tool set (matching his current stdio/Claude Code experience), including the shared-mailbox and mail-rules tools excluded from the default surface — satisfying the Problem Frame's parity goal.

**Operations, rollout, and cost**
- R9. Deployment is on Azure, codified in jp-infrastructure (Terraform/Terramate) — no ad-hoc portal changes.
- R10. A monthly running-cost estimate is produced before deployment, covering pilot scale (~3 users) and projected full JP rollout; Joel sets a budget ceiling once numbers exist.
- R11. Rollout is staged: Joel plus 1–2 JP pilot users first; broader JP invitation only after the pilot meets explicit exit criteria: each pilot user completes F2 end-to-end without Joel's live help, zero cross-user data incidents (AE1 holds in practice), and observed cost tracks the R10 estimate over the pilot period.
- R16. Every write/destructive tool invocation is logged server-side with user identity, tool name, target resource, timestamp, and prepare/confirm outcome, retained at least through the pilot; the R11 pilot-exit review reads this log.

**Documentation**
- R12. (reserved — merged into R13)
- R13. End-user setup documentation exists for JP staff: how to add the connector and sign in, written to support the training goal (assume no MCP familiarity).

---

## Acceptance Examples

- AE1. **Covers R5.** Given two signed-in users A and B, when B invokes a mail tool, results come exclusively from B's mailbox; A's tokens are never used.
- AE2. **Covers R4.** Given a user signing in with a non-JP account (e.g., a jbc.dev identity), the sign-in is rejected and no tokens are stored.
- AE3. **Covers R8.** Given a user deletes a file via the connector, when the delete tool is invoked, the two-phase prepare/confirm flow is still required before anything is deleted.
- AE4. **Covers R6, R7.** Given a user whose config entry excludes Planner, when their connector session lists tools, Planner tools are absent (not merely erroring on call).
- AE5. **Covers R1.** Given the remote mode ships, when a local user runs `npx @jbctechsolutions/mcp-office365` with no new flags, behavior is unchanged from today.
- AE6. **Covers R4.** Given a B2B guest account that exists in the JP directory (userType Guest / foreign home tenant), when they attempt sign-in, the sign-in is rejected and no tokens are stored.
- AE7. **Covers R15.** Given a user with stored tokens, when Joel triggers the revoke action for that user, their tokens are deleted and their next tool call requires a fresh sign-in (which R4 gates).

---

## Success Criteria

- Joel can draft a document in claude.ai web and land it in a JP SharePoint library without a Teams-forwarding or copy-paste step.
- A JP pilot user can onboard end-to-end (add connector → sign in → first successful tool call) using only R13's documentation.
- A written cost estimate exists before the first deployment, and the running service stays within the ceiling Joel sets.
- ce-plan can proceed without inventing product behavior: user set, tool surface, tenancy, entitlement model, and rollout order are all decided here.

---

## Scope Boundaries

- Entitlement dashboard/admin UI — deferred until rollout passes ~10 users or config churn becomes roughly weekly; v1's server-side per-identity enforcement is the seam it would later manage.
- Self-host documentation for Vercel, Google Cloud, and AWS — planned later phase, enabled by R2's platform-agnostic design but not written at v1.
- Other tenants: Joel's jbc.dev identity and any multi-tenant support.
- Hosted SaaS for the public (mcp-office365.jbc.dev as a service) — not this project.
- Replacing Anthropic's native M365 connector for read/search use cases — this project's value is write-back, Planner, and deep tooling.

---

## Key Decisions

- **One shared instance + Microsoft identity + config entitlements** (over per-user instances or a dashboard-first entitlement service): no user DB or admin UI to build; Entra ID is the user database; entitlement changes are code-reviewed like the rest of JP infra. Chosen for a small, known user set.
- **JP tenant only at v1**: keeps the governance boundary clean (JP client, JP infra repo, JP data). Joel accepts that claude.ai reaches only his JP mail/files, never JBC's.
- **Scoped default surface instead of all 229 tools**: large tool lists bloat every claude.ai conversation and widen the blast radius; the JP use case needs SharePoint/files, Planner, mail, calendar.
- **Unconfigured JP-tenant users get the default surface** (not default-deny): tenant restriction already gates who can connect, and low-friction onboarding serves the training goal. Config exists to restrict or extend individuals. Revisable if the pilot shows otherwise.
- **Platform-agnostic core, Azure-specific infra layer**: preserves the later self-hosting phase without carrying multi-cloud cost at v1.

---

## Dependencies / Assumptions

- JP's claude.ai plan supports custom connectors with admin enablement — confirmed 2026-07-11.
- An Entra app registration in the JP tenant with the delegated Graph scopes the v1 surface needs (some, like `Sites.ReadWrite.All`, may require admin consent).
- Joel has a JP-tenant account (his sign-in path at v1).
- Anthropic's native M365 connector may eventually add SharePoint write support, narrowing this project's moat to Planner/deep tooling — accepted risk for a self-hosted, training-oriented deployment.

---

## Outstanding Questions

### Deferred to Planning

- [Affects R3][Needs research] The exact OAuth pattern claude.ai custom connectors require (e.g., dynamic client registration) and how it composes with Entra ID — likely the project's highest-risk unknown.
- [Affects R5, R8][Technical] Server-side storage backend and encryption-at-rest approach for all per-user durable state — Graph tokens, two-phase approval tokens, and the durable-ID alias store (all currently in local `~/.mcp-office365` files) — including behavior across restarts and replicas (a `prepare_*` on one instance must be confirmable on another, or replicas must be ruled out).
- [Affects R9, R10][Needs research] Azure service choice and the cost-model inputs for the estimate.
- [Affects R6][Technical] How per-user preset enforcement composes with the existing per-process tool registry.
