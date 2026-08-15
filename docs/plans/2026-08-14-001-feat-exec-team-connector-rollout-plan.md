---
title: "feat: Exec team rollout — JP remote connector"
type: feat
status: active
date: 2026-08-14
origin: docs/brainstorms/2026-08-13-exec-team-connector-rollout-requirements.md
---

# feat: Exec team rollout — JP remote connector

> **Cross-repo plan.** Units are labeled with their target repo. `mcp-office365` paths are relative to this repo; `jp-infrastructure` paths are relative to the `joshua-project/jp-infrastructure` checkout. Two units are tenant-admin actions with no repo at all.

## Summary

Ship exec-team access as one coordinated rollout: a Terraform-managed security group replaces hand-maintained assignment, the connector API app gains the delegated Teams scopes, and the pinned tool surface widens to cover Teams discovery, reading, sending, and reactions. Ordering is the hard part — consent must land before the new image deploys so the fresh revision starts with a clean OBO cache.

---

## Problem Frame

Nine ELT members need connector access that today is maintained by hand, and the surface they'd reach for most — Teams — isn't reachable at all. See origin for the full framing.

Plan-specific: the work is almost entirely permissions and configuration. No new server behavior is being built, which makes sequencing and tenant-state correctness the whole risk surface rather than code correctness.

---

## Requirements

- R1. A security group governs connector access; its app assignment is codified, and membership changes do not require Terraform.
- R2. All nine ELT members are in that group.
- R3. Bud's portal-applied assignment is absorbed and removed so codified state is the only state.
- R4. The connector API app requests and holds tenant consent for the delegated Teams scopes.
- R5. The pinned default surface includes Teams discovery, message reading, sending (channel post, channel reply, chat send), and reactions.
- R6. Channel lifecycle tools stay out of the default surface.
- R7. Teams sends remain two-phase.
- R8. Pilot exit criteria are assessed and the verdict recorded, distinguishing evidence-backed from accepted-without-evidence.
- R9. Member-only enforcement and `revoke` are unchanged.
- R10. The prompt-injection posture review is recorded as reviewed and accepted.
- R11. The container is sized for the larger user set.
- R12. Onboarding communication points at the user guide and states what the connector can and cannot do.

**Origin actors:** A1 (Operator/Joel), A2 (ELT member), A3 (JP tenant admin)
**Origin flows:** F1 (Exec onboarding), F2 (Teams surface activation)
**Origin acceptance examples:** AE1 (covers R1, R3), AE2 (covers R2, R9), AE3 (covers R4, R5), AE4 (covers R6), AE5 (covers R7)

---

## Scope Boundaries

- No new server code. Every Teams tool already exists and is declared for the Graph backend; this plan changes an allow-list, not behavior.
- No Postgres store, no horizontal scale — the deployment stays at exactly one replica.
- No OBO certificate migration; the client secret (expires 2027-07-18) stays.
- No per-user entitlement tuning — all nine get the same surface.
- No rollout beyond the ELT.
- No change to downloads, shared-mailbox, mail-rules, `delete_plan`, or `update_plan_sharing` exclusions.
- `list_team_members` is dropped from the proposed surface rather than adding a permission that no other tool needs.

### Deferred to Follow-Up Work

- Presence and people tools on the remote surface: separate decision, would pull in `Presence.Read.All`.
- Codifying group membership itself in Terraform: deliberately left out-of-band so membership edits stay cheap.

---

## Context & Research

### Relevant Code and Patterns

- `src/remote/entitlements.ts` — `DEFAULT_TOOL_SURFACE`, the pinned allow-list (142 entries today). The header comment documents why it's explicit rather than preset-expanded, and records the Planner blast-radius exclusions; the same reasoning applies to the channel-lifecycle exclusion here.
- `tests/unit/remote/entitlements.test.ts` — the registry drift guard plus two negative guards (shared-mailbox/downloads/photos, and the Planner exclusions). The new exclusion follows the second pattern exactly.
- `src/tools/teams.ts` — all Teams tools, each `backends: ['graph']` and each two-phase where it writes.
- `jp-infrastructure` `stacks/azure/entra/jp-prompt-library-mcp/main.tf` — the direct precedent for this plan's access group: `azuread_group` with seeded members and `lifecycle { ignore_changes = [members] }`, wired to the service principal via `azuread_app_role_assignment` on the group's object id.
- `jp-infrastructure` `stacks/azure/entra/mcp-office365-connector/main.tf` — the connector's two app registrations, the existing `required_resource_access` block, and the seeded `client_superadmin` assignment.

### Institutional Learnings

- `docs/solutions/integration-issues/claude-ai-entra-oauth-remote-mcp-connector-2026-07-17.md` — the five fixes the handshake depends on. Nothing here touches them, but any change to the API app's identifier URIs or scopes risks regressing fix 5 (fully-qualified scope).
- `docs/solutions/conventions/adversarial-review-as-primary-gate.md` — auth/identity changes take adversarial review as the merge gate. The surface change is an auth-adjacent allow-list, so it qualifies.
- Vault log 2026-08-08 (Planner rollout): consent-before-deploy meant the new revision started with a clean MSAL OBO cache and no 90-minute stale window. The 2026-08-06 attempt, which consented after deploy, needed an explicit replica restart. This plan uses the working order.
- Vault log 2026-08-06: `az ad app permission admin-consent` returned `Authorization_RequestDenied` from a plain CLI session; consent needed PIM activation or the portal.

### External References

- Graph `chatMessage: setReaction` — delegated permissions are `ChannelMessage.Send` (channel) and `Chat.ReadWrite` / `ChatMessage.Send` (chat). Reactions require no permission beyond what sending already needs.
- Graph `channel: post messages` — least-privileged delegated is `ChannelMessage.Send`. Protected-API and migration constraints apply to *application* permissions only.
- Graph `channel: list messages` — least-privileged delegated is `ChannelMessage.Read.All`. Resource-specific consent applies to the application permission variant, not delegated.
- Graph `chat: list` — least-privileged delegated is `Chat.ReadBasic`; `Chat.Read` covers listing plus reading messages and members, and `ChatMessage.Send` covers sending and reactions. That pair is narrower than the single `Chat.ReadWrite` that would serve both.

---

## Key Technical Decisions

- **Consent before deploy, not deploy then restart.** The MSAL OBO cache is in-process, so a revision created *after* consent never serves pre-consent tokens. This removes the replica-restart step and its ~90-minute failure window entirely.
- **`Chat.Read` + `ChatMessage.Send` for chats, not `Chat.ReadWrite`.** One scope would have covered both, but it is the higher-privileged option — the same trade rejected for `TeamMember.Read.All`. Corrected in review after an earlier revision took the convenient scope.
- **Reactions included.** Verified to cost no additional permission. The only argument against was scope creep, and that argument doesn't survive the permission check.
- **`list_team_members` dropped.** It's the sole tool needing `TeamMember.Read.All`. Adding a directory-read permission for one convenience tool is a poor trade against least privilege.
- **Channel lifecycle exclusion pinned by a negative test.** Mirrors how the Planner exclusions are enforced, so re-adding them has to be a deliberate, reviewed edit rather than an accident.
- **Group membership left out of Terraform state.** Seeded at create, then `ignore_changes` — matching the prompt-library group. This is what makes R1's "membership change, not a Terraform change" true rather than aspirational.

---

## Open Questions

### Resolved During Planning

- What is the least-privileged chat scope set? `Chat.Read` + `ChatMessage.Send`. `Chat.ReadWrite` would cover both in one scope but grants strictly more.
- Do reactions require a broader scope? No — `setReaction` needs only the send permission.
- Is `ChannelMessage.Read.All` subject to protected-API approval? Not on the delegated flow.
- Does the container need resizing? No. Production already runs 0.5 vCPU / 1 GiB with `min_replicas = max_replicas = 1` — the cost estimate's "full JP" figure. R11 becomes verification, not change.
- Does a suitable group already exist? No connector group exists; create one following the prompt-library pattern.

### Deferred to Implementation

- Wilson Geisler's UPN: the other eight follow `<first>@joshuaproject.net`, but his was not found in existing records. Resolve against the directory at U1 time rather than guessing.
- Whether `find_chat` should stay in the surface: it's annotated `readOnlyHint: false` because resolving a chat by participant can create one. Low blast radius, but confirm the behavior against a real tenant before treating it as a read tool in the user guide.

---

## Implementation Units

### U1. Access group and codified assignment

**Repo:** `jp-infrastructure`

**Goal:** A security group governs connector sign-in, seeded with the nine ELT members, with the app assignment codified.

**Requirements:** R1, R2, R3, R9

**Dependencies:** None

**Files:**
- Modify: `stacks/azure/entra/mcp-office365-connector/main.tf`
- Modify: `stacks/azure/entra/mcp-office365-connector/outputs.tf`

**Approach:**
- Add an `azuread_group` for connector users, seeded with the nine members resolved as `azuread_user` data sources, with `lifecycle { ignore_changes = [members] }` so later membership edits survive apply.
- Add an `azuread_app_role_assignment` binding the group's object id to the Client service principal, alongside the existing seeded operator assignment. Keep the operator assignment — it is the break-glass path if group resolution ever fails.
- Export the group object id so onboarding docs and future stacks can reference it.
- After apply, remove Bud's portal-applied user assignment. Order matters: the group assignment must be live first, or Bud loses access between the two steps.
- Resolve Wilson's UPN against the directory before writing the data source; do not assume the naming pattern.

**Patterns to follow:**
- `stacks/azure/entra/jp-prompt-library-mcp/main.tf` — group + `ignore_changes` + group-scoped `azuread_app_role_assignment`.

**Test scenarios:**
- Test expectation: none (Terraform config; no test harness in this repo). Verified by plan/apply output and tenant state.

**Verification:**
- `terraform plan` shows the group and the group assignment as the only additions to this stack; no destroy on the existing operator assignment.
- After apply, the Client enterprise app lists the group under Users and groups, and Bud's individual assignment is gone.
- A group member who is not separately assigned can complete sign-in; a non-member still fails `not_member`.

---

### U2. Delegated Teams scopes on the connector API app

**Repo:** `jp-infrastructure`

**Goal:** The API app requests the delegated Graph permissions the Teams tools need — and nothing more.

**Requirements:** R4

**Dependencies:** None (independent of U1; same file, so likely the same PR)

**Files:**
- Modify: `stacks/azure/entra/mcp-office365-connector/main.tf`

**Approach:**
- Add six `resource_access` entries to the existing Graph `required_resource_access` block: `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `Chat.Read`, and `ChatMessage.Send`.
- Do not add `Chat.ReadWrite` (the higher-privileged single scope that `Chat.Read` + `ChatMessage.Send` replaces) or `TeamMember.Read.All` (dropped with `list_team_members`).
- Follow the existing comment convention in that block: record why each scope is the least-privileged choice and cite the date, as the Planner block does.
- Leave identifier URIs, pre-authorization, and the client app untouched — those carry the handshake fixes.

**Patterns to follow:**
- The `Tasks.ReadWrite` addition in the same block (2026-08-06), including its inline rationale comment and its pointer to the consent runbook.

**Test scenarios:**
- Test expectation: none (Terraform config).

**Verification:**
- `terraform plan` shows only additive `resource_access` entries; no changes to identifier URIs, the client app, or existing scopes.
- After apply, the API app's API permissions list shows the five new rows in a "Not granted" state, ready for U3.

---

### U3. Tenant admin consent for the new scopes

**Repo:** none — tenant action

**Goal:** The new Teams scopes are consented tenant-wide, before any new revision deploys.

**Requirements:** R4

**Dependencies:** U2

**Approach:**
- Grant admin consent for the API app. Expect the CLI path to fail from a plain session — budget for portal or PIM-activated escalation, since that is what happened on 2026-08-06.
- Verify via the grants list that all five new scopes read as granted, not just that the command exited zero. The 2026-08-06 failure looked successful until the grants list was checked.
- This unit must complete before U5 deploys. That ordering is the whole reason no replica restart appears in this plan.

**Test scenarios:**
- Test expectation: none (tenant state change).

**Verification:**
- The permission grants list shows `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `Chat.Read`, and `ChatMessage.Send` as granted for the tenant.

---

### U4. Widen the pinned default tool surface

**Repo:** `mcp-office365`

**Goal:** Teams discovery, reading, sending, and reactions become reachable for default-surface users; channel lifecycle tools stay out and are pinned out by test.

**Requirements:** R5, R6, R7

**Dependencies:** None for the code change; must not deploy before U3

**Files:**
- Modify: `src/remote/entitlements.ts`
- Modify: `tests/unit/remote/entitlements.test.ts`

**Approach:**
- Add to `DEFAULT_TOOL_SURFACE`: `list_teams`, `list_channels`, `get_channel`, `list_channel_messages`, `get_channel_message`, `prepare_send_channel_message`, `confirm_send_channel_message`, `prepare_reply_channel_message`, `confirm_reply_channel_message`, `list_chats`, `get_chat`, `find_chat`, `list_chat_messages`, `list_chat_members`, `prepare_send_chat_message`, `confirm_send_chat_message`, `list_message_reactions`, `prepare_add_message_reaction`, `confirm_add_message_reaction`, `remove_message_reaction`.
- Deliberately omit `create_channel`, `update_channel`, `prepare_delete_channel`, `confirm_delete_channel`, and `list_team_members`.
- Extend the header comment to record the Teams exclusions and their date, matching how the Planner exclusions are documented — the comment is what makes the next reader treat the omission as intentional.
- Add a negative guard test for the excluded channel-lifecycle names, mirroring the existing Planner exclusion test.

**Execution note:** Write the negative guard test before adding the names, so it is proven to fail against a surface that wrongly includes lifecycle tools.

**Patterns to follow:**
- The Planner blast-radius exclusion block in `src/remote/entitlements.ts` and its paired test in `tests/unit/remote/entitlements.test.ts`.

**Test scenarios:**
- Happy path: every newly pinned name resolves in the registry — the existing drift guard covers this and must stay green with the larger list.
- Happy path: a user with no entitlement entry resolves to the widened default surface.
- Edge case: `Covers AE4.` the surface contains none of `create_channel`, `update_channel`, `prepare_delete_channel`, `confirm_delete_channel`.
- Edge case: the surface contains no `list_team_members`, guarding the least-privilege decision against a well-meaning re-add.
- Integration: `Covers AE5.` a default-surface user sees both halves of each two-phase send pair, so no send is reachable without its prepare step.
- Integration: an entitlement entry with an explicit `allow` list still overrides the widened default rather than unioning with it.
- Edge case: the existing shared-mailbox / downloads / photos negative guard still passes — none of the added Teams names trip its patterns.

**Verification:**
- Full suite green, including all three existing contract tests plus the new guard.
- The surface count rises from 142 to 162.

---

### U5. Release and deploy to the JP connector

**Repo:** `mcp-office365`

**Goal:** The widened surface reaches the JP deployment on a revision that starts with a clean OBO cache.

**Requirements:** R5, R11

**Dependencies:** U3, U4

**Approach:**
- Cut a release from `main` and publish, following the established release-PR-then-tag flow.
- Push `main` to `joshua-project/mcp-office365` with a plain `git push` — it is not a GitHub fork and `gh repo sync` 422s.
- Let the deploy workflow build and roll the revision. Because U3 already landed, the new replica's MSAL cache starts empty and there is no stale-token window; do not add a separate restart step.
- Confirm the container is still at 0.5 vCPU / 1 GiB with `min_replicas = max_replicas = 1`. This is R11's verification — no change is expected.

**Patterns to follow:**
- The v4.5.0 and v5.0.0 deploy sequences recorded in the vault log: release, push to the JP remote, watch the deploy run, then health-gate.

**Test scenarios:**
- Test expectation: none beyond U4's suite — this unit ships existing code.

**Verification:**
- Deploy run green including build-and-push; the running image matches the new commit.
- `/healthz` returns ok, the PRM metadata route resolves, and `/mcp` still 401-challenges an unauthenticated request.
- A live Teams read and a live Teams send both succeed for an assigned member; a send does not post until confirmed.

---

### U6. Pilot closure record, docs, and onboarding

**Repo:** `mcp-office365`

**Goal:** The pilot ends on an honest written record, and the nine have what they need to onboard unaided.

**Requirements:** R8, R10, R12

**Dependencies:** U5

**Files:**
- Modify: `docs/remote/pilot-runbook.md`
- Modify: `docs/remote/provisioning.md`
- Modify: `docs/remote/user-guide.md`

**Approach:**
- Add a closure section to the pilot runbook recording the verdict against each of the seven exit criteria, and mark explicitly which were satisfied by evidence and which were accepted without it. Criterion 3 (throttling) was never exercised past two users — say so plainly rather than marking it passed.
- Record the criterion-7 prompt-injection review as reviewed and accepted, naming two-phase confirmation, the curated surface, and the audit trail as the defense, and noting that Teams sending was added under that same judgment.
- Update the provisioning runbook so Step 2 is "add to the access group" rather than per-user portal assignment, and note that offboarding still requires `revoke` in addition to group removal — removing the connector in claude.ai does not clear server state.
- Add a Teams section to the user guide covering what is reachable, that sends require the user's own approval, and that channel creation and deletion are deliberately absent.
- Note the outstanding follow-up: throttling is now the thing to watch, and a spike at nine users is a hold signal for any further widening.

**Test scenarios:**
- Test expectation: none — documentation.

**Verification:**
- The pilot runbook states a verdict for all seven criteria with evidence status attached to each.
- Provisioning Step 2 describes the group path, and no step instructs a portal user assignment.
- A reader following the user guide alone can reach a Teams channel and send an approved message.

---

## System-Wide Impact

- **Interaction graph:** The widened allow-list flows through the registry's `matches()` intersection, so the `--preset` outer bound still composes by intersection and the elicit path re-checks it. Nothing in this plan changes that logic; it changes its input. Verified that the deployment does not set an outer bound — the container runs `serve --host 0.0.0.0 --port 8080` with no `--preset` or `--read-only`, and the Terraform does not override the args. Had a preset been set, the widened list would have been intersected away and the rollout would have silently done nothing.
- **API surface parity:** Local stdio is unaffected — it does not consult `DEFAULT_TOOL_SURFACE`. Joel's `fullAccess` entry already exposed these tools, so his surface does not change either.
- **Error propagation:** If U3 is skipped or partially consented, Teams tools appear in the client and fail at Graph with 403. That is a confusing failure mode for a non-technical user, which is the practical argument for keeping U5 strictly after U3.
- **State lifecycle risks:** Removing Bud's portal assignment before the group assignment is live would drop his access. The group must land first.
- **Integration coverage:** Unit tests prove the surface contains the right names; they cannot prove the Graph scopes are consented or that a real send posts. U5's live verification is the only thing that covers that seam.
- **Unchanged invariants:** Fail-closed token validation, `homeAccountId` per-user isolation, the deny-list, the degraded-store refusal, and the audit chokepoint are all untouched. Every added write tool still routes through the same two-phase audit path.

---

## Risks & Dependencies

| Risk | Mitigation |
|------|------------|
| Admin consent blocked by role escalation, as on 2026-08-06 | Treat U3 as the critical path and start it early; expect portal or PIM rather than CLI. U1 can land and give access meanwhile. |
| Pooled Graph throttling degrades at nine users | Accepted risk, recorded in U6. Throttling is the named watch item; a spike is a hold signal for further widening, not a reason to roll back exec access. |
| A model-driven Teams post embarrasses someone | Two-phase confirmation, curated surface, and audit trail. Recorded as the criterion-7 judgment rather than an unexamined default. |
| Consent lands but users still 403 | Almost always the stale OBO cache. This plan's ordering avoids it; if it appears anyway, the deploy ran before consent. |
| Bud loses access during the assignment migration | Ordering constraint in U1: group assignment live before portal assignment removed. |
| Surface growth degrades the claude.ai tool-picker UX at 162 tools | Watch during rollout. Entitlements are hot-reloaded, so narrowing is a config edit with no restart. |

---

## Documentation / Operational Notes

- Offboarding an exec now has two steps: remove from the group *and* `revoke` their oid. Group removal alone leaves server-side state.
- The OBO client secret expires 2027-07-18. Unchanged by this plan, but the expiry is a total outage for all users at once and now affects nine people instead of two.
- Narrowing one exec's tools remains an entitlement-config edit, hot-reloaded, no restart.
- Vault write-back is outstanding for this decision (pilot closed, widen to ELT) per the repo's logging convention.

---

## Sources & References

- **Origin document:** `docs/brainstorms/2026-08-13-exec-team-connector-rollout-requirements.md`
- Related code: `src/remote/entitlements.ts`, `tests/unit/remote/entitlements.test.ts`, `src/tools/teams.ts`
- Related runbooks: `docs/remote/provisioning.md`, `docs/remote/pilot-runbook.md`, `docs/remote/deployment.md`, `docs/remote/user-guide.md`
- Related learnings: `docs/solutions/integration-issues/claude-ai-entra-oauth-remote-mcp-connector-2026-07-17.md`, `docs/solutions/conventions/adversarial-review-as-primary-gate.md`
- Infrastructure: `jp-infrastructure` `stacks/azure/entra/mcp-office365-connector/`, `stacks/azure/entra/jp-prompt-library-mcp/`
- External: Microsoft Graph permission references for `chatMessage: setReaction`, `channel: post messages`, `channel: list messages`, `chat: list`
