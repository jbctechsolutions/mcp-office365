---
date: 2026-08-13
topic: exec-team-connector-rollout
---

# Exec Team Rollout — JP Remote Connector

## Summary

Roll the JP remote connector out to the full Executive Leadership Team (9 members), replacing ad-hoc per-user assignment with a Terraform-managed security group, and widen the pinned tool surface to include Teams channel and chat messaging. The pilot formally closes with a recorded verdict against its exit criteria.

---

## Problem Frame

The remote connector has been live on the JP tenant since 2026-07-17 and added org-wide to JP's claude.ai Team workspace on 2026-07-18. claude.ai visibility is therefore already universal — every JP staff member can see the connector — but sign-in is gated by Entra enterprise-app assignment, and only two people are assigned: Joel and Bud Houston. Everyone else who clicks it fails with `not_member`.

That gate is currently maintained by hand. The `jp-infrastructure` connector stack codifies exactly one assignment (the seeded operator); Bud was added through the portal, which is drift against the project's standing rule that infra changes are codified, never portal-applied. Each additional user compounds the drift.

The pilot was scoped to ~3 users and a couple of weeks, with a documented decision point at the end: widen, hold, or change course. It has run since 7/18 at two users. The exec team wants access now, which forces that decision rather than deferring it further.

Separately, the exec team's daily coordination happens in Teams, and the connector cannot touch Teams at all — not to read a channel, not to send a chat. The connector API app requests no Teams scopes, and no Teams tool appears in the pinned default surface. For the ELT specifically, that omission covers most of what they would reach for.

---

## Actors

- A1. Operator (Joel): owns the connector, the Terraform stacks, entitlement config, and revocation. The only user on `fullAccess`.
- A2. ELT member: signs in to the connector through claude.ai with their JP account and works inside the pinned default surface.
- A3. JP tenant admin: grants tenant-wide admin consent for new Graph scopes. Distinct from A1 in practice — the `az` CLI path returned `Authorization_RequestDenied` on 2026-08-06 and required portal/PIM escalation.

---

## Key Flows

- F1. Exec onboarding
  - **Trigger:** An ELT member is added to the access group.
  - **Actors:** A1, A2
  - **Steps:** A1 adds the member to the security group → group membership propagates to the Client enterprise app → A2 opens the connector in claude.ai, signs in with their JP account, consents → tools appear.
  - **Outcome:** The member can call tools in the pinned surface as themselves, with every write attributable to their Entra `oid` in the audit log.
  - **Escape path:** A non-member or unassigned account is rejected at token validation (`not_member`), never partially provisioned.
  - **Covered by:** R1, R2, R3, R9

- F2. Teams surface activation
  - **Trigger:** The Teams scope addition is merged and applied.
  - **Actors:** A1, A3
  - **Steps:** Scopes added to the connector API app in Terraform → applied → A3 grants admin consent → A1 restarts the connector replica to clear the MSAL OBO cache → A1 verifies one channel post and one chat send.
  - **Outcome:** Teams messaging tools resolve for all assigned users.
  - **Failure path:** Skipping the replica restart leaves users on pre-consent Graph tokens for ~60–90 minutes, producing 403s that look like a failed consent.
  - **Covered by:** R4, R5, R6, R12

---

## Requirements

**Access and gating**

- R1. A security group governs connector access. The group's assignment to the Client enterprise app is codified in `jp-infrastructure`; adding or removing a user is a membership change, not a Terraform change.
- R2. All nine ELT members are members of that group: Chris Clayman, Duane Frasier, Kelly Benthem, Dan Scribner, Ben Laws, Alan McMahan, Bud Houston, Rotimi Akinpelu, Wilson Geisler.
- R3. Bud's existing portal-applied assignment is absorbed by the group, and the portal-only assignment is removed so the codified state is the only state.
- R9. Member-only enforcement is unchanged: guests and unassigned accounts are still rejected, and `revoke` remains the immediate off-switch independent of group membership.

**Teams surface**

- R4. The connector API app requests the delegated Graph scopes required for Teams reading and messaging, and they are tenant-consented.
- R5. The pinned default tool surface includes Teams team/channel/chat discovery, message reading, and message sending — channel posts, channel replies, and chat sends.
- R6. Channel lifecycle tools (create, update, delete channel) stay out of the default surface, consistent with the existing Planner blast-radius exclusions.
- R7. All Teams sends remain two-phase — a `prepare_` call returns an approval token or elicitation, and nothing posts without an explicit user confirmation.

**Pilot closure**

- R8. The pilot's exit criteria are assessed and the verdict recorded in the pilot runbook, including which criteria were satisfied by evidence and which were accepted without it.
- R10. The prompt-injection posture review (exit criterion 7) is recorded as a reviewed and accepted risk, naming two-phase confirmation, the curated surface, and the audit trail as the defense.

**Operations**

- R11. The container is sized for the larger user set before the execs land, rather than reactively.
- R12. Exec onboarding communication points at the existing user guide and states plainly what the connector can and cannot do, including that sends require their approval.

---

## Acceptance Examples

- AE1. **Covers R1, R3.** Given Bud is assigned to the Client app through the portal, when the group-based assignment is applied and the portal assignment removed, Bud's access is uninterrupted and the codified state matches the live state.
- AE2. **Covers R2, R9.** Given a JP staff member who is not in the access group, when they open the connector in claude.ai and attempt sign-in, they are rejected with `not_member` and no server-side state is created for them.
- AE3. **Covers R4, R5.** Given Teams scopes are consented and the replica has been restarted, when an ELT member asks for their recent channel messages, the tools resolve and Graph returns data rather than a 403.
- AE4. **Covers R6.** Given an ELT member on the default surface, when they attempt to delete a Teams channel, the tool is not exposed to them.
- AE5. **Covers R7.** Given an ELT member drafts a channel post through the connector, when the model calls the send tool, nothing is posted until the member explicitly confirms.

---

## Success Criteria

- Every ELT member who wants access has it and completed sign-in using only the user guide, with no per-person hand-holding from Joel.
- Adding or removing the tenth user is a group membership edit, not a portal visit or a Terraform PR.
- Under real exec use, pooled Graph throttling stays rare enough that no one is regularly blocked — and if it doesn't, the signal is visible early rather than discovered through complaints.
- The audit log attributes every exec write to the right person, including Teams sends.
- `ce-plan` can sequence this without inventing which users, which gating mechanism, which tools, or what happens to the pilot.

---

## Scope Boundaries

- Postgres state store and horizontal scale stay deferred. The deployment remains pinned at exactly one replica; a capacity problem is answered with a vertical bump, not scale-out.
- Access does not wait on the AI Usage Policy signature or the 8/21 training rollout. Deliberate — the two were decoupled.
- No broader JP staff rollout in this pass. The group makes that a later membership change, but widening past the ELT is a separate decision.
- Download tools, shared-mailbox tools, mail rules, `delete_plan`, and `update_plan_sharing` stay out of the default surface, unchanged.
- OBO stays on the client secret (expires 2027-07-18). The certificate migration remains deferred.
- Per-user entitlement tuning is not part of this rollout — all nine get the same surface, and narrowing an individual later is a config edit.

---

## Key Decisions

- **Widen now and close the pilot, rather than expand it.** The exec team is large enough that calling it a pilot expansion would be a fiction. Closing it forces the runbook's decision point to be answered rather than deferred indefinitely.
- **Ending the pilot on a decision, not on evidence.** Exit criterion 3 (throttling under real load) was never exercised past two users. Going 2 → 10 on one shared app registration is an accepted bet, watched rather than measured. This is recorded rather than glossed because it is the one thing the pilot existed to learn and did not.
- **Security group over per-user assignment.** Per-user assignment is codifiable but makes every add a Terraform PR; the group makes the next widening a membership change. Matches the standing "infra codified, never portal" rule and clears the existing Bud drift in the same move.
- **Teams messaging in, on the strength of two-phase confirmation.** Model-drafted posts sent under an exec's name is the largest blast-radius addition here. The mitigation is that a human clicks before anything posts, plus the curated surface and the audit trail. Recorded as the answer to exit criterion 7, not a skipped criterion.
- **Channel lifecycle tools excluded.** Reading and posting is what execs asked for; creating and deleting channels is org-structure change with a much larger blast radius and no stated need. Same reasoning that kept `delete_plan` out.
- **Access not gated on the AI policy.** The connector is tenant-only, member-gated, curated, and audited. Those controls do not become stronger or weaker based on the policy's signature date.

---

## Dependencies / Assumptions

- Admin consent is on the critical path and has historically needed portal/PIM escalation rather than the `az` CLI. Teams tools will not work for anyone until it lands, so execs may have working access before Teams tools appear.
- The replica must be restarted after consent. MSAL's in-memory OBO cache serves pre-consent tokens for ~60–90 minutes otherwise.
- `ChannelMessage.Read.All` is a high-privilege delegated scope. Assumed grantable in the JP tenant on the delegated flow; Microsoft's protected-API restrictions on Teams message endpoints should be confirmed before relying on it.
- Wilson Geisler's UPN was not found in existing records, unlike the other eight (`<first>@joshuaproject.net`). Assumed to follow the same pattern — verify before assignment.
- The ELT roster is taken from the 2026-08-11 exec meeting attendee list, which that note flags as carried forward from 7/21 and not independently confirmed.
- Everything here assumes the connector stays on one replica with pooled Graph throttling across all users.

---

## Outstanding Questions

### Deferred to Planning

- [Affects R4][Needs research] Is `Chat.ReadWrite` sufficient for chat sends, or is `ChatMessage.Send` also required? The local stdio app requests both; the minimal delegated set for the connector should be established rather than copied.
- [Affects R5][Technical] Which Teams tools belong in the pinned list beyond messaging — reactions, team member listing, presence? Presence in particular may pull in additional scopes.
- [Affects R11][Technical] What size is right for ten users? The cost estimate suggests 0.5 vCPU / 1 GiB for full JP against the current 0.25 / 0.5, but nothing has been measured at load.
- [Affects R1][Technical] Does a suitable JP security group already exist, or does the rollout create one?
