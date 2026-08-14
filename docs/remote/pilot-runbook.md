# Remote connector — pilot runbook & exit criteria (JP)

> **Status: the pilot is CLOSED as of 2026-08-14.** The verdict against each exit
> criterion is recorded in [Pilot closure](#pilot-closure) at the bottom of this
> document. The sections below are kept as the operating reference — the watch
> list and the operating procedures still apply now that the connector is in
> steady use.

How the pilot runs and how it ends. The pilot exists to answer the questions unit
tests can't: does the real claude.ai handshake hold, does one shared app
registration throttle under real use, and is the curated tool surface right? Run
it consciously against the exit criteria below, then decide: widen, hold, or
change course.

Related: [`deployment.md`](./deployment.md) (infra), [`user-guide.md`](./user-guide.md)
(what pilot users get), [`provisioning.md`](./provisioning.md) (assign/offboard).

---

## Scope

- **Users:** start with Joel, then a small handful (~3) of JP staff assigned via
  the provisioning runbook (Step 2). Add users deliberately, one or two at a time.
- **Duration:** run long enough to see real weekly patterns (a couple of weeks),
  not just a demo day.
- **Surface:** the pinned default tool surface for staff; `fullAccess` for Joel.
  Tune the default list from what the pilot shows (it's a config change, not a
  code change).

---

## What to watch (observation checklist)

Check these through the pilot; each maps to a risk the plan flagged.

| Watch | Why | Signal / where |
|-------|-----|----------------|
| **Auth failure rate / 401 spikes** | A healthy pilot has near-zero auth failures after setup. A spike means expired sessions, a CA change, or a token/audience problem. | Server logs emit `auth denied: reason=…` (no token material). Watch for a rising rate or a sudden cluster. |
| **Security denials** | Guests/non-members and deny-listed users must be rejected. Confirm every denial is *expected*. | `reason=not_member` / `deny_listed` in logs. Review periodically. |
| **Graph throttling ("service busy")** | All users share **one** app registration, so throttling is pooled — one heavy user can slow everyone. This is the main thing the pilot is testing. | Users reporting "service busy" / retry errors; `THROTTLED`/`GRAPH_RATE_LIMITED` envelopes. Note frequency and which tools. |
| **Long-running tool timeouts** | claude.ai enforces a ~300s tool timeout; large uploads/downloads can exceed it. | Users reporting hung/failed large transfers. Download tools are excluded from the default surface for this reason — confirm that holds. |
| **Session keepalive / disconnects** | A known SDK keepalive issue can drop sessions behind ingress. | Users reporting the connector "dropping" mid-task; reconnect frequency. Stateless mode reduces exposure — verify in practice. |
| **OBO credential health** | Cert/secret expiry is a **total outage** (`AADSTS7000222`) for everyone at once. | Watch for a sudden all-users failure; keep the expiry reminder (deployment §6). |
| **Audit trail completeness** | R16: every write must be attributable. | Run `node dist/index.js audit --user <oid>` and confirm writes reconstruct correctly (see below). |
| **Store health** | A degraded store would disable the deny-list + audit. | The server refuses to serve on a degraded store, so this shows as a failed deploy / failing `/healthz` — confirm `/healthz` stays green. |

### Audit review step

Periodically, and at exit, run the audit CLI on the deployment host:

```bash
node dist/index.js audit                 # all write/destructive actions, newest first
node dist/index.js audit --user <oid>    # one user (Entra oid; from the logs/report)
node dist/index.js audit --since 2026-07-01
```

Each row is `time · oid · tool · phase · outcome · target · link`. A `prepare`
row and its `confirm` row share an approval-token **link**, so you can trace a
two-phase action end to end. Confirm the writes match what users report doing,
and that nothing unexpected appears.

---

## Exit criteria (R11) — decide at the end

The pilot **passes** (widen toward full JP) when all of these hold:

1. **Handshake is reliable.** New users add the connector and sign in using only
   the [user guide](./user-guide.md), with no hand-holding. (The origin success
   criterion: a JP user onboards end-to-end unaided.)
2. **Auth is clean.** Auth-failure rate is near zero after setup; every security
   denial reviewed was expected (guest/unassigned/revoked).
3. **Throttling is tolerable.** Under real pilot use, "service busy" is rare and
   self-resolves; no user is regularly blocked. If throttling is bad at 3 users,
   it will be worse at 30 — that's a "hold + revisit the shared-registration
   decision" signal, not a "widen" one.
4. **No session-stability regression.** Connector doesn't drop sessions often
   enough to disrupt work.
5. **The audit trail is trustworthy.** A scripted mixed read/write session
   reconstructs exactly, with correct per-user attribution.
6. **The tool surface feels right.** The default list isn't missing something
   staff need daily, nor so large the claude.ai UX suffers. Tune and re-confirm.
7. **Prompt-injection posture reviewed.** Because a prepare→confirm can be driven
   by model-read content, the client-side approval prompts + curated surface +
   audit trail are the defense. Consciously review whether that's sufficient for
   JP's data before widening — this is a judgment call, not a metric.

**Hold or change course** if throttling is bad at pilot scale, the handshake
needs manual intervention per user, or the prompt-injection review isn't
comfortable. Any of those is worth solving before more users depend on it.

---

## Operating during the pilot

- **Add a user:** add them to the connector access group (provisioning Step 2).
  Record their oid only if they need a non-default tool surface.
- **Remove a user now:** `node dist/index.js revoke --oid <oid> --reason "..."`
  (deny-lists + purges their server-side state) **and** remove them from the
  access group. Neither step alone is sufficient, and removing the connector in
  claude.ai does **not** clear server state.
- **Narrow a user's tools:** edit the entitlement config (hot-reloaded, no
  restart).
- **Incident (all users failing):** check `/healthz`, then the OBO credential
  (`AADSTS7000222` = expired cert → rotate per deployment §6), then Graph status.

---

## Pilot closure

**Decision (2026-08-14): widen to the Executive Leadership Team; the pilot ends.**

The pilot ran 2026-07-18 → 2026-08-14 with **two users** (Joel and Bud Houston),
not the ~3 it was scoped for. Access widens to all nine ELT members, gated by a
Terraform-managed security group, with Teams messaging added to the pinned
surface.

The verdict below marks each criterion as **evidence** (satisfied by observation)
or **accepted** (judged acceptable without the evidence the pilot was meant to
produce). The distinction matters more than the pass/fail — an accepted criterion
is a live bet, not a closed question.

| # | Criterion | Verdict | Basis |
|---|-----------|---------|-------|
| 1 | Handshake is reliable | **Evidence** | Bud onboarded from the user guide without hand-holding. One user is thin proof, but it is the criterion's actual bar. |
| 2 | Auth is clean | **Evidence** | No unexplained auth failures over the window. Every denial reviewed was expected (unassigned accounts). |
| 3 | Throttling is tolerable | **Accepted — not tested** | Two users never exercised the shared app registration. This is the one thing the pilot existed to measure and it did not. Going 2 → 10 is a bet, watched rather than proven. |
| 4 | No session-stability regression | **Evidence** | No session-drop reports over the window. Stateless transport appears to hold. |
| 5 | Audit trail is trustworthy | **Evidence** | Writes reconstruct with correct per-user attribution via the `audit` CLI. |
| 6 | Tool surface feels right | **Evidence, and acted on** | The gap users actually hit was Teams — absent entirely. That is what this rollout fixes. |
| 7 | Prompt-injection posture reviewed | **Accepted — reviewed** | See below. |

### Criterion 7 — the injection review, recorded

Reviewed and accepted 2026-08-13. A `prepare_*` → `confirm_*` action can be driven
by content the model read, so a malicious email or channel post can *propose* an
action. The defense is three layers: **client-side approval** on every write (a
human reads and clicks before anything happens), the **curated surface** (channel
lifecycle, plan deletion, plan sharing, downloads, and shared-mailbox access are
simply not reachable), and the **audit trail** (every write attributable to a
person).

Teams *sending* was added under this same judgment, with eyes open: it is the
largest blast-radius addition in this rollout, because a post goes out under an
exec's name to a channel other people read. It was accepted because the approval
step is unchanged — the model can draft, but only a person can send.

This is a judgment, not a metric. It should be re-reviewed before the surface
widens again, and it is not a permanent finding.

### What to watch now

**Throttling is the open question.** Criterion 3 closed on judgment, not data, so
the signal now comes from users rather than from the pilot. A rising rate of
"service busy" reports, or any user regularly blocked, is a **hold** signal — it
means the shared-app-registration decision needs revisiting before access widens
past the ELT, not that exec access should be rolled back. The user guide now asks
people to flag persistent throttling explicitly, since reports are the only
instrumentation for it.

Second-order: the OBO client secret expires **2027-07-18** and its failure mode is
a total outage for everyone at once. That now affects nine people rather than two.
