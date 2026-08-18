# Remote connector — pilot runbook & exit criteria (JP)

> **Status: the pilot is CLOSED as of 2026-08-14.** Read
> [Pilot closure](#pilot-closure) first — it carries the verdict and what to
> watch now.
>
> - **Still live:** [What to watch](#what-to-watch-observation-checklist) and
>   [Operating](#operating). These apply to steady-state use.
> - **Historical:** Scope and [Exit criteria](#exit-criteria-r11--decide-at-the-end).
>   The decision they describe has been made; they are kept so the closure record
>   has something to be a verdict *against*.

How the pilot ran and how it ended. The pilot existed to answer the questions
unit tests can't: does the real claude.ai handshake hold, does one shared app
registration throttle under real use, and is the curated tool surface right?

Related: [`deployment.md`](./deployment.md) (infra), [`user-guide.md`](./user-guide.md)
(what pilot users get), [`provisioning.md`](./provisioning.md) (assign/offboard).

---

## Scope

*Historical — this is what the pilot was scoped to, kept for context. What
actually happened is in [Pilot closure](#pilot-closure).*

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

*Historical — assessed 2026-08-14, see [Pilot closure](#pilot-closure).*

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

## Operating

- **Add a user:** add them to the connector access group (provisioning Step 2),
  and **record their Entra oid** — revocation takes the oid, so not having it on
  file is what turns an urgent offboard into a directory lookup under pressure.
  (`az ad user show --id <upn> --query id -o tsv` if you need to recover one.)
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

The pilot ran 2026-07-18 → 2026-08-14 with **three people able to sign in** —
Joel, Bud Houston, and Kelly Benthem — roughly the ~3 it was scoped for. Access
widens to all nine ELT members, gated by a security group whose app assignment is
Terraform-managed — its **membership deliberately is not**, so adding or removing
a person is a group edit rather than a Terraform run. Teams messaging is added to
the pinned surface.

> **Usage was measured on 2026-08-18** and is no longer an open question. The
> original text here said usage was unverified and told the reader to treat
> three as an upper bound on who *could* have loaded the shared registration.
> That was true when written and is now superseded — see
> [Measured usage](#measured-usage) below. Both Bud and Kelly were really using
> the connector during the pilot, not merely holding assignments.
>
> **The tenant state at rollout, precisely**, since the counts differ depending on
> what is being counted:
>
> | | |
> |---|---|
> | Individual assignments found | **4** — Joel's primary, Joel's admin account (`admin-joel.castillo@`), Bud, Kelly |
> | Distinct people | **3** — Joel held two |
> | Codified in Terraform | **1** — Joel's primary, the seeded break-glass |
> | Removed when the group took over | **3** — Joel's admin account, Bud, Kelly |
> | Remaining after cleanup | Joel's primary (break-glass) + the access group |
>
> Only Bud's assignment was in the written record. Joel's admin account and
> Kelly's were discovered at rollout by querying the tenant, not from any doc.
> That is the concrete cost of per-user portal assignment — the roster
> drifted away from the record without anyone noticing, and the closure record
> was written against the record rather than the tenant.

### Measured usage

Measured 2026-08-18 from Entra sign-in logs on the connector API app
(`484c0657-6a05-4aad-a175-dabac48acb05`). Each row is one OBO token exchange —
a real tool call, not a sign-in page.

| Window | Exchanges | Failures | Users |
|---|---|---|---|
| Pilot (retained portion, 08-11 → rollout) | 28 | **0** | Bud 26, Kelly 2 |
| Post-rollout (08-15 → 08-18) | 16 | **0** | Bud 14, Dan 2 |
| **Total** | **44** | **0** | 3 distinct people |

**The pilot figure is a floor, not a total.** Entra retains sign-in logs for 7
days on the Free tier, so 2026-07-18 → 08-10 — most of the pilot — is simply
gone. What is retained shows zero failures across every exchange.

Two things this establishes that the closure record could not:

- **Real usage, not just assignment.** Bud and Kelly were actively using the
  connector during the pilot. The distinction the original text drew — holding
  an assignment versus using the thing — resolves in favour of use.
- **Group-based access works end to end.** Dan appears only after the rollout.
  He was never individually assigned; his access came entirely through the
  security group, which is [F1](../brainstorms/2026-08-13-exec-team-connector-rollout-requirements.md)
  demonstrated against the tenant rather than assumed.

**How to re-run this** — note the log, not just the query. OBO exchanges are
**non-interactive** sign-ins, and the default `/auditLogs/signIns` endpoint
returns only interactive ones. Querying the wrong log returns an empty result
that looks exactly like "nobody used it":

```bash
az rest --method GET --url "https://graph.microsoft.com/beta/auditLogs/signIns?\$filter=appId eq '484c0657-6a05-4aad-a175-dabac48acb05' and signInEventTypes/any(t: t eq 'nonInteractiveUser') and createdDateTime ge 2026-08-11T00:00:00Z&\$top=500" \
  --query "value[].{time:createdDateTime,user:userPrincipalName,code:status.errorCode}" -o json
```

The verdict below marks each criterion as **evidence** (satisfied by observation)
or **accepted** (judged acceptable without the evidence the pilot was meant to
produce). The distinction matters more than the pass/fail — an accepted criterion
is a live bet, not a closed question.

| # | Criterion | Verdict | Basis |
|---|-----------|---------|-------|
| 1 | Handshake is reliable | **Evidence** | Bud onboarded from the user guide without hand-holding. One user is thin proof, but it is the criterion's actual bar. |
| 2 | Auth is clean | **Evidence, since quantified** | No unexplained auth failures over the window; every denial reviewed was expected (unassigned accounts). Quantified 2026-08-18: **0 failures in 44 OBO exchanges**. |
| 3 | Throttling is tolerable | **Accepted — now with early data** | Revised 2026-08-18. Originally closed as *not measured*; [Measured usage](#measured-usage) now shows **44 OBO exchanges, zero failures** across pilot and post-rollout, so the load is no longer unknown. It is still small — one heavy user (Bud, 40 of 44) and two light ones — so this is early evidence of no throttling at observed volume, not proof the shared registration holds at ten active users. Going 3 → 10 remains a bet, but a watched one with a baseline rather than a blind one. (Ten: the nine group members plus the operator, who holds a separate break-glass assignment.) |
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

**Throttling is still the open question, but no longer an unmeasured one.**
Criterion 3 originally closed on judgment; [Measured usage](#measured-usage) now
gives a baseline of 44 exchanges with zero failures. That is reassuring at
observed volume and says nothing about ten *active* users — most of the traffic
is one person. A rising rate of "service busy" reports, or any user regularly
blocked, is still a **hold** signal: it means the shared-app-registration
decision needs revisiting before access widens past the ELT, not that exec
access should be rolled back.

User reports are no longer the *only* instrumentation. Re-run the sign-in-log
query above periodically — it gives exchange volume and failure counts per user
without waiting for someone to complain. The user guide still asks people to
flag persistent throttling, which remains the faster signal for the subjective
"it feels slow" case that a zero-failure log will not show.

Second-order: the OBO client secret expires **2027-07-18** and its failure mode is
a total outage for everyone at once. That now affects nine people rather than two.
