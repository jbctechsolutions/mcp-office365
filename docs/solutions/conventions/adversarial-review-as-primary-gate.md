---
title: Adversarial review as the primary gate when the review bots are unavailable
date: 2026-07-09
category: conventions
module: engineering-process
problem_type: convention
component: development_workflow
severity: medium
applies_when:
  - CodeRabbit / Copilot / other automated reviewers are quota-limited or offline
  - A change is large or touches id-minting, schemas, URL matching, or upgrade boundaries
  - You need a substantive review pass, not a rubber stamp
tags: [code-review, adversarial-review, opus-review, migration, regression-guard]
---

# Adversarial review as the primary gate when the review bots are unavailable

## Context

For most of the v4.0.0 program the automated reviewers (CodeRabbit, Copilot) were quota-limited, so they could not be relied on as the review gate. Rather than merge on a light self-read, each change went through an explicit **adversarial** review pass — an Opus reviewer instructed to actively construct failure scenarios and try to break the diff, not to check it against a happy-path checklist. That pass caught real, shipping-blocking bugs the happy path missed.

## Guidance

When the bots are down (or the change is high-risk regardless), run a deliberate adversarial review before merge and treat it as the gate. The reviewer's job is to *find the input that breaks this*, and the recurring bug classes below are the checklist to hunt for — every one of them actually occurred in this program:

- **Missing empty-id guards in mappers.** A list mapper minted a token from an id that could be empty (`itemId.length > 0 ? mint : ''`). Half-minted ids silently produce malformed URLs like `/teams//channels/...`.
- **Cross-consumer schema breaks.** A parent id migrated `z.number()` → `z.string()` in one tool but a *different* consumer of the same id (a sub-entity tool, a list filter) was left on the old type, bricking that path. Grep every consumer of a changed id, not just the file you're editing.
- **Over-broad URL matching.** A middleware matched on the bare segment `'messages'`, so `/teams/.../messages` and `/me/chats/.../messages` both matched unintentionally. Anchor URL matchers on the *first collection after the user context*, not on a free-floating segment.
- **Upgrade-boundary type coercion.** Narrowing a persisted field `string|number` → `string` silently broke tokens written by the prior major version. Coerce at the persistence boundary (`String(target.targetId)`) so old rows stay usable across the upgrade.
- **Coverage gates that only one CI job runs.** A branch-coverage gate ran only on one OS matrix leg, so `npm test` locally passed while CI would have failed. Know which job enforces which gate.

## Why This Matters

Automated reviewers are a convenience, not a guarantee, and their availability is outside your control. A change that is "reviewed" only by its author reading the happy path will ship the failure modes above — several of which are silent (malformed URLs, wrong writes) rather than loud crashes, so tests that assert the happy path won't catch them. An adversarial pass that *tries to break the diff* is the difference between "looks fine" and "I found the input that makes it fail." In a wide migration, these bug classes recur across entities, so the same checklist pays off dozens of times.

## When to Apply

- Any time the automated reviewers are unavailable and you'd otherwise merge on a self-read.
- Large diffs (roughly ≥50 changed lines) or changes touching id-minting, schema types, URL/route matching, auth, or data-format upgrade boundaries — regardless of bot availability.

For a trivial one-line or comment-only change, the full adversarial pass is overkill; a self-read is fine.

## Examples

**What the pass caught (real):** the empty-id mapper guard (#47), the `task_id` sub-entity schema desync (#51), the `list_contacts` folder_id cross-consumer miss (#53), the Teams/chat `messages` over-match, and the empty-`If-Match` risk in Planner writes (#58).

**Framing that works:** instruct the reviewer to "construct concrete inputs that produce a wrong result or crash," not to "assess whether this looks correct." The first finds bugs; the second finds typos.

## Related
- [[multi-agent-fleet-coordination-for-a-migration-wave]] — every fleet PR passed through this gate
- [[test-external-api-assumptions-before-building-defenses]] — a related discipline: verify, don't assume
- [[fetch-before-update-for-mutable-etags]] — the empty-If-Match bug this gate caught
