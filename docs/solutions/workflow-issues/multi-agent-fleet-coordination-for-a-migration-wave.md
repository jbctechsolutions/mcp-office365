---
title: Coordinating a multi-agent fleet through a wide migration wave
date: 2026-07-09
category: workflow-issues
module: engineering-process
problem_type: workflow_issue
component: development_workflow
severity: medium
applies_when:
  - A migration touches many entities/files and could parallelize across several agents
  - Some work lands in a few hot shared files while other work is cleanly separable
  - Multiple agents may open PRs against the same base concurrently
tags: [multi-agent, fleet, migration, merge-sequencing, rebase, coordination]
---

# Coordinating a multi-agent fleet through a wide migration wave

## Context

The v4.0.0 durable-ID rollout spanned ~19 entity types and ~19 PRs. Naively fanning every entity out to its own agent looked attractive, but the migration had two very different kinds of work: changes concentrated in a handful of *hot* shared files (`repository.ts`, `ids/token.ts`, the e2e tool-count assertion) and *separable* additive features (SharePoint Lists, delta-sync, elicitation, OneNote). Treating them the same caused churn — one feature branched off a stale base and went CONFLICTING after two sibling PRs merged.

## Guidance

Partition the work by *file contention*, not by feature count:

1. **Single-thread the hot files.** Anything that edits `repository.ts`, the central id-type union, or the e2e tool-count assertion goes through one lane, sequentially. Parallel edits there guarantee conflicts and force serial rebasing anyway — so serialize up front.
2. **Parallelize the separable additive features.** New tool files, new modules, and new tests that don't touch the hot files run concurrently across agents with little risk.
3. **Sequence merges with an additive rebase.** Pick a single serialization point that every branch must reconcile against — here, the end-to-end **tool-count assertion**. Each branch rebases onto the latest main, updates the count to the new total, and merges; the next branch repeats. The assertion doubles as a tripwire: if two branches both think they added the "last" tool, the count won't reconcile and you catch the collision before merge.
4. **Give each agent its own worktree.** One agent worked directly in the lead's worktree and left uncommitted changes stranded; the recovery was a manual `git checkout -b` to rescue the diff. Isolated worktrees per agent prevent this entirely.
5. **Forbid subagents from touching shared side-effect state.** A research subagent wrote a premature "shipped" entry to the Cairn vault; it had to be `git reset --hard`'d. Subagent specs must explicitly bar vault/git/publish actions — those belong to the orchestrator.

## Why This Matters

The wall-clock win of a fleet comes entirely from the *separable* work; the hot-file work is latency-bound no matter how many agents you throw at it, because it must serialize. Recognizing that split up front avoids the worst outcome — several agents racing on the same file, each rebasing on the others, producing more coordination overhead than a single lane would have. A shared serialization point (the tool-count assertion) turns "did anyone collide?" from a hope into a mechanical check.

## When to Apply

- The change set is large enough that a single agent would be the bottleneck.
- You can cleanly separate "touches the shared core" from "adds new isolated surface."
- You have a natural invariant (a global count, a registry, a manifest) that every branch must update — use it as the merge serialization point.

Skip the fleet for changes that are mostly hot-file edits: the coordination tax exceeds the parallelism benefit, and one focused lane ships faster.

## Examples

**Contention-aware split:** id-union and repository edits → one sequential lane; SharePoint Lists / delta-sync / elicitation / OneNote → four parallel agents in separate worktrees.

**Additive-rebase merge:** branch A rebases onto main, bumps the e2e tool count to 240, merges; branch B rebases onto the new main, bumps to 241, merges. A branch that went stale (CONFLICTING after siblings merged) is recovered by rebasing and re-reconciling the count — not by re-cutting from an old base.

## Related
- [[adversarial-review-as-primary-gate]] — the gate every fleet PR passed through
- [[alias-backed-composite-durable-id-pattern]] — the technical pattern the fleet was rolling out
