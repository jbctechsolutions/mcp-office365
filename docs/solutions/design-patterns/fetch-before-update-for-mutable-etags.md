---
title: Fetch-before-update for mutable per-sub-resource ETags (don't cache the ETag)
date: 2026-07-09
category: design-patterns
module: graph-repository
problem_type: design_pattern
component: service_object
severity: medium
related_components: [graph-api, planner]
applies_when:
  - An API requires If-Match with a current ETag on PATCH/DELETE (optimistic concurrency)
  - The ETag changes on every write and you cannot safely hold a stale copy
  - You considered caching the ETag in a token or record and reusing it later
tags: [etag, optimistic-concurrency, if-match, graph-api, planner, 412]
---

# Fetch-before-update for mutable per-sub-resource ETags

## Context

Microsoft Graph Planner writes require `If-Match: <etag>` for optimistic concurrency, and the ETag rotates on every successful write. An early instinct was to capture the ETag at read time and stash it (e.g. in the operation's approval token) so the later write wouldn't need an extra round trip. That is unsafe: any intervening change — or the entity's own prior write in a multi-step flow — invalidates the cached ETag, and a stale or empty `If-Match` either 412s or, worse, silently sends `If-Match: ''`.

## Guidance

Do not cache mutable ETags. **Fetch the current ETag immediately before the write, inside the same operation**, and retry once on a 412 with a freshly re-fetched ETag:

```ts
withFreshEtag(fetchEtag, write)
// 1. GET the resource -> current etag
// 2. guard: throw if etag is empty (never send If-Match: '')
// 3. write with If-Match: etag
// 4. on 412 (precondition failed): re-fetch etag once, retry the write
```

The empty-ETag guard is not optional: an unverified assumption that the fetch always yields a non-empty ETag can otherwise put `If-Match: ''` on the wire, which some backends treat very differently from "no If-Match at all."

Semantics this buys you: **last-writer-wins.** The one 412 retry re-reads the latest state and overwrites — document that explicitly so callers know concurrent edits are not merged, they are clobbered by whoever writes last.

## Why This Matters

A cached ETag is a correctness bug waiting for a race. It looks like a harmless optimization (saves one GET) but trades a guaranteed extra round trip for an intermittent, hard-to-reproduce 412 (or silent wrong write) under exactly the conditions — concurrency, multi-step flows — where the concurrency control was supposed to protect you. Fetch-before-update makes the write self-consistent regardless of what happened since the caller last read the entity, and the single retry absorbs the benign race where the entity changed between our own fetch and write.

## When to Apply

- Any PATCH/DELETE that requires `If-Match` and where the ETag mutates per write.
- Multi-step operations that write the same entity more than once (the second write must re-fetch).
- Anywhere you were tempted to persist an ETag beyond the lifetime of a single write.

Do **not** reach for this when the API uses a stable version token that only changes on the write you control, or when the platform offers a true merge/patch that doesn't need client-supplied preconditions.

## Examples

**Anti-pattern:** read the Planner task, store its `@odata.etag` in the approval token, later PATCH with the stored value → 412 whenever anything touched the task in between, including a prior step of the same flow.

**Pattern:** `withFreshEtag(() => getTaskEtag(id), etag => patchTask(id, body, etag))` — GET the etag now, guard non-empty, PATCH with `If-Match`, and on 412 re-GET once and retry. Last write wins.

## Related
- [[alias-backed-composite-durable-id-pattern]] — companion pattern in the same repository layer
- [[test-external-api-assumptions-before-building-defenses]] — the empty-ETag guard exists because the "fetch always returns an etag" assumption was not verified
