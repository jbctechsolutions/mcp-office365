---
title: Alias-backed composite durable-ID pattern (self-encoding vs alias-backed tokens)
date: 2026-07-09
category: architecture-patterns
module: ids
problem_type: architecture_pattern
component: service_object
severity: high
related_components: [database, graph-repository]
applies_when:
  - Exposing third-party resource identifiers through a stable, user-facing token
  - A resource is addressed by a tuple (parent + child ids) rather than a single id
  - Graph ids are too long, ugly, or volatile to hand out directly
tags: [durable-ids, tokens, alias-table, composite-keys, sqlite, graph-api]
---

# Alias-backed composite durable-ID pattern

## Context

The v4.0.0 program replaced legacy numeric hash IDs with durable prefixed tokens (`em_ ev_ ct_ fd_ tm_ cn_ ...`) across ~19 Office 365 entity types. Some entities carry their Graph id cheaply; others are addressed by a *tuple* (e.g. a channel is `teamId + channelId`, a checklist item is `taskId + itemId`) and their raw ids are long and unfriendly. We needed one token scheme that handled both without leaking Graph internals to callers.

## Guidance

Split tokens into two families and pick per entity:

**Self-encoding tokens** (`em_ ev_ ct_ fd_ dr_ nb_ ns_ np_`) carry the base64url-encoded Graph id *inside the token*. They are "cold-durable": resolvable with no store, on any machine, forever. Use these when a single Graph id fully addresses the resource and it's acceptable to embed it.

**Alias-backed tokens** (`tm_ cn_ xm_ ch_ cm_ td_ tl_ ...`) are a short 70-bit digest that indexes a row in a local SQLite alias table holding the real Graph id (or tuple). Use these when the id is composite, or too long/volatile to embed. Tradeoff: alias tokens are **account-scoped and machine-scoped** — they only resolve where the alias store lives. This is the deliberate cost of a short, opaque, tuple-capable token.

Composite entities store the tuple as JSON in the alias `graphId` column and resolve back through a single generic helper:

```ts
// mint: store the tuple JSON-encoded under one alias
mintAliasComposite(entityType, parts)   // graphId: JSON.stringify(parts)

// resolve: validate EVERY key is present and non-empty, else IdUnknownError
toGraphParts<K extends string>(id, entityType, keys: readonly K[]): Record<K, string>
```

The generic `toGraphParts` means one code path serves every composite entity — add a new composite type by choosing a prefix and passing its key list, not by writing another bespoke decoder.

## Why This Matters

Handing out raw Graph ids couples every caller to Graph's id format and makes composite addressing (tuples) leak into tool schemas. Two token families with an explicit self-vs-alias decision keep the *durability tradeoff visible at design time*: you consciously choose "cold-durable but embeds the id" vs "short and opaque but store-bound." The generic composite helper prevents the pattern from degrading into N hand-rolled tuple parsers as entities are added — the single biggest source of drift in a wide migration.

The non-obvious constraint to document loudly: **alias tokens do not survive a machine/store change.** Anything that persists a token across environments (approval tokens, saved links) must account for this or use a self-encoding form.

## When to Apply

- The resource is addressed by more than one id → alias-backed composite.
- The single id is short and stable and embedding it is fine → self-encoding.
- You need tokens to resolve with zero local state (cross-machine, cold) → self-encoding only.
- You're about to write a second bespoke "parse this tuple id" function → route it through the generic composite helper instead.

## Examples

**Self-encoding** — an email token embeds its Graph id; `parseToken` + base64url-decode yields the id with no DB hit.

**Alias-backed composite** — a channel token `cn_<digest>` indexes `{ teamId, channelId }` stored as JSON; `toGraphParts('cn_...', 'channel', ['teamId','channelId'] as const)` returns both, throwing `IdUnknownError` if either is empty (guards against a half-minted alias silently producing `/teams//channels/...`).

## Related
- [[fetch-before-update-for-mutable-etags]] — companion concurrency pattern in the same repository layer
- [[test-external-api-assumptions-before-building-defenses]] — a discipline learned mid-program
- [[adversarial-review-as-primary-gate]] — the review pattern that guarded empty-id and half-mint bugs here
