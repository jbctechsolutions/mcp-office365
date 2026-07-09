---
title: Test an external-API assumption empirically before engineering defenses around it
date: 2026-07-09
category: best-practices
module: graph-client
problem_type: best_practice
component: development_workflow
severity: high
applies_when:
  - A design hinges on an unverified belief about how a third-party API behaves
  - You are about to add retry, fallback, or skip-marker scaffolding to guard against a failure mode you have not observed
  - A read-only probe against the live service is cheap and available
tags: [graph-api, assumptions, immutable-id, yagni, verification, empiricism]
---

# Test an external-API assumption empirically before engineering defenses around it

## Context

While building U5b-3 (Outlook immutable IDs), the design rested on an unverified premise: that sending Microsoft Graph the `Prefer: IdType="ImmutableId"` header would cause requests carrying a *default-format* id in the URL to fail. Believing that, we built a defensive apparatus — a retry that stripped the header, a skip-marker to avoid re-issuing, and an extension of the same guard into every write path — to survive the failure we assumed would happen. Two review rounds went into hardening scaffolding for a failure mode nobody had observed.

## Guidance

When a design depends on a specific belief about how an external service behaves, and a read-only probe is cheap, **run the probe before writing a single line of defensive code.** One live call settles what could otherwise drive rounds of speculative engineering.

The probe that resolved U5b-3 (paraphrased) authenticated with the existing silent-token flow and issued raw GETs:

```js
// default id in URL, no header        -> 200
// default id in URL + Prefer:Immutable -> 200  (response id reshaped: len 152 -> 68)
// immutable id in URL, no header       -> 200
```

The header turned out to be **response-shaping only**: Graph accepts either id form in the request URL regardless of the header; the header only controls the *format of the id returned in the response body*. The premise was false. The entire retry / skip-marker / writes-extension apparatus was deleted, and the middleware collapsed to "add the header, read the reshaped id."

## Why This Matters

Defensive scaffolding built on an unverified premise is pure downside: it is code you must write, test, review, and maintain to defend against something that may not exist. When the premise is false, every hour spent hardening it is wasted, and the extra branches become permanent surface area for real bugs. A single empirical test converts a speculative argument (which can run for rounds, because neither side can prove the other wrong) into a settled fact. The corrective prompt in this session was literally *"why don't you test it?"* — and the test ended the debate in one call.

This is YAGNI applied to *external behavior*, not just features: don't build for a failure you haven't seen the API produce.

## When to Apply

- The blast radius of the assumption is more than a few lines (retries, fallbacks, new middleware branches, guard flags).
- A read-only or sandbox probe against the real service is available and low-risk.
- You catch yourself arguing about what the API *would* do rather than what it *does*.
- A stricter reading of a spec is driving you to add code — confirm the service actually enforces that reading.

## Examples

**Before (assumption-driven):** middleware issues the request with the header; on the anticipated failure it retries without the header, sets a per-URL skip-marker so it won't re-add the header, and the same fetch-retry guard is threaded through every write method "to be safe."

**After (probe-driven):** the live probe shows both id forms return 200 under the header. Middleware becomes: attach `Prefer: IdType="ImmutableId"`, read the (reshaped) immutable id from the response. No retry, no skip-marker, no writes extension.

## Related
- [[alias-backed-composite-durable-id-pattern]] — the durable-ID program this work was part of
- Deferred to v4.1: translate-on-resolve fallback + `degraded_ids` surfacing (documented in CHANGELOG / v4-release-status)
