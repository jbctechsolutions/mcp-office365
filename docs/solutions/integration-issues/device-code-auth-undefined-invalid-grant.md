---
title: "Device-code auth shows undefined code + invalid_grant (dead/mismatched Azure app)"
date: 2026-07-11
category: docs/solutions/integration-issues
module: graph/auth
problem_type: integration_issue
component: authentication
symptoms:
  - "`auth` prints `undefined` for the sign-in URL and code"
  - "Authentication fails with `post_request_failed: invalid_grant`"
  - "No obvious error explaining which of the client ID / tenant / app config is wrong"
root_cause: config_error
resolution_type: config_change
severity: high
tags: [msal, device-code, azure-ad, aadsts50059, aadsts700016, single-tenant, invalid-grant]
---

# Device-code auth shows undefined code + invalid_grant (dead/mismatched Azure app)

## Problem
`npx @jbctechsolutions/mcp-office365 auth` printed `undefined` for the sign-in URL and the code, then failed with `post_request_failed: invalid_grant`. The embedded default client ID no longer existed, and even a valid single-tenant app cannot complete device-code sign-in against the `common`/`organizations` authorities.

## Symptoms
- `To sign in … open the page: undefined` and `enter the code: undefined`
- `Authentication failed: post_request_failed: invalid_grant`
- The failure looks like a token/refresh problem but is actually the very first device-code request.

## What Didn't Work
- Reading `device-code-flow.ts` — the callback (`response.userCode`, `response.verificationUri`) is correct; the response object itself was empty `{}`.
- Blaming the msal-node `^2 → 5` major bump — reproduced the empty `{}` on **both** versions, so it wasn't a library regression.
- Trusting the sandbox repro — the default Bash sandbox blocked outbound network, producing a *false* empty `{}` that happened to match the real symptom. Had to re-run with the sandbox disabled to get real Azure responses.

## Solution
Isolate library-vs-service with a raw HTTP `POST` (no msal), then fix the config:

```bash
# Ground truth — what Azure returns for the /devicecode request:
curl -sS -X POST "https://login.microsoftonline.com/<tenant>/oauth2/v2.0/devicecode" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  --data-urlencode "client_id=<client-id>" \
  --data-urlencode "scope=User.Read offline_access"
# 400 AADSTS50059  -> single-tenant app requested via common/organizations (tenant can't be resolved)
# 400 AADSTS700016 -> app not registered in that tenant (wrong tenant, or dead/wrong client ID)
# 200 + user_code  -> correct client-id + tenant pairing
```

Fixes applied (PR #75):
- **Require the app registration** instead of embedding a dead default: `loadGraphConfig()` throws a clear setup message if `OUTLOOK_MCP_CLIENT_ID` is unset (tenant still defaults to `common`; `OUTLOOK_MCP_TENANT_ID` required for a single-tenant app).
- **Surface the real failure**: in the msal `deviceCodeCallback`, throw an actionable error when `userCode`/`verificationUri` are missing, instead of letting the flow print `undefined` and die polling with `invalid_grant`.
- **Don't authenticate for unknown tools**: reject an unregistered tool name (`registry.has(name)`) before `ensureInitialized()`, so a typo'd tool never triggers a device-code sign-in.

## Why This Works
msal-node **swallows a 400 from the `/devicecode` endpoint into an empty `{}` device-code response** and invokes the callback anyway — so `userCode`/`verificationUri` come back `undefined`, and the subsequent token poll (with an undefined `device_code`) returns `invalid_grant`. The underlying 400 is a config/registration mismatch: a **single-tenant app cannot resolve a tenant** from `common`/`organizations` for device-code flow (`AADSTS50059`), and a wrong/dead client ID yields `AADSTS700016`. Because the raw `curl` fails identically with no msal in the path, the root cause is provably Azure-side, not the client library.

## Prevention
- When a Microsoft/OAuth flow returns empty or `undefined` fields, hit the endpoint directly with `curl` before touching client code — it separates "the service rejected us" from "our parsing is wrong" in one request.
- A single-tenant Azure app (`signInAudience: AzureADMyOrg`) must use its **specific tenant ID** as the authority for device-code; `common`/`organizations` will 400 with `AADSTS50059`.
- Don't embed a single app registration in a shared/public tool and assume it works for everyone — device-code needs an app that belongs to the tenant being signed into. Require `OUTLOOK_MCP_CLIENT_ID`/`OUTLOOK_MCP_TENANT_ID` and fail fast with setup guidance.
- Beware the agent Bash sandbox for network repros: a blocked request can masquerade as an empty API response. Confirm reachability (or disable the sandbox) before drawing conclusions.

## Related Issues
- PR #75 — require `OUTLOOK_MCP_CLIENT_ID`, empty-device-code diagnostic, unknown-tool auth-skip
- `docs/solutions/best-practices/test-external-api-assumptions-before-building-defenses.md` — same "verify the external service's real behavior first" discipline
