---
title: Connecting a claude.ai custom connector to an Entra-protected remote MCP server
date: 2026-07-17
category: integration-issues
module: remote-connector
problem_type: integration_issue
component: oauth_auth
severity: high
applies_when:
  - Adding a remote MCP server to claude.ai as a custom connector, with Microsoft Entra ID as the OAuth authorization server
  - "'Add' spins and does nothing, or errors with 'Automatic client registration isn't supported', or Entra returns AADSTS9010010 / AADSTS650053"
  - The MCP server is a resource server that validates Entra JWTs (RFC 9728 PRM + 401 handshake) and points its PRM authorization_servers at login.microsoftonline.com directly (no OAuth proxy)
tags: [claude-ai, entra-id, azure-ad, oauth, mcp, rfc9728, rfc8707, aadsts9010010, dcr, cors, connector]
---

## Problem

A remote MCP server (Entra-protected, "Approach A" — Entra directly as the OAuth
authorization server, no proxy) would not connect as a claude.ai custom
connector. It went through **five distinct failure modes** before working. Each
looked like a dead end; the direct Entra flow *does* work, but only with all
five fixes together. The MCP auth spec assumes DCR + a well-behaved AS; Entra
satisfies neither cleanly, so the resource server and the app registration must
compensate.

## Symptoms (in the order they surface)

1. **"Add" spins quickly, then nothing changes.** No error, no sign-in popup.
2. **"Automatic client registration isn't supported by <server>. Edit the
   connector and add an OAuth Client ID."** (`ofid_…` reference.)
3. **"Authorization with <server> failed. Check your credentials and
   permissions."** with an Entra Trace ID → sign-in log shows **AADSTS9010010:
   "The resource parameter provided in the request doesn't match with the
   requested scopes."**
4. After the resource fix, briefly **AADSTS650053: "The application asked for
   scope '…' that doesn't exist on the resource '…'."** (propagation/transient.)

## What didn't work / red herrings

- **Assuming a proxy was mandatory.** Much community writing says Entra needs an
  OAuth proxy (DCR shim / APIM) because it lacks Dynamic Client Registration.
  True that Entra has no `registration_endpoint`, but claude.ai's **Advanced
  settings let you supply a Client ID manually**, which bypasses DCR. No proxy
  was needed for a single Entra app we control.
- **Registering the MCP URL as an identifier URI alone** did not fix
  AADSTS9010010 — the *scope* also has to be fully qualified (see fix 5).
- **Reading the Entra sign-in logs** requires a directory role (Security/Reports
  Reader). A plain app-owner account gets `Authentication_RequestFromUnsupportedUserRole`
  and empty results — use an admin account.
- Granting admin consent for the client→API scope was **redundant** here — an
  `azuread_application_pre_authorized` on the API app already covers it. The real
  blocker was never consent; it was resource/scope shape.

## Solution — the five fixes (all required)

Let `RESOURCE = https://<host>/mcp` (the connector URL you enter in claude.ai).

1. **Supply the Client ID manually.** claude.ai → connector → Advanced settings →
   **OAuth Client ID** = the public client app's id. (Entra has no DCR;
   auto-registration fails otherwise.) No secret needed for a public/PKCE client.

2. **Register `RESOURCE` as an identifier URI on the API (resource) app.**
   claude.ai sends `RESOURCE` as the RFC 8707 `resource` parameter; Entra rejects
   the token request (AADSTS9010010) unless `resource` matches a registered
   Application ID URI. `az ad app update --id <api-app> --identifier-uris
   "api://…" "https://<host>/mcp"`. (An `*.azurecontainerapps.io` URL was accepted
   as an identifier URI in our tenant; not all tenants allow arbitrary https.)

3. **Fix the RFC 9728 `WWW-Authenticate` discovery URL.** The metadata URL is
   `/.well-known/oauth-protected-resource` inserted **between host and path**, NOT
   appended to the full resource URL:
   - correct → `https://<host>/.well-known/oauth-protected-resource/mcp` (serve + advertise this)
   - wrong → `https://<host>/mcp/.well-known/oauth-protected-resource` (404 → discovery dies → "Add does nothing")

   ```ts
   // resource_metadata in the 401 WWW-Authenticate header:
   const u = new URL(config.publicUrl);           // https://host/mcp
   const path = u.pathname === '/' ? '' : u.pathname;
   return `${u.origin}${PRM_PATH}${path}`;         // https://host/.well-known/oauth-protected-resource/mcp
   ```

4. **Serve CORS.** claude.ai reads discovery metadata / the 401 challenge from the
   browser. Send `Access-Control-Allow-Origin: *` (Authorization is a header, not a
   cookie, so wildcard is safe), **expose `WWW-Authenticate`**, allow the MCP
   headers, and answer `OPTIONS` with 204.

5. **Advertise a FULLY-QUALIFIED scope in the PRM, and accept the resource as an
   audience.** This is the actual AADSTS9010010 fix. Entra v2 requires the app
   referenced by `scope` to have an identifier URI equal to `resource`. A bare
   `scopes_supported: ["access_as_user"]` doesn't tie to the resource →
   AADSTS9010010. Advertise `["<RESOURCE>/access_as_user"]` instead:

   ```ts
   scopes_supported: [`${config.publicUrl}/${config.requiredScope}`],
   ```

   The token's `scp` claim still comes back as the **bare** `access_as_user`, so
   token verification is unchanged. But Entra sets `aud` to `RESOURCE`, so add
   `RESOURCE` to the server's accepted audiences:

   ```ts
   allowedAudiences: [apiClientId, appIdUri, publicUrl],
   ```

## Why this works

- claude.ai's connector OAuth is auth-code + PKCE against the AS in the PRM. It
  fetches the PRM (RFC 9728), reads `authorization_servers`, then the AS metadata
  (`.../.well-known/openid-configuration`), then tries DCR — which Entra lacks, so
  a **manual Client ID** is required (fix 1).
- Entra v2 is strict about the RFC 8707 `resource` param: `resource` must be a
  registered identifier URI (fix 2) AND the requested `scope` must reference an
  app whose identifier URI equals `resource` (fix 5). A bare scope name doesn't
  qualify.
- Discovery is browser-mediated and follows the `WWW-Authenticate` hint, so a
  malformed metadata URL (fix 3) or missing CORS (fix 4) kills "Add" silently
  before any of the OAuth logic runs.

## Prevention

- **Advertise fully-qualified scopes** in the PRM from the start (`<resource>/<scope>`),
  and keep verification on the bare `scp` name.
- **Build the RFC 9728 metadata URL by host+well-known+path**, never
  `fullUrl + well-known`. Add a test asserting the header points at the served
  200 location, not a 404.
- **Register the connector URL as an API identifier URI** (Terraform: a
  `mcp_public_url` variable folded into `identifier_uris` via `compact()`), and
  **accept it as a token audience**.
- **Ship CORS** on any MCP server intended for claude.ai (wildcard origin, expose
  `WWW-Authenticate`, 204 on OPTIONS).
- The direct-Entra approach is viable for a single app you control — an OAuth
  proxy is only forced when the resource is a Microsoft first-party app whose
  identifier URI you can't change (e.g. Azure DevOps; see
  anthropics/claude-code#55993).
- Use an **admin account** for Entra sign-in logs / consent / app edits; the CI
  service principal for Azure RBAC (Key Vault, Container App). They are different
  privilege planes.
