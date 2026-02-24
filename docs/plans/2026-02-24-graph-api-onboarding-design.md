# Graph API Onboarding Flow Design

## Problem

When a first-time user runs `USE_GRAPH_API=1 npx -y @jbctechsolutions/mcp-outlook-mac`, the server starts silently (stdio is reserved for JSON-RPC). On the first tool call, `initializeGraphBackend()` checks `isAuthenticated()`, finds no cached tokens, and throws `GraphAuthRequiredError`. The device code flow is never triggered — the user sees an error referencing "device code flow" but has no way to initiate it.

## Solution

Two-pronged onboarding: inline auth on first MCP tool call + standalone CLI `auth` subcommand.

## Component 1: Inline Auth on First Tool Call

When `initializeGraphBackend()` detects no cached tokens, instead of throwing `GraphAuthRequiredError`, it calls `getAccessToken()` which triggers `acquireTokenInteractive()`. MSAL's device code flow blocks until the user completes auth in their browser.

### Flow

```
User calls any tool (e.g., list_emails)
  -> ensureInitialized()
  -> isAuthenticated() returns false
  -> Call getAccessToken() (triggers device code flow)
  -> MSAL calls deviceCodeCallback with code + URL
  -> Server logs to stderr (visible in some clients)
  -> MSAL blocks waiting for user to complete browser auth
  -> User visits URL, enters code, grants permissions
  -> Tokens cached to ~/.outlook-mcp/tokens.json
  -> initializeGraphBackend() completes normally
  -> Original tool call proceeds and returns results
```

### Timeout Handling

- MSAL's `acquireTokenByDeviceCode` has a built-in timeout (~15 minutes from Microsoft)
- If the user doesn't complete auth in time, MSAL throws and we return a clear error
- Some MCP clients may have shorter tool call timeouts — the CLI `auth` subcommand is the fallback

### Concurrency

A mutex/promise lock ensures only one auth flow runs at a time. If multiple tool calls arrive simultaneously while unauthenticated, the first triggers auth and subsequent calls wait for it to complete.

## Component 2: CLI `auth` Subcommand

Standalone authentication for power users or when MCP client timeouts are too short.

### Usage

```bash
# Authenticate
npx @jbctechsolutions/mcp-outlook-mac auth

# Check auth status
npx @jbctechsolutions/mcp-outlook-mac auth --status

# Sign out (clear tokens)
npx @jbctechsolutions/mcp-outlook-mac auth --logout
```

### Implementation

In `main()`, check `process.argv` for the `auth` subcommand before starting the MCP server. If found, run the CLI auth flow directly (stdout/stderr available since we're not in MCP mode).

### Output

```
$ npx @jbctechsolutions/mcp-outlook-mac auth

Microsoft Graph API Authentication
====================================
To sign in, open your browser to:
  https://microsoft.com/devicelogin

And enter the code:
  ABCDEF123

Waiting for authentication...

Authenticated as user@example.com
Tokens saved to ~/.outlook-mcp/tokens.json
You can now configure the MCP server in your client.
```

### Status Check

```
$ npx @jbctechsolutions/mcp-outlook-mac auth --status

Authenticated as user@example.com
Token cache: ~/.outlook-mcp/tokens.json
```

or:

```
Not authenticated
Run: npx @jbctechsolutions/mcp-outlook-mac auth
```

## Component 3: Error Handling

- **Token expiry / refresh failure**: `acquireTokenSilent()` returns null, `getAccessToken()` falls through to `acquireTokenInteractive()` automatically
- **MSAL device code timeout**: Catch and return "Authentication timed out. Please try again or run `npx @jbctechsolutions/mcp-outlook-mac auth`"
- **Network error**: Display "Could not reach Microsoft servers. Check your internet connection."
- **User cancels**: Display "Authentication cancelled."
- **Concurrent tool calls during auth**: Mutex/promise lock — first call triggers auth, others wait

## Files to Modify

- `src/index.ts` — Change `initializeGraphBackend()` to call `getAccessToken()` instead of throwing; add CLI arg parsing in `main()`
- `src/graph/auth/device-code-flow.ts` — No changes needed (already has `getAccessToken()`)
- `src/utils/errors.ts` — May remove or repurpose `GraphAuthRequiredError`

## Files to Add

- `src/cli.ts` — CLI auth subcommand handler

## Tests

- Unit tests for CLI arg parsing and routing
- Unit tests for auth mutex behavior
- Unit tests for inline auth flow in `initializeGraphBackend()`
