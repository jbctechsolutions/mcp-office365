# Graph API Onboarding Flow Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Enable first-time Graph API users to authenticate seamlessly via inline auth on first tool call, plus a standalone CLI `auth` subcommand.

**Architecture:** Modify `initializeGraphBackend()` to call `getAccessToken()` instead of throwing when unauthenticated. Add CLI arg parsing in `main()` that routes `auth` / `auth --status` / `auth --logout` to a new `src/cli.ts` module. Use a promise-based mutex to prevent concurrent auth flows.

**Tech Stack:** TypeScript, MSAL Node (`@azure/msal-node`), Vitest

---

### Task 1: Create CLI Auth Module

**Files:**
- Create: `src/cli.ts`
- Test: `tests/unit/cli.test.ts`

**Step 1: Write the failing tests**

Create `tests/unit/cli.test.ts`:

```typescript
/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock the graph auth module
const mockGetAccessToken = vi.fn();
const mockIsAuthenticated = vi.fn();
const mockGetAccount = vi.fn();
const mockSignOut = vi.fn();
const mockGetTokenCacheFile = vi.fn(() => '/home/user/.outlook-mcp/tokens.json');

vi.mock('../../src/graph/index.js', () => ({
  getAccessToken: mockGetAccessToken,
  isAuthenticated: mockIsAuthenticated,
  getAccount: mockGetAccount,
  signOut: mockSignOut,
  getTokenCacheFile: mockGetTokenCacheFile,
}));

import { handleAuthCommand } from '../../src/cli.js';

describe('CLI Auth', () => {
  beforeEach(() => {
    vi.resetAllMocks();
  });

  describe('auth (no flags)', () => {
    it('runs device code flow and prints success', async () => {
      mockGetAccessToken.mockResolvedValue('test-token');
      mockGetAccount.mockResolvedValue({ username: 'user@example.com' });

      const output: string[] = [];
      const exit = await handleAuthCommand([], (msg) => output.push(msg));

      expect(mockGetAccessToken).toHaveBeenCalledOnce();
      expect(exit).toBe(0);
      expect(output.some(line => line.includes('user@example.com'))).toBe(true);
    });

    it('returns exit code 1 on auth failure', async () => {
      mockGetAccessToken.mockRejectedValue(new Error('Auth failed'));

      const output: string[] = [];
      const exit = await handleAuthCommand([], (msg) => output.push(msg));

      expect(exit).toBe(1);
      expect(output.some(line => line.includes('failed'))).toBe(true);
    });
  });

  describe('auth --status', () => {
    it('prints authenticated status when tokens exist', async () => {
      mockIsAuthenticated.mockResolvedValue(true);
      mockGetAccount.mockResolvedValue({ username: 'user@example.com' });

      const output: string[] = [];
      const exit = await handleAuthCommand(['--status'], (msg) => output.push(msg));

      expect(exit).toBe(0);
      expect(output.some(line => line.includes('Authenticated'))).toBe(true);
      expect(output.some(line => line.includes('user@example.com'))).toBe(true);
    });

    it('prints not authenticated when no tokens', async () => {
      mockIsAuthenticated.mockResolvedValue(false);

      const output: string[] = [];
      const exit = await handleAuthCommand(['--status'], (msg) => output.push(msg));

      expect(exit).toBe(1);
      expect(output.some(line => line.includes('Not authenticated'))).toBe(true);
    });
  });

  describe('auth --logout', () => {
    it('signs out and prints confirmation', async () => {
      mockSignOut.mockResolvedValue(undefined);

      const output: string[] = [];
      const exit = await handleAuthCommand(['--logout'], (msg) => output.push(msg));

      expect(mockSignOut).toHaveBeenCalledOnce();
      expect(exit).toBe(0);
      expect(output.some(line => line.includes('Signed out'))).toBe(true);
    });

    it('returns exit code 1 on signout failure', async () => {
      mockSignOut.mockRejectedValue(new Error('Signout failed'));

      const output: string[] = [];
      const exit = await handleAuthCommand(['--logout'], (msg) => output.push(msg));

      expect(exit).toBe(1);
    });
  });
});
```

**Step 2: Run test to verify it fails**

Run: `npx vitest run tests/unit/cli.test.ts`
Expected: FAIL — `handleAuthCommand` not found

**Step 3: Write minimal implementation**

Create `src/cli.ts`:

```typescript
/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * CLI command handlers for standalone authentication management.
 *
 * Usage:
 *   npx @jbctechsolutions/mcp-outlook-mac auth           # Authenticate
 *   npx @jbctechsolutions/mcp-outlook-mac auth --status   # Check status
 *   npx @jbctechsolutions/mcp-outlook-mac auth --logout   # Sign out
 */

import {
  getAccessToken,
  isAuthenticated,
  getAccount,
  signOut,
  getTokenCacheFile,
} from './graph/index.js';

type PrintFn = (message: string) => void;

/**
 * Handles the `auth` CLI subcommand.
 *
 * @param flags - CLI flags after "auth" (e.g., ["--status"])
 * @param print - Output function (defaults to console.log)
 * @returns Exit code (0 = success, 1 = failure)
 */
export async function handleAuthCommand(
  flags: string[] = [],
  print: PrintFn = console.log,
): Promise<number> {
  if (flags.includes('--status')) {
    return await handleStatus(print);
  }

  if (flags.includes('--logout')) {
    return await handleLogout(print);
  }

  return await handleAuth(print);
}

async function handleAuth(print: PrintFn): Promise<number> {
  print('');
  print('Microsoft Graph API Authentication');
  print('='.repeat(40));
  print('');

  try {
    await getAccessToken();
    const account = await getAccount();
    const username = account?.username ?? 'unknown';

    print('');
    print(`Authenticated as ${username}`);
    print(`Tokens saved to ${getTokenCacheFile()}`);
    print('You can now configure the MCP server in your client.');
    return 0;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    print('');
    print(`Authentication failed: ${message}`);
    return 1;
  }
}

async function handleStatus(print: PrintFn): Promise<number> {
  try {
    const authenticated = await isAuthenticated();

    if (authenticated) {
      const account = await getAccount();
      const username = account?.username ?? 'unknown';
      print(`Authenticated as ${username}`);
      print(`Token cache: ${getTokenCacheFile()}`);
      return 0;
    }

    print('Not authenticated');
    print('Run: npx @jbctechsolutions/mcp-outlook-mac auth');
    return 1;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    print(`Error checking status: ${message}`);
    return 1;
  }
}

async function handleLogout(print: PrintFn): Promise<number> {
  try {
    await signOut();
    print('Signed out successfully');
    return 0;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    print(`Sign out failed: ${message}`);
    return 1;
  }
}
```

**Step 4: Run test to verify it passes**

Run: `npx vitest run tests/unit/cli.test.ts`
Expected: PASS — all 6 tests

**Step 5: Commit**

```bash
git add src/cli.ts tests/unit/cli.test.ts
git commit -m "feat: Add CLI auth subcommand module"
```

---

### Task 2: Wire CLI into Main Entry Point

**Files:**
- Modify: `src/index.ts:3376-3398` (main entry point)

**Step 1: Write the failing test**

Add to `tests/unit/cli.test.ts`:

```typescript
describe('parseCliCommand', () => {
  it('returns null for no args (MCP server mode)', () => {
    expect(parseCliCommand([])).toBeNull();
  });

  it('returns auth command with no flags', () => {
    expect(parseCliCommand(['auth'])).toEqual({ command: 'auth', flags: [] });
  });

  it('returns auth command with --status flag', () => {
    expect(parseCliCommand(['auth', '--status'])).toEqual({ command: 'auth', flags: ['--status'] });
  });

  it('returns auth command with --logout flag', () => {
    expect(parseCliCommand(['auth', '--logout'])).toEqual({ command: 'auth', flags: ['--logout'] });
  });

  it('returns null for unknown commands', () => {
    expect(parseCliCommand(['unknown'])).toBeNull();
  });
});
```

Update import to include `parseCliCommand`.

**Step 2: Run test to verify it fails**

Run: `npx vitest run tests/unit/cli.test.ts`
Expected: FAIL — `parseCliCommand` not found

**Step 3: Write minimal implementation**

Add to `src/cli.ts`:

```typescript
export interface CliCommand {
  command: 'auth';
  flags: string[];
}

/**
 * Parses CLI arguments to determine if a subcommand was invoked.
 * Returns null if no subcommand (normal MCP server mode).
 */
export function parseCliCommand(args: string[]): CliCommand | null {
  if (args.length === 0) return null;

  const command = args[0];
  if (command === 'auth') {
    return { command: 'auth', flags: args.slice(1) };
  }

  return null;
}
```

Then modify `src/index.ts` main entry point:

Change the `main()` function (lines 3376-3381):

```typescript
import { parseCliCommand, handleAuthCommand } from './cli.js';

async function main(): Promise<void> {
  // Check for CLI subcommands before starting MCP server
  const cliCommand = parseCliCommand(process.argv.slice(2));
  if (cliCommand != null) {
    const exitCode = await handleAuthCommand(cliCommand.flags);
    process.exit(exitCode);
  }

  const server = createServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
}
```

**Step 4: Run test to verify it passes**

Run: `npx vitest run tests/unit/cli.test.ts`
Expected: PASS — all 11 tests

**Step 5: Run full test suite**

Run: `npx vitest run`
Expected: All tests pass

**Step 6: Commit**

```bash
git add src/cli.ts src/index.ts tests/unit/cli.test.ts
git commit -m "feat: Wire CLI auth subcommand into main entry point"
```

---

### Task 3: Inline Auth on First Tool Call

**Files:**
- Modify: `src/index.ts:1609-1624` (`initializeGraphBackend`)
- Modify: `src/index.ts:43` (import)

**Step 1: Write the failing test**

This task modifies `initializeGraphBackend()` which is an inner function of `createServer()` — not directly testable in isolation. The behavior change is:
- Before: throws `GraphAuthRequiredError` when not authenticated
- After: calls `getAccessToken()` to trigger device code flow, then proceeds

The E2E/integration tests already cover the server flow. Add a targeted test to verify the auth flow is invoked by adding to `tests/integration/server.test.ts` (or creating a focused test if the existing file doesn't cover it). However, since `initializeGraphBackend` is a closure, we primarily verify it indirectly.

Instead, add a test for the auth mutex behavior to `tests/unit/cli.test.ts`:

```typescript
describe('createAuthMutex', () => {
  it('only runs the auth function once for concurrent calls', async () => {
    const authFn = vi.fn().mockResolvedValue(undefined);
    const mutex = createAuthMutex(authFn);

    // Call 3 times concurrently
    const results = await Promise.all([mutex(), mutex(), mutex()]);

    expect(authFn).toHaveBeenCalledOnce();
    expect(results).toEqual([undefined, undefined, undefined]);
  });

  it('propagates errors to all waiters', async () => {
    const authFn = vi.fn().mockRejectedValue(new Error('Auth failed'));
    const mutex = createAuthMutex(authFn);

    const results = await Promise.allSettled([mutex(), mutex(), mutex()]);

    expect(authFn).toHaveBeenCalledOnce();
    expect(results.every(r => r.status === 'rejected')).toBe(true);
  });

  it('allows retry after failure', async () => {
    const authFn = vi.fn()
      .mockRejectedValueOnce(new Error('First attempt failed'))
      .mockResolvedValueOnce(undefined);
    const mutex = createAuthMutex(authFn);

    // First call fails
    await expect(mutex()).rejects.toThrow('First attempt failed');

    // Second call succeeds (new attempt)
    await expect(mutex()).resolves.toBeUndefined();

    expect(authFn).toHaveBeenCalledTimes(2);
  });
});
```

**Step 2: Run test to verify it fails**

Run: `npx vitest run tests/unit/cli.test.ts`
Expected: FAIL — `createAuthMutex` not found

**Step 3: Write minimal implementation**

Add to `src/cli.ts`:

```typescript
/**
 * Creates a mutex that ensures only one auth flow runs at a time.
 * Concurrent callers wait for the in-progress auth to complete.
 * After completion (success or failure), the mutex resets for future calls.
 */
export function createAuthMutex<T>(fn: () => Promise<T>): () => Promise<T> {
  let pending: Promise<T> | null = null;

  return () => {
    if (pending != null) {
      return pending;
    }

    pending = fn().finally(() => {
      pending = null;
    });

    return pending;
  };
}
```

Then modify `src/index.ts`:

1. Update import on line 43 — add `getAccessToken` alongside `isAuthenticated`:

```typescript
import {
  createGraphRepository,
  createGraphContentReadersWithClient,
  isAuthenticated,
  getAccessToken,
  GraphMailboxAdapter,
  type GraphRepository,
  type GraphContentReaders,
} from './graph/index.js';
```

2. Add import for `createAuthMutex` near other imports:

```typescript
import { parseCliCommand, handleAuthCommand, createAuthMutex } from './cli.js';
```

3. Replace `initializeGraphBackend()` (lines 1609-1624):

```typescript
  /**
   * Initializes Graph API backend.
   * If not authenticated, triggers the device code flow inline.
   */
  const initializeGraphBackend = createAuthMutex(async (): Promise<void> => {
    // Try to authenticate if needed (triggers device code flow for first-time users)
    const authenticated = await isAuthenticated();
    if (!authenticated) {
      await getAccessToken();
    }

    graphRepository = createGraphRepository();
    graphContentReaders = createGraphContentReadersWithClient(graphRepository.getClient());

    const adapter = new GraphMailboxAdapter(graphRepository);
    orgTools = createMailboxOrganizationTools(adapter, tokenManager);
    sendTools = createMailSendTools(graphRepository, tokenManager);

    initialized = true;
  });
```

4. Update `ensureInitialized()` since `initializeGraphBackend` is now a callable, not an async function declaration — no change needed, it's already called with `await`.

5. Remove the `GraphAuthRequiredError` import (line 117) if it's no longer used elsewhere in this file.

**Step 4: Run test to verify it passes**

Run: `npx vitest run tests/unit/cli.test.ts`
Expected: PASS — all 14 tests

**Step 5: Run full test suite**

Run: `npx vitest run`
Expected: All tests pass. Verify TypeScript: `npx tsc --noEmit`. Verify lint: `npx eslint src/cli.ts`.

**Step 6: Commit**

```bash
git add src/index.ts src/cli.ts tests/unit/cli.test.ts
git commit -m "feat: Inline auth on first tool call with mutex"
```

---

### Task 4: Update Error Message and Documentation

**Files:**
- Modify: `src/utils/errors.ts:224-233` (update `GraphAuthRequiredError` message)
- Modify: `README.md` (add CLI auth docs)
- Modify: `tests/unit/utils/errors.test.ts` (update test if message changed)

**Step 1: Update error message**

`GraphAuthRequiredError` is still useful as a fallback for unexpected states. Update its message to reference the CLI:

```typescript
export class GraphAuthRequiredError extends OutlookMcpError {
  readonly code = ErrorCode.GRAPH_AUTH_REQUIRED;

  constructor() {
    super(
      'Microsoft Graph authentication required. ' +
        'Run: npx @jbctechsolutions/mcp-outlook-mac auth'
    );
  }
}
```

**Step 2: Update error test**

In `tests/unit/utils/errors.test.ts`, update the test that checks the error message to match the new text.

**Step 3: Update README**

Add to the Quick Start section for Graph API (after line 71 in current README):

```markdown
**Pre-authenticate (optional):**
```bash
npx -y @jbctechsolutions/mcp-outlook-mac auth
```

Or just configure the server — it will prompt for authentication on first use.
```

Add a new section under "Troubleshooting > Graph API Backend" before the "Authentication required" subsection:

```markdown
#### Pre-authentication

You can authenticate before configuring the MCP server:

\`\`\`bash
# Authenticate
npx @jbctechsolutions/mcp-outlook-mac auth

# Check status
npx @jbctechsolutions/mcp-outlook-mac auth --status

# Sign out
npx @jbctechsolutions/mcp-outlook-mac auth --logout
\`\`\`
```

**Step 4: Run tests**

Run: `npx vitest run`
Expected: All tests pass

**Step 5: Commit**

```bash
git add src/utils/errors.ts tests/unit/utils/errors.test.ts README.md
git commit -m "docs: Add CLI auth documentation and update error message"
```
