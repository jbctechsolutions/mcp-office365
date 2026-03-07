/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * CLI command handlers for standalone authentication management.
 *
 * Usage:
 *   npx @jbctechsolutions/mcp-office365-mac auth           # Authenticate
 *   npx @jbctechsolutions/mcp-office365-mac auth --status   # Check status
 *   npx @jbctechsolutions/mcp-office365-mac auth --logout   # Sign out
 */

import {
  getAccessToken,
  isAuthenticated,
  getAccount,
  signOut,
  getTokenCacheFile,
} from './graph/index.js';

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
    print('Run: npx @jbctechsolutions/mcp-office365-mac auth');
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
