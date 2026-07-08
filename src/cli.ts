/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * CLI command handlers for standalone authentication management.
 *
 * Usage:
 *   npx @jbctechsolutions/mcp-office365 auth            # Authenticate
 *   npx @jbctechsolutions/mcp-office365 auth --status   # Check status
 *   npx @jbctechsolutions/mcp-office365 auth --force    # Re-authenticate (clears existing tokens)
 *   npx @jbctechsolutions/mcp-office365 auth --logout   # Sign out
 */

import {
  getAccessToken,
  isAuthenticated,
  getAccount,
  signOut,
  getTokenCacheFile,
} from './graph/index.js';
import type { Preset } from './registry/types.js';

export interface CliCommand {
  command: 'auth';
  flags: string[];
}

/** Valid `--preset` names, mirroring the domain modules. */
export const VALID_PRESETS: readonly Preset[] = [
  'mail',
  'calendar',
  'contacts',
  'tasks',
  'notes',
  'teams',
  'planner',
  'files',
  'sharepoint',
  'excel',
  'people',
  'meetings',
];

/** Server-mode CLI options parsed from argv (U10). */
export interface ServerCliOptions {
  /** Presets to include; omitted means the full surface (`all`). */
  readonly presets?: readonly Preset[];
  /** When true, only non-destructive tools are exposed. */
  readonly readOnly: boolean;
}

/**
 * Parses server-mode flags (`--preset <names>`, `--read-only`) from argv.
 * `--preset` accepts a comma-separated list and may repeat; `all` (or no
 * `--preset`) means the full surface. Unknown preset names throw with the valid
 * list so startup fails loudly rather than silently exposing nothing.
 */
export function parseServerOptions(args: string[]): ServerCliOptions {
  const presetNames: string[] = [];
  let readOnly = false;

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg == null) continue;
    if (arg === '--read-only') {
      readOnly = true;
    } else if (arg === '--preset') {
      const value = args[i + 1];
      if (value == null || value.startsWith('--')) {
        throw missingPresetValue();
      }
      presetNames.push(...requirePresetNames(value));
      i++;
    } else if (arg.startsWith('--preset=')) {
      presetNames.push(...requirePresetNames(arg.slice('--preset='.length)));
    }
    // Unknown args are ignored — argv may carry runner-injected entries.
  }

  // No --preset flag at all → full surface.
  if (presetNames.length === 0) {
    return { readOnly };
  }

  // Validate every name (the `all` keyword excepted) BEFORE honoring `all`, so a
  // typo mixed with `all` (e.g. `--preset all,mial`) fails loudly rather than
  // silently exposing the full surface.
  const invalid = presetNames.filter(
    (name) => name !== 'all' && !(VALID_PRESETS as readonly string[]).includes(name),
  );
  if (invalid.length > 0) {
    throw new Error(
      `Unknown preset(s): ${invalid.join(', ')}. ` +
        `Valid presets: ${VALID_PRESETS.join(', ')}, all.`,
    );
  }

  // `all` (alone or mixed with valid names) exposes the full surface.
  if (presetNames.includes('all')) {
    return { readOnly };
  }

  return { readOnly, presets: presetNames as Preset[] };
}

function missingPresetValue(): Error {
  return new Error('--preset requires a comma-separated list of preset names.');
}

/** Splits a preset value and rejects an empty/whitespace-only list. */
function requirePresetNames(value: string): string[] {
  const names = splitPresetList(value);
  if (names.length === 0) {
    throw missingPresetValue();
  }
  return names;
}

function splitPresetList(value: string): string[] {
  return value
    .split(',')
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
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

  if (flags.includes('--force')) {
    return await handleForceAuth(print);
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

async function handleForceAuth(print: PrintFn): Promise<number> {
  print('');
  print('Microsoft Graph API Authentication (force re-auth)');
  print('='.repeat(40));
  print('');

  try {
    await signOut();
    print('Cleared existing tokens.');
    print('');
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
    print('Run: npx @jbctechsolutions/mcp-office365 auth');
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
