/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * MSAL token cache plugin for persistent token storage.
 *
 * Stores tokens in ~/.mcp-office365/tokens.json for persistence
 * across application restarts.
 */

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import type { ICachePlugin, TokenCacheContext } from '@azure/msal-node';

/**
 * Directory where tokens are stored.
 */
const TOKEN_CACHE_DIR = join(homedir(), '.mcp-office365');

/**
 * Path to the token cache file.
 */
const TOKEN_CACHE_FILE = join(TOKEN_CACHE_DIR, 'tokens.json');

/**
 * Ensures the token cache directory exists.
 */
function ensureCacheDir(): void {
  if (!existsSync(TOKEN_CACHE_DIR)) {
    mkdirSync(TOKEN_CACHE_DIR, { recursive: true, mode: 0o700 });
  }
}

/**
 * MSAL cache plugin implementation that persists tokens to disk.
 */
export class FileTokenCachePlugin implements ICachePlugin {
  /**
   * Called by MSAL before accessing the cache.
   * Loads the cache from disk into MSAL's in-memory cache.
   */
  // eslint-disable-next-line @typescript-eslint/require-await
  async beforeCacheAccess(context: TokenCacheContext): Promise<void> {
    try {
      if (existsSync(TOKEN_CACHE_FILE)) {
        const data = readFileSync(TOKEN_CACHE_FILE, 'utf-8');
        context.tokenCache.deserialize(data);
      }
    } catch {
      // If we can't read the cache, start fresh
    }
  }

  /**
   * Called by MSAL after modifying the cache.
   * Persists the cache to disk.
   */
  // eslint-disable-next-line @typescript-eslint/require-await
  async afterCacheAccess(context: TokenCacheContext): Promise<void> {
    if (context.cacheHasChanged) {
      try {
        ensureCacheDir();
        const data = context.tokenCache.serialize();
        writeFileSync(TOKEN_CACHE_FILE, data, { mode: 0o600 });
      } catch {
        // If we can't write, tokens won't persist (user will need to re-auth)
      }
    }
  }
}

/**
 * Creates a new file-based token cache plugin.
 */
export function createTokenCachePlugin(): ICachePlugin {
  return new FileTokenCachePlugin();
}

/**
 * Checks if a token cache file exists.
 */
export function hasTokenCache(): boolean {
  return existsSync(TOKEN_CACHE_FILE);
}

/**
 * Clears the token cache file.
 */
export function clearTokenCache(): void {
  try {
    if (existsSync(TOKEN_CACHE_FILE)) {
      writeFileSync(TOKEN_CACHE_FILE, '{}', { mode: 0o600 });
    }
  } catch {
    // Ignore errors
  }
}

/**
 * Gets the token cache directory path.
 */
export function getTokenCacheDir(): string {
  return TOKEN_CACHE_DIR;
}

/**
 * Gets the token cache file path.
 */
export function getTokenCacheFile(): string {
  return TOKEN_CACHE_FILE;
}
