/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * In-memory cache for Graph API responses.
 *
 * Provides TTL-based caching to reduce API calls and improve performance.
 */

/**
 * Cache entry with expiration time.
 */
interface CacheEntry<T> {
  readonly value: T;
  readonly expiresAt: number;
}

/**
 * Default TTL values in milliseconds.
 */
export const CacheTTL = {
  /** Folders are relatively static */
  FOLDERS: 60_000, // 60 seconds
  /** Emails change more frequently */
  EMAILS: 30_000, // 30 seconds
  /** Events are fairly static */
  EVENTS: 30_000, // 30 seconds
  /** Contacts rarely change */
  CONTACTS: 120_000, // 2 minutes
  /** Tasks change frequently */
  TASKS: 30_000, // 30 seconds
} as const;

/**
 * Simple in-memory cache with TTL support.
 */
export class ResponseCache {
  private readonly cache = new Map<string, CacheEntry<unknown>>();
  private readonly defaultTtl: number;

  constructor(defaultTtl: number = CacheTTL.EMAILS) {
    this.defaultTtl = defaultTtl;
  }

  /**
   * Gets a cached value if it exists and hasn't expired.
   */
  get<T>(key: string): T | undefined {
    const entry = this.cache.get(key);

    if (entry == null) {
      return undefined;
    }

    if (Date.now() > entry.expiresAt) {
      this.cache.delete(key);
      return undefined;
    }

    return entry.value as T;
  }

  /**
   * Sets a cached value with optional TTL.
   */
  set<T>(key: string, value: T, ttl: number = this.defaultTtl): void {
    this.cache.set(key, {
      value,
      expiresAt: Date.now() + ttl,
    });
  }

  /**
   * Checks if a key exists and hasn't expired.
   */
  has(key: string): boolean {
    return this.get(key) !== undefined;
  }

  /**
   * Removes a specific key from the cache.
   */
  delete(key: string): boolean {
    return this.cache.delete(key);
  }

  /**
   * Clears all entries from the cache.
   */
  clear(): void {
    this.cache.clear();
  }

  /**
   * Removes all expired entries from the cache.
   */
  prune(): number {
    const now = Date.now();
    let count = 0;

    for (const [key, entry] of this.cache.entries()) {
      if (now > entry.expiresAt) {
        this.cache.delete(key);
        count++;
      }
    }

    return count;
  }

  /**
   * Gets the current number of entries in the cache.
   */
  get size(): number {
    return this.cache.size;
  }
}

/**
 * Creates a cache key from a method and parameters.
 */
export function createCacheKey(method: string, ...params: unknown[]): string {
  const paramStr = params
    .map((p) => (p === undefined ? '' : JSON.stringify(p)))
    .join(':');
  return `${method}:${paramStr}`;
}

/**
 * Invalidates cache entries matching a prefix.
 */
export function invalidateByPrefix(cache: ResponseCache, _prefix: string): void {
  // Note: This requires access to internal cache structure
  // For now, just clear the whole cache when invalidating
  cache.clear();
}
