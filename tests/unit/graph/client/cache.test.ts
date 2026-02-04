/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Graph API response cache.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  ResponseCache,
  CacheTTL,
  createCacheKey,
  invalidateByPrefix,
} from '../../../../src/graph/client/cache.js';

describe('graph/client/cache', () => {
  describe('CacheTTL', () => {
    it('has expected TTL values', () => {
      expect(CacheTTL.FOLDERS).toBe(60_000);
      expect(CacheTTL.EMAILS).toBe(30_000);
      expect(CacheTTL.EVENTS).toBe(30_000);
      expect(CacheTTL.CONTACTS).toBe(120_000);
      expect(CacheTTL.TASKS).toBe(30_000);
    });
  });

  describe('ResponseCache', () => {
    let cache: ResponseCache;

    beforeEach(() => {
      cache = new ResponseCache();
      vi.useFakeTimers();
    });

    afterEach(() => {
      vi.useRealTimers();
    });

    describe('get/set', () => {
      it('stores and retrieves value', () => {
        cache.set('key1', { data: 'test' });
        expect(cache.get('key1')).toEqual({ data: 'test' });
      });

      it('returns undefined for missing key', () => {
        expect(cache.get('nonexistent')).toBeUndefined();
      });

      it('returns undefined after TTL expires', () => {
        cache.set('key1', 'value', 1000);
        expect(cache.get('key1')).toBe('value');

        vi.advanceTimersByTime(1001);
        expect(cache.get('key1')).toBeUndefined();
      });

      it('uses default TTL when not specified', () => {
        cache.set('key1', 'value');
        expect(cache.get('key1')).toBe('value');

        vi.advanceTimersByTime(CacheTTL.EMAILS + 1);
        expect(cache.get('key1')).toBeUndefined();
      });

      it('uses custom default TTL from constructor', () => {
        const customCache = new ResponseCache(5000);
        customCache.set('key1', 'value');
        expect(customCache.get('key1')).toBe('value');

        vi.advanceTimersByTime(4999);
        expect(customCache.get('key1')).toBe('value');

        vi.advanceTimersByTime(2);
        expect(customCache.get('key1')).toBeUndefined();
      });

      it('overwrites existing value', () => {
        cache.set('key1', 'value1');
        cache.set('key1', 'value2');
        expect(cache.get('key1')).toBe('value2');
      });

      it('stores complex objects', () => {
        const complex = {
          array: [1, 2, 3],
          nested: { a: 1, b: { c: 2 } },
        };
        cache.set('key1', complex);
        expect(cache.get('key1')).toEqual(complex);
      });
    });

    describe('has', () => {
      it('returns true for existing key', () => {
        cache.set('key1', 'value');
        expect(cache.has('key1')).toBe(true);
      });

      it('returns false for missing key', () => {
        expect(cache.has('nonexistent')).toBe(false);
      });

      it('returns false after TTL expires', () => {
        cache.set('key1', 'value', 1000);
        expect(cache.has('key1')).toBe(true);

        vi.advanceTimersByTime(1001);
        expect(cache.has('key1')).toBe(false);
      });
    });

    describe('delete', () => {
      it('removes existing key', () => {
        cache.set('key1', 'value');
        expect(cache.delete('key1')).toBe(true);
        expect(cache.get('key1')).toBeUndefined();
      });

      it('returns false for missing key', () => {
        expect(cache.delete('nonexistent')).toBe(false);
      });
    });

    describe('clear', () => {
      it('removes all entries', () => {
        cache.set('key1', 'value1');
        cache.set('key2', 'value2');
        cache.set('key3', 'value3');

        cache.clear();

        expect(cache.get('key1')).toBeUndefined();
        expect(cache.get('key2')).toBeUndefined();
        expect(cache.get('key3')).toBeUndefined();
        expect(cache.size).toBe(0);
      });
    });

    describe('prune', () => {
      it('removes expired entries', () => {
        cache.set('key1', 'value1', 1000);
        cache.set('key2', 'value2', 2000);
        cache.set('key3', 'value3', 3000);

        vi.advanceTimersByTime(1500);
        const removed = cache.prune();

        expect(removed).toBe(1);
        expect(cache.get('key1')).toBeUndefined();
        expect(cache.get('key2')).toBe('value2');
        expect(cache.get('key3')).toBe('value3');
      });

      it('returns 0 when no entries expired', () => {
        cache.set('key1', 'value1', 10000);
        const removed = cache.prune();
        expect(removed).toBe(0);
      });

      it('removes all expired entries', () => {
        cache.set('key1', 'value1', 1000);
        cache.set('key2', 'value2', 1000);
        cache.set('key3', 'value3', 1000);

        vi.advanceTimersByTime(1001);
        const removed = cache.prune();

        expect(removed).toBe(3);
        expect(cache.size).toBe(0);
      });
    });

    describe('size', () => {
      it('returns 0 for empty cache', () => {
        expect(cache.size).toBe(0);
      });

      it('returns correct count', () => {
        cache.set('key1', 'value1');
        expect(cache.size).toBe(1);

        cache.set('key2', 'value2');
        expect(cache.size).toBe(2);

        cache.set('key3', 'value3');
        expect(cache.size).toBe(3);
      });

      it('includes expired entries until accessed', () => {
        cache.set('key1', 'value1', 1000);
        vi.advanceTimersByTime(1001);
        // Size still includes expired until they are accessed or pruned
        expect(cache.size).toBe(1);
      });
    });
  });

  describe('createCacheKey', () => {
    it('creates key from method name only', () => {
      expect(createCacheKey('listFolders')).toBe('listFolders:');
    });

    it('creates key with single param', () => {
      expect(createCacheKey('listEmails', 'inbox')).toBe('listEmails:"inbox"');
    });

    it('creates key with multiple params', () => {
      expect(createCacheKey('listEmails', 'inbox', 50, 0)).toBe('listEmails:"inbox":50:0');
    });

    it('handles undefined params', () => {
      expect(createCacheKey('method', undefined, 'value')).toBe('method::"value"');
    });

    it('handles object params', () => {
      expect(createCacheKey('method', { key: 'value' })).toBe('method:{"key":"value"}');
    });

    it('handles array params', () => {
      expect(createCacheKey('method', [1, 2, 3])).toBe('method:[1,2,3]');
    });

    it('handles null params', () => {
      expect(createCacheKey('method', null)).toBe('method:null');
    });

    it('handles boolean params', () => {
      expect(createCacheKey('method', true, false)).toBe('method:true:false');
    });
  });

  describe('invalidateByPrefix', () => {
    it('clears the entire cache', () => {
      const cache = new ResponseCache();
      cache.set('prefix1:key1', 'value1');
      cache.set('prefix1:key2', 'value2');
      cache.set('prefix2:key1', 'value3');

      invalidateByPrefix(cache, 'prefix1');

      // Note: current implementation clears entire cache
      expect(cache.size).toBe(0);
    });
  });
});
