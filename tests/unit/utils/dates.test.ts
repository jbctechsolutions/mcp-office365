/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect } from 'vitest';
import {
  APPLE_EPOCH_OFFSET,
  appleTimestampToIso,
  appleTimestampToDate,
  isoToAppleTimestamp,
  dateToAppleTimestamp,
} from '../../../src/utils/dates.js';

describe('dates', () => {
  describe('APPLE_EPOCH_OFFSET', () => {
    it('equals the correct number of seconds between Unix and Apple epochs', () => {
      // 2001-01-01 00:00:00 UTC in Unix time
      const expected = Date.UTC(2001, 0, 1) / 1000;
      expect(APPLE_EPOCH_OFFSET).toBe(expected);
      expect(APPLE_EPOCH_OFFSET).toBe(978307200);
    });
  });

  describe('appleTimestampToIso', () => {
    it('converts Apple epoch 0 to 2001-01-01T00:00:00.000Z', () => {
      const result = appleTimestampToIso(0);
      expect(result).toBe('2001-01-01T00:00:00.000Z');
    });

    it('converts a positive timestamp correctly', () => {
      // 1 day = 86400 seconds
      const result = appleTimestampToIso(86400);
      expect(result).toBe('2001-01-02T00:00:00.000Z');
    });

    it('converts a negative timestamp correctly', () => {
      // -1 day should be 2000-12-31
      const result = appleTimestampToIso(-86400);
      expect(result).toBe('2000-12-31T00:00:00.000Z');
    });

    it('returns null for null input', () => {
      expect(appleTimestampToIso(null)).toBeNull();
    });

    it('returns null for undefined input', () => {
      expect(appleTimestampToIso(undefined)).toBeNull();
    });

    it('handles fractional seconds', () => {
      const result = appleTimestampToIso(0.5);
      expect(result).toBe('2001-01-01T00:00:00.500Z');
    });
  });

  describe('appleTimestampToDate', () => {
    it('converts Apple epoch 0 to correct Date', () => {
      const result = appleTimestampToDate(0);
      expect(result).toBeInstanceOf(Date);
      expect(result?.toISOString()).toBe('2001-01-01T00:00:00.000Z');
    });

    it('returns null for null input', () => {
      expect(appleTimestampToDate(null)).toBeNull();
    });

    it('returns null for undefined input', () => {
      expect(appleTimestampToDate(undefined)).toBeNull();
    });
  });

  describe('isoToAppleTimestamp', () => {
    it('converts 2001-01-01T00:00:00.000Z to Apple epoch 0', () => {
      const result = isoToAppleTimestamp('2001-01-01T00:00:00.000Z');
      expect(result).toBe(0);
    });

    it('converts a later date correctly', () => {
      const result = isoToAppleTimestamp('2001-01-02T00:00:00.000Z');
      expect(result).toBe(86400);
    });

    it('converts an earlier date correctly', () => {
      const result = isoToAppleTimestamp('2000-12-31T00:00:00.000Z');
      expect(result).toBe(-86400);
    });

    it('returns null for null input', () => {
      expect(isoToAppleTimestamp(null)).toBeNull();
    });

    it('returns null for undefined input', () => {
      expect(isoToAppleTimestamp(undefined)).toBeNull();
    });

    it('returns null for invalid date string', () => {
      expect(isoToAppleTimestamp('not-a-date')).toBeNull();
    });
  });

  describe('dateToAppleTimestamp', () => {
    it('converts Date to Apple timestamp', () => {
      const date = new Date('2001-01-01T00:00:00.000Z');
      const result = dateToAppleTimestamp(date);
      expect(result).toBe(0);
    });

    it('returns null for null input', () => {
      expect(dateToAppleTimestamp(null)).toBeNull();
    });

    it('returns null for undefined input', () => {
      expect(dateToAppleTimestamp(undefined)).toBeNull();
    });
  });

  describe('round-trip conversions', () => {
    it('converts back and forth correctly for ISO strings', () => {
      const original = '2024-06-15T12:30:45.000Z';
      const appleTs = isoToAppleTimestamp(original);
      const backToIso = appleTimestampToIso(appleTs);
      expect(backToIso).toBe(original);
    });

    it('converts back and forth correctly for Apple timestamps', () => {
      const originalTs = 739584645; // Some arbitrary timestamp
      const iso = appleTimestampToIso(originalTs);
      const backToTs = isoToAppleTimestamp(iso);
      expect(backToTs).toBe(originalTs);
    });
  });
});
