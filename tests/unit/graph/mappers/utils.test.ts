/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Graph mapper utility functions.
 */

import { describe, it, expect, beforeAll } from 'vitest';
import {
  hashStringToNumber,
  isoToTimestamp,
  dateTimeTimeZoneToTimestamp,
  unixTimestampToIso,
  importanceToPriority,
  flagStatusToNumber,
  extractEmailAddress,
  extractDisplayName,
  formatRecipients,
  formatRecipientAddresses,
  createGraphContentPath,
} from '../../../../src/graph/mappers/utils.js';

describe('graph/mappers/utils', () => {
  describe('hashStringToNumber', () => {
    it('converts string to positive number', () => {
      const result = hashStringToNumber('test-uuid-123');
      expect(result).toBeTypeOf('number');
      expect(result).toBeGreaterThanOrEqual(0);
    });

    it('produces consistent results for same input', () => {
      const input = 'abc123def456';
      const result1 = hashStringToNumber(input);
      const result2 = hashStringToNumber(input);
      expect(result1).toBe(result2);
    });

    it('produces different results for different inputs', () => {
      const result1 = hashStringToNumber('input-a');
      const result2 = hashStringToNumber('input-b');
      expect(result1).not.toBe(result2);
    });

    it('handles empty string', () => {
      const result = hashStringToNumber('');
      expect(result).toBe(0);
    });

    it('handles special characters', () => {
      const result = hashStringToNumber('test@email.com!#$%');
      expect(result).toBeTypeOf('number');
      expect(result).toBeGreaterThanOrEqual(0);
    });
  });

  describe('isoToTimestamp', () => {
    it('converts ISO date string to Unix timestamp', () => {
      const result = isoToTimestamp('2024-01-15T10:30:00Z');
      expect(result).toBe(1705314600);
    });

    it('returns null for null input', () => {
      expect(isoToTimestamp(null)).toBeNull();
    });

    it('returns null for undefined input', () => {
      expect(isoToTimestamp(undefined)).toBeNull();
    });

    it('returns null for invalid date string', () => {
      expect(isoToTimestamp('not-a-date')).toBeNull();
    });

    it('handles date with timezone offset', () => {
      const result = isoToTimestamp('2024-01-15T10:30:00+05:00');
      expect(result).toBeTypeOf('number');
    });
  });

  describe('dateTimeTimeZoneToTimestamp', () => {
    it('converts DateTimeTimeZone object to timestamp', () => {
      const dt = { dateTime: '2024-01-15T10:30:00', timeZone: 'UTC' };
      const result = dateTimeTimeZoneToTimestamp(dt);
      expect(result).toBeTypeOf('number');
    });

    it('returns null for null input', () => {
      expect(dateTimeTimeZoneToTimestamp(null)).toBeNull();
    });

    it('returns null for undefined input', () => {
      expect(dateTimeTimeZoneToTimestamp(undefined)).toBeNull();
    });

    it('returns null when dateTime is undefined', () => {
      const dt = { timeZone: 'UTC' };
      expect(dateTimeTimeZoneToTimestamp(dt)).toBeNull();
    });

    it('returns null when dateTime is null', () => {
      const dt = { dateTime: undefined as unknown as string, timeZone: 'UTC' };
      expect(dateTimeTimeZoneToTimestamp(dt)).toBeNull();
    });

    it('treats dateTime as UTC when timeZone is UTC', () => {
      // Graph API returns { dateTime: "2026-02-23T16:00:00.0000000", timeZone: "UTC" }
      // 2026-02-23T16:00:00Z = 1771869600 Unix seconds
      const dt = { dateTime: '2026-02-23T16:00:00.0000000', timeZone: 'UTC' };
      const result = dateTimeTimeZoneToTimestamp(dt);
      expect(result).toBe(1771862400);
    });

    it('treats dateTime as UTC when timeZone is Etc/GMT', () => {
      const dt = { dateTime: '2026-02-23T16:00:00.0000000', timeZone: 'Etc/GMT' };
      const result = dateTimeTimeZoneToTimestamp(dt);
      expect(result).toBe(1771862400);
    });

    it('handles object without timeZone', () => {
      const dt = { dateTime: '2024-01-15T10:30:00' };
      const result = dateTimeTimeZoneToTimestamp(dt);
      expect(result).toBeTypeOf('number');
    });
  });

  describe('importanceToPriority', () => {
    it('returns 1 for high importance', () => {
      expect(importanceToPriority('high')).toBe(1);
      expect(importanceToPriority('HIGH')).toBe(1);
      expect(importanceToPriority('High')).toBe(1);
    });

    it('returns -1 for low importance', () => {
      expect(importanceToPriority('low')).toBe(-1);
      expect(importanceToPriority('LOW')).toBe(-1);
      expect(importanceToPriority('Low')).toBe(-1);
    });

    it('returns 0 for normal importance', () => {
      expect(importanceToPriority('normal')).toBe(0);
    });

    it('returns 0 for null', () => {
      expect(importanceToPriority(null)).toBe(0);
    });

    it('returns 0 for undefined', () => {
      expect(importanceToPriority(undefined)).toBe(0);
    });

    it('returns 0 for unknown importance', () => {
      expect(importanceToPriority('critical')).toBe(0);
    });
  });

  describe('flagStatusToNumber', () => {
    it('returns 1 for flagged', () => {
      expect(flagStatusToNumber({ flagStatus: 'flagged' })).toBe(1);
      expect(flagStatusToNumber({ flagStatus: 'Flagged' })).toBe(1);
      expect(flagStatusToNumber({ flagStatus: 'FLAGGED' })).toBe(1);
    });

    it('returns 2 for complete', () => {
      expect(flagStatusToNumber({ flagStatus: 'complete' })).toBe(2);
      expect(flagStatusToNumber({ flagStatus: 'Complete' })).toBe(2);
      expect(flagStatusToNumber({ flagStatus: 'COMPLETE' })).toBe(2);
    });

    it('returns 0 for notFlagged', () => {
      expect(flagStatusToNumber({ flagStatus: 'notFlagged' })).toBe(0);
    });

    it('returns 0 for null', () => {
      expect(flagStatusToNumber(null)).toBe(0);
    });

    it('returns 0 for undefined', () => {
      expect(flagStatusToNumber(undefined)).toBe(0);
    });

    it('returns 0 when flagStatus is undefined', () => {
      expect(flagStatusToNumber({})).toBe(0);
    });
  });

  describe('extractEmailAddress', () => {
    it('extracts email address from recipient', () => {
      const recipient = {
        emailAddress: { address: 'test@example.com', name: 'Test User' },
      };
      expect(extractEmailAddress(recipient)).toBe('test@example.com');
    });

    it('returns null for null recipient', () => {
      expect(extractEmailAddress(null)).toBeNull();
    });

    it('returns null for undefined recipient', () => {
      expect(extractEmailAddress(undefined)).toBeNull();
    });

    it('returns null when emailAddress is missing', () => {
      expect(extractEmailAddress({})).toBeNull();
    });

    it('returns null when address is missing', () => {
      const recipient = { emailAddress: { name: 'Test User' } };
      expect(extractEmailAddress(recipient)).toBeNull();
    });
  });

  describe('extractDisplayName', () => {
    it('extracts display name from recipient', () => {
      const recipient = {
        emailAddress: { address: 'test@example.com', name: 'Test User' },
      };
      expect(extractDisplayName(recipient)).toBe('Test User');
    });

    it('returns null for null recipient', () => {
      expect(extractDisplayName(null)).toBeNull();
    });

    it('returns null for undefined recipient', () => {
      expect(extractDisplayName(undefined)).toBeNull();
    });

    it('returns null when emailAddress is missing', () => {
      expect(extractDisplayName({})).toBeNull();
    });

    it('returns null when name is missing', () => {
      const recipient = { emailAddress: { address: 'test@example.com' } };
      expect(extractDisplayName(recipient)).toBeNull();
    });
  });

  describe('formatRecipients', () => {
    it('formats single recipient with name', () => {
      const recipients = [
        { emailAddress: { address: 'test@example.com', name: 'Test User' } },
      ];
      expect(formatRecipients(recipients)).toBe('Test User');
    });

    it('formats multiple recipients', () => {
      const recipients = [
        { emailAddress: { address: 'a@example.com', name: 'User A' } },
        { emailAddress: { address: 'b@example.com', name: 'User B' } },
      ];
      expect(formatRecipients(recipients)).toBe('User A, User B');
    });

    it('uses email address when name is missing', () => {
      const recipients = [
        { emailAddress: { address: 'test@example.com' } },
      ];
      expect(formatRecipients(recipients)).toBe('test@example.com');
    });

    it('returns null for null recipients', () => {
      expect(formatRecipients(null)).toBeNull();
    });

    it('returns null for undefined recipients', () => {
      expect(formatRecipients(undefined)).toBeNull();
    });

    it('returns null for empty array', () => {
      expect(formatRecipients([])).toBeNull();
    });

    it('filters out empty entries', () => {
      const recipients = [
        { emailAddress: { address: 'a@example.com', name: 'User A' } },
        { emailAddress: {} },
        { emailAddress: { address: 'b@example.com', name: 'User B' } },
      ];
      expect(formatRecipients(recipients)).toBe('User A, User B');
    });
  });

  describe('formatRecipientAddresses', () => {
    it('formats single recipient address', () => {
      const recipients = [
        { emailAddress: { address: 'test@example.com', name: 'Test User' } },
      ];
      expect(formatRecipientAddresses(recipients)).toBe('test@example.com');
    });

    it('formats multiple addresses', () => {
      const recipients = [
        { emailAddress: { address: 'a@example.com', name: 'User A' } },
        { emailAddress: { address: 'b@example.com', name: 'User B' } },
      ];
      expect(formatRecipientAddresses(recipients)).toBe('a@example.com, b@example.com');
    });

    it('returns null for null recipients', () => {
      expect(formatRecipientAddresses(null)).toBeNull();
    });

    it('returns null for undefined recipients', () => {
      expect(formatRecipientAddresses(undefined)).toBeNull();
    });

    it('returns null for empty array', () => {
      expect(formatRecipientAddresses([])).toBeNull();
    });

    it('filters out missing addresses', () => {
      const recipients = [
        { emailAddress: { address: 'a@example.com' } },
        { emailAddress: { name: 'No Email' } },
        { emailAddress: { address: 'b@example.com' } },
      ];
      expect(formatRecipientAddresses(recipients)).toBe('a@example.com, b@example.com');
    });
  });

  describe('unixTimestampToIso', () => {
    it('converts Unix timestamp to ISO string', () => {
      // 2026-02-23T15:00:00.000Z = 1771858800 seconds
      const result = unixTimestampToIso(1771858800);
      expect(result).toBe('2026-02-23T15:00:00.000Z');
    });

    it('returns null for null input', () => {
      expect(unixTimestampToIso(null)).toBeNull();
    });

    it('returns null for undefined input', () => {
      expect(unixTimestampToIso(undefined)).toBeNull();
    });

    it('does not add Apple epoch offset', () => {
      // Unix timestamp 0 = 1970-01-01T00:00:00Z (not 2001-01-01)
      const result = unixTimestampToIso(0);
      expect(result).toBe('1970-01-01T00:00:00.000Z');
    });
  });

  describe('unixTimestampToLocalIso', () => {
    // Function imported dynamically to allow RED phase (function doesn't exist yet)
    type LocalIsoFn = (ts: number | null | undefined) => string | null;
    let fn: LocalIsoFn;

    beforeAll(async () => {
      const mod = await import('../../../../src/graph/mappers/utils.js') as Record<string, unknown>;
      if (typeof mod.unixTimestampToLocalIso !== 'function') {
        throw new Error('unixTimestampToLocalIso is not exported from utils.ts');
      }
      fn = mod.unixTimestampToLocalIso as LocalIsoFn;
    });

    it('returns a string with timezone offset (not Z)', () => {
      const result = fn(1771858800);
      expect(result).toBeTypeOf('string');
      // Should NOT end with Z — should have offset like +00:00 or -05:00
      expect(result).not.toMatch(/Z$/);
      expect(result).toMatch(/[+-]\d{2}:\d{2}$/);
    });

    it('produces the correct UTC point in time', () => {
      // 1771858800 = 2026-02-23T15:00:00Z
      const result = fn(1771858800)!;
      // Parsing the result back should give the same instant
      const parsed = new Date(result);
      expect(parsed.getTime()).toBe(1771858800 * 1000);
    });

    it('returns null for null input', () => {
      expect(fn(null)).toBeNull();
    });

    it('returns null for undefined input', () => {
      expect(fn(undefined)).toBeNull();
    });

    it('includes date and time components', () => {
      const result = fn(1771858800)!;
      // Should match ISO-like format: YYYY-MM-DDTHH:MM:SS.sss±HH:MM
      expect(result).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}[+-]\d{2}:\d{2}$/);
    });

    it('does not add Apple epoch offset', () => {
      // Unix timestamp 0 = 1970-01-01T00:00:00Z — in local time this is
      // Dec 31, 1969 for negative UTC offsets or Jan 1, 1970 for non-negative
      const result = fn(0)!;
      const parsed = new Date(result);
      expect(parsed.getTime()).toBe(0);
      // The year should be 1969 or 1970, NOT 2001 (Apple epoch)
      expect(result).toMatch(/^19(69|70)-/);
    });
  });

  describe('createGraphContentPath', () => {
    it('creates path for email', () => {
      expect(createGraphContentPath('email', 'msg-123')).toBe('graph-email:msg-123');
    });

    it('creates path for event', () => {
      expect(createGraphContentPath('event', 'evt-456')).toBe('graph-event:evt-456');
    });

    it('creates path for contact', () => {
      expect(createGraphContentPath('contact', 'contact-789')).toBe('graph-contact:contact-789');
    });

    it('creates path for task', () => {
      expect(createGraphContentPath('task', 'list-1:task-1')).toBe('graph-task:list-1:task-1');
    });
  });
});
