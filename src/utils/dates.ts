/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Date conversion utilities for Outlook's Apple epoch timestamps.
 *
 * Outlook for Mac stores timestamps as seconds since the Apple epoch
 * (January 1, 2001, 00:00:00 UTC), while JavaScript uses milliseconds
 * since the Unix epoch (January 1, 1970, 00:00:00 UTC).
 */

/**
 * Seconds between Unix epoch (1970-01-01) and Apple epoch (2001-01-01).
 * Calculated as: Date.UTC(2001, 0, 1) / 1000 = 978307200
 */
export const APPLE_EPOCH_OFFSET = 978307200;

/**
 * Converts an Apple epoch timestamp to an ISO 8601 string.
 *
 * @param timestamp - Seconds since Apple epoch (2001-01-01), or null/undefined
 * @returns ISO 8601 formatted date string, or null if input is null/undefined
 *
 * @example
 * ```ts
 * appleTimestampToIso(0);
 * // Returns: '2001-01-01T00:00:00.000Z'
 *
 * appleTimestampToIso(null);
 * // Returns: null
 * ```
 */
export function appleTimestampToIso(
  timestamp: number | null | undefined
): string | null {
  if (timestamp === null || timestamp === undefined) {
    return null;
  }

  const unixTimestampMs = (timestamp + APPLE_EPOCH_OFFSET) * 1000;
  return new Date(unixTimestampMs).toISOString();
}

/**
 * Converts an Apple epoch timestamp to a JavaScript Date object.
 *
 * @param timestamp - Seconds since Apple epoch (2001-01-01), or null/undefined
 * @returns Date object, or null if input is null/undefined
 *
 * @example
 * ```ts
 * appleTimestampToDate(0);
 * // Returns: Date representing 2001-01-01T00:00:00.000Z
 * ```
 */
export function appleTimestampToDate(
  timestamp: number | null | undefined
): Date | null {
  if (timestamp === null || timestamp === undefined) {
    return null;
  }

  const unixTimestampMs = (timestamp + APPLE_EPOCH_OFFSET) * 1000;
  return new Date(unixTimestampMs);
}

/**
 * Converts an ISO 8601 string to an Apple epoch timestamp.
 *
 * @param isoString - ISO 8601 formatted date string, or null/undefined
 * @returns Seconds since Apple epoch (2001-01-01), or null if input is null/undefined
 *
 * @example
 * ```ts
 * isoToAppleTimestamp('2001-01-01T00:00:00.000Z');
 * // Returns: 0
 * ```
 */
export function isoToAppleTimestamp(
  isoString: string | null | undefined
): number | null {
  if (isoString === null || isoString === undefined) {
    return null;
  }

  const date = new Date(isoString);
  if (isNaN(date.getTime())) {
    return null;
  }

  const unixTimestampSec = Math.floor(date.getTime() / 1000);
  return unixTimestampSec - APPLE_EPOCH_OFFSET;
}

/**
 * Converts a JavaScript Date to an Apple epoch timestamp.
 *
 * @param date - Date object, or null/undefined
 * @returns Seconds since Apple epoch (2001-01-01), or null if input is null/undefined
 */
export function dateToAppleTimestamp(date: Date | null | undefined): number | null {
  if (date === null || date === undefined) {
    return null;
  }

  const unixTimestampSec = Math.floor(date.getTime() / 1000);
  return unixTimestampSec - APPLE_EPOCH_OFFSET;
}
