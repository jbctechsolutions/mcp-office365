/**
 * Utility functions for mappers.
 */

/**
 * Hashes a string ID to a numeric ID.
 *
 * Graph API uses string UUIDs while our row types use numeric IDs.
 * This creates a deterministic numeric ID from a string.
 *
 * Note: There's a small chance of collision, but it's acceptable
 * for our use case (display purposes only, not database operations).
 */
export function hashStringToNumber(str: string): number {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32-bit integer
  }
  // Ensure positive number
  return Math.abs(hash);
}

/**
 * Parses an ISO date string to a Unix timestamp (seconds).
 */
export function isoToTimestamp(isoDate: string | null | undefined): number | null {
  if (isoDate == null) {
    return null;
  }

  try {
    const date = new Date(isoDate);
    if (isNaN(date.getTime())) {
      return null;
    }
    // Return seconds since epoch (not milliseconds)
    return Math.floor(date.getTime() / 1000);
  } catch {
    return null;
  }
}

/**
 * Parses a Graph DateTimeTimeZone object to a Unix timestamp.
 */
export function dateTimeTimeZoneToTimestamp(
  dt: { dateTime?: string; timeZone?: string } | null | undefined
): number | null {
  if (dt?.dateTime == null) {
    return null;
  }

  try {
    // Graph API returns dates in the specified timezone
    // For simplicity, we parse as-is (usually UTC or local)
    const date = new Date(dt.dateTime);
    if (isNaN(date.getTime())) {
      return null;
    }
    return Math.floor(date.getTime() / 1000);
  } catch {
    return null;
  }
}

/**
 * Maps Graph importance to a priority number.
 */
export function importanceToPriority(importance: string | null | undefined): number {
  switch (importance?.toLowerCase()) {
    case 'high':
      return 1;
    case 'low':
      return -1;
    default:
      return 0;
  }
}

/**
 * Maps Graph flag status to a number.
 */
export function flagStatusToNumber(
  flag: { flagStatus?: string } | null | undefined
): number {
  switch (flag?.flagStatus?.toLowerCase()) {
    case 'flagged':
      return 1;
    case 'complete':
      return 2;
    default:
      return 0;
  }
}

/**
 * Extracts email address from a Graph EmailAddress object.
 */
export function extractEmailAddress(
  recipient: { emailAddress?: { address?: string; name?: string } } | null | undefined
): string | null {
  return recipient?.emailAddress?.address ?? null;
}

/**
 * Extracts display name from a Graph EmailAddress object.
 */
export function extractDisplayName(
  recipient: { emailAddress?: { address?: string; name?: string } } | null | undefined
): string | null {
  return recipient?.emailAddress?.name ?? null;
}

/**
 * Formats recipients array to a comma-separated string.
 */
export function formatRecipients(
  recipients: Array<{ emailAddress?: { address?: string; name?: string } }> | null | undefined
): string | null {
  if (recipients == null || recipients.length === 0) {
    return null;
  }

  return recipients
    .map((r) => r.emailAddress?.name ?? r.emailAddress?.address ?? '')
    .filter((s) => s.length > 0)
    .join(', ');
}

/**
 * Formats recipients array to email addresses string.
 */
export function formatRecipientAddresses(
  recipients: Array<{ emailAddress?: { address?: string; name?: string } }> | null | undefined
): string | null {
  if (recipients == null || recipients.length === 0) {
    return null;
  }

  return recipients
    .map((r) => r.emailAddress?.address ?? '')
    .filter((s) => s.length > 0)
    .join(', ');
}

/**
 * Creates a Graph content path from an entity type and ID.
 */
export function createGraphContentPath(type: string, id: string): string {
  return `graph-${type}:${id}`;
}
