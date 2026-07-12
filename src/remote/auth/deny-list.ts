/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Revocation deny-list (U4 reads it; U7 fills it). Purging a user's tokens alone
 * does not stop access — claude.ai may still hold a valid Entra token and an OBO
 * would silently re-mint. The middleware rejects deny-listed identities before
 * any OBO, per request from the store (never process-cached, so a later
 * perf-motivated cache can't reintroduce a staleness window).
 */

/** Checks whether an Entra object id is revoked. */
export interface DenyList {
  isDenied(oid: string): boolean;
}

/**
 * In-memory stub used until U7 wires a store-backed deny-list. Always empty —
 * no user is revoked yet — but present so the middleware's rejection path exists
 * and is testable from day one.
 */
export function createStubDenyList(): DenyList {
  const denied = new Set<string>();
  return { isDenied: (oid: string): boolean => denied.has(oid) };
}
