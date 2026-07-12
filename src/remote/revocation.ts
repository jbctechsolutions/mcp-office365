/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Offboarding / revocation (U7). Purging a user's tokens alone does not stop
 * access — claude.ai may still hold a valid Entra token and an OBO would
 * silently re-mint. So `revoke` does both, deny-list FIRST:
 *
 *   1. Add the oid to the deny-list (the auth middleware rejects it before OBO).
 *   2. Purge their durable state (approval tokens, aliases, delta cursors).
 *
 * Ordering matters: inserting the deny-list row before the purge closes the
 * window where a request already past middleware could write fresh state after
 * the purge — it would be rejected on its next call, and nothing new persists.
 * Entra account disablement is the independent backstop (OBO then fails).
 */

import type { StateStore } from '../state/store.js';

/** Result of a revoke. */
export interface RevokeResult {
  readonly oid: string;
  readonly homeAccountId: string;
  readonly purgedRows: number;
}

/**
 * Revokes a user: deny-list first, then purge their per-account durable state.
 * `homeAccountId` is `<oid>.<tenantId>` (the store's account key).
 */
export function revokeUser(
  store: StateStore,
  oid: string,
  homeAccountId: string,
  now: number,
  reason?: string,
): RevokeResult {
  store.denyUser(oid, reason ?? null, now); // deny-list BEFORE purge (ordering)
  const purgedRows = store.purgeAccount(homeAccountId);
  return { oid, homeAccountId, purgedRows };
}

/** Re-admits a previously revoked user (removes the deny-list entry). */
export function readmitUser(store: StateStore, oid: string): boolean {
  return store.readmitUser(oid);
}

/** Lists revoked identities. */
export function listRevoked(
  store: StateStore,
): Array<{ oid: string; reason: string | null; deniedAt: number }> {
  return store.listDenied();
}
