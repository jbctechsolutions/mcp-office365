/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Universal durable-ID resolver (U5 / D1, D2, D4, D7). Turns any ID a tool
 * receives — a self-encoding token, an alias-backed token, a raw Graph ID, or a
 * legacy numeric ID — into a live Graph ID, or a typed error.
 *
 * - self-encoding token → decode, zero storage (cold-state durable).
 * - alias-backed token   → alias table lookup (account-scoped). A cold/foreign
 *   miss is distinguished: ID_FOREIGN_ACCOUNT vs ID_UNKNOWN.
 * - numeric (v2 hash)    → NUMERIC_ID_UNSUPPORTED on Graph (lossy, D4).
 * - raw non-token string → passed through as an opaque Graph ID (a user may
 *   paste a real Graph ID; Graph IDs are opaque, so rejecting them would be
 *   hostile). Mutation of state lives in the alias table, not here.
 */

import type { StateStore } from '../state/store.js';
import {
  IdForeignAccountError,
  IdUnknownError,
  NumericIdUnsupportedError,
} from '../utils/errors.js';
import { parseToken } from './token.js';

/** A resolved Graph ID plus whether it came from a mutable (drift-prone) alias. */
export interface ResolvedId {
  graphId: string;
  /** True for `$search`-minted mutable alias rows that may need re-resolution. */
  mutable: boolean;
}

/**
 * Resolves an ID parameter to a Graph ID for the given signed-in account.
 * Throws a typed {@link import('../utils/errors.js').OutlookMcpError} on failure.
 */
export function resolveId(
  id: string | number,
  accountId: string,
  store: StateStore,
): ResolvedId {
  if (typeof id === 'number') {
    throw new NumericIdUnsupportedError(id);
  }

  const parsed = parseToken(id);
  if (parsed === null) {
    // Not a known token — treat as an opaque Graph ID the caller supplied.
    return { graphId: id, mutable: false };
  }

  if (parsed.kind === 'self') {
    // graphId is always present for a self-encoding parse.
    return { graphId: parsed.graphId as string, mutable: false };
  }

  // Alias-backed: resolve against the store, scoped to the account.
  const row = store.getAlias(id, accountId);
  if (row !== null) {
    return { graphId: row.graphId, mutable: row.mutable };
  }

  // Miss — distinguish a foreign-account token from a genuinely unknown one so
  // the agent gets an actionable message (D7).
  const owner = store.getAliasAccount(id);
  if (owner !== null && owner !== accountId) {
    throw new IdForeignAccountError(id);
  }
  throw new IdUnknownError(id);
}
