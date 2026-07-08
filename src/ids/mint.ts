/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Store-backed minting for durable-ID tokens (U5 / D1). Self-encoding tokens
 * need no storage — {@link mintSelfEncoded} alone suffices. Composite/mutable
 * entities mint a short digest token AND persist the token → Graph-ID mapping in
 * the alias table so the resolver can look it up later.
 */

import type { StateStore } from '../state/store.js';
import { IdCollisionError } from '../utils/errors.js';
import { canonicalKey, mintComposite, type EntityType } from './token.js';

export interface RegisterCompositeArgs {
  entityType: EntityType;
  /** Identifying tuple (e.g. `{ messageId, attachmentId }`). */
  parts: Readonly<Record<string, string>>;
  /** The live Graph ID this token resolves to. */
  graphId: string;
  accountId: string;
  /** True for `$search`-minted rows whose Graph ID may drift (D2). */
  mutable?: boolean;
}

/**
 * Mints an alias-backed token for a composite entity and records it. Idempotent
 * for the same entity (deterministic token + same Graph ID). If a *different*
 * entity's Graph ID already occupies the token, throws {@link IdCollisionError}
 * (D1a) rather than silently overwriting — the fixed-length digest never
 * mis-resolves.
 */
export function registerComposite(store: StateStore, args: RegisterCompositeArgs): string {
  // Fold the account into the key so a tenant-global resource (chat/team/channel)
  // mints a DISTINCT token per account. Otherwise two accounts listing the same
  // shared object would mint one token and putAlias would flip its account_id,
  // locking the first account out of its own token with ID_FOREIGN_ACCOUNT (D7).
  const key = canonicalKey(args.entityType, { ...args.parts, '@account': args.accountId });
  const token = mintComposite(args.entityType, key);

  // The collision check + write happen atomically inside the store (D1a): a
  // differing Graph ID for an immutable entity is a genuine digest collision and
  // is refused; a mutable ($search-minted) entity's drifted ID (D2) updates.
  const result = store.registerAlias({
    token,
    graphId: args.graphId,
    entityType: args.entityType,
    accountId: args.accountId,
    mutable: args.mutable === true,
  });
  if (result === 'collision') {
    throw new IdCollisionError(token);
  }
  return token;
}
