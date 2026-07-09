/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Durable-ID tokens (U5 / D1). Two shapes:
 *
 * - **Self-encoding** (`em_`, `ev_`, `ct_`, `fd_`, `dr_`, `nb_`, `ns_`,
 *   `np_`): the token carries the immutable Graph ID directly —
 *   `<prefix>_<base64url(graphId)>`.
 *   Resolution is a decode with zero storage, so it survives a cold/lost
 *   `state.db` and resolves on any machine (the core cold-state fix).
 *
 * - **Alias-backed** (`pl_`, `pt_`, `ch_`, `tm_`, `at_`, `td_`, `tl_`,
 *   composite tuples): the token is a short deterministic digest of the
 *   entity's canonical key — `<prefix>_<base32(sha256(canonicalKey))[0..13]>`
 *   (70 bits) — backed by the alias table (D3). These are machine-scoped: a
 *   cold store yields `ID_UNKNOWN`.
 *
 * Determinism is a hard requirement: the same Graph ID / canonical key always
 * mints a byte-identical token.
 */

import { createHash } from 'node:crypto';

/** Entity kinds addressable by a durable-ID token. */
export type EntityType =
  | 'message'
  | 'event'
  | 'contact'
  | 'folder'
  | 'driveItem'
  | 'task'
  | 'taskList'
  | 'plan'
  | 'plannerTask'
  | 'chat'
  | 'team'
  | 'attachment'
  | 'channel'
  | 'chatMessage'
  | 'channelMessage'
  | 'checklistItem'
  | 'noteNotebook'
  | 'noteSection'
  | 'notePage';

/** How a token encodes its target. */
export type TokenKind = 'self' | 'alias';

/**
 * Prefix → entity maps are NULL-PROTOTYPE so a caller-supplied prefix like
 * `constructor` or `toString` (from an arbitrary ID string) can never match an
 * inherited `Object.prototype` member and get misclassified as a known token —
 * `parseToken` must return null for those and let the resolver pass them through
 * as opaque Graph IDs.
 */

/** Self-encoding prefixes → entity type (the token carries the Graph ID). */
export const SELF_ENCODING_PREFIXES: Readonly<Record<string, EntityType>> = Object.assign(
  Object.create(null) as Record<string, EntityType>,
  {
    em: 'message',
    ev: 'event',
    ct: 'contact',
    fd: 'folder',
    dr: 'driveItem',
    nb: 'noteNotebook',
    ns: 'noteSection',
    np: 'notePage',
  } satisfies Record<string, EntityType>,
);

/** Alias-backed prefixes → entity type (the token is a digest; needs the store). */
export const ALIAS_PREFIXES: Readonly<Record<string, EntityType>> = Object.assign(
  Object.create(null) as Record<string, EntityType>,
  {
    pl: 'plan',
    pt: 'plannerTask',
    ch: 'chat',
    tm: 'team',
    at: 'attachment',
    cn: 'channel',
    cm: 'chatMessage',
    xm: 'channelMessage',
    ci: 'checklistItem',
    td: 'task',
    tl: 'taskList',
  } satisfies Record<string, EntityType>,
);

const ENTITY_TO_PREFIX: Readonly<Record<EntityType, string>> = buildReverse();

function buildReverse(): Record<EntityType, string> {
  const out = Object.create(null) as Record<EntityType, string>;
  for (const [prefix, entity] of Object.entries(SELF_ENCODING_PREFIXES)) {
    out[entity] = prefix;
  }
  for (const [prefix, entity] of Object.entries(ALIAS_PREFIXES)) {
    out[entity] = prefix;
  }
  return out;
}

/** Crockford-free RFC 4648 base32 alphabet (lowercase for compact tokens). */
const BASE32_ALPHABET = 'abcdefghijklmnopqrstuvwxyz234567';
/** 14 base32 chars × 5 bits = 70 bits of the sha256 digest (D1). */
const COMPOSITE_DIGEST_CHARS = 14;

/** A parsed durable-ID token. */
export interface ParsedToken {
  prefix: string;
  kind: TokenKind;
  entityType: EntityType;
  /** For self-encoding tokens: the decoded Graph ID. Absent for alias tokens. */
  graphId?: string;
}

/** True when a prefix is a known self-encoding or alias prefix. */
export function isKnownPrefix(prefix: string): boolean {
  return prefix in SELF_ENCODING_PREFIXES || prefix in ALIAS_PREFIXES;
}

/**
 * Mints a self-encoding token carrying the immutable Graph ID (D1). Determin­istic.
 */
export function mintSelfEncoded(entityType: EntityType, graphId: string): string {
  const prefix = ENTITY_TO_PREFIX[entityType];
  if (prefix == null || !(prefix in SELF_ENCODING_PREFIXES)) {
    throw new Error(`Entity type "${entityType}" is not self-encoding.`);
  }
  if (graphId.length === 0) {
    // An empty ID would encode to `<prefix>_` (empty payload), which parseToken
    // rejects — the token would then be mis-handled as a raw Graph ID. Fail fast.
    throw new Error('Cannot mint a self-encoding token for an empty Graph ID.');
  }
  return `${prefix}_${Buffer.from(graphId, 'utf8').toString('base64url')}`;
}

/**
 * Mints a short, deterministic alias-backed token from an entity's canonical
 * key (D1). The caller stores the token → key mapping in the alias table.
 */
export function mintComposite(entityType: EntityType, canonicalKey: string): string {
  const prefix = ENTITY_TO_PREFIX[entityType];
  if (prefix == null || !(prefix in ALIAS_PREFIXES)) {
    throw new Error(`Entity type "${entityType}" is not alias-backed.`);
  }
  return `${prefix}_${base32Digest(canonicalKey)}`;
}

/**
 * Builds a canonical key for a composite entity from its identifying tuple.
 * Keys are sorted so field order never changes the digest. Keys and values are
 * percent-encoded so a value containing the `&`/`=` delimiters (Graph IDs are
 * base64-ish and can) cannot forge a boundary and make two distinct tuples
 * produce the same key — which would collide their tokens without a hash
 * collision.
 */
export function canonicalKey(entityType: EntityType, parts: Readonly<Record<string, string>>): string {
  const sorted = Object.keys(parts)
    .sort()
    .map((k) => `${encodeURIComponent(k)}=${encodeURIComponent(parts[k] ?? '')}`)
    .join('&');
  return `${entityType}:${sorted}`;
}

/**
 * Parses a token into its prefix/kind/entity (and decoded Graph ID for
 * self-encoding tokens). Returns null when the value is not a well-formed known
 * token — the caller decides whether to treat that as a raw Graph ID or reject.
 */
export function parseToken(token: string): ParsedToken | null {
  const underscore = token.indexOf('_');
  if (underscore <= 0) {
    return null;
  }
  const prefix = token.slice(0, underscore);
  const payload = token.slice(underscore + 1);
  if (payload.length === 0) {
    return null;
  }

  const selfEntity = SELF_ENCODING_PREFIXES[prefix];
  if (selfEntity !== undefined) {
    const graphId = decodeBase64Url(payload);
    if (graphId == null) {
      return null;
    }
    return { prefix, kind: 'self', entityType: selfEntity, graphId };
  }

  const aliasEntity = ALIAS_PREFIXES[prefix];
  if (aliasEntity !== undefined) {
    return { prefix, kind: 'alias', entityType: aliasEntity };
  }

  return null;
}

/** True when a string looks like a durable-ID token (known prefix + payload). */
export function isToken(value: string): boolean {
  return parseToken(value) !== null;
}

function base32Digest(input: string): string {
  const hash = createHash('sha256').update(input, 'utf8').digest();
  let bits = 0;
  let value = 0;
  let output = '';
  for (const byte of hash) {
    value = (value << 8) | byte;
    bits += 8;
    while (bits >= 5) {
      output += BASE32_ALPHABET[(value >>> (bits - 5)) & 0x1f];
      bits -= 5;
      if (output.length === COMPOSITE_DIGEST_CHARS) {
        return output;
      }
    }
  }
  return output;
}

function decodeBase64Url(payload: string): string | null {
  // base64url is [A-Za-z0-9_-]; reject anything else so a stray alias-shaped
  // string isn't silently decoded into garbage.
  if (!/^[A-Za-z0-9_-]+$/.test(payload)) {
    return null;
  }
  let decoded: string;
  try {
    decoded = Buffer.from(payload, 'base64url').toString('utf8');
  } catch {
    return null;
  }
  if (decoded.length === 0) {
    return null;
  }
  // Reject NON-CANONICAL encodings: Buffer's base64url decode is lenient (em_QQ,
  // em_QR, em_QS all decode to "A"), which would give one Graph ID many valid
  // token strings. Requiring the payload to be the exact canonical encoding of
  // the decoded bytes means one Graph ID ↔ exactly one token string, so a later
  // unit keying a control on the token string can't be evaded by a variant.
  if (Buffer.from(decoded, 'utf8').toString('base64url') !== payload) {
    return null;
  }
  return decoded;
}
