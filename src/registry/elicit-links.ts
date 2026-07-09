/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Factory helpers for the declarative `onElicit` links (U11).
 *
 * A `prepare_*` tool opts into inline elicitation by pointing at its `confirm_*`
 * counterpart and describing how to build that confirm's params from the prepare
 * params + the minted token(s). The three confirm shapes in the codebase each
 * get a one-line factory so the ~35 tool defs stay declarative:
 *
 * - {@link approvalTokenLink} — confirm takes `{ approval_token }`.
 * - {@link tokenIdLink}       — confirm takes `{ token_id, ...copied ids }`.
 * - {@link batchLink}         — confirm takes `{ tokens: [{token_id, email_id}] }`.
 */

import type { ElicitLink, ToolResult } from './types.js';

/** Parses a prepare tool's single-text-block JSON result, or {} on any failure. */
function parseResult(result: ToolResult): Record<string, unknown> {
  const text = result.content[0]?.text;
  if (text == null) return {};
  try {
    return JSON.parse(text) as Record<string, unknown>;
  } catch {
    return {};
  }
}

/** The lone token a single-target prepare minted (`approval_token`/`token_id`). */
function singleToken(result: ToolResult): string | null {
  const parsed = parseResult(result);
  const token = parsed['approval_token'] ?? parsed['token_id'];
  return typeof token === 'string' && token.length > 0 ? token : null;
}

function singleTokenIds(result: ToolResult): string[] {
  const token = singleToken(result);
  return token != null ? [token] : [];
}

/** Copies the named keys from prepare params into the confirm params. */
function pick(params: unknown, keys: readonly string[]): Record<string, unknown> {
  const src = (params ?? {}) as Record<string, unknown>;
  const out: Record<string, unknown> = {};
  for (const key of keys) {
    if (src[key] !== undefined) out[key] = src[key];
  }
  return out;
}

/** Shape A: confirm tool takes a single `{ approval_token }`. */
export function approvalTokenLink(confirmTool: string): ElicitLink {
  return {
    confirmTool,
    collectTokenIds: singleTokenIds,
    buildParams: (_prepareParams, result) => ({ approval_token: singleToken(result) }),
  };
}

/**
 * Shape B: confirm tool takes `{ token_id, ...copyFields }`, where each copied
 * field is carried over verbatim from the prepare params (same field name).
 */
export function tokenIdLink(confirmTool: string, copyFields: readonly string[] = []): ElicitLink {
  return {
    confirmTool,
    collectTokenIds: singleTokenIds,
    buildParams: (prepareParams, result) => ({
      token_id: singleToken(result),
      ...pick(prepareParams, copyFields),
    }),
  };
}

/**
 * Shape C: the batch confirm takes `{ tokens: [{ token_id, email_id }, …] }`,
 * one pair per email. The batch prepare returns `{ tokens: [{ token_id, email:
 * { id, … } }] }` where the token's target is `email.id`, so `email_id` is read
 * from `email.id` (with a direct `email_id` fallback for robustness).
 *
 * `collectTokenIds` gathers EVERY minted `token_id` independently of the id
 * mapping, so a decline invalidates all of them even if a pair is malformed.
 */
export function batchLink(confirmTool: string): ElicitLink {
  return {
    confirmTool,
    collectTokenIds: (result) => batchTokenIds(result),
    buildParams: (_prepareParams, result) => ({ tokens: batchPairs(result) }),
  };
}

/** Raw `tokens` array from a batch prepare result, or []. */
function batchEntries(result: ToolResult): Array<Record<string, unknown>> {
  const tokens = parseResult(result)['tokens'];
  return Array.isArray(tokens) ? (tokens as Array<Record<string, unknown>>) : [];
}

/** Every minted token id in the batch result (authoritative for invalidation). */
function batchTokenIds(result: ToolResult): string[] {
  const ids: string[] = [];
  for (const entry of batchEntries(result)) {
    if (typeof entry['token_id'] === 'string') ids.push(entry['token_id']);
  }
  return ids;
}

interface TokenPair {
  token_id: string;
  email_id: string | number;
}

/** The `{ token_id, email_id }` pairs the batch confirm expects. */
function batchPairs(result: ToolResult): TokenPair[] {
  const pairs: TokenPair[] = [];
  for (const entry of batchEntries(result)) {
    const tokenId = entry['token_id'];
    const email = entry['email'] as Record<string, unknown> | undefined;
    const emailId = email?.['id'] ?? entry['email_id'];
    if (typeof tokenId === 'string' && (typeof emailId === 'string' || typeof emailId === 'number')) {
      pairs.push({ token_id: tokenId, email_id: emailId });
    }
  }
  return pairs;
}
