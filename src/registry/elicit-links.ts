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
 * one pair per email — read straight from the prepare result (the prepare mints
 * a token per id and returns the pairs).
 */
export function batchLink(confirmTool: string): ElicitLink {
  return {
    confirmTool,
    collectTokenIds: (result) => tokenPairs(result).map((pair) => pair.token_id),
    buildParams: (_prepareParams, result) => ({ tokens: tokenPairs(result) }),
  };
}

interface TokenPair {
  token_id: string;
  email_id: string | number;
}

/** Extracts the `{ token_id, email_id }` pairs from a batch prepare result. */
function tokenPairs(result: ToolResult): TokenPair[] {
  const tokens = parseResult(result)['tokens'];
  if (!Array.isArray(tokens)) return [];
  const pairs: TokenPair[] = [];
  for (const entry of tokens) {
    const rec = entry as Record<string, unknown>;
    const tokenId = rec['token_id'];
    const emailId = rec['email_id'];
    if (typeof tokenId === 'string' && (typeof emailId === 'string' || typeof emailId === 'number')) {
      pairs.push({ token_id: tokenId, email_id: emailId });
    }
  }
  return pairs;
}
