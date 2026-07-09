/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Inline confirmation via MCP elicitation (U11).
 *
 * Bridges the MCP `Server.elicitInput` request to the registry's backend-neutral
 * {@link Elicitor} contract. Capability-gated and fail-open-to-degrade: if the
 * client can't elicit, or the request times out or errors, the caller falls back
 * to the durable two-phase token flow — an inline confirmation never blocks or
 * fails a tool call.
 */

import type { Server } from '@modelcontextprotocol/sdk/server/index.js';
import type { Elicitor, ElicitOutcome } from './types.js';

/** How long to wait for the user's inline yes/no before degrading (U11). */
export const ELICIT_TIMEOUT_MS = 60_000;

/**
 * A minimal confirmation form. We key the decision off the elicitation `action`
 * ('accept' = yes), so the schema is a formality some clients render; a single
 * optional boolean keeps it valid and unobtrusive.
 */
const CONFIRM_SCHEMA = {
  type: 'object' as const,
  properties: {
    confirm: {
      type: 'boolean' as const,
      title: 'Confirm',
      description: 'Approve this destructive action.',
    },
  },
  required: [] as string[],
};

/**
 * Builds an {@link Elicitor} bound to a live MCP server. Returns 'degrade' —
 * never throws — when the client lacks the capability, cancels, or times out, so
 * the dispatch interceptor can hand back the durable token unchanged.
 */
export function createServerElicitor(server: Server): Elicitor {
  return async ({ message }): Promise<ElicitOutcome> => {
    if (server.getClientCapabilities()?.elicitation == null) {
      return 'degrade';
    }
    try {
      const result = await server.elicitInput(
        { message, requestedSchema: CONFIRM_SCHEMA },
        { timeout: ELICIT_TIMEOUT_MS },
      );
      switch (result.action) {
        case 'accept':
          return 'accept';
        case 'decline':
          return 'decline';
        default:
          // 'cancel' — the user dismissed without deciding; leave the token.
          return 'degrade';
      }
    } catch {
      // Timeout or transport error → degrade (fail-open to the token flow).
      return 'degrade';
    }
  };
}
