/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Runtime-context helpers for registry tool handlers.
 */

import type { AppleScriptToolsets, GraphToolsets, ToolContext } from './types.js';

/**
 * Resolves an initialized Graph-backend toolset from the runtime context, or
 * throws a clear error when the Graph backend is unavailable.
 *
 * Centralizes the per-domain "resolve toolset or throw" boilerplate so every
 * migrated domain shares one implementation instead of hand-rolling a null
 * check. The `key` is type-checked against the augmented `GraphToolsets`, and
 * the return type is the exact toolset class for that key.
 */
export function requireGraphToolset<K extends keyof GraphToolsets>(
  ctx: ToolContext,
  key: K,
): GraphToolsets[K] {
  if (ctx.graph == null) {
    throw new Error('This tool requires the Microsoft Graph API backend.');
  }
  return ctx.graph[key];
}

/**
 * Resolves an initialized AppleScript-backend toolset from the runtime context,
 * or throws when the AppleScript backend is unavailable. The dual-backend
 * counterpart to {@link requireGraphToolset}.
 */
export function requireAppleScriptToolset<K extends keyof AppleScriptToolsets>(
  ctx: ToolContext,
  key: K,
): AppleScriptToolsets[K] {
  if (ctx.applescript == null) {
    throw new Error('This tool requires the AppleScript (classic Outlook) backend.');
  }
  return ctx.applescript[key];
}
