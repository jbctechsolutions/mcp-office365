/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tool registry types (v3 registry-driven architecture, U1).
 *
 * A tool definition is the single source of truth for what the four
 * previously hand-maintained lists carried independently: the JSON-schema
 * `TOOLS` array, per-module Zod schemas, the dispatch switch, and
 * `GRAPH_ONLY_TOOL_NAMES`. Definitions register their static metadata
 * eagerly (before any backend is initialized, so ListTools works without
 * auth) and resolve their live handler dependencies lazily via ToolContext.
 */

import type { z } from 'zod';
import type { ApprovalTokenManager } from '../approval/index.js';
// Type-only imports — erased at runtime, so the registry ↔ tools cycle is safe.
import type { MailRulesTools } from '../tools/mail-rules.js';

/** Which backend(s) expose a tool. */
export type Backend = 'graph' | 'applescript';

/** Preset groups a client can select with --preset to shrink the surface. */
export type Preset =
  | 'mail'
  | 'calendar'
  | 'contacts'
  | 'tasks'
  | 'notes'
  | 'teams'
  | 'planner'
  | 'files'
  | 'sharepoint'
  | 'excel'
  | 'people'
  | 'meetings';

/**
 * MCP tool annotations (spec 2025-03-26+). Clients use these to auto-approve
 * safe reads and gate destructive calls.
 */
export interface ToolAnnotations {
  readonly title?: string;
  readonly readOnlyHint?: boolean;
  readonly destructiveHint?: boolean;
  readonly idempotentHint?: boolean;
  readonly openWorldHint?: boolean;
}

/** The MCP text-content result shape every tool returns. */
export interface ToolResult {
  content: Array<{ type: 'text'; text: string }>;
  isError?: boolean;
}

/**
 * Bag of initialized Graph-backend tool instances, resolved after the backend
 * is lazily initialized. Grows one field per domain as U2 migrates them.
 */
export interface GraphToolsets {
  readonly rules: MailRulesTools;
}

/**
 * Bag of initialized AppleScript-backend tool instances. The AppleScript
 * backend is frozen (v3): registry-listed, best-effort metadata, no new
 * migrations. Grows only if a frozen domain is registry-migrated.
 */
export interface AppleScriptToolsets {
  readonly _frozen?: never;
}

/**
 * Runtime context passed to a tool handler. Built once per call after the
 * backend is initialized; carries the live toolset instances so handlers no
 * longer need the 23-positional-parameter dispatch signature.
 */
export interface ToolContext {
  readonly backend: Backend;
  readonly tokenManager: ApprovalTokenManager;
  readonly graph: GraphToolsets | null;
  readonly applescript: AppleScriptToolsets | null;
}

/**
 * A single tool definition — static metadata plus a lazily-bound handler.
 *
 * @typeParam S - the Zod input schema; the handler receives its parsed output.
 */
export interface ToolDefinition<S extends z.ZodType = z.ZodType> {
  /** Tool name exposed over MCP (snake_case). */
  readonly name: string;
  /** One-line description shown to the model. */
  readonly description: string;
  /** Zod input schema; the advertised JSON Schema is derived from it. */
  readonly input: S;
  /** MCP annotations surfaced in ListTools. */
  readonly annotations: ToolAnnotations;
  /**
   * True when the tool participates in a destructive flow (delete/send/upload,
   * including the prepare_/confirm_ halves). Drives --read-only filtering.
   */
  readonly destructive: boolean;
  /** Presets that include this tool. Empty means "always exposed". */
  readonly presets: readonly Preset[];
  /** Backends that expose this tool. */
  readonly backends: readonly Backend[];
  /** Handler invoked with the runtime context and validated params. */
  readonly handler: (ctx: ToolContext, params: z.infer<S>) => Promise<ToolResult> | ToolResult;
}
