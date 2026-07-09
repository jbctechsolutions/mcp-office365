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

/** Which backend(s) expose a tool. Graph is the only backend (AppleScript was removed). */
export type Backend = 'graph';

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
  | 'meetings'
  | 'shared';

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
 * is lazily initialized.
 *
 * This is an augmentation target: each domain module declares its own field
 * via `declare module '../registry/types.js'`, so migrating a domain never
 * edits this file — avoiding a 25-domain merge-conflict hotspot. Handlers read
 * their instance through `requireGraphToolset(ctx, key)`.
 */
// eslint-disable-next-line @typescript-eslint/no-empty-interface
export interface GraphToolsets {}

/** How a destructive `prepare_*` step seeks confirmation (U11). */
export type ConfirmMode = 'token' | 'elicit';

/** Outcome of an inline elicitation (U11). */
export type ElicitOutcome =
  /** The user explicitly confirmed — execute now. */
  | 'accept'
  /** The user explicitly said no — abort AND invalidate the token. */
  | 'decline'
  /** Cancelled, timed out, or the client can't elicit — fall back to the token flow. */
  | 'degrade';

/** A yes/no confirmation request sent to the client (U11). */
export interface ElicitRequest {
  /** Human-readable description of the action being confirmed. */
  readonly message: string;
}

/** Sends an inline confirmation to the client, resolving to its outcome (U11). */
export type Elicitor = (request: ElicitRequest) => Promise<ElicitOutcome>;

/**
 * Runtime context passed to a tool handler. Built once per call after the
 * backend is initialized; carries the live toolset instances so handlers no
 * longer need the 23-positional-parameter dispatch signature.
 */
export interface ToolContext {
  readonly backend: Backend;
  readonly tokenManager: ApprovalTokenManager;
  readonly graph: GraphToolsets | null;
  /** Confirmation mode for destructive prepares (default 'token' when absent). */
  readonly confirmMode?: ConfirmMode;
  /** Inline elicitor, present only when the client advertised the capability. */
  readonly elicit?: Elicitor;
}

/**
 * Links a `prepare_*` tool to its `confirm_*` counterpart so the dispatch-level
 * elicitation interceptor (U11) can execute inline on an accepted confirmation.
 * Declarative on purpose: no handler body changes, and the confirm tool's exact
 * validation + execution is reused.
 *
 * @typeParam S - the prepare tool's Zod input schema.
 */
export interface ElicitLink<S extends z.ZodType = z.ZodType> {
  /** Name of the paired confirm tool (e.g. 'confirm_delete_channel'). */
  readonly confirmTool: string;
  /**
   * Builds the confirm tool's params from the prepare params + the prepare
   * result (which carries the freshly minted token(s)). Handles heterogeneous
   * confirm shapes: plain `approval_token`, `token_id` + target id, and the
   * batch confirm that takes an array of `{ token_id, email_id }` pairs.
   */
  readonly buildParams: (prepareParams: z.infer<S>, prepareResult: ToolResult) => unknown;
  /**
   * Token ids minted by this prepare, so a declined action can invalidate them.
   * Empty means "no token in the result" — dispatch then degrades without
   * eliciting.
   */
  readonly collectTokenIds: (prepareResult: ToolResult) => string[];
  /** Optional custom confirmation prompt; a generic one is derived when absent. */
  readonly message?: (prepareParams: z.infer<S>) => string;
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
  /**
   * Present on `prepare_*` tools to opt into inline elicitation (U11). When the
   * server runs with `--confirm elicit` and the client can elicit, dispatch
   * runs the prepare, asks the user inline, and on accept executes the linked
   * confirm tool — otherwise it degrades to the normal token response.
   *
   * `NoInfer` keeps this field out of `S` inference so the generic is fixed by
   * `input` alone — otherwise a helper returning the default `ElicitLink` would
   * widen `S` and collapse the handler's `params` to `unknown`.
   */
  readonly onElicit?: ElicitLink<NoInfer<S>>;
}
