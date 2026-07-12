/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * ToolRegistry — single source of truth for the MCP tool surface (U1).
 *
 * Holds tool definitions, derives the advertised JSON Schema from each Zod
 * schema, filters the surface by backend / preset / read-only, and dispatches
 * a call to the matching handler after validating its arguments.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import { ReadOnlyModeError } from '../utils/errors.js';
import type {
  Backend,
  Preset,
  ToolContext,
  ToolDefinition,
  ToolResult,
  ElicitLink,
  Elicitor,
} from './types.js';

/** Options controlling which tools are exposed to a client. */
export interface SurfaceOptions {
  readonly backend: Backend;
  /** Presets to include. Undefined/empty means the full surface. */
  readonly presets?: readonly Preset[];
  /** When true, only non-destructive tools are exposed. */
  readonly readOnly?: boolean;
  /**
   * Explicit allow-list of tool names (U6 entitlements). When set, the surface
   * is EXACTLY these tools — presets are ignored and the empty-presets
   * "always-exposed" bypass does not apply (a tool shows iff it is listed). This
   * is what makes a per-user entitlement pinnable: a server upgrade that adds a
   * tool cannot widen a user's surface unless the tool is added to their list.
   */
  readonly allow?: readonly string[];
  /** Tool names to remove from the surface — applies in both allow and preset modes. */
  readonly exclude?: readonly string[];
}

/**
 * Converts a Zod schema to an MCP-compatible `inputSchema`.
 *
 * Uses zod 4's native `z.toJSONSchema` (draft-07 target) and strips the
 * `$schema` key, which MCP's `Tool.inputSchema` does not carry. Transitional
 * alias keys (U6/D11) are handled by the input pipeline, not here; when a
 * schema opts into aliases it should expose them as optional properties so the
 * advertised schema stays validation-safe.
 */
export function toInputSchema(schema: z.ZodType): Tool['inputSchema'] {
  // `io: 'input'` — the advertised schema describes what a CLIENT sends, so a
  // field with a `.default()` is optional (the server fills the default at
  // parse time). Zod 4's default `io: 'output'` would mark defaulted fields as
  // required, misleading agents into always supplying them.
  const json = z.toJSONSchema(schema, { target: 'draft-7', io: 'input' }) as Record<string, unknown>;
  delete json['$schema'];
  if (json['type'] == null) {
    json['type'] = 'object';
  }
  return json as Tool['inputSchema'];
}

/**
 * In-memory registry of tool definitions.
 */
export class ToolRegistry {
  private readonly definitions = new Map<string, ToolDefinition>();

  /** Registers a batch of tool definitions. Throws on a duplicate name. */
  register(defs: readonly ToolDefinition[]): void {
    for (const def of defs) {
      if (this.definitions.has(def.name)) {
        throw new Error(`Duplicate tool registration: "${def.name}"`);
      }
      this.definitions.set(def.name, def);
    }
  }

  /** True when a tool with this name is registered. */
  has(name: string): boolean {
    return this.definitions.has(name);
  }

  /** All registered names, insertion order preserved. */
  names(): string[] {
    return [...this.definitions.keys()];
  }

  /** Returns a single definition, or undefined. */
  get(name: string): ToolDefinition | undefined {
    return this.definitions.get(name);
  }

  /**
   * Returns the MCP tool list for a given surface, filtered by backend,
   * preset membership, and read-only mode.
   */
  listTools(options: SurfaceOptions): Tool[] {
    return this.filtered(options).map((def) => {
      const tool: Tool = {
        name: def.name,
        description: def.description,
        inputSchema: toInputSchema(def.input),
      };
      if (Object.keys(def.annotations).length > 0) {
        return { ...tool, annotations: def.annotations };
      }
      return tool;
    });
  }

  /** True when a tool is exposed under the given surface. */
  isExposed(name: string, options: SurfaceOptions): boolean {
    const def = this.definitions.get(name);
    return def != null && this.matches(def, options);
  }

  /**
   * Validates arguments against the tool's schema and invokes its handler.
   * Returns undefined when no tool with this name is registered, so the caller
   * can fall back to the legacy dispatch during the U2 migration window.
   */
  async dispatch(
    name: string,
    args: unknown,
    ctx: ToolContext,
    options: SurfaceOptions,
  ): Promise<ToolResult | undefined> {
    const def = this.definitions.get(name);
    if (def == null) {
      return undefined;
    }
    if (!this.matches(def, options)) {
      // Distinguish the read-only rejection so it surfaces a stable
      // READ_ONLY_MODE envelope (D13) — but only when read-only is the *sole*
      // reason (the tool would otherwise be exposed). A tool also filtered by
      // backend/preset gets the generic error, avoiding misattribution.
      if (options.readOnly === true && this.matches(def, { ...options, readOnly: false })) {
        throw new ReadOnlyModeError(name);
      }
      throw new Error(`Tool "${name}" is not available in the current mode.`);
    }
    // MCP marks `arguments` optional; default to {} so no-input tools
    // (e.g. list_mail_rules) parse cleanly, matching the codebase-wide
    // `parse(args ?? {})` convention.
    const params = def.input.parse(args ?? {});

    // Inline-elicitation interceptor (U11): only for a prepare tool that opted in
    // (onElicit), when the server is in elicit mode and the client can elicit.
    // Everything else runs the handler directly — an unchanged fast path.
    if (def.onElicit != null && ctx.confirmMode === 'elicit' && ctx.elicit != null) {
      return this.dispatchWithElicitation(def, params, ctx, def.onElicit, ctx.elicit);
    }
    return def.handler(ctx, params);
  }

  /**
   * Runs a prepare tool, asks the user inline, and resolves the outcome (U11):
   * - accept  → execute now via the linked confirm tool (its exact validation +
   *             execution is reused; the just-minted token is consumed there).
   * - decline → invalidate the token (consume-without-execute) and report abort.
   * - degrade → return the prepare token response unchanged (cancel / timeout /
   *             no capability / unparseable token) — the normal two-phase flow.
   *
   * Fail-closed: the destructive action runs only on an explicit accept.
   */
  private async dispatchWithElicitation(
    def: ToolDefinition,
    prepareParams: unknown,
    ctx: ToolContext,
    link: ElicitLink,
    elicit: Elicitor,
  ): Promise<ToolResult> {
    const prepareResult = await def.handler(ctx, prepareParams);
    const tokenIds = link.collectTokenIds(prepareResult);
    // No token in the prepare output → nothing to gate/execute; degrade.
    if (tokenIds.length === 0) {
      return prepareResult;
    }

    const message = link.message?.(prepareParams) ?? defaultConfirmMessage(ctx, tokenIds[0]);
    const outcome = await elicit({ message });

    if (outcome === 'accept') {
      const confirmDef = this.definitions.get(link.confirmTool);
      // A misconfigured link must not silently execute nor lose the approval —
      // degrade to the token response so the two-phase path still works.
      if (confirmDef == null) {
        return prepareResult;
      }
      // Only the param mapping/validation is guarded: a bad buildParams must
      // degrade (return the still-live token) rather than error out an accepted
      // action. The confirm handler runs OUTSIDE the guard so a genuine
      // execution error (e.g. a Graph failure) surfaces normally, not masked.
      let confirmParams: unknown;
      try {
        confirmParams = confirmDef.input.parse(link.buildParams(prepareParams, prepareResult));
      } catch {
        return prepareResult;
      }
      return confirmDef.handler(ctx, confirmParams);
    }

    if (outcome === 'decline') {
      // Burn every token this prepare minted (batch mints one per item), so a
      // declined action can't be redeemed via the confirm flow afterward.
      for (const tokenId of tokenIds) {
        invalidateToken(ctx, tokenId);
      }
      return declinedResult(def.name);
    }

    // 'degrade' — hand back the durable token exactly as the token flow would.
    return prepareResult;
  }

  private filtered(options: SurfaceOptions): ToolDefinition[] {
    return [...this.definitions.values()].filter((def) => this.matches(def, options));
  }

  private matches(def: ToolDefinition, options: SurfaceOptions): boolean {
    if (!def.backends.includes(options.backend)) {
      return false;
    }
    if (options.readOnly === true && !isReadOnly(def)) {
      return false;
    }
    // Exclusion wins over any inclusion (U6).
    if (options.exclude != null && options.exclude.includes(def.name)) {
      return false;
    }
    // Explicit allow-list mode (U6): the surface is exactly the listed tools;
    // presets and the empty-presets bypass do not apply.
    if (options.allow != null) {
      return options.allow.includes(def.name);
    }
    const presets = options.presets;
    if (presets != null && presets.length > 0 && def.presets.length > 0) {
      return def.presets.some((p) => presets.includes(p));
    }
    return true;
  }
}

/**
 * A tool is read-only iff it explicitly advertises `readOnlyHint: true`. This is
 * stricter than `!destructive` — a non-destructive write (e.g. create_draft,
 * mark_email_read) still mutates state and must be excluded from `--read-only`.
 */
function isReadOnly(def: ToolDefinition): boolean {
  return def.annotations.readOnlyHint === true;
}

/**
 * Burns a token without executing its action, so a declined destructive request
 * can't be redeemed via the confirm flow afterward. Uses the token's own
 * operation/targetId (it was just minted, so this always validates), reusing the
 * store's atomic single-use consume — no new persistence machinery.
 */
function invalidateToken(ctx: ToolContext, tokenId: string): void {
  const token = ctx.tokenManager.lookupToken(tokenId);
  if (token != null) {
    ctx.tokenManager.consumeToken(tokenId, token.operation, token.targetId);
  }
}

/** A generic confirmation prompt derived from the token's operation. */
function defaultConfirmMessage(ctx: ToolContext, tokenId: string | undefined): string {
  const op = tokenId != null ? ctx.tokenManager.lookupToken(tokenId)?.operation : undefined;
  const action = op != null ? op.replace(/_/g, ' ') : 'this action';
  return `Confirm: ${action}? This cannot be undone.`;
}

/** The result returned when a user declines an inline confirmation. */
function declinedResult(toolName: string): ToolResult {
  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        success: false,
        declined: true,
        message: `Action cancelled: you declined the confirmation for ${toolName}. The approval token has been invalidated.`,
      }, null, 2),
    }],
  };
}
