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
} from './types.js';

/** Options controlling which tools are exposed to a client. */
export interface SurfaceOptions {
  readonly backend: Backend;
  /** Presets to include. Undefined/empty means the full surface. */
  readonly presets?: readonly Preset[];
  /** When true, only non-destructive tools are exposed. */
  readonly readOnly?: boolean;
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
    return def.handler(ctx, params);
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
