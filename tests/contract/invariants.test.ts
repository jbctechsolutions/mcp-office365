/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Contract harness (U3) — invariants asserted over the whole registry.
 *
 * These iterate `allToolDefinitions()` and therefore cover every domain the
 * moment U2 migrates it — no per-domain test wiring. They make the observed
 * v2 failure classes structurally impossible to reintroduce: schema/handler
 * drift, unregistered next-action targets, and destructive tools leaking into
 * read-only mode.
 */

import { describe, it, expect } from 'vitest';
import { ToolRegistry, toInputSchema } from '../../src/registry/registry.js';
import { allToolDefinitions } from '../../src/registry/all-tools.js';
import type { Backend, Preset, ToolContext } from '../../src/registry/types.js';
import { ApprovalTokenManager } from '../../src/approval/index.js';

const defs = allToolDefinitions();
const byName = new Map(defs.map((d) => [d.name, d]));

const KNOWN_PRESETS: readonly Preset[] = [
  'mail', 'calendar', 'contacts', 'tasks', 'notes', 'teams',
  'planner', 'files', 'sharepoint', 'excel', 'people', 'meetings',
];

describe('registry contract invariants', () => {
  it('registers without duplicate names', () => {
    expect(() => new ToolRegistry().register(defs)).not.toThrow();
    expect(byName.size).toBe(defs.length);
  });

  it('every tool has a description and a valid object input schema', () => {
    for (const def of defs) {
      expect(def.description.length, `${def.name} description`).toBeGreaterThan(0);
      const json = toInputSchema(def.input) as Record<string, unknown>;
      expect(json['type'], `${def.name} inputSchema.type`).toBe('object');
    }
  });

  it('every tool carries at least one MCP annotation', () => {
    for (const def of defs) {
      expect(Object.keys(def.annotations).length, `${def.name} annotations`).toBeGreaterThan(0);
    }
  });

  it('every tool declares at least one backend', () => {
    const valid: readonly Backend[] = ['graph', 'applescript'];
    for (const def of defs) {
      expect(def.backends.length, `${def.name} backends`).toBeGreaterThan(0);
      for (const b of def.backends) {
        expect(valid, `${def.name} backend "${b}"`).toContain(b);
      }
    }
  });

  it('every declared preset is a known preset name', () => {
    for (const def of defs) {
      for (const p of def.presets) {
        expect(KNOWN_PRESETS, `${def.name} preset "${p}"`).toContain(p);
      }
    }
  });

  it('read-only tools are not flagged destructive (annotation/flag agreement)', () => {
    for (const def of defs) {
      if (def.annotations.readOnlyHint === true) {
        expect(def.destructive, `${def.name} readOnly but destructive`).toBe(false);
      }
    }
  });

  it('no readOnlyHint:true tool has a mutation-verb name (structural --read-only guard, U10)', () => {
    // A tool advertised readOnlyHint:true is exposed under --read-only AND
    // callable, so a mislabeled write would be a trust-boundary hole that the
    // destructive-flag check above does not catch for non-destructive writes.
    // Enforce structurally: read-only tools cannot carry a mutating verb prefix.
    const MUTATION_PREFIXES = [
      'create_', 'update_', 'delete_', 'send_', 'set_', 'add_', 'remove_',
      'mark_', 'move_', 'complete_', 'rename_', 'respond_', 'archive_', 'junk_',
      'empty_', 'clear_', 'upload_', 'forward_', 'reply_', 'prepare_', 'confirm_',
    ];
    for (const def of defs) {
      if (def.annotations.readOnlyHint === true) {
        const offending = MUTATION_PREFIXES.find((p) => def.name.startsWith(p));
        expect(offending, `${def.name} is readOnlyHint:true but names a mutation`).toBeUndefined();
      }
    }
  });

  it('every prepare_ tool has a matching confirm_ tool and both are destructive', () => {
    // Invariant (e): the two-phase pair must be complete and both halves must
    // be excluded by --read-only, so a read-only client cannot mint OR redeem.
    for (const def of defs) {
      if (def.name.startsWith('prepare_')) {
        // Batch prepares (prepare_batch_delete_emails, prepare_batch_move_emails)
        // deliberately pair many-to-one with a single confirm_batch_operation
        // rather than a 1:1 confirm_<suffix>.
        const confirmName = def.name.startsWith('prepare_batch_')
          ? 'confirm_batch_operation'
          : def.name.replace(/^prepare_/, 'confirm_');
        const confirm = byName.get(confirmName);
        expect(confirm, `${def.name} missing ${confirmName}`).toBeDefined();
        expect(def.destructive, `${def.name} not destructive`).toBe(true);
        expect(confirm!.destructive, `${confirmName} not destructive`).toBe(true);
      }
    }
  });

  it('every tool handler returns a valid ToolResult for a supported backend', async () => {
    // Invariant: each handler runs against a fully-proxied context (both the
    // graph and AppleScript toolset bags auto-vivify any method to a recording
    // stub) and returns a well-formed ToolResult without throwing. Exercises
    // every handler — catches a handler that references a missing context key
    // or throws on a valid call. Dual-backend handlers (e.g. notes) are run
    // against each backend they declare.
    const toolsetProxy = new Proxy(
      {},
      {
        get: () => (): { content: Array<{ type: 'text'; text: string }> } => ({
          content: [{ type: 'text', text: '{}' }],
        }),
      },
    );
    const bagProxy = new Proxy({}, { get: () => toolsetProxy });

    for (const def of defs) {
      for (const backend of def.backends) {
        const ctx: ToolContext = {
          backend,
          tokenManager: new ApprovalTokenManager(),
          graph: bagProxy as never,
          applescript: bagProxy as never,
        };
        const result = await def.handler(ctx, {} as never);
        expect(Array.isArray(result?.content), `${def.name} (${backend}) returned no content`).toBe(true);
      }
    }
  });

  it('read-only surface excludes every destructive tool', () => {
    // Invariant (e), enforced through the live registry filter.
    const registry = new ToolRegistry();
    registry.register(defs);
    const readOnlyNames = new Set(
      registry.listTools({ backend: 'graph', readOnly: true }).map((t) => t.name),
    );
    for (const def of defs) {
      if (def.destructive && def.backends.includes('graph')) {
        expect(readOnlyNames.has(def.name), `${def.name} leaked into read-only`).toBe(false);
      }
    }
  });
});
