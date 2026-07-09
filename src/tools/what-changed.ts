/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Delta-sync "what changed" MCP tools (U12).
 *
 * `what_changed` runs Microsoft Graph delta queries against a local mirror and
 * reports what was added / updated / deleted since the previous call. The first
 * call per resource establishes a baseline (no per-item changes). Surfaced ids
 * are durable self-encoding tokens (`em_` mail, `ev_` events).
 *
 * `reset_change_tracking` clears the local cursor + mirror so the next call
 * re-baselines — a local-only operation that never touches the user's mailbox.
 */

import { z } from 'zod';
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolContext, ToolDefinition, ToolResult } from '../registry/types.js';
import type { GraphClient } from '../graph/client/index.js';
import type { StateStore } from '../state/store.js';
import {
  DeltaMirror,
  ALL_RESOURCES,
  type ResourceKey,
  type ChangeEntry,
  type ResourceChangeSet,
} from '../delta/mirror.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    delta: DeltaTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

const ResourceEnum = z.enum(['mail', 'calendar']);

export const WhatChangedInput = z.strictObject({
  resources: z
    .array(ResourceEnum)
    .nonempty()
    .optional()
    .describe('Which resources to check (default: all — mail and calendar).'),
  max_items_per_resource: z
    .number()
    .int()
    .min(1)
    .max(200)
    .optional()
    .describe('Cap the created/updated/deleted lists per resource (default 50).'),
});

export const ResetChangeTrackingInput = z.strictObject({
  resources: z
    .array(ResourceEnum)
    .nonempty()
    .optional()
    .describe('Which resources to reset (default: all). Local-only; no mailbox changes.'),
});

export type WhatChangedParams = z.infer<typeof WhatChangedInput>;
export type ResetChangeTrackingParams = z.infer<typeof ResetChangeTrackingInput>;

// =============================================================================
// Tools
// =============================================================================

/** Trims a change list to `max` entries, recording how many were dropped. */
function cap(entries: ChangeEntry[], max: number): { items: ChangeEntry[]; truncated: number } {
  if (entries.length <= max) return { items: entries, truncated: 0 };
  return { items: entries.slice(0, max), truncated: entries.length - max };
}

export class DeltaTools {
  private readonly mirror: DeltaMirror;

  constructor(client: GraphClient, store: StateStore, accountId: () => string) {
    this.mirror = new DeltaMirror(client, store, accountId);
  }

  async whatChanged(params: WhatChangedParams): Promise<ToolResult> {
    const resources: readonly ResourceKey[] = params.resources ?? ALL_RESOURCES;
    const max = params.max_items_per_resource ?? 50;

    const report = await this.mirror.sync(resources);

    let created = 0;
    let updated = 0;
    let deleted = 0;
    let baselines = 0;

    const resourceViews = report.resources.map((rc: ResourceChangeSet) => {
      created += rc.created.length;
      updated += rc.updated.length;
      deleted += rc.deleted.length;
      if (rc.baseline) baselines += 1;

      const c = cap(rc.created, max);
      const u = cap(rc.updated, max);
      const d = cap(rc.deleted, max);
      const view: Record<string, unknown> = {
        resource: rc.resource,
        baseline: rc.baseline,
        tracked_count: rc.trackedCount,
        created: c.items,
        updated: u.items,
        deleted: d.items,
      };
      const truncated = c.truncated + u.truncated + d.truncated;
      if (truncated > 0) view.truncated = truncated;
      if (rc.note != null) view.note = rc.note;
      return view;
    });

    const summary = baselines === report.resources.length && created + updated + deleted === 0
      ? `Baseline established for ${baselines} resource(s); call again to see changes.`
      : `${created} created, ${updated} updated, ${deleted} deleted since last sync.`;

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          synced_at: new Date(report.syncedAt).toISOString(),
          summary,
          resources: resourceViews,
        }, null, 2),
      }],
    };
  }

  resetChangeTracking(params: ResetChangeTrackingParams): ToolResult {
    const resources: readonly ResourceKey[] = params.resources ?? ALL_RESOURCES;
    this.mirror.reset(resources);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          success: true,
          reset: resources,
          message: 'Change tracking reset; the next what_changed call will re-baseline.',
        }, null, 2),
      }],
    };
  }
}

// =============================================================================
// Registry Definitions
// =============================================================================

/** Registry tool definitions for the delta-sync domain. */
export function deltaToolDefinitions(): ToolDefinition[] {
  const tools = (ctx: ToolContext): DeltaTools => requireGraphToolset(ctx, 'delta');

  return [
    defineTool({
      name: 'what_changed',
      description:
        'Report mailbox/calendar items added, updated, or deleted since the last check, using Graph delta sync against a local mirror. First call per resource sets a baseline. (Graph API)',
      input: WhatChangedInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['mail', 'calendar'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).whatChanged(params),
    }),
    defineTool({
      name: 'reset_change_tracking',
      description:
        'Clear the local delta-sync cursor and mirror so the next what_changed call re-baselines. Local-only; does not modify the mailbox. (Graph API)',
      input: ResetChangeTrackingInput,
      annotations: { readOnlyHint: false, idempotentHint: true },
      destructive: false,
      presets: ['mail', 'calendar'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).resetChangeTracking(params),
    }),
  ];
}
