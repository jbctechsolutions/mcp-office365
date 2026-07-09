/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Delta-sync mirror orchestration (U12).
 *
 * Drives Microsoft Graph delta queries for a small set of resources (inbox
 * mail, calendar events), maintains a local mirror of seen items via
 * {@link DeltaStore}, and computes a change set (added / updated / deleted)
 * since the previous sync. Ids it surfaces are minted as durable self-encoding
 * tokens (`em_` for mail, `ev_` for events) so a caller can address them with
 * the same handles every other tool returns.
 *
 * The first sync of a resource is a *baseline*: it records the current state
 * without reporting every existing item as "created", so `what_changed` only
 * reports real deltas from the second call onward.
 */

import type { GraphClient } from '../graph/client/index.js';
import type { StateStore } from '../state/store.js';
import type { MirrorItem } from '../state/delta-store.js';
import { mintSelfEncoded, type EntityType } from '../ids/token.js';

/** Resources the mirror can track. */
export type ResourceKey = 'mail' | 'calendar';

/** How far forward/back the initial calendar-view window spans (baked into the cursor). */
const CALENDAR_WINDOW_PAST_MS = 1 * 24 * 60 * 60 * 1000;
const CALENDAR_WINDOW_FUTURE_MS = 90 * 24 * 60 * 60 * 1000;
/**
 * The calendar cursor bakes an absolute time window that cannot advance, so it
 * is force-refreshed (re-baselined with a fresh window) once it is this old —
 * bounding how far the tracked window can drift from "now".
 */
const CALENDAR_CURSOR_MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000;

/** A single change surfaced to the caller. */
export interface ChangeEntry {
  token: string;
  graphId: string;
  summary: string;
  changeType: 'created' | 'updated' | 'deleted';
  detail?: Record<string, unknown>;
}

/** The change set for one resource in a sync round. */
export interface ResourceChangeSet {
  resource: ResourceKey;
  /** True when this call established the baseline (no per-item changes reported). */
  baseline: boolean;
  /** Total items now mirrored for the resource. */
  trackedCount: number;
  created: ChangeEntry[];
  updated: ChangeEntry[];
  deleted: ChangeEntry[];
  /** Optional operator note (e.g. cursor unavailable). */
  note?: string;
}

/** The full report from a {@link DeltaMirror.sync} call. */
export interface ChangeReport {
  syncedAt: number;
  resources: ResourceChangeSet[];
}

/** A normalized delta entry, backend-agnostic. */
interface RawDelta {
  graphId: string;
  removed: boolean;
  /** `@removed.reason`: 'deleted' (gone) vs 'changed' (left this view). */
  removedReason: string;
  summary: string;
  snapshot: Record<string, unknown>;
}

interface ResourceDescriptor {
  key: ResourceKey;
  /** Storage key in the delta tables (namespaced, room for future scoping). */
  storageKey: string;
  entityType: EntityType;
  /** When set, a cursor older than this is force-refreshed with a fresh window. */
  maxCursorAgeMs?: number;
  fetch(
    client: GraphClient,
    deltaLink: string | undefined,
    now: number,
  ): Promise<{ items: RawDelta[]; deltaLink: string }>;
}

/** Reads `@removed` (and its reason) regardless of the concrete Graph type. */
function removal(item: unknown): { removed: boolean; reason: string } {
  const marker = (item as Record<string, unknown>)['@removed'] as { reason?: string } | null | undefined;
  if (marker == null) return { removed: false, reason: '' };
  return { removed: true, reason: marker.reason ?? 'deleted' };
}

const MAIL_RESOURCE: ResourceDescriptor = {
  key: 'mail',
  storageKey: 'mail:inbox',
  entityType: 'message',
  async fetch(client, deltaLink) {
    const { messages, deltaLink: next } = await client.getMessagesDelta('inbox', deltaLink);
    const items: RawDelta[] = [];
    for (const m of messages) {
      const graphId = m.id ?? '';
      if (graphId.length === 0) continue;
      const { removed, reason } = removal(m);
      items.push({
        graphId,
        removed,
        removedReason: reason,
        summary: m.subject ?? '(no subject)',
        snapshot: {
          from: m.from?.emailAddress?.address ?? '',
          receivedDateTime: m.receivedDateTime ?? '',
          isRead: m.isRead ?? null,
        },
      });
    }
    return { items, deltaLink: next };
  },
};

const CALENDAR_RESOURCE: ResourceDescriptor = {
  key: 'calendar',
  storageKey: 'calendar:primary',
  entityType: 'event',
  maxCursorAgeMs: CALENDAR_CURSOR_MAX_AGE_MS,
  async fetch(client, deltaLink, now) {
    const start = new Date(now - CALENDAR_WINDOW_PAST_MS).toISOString();
    const end = new Date(now + CALENDAR_WINDOW_FUTURE_MS).toISOString();
    const { events, deltaLink: next } = await client.getCalendarViewDelta(start, end, deltaLink);
    const items: RawDelta[] = [];
    for (const e of events) {
      const graphId = e.id ?? '';
      if (graphId.length === 0) continue;
      const { removed, reason } = removal(e);
      items.push({
        graphId,
        removed,
        removedReason: reason,
        summary: e.subject ?? '(no subject)',
        snapshot: {
          start: e.start?.dateTime ?? '',
          end: e.end?.dateTime ?? '',
          organizer: e.organizer?.emailAddress?.address ?? '',
        },
      });
    }
    return { items, deltaLink: next };
  },
};

const RESOURCES: Record<ResourceKey, ResourceDescriptor> = {
  mail: MAIL_RESOURCE,
  calendar: CALENDAR_RESOURCE,
};

/** All resource keys, in a stable order. */
export const ALL_RESOURCES: readonly ResourceKey[] = ['mail', 'calendar'];

/**
 * Orchestrates delta sync + change classification against a {@link StateStore}
 * mirror. Stateless beyond its dependencies; safe to construct per call.
 */
export class DeltaMirror {
  constructor(
    private readonly client: GraphClient,
    private readonly store: StateStore,
    private readonly accountId: () => string,
    private readonly now: () => number = () => Date.now(),
  ) {}

  /** Syncs the given resources (all by default) and returns the change report. */
  async sync(resources: readonly ResourceKey[] = ALL_RESOURCES): Promise<ChangeReport> {
    const syncedAt = this.now();
    const out: ResourceChangeSet[] = [];
    for (const key of resources) {
      out.push(await this.syncResource(RESOURCES[key], syncedAt));
    }
    return { syncedAt, resources: out };
  }

  /** Clears tracking for the given resources (all by default) — local only. */
  reset(resources: readonly ResourceKey[] = ALL_RESOURCES): void {
    const accountId = this.accountId();
    for (const key of resources) {
      this.store.delta.reset(accountId, RESOURCES[key].storageKey);
    }
  }

  private async syncResource(desc: ResourceDescriptor, syncedAt: number): Promise<ResourceChangeSet> {
    const accountId = this.accountId();
    const cursor = this.store.delta.getCursor(accountId, desc.storageKey);

    // A stale calendar-style cursor (fixed window that can no longer advance) is
    // treated as absent so this round re-baselines with a fresh window.
    const stale =
      cursor != null &&
      desc.maxCursorAgeMs != null &&
      syncedAt - cursor.syncedAt > desc.maxCursorAgeMs;
    const prevLink = stale ? undefined : (cursor?.deltaLink ?? undefined);
    const baseline = prevLink == null;

    const { items, deltaLink } = await desc.fetch(this.client, prevLink, syncedAt);
    // Collapse duplicate ids within one response (last entry wins, per Graph
    // ordering) so a created-then-deleted or repeated id can't be double-counted
    // or leave the report inconsistent with the committed mirror.
    const latest = new Map<string, RawDelta>();
    for (const item of items) latest.set(item.graphId, item);

    const seen = baseline ? new Set<string>() : this.store.delta.getSeenIds(accountId, desc.storageKey);

    const upserts: MirrorItem[] = [];
    const deletes: string[] = [];
    const created: ChangeEntry[] = [];
    const updated: ChangeEntry[] = [];
    const deleted: ChangeEntry[] = [];

    for (const item of latest.values()) {
      const token = mintSelfEncoded(desc.entityType, item.graphId);
      if (item.removed) {
        // Drop from the mirror either way, but only report an actual deletion
        // ('deleted'). 'changed' means the item merely left our view (moved
        // folder / rescheduled out of window) — reporting it as deleted lies.
        deletes.push(item.graphId);
        if (!baseline && item.removedReason === 'deleted' && seen.has(item.graphId)) {
          const known = this.store.delta.getItem(accountId, desc.storageKey, item.graphId);
          deleted.push({ token, graphId: item.graphId, summary: known?.summary ?? '', changeType: 'deleted' });
        }
        continue;
      }

      upserts.push({ graphId: item.graphId, token, summary: item.summary, snapshot: JSON.stringify(item.snapshot) });
      if (!baseline) {
        const changeType = seen.has(item.graphId) ? 'updated' : 'created';
        const entry: ChangeEntry = { token, graphId: item.graphId, summary: item.summary, changeType, detail: item.snapshot };
        (changeType === 'updated' ? updated : created).push(entry);
      }
    }

    this.store.delta.commit({
      accountId,
      resource: desc.storageKey,
      deltaLink,
      syncedAt,
      upserts,
      deletes,
      replaceMirror: baseline,
    });

    const result: ResourceChangeSet = {
      resource: desc.key,
      baseline,
      trackedCount: this.store.delta.countItems(accountId, desc.storageKey),
      created,
      updated,
      deleted,
    };
    if (deltaLink.length === 0) {
      result.note =
        'Graph returned no delta cursor (large initial page); tracking will re-baseline on the next call.';
    }
    return result;
  }
}
