/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Delta-sync mirror storage (U12). Persists, per account + resource:
 * - a Graph `@odata.deltaLink` cursor (`delta_links`), and
 * - a light snapshot of every item seen so far (`delta_items`).
 *
 * The item mirror is what lets a subsequent delta round classify a returned
 * entry as *added* vs *updated*: Graph delta returns both as a plain `value`
 * entry, distinguishable only by whether we have seen the id before. Deletes
 * arrive as `@removed` entries and drop the mirror row.
 *
 * This is the storage layer only; the sync orchestration and change
 * classification live in {@link ../delta/mirror.ts}. It is a thin companion to
 * {@link StateStore} (sharing its single better-sqlite3 connection), kept in its
 * own module so the mirror's SQL does not bloat the core store.
 */

import type Database from 'better-sqlite3';

type DB = Database.Database;

/** A mirrored item's stored snapshot. */
export interface MirrorItem {
  graphId: string;
  /** Durable self-encoding token (e.g. `em_`, `ev_`) for the item. */
  token: string;
  /** Short human label (subject / title); never null (defaults to ''). */
  summary: string;
  /** Optional JSON blob of light display fields. */
  snapshot?: string | null;
}

/** A committed sync round for one resource. */
export interface DeltaCommit {
  accountId: string;
  resource: string;
  /** New cursor to persist; when empty the cursor row is cleared (re-baseline). */
  deltaLink: string;
  syncedAt: number;
  /** Items to insert or refresh in the mirror. */
  upserts: readonly MirrorItem[];
  /** Graph ids to drop from the mirror. */
  deletes: readonly string[];
}

interface RawItemRow {
  graph_id: string;
  token: string;
  summary: string;
  snapshot: string | null;
}

export class DeltaStore {
  private readonly db: DB;
  private readonly now: () => number;

  constructor(db: DB, now: () => number) {
    this.db = db;
    this.now = now;
  }

  /** Returns the stored delta cursor for a resource, or null when un-synced. */
  getDeltaLink(accountId: string, resource: string): string | null {
    const row = this.db
      .prepare('SELECT delta_link FROM delta_links WHERE account_id = ? AND resource = ?')
      .get(accountId, resource) as { delta_link: string } | undefined;
    return row?.delta_link ?? null;
  }

  /** Returns the set of graph ids already mirrored for a resource. */
  getSeenIds(accountId: string, resource: string): Set<string> {
    const rows = this.db
      .prepare('SELECT graph_id FROM delta_items WHERE account_id = ? AND resource = ?')
      .all(accountId, resource) as Array<{ graph_id: string }>;
    return new Set(rows.map((r) => r.graph_id));
  }

  /** Looks up a single mirrored item (for labelling a delete), or null. */
  getItem(accountId: string, resource: string, graphId: string): MirrorItem | null {
    const raw = this.db
      .prepare(
        'SELECT graph_id, token, summary, snapshot FROM delta_items WHERE account_id = ? AND resource = ? AND graph_id = ?',
      )
      .get(accountId, resource, graphId) as RawItemRow | undefined;
    if (raw === undefined) return null;
    return { graphId: raw.graph_id, token: raw.token, summary: raw.summary, snapshot: raw.snapshot };
  }

  /** Number of items mirrored for a resource (baseline / status reporting). */
  countItems(accountId: string, resource: string): number {
    const row = this.db
      .prepare('SELECT COUNT(*) AS n FROM delta_items WHERE account_id = ? AND resource = ?')
      .get(accountId, resource) as { n: number };
    return row.n;
  }

  /**
   * Applies a completed sync round atomically: refresh the cursor and apply all
   * upserts/deletes in one transaction, so a crash mid-write never leaves the
   * mirror ahead of (or behind) its cursor. An empty `deltaLink` clears the
   * cursor row, forcing the next call to re-baseline (used when Graph returns no
   * usable deltaLink).
   */
  commit(commit: DeltaCommit): void {
    const apply = this.db.transaction((c: DeltaCommit): void => {
      if (c.deltaLink.length > 0) {
        this.db
          .prepare(
            `INSERT INTO delta_links (account_id, resource, delta_link, synced_at)
             VALUES (@accountId, @resource, @deltaLink, @syncedAt)
             ON CONFLICT(account_id, resource) DO UPDATE SET
               delta_link = excluded.delta_link,
               synced_at = excluded.synced_at`,
          )
          .run({ accountId: c.accountId, resource: c.resource, deltaLink: c.deltaLink, syncedAt: c.syncedAt });
      } else {
        this.db
          .prepare('DELETE FROM delta_links WHERE account_id = ? AND resource = ?')
          .run(c.accountId, c.resource);
      }

      const upsert = this.db.prepare(
        `INSERT INTO delta_items (account_id, resource, graph_id, token, summary, snapshot, first_seen, last_changed)
         VALUES (@accountId, @resource, @graphId, @token, @summary, @snapshot, @now, @now)
         ON CONFLICT(account_id, resource, graph_id) DO UPDATE SET
           token = excluded.token,
           summary = excluded.summary,
           snapshot = excluded.snapshot,
           last_changed = excluded.last_changed`,
      );
      for (const item of c.upserts) {
        upsert.run({
          accountId: c.accountId,
          resource: c.resource,
          graphId: item.graphId,
          token: item.token,
          summary: item.summary,
          snapshot: item.snapshot ?? null,
          now: c.syncedAt,
        });
      }

      const del = this.db.prepare(
        'DELETE FROM delta_items WHERE account_id = ? AND resource = ? AND graph_id = ?',
      );
      for (const graphId of c.deletes) {
        del.run(c.accountId, c.resource, graphId);
      }
    });
    apply(commit);
  }

  /**
   * Clears the cursor and mirror for a resource (or every resource for the
   * account when `resource` is omitted), forcing a fresh baseline on the next
   * sync. Local-only; never touches the user's Graph data.
   */
  reset(accountId: string, resource?: string): void {
    const wipe = this.db.transaction((): void => {
      if (resource != null) {
        this.db.prepare('DELETE FROM delta_links WHERE account_id = ? AND resource = ?').run(accountId, resource);
        this.db.prepare('DELETE FROM delta_items WHERE account_id = ? AND resource = ?').run(accountId, resource);
      } else {
        this.db.prepare('DELETE FROM delta_links WHERE account_id = ?').run(accountId);
        this.db.prepare('DELETE FROM delta_items WHERE account_id = ?').run(accountId);
      }
    });
    wipe();
  }
}
