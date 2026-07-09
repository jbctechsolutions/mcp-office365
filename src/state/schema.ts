/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * SQLite schema for the durable state store (U4). Each entry in {@link MIGRATIONS}
 * moves the schema forward by one version; the applied version is tracked in the
 * `meta` table so boot migration is idempotent and ordered.
 *
 * Tables (D3/D7/D8):
 * - `aliases`          — durable-ID token → Graph ID map (composite/mutable IDs).
 * - `approval_tokens`  — two-phase approval tokens + redemption receipts.
 * - `meta`             — schema version + one-shot migration markers.
 * - `delta_links`      — per-resource Graph delta cursors (U12 mirror).
 * - `delta_items`      — local mirror of seen items, to classify add/update/delete.
 */

/** Ordered forward migrations. Index i migrates schema version i → i+1. */
export const MIGRATIONS: readonly string[] = [
  // v0 → v1: initial durable-state schema.
  `
  CREATE TABLE IF NOT EXISTS aliases (
    token       TEXT PRIMARY KEY,
    graph_id    TEXT NOT NULL,
    entity_type TEXT NOT NULL,
    account_id  TEXT NOT NULL,
    mutable     INTEGER NOT NULL DEFAULT 0,
    created_at  INTEGER NOT NULL
  );

  CREATE TABLE IF NOT EXISTS approval_tokens (
    token         TEXT PRIMARY KEY,
    operation_key TEXT UNIQUE,
    action        TEXT NOT NULL,
    target_json   TEXT NOT NULL,
    content_hash  TEXT,
    account_id    TEXT NOT NULL,
    expires_at    INTEGER NOT NULL,
    redeemed_at   INTEGER,
    receipt_json  TEXT,
    created_at    INTEGER NOT NULL
  );

  CREATE INDEX IF NOT EXISTS idx_approval_tokens_expires_at ON approval_tokens (expires_at);
  `,
  // v1 → v2: delta-sync mirror (U12). One cursor row per (account, resource),
  // plus a per-item mirror so a subsequent delta round can tell an added item
  // from an updated one (Graph delta reports both as a plain value entry).
  `
  CREATE TABLE IF NOT EXISTS delta_links (
    account_id  TEXT NOT NULL,
    resource    TEXT NOT NULL,
    delta_link  TEXT NOT NULL,
    synced_at   INTEGER NOT NULL,
    PRIMARY KEY (account_id, resource)
  );

  CREATE TABLE IF NOT EXISTS delta_items (
    account_id   TEXT NOT NULL,
    resource     TEXT NOT NULL,
    graph_id     TEXT NOT NULL,
    token        TEXT NOT NULL,
    summary      TEXT NOT NULL DEFAULT '',
    snapshot     TEXT,
    first_seen   INTEGER NOT NULL,
    last_changed INTEGER NOT NULL,
    PRIMARY KEY (account_id, resource, graph_id)
  );
  `,
];

/** The schema version this build expects (equals the migration count). */
export const SCHEMA_VERSION = MIGRATIONS.length;

/** meta keys. */
export const META_SCHEMA_VERSION = 'schema_version';
export const META_LEGACY_TOKENS_MIGRATED = 'legacy_tokens_migrated';

/** Retention window for redeemed/expired approval tokens (D8): 90 days. */
export const APPROVAL_RETENTION_MS = 90 * 24 * 60 * 60 * 1000;
