/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Boot-time migrations for the durable state store (U4):
 * - Schema migrations applied in order, tracked via the `meta` table (idempotent).
 * - One-shot legacy token copy from `~/.outlook-mcp` (D15), never overwriting an
 *   existing token cache, gated by a `meta` marker so it runs at most once.
 */

import { chmodSync, copyFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import type Database from 'better-sqlite3';
import {
  MIGRATIONS,
  META_SCHEMA_VERSION,
  META_LEGACY_TOKENS_MIGRATED,
} from './schema.js';

type DB = Database.Database;

/** Creates the meta table if absent (needed before reading the schema version). */
function ensureMetaTable(db: DB): void {
  db.exec('CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);');
}

function getMeta(db: DB, key: string): string | null {
  const row = db.prepare('SELECT value FROM meta WHERE key = ?').get(key) as
    | { value: string }
    | undefined;
  return row?.value ?? null;
}

function setMeta(db: DB, key: string, value: string): void {
  db.prepare(
    'INSERT INTO meta (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value',
  ).run(key, value);
}

/** Reads the applied schema version (0 when never migrated). */
export function getSchemaVersion(db: DB): number {
  const raw = getMeta(db, META_SCHEMA_VERSION);
  const parsed = raw != null ? Number.parseInt(raw, 10) : 0;
  return Number.isFinite(parsed) ? parsed : 0;
}

/**
 * Applies pending schema migrations in order, each in its own transaction so a
 * failure leaves the schema at the last good version. Running with no pending
 * migrations is a no-op.
 */
export function runMigrations(db: DB): void {
  ensureMetaTable(db);
  let current = getSchemaVersion(db);
  // Downgrade guard: a db migrated by a newer build (version beyond this build's
  // migration count) must not be operated on blindly against an unknown schema.
  // Throwing here routes the caller to its degraded in-memory fallback.
  if (current > MIGRATIONS.length) {
    throw new Error(
      `state.db schema version ${current} is newer than this build supports (${MIGRATIONS.length}).`,
    );
  }
  while (current < MIGRATIONS.length) {
    const sql = MIGRATIONS[current];
    if (sql == null) break;
    const apply = db.transaction((migrationSql: string, next: number) => {
      db.exec(migrationSql);
      setMeta(db, META_SCHEMA_VERSION, String(next));
    });
    apply(sql, current + 1);
    current += 1;
  }
}

/**
 * Copies a legacy `~/.outlook-mcp/tokens.json` into the new state dir exactly
 * once (D15): only when the marker is unset AND the new dir has no tokens.json.
 * The legacy dir is left untouched; the marker is set regardless so the check
 * never repeats.
 */
export function migrateLegacyTokens(
  db: DB,
  newDir: string,
  legacyDir: string,
): void {
  if (getMeta(db, META_LEGACY_TOKENS_MIGRATED) === '1') {
    return;
  }
  const newTokens = join(newDir, 'tokens.json');
  const legacyTokens = join(legacyDir, 'tokens.json');
  if (!existsSync(newTokens) && existsSync(legacyTokens)) {
    copyFileSync(legacyTokens, newTokens);
    // The copied cache holds OAuth material — restrict it like state.db (D18).
    try {
      chmodSync(newTokens, 0o600);
    } catch {
      // Best-effort; non-POSIX filesystems ignore modes.
    }
  }
  setMeta(db, META_LEGACY_TOKENS_MIGRATED, '1');
}
