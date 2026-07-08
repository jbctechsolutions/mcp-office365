/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Durable state store (U4) — a better-sqlite3-backed store at
 * `~/.mcp-office365/state.db` holding durable-ID aliases and two-phase approval
 * tokens. Provides the concurrency (WAL + busy_timeout + atomic consume, D7),
 * at-rest (0700/0600 permissions, D18), account-stamping (D7), and
 * degradation (in-memory fallback, D15) semantics the rest of v3 builds on.
 *
 * This unit is the storage layer only; wiring into the ID resolver (U5) and the
 * approval manager (U9) lands in those units.
 */

import Database from 'better-sqlite3';
import { chmodSync, existsSync, mkdirSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { runMigrations, migrateLegacyTokens } from './migrate.js';
import { APPROVAL_RETENTION_MS } from './schema.js';

type DB = Database.Database;

const DEFAULT_DIR = join(homedir(), '.mcp-office365');
const LEGACY_DIR = join(homedir(), '.outlook-mcp');
const DB_FILENAME = 'state.db';

/** A durable-ID alias row. */
export interface AliasRow {
  token: string;
  graphId: string;
  entityType: string;
  accountId: string;
  mutable: boolean;
  createdAt: number;
}

/** Input for minting an alias. */
export interface AliasInput {
  token: string;
  graphId: string;
  entityType: string;
  accountId: string;
  mutable?: boolean;
  createdAt?: number;
}

/** Input for staging a two-phase approval token. */
export interface ApprovalTokenInput {
  token: string;
  operationKey?: string | null;
  action: string;
  targetJson: string;
  contentHash?: string | null;
  accountId: string;
  expiresAt: number;
  createdAt?: number;
}

/** Outcome of an atomic approval-token consume. */
export type ConsumeResult =
  | { status: 'consumed'; receiptJson: string | null }
  | { status: 'already_redeemed'; receiptJson: string | null }
  | { status: 'foreign_account' }
  | { status: 'not_found' };

/** Options for {@link StateStore.open}. */
export interface StateStoreOptions {
  /** Override the state directory (defaults to `~/.mcp-office365`). */
  dir?: string;
  /** Override the legacy directory to migrate tokens from. */
  legacyDir?: string;
  /** Clock, for deterministic tests. */
  now?: () => number;
  /** Warning sink (defaults to stderr). */
  warn?: (message: string) => void;
}

interface RawAliasRow {
  token: string;
  graph_id: string;
  entity_type: string;
  account_id: string;
  mutable: number;
  created_at: number;
}

export class StateStore {
  /** True when running from an in-memory fallback (durability degraded). */
  readonly degraded: boolean;
  /** The db path, or ':memory:' when degraded. */
  readonly path: string;
  /** Effective `journal_mode` pragma (e.g. 'wal'). */
  readonly journalMode: string;
  /** Effective `busy_timeout` pragma in ms. */
  readonly busyTimeout: number;

  private readonly db: DB;
  private readonly now: () => number;

  private constructor(db: DB, path: string, degraded: boolean, now: () => number) {
    this.db = db;
    this.path = path;
    this.degraded = degraded;
    this.now = now;
    this.journalMode = String(db.pragma('journal_mode', { simple: true }));
    this.busyTimeout = Number(db.pragma('busy_timeout', { simple: true }));
  }

  /**
   * Opens (creating if needed) the durable store. On any failure to use the
   * on-disk db (corrupt/locked file, unwritable dir) it degrades to an in-memory
   * store, emits a stderr warning, and reports `degraded = true` — the server
   * stays usable, only durability is lost for the run.
   */
  static open(options: StateStoreOptions = {}): StateStore {
    const dir = options.dir ?? DEFAULT_DIR;
    const legacyDir = options.legacyDir ?? LEGACY_DIR;
    const now = options.now ?? ((): number => Date.now());
    const warn = options.warn ?? ((msg: string): void => void process.stderr.write(`${msg}\n`));

    try {
      // Directory: create 0700, or repair perms if it already exists (D18).
      if (!existsSync(dir)) {
        mkdirSync(dir, { recursive: true, mode: 0o700 });
      } else {
        safeChmod(dir, 0o700);
      }

      const dbPath = join(dir, DB_FILENAME);
      const db = new Database(dbPath);
      configurePragmas(db);
      runMigrations(db); // executes DDL — surfaces a corrupt file here
      migrateLegacyTokens(db, dir, legacyDir);
      // File may have just been created; enforce 0600 (and repair a loose one).
      safeChmod(dbPath, 0o600);

      const store = new StateStore(db, dbPath, false, now);
      store.purgeExpired(now());
      return store;
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      warn(`[mcp-office365] state store unavailable (${reason}); running in-memory (durability degraded).`);
      const mem = new Database(':memory:');
      configurePragmas(mem);
      runMigrations(mem);
      return new StateStore(mem, ':memory:', true, now);
    }
  }

  // ---- Aliases (D3) --------------------------------------------------------

  /** Inserts or replaces a durable-ID alias. */
  putAlias(input: AliasInput): void {
    this.db
      .prepare(
        `INSERT INTO aliases (token, graph_id, entity_type, account_id, mutable, created_at)
         VALUES (@token, @graphId, @entityType, @accountId, @mutable, @createdAt)
         ON CONFLICT(token) DO UPDATE SET
           graph_id = excluded.graph_id,
           entity_type = excluded.entity_type,
           account_id = excluded.account_id,
           mutable = excluded.mutable`,
      )
      .run({
        token: input.token,
        graphId: input.graphId,
        entityType: input.entityType,
        accountId: input.accountId,
        mutable: input.mutable === true ? 1 : 0,
        createdAt: input.createdAt ?? this.now(),
      });
  }

  /**
   * Resolves an alias token. When `accountId` is provided, rows stamped with a
   * different account are invisible (D7 account stamping).
   */
  getAlias(token: string, accountId?: string): AliasRow | null {
    const raw = (accountId != null
      ? this.db.prepare('SELECT * FROM aliases WHERE token = ? AND account_id = ?').get(token, accountId)
      : this.db.prepare('SELECT * FROM aliases WHERE token = ?').get(token)) as RawAliasRow | undefined;
    if (raw === undefined) {
      return null;
    }
    return {
      token: raw.token,
      graphId: raw.graph_id,
      entityType: raw.entity_type,
      accountId: raw.account_id,
      mutable: raw.mutable !== 0,
      createdAt: raw.created_at,
    };
  }

  // ---- Approval tokens (D7/D8) --------------------------------------------

  /** Stages a two-phase approval token. */
  putApprovalToken(input: ApprovalTokenInput): void {
    this.db
      .prepare(
        `INSERT INTO approval_tokens
           (token, operation_key, action, target_json, content_hash, account_id, expires_at, created_at)
         VALUES (@token, @operationKey, @action, @targetJson, @contentHash, @accountId, @expiresAt, @createdAt)`,
      )
      .run({
        token: input.token,
        operationKey: input.operationKey ?? null,
        action: input.action,
        targetJson: input.targetJson,
        contentHash: input.contentHash ?? null,
        accountId: input.accountId,
        expiresAt: input.expiresAt,
        createdAt: input.createdAt ?? this.now(),
      });
  }

  /**
   * Atomically consumes an approval token (D8). The guarded `UPDATE … WHERE
   * redeemed_at IS NULL RETURNING` means only one caller — across processes
   * sharing the db — can win the redemption; the loser sees `already_redeemed`
   * and can return the stored receipt. A foreign-account token is refused.
   */
  consumeApprovalToken(args: {
    token: string;
    accountId: string;
    receiptJson?: string | null;
    now?: number;
  }): ConsumeResult {
    const redeemedAt = args.now ?? this.now();
    const updated = this.db
      .prepare(
        `UPDATE approval_tokens
           SET redeemed_at = ?, receipt_json = ?
         WHERE token = ? AND account_id = ? AND redeemed_at IS NULL
         RETURNING receipt_json`,
      )
      .get(redeemedAt, args.receiptJson ?? null, args.token, args.accountId) as
      | { receipt_json: string | null }
      | undefined;

    if (updated !== undefined) {
      return { status: 'consumed', receiptJson: updated.receipt_json };
    }

    const row = this.db
      .prepare('SELECT account_id, redeemed_at, receipt_json FROM approval_tokens WHERE token = ?')
      .get(args.token) as
      | { account_id: string; redeemed_at: number | null; receipt_json: string | null }
      | undefined;
    if (row === undefined) {
      return { status: 'not_found' };
    }
    if (row.account_id !== args.accountId) {
      return { status: 'foreign_account' };
    }
    return { status: 'already_redeemed', receiptJson: row.receipt_json };
  }

  // ---- Maintenance ---------------------------------------------------------

  /**
   * Deletes approval tokens whose expiry is older than the retention window
   * (D8, 90 days). Recently-expired receipts linger so idempotent redemption
   * still returns them. Returns the number of rows purged.
   */
  purgeExpired(now: number): number {
    const cutoff = now - APPROVAL_RETENTION_MS;
    const result = this.db.prepare('DELETE FROM approval_tokens WHERE expires_at < ?').run(cutoff);
    return result.changes;
  }

  close(): void {
    this.db.close();
  }
}

function configurePragmas(db: DB): void {
  db.pragma('journal_mode = WAL');
  db.pragma('busy_timeout = 5000');
}

/** chmod that never throws (a no-op on platforms without POSIX perms). */
function safeChmod(target: string, mode: number): void {
  try {
    chmodSync(target, mode);
  } catch {
    // Best-effort — Windows and some filesystems don't honor POSIX modes.
  }
}
