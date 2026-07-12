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
import { DeltaStore } from './delta-store.js';

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

/** A stored approval-token row (read view). */
export interface ApprovalTokenRow {
  token: string;
  action: string;
  targetJson: string;
  contentHash: string | null;
  expiresAt: number;
  redeemedAt: number | null;
  receiptJson: string | null;
  createdAt: number;
}

/** Outcome of an atomic approval-token consume. */
export type ConsumeResult =
  | { status: 'consumed'; receiptJson: string | null }
  | { status: 'already_redeemed'; receiptJson: string | null }
  | { status: 'expired' }
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
  /** Delta-sync mirror storage (U12), sharing this store's connection. */
  readonly delta: DeltaStore;

  private readonly db: DB;
  private readonly now: () => number;

  private constructor(db: DB, path: string, degraded: boolean, now: () => number) {
    this.db = db;
    this.path = path;
    this.degraded = degraded;
    this.now = now;
    this.journalMode = String(db.pragma('journal_mode', { simple: true }));
    this.busyTimeout = Number(db.pragma('busy_timeout', { simple: true }));
    this.delta = new DeltaStore(db, now);
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

    let fileDb: DB | undefined;
    try {
      // Directory: create 0700, or repair perms if it already exists (D18).
      if (!existsSync(dir)) {
        mkdirSync(dir, { recursive: true, mode: 0o700 });
      } else {
        safeChmod(dir, 0o700, warn);
      }

      const dbPath = join(dir, DB_FILENAME);
      fileDb = new Database(dbPath);
      configurePragmas(fileDb);
      runMigrations(fileDb); // executes DDL — surfaces a corrupt/newer file here
      // The file (and its -wal/-shm sidecars) may have just been created;
      // enforce 0600 across all three (and repair a loose one).
      safeChmod(dbPath, 0o600, warn);
      safeChmod(`${dbPath}-wal`, 0o600, warn);
      safeChmod(`${dbPath}-shm`, 0o600, warn);

      const store = new StateStore(fileDb, dbPath, false, now);

      // The legacy token copy and boot purge are conveniences — a failure in
      // either must NOT discard an otherwise-healthy on-disk store, so they run
      // outside the degrade-governing path (warn-and-continue).
      try {
        migrateLegacyTokens(fileDb, dir, legacyDir);
      } catch (e) {
        warn(`[mcp-office365] legacy token migration skipped (${e instanceof Error ? e.message : String(e)}).`);
      }
      try {
        store.purgeExpired(now());
      } catch (e) {
        warn(`[mcp-office365] boot purge skipped (${e instanceof Error ? e.message : String(e)}).`);
      }
      return store;
    } catch (error) {
      // Release the on-disk handle (and its WAL lock) before degrading, so a
      // failed open does not leak a file descriptor / lock for the process.
      if (fileDb !== undefined) {
        try {
          fileDb.close();
        } catch {
          /* best-effort */
        }
      }
      const reason = error instanceof Error ? error.message : String(error);
      // The in-memory fallback is backed by the same native module, so a module
      // that cannot load (Node ABI mismatch, missing build) would just rethrow
      // the raw dlopen stack from the fallback. Surface remediation instead.
      if (isNativeLoadFailure(error)) {
        throw new Error(
          `better-sqlite3 native module failed to load — ABI mismatch or missing compiled ` +
            `binding (running Node.js ${process.version}), so neither the on-disk state ` +
            `store nor its in-memory fallback can start.\n` +
            `Fix one of:\n` +
            `  - npm rebuild better-sqlite3   # in the directory the server is installed in\n` +
            `  - rm -rf ~/.npm/_npx           # clear the npx cache so it recompiles on next run\n` +
            `  - run the server under the Node.js version that installed it\n` +
            `Original error: ${reason}`,
          { cause: error },
        );
      }
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
   * Resolves an alias token, scoped to `accountId` (D7 account stamping). The
   * account is a mandatory, fail-closed argument: a row minted under a different
   * account is invisible, so no caller can accidentally resolve a foreign
   * account's durable ID into a live Graph ID by omitting the scope.
   */
  getAlias(token: string, accountId: string): AliasRow | null {
    const raw = this.db
      .prepare('SELECT * FROM aliases WHERE token = ? AND account_id = ?')
      .get(token, accountId) as RawAliasRow | undefined;
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

  /**
   * Atomically registers an alias with D1a collision enforcement, in a single
   * IMMEDIATE transaction so the check-and-write is not a read-then-write race:
   * two concurrent immutable registrations of the same token cannot both pass
   * the check and then clobber each other. Returns `'collision'` when a
   * *different* Graph ID already occupies the token for an immutable entity
   * (neither the stored nor the new row is mutable); otherwise upserts and
   * returns `'ok'`.
   */
  registerAlias(input: AliasInput): 'ok' | 'collision' {
    const run = this.db.transaction((): 'ok' | 'collision' => {
      const existing = this.getAliasUnscoped(input.token);
      if (
        existing !== null &&
        existing.graphId !== input.graphId &&
        !existing.mutable &&
        input.mutable !== true
      ) {
        return 'collision';
      }
      this.putAlias(input);
      return 'ok';
    });
    return run.immediate();
  }

  /**
   * Returns an alias row regardless of account, or null. Reserved for two
   * internal uses that must see across accounts: foreign-account disambiguation
   * during resolution, and collision detection at mint time (D1a). It is NOT a
   * resolution path — {@link getAlias} (account-scoped) is the only one of those.
   */
  getAliasUnscoped(token: string): AliasRow | null {
    const raw = this.db.prepare('SELECT * FROM aliases WHERE token = ?').get(token) as
      | RawAliasRow
      | undefined;
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

  /**
   * Returns the account a token was minted under (any account), or null when the
   * token is unknown. Distinguishes "foreign account" from "unknown" during
   * resolution so a foreign token yields a typed ID_FOREIGN_ACCOUNT.
   */
  getAliasAccount(token: string): string | null {
    return this.getAliasUnscoped(token)?.accountId ?? null;
  }

  // ---- Approval tokens (D7/D8) --------------------------------------------

  /**
   * Reads an approval token scoped to `accountId` (or null). A read view for
   * validation/preview; redemption still goes through {@link consumeApprovalToken}
   * so the atomic guard is never bypassed.
   */
  getApprovalToken(token: string, accountId: string): ApprovalTokenRow | null {
    const raw = this.db
      .prepare(
        'SELECT token, action, target_json, content_hash, expires_at, redeemed_at, receipt_json, created_at FROM approval_tokens WHERE token = ? AND account_id = ?',
      )
      .get(token, accountId) as
      | {
          token: string;
          action: string;
          target_json: string;
          content_hash: string | null;
          expires_at: number;
          redeemed_at: number | null;
          receipt_json: string | null;
          created_at: number;
        }
      | undefined;
    if (raw === undefined) {
      return null;
    }
    return {
      token: raw.token,
      action: raw.action,
      targetJson: raw.target_json,
      contentHash: raw.content_hash,
      expiresAt: raw.expires_at,
      redeemedAt: raw.redeemed_at,
      receiptJson: raw.receipt_json,
      createdAt: raw.created_at,
    };
  }

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
   * redeemed_at IS NULL AND expires_at > now RETURNING` means only one caller —
   * across processes sharing the db — can win the redemption; the loser sees
   * `already_redeemed` and can return the stored receipt. Expiry is enforced
   * here at the trust boundary (not left to callers): an expired token that
   * still authorizes a send/delete would defeat the whole point of a time-boxed
   * two-phase approval. Foreign-account tokens are refused.
   */
  consumeApprovalToken(args: {
    token: string;
    accountId: string;
    receiptJson?: string | null;
    now?: number;
  }): ConsumeResult {
    const now = args.now ?? this.now();
    const updated = this.db
      .prepare(
        `UPDATE approval_tokens
           SET redeemed_at = ?, receipt_json = ?
         WHERE token = ? AND account_id = ? AND redeemed_at IS NULL AND expires_at > ?
         RETURNING receipt_json`,
      )
      .get(now, args.receiptJson ?? null, args.token, args.accountId, now) as
      | { receipt_json: string | null }
      | undefined;

    if (updated !== undefined) {
      return { status: 'consumed', receiptJson: updated.receipt_json };
    }

    // Classify why the guarded update matched nothing.
    const row = this.db
      .prepare(
        'SELECT account_id, redeemed_at, receipt_json, expires_at FROM approval_tokens WHERE token = ?',
      )
      .get(args.token) as
      | { account_id: string; redeemed_at: number | null; receipt_json: string | null; expires_at: number }
      | undefined;
    if (row === undefined) {
      return { status: 'not_found' };
    }
    if (row.account_id !== args.accountId) {
      return { status: 'foreign_account' };
    }
    // Redemption takes precedence over expiry so an idempotent re-consume of an
    // already-redeemed token still returns its receipt.
    if (row.redeemed_at !== null) {
      return { status: 'already_redeemed', receiptJson: row.receipt_json };
    }
    return { status: 'expired' };
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

/**
 * True when the error means the better-sqlite3 native binding itself cannot be
 * loaded (as opposed to a bad/locked db file): dlopen ABI rejection, or the
 * compiled `.node` artifact missing entirely.
 */
function isNativeLoadFailure(error: unknown): boolean {
  if (!(error instanceof Error)) return false;
  const code = (error as NodeJS.ErrnoException).code;
  if (code === 'ERR_DLOPEN_FAILED') return true;
  if (/NODE_MODULE_VERSION|was compiled against a different Node\.js version/.test(error.message)) {
    return true;
  }
  // The `bindings` package throws a code-less plain Error when the compiled
  // artifact is missing entirely (never built / pruned).
  if (/Could not locate the bindings file/.test(error.message)) return true;
  return code === 'MODULE_NOT_FOUND' && error.message.includes('better_sqlite3.node');
}

function configurePragmas(db: DB): void {
  db.pragma('journal_mode = WAL');
  db.pragma('busy_timeout = 5000');
}

/**
 * chmod that never throws. Windows/filesystems without POSIX modes fail benignly
 * and are ignored; a genuine failure on a POSIX host (e.g. EPERM) means the file
 * may sit with looser-than-intended permissions, so we surface it via `warn`
 * rather than swallow silently — the at-rest exposure should be observable.
 * A missing sidecar (-wal/-shm not yet created) is not a failure worth noting.
 */
function safeChmod(target: string, mode: number, warn?: (message: string) => void): void {
  try {
    chmodSync(target, mode);
  } catch (error) {
    const code = (error as { code?: string }).code;
    if (process.platform !== 'win32' && code !== 'ENOENT' && warn != null) {
      warn(`[mcp-office365] could not set permissions on ${target} (${code ?? 'error'}); it may be readable by other local users.`);
    }
  }
}
