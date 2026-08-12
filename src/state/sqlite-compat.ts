/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * The two conveniences `node:sqlite` does not provide that better-sqlite3 did:
 * pragma access and transaction wrapping (#108).
 *
 * Both are deliberately thin. They exist so the state modules read the same as
 * they did under better-sqlite3, not to build a driver abstraction — there is
 * one driver and it is the standard library.
 */

import type { DatabaseSync } from 'node:sqlite';

/**
 * Reads a pragma and returns its scalar value.
 *
 * Take the row's first value rather than indexing by pragma name: SQLite names
 * the returned column after the *result*, not the pragma, so
 * `PRAGMA busy_timeout` comes back as `{ timeout: 5000 }`. Indexing by pragma
 * name works for `journal_mode` and silently yields `undefined` for
 * `busy_timeout` — a wrong value with no error, which is the worst shape a bug
 * in a storage layer can take.
 */
export function pragmaGet(db: DatabaseSync, name: string): unknown {
  const row = db.prepare(`PRAGMA ${name}`).get();
  if (row == null) return undefined;
  return Object.values(row)[0];
}

/**
 * Applies a pragma assignment, e.g. `busy_timeout = 5000`.
 *
 * Routed through `exec` because assignment pragmas return no rows and some
 * (notably `journal_mode`) must run outside a prepared-statement context.
 */
export function pragmaSet(db: DatabaseSync, assignment: string): void {
  db.exec(`PRAGMA ${assignment}`);
}

/**
 * Runs `fn` inside a transaction, committing on return and rolling back on
 * throw. Returns whatever `fn` returns.
 *
 * Unlike better-sqlite3's `db.transaction(fn)`, this executes immediately
 * rather than returning a callable. Call sites that previously passed arguments
 * to the returned function close over them instead — no variadic form is
 * needed, and the transaction boundary is visible at the call site rather than
 * deferred to a later invocation.
 *
 * A failed ROLLBACK never masks the original error: the reason the transaction
 * failed is more useful than the reason the cleanup did.
 */
export function transaction<T>(db: DatabaseSync, fn: () => T): T {
  return runInTransaction(db, 'BEGIN', fn);
}

/**
 * As {@link transaction}, but opens with `BEGIN IMMEDIATE` — the write lock is
 * taken up front instead of on first write.
 *
 * This matters for read-then-write sequences under WAL with more than one
 * process: a deferred transaction lets two readers both observe "no conflict"
 * and then both write. Preserves the `.immediate()` semantics the alias
 * registration relied on under better-sqlite3.
 */
export function immediateTransaction<T>(db: DatabaseSync, fn: () => T): T {
  return runInTransaction(db, 'BEGIN IMMEDIATE', fn);
}

function runInTransaction<T>(db: DatabaseSync, begin: string, fn: () => T): T {
  db.exec(begin);
  let result: T;
  try {
    result = fn();
  } catch (error) {
    try {
      db.exec('ROLLBACK');
    } catch {
      /* Preserve the original failure; the rollback error is noise beside it. */
    }
    throw error;
  }
  db.exec('COMMIT');
  return result;
}
