/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { DatabaseSync } from 'node:sqlite';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pragmaGet, pragmaSet, transaction, immediateTransaction } from '../../../src/state/sqlite-compat.js';

describe('state/sqlite-compat', () => {
  let dir: string;
  let db: DatabaseSync;

  beforeEach(() => {
    // A file db, not ':memory:' — WAL is a no-op on memory databases, so the
    // journal_mode assertions below would pass vacuously.
    dir = mkdtempSync(join(tmpdir(), 'sqlite-compat-'));
    db = new DatabaseSync(join(dir, 'test.db'));
    db.exec('CREATE TABLE t (id INTEGER PRIMARY KEY, name TEXT)');
  });

  afterEach(() => {
    try {
      db.close();
    } catch {
      /* already closed by a test */
    }
    rmSync(dir, { recursive: true, force: true });
  });

  describe('pragmaGet', () => {
    it('reads journal_mode as a string', () => {
      pragmaSet(db, 'journal_mode = WAL');
      expect(pragmaGet(db, 'journal_mode')).toBe('wal');
    });

    it('reads busy_timeout by position, not by pragma name', () => {
      // Regression guard: SQLite returns `{ timeout: 5000 }` for this pragma —
      // keyed by result column, not pragma name. A shim indexing by name
      // returns undefined here with no error.
      pragmaSet(db, 'busy_timeout = 5000');
      expect(pragmaGet(db, 'busy_timeout')).toBe(5000);
    });

    it('returns undefined for a pragma that yields no row', () => {
      expect(pragmaGet(db, 'optimize')).toBeUndefined();
    });
  });

  describe('pragmaSet', () => {
    it('applies an assignment that a subsequent read reflects', () => {
      pragmaSet(db, 'busy_timeout = 1234');
      expect(pragmaGet(db, 'busy_timeout')).toBe(1234);
    });

    it('applies journal_mode, which cannot run as a prepared statement', () => {
      pragmaSet(db, 'journal_mode = WAL');
      expect(pragmaGet(db, 'journal_mode')).toBe('wal');
    });
  });

  describe('transaction', () => {
    it('returns the callback value unchanged', () => {
      expect(transaction(db, () => 'ok' as const)).toBe('ok');
      expect(transaction(db, () => 42)).toBe(42);
    });

    it('commits writes made inside', () => {
      transaction(db, () => {
        db.prepare('INSERT INTO t (name) VALUES (?)').run('committed');
      });
      expect(db.prepare('SELECT count(*) AS c FROM t').get()).toEqual({ c: 1 });
    });

    it('rolls back on throw and propagates the original error', () => {
      const boom = new Error('boom');
      expect(() =>
        transaction(db, () => {
          db.prepare('INSERT INTO t (name) VALUES (?)').run('rolled-back');
          throw boom;
        }),
      ).toThrow(boom);
      expect(db.prepare('SELECT count(*) AS c FROM t').get()).toEqual({ c: 0 });
    });

    it('supports consecutive transactions on the same handle', () => {
      // Consecutive, not nested: BEGIN cannot run inside an active transaction,
      // and no call site nests. Nesting would require SAVEPOINT support.
      transaction(db, () => {
        db.prepare('INSERT INTO t (name) VALUES (?)').run('first');
      });
      transaction(db, () => {
        db.prepare('INSERT INTO t (name) VALUES (?)').run('second');
      });
      expect(db.prepare('SELECT count(*) AS c FROM t').get()).toEqual({ c: 2 });
    });

    it('leaves no transaction open when COMMIT itself fails', () => {
      // A DEFERRABLE INITIALLY DEFERRED constraint is only checked at COMMIT,
      // so this makes COMMIT — not the callback — throw. If COMMIT ran outside
      // the protected block, the transaction would stay open and the *next*
      // BEGIN would fail, surfacing the problem two operations downstream.
      const fk = new DatabaseSync(join(dir, 'fk.db'), { enableForeignKeyConstraints: true });
      try {
        fk.exec('CREATE TABLE parent (id INTEGER PRIMARY KEY)');
        fk.exec(
          'CREATE TABLE child (id INTEGER PRIMARY KEY, parent_id INTEGER REFERENCES parent(id) DEFERRABLE INITIALLY DEFERRED)',
        );

        expect(() =>
          transaction(fk, () => {
            fk.prepare('INSERT INTO child (parent_id) VALUES (?)').run(999);
          }),
        ).toThrow(/FOREIGN KEY/i);

        // The proof: a subsequent transaction still works.
        transaction(fk, () => {
          fk.prepare('INSERT INTO parent (id) VALUES (?)').run(1);
        });
        expect(fk.prepare('SELECT count(*) AS c FROM parent').get()).toEqual({ c: 1 });
        expect(fk.prepare('SELECT count(*) AS c FROM child').get()).toEqual({ c: 0 });
      } finally {
        fk.close();
      }
    });

    it('leaves no transaction open after a rollback, so the next one succeeds', () => {
      expect(() =>
        transaction(db, () => {
          throw new Error('first fails');
        }),
      ).toThrow('first fails');
      // Would throw "cannot start a transaction within a transaction" if the
      // failed transaction had not been rolled back.
      transaction(db, () => {
        db.prepare('INSERT INTO t (name) VALUES (?)').run('after-rollback');
      });
      expect(db.prepare('SELECT count(*) AS c FROM t').get()).toEqual({ c: 1 });
    });
  });

  describe('immediateTransaction', () => {
    it('commits and returns the callback value', () => {
      const result = immediateTransaction(db, () => {
        db.prepare('INSERT INTO t (name) VALUES (?)').run('immediate');
        return 'done' as const;
      });
      expect(result).toBe('done');
      expect(db.prepare('SELECT count(*) AS c FROM t').get()).toEqual({ c: 1 });
    });

    it('rolls back on throw', () => {
      expect(() =>
        immediateTransaction(db, () => {
          db.prepare('INSERT INTO t (name) VALUES (?)').run('nope');
          throw new Error('fail');
        }),
      ).toThrow('fail');
      expect(db.prepare('SELECT count(*) AS c FROM t').get()).toEqual({ c: 0 });
    });

    it('takes the write lock up front, unlike a deferred transaction', () => {
      // Two connections to the same file. With BEGIN IMMEDIATE held by the
      // first, a second writer is locked out at BEGIN IMMEDIATE rather than
      // being allowed to proceed and collide at write time.
      pragmaSet(db, 'journal_mode = WAL');
      const other = new DatabaseSync(join(dir, 'test.db'));
      try {
        db.exec('BEGIN IMMEDIATE');
        expect(() => other.exec('BEGIN IMMEDIATE')).toThrow(/lock/i);
        db.exec('ROLLBACK');
      } finally {
        other.close();
      }
    });
  });
});
