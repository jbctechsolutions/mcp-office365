/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { mkdtempSync, rmSync, writeFileSync, readFileSync, existsSync, statSync, chmodSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import Database from 'better-sqlite3';
import { StateStore } from '../../../src/state/store.js';
import { SCHEMA_VERSION, APPROVAL_RETENTION_MS } from '../../../src/state/schema.js';

const isPosix = process.platform !== 'win32';

let dir: string;
let legacyDir: string;
const open = new Set<StateStore>();

function openStore(now?: () => number): StateStore {
  const store = StateStore.open({ dir, legacyDir, now, warn: () => {} });
  open.add(store);
  return store;
}

function approval(overrides: Partial<Parameters<StateStore['putApprovalToken']>[0]> = {}): Parameters<StateStore['putApprovalToken']>[0] {
  return {
    token: 'ap_1',
    action: 'send_email',
    targetJson: '{"to":"a@b.com"}',
    accountId: 'acct-X',
    // Far future so a boot purge (which runs at real Date.now() for stores
    // opened without a fixed clock) never removes it mid-test.
    expiresAt: 4_000_000_000_000,
    ...overrides,
  };
}

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'mcp-state-'));
  legacyDir = mkdtempSync(join(tmpdir(), 'mcp-legacy-'));
});

afterEach(() => {
  for (const s of open) {
    try {
      s.close();
    } catch {
      /* already closed */
    }
  }
  open.clear();
  rmSync(dir, { recursive: true, force: true });
  rmSync(legacyDir, { recursive: true, force: true });
});

describe('StateStore.open', () => {
  it('sets WAL journal mode and a busy_timeout on open', () => {
    const store = openStore();
    expect(store.degraded).toBe(false);
    expect(store.journalMode).toBe('wal');
    expect(store.busyTimeout).toBe(5000);
  });

  it('creates the db with the current schema version and expected tables', () => {
    openStore();
    const raw = new Database(join(dir, 'state.db'));
    const version = raw.prepare("SELECT value FROM meta WHERE key = 'schema_version'").get() as { value: string };
    expect(Number(version.value)).toBe(SCHEMA_VERSION);
    const tables = raw
      .prepare("SELECT name FROM sqlite_master WHERE type = 'table'")
      .all()
      .map((r) => (r as { name: string }).name);
    expect(tables).toEqual(expect.arrayContaining(['aliases', 'approval_tokens', 'meta']));
    // delta_links / delta_items ship with U12 (the delta-sync mirror).
    expect(tables).toEqual(expect.arrayContaining(['delta_links', 'delta_items']));
    raw.close();
  });

  it.runIf(isPosix)('creates the dir 0700 and state.db 0600, repairing a loose file on boot', () => {
    openStore();
    expect(statSync(dir).mode & 0o777).toBe(0o700);
    const dbPath = join(dir, 'state.db');
    expect(statSync(dbPath).mode & 0o777).toBe(0o600);

    // Loosen the file, reopen → repaired back to 0600.
    chmodSync(dbPath, 0o644);
    openStore();
    expect(statSync(dbPath).mode & 0o777).toBe(0o600);
  });

  it.runIf(isPosix)('restricts the -wal/-shm sidecars to 0600 (they hold token data too)', () => {
    openStore();
    // WAL sidecars are created during the migration writes at open.
    for (const suffix of ['-wal', '-shm']) {
      const p = join(dir, `state.db${suffix}`);
      if (existsSync(p)) {
        expect(statSync(p).mode & 0o777, `${suffix} perms`).toBe(0o600);
      }
    }
  });

  it('degrades to in-memory when the db was migrated by a newer build (downgrade guard)', () => {
    // Simulate a future build: stamp a schema_version beyond this build's count.
    const raw = new Database(join(dir, 'state.db'));
    raw.exec('CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);');
    raw.prepare("INSERT INTO meta (key, value) VALUES ('schema_version', ?)").run(String(SCHEMA_VERSION + 5));
    raw.close();

    const warn = vi.fn();
    const store = StateStore.open({ dir, legacyDir, warn });
    open.add(store);
    expect(store.degraded).toBe(true);
    expect(warn).toHaveBeenCalledOnce();
  });
});

describe('StateStore aliases (account stamping)', () => {
  it('resolves an alias and hides rows from a foreign account', () => {
    const store = openStore();
    store.putAlias({ token: 'em_abc', graphId: 'AAA', entityType: 'message', accountId: 'acct-X', mutable: false });

    expect(store.getAlias('em_abc', 'acct-X')?.graphId).toBe('AAA');
    // Scoped to a different account → invisible (account is mandatory / fail-closed).
    expect(store.getAlias('em_abc', 'acct-Y')).toBeNull();
  });
});

describe('StateStore approval tokens (atomic consume)', () => {
  it('a token consumed in one instance cannot be consumed in another', () => {
    const a = openStore();
    a.putApprovalToken(approval({ token: 'ap_x' }));

    const b = openStore(); // second connection on the same db file
    const first = a.consumeApprovalToken({ token: 'ap_x', accountId: 'acct-X', receiptJson: '{"ok":true}', now: 1 });
    expect(first.status).toBe('consumed');

    const second = b.consumeApprovalToken({ token: 'ap_x', accountId: 'acct-X', now: 2 });
    expect(second.status).toBe('already_redeemed');
    // The loser can still recover the receipt written by the winner.
    if (second.status === 'already_redeemed') {
      expect(second.receiptJson).toBe('{"ok":true}');
    }
  });

  it('refuses a token belonging to a different account', () => {
    const store = openStore();
    store.putApprovalToken(approval({ token: 'ap_y', accountId: 'acct-X' }));
    expect(store.consumeApprovalToken({ token: 'ap_y', accountId: 'acct-Y', now: 1 }).status).toBe('foreign_account');
    // ...and the real owner can still consume it.
    expect(store.consumeApprovalToken({ token: 'ap_y', accountId: 'acct-X', now: 2 }).status).toBe('consumed');
  });

  it('returns not_found for an unknown token', () => {
    expect(openStore().consumeApprovalToken({ token: 'nope', accountId: 'acct-X', now: 1 }).status).toBe('not_found');
  });

  it('refuses to consume an expired-but-unpurged token (expiry enforced at the store)', () => {
    const store = openStore();
    store.putApprovalToken(approval({ token: 'ap_exp', expiresAt: 5000 }));
    // now is past expiry but the row still exists (not yet purged).
    expect(store.consumeApprovalToken({ token: 'ap_exp', accountId: 'acct-X', now: 6000 }).status).toBe('expired');
    // A still-valid consume of a fresh token works.
    store.putApprovalToken(approval({ token: 'ap_ok', operationKey: 'op-ok', expiresAt: 9000 }));
    expect(store.consumeApprovalToken({ token: 'ap_ok', accountId: 'acct-X', now: 6000 }).status).toBe('consumed');
  });

  it('throws on a duplicate token (PK) — the U9 mint contract must not silently overwrite', () => {
    const store = openStore();
    store.putApprovalToken(approval({ token: 'dup' }));
    expect(() => store.putApprovalToken(approval({ token: 'dup' }))).toThrow();
  });
});

describe('StateStore degraded mode', () => {
  it('falls back to in-memory on a corrupt db file, warns, and still works', () => {
    // Write garbage where state.db would live so opening the file fails.
    writeFileSync(join(dir, 'state.db'), 'this is not a sqlite database');
    const warn = vi.fn();
    const store = StateStore.open({ dir, legacyDir, warn });
    open.add(store);

    expect(store.degraded).toBe(true);
    expect(store.path).toBe(':memory:');
    expect(warn).toHaveBeenCalledOnce();

    // Operations still succeed against the in-memory schema.
    store.putAlias({ token: 'em_1', graphId: 'G', entityType: 'message', accountId: 'acct-X' });
    expect(store.getAlias('em_1', 'acct-X')?.graphId).toBe('G');
  });
});

describe('StateStore legacy migration', () => {
  it('copies legacy tokens.json once when the new dir has none, and never overwrites', () => {
    writeFileSync(join(legacyDir, 'tokens.json'), 'LEGACY');
    openStore();
    expect(readFileSync(join(dir, 'tokens.json'), 'utf-8')).toBe('LEGACY');

    // Simulate the user re-authenticating (new token cache), then reopen:
    // the one-shot marker must prevent re-copy/overwrite.
    writeFileSync(join(dir, 'tokens.json'), 'FRESH');
    openStore();
    expect(readFileSync(join(dir, 'tokens.json'), 'utf-8')).toBe('FRESH');
  });

  it('does not overwrite an existing new tokens.json even on first boot', () => {
    writeFileSync(join(legacyDir, 'tokens.json'), 'LEGACY');
    writeFileSync(join(dir, 'tokens.json'), 'EXISTING');
    openStore();
    expect(readFileSync(join(dir, 'tokens.json'), 'utf-8')).toBe('EXISTING');
  });

  it('is a no-op when there is no legacy tokens.json', () => {
    openStore();
    expect(existsSync(join(dir, 'tokens.json'))).toBe(false);
  });
});

describe('StateStore purge (90-day retention)', () => {
  it('purges tokens expired beyond the retention window and keeps the rest', () => {
    const now = 1_000_000_000_000;
    const store = openStore(() => now);
    // Expired long ago (beyond 90 days) → purged.
    store.putApprovalToken(approval({ token: 'old', expiresAt: now - APPROVAL_RETENTION_MS - 1 }));
    // Expired recently (within retention) → kept as an idempotency receipt.
    store.putApprovalToken(approval({ token: 'recent', operationKey: 'op-recent', expiresAt: now - 1000 }));
    // Not yet expired → kept.
    store.putApprovalToken(approval({ token: 'future', operationKey: 'op-future', expiresAt: now + 100_000 }));

    const purged = store.purgeExpired(now);
    expect(purged).toBe(1);
    expect(store.consumeApprovalToken({ token: 'old', accountId: 'acct-X', now }).status).toBe('not_found');
    // 'recent' is retained (within 90d) but expired → consume refuses it.
    expect(store.consumeApprovalToken({ token: 'recent', accountId: 'acct-X', now }).status).toBe('expired');
    expect(store.consumeApprovalToken({ token: 'future', accountId: 'acct-X', now }).status).toBe('consumed');
  });

  it('runs the boot purge automatically on open', () => {
    const now = 2_000_000_000_000;
    // Seed an old row via one store, close, then reopen → boot purge removes it.
    const seed = openStore(() => now);
    seed.putApprovalToken(approval({ token: 'stale', expiresAt: now - APPROVAL_RETENTION_MS - 1 }));
    seed.close();
    open.delete(seed);

    const reopened = openStore(() => now);
    expect(reopened.consumeApprovalToken({ token: 'stale', accountId: 'acct-X', now }).status).toBe('not_found');
  });
});
