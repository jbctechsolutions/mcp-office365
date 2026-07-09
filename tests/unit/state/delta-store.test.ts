/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { StateStore } from '../../../src/state/store.js';
import type { MirrorItem } from '../../../src/state/delta-store.js';

let dir: string;
let legacyDir: string;
let store: StateStore;

const ACCT = 'acct-1';
const RES = 'mail:inbox';

function item(id: string, summary = `subject-${id}`): MirrorItem {
  return { graphId: id, token: `em_${id}`, summary, snapshot: JSON.stringify({ id }) };
}

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'mcp-delta-'));
  legacyDir = mkdtempSync(join(tmpdir(), 'mcp-delta-legacy-'));
  store = StateStore.open({ dir, legacyDir, warn: () => {} });
});

afterEach(() => {
  try { store.close(); } catch { /* already closed */ }
  rmSync(dir, { recursive: true, force: true });
  rmSync(legacyDir, { recursive: true, force: true });
});

describe('DeltaStore', () => {
  it('reports no cursor and no seen ids before any sync', () => {
    expect(store.delta.getDeltaLink(ACCT, RES)).toBeNull();
    expect(store.delta.getSeenIds(ACCT, RES).size).toBe(0);
    expect(store.delta.countItems(ACCT, RES)).toBe(0);
  });

  it('persists the cursor and upserted items on commit', () => {
    store.delta.commit({
      accountId: ACCT,
      resource: RES,
      deltaLink: 'https://graph/delta?token=abc',
      syncedAt: 1000,
      upserts: [item('a'), item('b')],
      deletes: [],
    });

    expect(store.delta.getDeltaLink(ACCT, RES)).toBe('https://graph/delta?token=abc');
    expect([...store.delta.getSeenIds(ACCT, RES)].sort()).toEqual(['a', 'b']);
    expect(store.delta.countItems(ACCT, RES)).toBe(2);
    expect(store.delta.getItem(ACCT, RES, 'a')?.summary).toBe('subject-a');
  });

  it('updates an existing item and removes deleted ones', () => {
    store.delta.commit({ accountId: ACCT, resource: RES, deltaLink: 'l1', syncedAt: 1, upserts: [item('a'), item('b')], deletes: [] });
    store.delta.commit({
      accountId: ACCT,
      resource: RES,
      deltaLink: 'l2',
      syncedAt: 2,
      upserts: [item('a', 'renamed')],
      deletes: ['b'],
    });

    expect(store.delta.getItem(ACCT, RES, 'a')?.summary).toBe('renamed');
    expect(store.delta.getItem(ACCT, RES, 'b')).toBeNull();
    expect(store.delta.getDeltaLink(ACCT, RES)).toBe('l2');
  });

  it('clears the cursor (forcing a re-baseline) when the deltaLink is empty', () => {
    store.delta.commit({ accountId: ACCT, resource: RES, deltaLink: 'l1', syncedAt: 1, upserts: [item('a')], deletes: [] });
    store.delta.commit({ accountId: ACCT, resource: RES, deltaLink: '', syncedAt: 2, upserts: [item('b')], deletes: [] });

    // Cursor gone, but mirror rows retained.
    expect(store.delta.getDeltaLink(ACCT, RES)).toBeNull();
    expect(store.delta.countItems(ACCT, RES)).toBe(2);
  });

  it('scopes cursors and items by account', () => {
    store.delta.commit({ accountId: ACCT, resource: RES, deltaLink: 'l1', syncedAt: 1, upserts: [item('a')], deletes: [] });
    expect(store.delta.getDeltaLink('other', RES)).toBeNull();
    expect(store.delta.getSeenIds('other', RES).size).toBe(0);
  });

  it('reset(resource) drops only that resource; reset(account) drops all', () => {
    store.delta.commit({ accountId: ACCT, resource: 'mail:inbox', deltaLink: 'l1', syncedAt: 1, upserts: [item('a')], deletes: [] });
    store.delta.commit({ accountId: ACCT, resource: 'calendar:primary', deltaLink: 'l2', syncedAt: 1, upserts: [item('c')], deletes: [] });

    store.delta.reset(ACCT, 'mail:inbox');
    expect(store.delta.getDeltaLink(ACCT, 'mail:inbox')).toBeNull();
    expect(store.delta.countItems(ACCT, 'mail:inbox')).toBe(0);
    expect(store.delta.getDeltaLink(ACCT, 'calendar:primary')).toBe('l2');

    store.delta.reset(ACCT);
    expect(store.delta.getDeltaLink(ACCT, 'calendar:primary')).toBeNull();
    expect(store.delta.countItems(ACCT, 'calendar:primary')).toBe(0);
  });
});
