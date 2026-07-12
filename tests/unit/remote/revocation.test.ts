/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/** U7 revocation: deny-list store, per-account purge, and revoke ordering. */

import { afterEach, describe, expect, it } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { StateStore } from '../../../src/state/store.js';
import { createStoreDenyList } from '../../../src/remote/auth/deny-list.js';
import { listRevoked, readmitUser, revokeUser } from '../../../src/remote/revocation.js';

const dirs: string[] = [];
const stores: StateStore[] = [];
function store(): StateStore {
  const dir = mkdtempSync(join(tmpdir(), 'mcp-u7-'));
  dirs.push(dir);
  const s = StateStore.open({ dir });
  stores.push(s);
  return s;
}
afterEach(() => {
  while (stores.length > 0) stores.pop()?.close();
  while (dirs.length > 0) {
    const d = dirs.pop();
    if (d != null) rmSync(d, { recursive: true, force: true });
  }
});

describe('deny-list store (U7)', () => {
  it('deny / isDenied / readmit / list', () => {
    const s = store();
    expect(s.isDenied('oid-1')).toBe(false);
    s.denyUser('oid-1', 'left the org', 1000);
    expect(s.isDenied('oid-1')).toBe(true);
    expect(s.listDenied()).toEqual([{ oid: 'oid-1', reason: 'left the org', deniedAt: 1000 }]);
    expect(s.readmitUser('oid-1')).toBe(true);
    expect(s.isDenied('oid-1')).toBe(false);
    expect(s.readmitUser('oid-1')).toBe(false); // idempotent
  });

  it('createStoreDenyList reflects the store per call', () => {
    const s = store();
    const dl = createStoreDenyList(s);
    expect(dl.isDenied('oid-2')).toBe(false);
    s.denyUser('oid-2', null, 1);
    expect(dl.isDenied('oid-2')).toBe(true); // no process cache
  });
});

describe('purgeAccount (U7)', () => {
  it('removes only the target account’s durable state', () => {
    const s = store();
    s.putAlias({ token: 'tm_a', graphId: 'g1', entityType: 'team', accountId: 'user-a.tid' });
    s.putAlias({ token: 'tm_b', graphId: 'g2', entityType: 'team', accountId: 'user-b.tid' });
    s.putApprovalToken({
      token: 'ap_a', action: 'delete', targetJson: '{}', accountId: 'user-a.tid', expiresAt: Date.now() + 1e6,
    });

    const purged = s.purgeAccount('user-a.tid');
    expect(purged).toBeGreaterThanOrEqual(2); // alias + approval token
    expect(s.getAlias('tm_a', 'user-a.tid')).toBeNull();
    expect(s.getApprovalToken('ap_a', 'user-a.tid')).toBeNull();
    // User B untouched.
    expect(s.getAlias('tm_b', 'user-b.tid')).not.toBeNull();
  });
});

describe('revokeUser (U7)', () => {
  it('deny-lists AND purges, deny-list before purge', () => {
    const s = store();
    s.putAlias({ token: 'tm_x', graphId: 'g', entityType: 'team', accountId: 'oid-9.tid' });
    const result = revokeUser(s, 'oid-9', 'oid-9.tid', 5000, 'departed');
    expect(s.isDenied('oid-9')).toBe(true);
    expect(result.purgedRows).toBeGreaterThanOrEqual(1);
    expect(s.getAlias('tm_x', 'oid-9.tid')).toBeNull();
    expect(listRevoked(s)[0]).toMatchObject({ oid: 'oid-9', reason: 'departed' });
  });

  it('readmit removes the deny-list entry', () => {
    const s = store();
    revokeUser(s, 'oid-10', 'oid-10.tid', 1, undefined);
    expect(s.isDenied('oid-10')).toBe(true);
    expect(readmitUser(s, 'oid-10')).toBe(true);
    expect(s.isDenied('oid-10')).toBe(false);
  });
});
