/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { StateStore } from '../../../src/state/store.js';
import { resolveId } from '../../../src/ids/resolver.js';
import { registerComposite } from '../../../src/ids/mint.js';
import { mintSelfEncoded, canonicalKey, mintComposite } from '../../../src/ids/token.js';
import { ErrorCode } from '../../../src/utils/errors.js';

let dir: string;
let store: StateStore;
const ACCOUNT = 'acct-X';

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'mcp-ids-'));
  store = StateStore.open({ dir, legacyDir: join(dir, 'legacy'), warn: () => {} });
});

afterEach(() => {
  store.close();
  rmSync(dir, { recursive: true, force: true });
});

describe('resolveId — self-encoding (cold-state durable)', () => {
  it('resolves a self-encoded token with zero storage, even on a fresh store', () => {
    const graphId = 'AAMkAGI2-immutable-id';
    const token = mintSelfEncoded('message', graphId);
    // A brand-new store instance (empty alias table) still resolves it.
    expect(resolveId(token, ACCOUNT, store)).toEqual({ graphId, mutable: false });
  });

  it('survives a discarded state.db between processes (the core cold-state fix)', () => {
    const token = mintSelfEncoded('event', 'EVT-123');
    store.close();
    rmSync(join(dir, 'state.db'), { force: true });
    const fresh = StateStore.open({ dir, legacyDir: join(dir, 'legacy'), warn: () => {} });
    try {
      expect(resolveId(token, ACCOUNT, fresh).graphId).toBe('EVT-123');
    } finally {
      fresh.close();
    }
  });
});

describe('resolveId — alias-backed', () => {
  it('mints, stores, and resolves a composite token', () => {
    const token = registerComposite(store, {
      entityType: 'attachment',
      parts: { messageId: 'M', attachmentId: 'A' },
      graphId: 'ATT-GRAPH-ID',
      accountId: ACCOUNT,
    });
    expect(resolveId(token, ACCOUNT, store)).toEqual({ graphId: 'ATT-GRAPH-ID', mutable: false });
  });

  it('carries the mutable flag through to the resolved result', () => {
    const token = registerComposite(store, {
      entityType: 'plannerTask',
      parts: { id: 'PT1' },
      graphId: 'PT-GRAPH',
      accountId: ACCOUNT,
      mutable: true,
    });
    expect(resolveId(token, ACCOUNT, store).mutable).toBe(true);
  });

  it('a cold composite token (empty alias table) yields ID_UNKNOWN with a re-list hint', () => {
    const token = mintComposite('attachment', canonicalKey('attachment', { messageId: 'M', attachmentId: 'A' }));
    try {
      resolveId(token, ACCOUNT, store);
      expect.unreachable('should throw');
    } catch (e) {
      expect((e as { code?: string }).code).toBe(ErrorCode.ID_UNKNOWN);
      expect((e as { suggestion?: string }).suggestion).toMatch(/re-list/i);
    }
  });

  it('a token minted under a different account yields ID_FOREIGN_ACCOUNT', () => {
    const token = registerComposite(store, {
      entityType: 'chat',
      parts: { id: 'C1' },
      graphId: 'CHAT-GRAPH',
      accountId: 'acct-OTHER',
    });
    try {
      resolveId(token, ACCOUNT, store);
      expect.unreachable('should throw');
    } catch (e) {
      expect((e as { code?: string }).code).toBe(ErrorCode.ID_FOREIGN_ACCOUNT);
    }
  });
});

describe('resolveId — legacy numeric + raw pass-through', () => {
  it('rejects a legacy numeric (v2 hash) ID with NUMERIC_ID_UNSUPPORTED', () => {
    try {
      resolveId(123456, ACCOUNT, store);
      expect.unreachable('should throw');
    } catch (e) {
      expect((e as { code?: string }).code).toBe(ErrorCode.NUMERIC_ID_UNSUPPORTED);
    }
  });

  it('passes a raw non-token string through as an opaque Graph ID', () => {
    expect(resolveId('AAMkAGI2rawgraphid', ACCOUNT, store)).toEqual({
      graphId: 'AAMkAGI2rawgraphid',
      mutable: false,
    });
  });
});

describe('resolveId — entity-type guard', () => {
  it('rejects a token for the wrong entity kind with ID_ENTITY_MISMATCH', () => {
    const folderToken = mintSelfEncoded('folder', 'FOLDER-1');
    try {
      resolveId(folderToken, ACCOUNT, store, 'contact');
      expect.unreachable('should throw');
    } catch (e) {
      expect((e as { code?: string }).code).toBe(ErrorCode.ID_ENTITY_MISMATCH);
    }
  });

  it('accepts a token whose entity kind matches the expected type', () => {
    const contactToken = mintSelfEncoded('contact', 'CONTACT-1');
    expect(resolveId(contactToken, ACCOUNT, store, 'contact').graphId).toBe('CONTACT-1');
  });

  it('still passes a raw Graph ID through when an entity type is expected', () => {
    expect(resolveId('rawid', ACCOUNT, store, 'contact').graphId).toBe('rawid');
  });
});

describe('registerComposite — collision policy (D1a)', () => {
  it('is idempotent for the same entity (re-mint returns the same token, no throw)', () => {
    const args = {
      entityType: 'attachment' as const,
      parts: { messageId: 'M', attachmentId: 'A' },
      graphId: 'ATT-1',
      accountId: ACCOUNT,
    };
    const t1 = registerComposite(store, args);
    const t2 = registerComposite(store, args);
    expect(t1).toBe(t2);
  });

  it('does NOT throw on graphId drift of a mutable entity — it updates (D2)', () => {
    const args = {
      entityType: 'plannerTask' as const,
      parts: { id: 'PT1' },
      graphId: 'PT-OLD',
      accountId: ACCOUNT,
      mutable: true,
    };
    const token = registerComposite(store, args);
    // The same mutable entity re-listed after its Graph ID drifted → update.
    const token2 = registerComposite(store, { ...args, graphId: 'PT-NEW' });
    expect(token2).toBe(token);
    expect(resolveId(token, ACCOUNT, store).graphId).toBe('PT-NEW');
  });

  it('gives each account a distinct token for a shared resource — no ownership churn', () => {
    const parts = { id: 'SHARED-CHAT' };
    const tokenA = registerComposite(store, { entityType: 'chat', parts, graphId: 'CHAT-G', accountId: 'acct-A' });
    const tokenB = registerComposite(store, { entityType: 'chat', parts, graphId: 'CHAT-G', accountId: 'acct-B' });
    // Distinct tokens → account B listing the shared chat does not lock A out.
    expect(tokenA).not.toBe(tokenB);
    expect(resolveId(tokenA, 'acct-A', store).graphId).toBe('CHAT-G');
    expect(resolveId(tokenB, 'acct-B', store).graphId).toBe('CHAT-G');
  });

  it('throws ID_COLLISION when a different Graph ID already occupies an immutable token', () => {
    // Compute the token exactly as registerComposite does (account folded in),
    // then pre-seed it with a DIFFERENT graph id — simulating a digest collision
    // from a different canonical key.
    const token = mintComposite(
      'attachment',
      canonicalKey('attachment', { messageId: 'M', attachmentId: 'A', '@account': ACCOUNT }),
    );
    store.putAlias({ token, graphId: 'DIFFERENT', entityType: 'attachment', accountId: ACCOUNT });
    try {
      registerComposite(store, {
        entityType: 'attachment',
        parts: { messageId: 'M', attachmentId: 'A' },
        graphId: 'ATT-1',
        accountId: ACCOUNT,
      });
      expect.unreachable('should throw');
    } catch (e) {
      expect((e as { code?: string }).code).toBe(ErrorCode.ID_COLLISION);
    }
  });
});
