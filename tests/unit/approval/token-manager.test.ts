/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

import {
  hashEmailForApproval,
  hashFolderForApproval,
  hashDraftForSend,
  hashDirectSendForApproval,
  hashReplyForApproval,
  hashForwardForApproval,
  hashEventForApproval,
  hashContactForApproval,
  hashTaskForApproval,
} from '../../../src/approval/hash.js';
import { ApprovalTokenManager } from '../../../src/approval/token-manager.js';
import { StateStore } from '../../../src/state/store.js';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

// =============================================================================
// hashEmailForApproval
// =============================================================================

describe('hashEmailForApproval', () => {
  const email = { id: 1, subject: 'Hello', folderId: 10, timeReceived: 1700000000 };

  it('returns a consistent hash for the same input', () => {
    const hash1 = hashEmailForApproval(email);
    const hash2 = hashEmailForApproval(email);
    expect(hash1).toBe(hash2);
  });

  it('returns different hashes for different inputs', () => {
    const other = { id: 2, subject: 'Goodbye', folderId: 20, timeReceived: 1700000001 };
    expect(hashEmailForApproval(email)).not.toBe(hashEmailForApproval(other));
  });

  it('returns a 16-character hex string', () => {
    const hash = hashEmailForApproval(email);
    expect(hash).toMatch(/^[0-9a-f]{16}$/);
  });

  it('handles null subject', () => {
    const emailNullSubject = { id: 1, subject: null, folderId: 10, timeReceived: 1700000000 };
    const hash = hashEmailForApproval(emailNullSubject);
    expect(hash).toMatch(/^[0-9a-f]{16}$/);
    // null subject should produce a different hash than a non-null subject
    expect(hash).not.toBe(hashEmailForApproval(email));
  });

  it('handles null timeReceived', () => {
    const emailNullTime = { id: 1, subject: 'Hello', folderId: 10, timeReceived: null };
    const hash = hashEmailForApproval(emailNullTime);
    expect(hash).toMatch(/^[0-9a-f]{16}$/);
    // null timeReceived should produce a different hash than a non-null timeReceived
    expect(hash).not.toBe(hashEmailForApproval(email));
  });
});

// =============================================================================
// hashFolderForApproval
// =============================================================================

describe('hashFolderForApproval', () => {
  const folder = { id: 5, name: 'Inbox', messageCount: 42 };

  it('returns a consistent hash for the same input', () => {
    const hash1 = hashFolderForApproval(folder);
    const hash2 = hashFolderForApproval(folder);
    expect(hash1).toBe(hash2);
  });

  it('returns different hashes for different inputs', () => {
    const other = { id: 6, name: 'Trash', messageCount: 0 };
    expect(hashFolderForApproval(folder)).not.toBe(hashFolderForApproval(other));
  });

  it('returns a 16-character hex string', () => {
    const hash = hashFolderForApproval(folder);
    expect(hash).toMatch(/^[0-9a-f]{16}$/);
  });

  it('handles null name', () => {
    const folderNullName = { id: 5, name: null, messageCount: 42 };
    const hash = hashFolderForApproval(folderNullName);
    expect(hash).toMatch(/^[0-9a-f]{16}$/);
    // null name should produce a different hash than a non-null name
    expect(hash).not.toBe(hashFolderForApproval(folder));
  });
});

// =============================================================================
// hashDraftForSend
// =============================================================================

describe('hashDraftForSend', () => {
  const draft = { id: 1, subject: 'Draft email', recipientCount: 3 };

  it('returns a consistent hash for the same input', () => {
    expect(hashDraftForSend(draft)).toBe(hashDraftForSend(draft));
  });

  it('returns different hashes for different inputs', () => {
    const other = { id: 2, subject: 'Other draft', recipientCount: 1 };
    expect(hashDraftForSend(draft)).not.toBe(hashDraftForSend(other));
  });

  it('returns a 16-character hex string', () => {
    expect(hashDraftForSend(draft)).toMatch(/^[0-9a-f]{16}$/);
  });

  it('handles null subject', () => {
    const nullSubject = { id: 1, subject: null, recipientCount: 3 };
    const hash = hashDraftForSend(nullSubject);
    expect(hash).toMatch(/^[0-9a-f]{16}$/);
    expect(hash).not.toBe(hashDraftForSend(draft));
  });
});

// =============================================================================
// hashDirectSendForApproval
// =============================================================================

describe('hashDirectSendForApproval', () => {
  const params = { subject: 'Hello', toCount: 2, ccCount: 1, bccCount: 0 };

  it('returns a consistent hash for the same input', () => {
    expect(hashDirectSendForApproval(params)).toBe(hashDirectSendForApproval(params));
  });

  it('returns different hashes for different inputs', () => {
    const other = { subject: 'Goodbye', toCount: 1, ccCount: 0, bccCount: 3 };
    expect(hashDirectSendForApproval(params)).not.toBe(hashDirectSendForApproval(other));
  });

  it('returns a 16-character hex string', () => {
    expect(hashDirectSendForApproval(params)).toMatch(/^[0-9a-f]{16}$/);
  });
});

// =============================================================================
// hashReplyForApproval
// =============================================================================

describe('hashReplyForApproval', () => {
  const params = { originalId: 42, commentLength: 100, replyAll: false };

  it('returns a consistent hash for the same input', () => {
    expect(hashReplyForApproval(params)).toBe(hashReplyForApproval(params));
  });

  it('returns different hashes for different inputs', () => {
    const other = { originalId: 43, commentLength: 200, replyAll: true };
    expect(hashReplyForApproval(params)).not.toBe(hashReplyForApproval(other));
  });

  it('returns a 16-character hex string', () => {
    expect(hashReplyForApproval(params)).toMatch(/^[0-9a-f]{16}$/);
  });

  it('distinguishes replyAll true from false', () => {
    const replyAllTrue = { ...params, replyAll: true };
    expect(hashReplyForApproval(params)).not.toBe(hashReplyForApproval(replyAllTrue));
  });
});

// =============================================================================
// hashForwardForApproval
// =============================================================================

describe('hashForwardForApproval', () => {
  const params = { originalId: 42, recipientCount: 3 };

  it('returns a consistent hash for the same input', () => {
    expect(hashForwardForApproval(params)).toBe(hashForwardForApproval(params));
  });

  it('returns different hashes for different inputs', () => {
    const other = { originalId: 43, recipientCount: 1 };
    expect(hashForwardForApproval(params)).not.toBe(hashForwardForApproval(other));
  });

  it('returns a 16-character hex string', () => {
    expect(hashForwardForApproval(params)).toMatch(/^[0-9a-f]{16}$/);
  });
});

// =============================================================================
// hashEventForApproval
// =============================================================================

describe('hashEventForApproval', () => {
  const event = { id: 10, subject: 'Meeting', startDateTime: '2026-02-23T10:00:00' };

  it('returns a consistent hash for the same input', () => {
    expect(hashEventForApproval(event)).toBe(hashEventForApproval(event));
  });

  it('returns different hashes for different inputs', () => {
    const other = { id: 11, subject: 'Standup', startDateTime: '2026-02-24T09:00:00' };
    expect(hashEventForApproval(event)).not.toBe(hashEventForApproval(other));
  });

  it('returns a 16-character hex string', () => {
    expect(hashEventForApproval(event)).toMatch(/^[0-9a-f]{16}$/);
  });

  it('handles null subject', () => {
    const nullSubject = { id: 10, subject: null, startDateTime: '2026-02-23T10:00:00' };
    const hash = hashEventForApproval(nullSubject);
    expect(hash).toMatch(/^[0-9a-f]{16}$/);
    expect(hash).not.toBe(hashEventForApproval(event));
  });

  it('handles null startDateTime', () => {
    const nullStart = { id: 10, subject: 'Meeting', startDateTime: null };
    const hash = hashEventForApproval(nullStart);
    expect(hash).toMatch(/^[0-9a-f]{16}$/);
    expect(hash).not.toBe(hashEventForApproval(event));
  });
});

// =============================================================================
// hashContactForApproval
// =============================================================================

describe('hashContactForApproval', () => {
  const contact = { id: 20, displayName: 'John Doe', emailAddress: 'john@example.com' };

  it('returns a consistent hash for the same input', () => {
    expect(hashContactForApproval(contact)).toBe(hashContactForApproval(contact));
  });

  it('returns different hashes for different inputs', () => {
    const other = { id: 21, displayName: 'Jane Smith', emailAddress: 'jane@example.com' };
    expect(hashContactForApproval(contact)).not.toBe(hashContactForApproval(other));
  });

  it('returns a 16-character hex string', () => {
    expect(hashContactForApproval(contact)).toMatch(/^[0-9a-f]{16}$/);
  });

  it('handles null displayName', () => {
    const nullName = { id: 20, displayName: null, emailAddress: 'john@example.com' };
    const hash = hashContactForApproval(nullName);
    expect(hash).toMatch(/^[0-9a-f]{16}$/);
    expect(hash).not.toBe(hashContactForApproval(contact));
  });

  it('handles null emailAddress', () => {
    const nullEmail = { id: 20, displayName: 'John Doe', emailAddress: null };
    const hash = hashContactForApproval(nullEmail);
    expect(hash).toMatch(/^[0-9a-f]{16}$/);
    expect(hash).not.toBe(hashContactForApproval(contact));
  });
});

// =============================================================================
// hashTaskForApproval
// =============================================================================

describe('hashTaskForApproval', () => {
  const task = { taskId: 'task-abc-123', title: 'Buy groceries', listId: 'list-xyz' };

  it('returns a consistent hash for the same input', () => {
    expect(hashTaskForApproval(task)).toBe(hashTaskForApproval(task));
  });

  it('returns different hashes for different inputs', () => {
    const other = { taskId: 'task-def-456', title: 'Clean house', listId: 'list-uvw' };
    expect(hashTaskForApproval(task)).not.toBe(hashTaskForApproval(other));
  });

  it('returns a 16-character hex string', () => {
    expect(hashTaskForApproval(task)).toMatch(/^[0-9a-f]{16}$/);
  });

  it('handles null title', () => {
    const nullTitle = { taskId: 'task-abc-123', title: null, listId: 'list-xyz' };
    const hash = hashTaskForApproval(nullTitle);
    expect(hash).toMatch(/^[0-9a-f]{16}$/);
    expect(hash).not.toBe(hashTaskForApproval(task));
  });
});

// =============================================================================
// ApprovalTokenManager
// =============================================================================

describe('ApprovalTokenManager', () => {
  let manager: ApprovalTokenManager;

  const defaultParams = {
    operation: 'delete_email' as const,
    targetType: 'email' as const,
    targetId: 1,
    targetHash: 'abc123',
  };

  beforeEach(() => {
    manager = new ApprovalTokenManager();
  });

  // ---------------------------------------------------------------------------
  // generateToken
  // ---------------------------------------------------------------------------

  describe('generateToken', () => {
    it('creates a token with correct fields', () => {
      const token = manager.generateToken(defaultParams);

      expect(token.tokenId).toEqual(expect.any(String));
      expect(token.tokenId.length).toBeGreaterThan(0);
      expect(token.operation).toBe('delete_email');
      expect(token.targetType).toBe('email');
      expect(token.targetId).toBe(1);
      expect(token.targetHash).toBe('abc123');
      expect(token.createdAt).toEqual(expect.any(Number));
      expect(token.expiresAt).toBeGreaterThan(token.createdAt);
      expect(token.metadata).toBeDefined();
    });

    it('returns unique token IDs', () => {
      const token1 = manager.generateToken(defaultParams);
      const token2 = manager.generateToken(defaultParams);
      expect(token1.tokenId).not.toBe(token2.tokenId);
    });
  });

  // ---------------------------------------------------------------------------
  // validateToken
  // ---------------------------------------------------------------------------

  describe('validateToken', () => {
    it('returns valid for a fresh token', () => {
      const token = manager.generateToken(defaultParams);
      const result = manager.validateToken(token.tokenId, 'delete_email', 1);

      expect(result.valid).toBe(true);
      expect(result.error).toBeUndefined();
      expect(result.token).toBeDefined();
      expect(result.token!.tokenId).toBe(token.tokenId);
    });

    it('returns EXPIRED for an expired token', () => {
      vi.useFakeTimers();
      try {
        const token = manager.generateToken(defaultParams);

        // Advance time past the default 24h TTL
        vi.advanceTimersByTime(24 * 60 * 60 * 1000 + 1);

        const result = manager.validateToken(token.tokenId, 'delete_email', 1);
        expect(result.valid).toBe(false);
        expect(result.error).toBe('EXPIRED');
      } finally {
        vi.useRealTimers();
      }
    });

    it('does not remove expired tokens from the store', () => {
      vi.useFakeTimers();
      try {
        const token = manager.generateToken(defaultParams);
        expect(manager.size).toBe(1);

        vi.advanceTimersByTime(24 * 60 * 60 * 1000 + 1);

        const result = manager.validateToken(token.tokenId, 'delete_email', 1);
        expect(result.valid).toBe(false);
        expect(result.error).toBe('EXPIRED');
        // Token should still be in the store (not eagerly evicted)
        expect(manager.size).toBe(1);
      } finally {
        vi.useRealTimers();
      }
    });

    it('returns NOT_FOUND for an unknown token', () => {
      const result = manager.validateToken('nonexistent-id', 'delete_email', 1);
      expect(result.valid).toBe(false);
      expect(result.error).toBe('NOT_FOUND');
    });

    it('returns OPERATION_MISMATCH for wrong operation', () => {
      const token = manager.generateToken(defaultParams);
      const result = manager.validateToken(token.tokenId, 'move_email', 1);

      expect(result.valid).toBe(false);
      expect(result.error).toBe('OPERATION_MISMATCH');
    });

    it('returns TARGET_MISMATCH for wrong targetId', () => {
      const token = manager.generateToken(defaultParams);
      const result = manager.validateToken(token.tokenId, 'delete_email', 999);

      expect(result.valid).toBe(false);
      expect(result.error).toBe('TARGET_MISMATCH');
    });
  });

  // ---------------------------------------------------------------------------
  // consumeToken
  // ---------------------------------------------------------------------------

  describe('consumeToken', () => {
    it('removes the token after consumption', () => {
      const token = manager.generateToken(defaultParams);
      expect(manager.size).toBe(1);

      const result = manager.consumeToken(token.tokenId, 'delete_email', 1);
      expect(result.valid).toBe(true);
      expect(manager.size).toBe(0);
    });

    it('cannot consume the same token twice (returns NOT_FOUND)', () => {
      const token = manager.generateToken(defaultParams);

      const first = manager.consumeToken(token.tokenId, 'delete_email', 1);
      expect(first.valid).toBe(true);

      const second = manager.consumeToken(token.tokenId, 'delete_email', 1);
      expect(second.valid).toBe(false);
      expect(second.error).toBe('NOT_FOUND');
    });
  });

  // ---------------------------------------------------------------------------
  // cleanupExpiredTokens
  // ---------------------------------------------------------------------------

  describe('cleanupExpiredTokens', () => {
    it('removes expired tokens', () => {
      vi.useFakeTimers();
      try {
        manager.generateToken(defaultParams);
        manager.generateToken({ ...defaultParams, targetId: 2 });
        expect(manager.size).toBe(2);

        // Advance past expiry
        vi.advanceTimersByTime(24 * 60 * 60 * 1000 + 1);

        manager.cleanupExpiredTokens();
        expect(manager.size).toBe(0);
      } finally {
        vi.useRealTimers();
      }
    });

    it('keeps non-expired tokens during cleanup', () => {
      vi.useFakeTimers();
      try {
        // Explicit 5-minute TTL so the relative timings below are meaningful.
        const m = new ApprovalTokenManager(5 * 60 * 1000);
        m.generateToken(defaultParams);

        // Advance time but stay within TTL
        vi.advanceTimersByTime(2 * 60 * 1000);

        // Add another token at the later time
        m.generateToken({ ...defaultParams, targetId: 2 });

        // Advance so first token expires but second does not
        vi.advanceTimersByTime(3 * 60 * 1000 + 1);

        m.cleanupExpiredTokens();
        expect(m.size).toBe(1);
      } finally {
        vi.useRealTimers();
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Auto-cleanup threshold
  // ---------------------------------------------------------------------------

  describe('auto-cleanup on threshold', () => {
    it('triggers cleanup when exceeding 100 tokens', () => {
      vi.useFakeTimers();
      try {
        // Use a short TTL so tokens expire quickly
        const shortManager = new ApprovalTokenManager(1);

        // Generate 101 tokens (all will be immediately expired after time advance)
        for (let i = 0; i < 101; i++) {
          shortManager.generateToken({ ...defaultParams, targetId: i });
        }
        expect(shortManager.size).toBe(101);

        // Advance time so all existing tokens expire
        vi.advanceTimersByTime(2);

        // The next generateToken call should trigger cleanup because size > 100
        shortManager.generateToken({ ...defaultParams, targetId: 999 });

        // All 101 expired tokens should have been cleaned up, leaving only the new one
        expect(shortManager.size).toBe(1);
      } finally {
        vi.useRealTimers();
      }
    });
  });

  // ---------------------------------------------------------------------------
  // size property
  // ---------------------------------------------------------------------------

  describe('size', () => {
    it('returns correct count', () => {
      expect(manager.size).toBe(0);

      manager.generateToken(defaultParams);
      expect(manager.size).toBe(1);

      manager.generateToken({ ...defaultParams, targetId: 2 });
      expect(manager.size).toBe(2);

      manager.generateToken({ ...defaultParams, targetId: 3 });
      expect(manager.size).toBe(3);
    });
  });

  // ---------------------------------------------------------------------------
  // metadata
  // ---------------------------------------------------------------------------

  describe('metadata', () => {
    it('is frozen and passed through correctly', () => {
      const metadata = { reason: 'test cleanup', count: 5 };
      const token = manager.generateToken({ ...defaultParams, metadata });

      expect(token.metadata).toEqual({ reason: 'test cleanup', count: 5 });

      // Verify the metadata object is frozen (Object.freeze was applied)
      expect(Object.isFrozen(token.metadata)).toBe(true);

      // Attempting to mutate should have no effect (strict mode throws, sloppy is silent)
      expect(() => {
        (token.metadata as Record<string, unknown>)['newKey'] = 'should fail';
      }).toThrow();
    });

    it('defaults to an empty frozen object when metadata is not provided', () => {
      const token = manager.generateToken(defaultParams);
      expect(token.metadata).toEqual({});
      expect(Object.isFrozen(token.metadata)).toBe(true);
    });

    it('does not share reference with the original metadata object', () => {
      const metadata: Record<string, unknown> = { key: 'value' };
      const token = manager.generateToken({ ...defaultParams, metadata });

      // Mutating the original should not affect the token's metadata
      metadata['key'] = 'changed';
      expect(token.metadata['key']).toBe('value');
    });
  });

  // ---------------------------------------------------------------------------
  // Send operation types
  // ---------------------------------------------------------------------------

  describe('send operation types', () => {
    it('should accept send_draft operation', () => {
      const token = manager.generateToken({
        operation: 'send_draft',
        targetType: 'email',
        targetId: 1,
        targetHash: 'hash1',
      });
      expect(token.operation).toBe('send_draft');

      const result = manager.validateToken(token.tokenId, 'send_draft', 1);
      expect(result.valid).toBe(true);
    });

    it('should accept send_email operation', () => {
      const token = manager.generateToken({
        operation: 'send_email',
        targetType: 'email',
        targetId: 0,
        targetHash: 'hash2',
      });
      expect(token.operation).toBe('send_email');

      const result = manager.validateToken(token.tokenId, 'send_email', 0);
      expect(result.valid).toBe(true);
    });

    it('should accept reply_email operation', () => {
      const token = manager.generateToken({
        operation: 'reply_email',
        targetType: 'email',
        targetId: 5,
        targetHash: 'hash3',
      });
      expect(token.operation).toBe('reply_email');

      const result = manager.validateToken(token.tokenId, 'reply_email', 5);
      expect(result.valid).toBe(true);
    });

    it('should accept forward_email operation', () => {
      const token = manager.generateToken({
        operation: 'forward_email',
        targetType: 'email',
        targetId: 7,
        targetHash: 'hash4',
      });
      expect(token.operation).toBe('forward_email');

      const result = manager.validateToken(token.tokenId, 'forward_email', 7);
      expect(result.valid).toBe(true);
    });
  });

  // ---------------------------------------------------------------------------
  // Delete operation types for events, contacts, tasks
  // ---------------------------------------------------------------------------

  describe('delete operation types for events, contacts, tasks', () => {
    it('should accept delete_event operation with event target type', () => {
      const token = manager.generateToken({
        operation: 'delete_event',
        targetType: 'event',
        targetId: 10,
        targetHash: 'evhash',
      });
      expect(token.operation).toBe('delete_event');
      expect(token.targetType).toBe('event');

      const result = manager.validateToken(token.tokenId, 'delete_event', 10);
      expect(result.valid).toBe(true);
    });

    it('should accept delete_contact operation with contact target type', () => {
      const token = manager.generateToken({
        operation: 'delete_contact',
        targetType: 'contact',
        targetId: 20,
        targetHash: 'cthash',
      });
      expect(token.operation).toBe('delete_contact');
      expect(token.targetType).toBe('contact');

      const result = manager.validateToken(token.tokenId, 'delete_contact', 20);
      expect(result.valid).toBe(true);
    });

    it('should accept delete_task operation with task target type', () => {
      const token = manager.generateToken({
        operation: 'delete_task',
        targetType: 'task',
        targetId: 30,
        targetHash: 'tkhash',
      });
      expect(token.operation).toBe('delete_task');
      expect(token.targetType).toBe('task');

      const result = manager.validateToken(token.tokenId, 'delete_task', 30);
      expect(result.valid).toBe(true);
    });
  });

  // ---------------------------------------------------------------------------
  // Durable (store-backed) mode — U9b
  // ---------------------------------------------------------------------------

  describe('durable (store-backed) mode', () => {
    let dir: string;
    let store: StateStore;
    const genParams = {
      operation: 'delete_email' as const,
      targetType: 'email' as const,
      targetId: '42',
      targetHash: 'seal-abc',
      metadata: { subject: 'Q3' },
    };

    beforeEach(() => {
      dir = mkdtempSync(join(tmpdir(), 'mcp-approval-'));
      store = StateStore.open({ dir, legacyDir: join(dir, 'legacy'), warn: () => {} });
    });

    afterEach(() => {
      store.close();
      rmSync(dir, { recursive: true, force: true });
    });

    it('persists a token that survives a fresh manager on the same store (restart)', () => {
      const a = new ApprovalTokenManager({ store });
      const token = a.generateToken(genParams);

      // A new manager instance (simulating a server restart) resolves it.
      const b = new ApprovalTokenManager({ store });
      expect(b.lookupToken(token.tokenId)?.targetId).toBe('42');
      expect(b.validateToken(token.tokenId, 'delete_email', '42').valid).toBe(true);
    });

    it('round-trips targetType/targetHash/metadata through the store', () => {
      const m = new ApprovalTokenManager({ store });
      const token = m.generateToken(genParams);
      const looked = new ApprovalTokenManager({ store }).lookupToken(token.tokenId);
      expect(looked?.targetType).toBe('email');
      expect(looked?.targetHash).toBe('seal-abc');
      expect(looked?.metadata).toEqual({ subject: 'Q3' });
    });

    it('consume is atomic + idempotent across manager instances (D8)', () => {
      const a = new ApprovalTokenManager({ store });
      const token = a.generateToken(genParams);

      const b = new ApprovalTokenManager({ store }); // e.g. a second window
      expect(a.consumeToken(token.tokenId, 'delete_email', '42').valid).toBe(true);
      // The second consume — same or different instance — is refused.
      const second = b.consumeToken(token.tokenId, 'delete_email', '42');
      expect(second.valid).toBe(false);
      expect(second.error).toBe('ALREADY_CONSUMED');
    });

    it('validate reports ALREADY_CONSUMED after redemption', () => {
      const m = new ApprovalTokenManager({ store });
      const token = m.generateToken(genParams);
      m.consumeToken(token.tokenId, 'delete_email', '42');
      const v = m.validateToken(token.tokenId, 'delete_email', '42');
      expect(v.error).toBe('ALREADY_CONSUMED');
    });

    it('enforces operation and target match against the persisted token', () => {
      const m = new ApprovalTokenManager({ store });
      const token = m.generateToken(genParams);
      expect(m.validateToken(token.tokenId, 'delete_folder', '42').error).toBe('OPERATION_MISMATCH');
      expect(m.validateToken(token.tokenId, 'delete_email', '999').error).toBe('TARGET_MISMATCH');
    });

    it('expires a durable token past its 24h TTL', () => {
      let clock = 1_000_000_000_000;
      const m = new ApprovalTokenManager({ store, now: () => clock });
      const token = m.generateToken(genParams);
      clock += 24 * 60 * 60 * 1000 + 1;
      expect(m.validateToken(token.tokenId, 'delete_email', '42').error).toBe('EXPIRED');
      // And a consume of an expired token is refused.
      expect(m.consumeToken(token.tokenId, 'delete_email', '42').valid).toBe(false);
    });

    it('coerces a legacy numeric targetId (pre-v4 token) to string so it still validates', () => {
      // Simulate a token persisted by a pre-v4 build: the send_email/upload_file
      // sentinel wrote `targetId` as the JS number 0. v4 consumes with '0'.
      store.putApprovalToken({
        token: 'ap-legacy',
        action: 'send_email',
        targetJson: JSON.stringify({ targetType: 'email', targetId: 0, targetHash: 'seal', metadata: {} }),
        contentHash: null,
        accountId: 'default',
        expiresAt: 4_000_000_000_000,
      });
      const m = new ApprovalTokenManager({ store });
      expect(m.lookupToken('ap-legacy')?.targetId).toBe('0');
      expect(m.validateToken('ap-legacy', 'send_email', '0').valid).toBe(true);
    });

    it('fails closed (NOT_FOUND, no throw) on a corrupt persisted target_json', () => {
      // Simulate a corrupt/tampered row: valid token record but unparseable target.
      store.putApprovalToken({
        token: 'ap-corrupt',
        action: 'delete_email',
        targetJson: '{not valid json',
        contentHash: null,
        accountId: 'default',
        expiresAt: 4_000_000_000_000,
      });
      const m = new ApprovalTokenManager({ store });
      expect(m.lookupToken('ap-corrupt')).toBeUndefined();
      expect(() => m.validateToken('ap-corrupt', 'delete_email', '1')).not.toThrow();
      expect(m.validateToken('ap-corrupt', 'delete_email', '1').error).toBe('NOT_FOUND');
    });

    it('does not resolve a token minted under a different account (D7)', () => {
      const owner = new ApprovalTokenManager({ store, accountId: 'acct-A' });
      const token = owner.generateToken(genParams);
      const other = new ApprovalTokenManager({ store, accountId: 'acct-B' });
      expect(other.lookupToken(token.tokenId)).toBeUndefined();
      expect(other.validateToken(token.tokenId, 'delete_email', '42').error).toBe('NOT_FOUND');
    });

    it('round-trips a string targetId (durable token) through the store and matches on consume', () => {
      // A migrated Graph entity (U5) seals its approval on a ct_ token string,
      // not a number. Persist under one manager, redeem under a fresh one.
      const a = new ApprovalTokenManager({ store });
      const token = a.generateToken({
        operation: 'delete_contact',
        targetType: 'contact',
        targetId: 'ct_Y29udGFjdC0x',
        targetHash: 'seal-ct',
      });

      const b = new ApprovalTokenManager({ store });
      // A different (numeric) targetId must NOT match the string seal.
      expect(b.validateToken(token.tokenId, 'delete_contact', 42).error).toBe('TARGET_MISMATCH');
      // The exact string targetId matches and consumes once.
      expect(b.consumeToken(token.tokenId, 'delete_contact', 'ct_Y29udGFjdC0x').valid).toBe(true);
      expect(b.consumeToken(token.tokenId, 'delete_contact', 'ct_Y29udGFjdC0x').error).toBe('ALREADY_CONSUMED');
    });

    it('resolves accountId lazily (thunk), so an account known only after auth still scopes', () => {
      // The manager is built before sign-in; the account arrives later. A thunk
      // returning the fallback at generate time and the real id at consume time
      // must still key both ops the same when the id is stable across the call.
      let account = 'acct-lazy';
      const m = new ApprovalTokenManager({ store, accountId: () => account });
      const token = m.generateToken(genParams);

      // A second manager reading the same live thunk value resolves it.
      const same = new ApprovalTokenManager({ store, accountId: () => account });
      expect(same.lookupToken(token.tokenId)?.tokenId).toBe(token.tokenId);

      // Once the thunk reports a different account, the token is out of scope.
      account = 'acct-other';
      expect(same.lookupToken(token.tokenId)).toBeUndefined();
    });
  });
});
