import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

import { hashEmailForApproval, hashFolderForApproval } from '../../../src/approval/hash.js';
import { ApprovalTokenManager } from '../../../src/approval/token-manager.js';

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

        // Advance time past the default 5-minute TTL
        vi.advanceTimersByTime(5 * 60 * 1000 + 1);

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

        vi.advanceTimersByTime(5 * 60 * 1000 + 1);

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
        vi.advanceTimersByTime(5 * 60 * 1000 + 1);

        manager.cleanupExpiredTokens();
        expect(manager.size).toBe(0);
      } finally {
        vi.useRealTimers();
      }
    });

    it('keeps non-expired tokens during cleanup', () => {
      vi.useFakeTimers();
      try {
        manager.generateToken(defaultParams);

        // Advance time but stay within TTL
        vi.advanceTimersByTime(2 * 60 * 1000);

        // Add another token at the later time
        manager.generateToken({ ...defaultParams, targetId: 2 });

        // Advance so first token expires but second does not
        vi.advanceTimersByTime(3 * 60 * 1000 + 1);

        manager.cleanupExpiredTokens();
        expect(manager.size).toBe(1);
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
});
