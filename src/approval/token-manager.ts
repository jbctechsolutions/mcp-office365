/**
 * Approval token manager.
 *
 * Manages the lifecycle of approval tokens: generation, validation,
 * consumption (one-time use), and expiry cleanup.
 */

import { randomUUID } from 'node:crypto';
import type { OperationType, TargetType, ApprovalToken, ValidationResult } from './types.js';

// =============================================================================
// Constants
// =============================================================================

const DEFAULT_TTL_MS = 5 * 60 * 1000; // 5 minutes
const CLEANUP_THRESHOLD = 100;

// =============================================================================
// Token Manager
// =============================================================================

/**
 * Manages approval tokens for destructive operations.
 *
 * Tokens are stored in memory and are single-use. They expire
 * after a configurable TTL (default 5 minutes).
 */
export class ApprovalTokenManager {
  private readonly tokens = new Map<string, ApprovalToken>();
  private readonly ttlMs: number;

  constructor(ttlMs: number = DEFAULT_TTL_MS) {
    this.ttlMs = ttlMs;
  }

  /**
   * Generates a new approval token for a destructive operation.
   */
  generateToken(params: {
    operation: OperationType;
    targetType: TargetType;
    targetId: number;
    targetHash: string;
    metadata?: Record<string, unknown>;
  }): ApprovalToken {
    // Clean up expired tokens if the map is getting large
    if (this.tokens.size > CLEANUP_THRESHOLD) {
      this.cleanupExpiredTokens();
    }

    const now = Date.now();
    const token: ApprovalToken = {
      tokenId: randomUUID(),
      operation: params.operation,
      targetType: params.targetType,
      targetId: params.targetId,
      targetHash: params.targetHash,
      createdAt: now,
      expiresAt: now + this.ttlMs,
      metadata: Object.freeze({ ...params.metadata }),
    };

    this.tokens.set(token.tokenId, token);
    return token;
  }

  /**
   * Validates a token without consuming it or modifying state.
   * Checks existence, expiry, operation match, and target match.
   */
  validateToken(
    tokenId: string,
    operation: OperationType,
    targetId: number
  ): ValidationResult {
    const token = this.tokens.get(tokenId);

    if (token == null) {
      return { valid: false, error: 'NOT_FOUND' };
    }

    if (Date.now() > token.expiresAt) {
      return { valid: false, error: 'EXPIRED' };
    }

    if (token.operation !== operation) {
      return { valid: false, error: 'OPERATION_MISMATCH' };
    }

    if (token.targetId !== targetId) {
      return { valid: false, error: 'TARGET_MISMATCH' };
    }

    return { valid: true, token };
  }

  /**
   * Validates and consumes a token (one-time use).
   * On success, the token is removed from the store.
   */
  consumeToken(
    tokenId: string,
    operation: OperationType,
    targetId: number
  ): ValidationResult {
    const result = this.validateToken(tokenId, operation, targetId);

    if (result.valid) {
      this.tokens.delete(tokenId);
    }

    return result;
  }

  /**
   * Removes all expired tokens from the store.
   */
  cleanupExpiredTokens(): void {
    const now = Date.now();
    for (const [tokenId, token] of this.tokens) {
      if (now > token.expiresAt) {
        this.tokens.delete(tokenId);
      }
    }
  }

  /**
   * Returns the number of active (non-expired) tokens.
   * Useful for testing and monitoring.
   */
  get size(): number {
    return this.tokens.size;
  }
}
