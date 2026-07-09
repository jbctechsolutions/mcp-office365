/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Approval token manager (U9). Manages the lifecycle of two-phase approval
 * tokens: generation, validation, and atomic single-use consumption.
 *
 * When a durable {@link StateStore} is supplied (production), tokens persist to
 * SQLite with a 24h TTL and are consumed via an atomic guarded write (D8), so an
 * approval prepared in one process/session survives a restart and can be
 * redeemed exactly once — even across concurrent Claude Code windows. Without a
 * store (tests / degraded mode) it falls back to an in-memory map, preserving
 * the original behavior and public surface.
 */

import { randomUUID } from 'node:crypto';
import type { StateStore, ApprovalTokenRow } from '../state/store.js';
import type { OperationType, TargetType, ApprovalToken, ValidationResult } from './types.js';

const DEFAULT_TTL_MS = 24 * 60 * 60 * 1000; // 24h (D8, up from 5 minutes)
const CLEANUP_THRESHOLD = 100;

/** The target tuple persisted in the store's `target_json` column. */
interface StoredTarget {
  targetType: TargetType;
  targetId: string;
  targetHash: string;
  metadata: Record<string, unknown>;
}

/** Options for {@link ApprovalTokenManager}. */
export interface ApprovalTokenManagerOptions {
  /** Token lifetime (default 24h). */
  ttlMs?: number;
  /** Durable store; omit for in-memory (tests / degraded mode). */
  store?: StateStore;
  /**
   * Account stamp for stored tokens (D7). May be a function, since the signed-in
   * account (MSAL `homeAccountId`) is only known after auth — later than when
   * this manager is constructed — and is resolved lazily at each token op.
   */
  accountId?: string | (() => string);
  /** Clock, for deterministic tests. */
  now?: () => number;
}

export class ApprovalTokenManager {
  private readonly tokens = new Map<string, ApprovalToken>();
  private readonly ttlMs: number;
  private readonly store: StateStore | undefined;
  private readonly resolveAccountId: () => string;
  private readonly now: () => number;

  constructor(options: ApprovalTokenManagerOptions | number = {}) {
    // Back-compat: some call sites/tests pass a bare ttlMs number.
    const opts: ApprovalTokenManagerOptions = typeof options === 'number' ? { ttlMs: options } : options;
    this.ttlMs = opts.ttlMs ?? DEFAULT_TTL_MS;
    this.store = opts.store;
    const account = opts.accountId ?? 'default';
    this.resolveAccountId = typeof account === 'function' ? account : ((): string => account);
    this.now = opts.now ?? ((): number => Date.now());
  }

  /** The account scope for store ops, resolved fresh each call (D7). */
  private get accountId(): string {
    return this.resolveAccountId();
  }

  /**
   * Generates a new approval token for a destructive operation.
   */
  generateToken(params: {
    operation: OperationType;
    targetType: TargetType;
    targetId: string;
    targetHash: string;
    metadata?: Record<string, unknown>;
  }): ApprovalToken {
    const now = this.now();
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

    if (this.store != null) {
      const target: StoredTarget = {
        targetType: token.targetType,
        targetId: token.targetId,
        targetHash: token.targetHash,
        metadata: { ...token.metadata },
      };
      this.store.putApprovalToken({
        token: token.tokenId,
        action: token.operation,
        targetJson: JSON.stringify(target),
        contentHash: token.targetHash,
        accountId: this.accountId,
        expiresAt: token.expiresAt,
        createdAt: token.createdAt,
      });
    } else {
      if (this.tokens.size > CLEANUP_THRESHOLD) {
        this.cleanupExpiredTokens();
      }
      this.tokens.set(token.tokenId, token);
    }
    return token;
  }

  /**
   * Looks up a token by ID without consuming or validating it. Returns undefined
   * if the token does not exist (or belongs to another account, when durable).
   */
  lookupToken(tokenId: string): ApprovalToken | undefined {
    if (this.store != null) {
      const row = this.store.getApprovalToken(tokenId, this.accountId);
      return row != null ? (rowToToken(row) ?? undefined) : undefined;
    }
    return this.tokens.get(tokenId);
  }

  /**
   * Validates a token without consuming it. Checks existence, prior redemption,
   * expiry, operation match, and target match.
   */
  validateToken(tokenId: string, operation: OperationType, targetId: string): ValidationResult {
    if (this.store != null) {
      const row = this.store.getApprovalToken(tokenId, this.accountId);
      if (row == null) {
        return { valid: false, error: 'NOT_FOUND' };
      }
      if (row.redeemedAt != null) {
        return { valid: false, error: 'ALREADY_CONSUMED' };
      }
      const token = rowToToken(row);
      // A corrupt/tampered target_json fails closed as NOT_FOUND rather than
      // throwing out of validate — the token simply can't be redeemed.
      if (token == null) {
        return { valid: false, error: 'NOT_FOUND' };
      }
      return this.check(token, operation, targetId);
    }

    const token = this.tokens.get(tokenId);
    if (token == null) {
      return { valid: false, error: 'NOT_FOUND' };
    }
    return this.check(token, operation, targetId);
  }

  private check(token: ApprovalToken, operation: OperationType, targetId: string): ValidationResult {
    // `>=` so validate agrees with the store's consume guard (expires_at > now):
    // a token is expired at the exact expiry instant, on both paths.
    if (this.now() >= token.expiresAt) {
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
   * Validates and consumes a token (one-time use). Durable stores use an atomic
   * guarded consume so only one caller — across processes sharing the db — can
   * redeem it; a losing/repeat caller gets `ALREADY_CONSUMED` (D8).
   */
  consumeToken(tokenId: string, operation: OperationType, targetId: string): ValidationResult {
    const result = this.validateToken(tokenId, operation, targetId);
    if (!result.valid) {
      return result;
    }

    if (this.store != null) {
      const consumed = this.store.consumeApprovalToken({
        token: tokenId,
        accountId: this.accountId,
        now: this.now(),
      });
      switch (consumed.status) {
        case 'consumed':
          return result;
        case 'already_redeemed':
          return { valid: false, error: 'ALREADY_CONSUMED' };
        case 'expired':
          return { valid: false, error: 'EXPIRED' };
        default:
          return { valid: false, error: 'NOT_FOUND' };
      }
    }

    this.tokens.delete(tokenId);
    return result;
  }

  /**
   * Removes expired tokens from the in-memory store. Durable stores purge on
   * boot (90-day retention), so this is a no-op there.
   */
  cleanupExpiredTokens(): void {
    if (this.store != null) {
      return;
    }
    const now = this.now();
    for (const [tokenId, token] of this.tokens) {
      if (now > token.expiresAt) {
        this.tokens.delete(tokenId);
      }
    }
  }

  /**
   * Number of active in-memory tokens (testing/monitoring). Durable tokens are
   * not counted here — they live in the store.
   */
  get size(): number {
    return this.tokens.size;
  }
}

/** Reconstructs an {@link ApprovalToken} from a stored row, or null if the
 * persisted target JSON is corrupt/unparseable (fail closed). */
function rowToToken(row: ApprovalTokenRow): ApprovalToken | null {
  let target: StoredTarget;
  try {
    target = JSON.parse(row.targetJson) as StoredTarget;
  } catch {
    return null;
  }
  return {
    tokenId: row.token,
    operation: row.action as OperationType,
    targetType: target.targetType,
    targetId: target.targetId,
    targetHash: target.targetHash ?? row.contentHash ?? '',
    createdAt: row.createdAt,
    expiresAt: row.expiresAt,
    metadata: Object.freeze({ ...target.metadata }),
  };
}
