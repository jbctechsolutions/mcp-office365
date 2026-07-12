/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Entra JWT validation for remote mode (U4). Fail-closed member-only enforcement:
 * a token is accepted only on a positive delegated-member signal, and any absent
 * optional claim is treated as rejection — never as pass. Signature keys come
 * from the configured tenant JWKS only (never the token's own `iss`).
 */

import { createRemoteJWKSet, jwtVerify } from 'jose';
import type { RemoteAuthConfig } from '../config.js';

/** Clock skew tolerance for exp/nbf/iat (seconds) — tight on synced cloud clocks. */
const CLOCK_TOLERANCE_S = 120;

/** Validated caller identity attached to the request for downstream units. */
export interface RemoteIdentity {
  /** Entra object id (stable user id). */
  readonly oid: string;
  /** Tenant id. */
  readonly tid: string;
  /** MSAL homeAccountId form (`<oid>.<tid>`) — the per-user state key (U5). */
  readonly homeAccountId: string;
  /** Delegated scopes on the token. */
  readonly scopes: readonly string[];
  /** Expiry (seconds since epoch), when present. */
  readonly expiresAt?: number;
}

/** Token failed validation — respond 401 so the client re-authenticates. */
export class AuthChallengeError extends Error {
  readonly reason: string;
  constructor(reason: string, message?: string) {
    super(message ?? `invalid_token: ${reason}`);
    this.name = 'AuthChallengeError';
    this.reason = reason;
  }
}

/** Authenticated but not authorized (guest, app-only, wrong scope, deny-listed) — respond 403. */
export class AuthForbiddenError extends Error {
  readonly reason: string;
  constructor(reason: string, message?: string) {
    super(message ?? `forbidden: ${reason}`);
    this.name = 'AuthForbiddenError';
    this.reason = reason;
  }
}

/** Key source — the tenant JWKS by default; tests inject a local key set. */
export type KeySource = Parameters<typeof jwtVerify>[1];

/**
 * Builds a verifier that resolves a bearer token to a {@link RemoteIdentity} or
 * throws {@link AuthChallengeError} (401) / {@link AuthForbiddenError} (403).
 */
export function createTokenVerifier(
  config: RemoteAuthConfig,
  keySource?: KeySource,
): (token: string) => Promise<RemoteIdentity> {
  // createRemoteJWKSet caches keys and refetches on an unknown `kid` with a
  // built-in cooldown (bounds JWKS fetches; fails closed on outage).
  const keys: KeySource = keySource ?? createRemoteJWKSet(new URL(config.jwksUri));

  return async function verify(token: string): Promise<RemoteIdentity> {
    let payload: Record<string, unknown>;
    try {
      const result = await jwtVerify(token, keys, {
        issuer: config.issuer,
        audience: [...config.allowedAudiences],
        algorithms: ['RS256'], // pin RS256; reject none/HS*
        clockTolerance: CLOCK_TOLERANCE_S,
      });
      payload = result.payload as Record<string, unknown>;
    } catch (error) {
      // Signature, iss, aud, exp, nbf, alg failures — all token-invalid.
      throw new AuthChallengeError(
        'verification_failed',
        `invalid_token: ${error instanceof Error ? error.name : 'verification failed'}`,
      );
    }

    // Fail-closed member-only enforcement — a positive signal is required.
    if (payload.tid !== config.tenantId) {
      throw new AuthForbiddenError('foreign_tenant');
    }
    // idtyp 'app' or a roles-without-scp shape = app-only token; reject.
    if (payload.idtyp === 'app') {
      throw new AuthForbiddenError('app_only_token');
    }
    const scopes = typeof payload.scp === 'string' ? payload.scp.split(' ').filter(Boolean) : [];
    if (scopes.length === 0) {
      // No delegated scope claim — not a delegated-user token. Fail closed.
      throw new AuthForbiddenError('no_delegated_scope');
    }
    // acct: 0 = member, 1 = guest. Must be present AND 0.
    if (payload.acct !== 0) {
      throw new AuthForbiddenError('not_member');
    }
    if (!scopes.includes(config.requiredScope)) {
      throw new AuthForbiddenError('insufficient_scope');
    }

    const oid = typeof payload.oid === 'string' ? payload.oid : '';
    const tid = typeof payload.tid === 'string' ? payload.tid : '';
    if (oid === '' || tid === '') {
      throw new AuthChallengeError('missing_identity');
    }

    return {
      oid,
      tid,
      homeAccountId: `${oid}.${tid}`,
      scopes,
      ...(typeof payload.exp === 'number' ? { expiresAt: payload.exp } : {}),
    };
  };
}
