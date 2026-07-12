/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Bearer-auth middleware for `/mcp` (U4). Verifies the Entra JWT on every
 * request (never trusts a session for auth), rejects deny-listed identities, and
 * attaches an AuthInfo-compatible identity to the request for downstream units
 * (U5 OBO). Authorization denials are logged as security events — never with any
 * token material — as the only detection surface during the pilot.
 */

import type { NextFunction, Request, Response } from 'express';
import type { AuthInfo } from '@modelcontextprotocol/sdk/server/auth/types.js';
import type { RemoteAuthConfig } from '../config.js';
import type { DenyList } from './deny-list.js';
import { AuthChallengeError, AuthForbiddenError, type RemoteIdentity } from './verify.js';
import { wwwAuthenticate } from './metadata.js';

/** Attach the validated identity (and SDK-compatible auth info) to the request. */
declare module 'express-serve-static-core' {
  interface Request {
    remoteIdentity?: RemoteIdentity;
    auth?: AuthInfo;
  }
}

type Verify = (token: string) => Promise<RemoteIdentity>;

/** Logs an authorization denial (reason + oid when known, no token material). */
function logDenial(reason: string, oid?: string): void {
  const who = oid != null ? ` oid=${oid}` : '';
  process.stderr.write(
    `[mcp-office365] auth denied: reason=${reason}${who} at=${new Date().toISOString()}\n`,
  );
}

function challenge(res: Response, config: RemoteAuthConfig, reason: string): void {
  logDenial(reason);
  res
    .status(401)
    .set('WWW-Authenticate', wwwAuthenticate(config, 'invalid_token'))
    .json({ jsonrpc: '2.0', error: { code: -32001, message: 'Unauthorized.' }, id: null });
}

function forbidden(res: Response, reason: string, oid?: string): void {
  logDenial(reason, oid);
  res
    .status(403)
    .json({ jsonrpc: '2.0', error: { code: -32002, message: 'Forbidden.' }, id: null });
}

/**
 * Express middleware requiring a valid, member, non-deny-listed Entra token.
 * A missing/invalid token → 401 + WWW-Authenticate (RFC 9728 discovery); an
 * authenticated-but-unauthorized identity → 403 (no re-auth loop).
 */
export function createAuthMiddleware(
  config: RemoteAuthConfig,
  verify: Verify,
  denyList: DenyList,
) {
  return async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    const header = req.headers.authorization;
    const token =
      header != null && header.startsWith('Bearer ') ? header.slice('Bearer '.length) : undefined;
    if (token == null || token === '') {
      challenge(res, config, 'missing_token');
      return;
    }

    let identity: RemoteIdentity;
    try {
      identity = await verify(token);
    } catch (error) {
      if (error instanceof AuthForbiddenError) {
        forbidden(res, error.reason);
      } else if (error instanceof AuthChallengeError) {
        challenge(res, config, error.reason);
      } else {
        challenge(res, config, 'verification_error');
      }
      return;
    }

    // Deny-list is read per request from the store (U7) — never process-cached.
    if (denyList.isDenied(identity.oid)) {
      forbidden(res, 'deny_listed', identity.oid);
      return;
    }

    req.remoteIdentity = identity;
    // AuthInfo-compatible shape the SDK exposes to tool handlers as extra.authInfo.
    req.auth = {
      token,
      clientId: '',
      scopes: [...identity.scopes],
      ...(identity.expiresAt != null ? { expiresAt: identity.expiresAt } : {}),
      resource: new URL(config.publicUrl),
      extra: { oid: identity.oid, tid: identity.tid, homeAccountId: identity.homeAccountId },
    };
    next();
  };
}
