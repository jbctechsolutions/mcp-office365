/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * U4 JWT validation — the security-critical unit. Mints real RS256 tokens with a
 * local keypair and drives every accept/reject path, including the fail-closed
 * member checks (absent optional claims must reject, never pass).
 */

import { beforeAll, describe, expect, it } from 'vitest';
import { SignJWT, generateKeyPair, importJWK, type CryptoKey } from 'jose';
import { loadRemoteAuthConfig, type RemoteAuthConfig } from '../../../../src/remote/config.js';
import {
  AuthChallengeError,
  AuthForbiddenError,
  createTokenVerifier,
} from '../../../../src/remote/auth/verify.js';

const TID = '761e2c5f-0000-4000-8000-000000000001';
const OID = 'aaaaaaaa-0000-4000-8000-000000000002';
const API_ID = '484c0657-0000-4000-8000-000000000003';

const config: RemoteAuthConfig = loadRemoteAuthConfig({
  OUTLOOK_MCP_TENANT_ID: TID,
  OUTLOOK_MCP_CONNECTOR_API_ID: API_ID,
  OUTLOOK_MCP_CONNECTOR_URL: 'https://mcp.example.com/mcp',
} as NodeJS.ProcessEnv);

let privateKey: CryptoKey;
let publicKey: CryptoKey;
let hmacKey: CryptoKey;
let verify: (token: string) => Promise<unknown>;

/** Base valid member-token claims; individual tests override fields. */
function claims(over: Record<string, unknown> = {}): Record<string, unknown> {
  return { tid: TID, oid: OID, acct: 0, idtyp: 'user', scp: 'access_as_user', ...over };
}

async function sign(
  payload: Record<string, unknown>,
  opts: { aud?: string; iss?: string; alg?: string; expiresIn?: string; key?: CryptoKey } = {},
): Promise<string> {
  return new SignJWT(payload)
    .setProtectedHeader({ alg: opts.alg ?? 'RS256' })
    .setIssuer(opts.iss ?? config.issuer)
    .setAudience(opts.aud ?? API_ID)
    .setIssuedAt()
    .setExpirationTime(opts.expiresIn ?? '5m')
    .sign(opts.key ?? privateKey);
}

beforeAll(async () => {
  const pair = await generateKeyPair('RS256');
  privateKey = pair.privateKey;
  publicKey = pair.publicKey;
  // Re-import a symmetric key for the HS256 downgrade test.
  hmacKey = (await importJWK(
    { kty: 'oct', k: 'c2VjcmV0LXNlY3JldC1zZWNyZXQtc2VjcmV0LXNlY3JldA' },
    'HS256',
  )) as CryptoKey;
  // Verifier keyed on the local public key (no network JWKS in tests).
  verify = createTokenVerifier(config, publicKey);
});

describe('createTokenVerifier (U4)', () => {
  it('accepts a valid member token and returns identity', async () => {
    const id = (await verify(await sign(claims()))) as {
      oid: string;
      tid: string;
      homeAccountId: string;
      scopes: string[];
    };
    expect(id.oid).toBe(OID);
    expect(id.tid).toBe(TID);
    expect(id.homeAccountId).toBe(`${OID}.${TID}`);
    expect(id.scopes).toContain('access_as_user');
  });

  it('lowercases oid/tid so deny-list and account keys match byte-for-byte', async () => {
    const id = (await verify(await sign(claims({ oid: OID.toUpperCase() })))) as {
      oid: string;
      homeAccountId: string;
    };
    expect(id.oid).toBe(OID); // lowercased
    expect(id.homeAccountId).toBe(`${OID}.${TID}`);
  });

  it('accepts the api:// audience form as well as the GUID', async () => {
    await expect(verify(await sign(claims(), { aud: 'api://mcp-office365-connector' }))).resolves
      .toBeDefined();
  });

  it('rejects a guest (acct=1) as forbidden', async () => {
    await expect(verify(await sign(claims({ acct: 1 })))).rejects.toBeInstanceOf(AuthForbiddenError);
  });

  it('rejects a token with acct absent (fail closed, not member)', async () => {
    const c = claims();
    delete c.acct;
    await expect(verify(await sign(c))).rejects.toMatchObject({ reason: 'not_member' });
  });

  it('rejects an app-only token (idtyp=app)', async () => {
    await expect(verify(await sign(claims({ idtyp: 'app' })))).rejects.toMatchObject({
      reason: 'app_only_token',
    });
  });

  it('rejects a token with no delegated scope', async () => {
    const c = claims();
    delete c.scp;
    await expect(verify(await sign(c))).rejects.toMatchObject({ reason: 'no_delegated_scope' });
  });

  it('rejects a token whose scope lacks access_as_user', async () => {
    await expect(verify(await sign(claims({ scp: 'User.Read' })))).rejects.toMatchObject({
      reason: 'insufficient_scope',
    });
  });

  it('rejects a foreign-tenant tid claim', async () => {
    await expect(
      verify(await sign(claims({ tid: '00000000-0000-4000-8000-0000000000ff' }))),
    ).rejects.toMatchObject({ reason: 'foreign_tenant' });
  });

  it('rejects a wrong audience with a 401 challenge', async () => {
    await expect(verify(await sign(claims(), { aud: 'api://someone-else' }))).rejects.toBeInstanceOf(
      AuthChallengeError,
    );
  });

  it('rejects a wrong issuer with a 401 challenge', async () => {
    await expect(
      verify(await sign(claims(), { iss: 'https://login.microsoftonline.com/evil/v2.0' })),
    ).rejects.toBeInstanceOf(AuthChallengeError);
  });

  it('rejects an expired token with a 401 challenge', async () => {
    await expect(verify(await sign(claims(), { expiresIn: '-10m' }))).rejects.toBeInstanceOf(
      AuthChallengeError,
    );
  });

  it('rejects an HS256-signed token (algorithm pinned to RS256)', async () => {
    await expect(
      verify(await sign(claims(), { alg: 'HS256', key: hmacKey })),
    ).rejects.toBeInstanceOf(AuthChallengeError);
  });

  it('rejects an alg=none (unsigned) token', async () => {
    const b64 = (o: unknown): string => Buffer.from(JSON.stringify(o)).toString('base64url');
    const none = `${b64({ alg: 'none', typ: 'JWT' })}.${b64(claims())}.`;
    await expect(verify(none)).rejects.toBeInstanceOf(AuthChallengeError);
  });

  it('rejects a token minted for Microsoft Graph (audience confusion)', async () => {
    await expect(
      verify(await sign(claims(), { aud: '00000003-0000-0000-c000-000000000000' })),
    ).rejects.toBeInstanceOf(AuthChallengeError);
  });

  it('rejects acct as the string "0" (fail closed — only integer 0 is a member)', async () => {
    await expect(verify(await sign(claims({ acct: '0' })))).rejects.toMatchObject({
      reason: 'not_member',
    });
  });
});
