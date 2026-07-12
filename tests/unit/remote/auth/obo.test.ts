/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/** U5 On-Behalf-Of: credential loading, token exchange, and error mapping. */

import { beforeEach, describe, expect, it, vi } from 'vitest';

// vi.hoisted so the (hoisted) mock factory can reference the spy.
const { acquireMock } = vi.hoisted(() => ({ acquireMock: vi.fn() }));
vi.mock('@azure/msal-node', () => ({
  ConfidentialClientApplication: class {
    acquireTokenOnBehalfOf(...args: unknown[]): unknown {
      return acquireMock(...args);
    }
  },
}));

import { createOboClient, loadOboCredential, mapOboError } from '../../../../src/remote/auth/obo.js';
import { loadRemoteAuthConfig } from '../../../../src/remote/config.js';
import { GraphAuthRequiredError, GraphError } from '../../../../src/utils/errors.js';

const config = loadRemoteAuthConfig({
  OUTLOOK_MCP_TENANT_ID: '761e2c5f-0000-4000-8000-000000000001',
  OUTLOOK_MCP_CONNECTOR_API_ID: 'api-guid',
  OUTLOOK_MCP_CONNECTOR_URL: 'https://mcp.example.com/mcp',
} as NodeJS.ProcessEnv);

beforeEach(() => acquireMock.mockReset());

/** Await a rejection and return the thrown error (robust vs .rejects matchers). */
async function caught(p: Promise<unknown>): Promise<Error> {
  try {
    await p;
  } catch (e) {
    return e as Error;
  }
  throw new Error('expected a rejection');
}

describe('loadOboCredential', () => {
  it('prefers a certificate when both cert and secret are set', () => {
    const cred = loadOboCredential({
      OUTLOOK_MCP_CONNECTOR_CERT_THUMBPRINT: 'THUMB',
      OUTLOOK_MCP_CONNECTOR_CERT_KEY: '-----BEGIN PRIVATE KEY-----\n...',
      OUTLOOK_MCP_CONNECTOR_CLIENT_SECRET: 'sekret',
    } as NodeJS.ProcessEnv);
    expect(cred).toMatchObject({ kind: 'certificate', thumbprint: 'THUMB' });
  });

  it('falls back to a client secret', () => {
    expect(loadOboCredential({ OUTLOOK_MCP_CONNECTOR_CLIENT_SECRET: 'sekret' } as NodeJS.ProcessEnv))
      .toMatchObject({ kind: 'secret', clientSecret: 'sekret' });
  });

  it('returns null when nothing is configured (OBO disabled)', () => {
    expect(loadOboCredential({} as NodeJS.ProcessEnv)).toBeNull();
  });
});

describe('createOboClient.acquireGraphToken', () => {
  const obo = createOboClient(config, { kind: 'secret', clientSecret: 'sekret' });

  it('exchanges the inbound assertion for a Graph token', async () => {
    acquireMock.mockResolvedValue({ accessToken: 'graph-token' });
    await expect(obo.acquireGraphToken('inbound-jwt')).resolves.toBe('graph-token');
    expect(acquireMock).toHaveBeenCalledWith(
      expect.objectContaining({
        oboAssertion: 'inbound-jwt',
        scopes: ['https://graph.microsoft.com/.default'],
      }),
    );
  });

  it('throws when MSAL returns no access token', async () => {
    acquireMock.mockResolvedValue({ accessToken: '' });
    expect(await caught(obo.acquireGraphToken('x'))).toBeInstanceOf(GraphAuthRequiredError);
  });

});

describe('mapOboError', () => {
  it('maps missing consent (AADSTS65001) to an actionable GraphError', () => {
    const e = mapOboError(new Error('AADSTS65001: The user or admin has not consented'));
    expect(e).toBeInstanceOf(GraphError);
    expect(e.message).toContain('consent');
  });

  it('maps invalid_grant / interaction_required to a re-auth error', () => {
    expect(mapOboError(new Error('invalid_grant: token expired'))).toBeInstanceOf(
      GraphAuthRequiredError,
    );
    expect(mapOboError(new Error('AADSTS50076: interaction_required'))).toBeInstanceOf(
      GraphAuthRequiredError,
    );
  });

  it('maps an expired confidential credential (AADSTS7000222) to a rotation error', () => {
    const e = mapOboError(new Error('AADSTS7000222: client secret expired'));
    expect(e).toBeInstanceOf(GraphError);
    expect(e.message).toContain('credential is invalid or');
  });

  it('maps an unknown failure to a generic GraphError', () => {
    expect(mapOboError(new Error('something odd'))).toBeInstanceOf(GraphError);
  });
});
