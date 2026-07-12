/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/** U4 config fail-fast + PRM/WWW-Authenticate shape. */

import { describe, expect, it } from 'vitest';
import {
  loadRemoteAuthConfig,
  normalizeResourceUrl,
  type RemoteAuthConfig,
} from '../../../../src/remote/config.js';
import {
  PRM_PATH,
  protectedResourceMetadata,
  wwwAuthenticate,
} from '../../../../src/remote/auth/metadata.js';

const TID = '761e2c5f-0000-4000-8000-000000000001';

function env(over: Record<string, string | undefined> = {}): NodeJS.ProcessEnv {
  return {
    OUTLOOK_MCP_TENANT_ID: TID,
    OUTLOOK_MCP_CONNECTOR_API_ID: 'api-guid',
    OUTLOOK_MCP_CONNECTOR_URL: 'https://mcp.example.com/mcp',
    ...over,
  } as NodeJS.ProcessEnv;
}

describe('loadRemoteAuthConfig (U4)', () => {
  it('derives issuer, jwks, and audiences from the tenant + api id', () => {
    const c = loadRemoteAuthConfig(env());
    expect(c.issuer).toBe(`https://login.microsoftonline.com/${TID}/v2.0`);
    expect(c.jwksUri).toBe(`https://login.microsoftonline.com/${TID}/discovery/v2.0/keys`);
    expect(c.allowedAudiences).toEqual(['api-guid', 'api://mcp-office365-connector']);
    expect(c.requiredScope).toBe('access_as_user');
  });

  it('fails fast on a placeholder tenant (member enforcement needs one directory)', () => {
    expect(() => loadRemoteAuthConfig(env({ OUTLOOK_MCP_TENANT_ID: 'common' }))).toThrow(
      /specific tenant GUID/,
    );
  });

  it('fails fast when required env is missing', () => {
    expect(() => loadRemoteAuthConfig(env({ OUTLOOK_MCP_CONNECTOR_URL: undefined }))).toThrow(
      /OUTLOOK_MCP_CONNECTOR_URL is required/,
    );
    expect(() => loadRemoteAuthConfig(env({ OUTLOOK_MCP_CONNECTOR_API_ID: undefined }))).toThrow(
      /OUTLOOK_MCP_CONNECTOR_API_ID is required/,
    );
  });

  it('rejects a non-https connector URL', () => {
    expect(() => loadRemoteAuthConfig(env({ OUTLOOK_MCP_CONNECTOR_URL: 'http://evil.example/mcp' })))
      .toThrow(/must be https/);
  });

  it('never puts an empty string into allowedAudiences (empty override → default URI)', () => {
    const c = loadRemoteAuthConfig(env({ OUTLOOK_MCP_CONNECTOR_APP_ID_URI: '   ' }));
    expect(c.allowedAudiences).toEqual(['api-guid', 'api://mcp-office365-connector']);
    expect(c.allowedAudiences).not.toContain('');
  });

  it('lowercases an uppercase tenant GUID (issuer-match robustness)', () => {
    const c = loadRemoteAuthConfig(env({ OUTLOOK_MCP_TENANT_ID: TID.toUpperCase() }));
    expect(c.issuer).toBe(`https://login.microsoftonline.com/${TID}/v2.0`);
  });
});

describe('normalizeResourceUrl', () => {
  it('lowercases host, strips default port and trailing slash', () => {
    expect(normalizeResourceUrl('https://MCP.Example.com:443/mcp/')).toBe('https://mcp.example.com/mcp');
    expect(normalizeResourceUrl('https://mcp.example.com/')).toBe('https://mcp.example.com');
    expect(normalizeResourceUrl('https://mcp.example.com:8443/mcp')).toBe(
      'https://mcp.example.com:8443/mcp',
    );
  });
});

describe('PRM + WWW-Authenticate (U4)', () => {
  const config: RemoteAuthConfig = loadRemoteAuthConfig(env());

  it('serves a spec-shaped PRM document pinned to config values', () => {
    const prm = protectedResourceMetadata(config);
    expect(prm.resource).toBe('https://mcp.example.com/mcp');
    expect(prm.authorization_servers).toEqual([config.issuer]);
    expect(prm.scopes_supported).toEqual(['access_as_user']);
    expect(prm.bearer_methods_supported).toEqual(['header']);
  });

  it('builds a WWW-Authenticate header pointing at the PRM URL', () => {
    const header = wwwAuthenticate(config, 'invalid_token');
    expect(header).toContain('Bearer ');
    expect(header).toContain(`resource_metadata="https://mcp.example.com/mcp${PRM_PATH}"`);
    expect(header).toContain('error="invalid_token"');
  });
});
