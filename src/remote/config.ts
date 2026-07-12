/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Remote-mode auth configuration (U4). Loaded once at `serve` startup with
 * fail-fast diagnostics (the v4.2.0 pattern) so a misconfigured deployment
 * refuses to start rather than silently accepting the wrong tokens.
 *
 * The server is a pure resource server: it validates Entra-issued JWTs whose
 * audience is the connector's API app. It never mints tokens and holds no
 * client secret here (the On-Behalf-Of credential lives in U5).
 */

const TENANT_PLACEHOLDERS = new Set(['common', 'organizations', 'consumers']);
const DEFAULT_APP_ID_URI = 'api://mcp-office365-connector';

/** Resolved remote-mode auth configuration. */
export interface RemoteAuthConfig {
  /** JP tenant GUID — sign-in is single-tenant to this directory. */
  readonly tenantId: string;
  /** v2 token issuer that JWTs must carry. */
  readonly issuer: string;
  /** Tenant JWKS endpoint (signing keys). */
  readonly jwksUri: string;
  /** Accepted `aud` values: the API app's GUID and its Application ID URI. */
  readonly allowedAudiences: readonly string[];
  /** Canonical public MCP URL (RFC 8707 resource / RFC 9728 PRM resource). */
  readonly publicUrl: string;
  /** Delegated scope the token must carry. */
  readonly requiredScope: string;
}

function requireEnv(name: string, value: string | undefined): string {
  if (value == null || value.trim() === '') {
    throw new Error(
      `${name} is required for remote (serve) mode. See docs/remote/provisioning.md ` +
        `for the connector's app-registration values.`,
    );
  }
  return value.trim();
}

/**
 * Loads and validates remote auth config from the environment. Throws (fail
 * fast) on any missing or invalid value.
 */
export function loadRemoteAuthConfig(env: NodeJS.ProcessEnv = process.env): RemoteAuthConfig {
  const tenantId = requireEnv('OUTLOOK_MCP_TENANT_ID', env.OUTLOOK_MCP_TENANT_ID);
  if (TENANT_PLACEHOLDERS.has(tenantId.toLowerCase())) {
    throw new Error(
      `OUTLOOK_MCP_TENANT_ID must be a specific tenant GUID for remote mode, not ` +
        `'${tenantId}'. Member-vs-guest enforcement requires a single directory.`,
    );
  }

  const apiClientId = requireEnv(
    'OUTLOOK_MCP_CONNECTOR_API_ID',
    env.OUTLOOK_MCP_CONNECTOR_API_ID,
  );
  const publicUrl = normalizeResourceUrl(
    requireEnv('OUTLOOK_MCP_CONNECTOR_URL', env.OUTLOOK_MCP_CONNECTOR_URL),
  );
  const appIdUri = (env.OUTLOOK_MCP_CONNECTOR_APP_ID_URI ?? DEFAULT_APP_ID_URI).trim();

  return {
    tenantId,
    issuer: `https://login.microsoftonline.com/${tenantId}/v2.0`,
    jwksUri: `https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`,
    // Accept both the bare GUID and the api:// identifier form (tokens may carry
    // either); reject anything else (anti-token-passthrough).
    allowedAudiences: [apiClientId, appIdUri],
    publicUrl,
    requiredScope: 'access_as_user',
  };
}

/**
 * Canonical RFC 8707 resource form: lowercase scheme/host, strip a trailing
 * slash and default port. Path is preserved. claude.ai sends this exact form as
 * `resource`, and the PRM `resource` must match it.
 */
export function normalizeResourceUrl(raw: string): string {
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    throw new Error(`OUTLOOK_MCP_CONNECTOR_URL is not a valid URL: ${raw}`);
  }
  if (url.protocol !== 'https:' && url.hostname !== 'localhost' && url.hostname !== '127.0.0.1') {
    throw new Error(`OUTLOOK_MCP_CONNECTOR_URL must be https (got ${url.protocol}//).`);
  }
  const port =
    (url.protocol === 'https:' && url.port === '443') ||
    (url.protocol === 'http:' && url.port === '80')
      ? ''
      : url.port;
  const host = port === '' ? url.hostname.toLowerCase() : `${url.hostname.toLowerCase()}:${port}`;
  const path = url.pathname === '/' ? '' : url.pathname.replace(/\/$/, '');
  return `${url.protocol}//${host}${path}`;
}
