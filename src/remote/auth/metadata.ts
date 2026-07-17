/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * RFC 9728 Protected Resource Metadata + the 401 `WWW-Authenticate` challenge
 * (U4). This is the discovery handshake a claude.ai custom connector follows:
 * an unauthenticated request gets a 401 pointing at the PRM document, which in
 * turn points at the Entra authorization server. All values come from config —
 * never from request headers (a `Host`-header value here would let a caller
 * poison discovery).
 */

import type { RemoteAuthConfig } from '../config.js';

/** RFC 9728 metadata path, and the path-suffixed variant claude.ai also probes. */
export const PRM_PATH = '/.well-known/oauth-protected-resource';

/**
 * RFC 9728 §3.1 metadata URL for the resource: `/.well-known/oauth-protected-
 * resource` is inserted between the host and the resource's path, NOT appended
 * to the full resource URL. For `https://host/mcp` the metadata lives at
 * `https://host/.well-known/oauth-protected-resource/mcp` — which is exactly the
 * path-suffixed route the server serves. (Appending to the full URL instead —
 * `https://host/mcp/.well-known/oauth-protected-resource` — 404s, so a client
 * following the WWW-Authenticate hint can't discover the AS.)
 */
export function resourceMetadataUrl(config: RemoteAuthConfig): string {
  const url = new URL(config.publicUrl);
  const path = url.pathname === '/' ? '' : url.pathname;
  return `${url.origin}${PRM_PATH}${path}`;
}

/** The Protected Resource Metadata document (RFC 9728). */
export function protectedResourceMetadata(config: RemoteAuthConfig): Record<string, unknown> {
  return {
    resource: config.publicUrl,
    // claude.ai uses only the first entry; pin it to the JP tenant issuer.
    authorization_servers: [config.issuer],
    scopes_supported: [config.requiredScope],
    bearer_methods_supported: ['header'],
  };
}

/**
 * The `WWW-Authenticate` header value for a 401, carrying the PRM URL so the
 * client can discover the authorization server (RFC 9728 §5.1).
 */
export function wwwAuthenticate(config: RemoteAuthConfig, error?: string): string {
  const params = [`resource_metadata="${resourceMetadataUrl(config)}"`];
  if (error != null) {
    params.unshift(`error="${error}"`);
  }
  return `Bearer ${params.join(', ')}`;
}
