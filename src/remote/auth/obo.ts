/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * On-Behalf-Of token exchange (U5). The API app is a confidential client: it
 * exchanges the inbound user assertion (the Entra JWT U4 validated) for a
 * delegated Microsoft Graph token for that same user. Per-user isolation comes
 * from the assertion — MSAL caches each user's OBO result separately, keyed by
 * their identity, so one user's token can never satisfy another's request.
 *
 * The confidential credential (certificate preferred; secret supported for the
 * pilot) lives only in the environment. When neither is configured, OBO fails
 * fast with actionable guidance rather than silently degrading.
 */

import { ConfidentialClientApplication, type Configuration } from '@azure/msal-node';
import type { RemoteAuthConfig } from '../config.js';
import { GraphAuthRequiredError, GraphError, type OutlookMcpError } from '../../utils/errors.js';

/** `.default` returns exactly the app's admin-consented Graph scopes. */
const GRAPH_DEFAULT_SCOPE = ['https://graph.microsoft.com/.default'];

/** Confidential-client credential for OBO. Certificate is the production choice. */
export type OboCredential =
  | { readonly kind: 'secret'; readonly clientSecret: string }
  | { readonly kind: 'certificate'; readonly thumbprint: string; readonly privateKey: string };

/** Exchanges a validated inbound assertion for a per-user Graph token. */
export interface OboClient {
  acquireGraphToken(userAssertion: string): Promise<string>;
}

/**
 * Loads the OBO credential from the environment, preferring a certificate.
 * Returns null when none is configured (OBO disabled — U4 auth still works, but
 * tool calls fail fast until the credential is provisioned in U9).
 */
export function loadOboCredential(env: NodeJS.ProcessEnv = process.env): OboCredential | null {
  const thumbprint = env.OUTLOOK_MCP_CONNECTOR_CERT_THUMBPRINT?.trim();
  const privateKey = env.OUTLOOK_MCP_CONNECTOR_CERT_KEY?.trim();
  if (thumbprint != null && thumbprint !== '' && privateKey != null && privateKey !== '') {
    return { kind: 'certificate', thumbprint, privateKey };
  }
  const secret = env.OUTLOOK_MCP_CONNECTOR_CLIENT_SECRET?.trim();
  if (secret != null && secret !== '') {
    return { kind: 'secret', clientSecret: secret };
  }
  return null;
}

/**
 * Builds the process-wide OBO client (one confidential-client app; per-user
 * isolation is per-assertion, not per-instance).
 */
export function createOboClient(config: RemoteAuthConfig, credential: OboCredential): OboClient {
  const authConfig: Configuration['auth'] = {
    clientId: config.apiClientId,
    authority: `https://login.microsoftonline.com/${config.tenantId}`,
    ...(credential.kind === 'secret'
      ? { clientSecret: credential.clientSecret }
      : {
          clientCertificate: {
            thumbprint: credential.thumbprint,
            privateKey: credential.privateKey,
          },
        }),
  };
  const cca = new ConfidentialClientApplication({ auth: authConfig });

  return {
    async acquireGraphToken(userAssertion: string): Promise<string> {
      let result;
      try {
        result = await cca.acquireTokenOnBehalfOf({
          oboAssertion: userAssertion,
          scopes: GRAPH_DEFAULT_SCOPE,
        });
      } catch (error) {
        throw mapOboError(error);
      }
      if (result?.accessToken == null || result.accessToken === '') {
        throw new GraphAuthRequiredError('session_expired');
      }
      return result.accessToken;
    },
  };
}

/**
 * Maps an OBO failure to a typed error. Consent/interaction failures need the
 * user (or admin) to act; transient failures are retriable. The raw AADSTS code
 * is surfaced in the message (never token material) so a broken flow is
 * diagnosable — MSAL otherwise swallows 4xx detail.
 */
export function mapOboError(error: unknown): OutlookMcpError {
  const code = extractAadsts(error);
  const message = error instanceof Error ? error.message : String(error);

  // Missing consent for a scope — needs admin consent (see provisioning runbook).
  if (code === 'AADSTS65001') {
    return new GraphError(
      `On-Behalf-Of failed: consent missing for a Graph scope (${code}). Grant admin ` +
        `consent for the connector API app — see docs/remote/provisioning.md.`,
    );
  }
  // Conditional Access / MFA / password change / revoked — the user must re-auth.
  if (code === 'AADSTS50076' || code === 'AADSTS50079' || /interaction_required|invalid_grant/i.test(message)) {
    return new GraphAuthRequiredError('session_expired');
  }
  // Expired confidential credential — total outage until the cert/secret rotates.
  if (code === 'AADSTS7000222' || code === 'AADSTS7000215') {
    return new GraphError(
      `On-Behalf-Of failed: the connector's confidential credential is invalid or ` +
        `expired (${code}). Rotate it per docs/remote/provisioning.md.`,
    );
  }
  return new GraphError(`On-Behalf-Of exchange failed${code != null ? ` (${code})` : ''}.`);
}

/** Pulls the first AADSTSxxxxx code from an MSAL error, if present. */
function extractAadsts(error: unknown): string | undefined {
  const text = error instanceof Error ? `${error.message}` : String(error);
  return /AADSTS\d+/.exec(text)?.[0];
}
