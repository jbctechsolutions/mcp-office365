/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Microsoft Authentication Library (MSAL) device code flow implementation.
 *
 * Uses the device code flow for authentication, which is ideal for CLI tools:
 * 1. User is shown a code and URL
 * 2. User visits URL in browser and enters code
 * 3. User authenticates with Microsoft
 * 4. Application receives tokens
 */

import {
  PublicClientApplication,
  InteractionRequiredAuthError,
  type AuthenticationResult,
  type DeviceCodeRequest,
  type AccountInfo,
} from '@azure/msal-node';

import { loadGraphConfig, getAuthorityUrl, type GraphAuthConfig } from './config.js';
import { createTokenCachePlugin, hasTokenCache } from './token-cache.js';
import { GraphAuthRequiredError } from '../../utils/errors.js';

/**
 * Singleton MSAL application instance.
 */
let msalInstance: PublicClientApplication | null = null;

/**
 * Cached configuration.
 */
let cachedConfig: GraphAuthConfig | null = null;

/**
 * Gets or creates the MSAL PublicClientApplication instance.
 */
function getMsalInstance(): PublicClientApplication {
  if (msalInstance == null) {
    cachedConfig = loadGraphConfig();

    msalInstance = new PublicClientApplication({
      auth: {
        clientId: cachedConfig.clientId,
        authority: getAuthorityUrl(cachedConfig),
      },
      cache: {
        cachePlugin: createTokenCachePlugin(),
      },
    });
  }

  return msalInstance;
}

/**
 * Gets the cached configuration.
 */
function getConfig(): GraphAuthConfig {
  if (cachedConfig == null) {
    getMsalInstance(); // This will initialize cachedConfig
  }
  return cachedConfig!;
}

/**
 * Callback type for device code messages.
 */
export type DeviceCodeCallback = (userCode: string, verificationUri: string, message: string) => void;

/**
 * Default callback that outputs to stderr.
 * stderr is used because stdout is reserved for MCP JSON-RPC communication.
 */
const defaultDeviceCodeCallback: DeviceCodeCallback = (userCode, verificationUri, _message) => {
  console.error('\n' + '='.repeat(60));
  console.error('Microsoft Graph API Authentication Required');
  console.error('='.repeat(60));
  console.error(`\nTo sign in, use a web browser to open the page:`);
  console.error(`  ${verificationUri}`);
  console.error(`\nAnd enter the code:`);
  console.error(`  ${userCode}`);
  console.error('\n' + '='.repeat(60) + '\n');
};

/**
 * Acquires a token using the device code flow.
 *
 * This will prompt the user to authenticate if needed.
 */
export async function acquireTokenInteractive(
  deviceCodeCallback: DeviceCodeCallback = defaultDeviceCodeCallback
): Promise<AuthenticationResult> {
  const msal = getMsalInstance();
  const config = getConfig();

  const request: DeviceCodeRequest = {
    scopes: [...config.scopes],
    deviceCodeCallback: (response) => {
      // MSAL swallows a 400 from the /devicecode endpoint (e.g. AADSTS50059 when a
      // single-tenant app is requested via `common`/`organizations`, or AADSTS700016
      // when the app is absent from the tenant) into an EMPTY response object and
      // invokes this callback anyway. Left alone it prints `undefined` for the code
      // and URL, then dies polling with a misleading `invalid_grant`. Fail fast with
      // an actionable diagnosis instead.
      if (!response.userCode || !response.verificationUri) {
        throw new Error(
          `Microsoft returned an empty device-code response for tenant "${config.tenantId}". ` +
            'This usually means the Azure app registration cannot complete a device-code sign-in ' +
            'against that tenant — e.g. a single-tenant app requested via "common"/"organizations", ' +
            'or an app that is not registered in the tenant. Set OUTLOOK_MCP_TENANT_ID to the tenant ' +
            'where the app is registered (or use an account that belongs to it).'
        );
      }
      deviceCodeCallback(response.userCode, response.verificationUri, response.message);
    },
  };

  const result = await msal.acquireTokenByDeviceCode(request);
  if (result == null) {
    throw new Error('Device code authentication failed');
  }
  return result;
}

/**
 * Acquires a token silently using cached credentials.
 *
 * Returns null if no cached token is available or if refresh fails.
 */
export async function acquireTokenSilent(): Promise<AuthenticationResult | null> {
  try {
    const outcome = await acquireTokenSilentDetailed();
    return outcome.status === 'ok' ? outcome.result : null;
  } catch {
    // Preserve this wrapper's "null on any failure" contract even for the
    // transient errors acquireTokenSilentDetailed now rethrows.
    return null;
  }
}

/**
 * Silent acquisition that distinguishes *why* it failed: no cached account
 * (never authenticated) vs. a cached account whose refresh token is expired or
 * revoked (session expired mid-run — the `invalid_grant` case). The caller uses
 * this to avoid falling into an unwatchable re-interactive device-code flow for
 * an expired session, surfacing a typed AUTH_EXPIRED instead (U9).
 */
type SilentOutcome =
  | { status: 'ok'; result: AuthenticationResult }
  | { status: 'no_account' }
  | { status: 'refresh_failed'; cause: unknown };

async function acquireTokenSilentDetailed(): Promise<SilentOutcome> {
  const msal = getMsalInstance();
  const config = getConfig();

  const accounts = await msal.getTokenCache().getAllAccounts();
  const account = accounts[0];
  if (account == null) {
    return { status: 'no_account' };
  }

  try {
    const result = await msal.acquireTokenSilent({
      account,
      scopes: [...config.scopes],
    });
    return { status: 'ok', result };
  } catch (cause) {
    // Only an interaction-required failure (expired/revoked refresh token,
    // invalid_grant) means a re-auth is needed. Transient failures (network,
    // throttling, server errors) are NOT auth problems — rethrow so the caller
    // surfaces them as retriable errors rather than a misleading AUTH_EXPIRED or
    // an unnecessary device-code prompt.
    if (cause instanceof InteractionRequiredAuthError) {
      return { status: 'refresh_failed', cause };
    }
    throw cause;
  }
}

/**
 * Gets an access token, prompting for login if needed.
 *
 * This is the main entry point for getting a token:
 * 1. First tries to get a cached/refreshed token silently
 * 2. If that fails, prompts for device code authentication
 */
export async function getAccessToken(
  deviceCodeCallback: DeviceCodeCallback = defaultDeviceCodeCallback,
  options: { interactiveOnExpired?: boolean } = {}
): Promise<string> {
  const silent = await acquireTokenSilentDetailed();
  if (silent.status === 'ok') {
    return silent.result.accessToken;
  }

  // A cached session that failed to refresh (expired/revoked, `invalid_grant`):
  // in the MCP server (default) we must NOT fall into a fresh interactive
  // device-code flow — it prints to an unwatched stderr and the call hangs until
  // the code expires — so surface a typed AUTH_EXPIRED with the re-auth hint (U9).
  // The CLI `auth` command sets interactiveOnExpired so a user at a terminal can
  // actually recover by re-authenticating (otherwise the hint would be circular).
  if (silent.status === 'refresh_failed' && options.interactiveOnExpired !== true) {
    throw new GraphAuthRequiredError('session_expired');
  }

  // No cached account (genuine first-time), or an interactive-capable caller on
  // an expired session: the device code is appropriate.
  const result = await acquireTokenInteractive(deviceCodeCallback);
  return result.accessToken;
}

/**
 * Checks if the user is authenticated (has cached tokens).
 */
export async function isAuthenticated(): Promise<boolean> {
  if (!hasTokenCache()) {
    return false;
  }

  try {
    const msal = getMsalInstance();
    const accounts = await msal.getTokenCache().getAllAccounts();
    return accounts.length > 0;
  } catch {
    return false;
  }
}

/**
 * Gets the currently authenticated account info.
 */
export async function getAccount(): Promise<AccountInfo | null> {
  try {
    const msal = getMsalInstance();
    const accounts = await msal.getTokenCache().getAllAccounts();
    const account = accounts[0];
    return account != null ? account : null;
  } catch {
    return null;
  }
}

/**
 * Signs out by clearing the token cache.
 */
export async function signOut(): Promise<void> {
  try {
    const msal = getMsalInstance();
    const accounts = await msal.getTokenCache().getAllAccounts();

    for (const account of accounts) {
      await msal.getTokenCache().removeAccount(account);
    }
  } catch {
    // Ignore errors during sign out
  }
}

/**
 * Resets the MSAL instance (mainly for testing).
 */
export function resetMsalInstance(): void {
  msalInstance = null;
  cachedConfig = null;
}
