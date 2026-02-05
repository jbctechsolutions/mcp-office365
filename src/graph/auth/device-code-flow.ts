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
  type AuthenticationResult,
  type DeviceCodeRequest,
  type AccountInfo,
} from '@azure/msal-node';

import { loadGraphConfig, getAuthorityUrl, type GraphAuthConfig } from './config.js';
import { createTokenCachePlugin, hasTokenCache } from './token-cache.js';

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
  const msal = getMsalInstance();
  const config = getConfig();

  // Get accounts from cache
  const accounts = await msal.getTokenCache().getAllAccounts();

  // Use the first account (most common case)
  const account = accounts[0];
  if (account == null) {
    return null;
  }

  try {
    return await msal.acquireTokenSilent({
      account,
      scopes: [...config.scopes],
    });
  } catch {
    // Token refresh failed, will need interactive auth
    return null;
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
  deviceCodeCallback: DeviceCodeCallback = defaultDeviceCodeCallback
): Promise<string> {
  // Try silent first
  let result = await acquireTokenSilent();

  // If silent fails, do interactive
  if (result == null) {
    result = await acquireTokenInteractive(deviceCodeCallback);
  }

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
