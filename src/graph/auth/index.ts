/**
 * Microsoft Graph authentication module.
 *
 * Exports authentication utilities for the device code flow.
 */

export { loadGraphConfig, getAuthorityUrl, GRAPH_SCOPES, type GraphAuthConfig } from './config.js';

export {
  createTokenCachePlugin,
  hasTokenCache,
  clearTokenCache,
  getTokenCacheDir,
  getTokenCacheFile,
} from './token-cache.js';

export {
  acquireTokenInteractive,
  acquireTokenSilent,
  getAccessToken,
  isAuthenticated,
  getAccount,
  signOut,
  resetMsalInstance,
  type DeviceCodeCallback,
} from './device-code-flow.js';
