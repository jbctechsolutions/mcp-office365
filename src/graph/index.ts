/**
 * Microsoft Graph API module.
 *
 * Provides an alternative backend for the Outlook MCP server that uses
 * Microsoft Graph API instead of AppleScript. This is required for
 * "new Outlook" for Mac which doesn't expose data via AppleScript.
 */

// Auth exports
export {
  loadGraphConfig,
  getAuthorityUrl,
  GRAPH_SCOPES,
  type GraphAuthConfig,
  createTokenCachePlugin,
  hasTokenCache,
  clearTokenCache,
  getTokenCacheDir,
  getTokenCacheFile,
  acquireTokenInteractive,
  acquireTokenSilent,
  getAccessToken,
  isAuthenticated,
  getAccount,
  signOut,
  resetMsalInstance,
  type DeviceCodeCallback,
} from './auth/index.js';

// Client exports
export {
  GraphClient,
  createGraphClient,
  ResponseCache,
  CacheTTL,
  createCacheKey,
  invalidateByPrefix,
} from './client/index.js';

// Mapper exports
export {
  mapMailFolderToRow,
  mapCalendarToFolderRow,
  mapTaskListToFolderRow,
  mapMessageToEmailRow,
  mapEventToEventRow,
  mapContactToContactRow,
  mapTaskToTaskRow,
  hashStringToNumber,
  isoToTimestamp,
  dateTimeTimeZoneToTimestamp,
  type TodoTaskWithList,
} from './mappers/index.js';

// Repository exports
export { GraphRepository, createGraphRepository } from './repository.js';

// Content reader exports
export {
  GraphEmailContentReader,
  GraphEventContentReader,
  GraphContactContentReader,
  GraphTaskContentReader,
  GraphNoteContentReader,
  createGraphContentReaders,
  createGraphContentReadersWithClient,
  type GraphContentReaders,
  GRAPH_EMAIL_PATH_PREFIX,
  GRAPH_EVENT_PATH_PREFIX,
  GRAPH_CONTACT_PATH_PREFIX,
  GRAPH_TASK_PATH_PREFIX,
} from './content-readers.js';
