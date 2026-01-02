/**
 * Database module for Outlook MCP server.
 */

export {
  type IConnection,
  type ConnectionOptions,
  DEFAULT_CONNECTION_OPTIONS,
  OutlookConnection,
  createConnection,
} from './connection.js';

export * as queries from './queries.js';

export {
  type IRepository,
  type FolderRow,
  type EmailRow,
  type EventRow,
  type ContactRow,
  type TaskRow,
  type NoteRow,
  OutlookRepository,
  createRepository,
} from './repository.js';
