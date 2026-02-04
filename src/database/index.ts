/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

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
  type IWriteableRepository,
  type IMailboxRepository,
  type MaybePromise,
  type FolderRow,
  type EmailRow,
  type EventRow,
  type ContactRow,
  type TaskRow,
  type NoteRow,
  OutlookRepository,
  createRepository,
} from './repository.js';
