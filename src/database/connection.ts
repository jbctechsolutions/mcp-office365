/**
 * SQLite database connection manager for Outlook database.
 *
 * Provides read-only access with retry logic for handling locked database.
 */

import Database, { type Database as DatabaseType } from 'better-sqlite3';
import { existsSync } from 'node:fs';
import {
  DatabaseNotFoundError,
  DatabaseLockedError,
  DatabaseError,
} from '../utils/errors.js';

/**
 * Options for database connection.
 */
export interface ConnectionOptions {
  /** Maximum number of retry attempts for locked database */
  readonly maxRetries: number;
  /** Initial delay in milliseconds between retries */
  readonly retryDelayMs: number;
  /** Busy timeout in milliseconds */
  readonly busyTimeoutMs: number;
}

/**
 * Default connection options.
 */
export const DEFAULT_CONNECTION_OPTIONS: ConnectionOptions = {
  maxRetries: 3,
  retryDelayMs: 500,
  busyTimeoutMs: 5000,
};

/**
 * Interface for database connection (for dependency injection).
 */
export interface IConnection {
  /**
   * Executes a function with a database connection.
   * Handles connection lifecycle and error handling.
   */
  execute<T>(fn: (db: DatabaseType) => T): T;

  /**
   * Closes any open connections.
   */
  close(): void;
}

/**
 * Manages read-only connections to the Outlook SQLite database.
 */
export class OutlookConnection implements IConnection {
  private db: DatabaseType | null = null;
  private readonly options: ConnectionOptions;

  constructor(
    private readonly dbPath: string,
    options: Partial<ConnectionOptions> = {}
  ) {
    this.options = { ...DEFAULT_CONNECTION_OPTIONS, ...options };
  }

  /**
   * Validates that the database file exists.
   */
  private validateDatabase(): void {
    if (!existsSync(this.dbPath)) {
      throw new DatabaseNotFoundError(this.dbPath);
    }
  }

  /**
   * Opens a connection to the database.
   */
  private open(): DatabaseType {
    if (this.db !== null) {
      return this.db;
    }

    this.validateDatabase();

    try {
      this.db = new Database(this.dbPath, {
        readonly: true,
        fileMustExist: true,
      });

      // Set pragmas for safe read-only access
      this.db.pragma(`busy_timeout = ${this.options.busyTimeoutMs}`);

      return this.db;
    } catch (error) {
      throw this.handleError(error);
    }
  }

  /**
   * Handles database errors, converting them to appropriate error types.
   */
  private handleError(error: unknown): OutlookError {
    if (error instanceof Error) {
      const message = error.message.toLowerCase();

      if (message.includes('database is locked')) {
        return new DatabaseLockedError();
      }

      if (
        message.includes('no such file') ||
        message.includes('unable to open')
      ) {
        return new DatabaseNotFoundError(this.dbPath);
      }

      return new DatabaseError(error.message, error);
    }

    return new DatabaseError('Unknown database error');
  }

  /**
   * Sleeps for the specified duration.
   */
  private sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  /**
   * Executes a function with a database connection, with retry logic.
   */
  execute<T>(fn: (db: DatabaseType) => T): T {
    let lastError: Error | null = null;

    for (let attempt = 0; attempt < this.options.maxRetries; attempt++) {
      try {
        const db = this.open();
        return fn(db);
      } catch (error) {
        lastError = error instanceof Error ? error : new Error(String(error));

        // Only retry on locked database
        if (!(error instanceof DatabaseLockedError)) {
          throw error;
        }

        // Don't sleep on last attempt
        if (attempt < this.options.maxRetries - 1) {
          // Synchronous sleep using busy-wait for simplicity
          // In production, you might want to make execute async
          const delay = this.options.retryDelayMs * Math.pow(2, attempt);
          const end = Date.now() + delay;
          while (Date.now() < end) {
            // Busy wait
          }
        }
      }
    }

    throw lastError ?? new DatabaseLockedError();
  }

  /**
   * Closes the database connection.
   */
  close(): void {
    if (this.db !== null) {
      this.db.close();
      this.db = null;
    }
  }
}

// Type alias for error handling
type OutlookError = DatabaseNotFoundError | DatabaseLockedError | DatabaseError;

/**
 * Creates a connection to the Outlook database.
 */
export function createConnection(
  dbPath: string,
  options?: Partial<ConnectionOptions>
): IConnection {
  return new OutlookConnection(dbPath, options);
}
