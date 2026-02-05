/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { createTestDatabase } from '../../fixtures/database.js';
import {
  OutlookConnection,
  createConnection,
  DEFAULT_CONNECTION_OPTIONS,
} from '../../../src/database/connection.js';
import {
  DatabaseNotFoundError,
} from '../../../src/utils/errors.js';

describe('OutlookConnection', () => {
  let testDb: { path: string; cleanup: () => void };

  beforeEach(() => {
    testDb = createTestDatabase();
  });

  afterEach(() => {
    testDb.cleanup();
  });

  describe('constructor', () => {
    it('creates connection with default options', () => {
      const conn = new OutlookConnection(testDb.path);
      expect(conn).toBeInstanceOf(OutlookConnection);
      conn.close();
    });

    it('creates connection with custom options', () => {
      const conn = new OutlookConnection(testDb.path, {
        maxRetries: 5,
        retryDelayMs: 1000,
      });
      expect(conn).toBeInstanceOf(OutlookConnection);
      conn.close();
    });
  });

  describe('execute', () => {
    it('executes a query successfully', () => {
      const conn = createConnection(testDb.path);

      const result = conn.execute((db) => {
        const stmt = db.prepare('SELECT COUNT(*) as count FROM Folders');
        return (stmt.get() as { count: number }).count;
      });

      expect(result).toBeGreaterThan(0);
      conn.close();
    });

    it('can execute multiple queries', () => {
      const conn = createConnection(testDb.path);

      const count1 = conn.execute((db) => {
        const stmt = db.prepare('SELECT COUNT(*) as count FROM Folders');
        return (stmt.get() as { count: number }).count;
      });

      const count2 = conn.execute((db) => {
        const stmt = db.prepare('SELECT COUNT(*) as count FROM Mail');
        return (stmt.get() as { count: number }).count;
      });

      expect(count1).toBeGreaterThan(0);
      expect(count2).toBeGreaterThan(0);
      conn.close();
    });

    it('throws DatabaseNotFoundError for non-existent database', () => {
      const conn = createConnection('/nonexistent/path/to/db.sqlite');

      expect(() => {
        conn.execute((db) => db.prepare('SELECT 1').get());
      }).toThrow(DatabaseNotFoundError);
    });
  });

  describe('close', () => {
    it('closes the connection without error', () => {
      const conn = createConnection(testDb.path);

      // Execute something to open the connection
      conn.execute((db) => db.prepare('SELECT 1').get());

      // Close should not throw
      expect(() => conn.close()).not.toThrow();
    });

    it('can be called multiple times safely', () => {
      const conn = createConnection(testDb.path);

      conn.execute((db) => db.prepare('SELECT 1').get());
      conn.close();
      conn.close(); // Second close should not throw

      expect(true).toBe(true);
    });
  });

  describe('DEFAULT_CONNECTION_OPTIONS', () => {
    it('has reasonable defaults', () => {
      expect(DEFAULT_CONNECTION_OPTIONS.maxRetries).toBe(3);
      expect(DEFAULT_CONNECTION_OPTIONS.retryDelayMs).toBe(500);
      expect(DEFAULT_CONNECTION_OPTIONS.busyTimeoutMs).toBe(5000);
    });
  });
});
