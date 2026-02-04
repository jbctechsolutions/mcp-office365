/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { createTestDatabase, SAMPLE_COUNTS } from '../../fixtures/database.js';
import { createConnection, type IConnection } from '../../../src/database/connection.js';
import { createRepository, type IRepository } from '../../../src/database/repository.js';
import {
  MailTools,
  createMailTools,
  ListFoldersInput,
  ListEmailsInput,
  SearchEmailsInput,
  GetEmailInput,
  GetUnreadCountInput,
  type IContentReader,
} from '../../../src/tools/mail.js';

describe('MailTools', () => {
  let testDb: { path: string; cleanup: () => void };
  let connection: IConnection;
  let repository: IRepository;
  let mailTools: MailTools;

  beforeEach(() => {
    testDb = createTestDatabase();
    connection = createConnection(testDb.path);
    repository = createRepository(connection);
    mailTools = createMailTools(repository);
  });

  afterEach(() => {
    connection.close();
    testDb.cleanup();
  });

  // ---------------------------------------------------------------------------
  // Input Validation
  // ---------------------------------------------------------------------------

  describe('input validation', () => {
    it('validates ListFoldersInput', () => {
      expect(() => ListFoldersInput.parse({})).not.toThrow();
      expect(() => ListFoldersInput.parse({ extra: 'field' })).toThrow();
    });

    it('validates ListEmailsInput', () => {
      const valid = { folder_id: 1 };
      const parsed = ListEmailsInput.parse(valid);
      expect(parsed.folder_id).toBe(1);
      expect(parsed.limit).toBe(50); // default
      expect(parsed.offset).toBe(0); // default
      expect(parsed.unread_only).toBe(false); // default
    });

    it('validates ListEmailsInput with all options', () => {
      const input = {
        folder_id: 2,
        limit: 25,
        offset: 10,
        unread_only: true,
      };
      const parsed = ListEmailsInput.parse(input);
      expect(parsed).toEqual(input);
    });

    it('rejects invalid ListEmailsInput', () => {
      expect(() => ListEmailsInput.parse({ folder_id: 'abc' })).toThrow();
      expect(() => ListEmailsInput.parse({ folder_id: -1 })).toThrow();
      expect(() => ListEmailsInput.parse({ folder_id: 1, limit: 0 })).toThrow();
      expect(() => ListEmailsInput.parse({ folder_id: 1, limit: 101 })).toThrow();
    });

    it('validates SearchEmailsInput', () => {
      const parsed = SearchEmailsInput.parse({ query: 'test' });
      expect(parsed.query).toBe('test');
      expect(parsed.limit).toBe(50);
      expect(parsed.folder_id).toBeUndefined();
    });

    it('rejects empty search query', () => {
      expect(() => SearchEmailsInput.parse({ query: '' })).toThrow();
    });

    it('validates GetEmailInput', () => {
      const parsed = GetEmailInput.parse({ email_id: 1 });
      expect(parsed.email_id).toBe(1);
      expect(parsed.include_body).toBe(true);
      expect(parsed.strip_html).toBe(true);
    });

    it('validates GetUnreadCountInput', () => {
      const parsed = GetUnreadCountInput.parse({});
      expect(parsed.folder_id).toBeUndefined();

      const withFolder = GetUnreadCountInput.parse({ folder_id: 1 });
      expect(withFolder.folder_id).toBe(1);
    });
  });

  // ---------------------------------------------------------------------------
  // listFolders
  // ---------------------------------------------------------------------------

  describe('listFolders', () => {
    it('returns all mail folders', () => {
      const folders = mailTools.listFolders({});
      expect(folders.length).toBe(SAMPLE_COUNTS.mailFolders);
    });

    it('returns folders with correct structure', () => {
      const folders = mailTools.listFolders({});
      const inbox = folders.find((f) => f.name === 'Inbox');

      expect(inbox).toBeDefined();
      expect(inbox).toHaveProperty('id');
      expect(inbox).toHaveProperty('name');
      expect(inbox).toHaveProperty('parentId');
      expect(inbox).toHaveProperty('specialType');
      expect(inbox).toHaveProperty('folderType');
      expect(inbox).toHaveProperty('accountId');
      expect(inbox).toHaveProperty('messageCount');
      expect(inbox).toHaveProperty('unreadCount');
    });

    it('includes message and unread counts', () => {
      const folders = mailTools.listFolders({});
      const inbox = folders.find((f) => f.name === 'Inbox');

      expect(inbox?.messageCount).toBe(SAMPLE_COUNTS.inboxEmails);
      expect(inbox?.unreadCount).toBe(SAMPLE_COUNTS.unreadEmails);
    });
  });

  // ---------------------------------------------------------------------------
  // listEmails
  // ---------------------------------------------------------------------------

  describe('listEmails', () => {
    it('returns emails in folder', () => {
      const folders = mailTools.listFolders({});
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const emails = mailTools.listEmails({
          folder_id: inbox.id,
          limit: 50,
          offset: 0,
          unread_only: false,
        });
        expect(emails.length).toBe(SAMPLE_COUNTS.inboxEmails);
      }
    });

    it('returns emails with correct structure', () => {
      const folders = mailTools.listFolders({});
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const emails = mailTools.listEmails({
          folder_id: inbox.id,
          limit: 1,
          offset: 0,
          unread_only: false,
        });
        const email = emails[0];

        expect(email).toHaveProperty('id');
        expect(email).toHaveProperty('folderId');
        expect(email).toHaveProperty('subject');
        expect(email).toHaveProperty('sender');
        expect(email).toHaveProperty('isRead');
        expect(email).toHaveProperty('timeReceived');
        expect(email).toHaveProperty('hasAttachment');
        expect(typeof email?.isRead).toBe('boolean');
        expect(typeof email?.hasAttachment).toBe('boolean');
      }
    });

    it('converts timestamps to ISO format', () => {
      const folders = mailTools.listFolders({});
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const emails = mailTools.listEmails({
          folder_id: inbox.id,
          limit: 1,
          offset: 0,
          unread_only: false,
        });
        const email = emails[0];

        // Should be ISO string format
        if (email?.timeReceived) {
          expect(email.timeReceived).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/);
        }
      }
    });

    it('respects limit parameter', () => {
      const folders = mailTools.listFolders({});
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const emails = mailTools.listEmails({
          folder_id: inbox.id,
          limit: 1,
          offset: 0,
          unread_only: false,
        });
        expect(emails.length).toBe(1);
      }
    });

    it('respects offset parameter', () => {
      const folders = mailTools.listFolders({});
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const allEmails = mailTools.listEmails({
          folder_id: inbox.id,
          limit: 50,
          offset: 0,
          unread_only: false,
        });
        const offsetEmails = mailTools.listEmails({
          folder_id: inbox.id,
          limit: 50,
          offset: 1,
          unread_only: false,
        });
        expect(offsetEmails.length).toBe(allEmails.length - 1);
      }
    });

    it('filters unread emails when unread_only is true', () => {
      const folders = mailTools.listFolders({});
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const emails = mailTools.listEmails({
          folder_id: inbox.id,
          limit: 50,
          offset: 0,
          unread_only: true,
        });
        expect(emails.length).toBe(SAMPLE_COUNTS.unreadEmails);
        expect(emails.every((e) => e.isRead === false)).toBe(true);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // searchEmails
  // ---------------------------------------------------------------------------

  describe('searchEmails', () => {
    it('finds emails by subject', () => {
      const emails = mailTools.searchEmails({
        query: 'Meeting',
        limit: 50,
      });
      expect(emails.length).toBeGreaterThan(0);
      expect(emails.some((e) => e.subject?.includes('Meeting'))).toBe(true);
    });

    it('finds emails by sender', () => {
      const emails = mailTools.searchEmails({
        query: 'John',
        limit: 50,
      });
      expect(emails.length).toBeGreaterThan(0);
    });

    it('limits results to specified folder', () => {
      const folders = mailTools.listFolders({});
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const emails = mailTools.searchEmails({
          query: 'Meeting',
          folder_id: inbox.id,
          limit: 50,
        });
        expect(emails.every((e) => e.folderId === inbox.id)).toBe(true);
      }
    });

    it('returns empty array for no matches', () => {
      const emails = mailTools.searchEmails({
        query: 'xyznonexistent',
        limit: 50,
      });
      expect(emails.length).toBe(0);
    });
  });

  // ---------------------------------------------------------------------------
  // getEmail
  // ---------------------------------------------------------------------------

  describe('getEmail', () => {
    it('returns email by ID', () => {
      const folders = mailTools.listFolders({});
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const emails = mailTools.listEmails({
          folder_id: inbox.id,
          limit: 1,
          offset: 0,
          unread_only: false,
        });
        const firstEmail = emails[0];

        if (firstEmail) {
          const email = mailTools.getEmail({
            email_id: firstEmail.id,
            include_body: false,
            strip_html: true,
          });
          expect(email).not.toBeNull();
          expect(email?.id).toBe(firstEmail.id);
        }
      }
    });

    it('returns null for non-existent ID', () => {
      const email = mailTools.getEmail({
        email_id: 99999,
        include_body: false,
        strip_html: true,
      });
      expect(email).toBeNull();
    });

    it('includes additional fields in full email', () => {
      const folders = mailTools.listFolders({});
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const emails = mailTools.listEmails({
          folder_id: inbox.id,
          limit: 1,
          offset: 0,
          unread_only: false,
        });
        const firstEmail = emails[0];

        if (firstEmail) {
          const email = mailTools.getEmail({
            email_id: firstEmail.id,
            include_body: true,
            strip_html: true,
          });
          expect(email).toHaveProperty('recipients');
          expect(email).toHaveProperty('displayTo');
          expect(email).toHaveProperty('size');
          expect(email).toHaveProperty('body');
        }
      }
    });
  });

  // ---------------------------------------------------------------------------
  // getUnreadCount
  // ---------------------------------------------------------------------------

  describe('getUnreadCount', () => {
    it('returns total unread count', () => {
      const result = mailTools.getUnreadCount({});
      expect(result.count).toBe(SAMPLE_COUNTS.unreadEmails);
    });

    it('returns unread count for specific folder', () => {
      const folders = mailTools.listFolders({});
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const result = mailTools.getUnreadCount({ folder_id: inbox.id });
        expect(result.count).toBe(SAMPLE_COUNTS.unreadEmails);
      }
    });

    it('returns 0 for folder with no unread emails', () => {
      const folders = mailTools.listFolders({});
      const sent = folders.find((f) => f.name === 'Sent Items');

      if (sent) {
        const result = mailTools.getUnreadCount({ folder_id: sent.id });
        expect(result.count).toBe(0);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Content Reader Integration
  // ---------------------------------------------------------------------------

  describe('content reader integration', () => {
    it('uses custom content reader for body', () => {
      const mockContentReader: IContentReader = {
        readEmailBody: (path) => {
          if (path) {
            return '<html><body>Test email body</body></html>';
          }
          return null;
        },
      };

      const toolsWithReader = createMailTools(repository, mockContentReader);
      const folders = toolsWithReader.listFolders({});
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const emails = toolsWithReader.listEmails({
          folder_id: inbox.id,
          limit: 1,
          offset: 0,
          unread_only: false,
        });
        const firstEmail = emails[0];

        if (firstEmail) {
          const email = toolsWithReader.getEmail({
            email_id: firstEmail.id,
            include_body: true,
            strip_html: true,
          });
          expect(email?.body).toBe('Test email body');
        }
      }
    });

    it('preserves HTML when strip_html is false', () => {
      const mockContentReader: IContentReader = {
        readEmailBody: () => '<html><body><p>Test</p></body></html>',
      };

      const toolsWithReader = createMailTools(repository, mockContentReader);
      const folders = toolsWithReader.listFolders({});
      const inbox = folders.find((f) => f.name === 'Inbox');

      if (inbox) {
        const emails = toolsWithReader.listEmails({
          folder_id: inbox.id,
          limit: 1,
          offset: 0,
          unread_only: false,
        });
        const firstEmail = emails[0];

        if (firstEmail) {
          const email = toolsWithReader.getEmail({
            email_id: firstEmail.id,
            include_body: true,
            strip_html: false,
          });
          expect(email?.body).toContain('<html>');
          expect(email?.htmlBody).toContain('<html>');
        }
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Factory Function
  // ---------------------------------------------------------------------------

  describe('createMailTools', () => {
    it('creates a MailTools instance', () => {
      const tools = createMailTools(repository);
      expect(tools).toBeInstanceOf(MailTools);
    });
  });
});
