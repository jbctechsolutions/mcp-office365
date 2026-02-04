/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { writeFileSync, mkdirSync, rmSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { randomBytes } from 'node:crypto';
import {
  parseOlk15File,
  getDefaultDataPath,
  Olk15EmailContentReader,
  Olk15EventContentReader,
  Olk15ContactContentReader,
  Olk15TaskContentReader,
  Olk15NoteContentReader,
  createContentReaders,
} from '../../../src/parsers/olk15.js';

describe('OLK15 Parser', () => {
  let testDir: string;

  beforeEach(() => {
    testDir = join(tmpdir(), `olk15-test-${randomBytes(8).toString('hex')}`);
    mkdirSync(testDir, { recursive: true });
  });

  afterEach(() => {
    if (existsSync(testDir)) {
      rmSync(testDir, { recursive: true, force: true });
    }
  });

  // ---------------------------------------------------------------------------
  // parseOlk15File
  // ---------------------------------------------------------------------------

  describe('parseOlk15File', () => {
    it('returns error for non-existent file', () => {
      const result = parseOlk15File('/nonexistent/path/file.olk15Message');
      expect(result.success).toBe(false);
      expect(result.error).toBe('File not found');
    });

    it('returns error for empty file', () => {
      const filePath = join(testDir, 'empty.olk15Message');
      writeFileSync(filePath, Buffer.alloc(0));

      const result = parseOlk15File(filePath);
      expect(result.success).toBe(false);
      expect(result.error).toBe('Empty file');
    });

    it('extracts HTML content from file', () => {
      const html = '<html><body><p>Hello World</p></body></html>';
      const filePath = join(testDir, 'test.olk15Message');
      writeFileSync(filePath, Buffer.from(html));

      const result = parseOlk15File(filePath);
      expect(result.success).toBe(true);
      expect(result.html).toContain('<html>');
      expect(result.html).toContain('Hello World');
    });

    it('extracts body content without full html wrapper', () => {
      const html = 'PREFIX<body><p>Content here</p></body>SUFFIX';
      const filePath = join(testDir, 'test.olk15Message');
      writeFileSync(filePath, Buffer.from(html));

      const result = parseOlk15File(filePath);
      expect(result.success).toBe(true);
      expect(result.html).toContain('<body>');
    });

    it('extracts UTF-8 text from binary', () => {
      // Create binary data with embedded UTF-8 text
      const prefix = Buffer.from([0x00, 0x01, 0x02, 0x03, 0x04]);
      const text = Buffer.from('This is a test email body with some content that should be extracted.');
      const suffix = Buffer.from([0x00, 0x00, 0x00]);
      const buffer = Buffer.concat([prefix, text, suffix]);

      const filePath = join(testDir, 'utf8.olk15Message');
      writeFileSync(filePath, buffer);

      const result = parseOlk15File(filePath);
      expect(result.success).toBe(true);
      expect(result.text).toContain('test email body');
    });

    it('extracts UTF-16LE text from binary', () => {
      // Create UTF-16LE encoded text with proper 2-byte alignment
      const text = 'This is UTF-16 encoded text for testing purposes.';
      const utf16Buffer = Buffer.alloc(text.length * 2);
      for (let i = 0; i < text.length; i++) {
        utf16Buffer.writeUInt16LE(text.charCodeAt(i), i * 2);
      }

      // Use even-length prefix to maintain 2-byte alignment
      const prefix = Buffer.from([0x00, 0x00, 0x01, 0x00]);
      const suffix = Buffer.from([0x00, 0x00, 0x00, 0x00]);
      const buffer = Buffer.concat([prefix, utf16Buffer, suffix]);

      const filePath = join(testDir, 'utf16.olk15Message');
      writeFileSync(filePath, buffer);

      const result = parseOlk15File(filePath);
      expect(result.success).toBe(true);
      expect(result.text).toContain('UTF-16');
    });

    it('returns no content for binary-only file', () => {
      // Create purely binary data with no recognizable text
      const buffer = Buffer.from([
        0xff, 0xfe, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08,
      ]);

      const filePath = join(testDir, 'binary.olk15Message');
      writeFileSync(filePath, buffer);

      const result = parseOlk15File(filePath);
      expect(result.success).toBe(false);
      expect(result.error).toBe('No text content found');
    });
  });

  // ---------------------------------------------------------------------------
  // getDefaultDataPath
  // ---------------------------------------------------------------------------

  describe('getDefaultDataPath', () => {
    it('returns path with default profile', () => {
      const path = getDefaultDataPath();
      expect(path).toContain('Main Profile');
      expect(path).toContain('Outlook 15 Profiles');
      expect(path).toContain('Data');
    });

    it('returns path with custom profile', () => {
      const path = getDefaultDataPath('Work Profile');
      expect(path).toContain('Work Profile');
      expect(path).not.toContain('Main Profile');
    });
  });

  // ---------------------------------------------------------------------------
  // Olk15EmailContentReader
  // ---------------------------------------------------------------------------

  describe('Olk15EmailContentReader', () => {
    it('returns null for null path', () => {
      const reader = new Olk15EmailContentReader(testDir);
      expect(reader.readEmailBody(null)).toBeNull();
    });

    it('returns null for non-existent file', () => {
      const reader = new Olk15EmailContentReader(testDir);
      expect(reader.readEmailBody('nonexistent.olk15Message')).toBeNull();
    });

    it('reads email body from file', () => {
      const html = '<html><body>Email content here</body></html>';
      const filePath = 'test.olk15Message';
      writeFileSync(join(testDir, filePath), Buffer.from(html));

      const reader = new Olk15EmailContentReader(testDir);
      const body = reader.readEmailBody(filePath);

      expect(body).toContain('Email content');
    });
  });

  // ---------------------------------------------------------------------------
  // Olk15EventContentReader
  // ---------------------------------------------------------------------------

  describe('Olk15EventContentReader', () => {
    it('returns null for null path', () => {
      const reader = new Olk15EventContentReader(testDir);
      expect(reader.readEventDetails(null)).toBeNull();
    });

    it('extracts event details from file', () => {
      const content = 'Subject: Team Meeting\nLocation: Room 101\nOrganizer: boss@example.com';
      const filePath = 'event.olk15Event';
      writeFileSync(join(testDir, filePath), Buffer.from(content));

      const reader = new Olk15EventContentReader(testDir);
      const details = reader.readEventDetails(filePath);

      expect(details).not.toBeNull();
      expect(details?.title).toBe('Team Meeting');
      expect(details?.location).toBe('Room 101');
      expect(details?.organizer).toBe('boss@example.com');
    });

    it('uses first line as title when no Subject field', () => {
      const content = 'Weekly Standup\nDiscussion of sprint progress';
      const filePath = 'event2.olk15Event';
      writeFileSync(join(testDir, filePath), Buffer.from(content));

      const reader = new Olk15EventContentReader(testDir);
      const details = reader.readEventDetails(filePath);

      expect(details?.title).toBe('Weekly Standup');
    });
  });

  // ---------------------------------------------------------------------------
  // Olk15ContactContentReader
  // ---------------------------------------------------------------------------

  describe('Olk15ContactContentReader', () => {
    it('returns null for null path', () => {
      const reader = new Olk15ContactContentReader(testDir);
      expect(reader.readContactDetails(null)).toBeNull();
    });

    it('reads contact data from file', () => {
      const content = 'John Doe\njohn@example.com\n555-1234';
      const filePath = 'contact.olk15Contact';
      writeFileSync(join(testDir, filePath), Buffer.from(content));

      const reader = new Olk15ContactContentReader(testDir);
      const details = reader.readContactDetails(filePath);

      expect(details).not.toBeNull();
      expect(details?.notes).toContain('John Doe');
    });
  });

  // ---------------------------------------------------------------------------
  // Olk15TaskContentReader
  // ---------------------------------------------------------------------------

  describe('Olk15TaskContentReader', () => {
    it('returns null for null path', () => {
      const reader = new Olk15TaskContentReader(testDir);
      expect(reader.readTaskDetails(null)).toBeNull();
    });

    it('reads task content from file', () => {
      const content = 'Complete the quarterly report\nDue by end of week';
      const filePath = 'task.olk15Task';
      writeFileSync(join(testDir, filePath), Buffer.from(content));

      const reader = new Olk15TaskContentReader(testDir);
      const details = reader.readTaskDetails(filePath);

      expect(details).not.toBeNull();
      expect(details?.body).toContain('quarterly report');
    });
  });

  // ---------------------------------------------------------------------------
  // Olk15NoteContentReader
  // ---------------------------------------------------------------------------

  describe('Olk15NoteContentReader', () => {
    it('returns null for null path', () => {
      const reader = new Olk15NoteContentReader(testDir);
      expect(reader.readNoteDetails(null)).toBeNull();
    });

    it('reads note content from file', () => {
      const content = 'Meeting Notes\nDiscussed project timeline\nAction items assigned';
      const filePath = 'note.olk15Note';
      writeFileSync(join(testDir, filePath), Buffer.from(content));

      const reader = new Olk15NoteContentReader(testDir);
      const details = reader.readNoteDetails(filePath);

      expect(details).not.toBeNull();
      expect(details?.title).toBe('Meeting Notes');
      expect(details?.body).toContain('project timeline');
      expect(details?.preview?.length).toBeLessThanOrEqual(200);
    });
  });

  // ---------------------------------------------------------------------------
  // createContentReaders
  // ---------------------------------------------------------------------------

  describe('createContentReaders', () => {
    it('creates all content readers', () => {
      const readers = createContentReaders(testDir);

      expect(readers.email).toBeInstanceOf(Olk15EmailContentReader);
      expect(readers.event).toBeInstanceOf(Olk15EventContentReader);
      expect(readers.contact).toBeInstanceOf(Olk15ContactContentReader);
      expect(readers.task).toBeInstanceOf(Olk15TaskContentReader);
      expect(readers.note).toBeInstanceOf(Olk15NoteContentReader);
    });
  });
});
