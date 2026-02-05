/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * OLK15 binary file parser.
 *
 * Parses Outlook for Mac's proprietary binary content files.
 * These files contain email bodies, event details, contact information, etc.
 *
 * File types:
 * - .olk15MsgSource / .olk15Message - Email content
 * - .olk15Event - Calendar event details
 * - .olk15Contact - Contact information
 * - .olk15Note - Note content
 * - .olk15Task - Task details
 *
 * The format is proprietary and not documented. This parser uses heuristics
 * to extract text content and provides graceful fallbacks when parsing fails.
 */

import { readFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import type { IContentReader } from '../tools/mail.js';
import type { IEventContentReader, EventDetails } from '../tools/calendar.js';
import type { IContactContentReader, ContactDetails } from '../tools/contacts.js';
import type { ITaskContentReader, TaskDetails } from '../tools/tasks.js';
import type { INoteContentReader, NoteDetails } from '../tools/notes.js';

// =============================================================================
// Configuration
// =============================================================================

/**
 * Default profile data directory path.
 */
export function getDefaultDataPath(profileName: string = 'Main Profile'): string {
  const home = process.env.HOME ?? '';
  return join(
    home,
    'Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles',
    profileName,
    'Data'
  );
}

// =============================================================================
// Binary Parsing Utilities
// =============================================================================

/**
 * Finds UTF-16LE encoded text in binary data.
 * OLK15 files often store text as UTF-16LE.
 */
function findUtf16Text(buffer: Buffer, minLength: number = 10): string[] {
  const results: string[] = [];
  let start = -1;
  let chars: string[] = [];

  for (let i = 0; i < buffer.length - 1; i += 2) {
    const char = buffer.readUInt16LE(i);

    // Check if it's a printable character or common whitespace
    if ((char >= 0x20 && char < 0x7f) || char === 0x0a || char === 0x0d || char === 0x09) {
      if (start === -1) start = i;
      chars.push(String.fromCharCode(char));
    } else if (char === 0) {
      // Null character - end of string
      if (chars.length >= minLength) {
        results.push(chars.join(''));
      }
      chars = [];
      start = -1;
    } else {
      // Non-printable, non-null - might be end of text section
      if (chars.length >= minLength) {
        results.push(chars.join(''));
      }
      chars = [];
      start = -1;
    }
  }

  // Don't forget the last string
  if (chars.length >= minLength) {
    results.push(chars.join(''));
  }

  return results;
}

/**
 * Finds UTF-8 encoded text in binary data.
 */
function findUtf8Text(buffer: Buffer, minLength: number = 10): string[] {
  const results: string[] = [];
  let start = -1;
  let chars: number[] = [];

  for (let i = 0; i < buffer.length; i++) {
    const byte = buffer[i];

    if (byte !== undefined) {
      // Check if it's a printable ASCII or common whitespace
      if ((byte >= 0x20 && byte < 0x7f) || byte === 0x0a || byte === 0x0d || byte === 0x09) {
        if (start === -1) start = i;
        chars.push(byte);
      } else if (byte === 0) {
        // Null terminator
        if (chars.length >= minLength) {
          results.push(Buffer.from(chars).toString('utf8'));
        }
        chars = [];
        start = -1;
      } else if (byte >= 0x80) {
        // UTF-8 multi-byte - try to include
        chars.push(byte);
      } else {
        // Control character - end of text
        if (chars.length >= minLength) {
          results.push(Buffer.from(chars).toString('utf8'));
        }
        chars = [];
        start = -1;
      }
    }
  }

  if (chars.length >= minLength) {
    results.push(Buffer.from(chars).toString('utf8'));
  }

  return results;
}

/**
 * Extracts the longest text block that looks like content.
 */
function extractPrimaryText(buffer: Buffer): string | null {
  // Try UTF-16LE first (common in OLK15)
  const utf16Texts = findUtf16Text(buffer, 20);

  // Try UTF-8 as fallback
  const utf8Texts = findUtf8Text(buffer, 20);

  // Combine and find the longest meaningful text
  const allTexts = [...utf16Texts, ...utf8Texts]
    .filter((t) => t.trim().length > 0)
    // eslint-disable-next-line no-control-regex
    .filter((t) => !t.match(/^[\x00-\x1f\x7f-\x9f]+$/)) // Filter control chars
    .sort((a, b) => b.length - a.length);

  return allTexts[0] ?? null;
}

/**
 * Checks if text looks like HTML content.
 */
function looksLikeHtml(text: string): boolean {
  return /<\/?[a-z][\s\S]*>/i.test(text);
}

/**
 * Finds HTML content in buffer.
 */
function findHtmlContent(buffer: Buffer): string | null {
  const text = buffer.toString('utf8', 0, Math.min(buffer.length, 1024 * 1024));

  // Look for HTML patterns
  const htmlStart = text.search(/<html[^>]*>/i);
  if (htmlStart >= 0) {
    const htmlEnd = text.indexOf('</html>', htmlStart);
    if (htmlEnd > htmlStart) {
      return text.substring(htmlStart, htmlEnd + 7);
    }
  }

  // Look for body content
  const bodyStart = text.search(/<body[^>]*>/i);
  if (bodyStart >= 0) {
    const bodyEnd = text.indexOf('</body>', bodyStart);
    if (bodyEnd > bodyStart) {
      return text.substring(bodyStart, bodyEnd + 7);
    }
  }

  return null;
}

// =============================================================================
// OLK15 Parser Implementation
// =============================================================================

/**
 * Result of parsing an OLK15 file.
 */
export interface Olk15ParseResult {
  readonly success: boolean;
  readonly text: string | null;
  readonly html: string | null;
  readonly error?: string;
}

/**
 * Parses an OLK15 file and extracts text content.
 */
export function parseOlk15File(filePath: string): Olk15ParseResult {
  try {
    if (!existsSync(filePath)) {
      return { success: false, text: null, html: null, error: 'File not found' };
    }

    const buffer = readFileSync(filePath);

    if (buffer.length === 0) {
      return { success: false, text: null, html: null, error: 'Empty file' };
    }

    // Try to find HTML content first
    const html = findHtmlContent(buffer);
    if (html != null) {
      return { success: true, text: null, html };
    }

    // Extract text content
    const text = extractPrimaryText(buffer);
    if (text != null) {
      // Check if the text is actually HTML
      if (looksLikeHtml(text)) {
        return { success: true, text: null, html: text };
      }
      return { success: true, text, html: null };
    }

    return { success: false, text: null, html: null, error: 'No text content found' };
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    return { success: false, text: null, html: null, error: message };
  }
}

// =============================================================================
// Content Reader Implementations
// =============================================================================

/**
 * OLK15 content reader for email bodies.
 */
export class Olk15EmailContentReader implements IContentReader {
  constructor(private readonly dataPath: string) {}

  readEmailBody(dataFilePath: string | null): string | null {
    if (dataFilePath == null) {
      return null;
    }

    const fullPath = join(this.dataPath, dataFilePath);
    const result = parseOlk15File(fullPath);

    if (result.success) {
      return result.html ?? result.text;
    }

    return null;
  }
}

/**
 * OLK15 content reader for calendar events.
 */
export class Olk15EventContentReader implements IEventContentReader {
  constructor(private readonly dataPath: string) {}

  readEventDetails(dataFilePath: string | null): EventDetails | null {
    if (dataFilePath == null) {
      return null;
    }

    const fullPath = join(this.dataPath, dataFilePath);
    const result = parseOlk15File(fullPath);

    if (!result.success) {
      return null;
    }

    const text = result.text ?? result.html ?? '';

    // Basic extraction - can be improved with format knowledge
    return {
      title: this.extractTitle(text),
      location: this.extractField(text, 'Location'),
      description: text,
      organizer: this.extractField(text, 'Organizer'),
      attendees: [],
    };
  }

  private extractTitle(text: string): string | null {
    // Try to find a subject/title line
    const subjectMatch = text.match(/Subject:\s*(.+?)(?:\r?\n|$)/i);
    if (subjectMatch?.[1] != null && subjectMatch[1].length > 0) {
      return subjectMatch[1].trim();
    }

    // Use first line as title
    const firstLine = text.split(/\r?\n/)[0]?.trim();
    return firstLine != null && firstLine.length > 0 ? firstLine : null;
  }

  private extractField(text: string, fieldName: string): string | null {
    const regex = new RegExp(`${fieldName}:\\s*(.+?)(?:\\r?\\n|$)`, 'i');
    const match = text.match(regex);
    return match?.[1]?.trim() ?? null;
  }
}

/**
 * OLK15 content reader for contacts.
 */
export class Olk15ContactContentReader implements IContactContentReader {
  constructor(private readonly dataPath: string) {}

  readContactDetails(dataFilePath: string | null): ContactDetails | null {
    if (dataFilePath == null) {
      return null;
    }

    const fullPath = join(this.dataPath, dataFilePath);
    const result = parseOlk15File(fullPath);

    if (!result.success) {
      return null;
    }

    // Contact parsing would require understanding the binary format
    // For now, return a basic structure
    return {
      firstName: null,
      lastName: null,
      middleName: null,
      nickname: null,
      company: null,
      jobTitle: null,
      department: null,
      emails: [],
      phones: [],
      addresses: [],
      notes: result.text,
    };
  }
}

/**
 * OLK15 content reader for tasks.
 */
export class Olk15TaskContentReader implements ITaskContentReader {
  constructor(private readonly dataPath: string) {}

  readTaskDetails(dataFilePath: string | null): TaskDetails | null {
    if (dataFilePath == null) {
      return null;
    }

    const fullPath = join(this.dataPath, dataFilePath);
    const result = parseOlk15File(fullPath);

    if (!result.success) {
      return null;
    }

    return {
      body: result.text ?? result.html,
      completedDate: null,
      reminderDate: null,
      categories: [],
    };
  }
}

/**
 * OLK15 content reader for notes.
 */
export class Olk15NoteContentReader implements INoteContentReader {
  constructor(private readonly dataPath: string) {}

  readNoteDetails(dataFilePath: string | null): NoteDetails | null {
    if (dataFilePath == null) {
      return null;
    }

    const fullPath = join(this.dataPath, dataFilePath);
    const result = parseOlk15File(fullPath);

    if (!result.success) {
      return null;
    }

    const body = result.text ?? result.html ?? '';
    const lines = body.split(/\r?\n/).filter((l) => l.trim().length > 0);

    return {
      title: lines[0] ?? null,
      body,
      preview: body.substring(0, 200),
      createdDate: null,
      categories: [],
    };
  }
}

// =============================================================================
// Factory Functions
// =============================================================================

/**
 * Creates all content readers for a given data path.
 */
export interface ContentReaders {
  readonly email: IContentReader;
  readonly event: IEventContentReader;
  readonly contact: IContactContentReader;
  readonly task: ITaskContentReader;
  readonly note: INoteContentReader;
}

/**
 * Creates content readers for all Outlook data types.
 */
export function createContentReaders(dataPath: string): ContentReaders {
  return {
    email: new Olk15EmailContentReader(dataPath),
    event: new Olk15EventContentReader(dataPath),
    contact: new Olk15ContactContentReader(dataPath),
    task: new Olk15TaskContentReader(dataPath),
    note: new Olk15NoteContentReader(dataPath),
  };
}

/**
 * Creates content readers using the default profile path.
 */
export function createDefaultContentReaders(profileName?: string): ContentReaders {
  const dataPath = getDefaultDataPath(profileName);
  return createContentReaders(dataPath);
}
