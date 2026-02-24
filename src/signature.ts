/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Email signature storage and auto-append logic.
 *
 * Signatures are stored as HTML at ~/.outlook-mcp/signature.html
 * and auto-appended to email bodies when creating/sending emails.
 */

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

const SIGNATURE_DIR = join(homedir(), '.outlook-mcp');
const SIGNATURE_FILE = join(SIGNATURE_DIR, 'signature.html');

/**
 * Ensures the signature directory exists.
 */
function ensureDir(): void {
  if (!existsSync(SIGNATURE_DIR)) {
    mkdirSync(SIGNATURE_DIR, { recursive: true, mode: 0o700 });
  }
}

/**
 * Strips HTML tags from a string and converts <br> to newlines.
 */
function stripHtml(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]*>/g, '');
}

/**
 * Reads the stored signature. Returns null if no signature file exists.
 */
export function readSignature(): string | null {
  if (!existsSync(SIGNATURE_FILE)) return null;
  return readFileSync(SIGNATURE_FILE, 'utf-8');
}

/**
 * Writes a signature to disk.
 * If contentType is 'text', wraps in <pre> tag for HTML storage.
 */
export function writeSignature(content: string, contentType: 'html' | 'text' = 'html'): void {
  ensureDir();
  const html = contentType === 'text' ? `<pre>${content}</pre>` : content;
  writeFileSync(SIGNATURE_FILE, html, { encoding: 'utf-8', mode: 0o600 });
}

/**
 * Appends the stored signature to an email body.
 *
 * For HTML bodies: appends with <br><br> separator.
 * For text bodies: appends with \n\n--\n separator and strips HTML from signature.
 * Returns the body unchanged if includeSignature is false or no signature exists.
 */
export function appendSignature(
  body: string,
  bodyType: 'html' | 'text',
  includeSignature: boolean
): string {
  if (!includeSignature) return body;

  const signature = readSignature();
  if (signature == null) return body;

  if (bodyType === 'html') {
    return `${body}<br><br>${signature}`;
  }

  return `${body}\n\n--\n${stripHtml(signature)}`;
}
