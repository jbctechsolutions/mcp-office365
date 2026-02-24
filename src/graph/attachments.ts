/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Attachment upload and download helpers for Microsoft Graph API.
 *
 * Handles:
 * - File I/O for reading/writing attachments
 * - Size-based routing (inline vs chunked upload)
 * - Chunked upload sessions for large files (> 3MB)
 * - Safe filename resolution with path traversal prevention
 */

import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

import type { GraphClient } from './client/index.js';

/** 3MB threshold for inline vs chunked upload. */
const INLINE_MAX_BYTES = 3 * 1024 * 1024;

/** Microsoft's recommended chunk size (~3.75MB). */
const CHUNK_SIZE = 3932160;

/**
 * Returns the download directory for attachments.
 *
 * Reads `MCP_OUTLOOK_DOWNLOAD_DIR` env var, falls back to `os.tmpdir()`.
 * Creates a subdirectory `mcp-outlook-attachments` inside it.
 */
export function getDownloadDir(): string {
  const base = process.env['MCP_OUTLOOK_DOWNLOAD_DIR'] ?? os.tmpdir();
  const dir = path.join(base, 'mcp-outlook-attachments');
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

/**
 * Sanitizes a filename to prevent path traversal attacks.
 *
 * Strips directory separators (`/`, `\`), strips `..` segments,
 * and trims whitespace. Falls back to `'attachment'` if result is empty.
 */
export function sanitizeFilename(name: string): string {
  // Strip .. segments
  let sanitized = name.replace(/\.\./g, '');
  // Take only the last component after any path separator
  const forwardSlashIdx = sanitized.lastIndexOf('/');
  if (forwardSlashIdx >= 0) {
    sanitized = sanitized.substring(forwardSlashIdx + 1);
  }
  const backslashIdx = sanitized.lastIndexOf('\\');
  if (backslashIdx >= 0) {
    sanitized = sanitized.substring(backslashIdx + 1);
  }
  sanitized = sanitized.trim();
  return sanitized.length > 0 ? sanitized : 'attachment';
}

/**
 * Generates a unique file path in the given directory.
 *
 * Sanitizes the filename, then checks if it already exists.
 * If it does, appends a numeric suffix (e.g., `file(1).txt`, `file(2).txt`).
 */
export function resolveFilePath(downloadDir: string, filename: string): string {
  const safeName = sanitizeFilename(filename);
  let candidate = path.join(downloadDir, safeName);

  if (!fs.existsSync(candidate)) {
    return candidate;
  }

  const ext = path.extname(safeName);
  const base = safeName.substring(0, safeName.length - ext.length);

  let counter = 1;
  for (;;) { // eslint-disable-line no-constant-condition
    candidate = path.join(downloadDir, `${base}(${counter})${ext}`);
    if (!fs.existsSync(candidate)) {
      return candidate;
    }
    counter++;
  }
}

/**
 * Uploads a file as an attachment to a message.
 *
 * Routes based on file size:
 * - <= 3MB: Inline base64 upload via `client.addAttachment`
 * - > 3MB: Chunked upload session via `client.createUploadSession` + raw fetch
 *
 * @param client - The Graph API client
 * @param messageId - The message to attach the file to
 * @param filePath - Local path to the file to upload
 * @param name - Optional display name (defaults to basename of filePath)
 * @param contentType - Optional MIME type (defaults to 'application/octet-stream')
 */
export async function uploadAttachment(
  client: GraphClient,
  messageId: string,
  filePath: string,
  name?: string,
  contentType?: string,
): Promise<void> {
  const fileName = name ?? path.basename(filePath);
  const mimeType = contentType ?? 'application/octet-stream';
  const stat = fs.statSync(filePath);
  const fileSize = stat.size;

  if (fileSize <= INLINE_MAX_BYTES) {
    // Inline upload: read file, base64 encode, POST as attachment
    const fileContent = fs.readFileSync(filePath);
    const base64 = fileContent.toString('base64');

    await client.addAttachment(messageId, {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: fileName,
      contentBytes: base64,
      contentType: mimeType,
    });
  } else {
    // Chunked upload session for large files
    const session = await client.createUploadSession(messageId, {
      AttachmentItem: {
        attachmentType: 'file',
        name: fileName,
        size: fileSize,
      },
    });

    const uploadUrl = session.uploadUrl;
    const fileContent = fs.readFileSync(filePath);

    let offset = 0;
    while (offset < fileSize) {
      const end = Math.min(offset + CHUNK_SIZE, fileSize);
      const chunk = fileContent.subarray(offset, end);
      const contentRange = `bytes ${offset}-${end - 1}/${fileSize}`;

      const response = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Range': contentRange,
          'Content-Length': String(chunk.length),
        },
        body: chunk,
      });

      if (!response.ok) {
        throw new Error(
          `Upload chunk failed: ${response.status} ${response.statusText}`
        );
      }

      offset = end;
    }
  }
}

/**
 * Downloads an attachment from a message and saves it to disk.
 *
 * Decodes `contentBytes` from base64 and writes the file to the download directory.
 *
 * @param client - The Graph API client
 * @param messageId - The message containing the attachment
 * @param attachmentId - The attachment to download
 * @returns Metadata about the downloaded file including its local path
 */
export async function downloadAttachment(
  client: GraphClient,
  messageId: string,
  attachmentId: string,
): Promise<{ filePath: string; name: string; size: number; contentType: string }> {
  const attachment = await client.getAttachment(messageId, attachmentId);

  const rawName = attachment.name ?? '';
  const safeName = sanitizeFilename(rawName);
  const size = attachment.size ?? 0;
  const attachmentContentType = attachment.contentType ?? 'application/octet-stream';

  const downloadDir = getDownloadDir();
  const filePath = resolveFilePath(downloadDir, safeName);

  const contentBytes = (attachment as { contentBytes?: string }).contentBytes ?? '';
  const buffer = Buffer.from(contentBytes, 'base64');
  fs.writeFileSync(filePath, buffer);

  return {
    filePath,
    name: safeName,
    size,
    contentType: attachmentContentType,
  };
}
