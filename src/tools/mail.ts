/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Mail-related MCP tools.
 *
 * Provides tools for listing folders, emails, and searching.
 */

import { existsSync } from 'fs';
import { dirname } from 'path';
import { z } from 'zod';
import type { IRepository, EmailRow, FolderRow } from '../database/repository.js';
import type { Folder, EmailSummary, Email, AttachmentInfo, PriorityValue, FlagStatusValue } from '../types/index.js';
import { MAX_ATTACHMENT_DOWNLOAD_SIZE } from '../types/mail.js';
import type { IAttachmentReader } from '../applescript/content-readers.js';
import type { SaveAttachmentResult } from '../applescript/parser.js';
import { appleTimestampToIso } from '../utils/dates.js';
import { extractPlainText } from '../parsers/html-stripper.js';
import { NotFoundError, ValidationError, AttachmentTooLargeError, AttachmentSaveError } from '../utils/errors.js';

// =============================================================================
// Input Schemas
// =============================================================================

export const ListFoldersInput = z.strictObject({});

export const ListEmailsInput = z.strictObject({
  folder_id: z.number().int().positive().describe('The folder ID to list emails from'),
  limit: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(50)
    .describe('Maximum number of emails to return (1-100)'),
  offset: z.number().int().min(0).default(0).describe('Number of emails to skip'),
  unread_only: z.boolean().default(false).describe('Only return unread emails'),
});

export const SearchEmailsInput = z.strictObject({
  query: z.string().min(1).describe('Search query (searches subject, sender, and preview)'),
  folder_id: z
    .number()
    .int()
    .positive()
    .optional()
    .describe('Optional folder ID to limit search to'),
  limit: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(50)
    .describe('Maximum number of emails to return (1-100)'),
});

export const SearchEmailsAdvancedInput = z.strictObject({
  query: z.string().min(1).describe(
    'KQL search query. Examples: from:alice, subject:"quarterly report", hasAttachments:true, received>=2024-01-01. Combine with AND/OR.'
  ),
  folder_id: z.number().int().positive().optional().describe('Optional folder ID to search within'),
  limit: z.number().int().min(1).max(100).default(50).describe('Maximum results (1-100)'),
});

export const GetEmailInput = z.strictObject({
  email_id: z.number().int().positive().describe('The email ID to retrieve'),
  include_body: z.boolean().default(true).describe('Include the email body in the response'),
  strip_html: z.boolean().default(true).describe('Strip HTML tags from the body'),
});

export const GetEmailsInput = z.strictObject({
  email_ids: z.array(z.number().int().positive()).min(1).max(25)
    .describe('Array of email IDs to fetch (max 25)'),
  include_body: z.boolean().default(false).describe('Include full email body'),
  strip_html: z.boolean().default(false).describe('Strip HTML tags from body'),
});

export const ListConversationInput = z.strictObject({
  message_id: z.number().int().positive().describe('Any message ID from the conversation thread'),
  limit: z.number().int().min(1).max(100).default(25).describe('Maximum messages to return'),
});

export const GetUnreadCountInput = z.strictObject({
  folder_id: z
    .number()
    .int()
    .positive()
    .optional()
    .describe('Optional folder ID to get unread count for'),
});

export const ListAttachmentsInput = z.strictObject({
  email_id: z.number().int().positive().describe('The email ID to list attachments for'),
});

export const DownloadAttachmentInput = z.strictObject({
  email_id: z.number().int().positive().describe('The email ID containing the attachment'),
  attachment_index: z.number().int().positive().describe('The 1-based index of the attachment (from list_attachments)'),
  save_path: z.string().min(1).describe('Absolute file path where the attachment should be saved'),
});

export const CheckNewEmailsInput = z.strictObject({
  folder_id: z.number().int().positive().describe('Folder ID to check for new emails'),
});

// =============================================================================
// Type Definitions
// =============================================================================

export type ListFoldersParams = z.infer<typeof ListFoldersInput>;
export type ListEmailsParams = z.infer<typeof ListEmailsInput>;
export type SearchEmailsParams = z.infer<typeof SearchEmailsInput>;
export type GetEmailParams = z.infer<typeof GetEmailInput>;
export type GetEmailsParams = z.infer<typeof GetEmailsInput>;
export type GetUnreadCountParams = z.infer<typeof GetUnreadCountInput>;
export type ListAttachmentsParams = z.infer<typeof ListAttachmentsInput>;
export type DownloadAttachmentParams = z.infer<typeof DownloadAttachmentInput>;
export type CheckNewEmailsParams = z.infer<typeof CheckNewEmailsInput>;

// =============================================================================
// Transformers
// =============================================================================

/**
 * Transforms a database folder row to domain Folder type.
 */
function transformFolder(row: FolderRow): Folder {
  return {
    id: row.id,
    name: row.name ?? 'Unnamed',
    parentId: row.parentId,
    specialType: row.specialType,
    folderType: row.folderType,
    accountId: row.accountId,
    messageCount: row.messageCount,
    unreadCount: row.unreadCount,
  };
}

/**
 * Transforms a database email row to domain EmailSummary type.
 */
function transformEmailSummary(row: EmailRow): EmailSummary {
  return {
    id: row.id,
    folderId: row.folderId,
    subject: row.subject,
    sender: row.sender,
    senderAddress: row.senderAddress,
    preview: row.preview,
    isRead: row.isRead === 1,
    timeReceived: appleTimestampToIso(row.timeReceived),
    timeSent: appleTimestampToIso(row.timeSent),
    hasAttachment: row.hasAttachment === 1,
    priority: row.priority as PriorityValue,
    flagStatus: row.flagStatus as FlagStatusValue,
    categories: parseCategories(row.categories),
  };
}

/**
 * Parses categories from the database buffer.
 * Outlook stores categories as a null-delimited or comma-delimited buffer.
 */
function parseCategories(buffer: Buffer | null): readonly string[] {
  if (buffer == null || buffer.length === 0) {
    return [];
  }

  try {
    const text = buffer.toString('utf-8');
    // Categories may be stored as null-delimited or comma-delimited strings
    const categories = text.includes('\0')
      ? text.split('\0').filter(s => s.length > 0)
      : text.split(',').map(s => s.trim()).filter(s => s.length > 0);
    return categories;
  } catch {
    return [];
  }
}

/**
 * Transforms a database email row to domain Email type.
 */
function transformEmail(row: EmailRow, body: string | null, stripHtml: boolean): Email {
  const summary = transformEmailSummary(row);

  // Process the body
  let processedBody: string | null = null;
  let htmlBody: string | null = null;

  if (body != null) {
    htmlBody = body;
    processedBody = stripHtml ? extractPlainText(body) : body;
  }

  return {
    ...summary,
    recipients: row.recipients,
    displayTo: row.displayTo,
    toAddresses: row.toAddresses,
    ccAddresses: row.ccAddresses,
    size: row.size,
    messageId: row.messageId ?? null,
    conversationId: row.conversationId ?? null,
    body: processedBody,
    htmlBody: stripHtml ? null : htmlBody,
  };
}

// =============================================================================
// Parser Interface (for body content)
// =============================================================================

/**
 * Interface for reading email body content from data files.
 */
export interface IContentReader {
  /**
   * Reads the body content from the given data file path.
   * Returns null if the file cannot be read.
   */
  readEmailBody(dataFilePath: string | null): string | null;
}

/**
 * Default content reader that returns null (body reading not implemented yet).
 * Will be replaced with OLK15 parser implementation.
 */
export const nullContentReader: IContentReader = {
  readEmailBody: (): string | null => null,
};

// =============================================================================
// Mail Tools Class
// =============================================================================

/**
 * Mail tools implementation with dependency injection.
 */
export class MailTools {
  constructor(
    private readonly repository: IRepository,
    private readonly contentReader: IContentReader = nullContentReader,
    private readonly attachmentReader?: IAttachmentReader
  ) {}

  /**
   * Lists all mail folders with message and unread counts.
   */
  listFolders(_params: ListFoldersParams): Folder[] {
    const rows = this.repository.listFolders();
    return rows.map(transformFolder);
  }

  /**
   * Lists emails in a folder with pagination.
   */
  listEmails(params: ListEmailsParams): EmailSummary[] {
    const { folder_id, limit, offset, unread_only } = params;

    const rows = unread_only
      ? this.repository.listUnreadEmails(folder_id, limit, offset)
      : this.repository.listEmails(folder_id, limit, offset);

    return rows.map(transformEmailSummary);
  }

  /**
   * Searches emails by subject, sender, or preview.
   */
  searchEmails(params: SearchEmailsParams): EmailSummary[] {
    const { query, folder_id, limit } = params;

    const rows =
      folder_id != null
        ? this.repository.searchEmailsInFolder(folder_id, query, limit)
        : this.repository.searchEmails(query, limit);

    return rows.map(transformEmailSummary);
  }

  /**
   * Gets a single email by ID with optional body content.
   */
  getEmail(params: GetEmailParams): (Email & { attachments: AttachmentInfo[] }) | null {
    const { email_id, include_body, strip_html } = params;

    const row = this.repository.getEmail(email_id);
    if (row == null) {
      return null;
    }

    // Get body content if requested
    let body: string | null = null;
    if (include_body && row.dataFilePath != null) {
      body = this.contentReader.readEmailBody(row.dataFilePath);
    }

    // Get attachment metadata if available
    let attachments: AttachmentInfo[] = [];
    if (this.attachmentReader != null && row.hasAttachment === 1) {
      attachments = this.attachmentReader.listAttachments(email_id);
    }

    return {
      ...transformEmail(row, body, strip_html),
      attachments,
    };
  }

  /**
   * Gets the unread email count.
   */
  getUnreadCount(params: GetUnreadCountParams): { count: number } {
    const { folder_id } = params;

    const count =
      folder_id != null
        ? this.repository.getUnreadCountByFolder(folder_id)
        : this.repository.getUnreadCount();

    return { count };
  }

  /**
   * Lists attachment metadata for an email.
   */
  listAttachments(params: ListAttachmentsParams): AttachmentInfo[] {
    if (this.attachmentReader == null) {
      return [];
    }

    const { email_id } = params;

    const row = this.repository.getEmail(email_id);
    if (row == null) {
      throw new NotFoundError('Email', email_id);
    }

    return this.attachmentReader.listAttachments(email_id);
  }

  /**
   * Downloads/saves an attachment to disk.
   */
  downloadAttachment(params: DownloadAttachmentParams): {
    name: string;
    savedTo: string;
    size: number;
  } {
    if (this.attachmentReader == null) {
      throw new ValidationError('Attachment reader not available');
    }

    const { email_id, attachment_index, save_path } = params;

    const row = this.repository.getEmail(email_id);
    if (row == null) {
      throw new NotFoundError('Email', email_id);
    }

    // Validate save path directory exists
    const dir = dirname(save_path);
    if (!existsSync(dir)) {
      throw new ValidationError(`Directory does not exist: ${dir}`);
    }

    // Get attachment list to validate index and check size
    const attachments = this.attachmentReader.listAttachments(email_id);
    const attachment = attachments.find(a => a.index === attachment_index);

    if (attachment == null) {
      throw new NotFoundError('Attachment', attachment_index);
    }

    if (attachment.size > MAX_ATTACHMENT_DOWNLOAD_SIZE) {
      throw new AttachmentTooLargeError(
        attachment.name,
        attachment.size,
        MAX_ATTACHMENT_DOWNLOAD_SIZE
      );
    }

    const result: SaveAttachmentResult = this.attachmentReader.saveAttachment(
      email_id,
      attachment_index,
      save_path
    );

    if (!result.success) {
      throw new AttachmentSaveError(
        attachment.name,
        result.error ?? 'Unknown error'
      );
    }

    return {
      name: result.name ?? attachment.name,
      savedTo: result.savedTo ?? save_path,
      size: result.fileSize ?? attachment.size,
    };
  }
}

/**
 * Creates mail tools with the given repository.
 */
export function createMailTools(
  repository: IRepository,
  contentReader: IContentReader = nullContentReader,
  attachmentReader?: IAttachmentReader
): MailTools {
  return new MailTools(repository, contentReader, attachmentReader);
}
