/**
 * Mail-related MCP tools.
 *
 * Provides tools for listing folders, emails, and searching.
 */

import { z } from 'zod';
import type { IRepository, EmailRow, FolderRow } from '../database/repository.js';
import type { Folder, EmailSummary, Email, PriorityValue, FlagStatusValue } from '../types/index.js';
import { appleTimestampToIso } from '../utils/dates.js';
import { extractPlainText } from '../parsers/html-stripper.js';

// =============================================================================
// Input Schemas
// =============================================================================

export const ListFoldersInput = z.object({}).strict();

export const ListEmailsInput = z
  .object({
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
  })
  .strict();

export const SearchEmailsInput = z
  .object({
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
  })
  .strict();

export const GetEmailInput = z
  .object({
    email_id: z.number().int().positive().describe('The email ID to retrieve'),
    include_body: z.boolean().default(true).describe('Include the email body in the response'),
    strip_html: z.boolean().default(true).describe('Strip HTML tags from the body'),
  })
  .strict();

export const GetUnreadCountInput = z
  .object({
    folder_id: z
      .number()
      .int()
      .positive()
      .optional()
      .describe('Optional folder ID to get unread count for'),
  })
  .strict();

// =============================================================================
// Type Definitions
// =============================================================================

export type ListFoldersParams = z.infer<typeof ListFoldersInput>;
export type ListEmailsParams = z.infer<typeof ListEmailsInput>;
export type SearchEmailsParams = z.infer<typeof SearchEmailsInput>;
export type GetEmailParams = z.infer<typeof GetEmailInput>;
export type GetUnreadCountParams = z.infer<typeof GetUnreadCountInput>;

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
  };
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
    private readonly contentReader: IContentReader = nullContentReader
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
  getEmail(params: GetEmailParams): Email | null {
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

    return transformEmail(row, body, strip_html);
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
}

/**
 * Creates mail tools with the given repository.
 */
export function createMailTools(
  repository: IRepository,
  contentReader: IContentReader = nullContentReader
): MailTools {
  return new MailTools(repository, contentReader);
}
