/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Graph-backend mail tools (v3 registry-driven architecture, U2 — dual
 * backend). Holds the mail READ logic that previously lived inline in the
 * `handleGraphToolCall` switch, so the registry handlers stay thin and branch
 * on `ctx.backend`.
 */

import type { GraphRepository } from '../graph/repository.js';
import type { GraphContentReaders } from '../graph/content-readers.js';
import type { FolderRow, EmailRow } from '../database/repository.js';
import { unixTimestampToLocalIso } from '../graph/mappers/utils.js';
import { compileEmailSearch } from '../search/compiler.js';
import type { ToolResult } from '../registry/types.js';
import type {
  ListFoldersToolParams,
  ListEmailsParams,
  SearchEmailsParams,
  GetEmailParams,
  GetEmailsParams,
  GetUnreadCountParams,
  ListAttachmentsParams,
  DownloadAttachmentParams,
  SearchEmailsAdvancedParams,
  CheckNewEmailsParams,
  ListConversationParams,
  GetMessageHeadersParams,
  GetMessageMimeParams,
  GetMailTipsParams,
} from './mail.js';

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Transforms a folder row into the shape returned by the graph backend's
 * `list_folders` tool. Self-contained copy of the legacy graph-mode transform.
 */
function transformFolderRow(row: FolderRow): {
  id: number;
  name: string;
  parentId: number | null;
  specialType: number;
  folderType: number;
  accountId: number;
  messageCount: number;
  unreadCount: number;
} {
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
 * Transforms an email row into the shape returned by the graph backend's mail
 * READ tools. Uses Unix timestamps (not Apple epoch). Self-contained copy of
 * the legacy graph-mode transform.
 */
function transformEmailRow(row: EmailRow): {
  id: string | number;
  folderId: number | null;
  subject: string | null;
  sender: string | null;
  senderAddress: string | null;
  preview: string | null;
  isRead: boolean;
  timeReceived: string | null;
  timeSent: string | null;
  hasAttachment: boolean;
  priority: number | null;
  flagStatus: number | null;
  categories: readonly string[];
} {
  return {
    id: row.id,
    folderId: row.folderId,
    subject: row.subject,
    sender: row.sender,
    senderAddress: row.senderAddress,
    preview: row.preview,
    isRead: row.isRead === 1,
    timeReceived: unixTimestampToLocalIso(row.timeReceived),
    timeSent: unixTimestampToLocalIso(row.timeSent),
    hasAttachment: row.hasAttachment === 1,
    priority: row.priority,
    flagStatus: row.flagStatus,
    categories: parseEmailCategories(row.categories),
  };
}

function parseEmailCategories(buffer: Buffer | null): string[] {
  if (buffer == null || buffer.length === 0) return [];
  try {
    const text = buffer.toString('utf-8');
    return text.includes('\0')
      ? text.split('\0').filter(s => s.length > 0)
      : text.split(',').map(s => s.trim()).filter(s => s.length > 0);
  } catch {
    return [];
  }
}

function stripHtml(html: string): string {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Graph mail tools. Each method mirrors the extracted inline graph case body
 * and returns an MCP `ToolResult`.
 */
export class GraphMailTools {
  constructor(
    private readonly repository: GraphRepository,
    private readonly contentReaders: GraphContentReaders
  ) {}

  async listFolders(_params: ListFoldersToolParams): Promise<ToolResult> {
    const folders = await this.repository.listFoldersAsync();
    return jsonResult({ folders: folders.map(transformFolderRow) });
  }

  async listEmails(params: ListEmailsParams): Promise<ToolResult> {
    const emails = params.unread_only
      ? await this.repository.listUnreadEmailsAsync(params.folder_id, params.limit, params.offset)
      : await this.repository.listEmailsAsync(params.folder_id, params.limit, params.offset);
    return jsonResult({ emails: emails.map(transformEmailRow) });
  }

  async searchEmails(params: SearchEmailsParams): Promise<ToolResult> {
    const emails = params.folder_id != null
      ? await this.repository.searchEmailsInFolderAsync(params.folder_id, params.query, params.limit)
      : await this.repository.searchEmailsAsync(params.query, params.limit);
    return jsonResult({ emails: emails.map(transformEmailRow) });
  }

  async getEmail(params: GetEmailParams): Promise<ToolResult> {
    const email = await this.repository.getEmailAsync(params.email_id);
    if (email == null) {
      return { content: [{ type: 'text', text: 'Email not found' }], isError: true };
    }

    let body: string | null = null;
    if (params.include_body) {
      body = await this.contentReaders.email.readEmailBodyAsync(email.dataFilePath);
      if (params.strip_html && body != null) {
        body = stripHtml(body);
      }
    }

    return jsonResult({ ...transformEmailRow(email), body });
  }

  async getEmails(params: GetEmailsParams): Promise<ToolResult> {
    const results = await Promise.all(
      params.email_ids.map(async (id) => {
        const email = await this.repository.getEmailAsync(id);
        if (email == null) return { id, error: 'Not found' };
        let body: string | null = null;
        if (params.include_body) {
          body = await this.contentReaders.email.readEmailBodyAsync(email.dataFilePath);
          if (params.strip_html && body != null) body = stripHtml(body);
        }
        return { ...transformEmailRow(email), body };
      })
    );
    return jsonResult({ emails: results });
  }

  async getUnreadCount(params: GetUnreadCountParams): Promise<ToolResult> {
    const count = params.folder_id != null
      ? await this.repository.getUnreadCountByFolderAsync(params.folder_id)
      : await this.repository.getUnreadCountAsync();
    return jsonResult({ total: count });
  }

  async listAttachments(params: ListAttachmentsParams): Promise<ToolResult> {
    const attachments = await this.repository.listAttachmentsAsync(params.email_id);
    return jsonResult({ attachments });
  }

  async downloadAttachment(params: DownloadAttachmentParams): Promise<ToolResult> {
    const result = await this.repository.downloadAttachmentAsync(params.attachment_index);
    return jsonResult(result);
  }

  async searchEmailsAdvanced(params: SearchEmailsAdvancedParams): Promise<ToolResult> {
    const { limit, ...criteria } = params;
    // Compile structured criteria to the correct Graph mechanism (U7/D9). An
    // empty or malformed-date query throws a typed VALIDATION error → envelope.
    const compiled = compileEmailSearch(criteria);
    const emails = await this.repository.searchEmailsStructuredAsync(compiled, limit);
    return jsonResult({ emails: emails.map(transformEmailRow) });
  }

  async checkNewEmails(params: CheckNewEmailsParams): Promise<ToolResult> {
    const deltaResult = await this.repository.checkNewEmailsAsync(params.folder_id);
    return jsonResult({
      emails: deltaResult.emails.map(transformEmailRow),
      is_initial_sync: deltaResult.isInitialSync,
      count: deltaResult.emails.length,
    });
  }

  async listConversation(params: ListConversationParams): Promise<ToolResult> {
    const emails = await this.repository.listConversationAsync(params.message_id, params.limit);
    return jsonResult({ emails: emails.map(transformEmailRow) });
  }

  async getMessageHeaders(params: GetMessageHeadersParams): Promise<ToolResult> {
    const headers = await this.repository.getMessageHeadersAsync(params.email_id);
    return jsonResult({ headers });
  }

  async getMessageMime(params: GetMessageMimeParams): Promise<ToolResult> {
    const result = await this.repository.getMessageMimeAsync(params.email_id);
    return jsonResult({ success: true, file_path: result.filePath });
  }

  async getMailTips(params: GetMailTipsParams): Promise<ToolResult> {
    const tips = await this.repository.getMailTipsAsync(params.email_addresses);
    return jsonResult({ mail_tips: tips });
  }
}
