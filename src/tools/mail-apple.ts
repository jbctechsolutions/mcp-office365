/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * AppleScript-backend mail tools (v3 registry-driven architecture, U2 — dual
 * backend). Holds the mail READ logic that previously lived inline in the
 * `handleAppleScriptToolCall` switch, so the registry handlers stay thin and
 * branch on `ctx.backend`.
 *
 * The advertised (canonical) schemas are Graph-shaped (Graph is the default
 * backend). This backend receives the superset params and maps only the fields
 * it supports, exactly as the pre-registry dispatch did.
 */

import type { MailTools } from './mail.js';
import type {
  ListFoldersToolParams,
  ListEmailsParams,
  SearchEmailsParams,
  GetEmailParams,
  GetEmailsParams,
  GetUnreadCountParams,
  ListAttachmentsParams,
  DownloadAttachmentParams,
} from './mail.js';
import type { IAccountRepository, IMailSender, MailSenderSendEmailParams } from '../applescript/index.js';
import type { ToolResult } from '../registry/types.js';
import type { SendEmailParams } from './mail.js';

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Resolves an account_id parameter to an array of account IDs. Self-contained
 * copy of the legacy dispatch helper (the original stays in index.ts, where it
 * is still used by the account tools).
 *
 * - undefined → [defaultAccountId]
 * - "all" → all account IDs
 * - number → [number]
 * - number[] → number[]
 */
function resolveAccountIds(
  accountId: number | number[] | 'all' | undefined,
  accountRepository: IAccountRepository
): number[] {
  if (accountId === undefined) {
    const defaultId = accountRepository.getDefaultAccountId();
    return defaultId !== null ? [defaultId] : [];
  }

  if (accountId === 'all') {
    const accounts = accountRepository.listAccounts();
    return accounts.map(acc => acc.id);
  }

  if (typeof accountId === 'number') {
    return [accountId];
  }

  if (Array.isArray(accountId)) {
    return accountId;
  }

  const defaultId = accountRepository.getDefaultAccountId();
  return defaultId !== null ? [defaultId] : [];
}

/**
 * AppleScript mail tools. Each method mirrors the extracted inline AppleScript
 * case body and returns an MCP `ToolResult`.
 */
export class AppleMailTools {
  constructor(
    private readonly mailTools: MailTools,
    private readonly accountRepository: IAccountRepository,
    private readonly mailSender: IMailSender | null,
  ) {}

  /** Single-shot send via the AppleScript mail sender (Graph uses two-phase). */
  sendEmail(params: SendEmailParams): ToolResult {
    if (this.mailSender == null) {
      return { content: [{ type: 'text', text: 'Email sending is not available' }], isError: true };
    }
    let sendParams: MailSenderSendEmailParams = {
      to: params.to,
      subject: params.subject,
      body: params.body,
      bodyType: params.body_type ?? 'plain',
    };
    if (params.cc != null) sendParams = { ...sendParams, cc: params.cc };
    if (params.bcc != null) sendParams = { ...sendParams, bcc: params.bcc };
    if (params.reply_to != null) sendParams = { ...sendParams, replyTo: params.reply_to };
    if (params.attachments != null) {
      sendParams = {
        ...sendParams,
        attachments: params.attachments.map((a) => (a.name != null ? { path: a.path, name: a.name } : { path: a.path })),
      };
    }
    if (params.inline_images != null) {
      sendParams = {
        ...sendParams,
        inlineImages: params.inline_images.map((img) => ({ path: img.path, contentId: img.content_id })),
      };
    }
    if (params.account_id != null) sendParams = { ...sendParams, accountId: params.account_id };

    const sent = this.mailSender.sendEmail(sendParams);
    return {
      content: [{ type: 'text', text: JSON.stringify({ message_id: sent.messageId, sent_at: sent.sentAt, status: 'sent' }, null, 2) }],
    };
  }

  listFolders(params: ListFoldersToolParams): ToolResult {
    const accountIds = resolveAccountIds(params.account_id, this.accountRepository);

    // If querying multiple accounts, use grouped format
    if (accountIds.length > 1 || params.account_id === 'all') {
      const foldersWithAccount = this.accountRepository.listMailFoldersByAccounts(accountIds);
      const accounts = this.accountRepository.listAccounts();

      // Group folders by account
      const groupedByAccount = accountIds.map(accountId => {
        const account = accounts.find(a => a.id === accountId);
        const folders = foldersWithAccount
          .filter(f => f.accountId === accountId)
          .map(f => ({
            id: f.id,
            name: f.name,
            unreadCount: f.unreadCount,
            messageCount: f.messageCount,
          }));

        return {
          account_id: accountId,
          account_name: account?.name ?? null,
          account_email: account?.email ?? null,
          folders,
        };
      });

      return jsonResult({ accounts: groupedByAccount });
    }

    // Single account - use existing format for backward compatibility
    return jsonResult(this.mailTools.listFolders({}));
  }

  listEmails(params: ListEmailsParams): ToolResult {
    return jsonResult(this.mailTools.listEmails(params));
  }

  searchEmails(params: SearchEmailsParams): ToolResult {
    return jsonResult(this.mailTools.searchEmails(params));
  }

  getEmail(params: GetEmailParams): ToolResult {
    const result = this.mailTools.getEmail(params);
    if (result == null) {
      return { content: [{ type: 'text', text: 'Email not found' }], isError: true };
    }
    return jsonResult(result);
  }

  getEmails(params: GetEmailsParams): ToolResult {
    const results = params.email_ids.map((id) => {
      try {
        const email = this.mailTools.getEmail({ email_id: id, include_body: params.include_body, strip_html: params.strip_html });
        if (email == null) return { id, error: 'Not found' };
        return email;
      } catch (err) {
        // Per-id isolation — one bad id must not abort the whole batch.
        return { id, error: err instanceof Error ? err.message : 'Unresolvable id' };
      }
    });
    return jsonResult({ emails: results });
  }

  getUnreadCount(params: GetUnreadCountParams): ToolResult {
    return jsonResult(this.mailTools.getUnreadCount(params));
  }

  listAttachments(params: ListAttachmentsParams): ToolResult {
    return jsonResult(this.mailTools.listAttachments(params));
  }

  downloadAttachment(params: DownloadAttachmentParams): ToolResult {
    return jsonResult(this.mailTools.downloadAttachment(params));
  }
}
