/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Mail-related MCP tools.
 *
 * Provides tools for listing folders, emails, and searching.
 */

import { z } from 'zod';
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset } from '../registry/context.js';
import { Id } from '../ids/schema.js';
import type { ToolDefinition } from '../registry/types.js';
import type { GraphMailTools } from './mail-graph.js';

// The advertised (canonical) schemas below are Graph-shaped — Graph is the
// only backend.
declare module '../registry/types.js' {
  interface GraphToolsets {
    mailGraph: GraphMailTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

/**
 * Shared email/message id schema (U6): the canonical `em_` message-id schema.
 * A raw Graph id is also accepted; a legacy numeric id is a validation error
 * (numeric strings still route to resolveId → NUMERIC_ID_UNSUPPORTED).
 */
const EmailIdSchema = Id.message;

export const ListFoldersInput = z.strictObject({});

/**
 * Canonical (advertised) input for the `list_folders` tool. The Graph backend
 * ignores `account_id` and returns the default mailbox's folders.
 */
export const ListFoldersToolInput = z.strictObject({
  account_id: z
    .union([z.number(), z.array(z.number()), z.literal('all')])
    .optional()
    .describe('Account filter: number (specific account), array (multiple accounts), "all" (all accounts), or omit for default account'),
});

export const ListEmailsInput = z.strictObject({
  folder_id: Id.folder.describe('The folder to list emails from — a `fd_` token from list_folders.'),
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
  folder_id: Id.folder.optional().describe('Optional folder to limit search to — a `fd_` token from list_folders.'),
  limit: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(50)
    .describe('Maximum number of emails to return (1-100)'),
});

/**
 * Structured email search (U7 / D9). Replaces raw KQL — the #1 failure class —
 * with typed criteria compiled server-side to the correct Graph mechanism
 * ($filter / quoted $search / /search/query), so operator syntax, quoting, and
 * date formatting are handled for you. Provide at least one criterion.
 */
export const SearchEmailsAdvancedInput = z.strictObject({
  from: z.string().min(1).optional().describe('Sender email address (exact match)'),
  to: z.string().min(1).optional().describe('Recipient email address (exact match)'),
  subject_contains: z.string().min(1).optional().describe('Text the subject must contain'),
  body_contains: z.string().min(1).optional().describe('Text the body must contain'),
  text: z.string().min(1).optional().describe('Free-text search across the whole message'),
  received_after: z.string().min(1).optional().describe('Received on/after this ISO date (YYYY-MM-DD) or datetime'),
  received_before: z.string().min(1).optional().describe('Received on/before this ISO date (YYYY-MM-DD) or datetime'),
  has_attachments: z.boolean().optional().describe('Only messages with attachments'),
  is_unread: z.boolean().optional().describe('Only unread messages'),
  importance: z.enum(['low', 'normal', 'high']).optional().describe('Importance level'),
  limit: z.number().int().min(1).max(100).default(50).describe('Maximum results (1-100)'),
});

export const GetEmailInput = z.strictObject({
  email_id: EmailIdSchema.describe('The email ID to retrieve'),
  include_body: z.boolean().default(true).describe('Include the email body in the response'),
  strip_html: z.boolean().default(true).describe('Strip HTML tags from the body'),
});

export const GetEmailsInput = z.strictObject({
  email_ids: z.array(EmailIdSchema).min(1).max(25)
    .describe('Array of email IDs to fetch (max 25)'),
  include_body: z.boolean().default(false).describe('Include full email body'),
  strip_html: z.boolean().default(false).describe('Strip HTML tags from body'),
});

export const ListConversationInput = z.strictObject({
  message_id: EmailIdSchema.describe('Any message ID from the conversation thread'),
  limit: z.number().int().min(1).max(100).default(25).describe('Maximum messages to return'),
});

export const GetUnreadCountInput = z.strictObject({
  folder_id: z
    .string()
    .min(1)
    .optional()
    .describe('Optional folder ID to get unread count for'),
});

export const ListAttachmentsInput = z.strictObject({
  email_id: EmailIdSchema.describe('The email ID to list attachments for'),
});

export const DownloadAttachmentInput = z.strictObject({
  email_id: EmailIdSchema.describe('The email ID containing the attachment'),
  attachment_id: Id.attachment,
  save_path: z.string().min(1).describe('Absolute file path where the attachment should be saved'),
});

export const CheckNewEmailsInput = z.strictObject({
  folder_id: Id.folder.describe('Folder to check for new emails — a `fd_` token from list_folders.'),
});

export const GetMessageHeadersInput = z.strictObject({
  email_id: EmailIdSchema.describe('Email ID'),
});

export const GetMessageMimeInput = z.strictObject({
  email_id: EmailIdSchema.describe('Email ID'),
});

export const GetMailTipsInput = z.strictObject({
  email_addresses: z.array(z.string().email()).min(1).max(20).describe('Email addresses to check'),
});

// =============================================================================
// Type Definitions
// =============================================================================

export type ListFoldersParams = z.infer<typeof ListFoldersInput>;
export type ListFoldersToolParams = z.infer<typeof ListFoldersToolInput>;
export type ListEmailsParams = z.infer<typeof ListEmailsInput>;
export type SearchEmailsParams = z.infer<typeof SearchEmailsInput>;
export type GetEmailParams = z.infer<typeof GetEmailInput>;
export type GetEmailsParams = z.infer<typeof GetEmailsInput>;
export type GetUnreadCountParams = z.infer<typeof GetUnreadCountInput>;
export type ListAttachmentsParams = z.infer<typeof ListAttachmentsInput>;
export type DownloadAttachmentParams = z.infer<typeof DownloadAttachmentInput>;
export type CheckNewEmailsParams = z.infer<typeof CheckNewEmailsInput>;
export type SearchEmailsAdvancedParams = z.infer<typeof SearchEmailsAdvancedInput>;
export type ListConversationParams = z.infer<typeof ListConversationInput>;
export type GetMessageHeadersParams = z.infer<typeof GetMessageHeadersInput>;
export type GetMessageMimeParams = z.infer<typeof GetMessageMimeInput>;
export type GetMailTipsParams = z.infer<typeof GetMailTipsInput>;

// =============================================================================
// Parser Interface (for body content)
// =============================================================================

/**
 * Interface for reading email body content from data files. Implemented by the
 * Graph content reader.
 */
export interface IContentReader {
  /**
   * Reads the body content from the given data file path.
   * Returns null if the file cannot be read.
   */
  readEmailBody(dataFilePath: string | null): string | null;
}

// =============================================================================
// Registry Definitions (v3 registry-driven architecture)
// =============================================================================

/**
 * Registry tool definitions for the mail READ domain. Each handler delegates
 * to GraphMailTools, which returns MCP content directly.
 */
export const SendEmailInput = z.strictObject({
  to: z.array(z.string()).min(1).describe('Recipient email addresses'),
  subject: z.string().min(1).describe('Email subject'),
  body: z.string().describe('Email body content'),
  body_type: z.enum(['plain', 'html']).default('plain').describe('Body content type (default: plain)'),
  cc: z.array(z.string()).optional().describe('CC recipients'),
  bcc: z.array(z.string()).optional().describe('BCC recipients'),
  reply_to: z.string().optional().describe('Reply-to address'),
  attachments: z.array(z.strictObject({
    path: z.string().describe('Absolute file path to attachment'),
    name: z.string().optional().describe('Optional display name for the attachment'),
  })).optional().describe('File attachments'),
  inline_images: z.array(z.strictObject({
    path: z.string().describe('Absolute file path to the inline image'),
    content_id: z.string().describe('Content-ID referenced in the HTML body'),
  })).optional().describe('Inline images referenced by the HTML body'),
  account_id: z.number().int().positive().optional().describe('Optional account ID to send from'),
});
export type SendEmailParams = z.infer<typeof SendEmailInput>;

export function mailToolDefinitions(): ToolDefinition[] {
  return [
    defineTool({
      name: 'send_email',
      description: 'Send an email with optional CC, BCC, attachments, and HTML formatting. Returns the sent message ID and timestamp.',
      input: SendEmailInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['mail'],
      backends: ['graph'],
      // Direct send is not available on the Graph backend — clients use the
      // two-phase prepare_send_email/confirm_send_email.
      handler: () => ({ content: [{ type: 'text' as const, text: 'Direct send_email is not available on the Graph backend; use prepare_send_email then confirm_send_email.' }], isError: true }),
    }),
    defineTool({
      name: 'list_folders',
      description: 'List all mail folders with message and unread counts. Can filter by account.',
      input: ListFoldersToolInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'mailGraph').listFolders(params),
    }),
    defineTool({
      name: 'list_emails',
      description: 'List emails in a folder with pagination',
      input: ListEmailsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'mailGraph').listEmails(params),
    }),
    defineTool({
      name: 'search_emails',
      description: 'Search emails by subject, sender, or content',
      input: SearchEmailsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'mailGraph').searchEmails(params),
    }),
    defineTool({
      name: 'get_email',
      description: 'Get full email details including body',
      input: GetEmailInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'mailGraph').getEmail(params),
    }),
    defineTool({
      name: 'get_emails',
      description: 'Get multiple emails by ID in a single call (max 25). Useful for batch operations or summarizing threads.',
      input: GetEmailsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'mailGraph').getEmails(params),
    }),
    defineTool({
      name: 'get_unread_count',
      description: 'Get unread email count',
      input: GetUnreadCountInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'mailGraph').getUnreadCount(params),
    }),
    defineTool({
      name: 'list_attachments',
      description: 'List attachment metadata (name, size, type) for an email',
      input: ListAttachmentsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'mailGraph').listAttachments(params),
    }),
    defineTool({
      name: 'download_attachment',
      description: 'Download/save an email attachment to a file on disk. Returns the saved file path and size.',
      input: DownloadAttachmentInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'mailGraph').downloadAttachment(params),
    }),
    // ---- Graph-only reads ----
    defineTool({
      name: 'search_emails_advanced',
      description: 'Search emails with structured criteria (from, to, subject_contains, body_contains, text, received_after/before, has_attachments, is_unread, importance). Compiled server-side to the correct Graph query — no KQL syntax to get wrong. Dates are ISO (YYYY-MM-DD); mixed criteria use day-granular dates. Provide at least one criterion. (Graph API)',
      input: SearchEmailsAdvancedInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'mailGraph').searchEmailsAdvanced(params),
    }),
    defineTool({
      name: 'check_new_emails',
      description: 'Check for new or changed emails since last check using delta sync. First call returns recent messages (initial sync). Subsequent calls return only new/changed messages.',
      input: CheckNewEmailsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'mailGraph').checkNewEmails(params),
    }),
    defineTool({
      name: 'list_conversation',
      description: 'List all messages in an email conversation/thread, ordered chronologically. Provide any message ID from the thread.',
      input: ListConversationInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'mailGraph').listConversation(params),
    }),
    defineTool({
      name: 'get_message_headers',
      description: 'Get internet message headers (SPF, DKIM, routing, etc.) for an email (Graph API)',
      input: GetMessageHeadersInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'mailGraph').getMessageHeaders(params),
    }),
    defineTool({
      name: 'get_message_mime',
      description: 'Download the full MIME content (.eml) of an email to a local file (Graph API)',
      input: GetMessageMimeInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'mailGraph').getMessageMime(params),
    }),
    defineTool({
      name: 'get_mail_tips',
      description: 'Get mail tips (automatic replies, mailbox full, delivery restrictions, max message size) for email addresses (Graph API)',
      input: GetMailTipsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['mail'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'mailGraph').getMailTips(params),
    }),
  ];
}
