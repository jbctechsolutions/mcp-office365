/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Shared-mailbox / delegate-access MCP tools (#40).
 *
 * Every other domain scopes its Graph calls to `/me/...`. These tools target
 * another user's mailbox, calendar, or drive via `/users/{upn}/...`, relying on
 * the signed-in user having delegate / shared access (or the delegated
 * `Mail.Read.Shared` / `Calendars.Read.Shared` / `Files.Read.All` scopes). They
 * are READ-ONLY (issue #40: start read-only mail + calendar + files, expand
 * later) and self-contained: they hit Graph directly rather than the `/me`
 * local mirror, and return RAW Graph ids.
 *
 * Why raw ids (not durable tokens): durable tokens are `/me`- and
 * account-scoped (D7) — the same message in a shared mailbox has a different
 * Graph id than it would under `/me`, so a token minted for `/me` cannot
 * address another mailbox. Each `get_*` tool therefore takes the raw id echoed
 * by its `list_*` / `search_*` counterpart plus the `mailbox` it came from.
 * Passing a durable token here is rejected with a clear validation error.
 */

import { z } from 'zod';
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset } from '../registry/context.js';
import { isToken } from '../ids/token.js';
import { ValidationError } from '../utils/errors.js';
import type { ToolContext, ToolDefinition, ToolResult } from '../registry/types.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    sharedMailbox: SharedMailboxTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

const mailbox = z
  .string()
  .min(1)
  .describe("Target mailbox: the other user's email address (UPN) or user ID you have shared/delegate access to");

export const ListSharedMailboxFoldersInput = z.strictObject({
  mailbox,
});

export const ListSharedMailboxEmailsInput = z.strictObject({
  mailbox,
  folder_id: z.string().min(1).optional().describe('Raw Graph folder ID from list_shared_mailbox_folders (defaults to the whole mailbox)'),
  limit: z.number().int().min(1).max(100).optional().describe('Max messages to return (default 25, max 100)'),
  unread_only: z.boolean().optional().describe('Only return unread messages (default false)'),
});

export const GetSharedMailboxEmailInput = z.strictObject({
  mailbox,
  email_id: z.string().min(1).describe('Raw Graph message ID from list_shared_mailbox_emails / search_shared_mailbox_emails'),
  include_body: z.boolean().optional().describe('Include the full message body (default false)'),
  strip_html: z.boolean().optional().describe('Strip HTML from the body when include_body is true (default false)'),
});

export const SearchSharedMailboxEmailsInput = z.strictObject({
  mailbox,
  query: z.string().min(1).describe('Free-text search over the shared mailbox (Graph $search)'),
  limit: z.number().int().min(1).max(100).optional().describe('Max messages to return (default 25, max 100)'),
});

export const ListSharedCalendarEventsInput = z.strictObject({
  mailbox,
  start: z.string().min(1).optional().describe('ISO 8601 start of the window (e.g. 2026-07-01T00:00:00Z); requires end'),
  end: z.string().min(1).optional().describe('ISO 8601 end of the window; requires start'),
  limit: z.number().int().min(1).max(100).optional().describe('Max events to return (default 25, max 100)'),
});

export const GetSharedCalendarEventInput = z.strictObject({
  mailbox,
  event_id: z.string().min(1).describe('Raw Graph event ID from list_shared_calendar_events'),
});

export const ListSharedUserDriveItemsInput = z.strictObject({
  mailbox,
  item_id: z.string().min(1).optional().describe('Raw Graph driveItem ID to list children of (defaults to the drive root)'),
});

export const SearchSharedUserDriveItemsInput = z.strictObject({
  mailbox,
  query: z.string().min(1).describe('Free-text search over the shared user drive'),
  limit: z.number().int().min(1).max(100).optional().describe('Max items to return (default 25, max 100)'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListSharedMailboxFoldersParams = z.infer<typeof ListSharedMailboxFoldersInput>;
export type ListSharedMailboxEmailsParams = z.infer<typeof ListSharedMailboxEmailsInput>;
export type GetSharedMailboxEmailParams = z.infer<typeof GetSharedMailboxEmailInput>;
export type SearchSharedMailboxEmailsParams = z.infer<typeof SearchSharedMailboxEmailsInput>;
export type ListSharedCalendarEventsParams = z.infer<typeof ListSharedCalendarEventsInput>;
export type GetSharedCalendarEventParams = z.infer<typeof GetSharedCalendarEventInput>;
export type ListSharedUserDriveItemsParams = z.infer<typeof ListSharedUserDriveItemsInput>;
export type SearchSharedUserDriveItemsParams = z.infer<typeof SearchSharedUserDriveItemsInput>;

// =============================================================================
// Client Interface (implemented by GraphClient)
// =============================================================================

interface GraphRecipient {
  emailAddress?: { name?: string | null; address?: string | null } | null;
}

interface GraphMessage {
  id?: string | null;
  subject?: string | null;
  from?: GraphRecipient | null;
  toRecipients?: GraphRecipient[] | null;
  ccRecipients?: GraphRecipient[] | null;
  receivedDateTime?: string | null;
  sentDateTime?: string | null;
  isRead?: boolean | null;
  hasAttachments?: boolean | null;
  importance?: string | null;
  bodyPreview?: string | null;
  conversationId?: string | null;
  parentFolderId?: string | null;
  body?: { contentType?: string | null; content?: string | null } | null;
}

interface GraphMailFolder {
  id?: string | null;
  displayName?: string | null;
  parentFolderId?: string | null;
  totalItemCount?: number | null;
  unreadItemCount?: number | null;
}

interface GraphEventLike {
  id?: string | null;
  subject?: string | null;
  start?: { dateTime?: string | null; timeZone?: string | null } | null;
  end?: { dateTime?: string | null; timeZone?: string | null } | null;
  location?: { displayName?: string | null } | null;
  isAllDay?: boolean | null;
  organizer?: GraphRecipient | null;
  attendees?: Array<{ emailAddress?: { name?: string | null; address?: string | null } | null; status?: { response?: string | null } | null }> | null;
  bodyPreview?: string | null;
  body?: { contentType?: string | null; content?: string | null } | null;
}

export interface ISharedMailboxClient {
  listSharedMailFolders(mailbox: string): Promise<GraphMailFolder[]>;
  listSharedMessages(mailbox: string, folderId?: string, limit?: number, unreadOnly?: boolean): Promise<GraphMessage[]>;
  getSharedMessage(mailbox: string, messageId: string, includeBody?: boolean): Promise<GraphMessage>;
  searchSharedMessages(mailbox: string, query: string, limit?: number): Promise<GraphMessage[]>;
  listSharedEvents(mailbox: string, limit?: number, startDate?: Date, endDate?: Date): Promise<GraphEventLike[]>;
  getSharedEvent(mailbox: string, eventId: string): Promise<GraphEventLike>;
  listSharedDriveItems(mailbox: string, itemId?: string): Promise<Array<Record<string, unknown>>>;
  searchSharedDriveItems(mailbox: string, query: string, limit?: number): Promise<Array<Record<string, unknown>>>;
}

// =============================================================================
// Helpers
// =============================================================================

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Rejects a durable token where a raw Graph id is required. A durable token is
 * `/me`- and account-scoped and cannot address another mailbox's item, so
 * failing loudly here is clearer than a downstream Graph 404.
 */
function requireRawGraphId(value: string, field: string): void {
  if (isToken(value)) {
    throw new ValidationError(
      `${field} must be the raw Graph id returned by the shared-mailbox list/search tools, not a durable token. Durable tokens are scoped to your own (/me) mailbox and cannot address a shared mailbox.`,
    );
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

function mapRecipient(r?: GraphRecipient | null): string | null {
  return r?.emailAddress?.address ?? null;
}

function mapMessage(m: GraphMessage): {
  id: string | null;
  subject: string | null;
  from: string | null;
  to: string[];
  cc: string[];
  receivedDateTime: string | null;
  sentDateTime: string | null;
  isRead: boolean;
  hasAttachments: boolean;
  importance: string | null;
  preview: string | null;
  conversationId: string | null;
} {
  return {
    id: m.id ?? null,
    subject: m.subject ?? null,
    from: mapRecipient(m.from),
    to: (m.toRecipients ?? []).map(mapRecipient).filter((a): a is string => a != null),
    cc: (m.ccRecipients ?? []).map(mapRecipient).filter((a): a is string => a != null),
    receivedDateTime: m.receivedDateTime ?? null,
    sentDateTime: m.sentDateTime ?? null,
    isRead: m.isRead ?? false,
    hasAttachments: m.hasAttachments ?? false,
    importance: m.importance ?? null,
    preview: m.bodyPreview ?? null,
    conversationId: m.conversationId ?? null,
  };
}

function mapEvent(e: GraphEventLike): {
  id: string | null;
  subject: string | null;
  start: string | null;
  end: string | null;
  location: string | null;
  isAllDay: boolean;
  organizer: string | null;
  attendees: string[];
  preview: string | null;
} {
  return {
    id: e.id ?? null,
    subject: e.subject ?? null,
    start: e.start?.dateTime ?? null,
    end: e.end?.dateTime ?? null,
    location: e.location?.displayName ?? null,
    isAllDay: e.isAllDay ?? false,
    organizer: e.organizer?.emailAddress?.address ?? null,
    attendees: (e.attendees ?? [])
      .map((a) => a.emailAddress?.address)
      .filter((a): a is string => a != null),
    preview: e.bodyPreview ?? null,
  };
}

function mapDriveItem(item: Record<string, unknown>): {
  id: string | null;
  name: string | null;
  size: number | null;
  webUrl: string | null;
  lastModifiedDateTime: string | null;
  isFolder: boolean;
} {
  return {
    id: (item.id as string | undefined) ?? null,
    name: (item.name as string | undefined) ?? null,
    size: (item.size as number | undefined) ?? null,
    webUrl: (item.webUrl as string | undefined) ?? null,
    lastModifiedDateTime: (item.lastModifiedDateTime as string | undefined) ?? null,
    isFolder: item.folder != null,
  };
}

// =============================================================================
// Shared Mailbox Tools
// =============================================================================

export class SharedMailboxTools {
  constructor(private readonly client: ISharedMailboxClient) {}

  async listFolders(params: ListSharedMailboxFoldersParams): Promise<ToolResult> {
    const folders = await this.client.listSharedMailFolders(params.mailbox);
    return jsonResult({
      mailbox: params.mailbox,
      folders: folders.map((f) => ({
        id: f.id ?? null,
        name: f.displayName ?? null,
        parentFolderId: f.parentFolderId ?? null,
        totalItemCount: f.totalItemCount ?? 0,
        unreadItemCount: f.unreadItemCount ?? 0,
      })),
    });
  }

  async listEmails(params: ListSharedMailboxEmailsParams): Promise<ToolResult> {
    if (params.folder_id != null) requireRawGraphId(params.folder_id, 'folder_id');
    const messages = await this.client.listSharedMessages(
      params.mailbox,
      params.folder_id,
      params.limit ?? 25,
      params.unread_only ?? false,
    );
    return jsonResult({ mailbox: params.mailbox, emails: messages.map(mapMessage) });
  }

  async getEmail(params: GetSharedMailboxEmailParams): Promise<ToolResult> {
    requireRawGraphId(params.email_id, 'email_id');
    const message = await this.client.getSharedMessage(params.mailbox, params.email_id, params.include_body ?? false);
    const base = mapMessage(message);
    let body: string | null = null;
    if (params.include_body === true) {
      body = message.body?.content ?? null;
      if (params.strip_html === true && body != null) body = stripHtml(body);
    }
    return jsonResult({ mailbox: params.mailbox, ...base, body });
  }

  async searchEmails(params: SearchSharedMailboxEmailsParams): Promise<ToolResult> {
    const messages = await this.client.searchSharedMessages(params.mailbox, params.query, params.limit ?? 25);
    return jsonResult({ mailbox: params.mailbox, emails: messages.map(mapMessage) });
  }

  async listEvents(params: ListSharedCalendarEventsParams): Promise<ToolResult> {
    const hasStart = params.start != null;
    const hasEnd = params.end != null;
    if (hasStart !== hasEnd) {
      throw new ValidationError('start and end must be provided together to query a calendar window.');
    }
    let startDate: Date | undefined;
    let endDate: Date | undefined;
    if (hasStart && hasEnd) {
      startDate = new Date(params.start!);
      endDate = new Date(params.end!);
      if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime())) {
        throw new ValidationError('start and end must be valid ISO 8601 date-times.');
      }
      if (startDate.getTime() > endDate.getTime()) {
        throw new ValidationError('start must be on or before end.');
      }
    }
    const events = await this.client.listSharedEvents(params.mailbox, params.limit ?? 25, startDate, endDate);
    return jsonResult({ mailbox: params.mailbox, events: events.map(mapEvent) });
  }

  async getEvent(params: GetSharedCalendarEventParams): Promise<ToolResult> {
    requireRawGraphId(params.event_id, 'event_id');
    const event = await this.client.getSharedEvent(params.mailbox, params.event_id);
    const base = mapEvent(event);
    const body = event.body?.content ?? null;
    return jsonResult({ mailbox: params.mailbox, ...base, body });
  }

  async listDriveItems(params: ListSharedUserDriveItemsParams): Promise<ToolResult> {
    if (params.item_id != null) requireRawGraphId(params.item_id, 'item_id');
    const items = await this.client.listSharedDriveItems(params.mailbox, params.item_id);
    return jsonResult({ mailbox: params.mailbox, items: items.map(mapDriveItem) });
  }

  async searchDriveItems(params: SearchSharedUserDriveItemsParams): Promise<ToolResult> {
    const items = await this.client.searchSharedDriveItems(params.mailbox, params.query, params.limit ?? 25);
    return jsonResult({ mailbox: params.mailbox, items: items.map(mapDriveItem) });
  }
}

// =============================================================================
// Registry Definitions
// =============================================================================

export function sharedMailboxToolDefinitions(): ToolDefinition[] {
  const tools = (ctx: ToolContext): SharedMailboxTools => requireGraphToolset(ctx, 'sharedMailbox');

  return [
    defineTool({
      name: 'list_shared_mailbox_folders',
      description: 'List mail folders in a shared/delegated mailbox (Graph API, /users/{upn}/mailFolders). Read-only.',
      input: ListSharedMailboxFoldersInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['shared', 'mail'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listFolders(params),
    }),
    defineTool({
      name: 'list_shared_mailbox_emails',
      description: 'List emails in a shared/delegated mailbox, optionally within a folder (Graph API). Read-only. Returns raw Graph message ids.',
      input: ListSharedMailboxEmailsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['shared', 'mail'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listEmails(params),
    }),
    defineTool({
      name: 'get_shared_mailbox_email',
      description: 'Get a single email from a shared/delegated mailbox by raw Graph message id (Graph API). Read-only.',
      input: GetSharedMailboxEmailInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['shared', 'mail'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getEmail(params),
    }),
    defineTool({
      name: 'search_shared_mailbox_emails',
      description: 'Full-text search a shared/delegated mailbox (Graph $search). Read-only. Returns raw Graph message ids.',
      input: SearchSharedMailboxEmailsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['shared', 'mail'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).searchEmails(params),
    }),
    defineTool({
      name: 'list_shared_calendar_events',
      description: "List events from another user's calendar, optionally within a start/end window (Graph API). Read-only. Returns raw Graph event ids.",
      input: ListSharedCalendarEventsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['shared', 'calendar'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listEvents(params),
    }),
    defineTool({
      name: 'get_shared_calendar_event',
      description: "Get a single event from another user's calendar by raw Graph event id (Graph API). Read-only.",
      input: GetSharedCalendarEventInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['shared', 'calendar'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getEvent(params),
    }),
    defineTool({
      name: 'list_shared_user_drive_items',
      description: "List items in another user's OneDrive (root or a folder's children) (Graph API). Read-only. Returns raw Graph driveItem ids.",
      input: ListSharedUserDriveItemsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['shared', 'files'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listDriveItems(params),
    }),
    defineTool({
      name: 'search_shared_user_drive_items',
      description: "Search another user's OneDrive (Graph API). Read-only. Returns raw Graph driveItem ids.",
      input: SearchSharedUserDriveItemsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['shared', 'files'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).searchDriveItems(params),
    }),
  ];
}
