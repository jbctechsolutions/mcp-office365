/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Microsoft Graph API client wrapper.
 *
 * Provides a typed interface to the Graph API with:
 * - Automatic token management
 * - Response caching
 * - Pagination support
 * - Error handling
 */

import 'isomorphic-fetch';
import {
  Client,
  ResponseType,
  AuthenticationHandler,
  RetryHandler,
  RetryHandlerOptions,
  RedirectHandler,
  HTTPMessageHandler,
  type AuthenticationProvider,
  type ShouldRetry,
  type PageCollection,
} from '@microsoft/microsoft-graph-client';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { getAccessToken, type DeviceCodeCallback } from '../auth/index.js';
import { type BatchRequest, type BatchResponseItem, buildBatchPayload, splitIntoBatches, parseBatchResponse } from './batch.js';
import { ResponseCache, CacheTTL, createCacheKey } from './cache.js';
import { ImmutableIdMiddleware } from './immutable-id-middleware.js';

/** Generic shape for a Graph API response containing a `.value` array. */
interface GraphCollectionResponse<T> { value: T[] }

// The SDK's RetryHandler only considers 429/503/504 retriable (its private
// static RETRY_STATUS_CODES), gating `this.isRetry(response)` BEFORE our
// shouldRetry callback runs — so a callback can only *narrow* that set, never
// widen it. D5 also wants 502 Bad Gateway retried, so we widen the SDK's set
// once at load. Our shouldRetry still narrows to idempotent GET/HEAD/OPTIONS,
// and this server builds only its own client, so the shared-static mutation
// has no cross-tenant effect.
const RETRY_STATUS_CODES = (RetryHandler as unknown as { RETRY_STATUS_CODES: number[] })
  .RETRY_STATUS_CODES;
if (Array.isArray(RETRY_STATUS_CODES) && !RETRY_STATUS_CODES.includes(502)) {
  RETRY_STATUS_CODES.push(502);
}

/**
 * D5 retry policy (exported for testing). Retries only idempotent reads on
 * transient failures — 429 plus 502/503/504 — honoring Retry-After via the
 * SDK-computed delay. NEVER retries a send/write once the body is on the wire:
 * an ambiguous 429/5xx on a POST could double-send. Paired with the module-load
 * widening above (the SDK gates 502 out by default), this makes 502 genuinely
 * retriable for idempotent reads while leaving writes/sendMail untouched.
 */
export function shouldRetryGraphRequest(
  method: string,
  url: string,
  status: number
): boolean {
  const m = method.toUpperCase();
  if (m !== 'GET' && m !== 'HEAD' && m !== 'OPTIONS') {
    return false;
  }
  // Never retry an OData action (POST-shaped operation like /sendMail, /reply,
  // /createReply of writes) even if the transport surfaces it as a GET. Anchor
  // the action to the end of the path (or a query string) so a drive item
  // literally named e.g. "reply.docx" read via GET is still retriable.
  if (/\/(sendMail|reply|forward|createReply|createForward)(?:\?|$)/i.test(url)) {
    return false;
  }
  return status === 429 || status === 502 || status === 503 || status === 504;
}

/** Generic Graph API entity (untyped). */
type GraphEntity = Record<string, unknown>;

/** Fields selected for message search results (shared across search mechanisms). */
const MESSAGE_SEARCH_SELECT =
  'id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId,parentFolderId';

/** Shape of a `POST /search/query` response (only the parts we read). */
interface SearchQueryResponse {
  value?: Array<{
    hitsContainers?: Array<{
      hits?: Array<{ resource: MicrosoftGraph.Message }>;
    }>;
  }>;
}

/**
 * Graph client wrapper with caching and token management.
 */
export class GraphClient {
  private client: Client | null = null;
  private readonly cache = new ResponseCache();
  private readonly deviceCodeCallback: DeviceCodeCallback | undefined;
  private readonly tokenProvider: (() => Promise<string>) | undefined;

  /**
   * @param deviceCodeCallback stdio device-code flow (default backend).
   * @param tokenProvider remote mode (U5): supplies a per-user Graph token
   *   (e.g. via On-Behalf-Of). When set, it replaces the process-global
   *   device-code token acquisition for this client instance.
   */
  constructor(deviceCodeCallback?: DeviceCodeCallback, tokenProvider?: () => Promise<string>) {
    this.deviceCodeCallback = deviceCodeCallback;
    this.tokenProvider = tokenProvider;
  }

  /**
   * Gets or creates the Graph client instance.
   */
  // eslint-disable-next-line @typescript-eslint/require-await
  private async getClient(): Promise<Client> {
    if (this.client == null) {
      const authProvider: AuthenticationProvider = {
        getAccessToken: () =>
          this.tokenProvider != null
            ? this.tokenProvider()
            : getAccessToken(this.deviceCodeCallback),
      };

      const shouldRetry: ShouldRetry = (_delay, _attempt, request, options, response) => {
        const method = (options?.method ?? 'GET').toString();
        // The SDK types `request` as RequestInfo, which resolves to `any` under
        // this repo's lib config; cast to the shape we read to stay lint-clean.
        const url = typeof request === 'string' ? request : String((request as { url?: unknown }).url ?? '');
        return shouldRetryGraphRequest(method, url, response?.status ?? 0);
      };

      // Chain mirrors the SDK default order (auth → retry → redirect → http),
      // minus the cosmetic TelemetryHandler, plus the immutable-ID preference
      // (U5b-3) sitting right after auth so its Prefer header rides every retry
      // attempt without being re-appended. RedirectHandler matters: the binary
      // download endpoints (/content, /photo/$value, meeting recordings) 302 to
      // a pre-authenticated CDN URL, so the redirect must be followed explicitly
      // rather than relying on the underlying fetch default.
      const auth = new AuthenticationHandler(authProvider);
      const immutableId = new ImmutableIdMiddleware();
      const retry = new RetryHandler(new RetryHandlerOptions(undefined, undefined, shouldRetry));
      const redirect = new RedirectHandler();
      const http = new HTTPMessageHandler();
      auth.setNext(immutableId);
      immutableId.setNext(retry);
      retry.setNext(redirect);
      redirect.setNext(http);

      this.client = Client.initWithMiddleware({ middleware: auth });
    }
    return this.client;
  }

  /**
   * Clears the response cache.
   */
  clearCache(): void {
    this.cache.clear();
  }

  // ===========================================================================
  // Mail Folders
  // ===========================================================================

  /**
   * Lists all mail folders.
   */
  async listMailFolders(): Promise<MicrosoftGraph.MailFolder[]> {
    const cacheKey = createCacheKey('listMailFolders');
    const cached = this.cache.get<MicrosoftGraph.MailFolder[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();
    const result: MicrosoftGraph.MailFolder[] = [];

    // Get top-level folders with pagination
    let response = await client
      .api('/me/mailFolders')
      .select('id,displayName,parentFolderId,totalItemCount,unreadItemCount')
      .top(100)
      .get() as PageCollection;

    result.push(...(response.value as MicrosoftGraph.MailFolder[]));

    // Handle pagination
    while (response['@odata.nextLink'] != null) {
      response = await client.api(response['@odata.nextLink']).get() as PageCollection;
      result.push(...(response.value as MicrosoftGraph.MailFolder[]));
    }

    // Also get child folders (one level deep)
    for (const folder of [...result]) {
      try {
        const children = await client
          .api(`/me/mailFolders/${folder.id}/childFolders`)
          .select('id,displayName,parentFolderId,totalItemCount,unreadItemCount')
          .get() as PageCollection;

        result.push(...(children.value as MicrosoftGraph.MailFolder[]));
      } catch {
        // Some folders may not have children or may not be accessible
      }
    }

    this.cache.set(cacheKey, result, CacheTTL.FOLDERS);
    return result;
  }

  /**
   * Gets a specific mail folder by ID.
   */
  async getMailFolder(folderId: string): Promise<MicrosoftGraph.MailFolder | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/mailFolders/${folderId}`)
        .select('id,displayName,parentFolderId,totalItemCount,unreadItemCount')
        .get() as MicrosoftGraph.MailFolder;
    } catch {
      return null;
    }
  }

  // ===========================================================================
  // Messages (Emails)
  // ===========================================================================

  /**
   * Lists messages in a folder with pagination.
   */
  async listMessages(
    folderId: string,
    limit: number = 50,
    skip: number = 0
  ): Promise<MicrosoftGraph.Message[]> {
    const cacheKey = createCacheKey('listMessages', folderId, limit, skip);
    const cached = this.cache.get<MicrosoftGraph.Message[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();

    const response = await client
      .api(`/me/mailFolders/${folderId}/messages`)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId')
      .orderby('receivedDateTime desc')
      .top(limit)
      .skip(skip)
      .get() as PageCollection;

    const result = response.value as MicrosoftGraph.Message[];
    this.cache.set(cacheKey, result, CacheTTL.EMAILS);
    return result;
  }

  /**
   * Lists unread messages in a folder.
   */
  async listUnreadMessages(
    folderId: string,
    limit: number = 50,
    skip: number = 0
  ): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();

    const response = await client
      .api(`/me/mailFolders/${folderId}/messages`)
      .filter('isRead eq false')
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId')
      .orderby('receivedDateTime desc')
      .top(limit)
      .skip(skip)
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Message[];
  }

  /**
   * Searches messages across all folders.
   */
  async searchMessages(query: string, limit: number = 50): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();

    const response = await client
      .api('/me/messages')
      .search(`"${query}"`)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId,parentFolderId')
      .top(limit)
      .get() as PageCollection;

    const messages = response.value as MicrosoftGraph.Message[];
    await this.upgradeMessageIdsToImmutable(messages);
    return messages;
  }

  /**
   * Searches messages in a specific folder.
   */
  async searchMessagesInFolder(
    folderId: string,
    query: string,
    limit: number = 50
  ): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();

    const response = await client
      .api(`/me/mailFolders/${folderId}/messages`)
      .search(`"${query}"`)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId')
      .top(limit)
      .get() as PageCollection;

    const messages = response.value as MicrosoftGraph.Message[];
    await this.upgradeMessageIdsToImmutable(messages);
    return messages;
  }

  /**
   * Property-only structured search (U7 / D9): `$filter` on messages. The filter
   * string is built server-side by the search compiler (validated OData syntax).
   */
  async searchMessagesFilter(filter: string, limit: number = 50): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();
    // No $orderby: Graph returns InefficientFilter when a sender/flag $filter is
    // combined with $orderby (D9 spike), and /me/messages already defaults to
    // receivedDateTime desc, so results stay recent-first without it.
    const response = await client
      .api('/me/messages')
      .filter(filter)
      .select(MESSAGE_SEARCH_SELECT)
      .top(limit)
      .get() as PageCollection;
    return response.value as MicrosoftGraph.Message[];
  }

  /**
   * Free-text structured search (U7 / D9): quoted `$search` on messages. The
   * search value is built by the compiler already correctly quoted (e.g.
   * `"subject:report"`), which the D9 spike confirmed property-scopes correctly.
   */
  async searchMessagesSearchValue(searchValue: string, limit: number = 50): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();
    const response = await client
      .api('/me/messages')
      .search(searchValue)
      .select(MESSAGE_SEARCH_SELECT)
      .top(limit)
      .get() as PageCollection;
    const messages = response.value as MicrosoftGraph.Message[];
    await this.upgradeMessageIdsToImmutable(messages);
    return messages;
  }

  /**
   * Mixed property + free-text structured search (U7 / D9): `POST /search/query`
   * with a server-built KQL string — the only single-request path when both are
   * present. Normalizes the hitsContainers response to a Message[].
   *
   * Limitation (verified in the D9 spike): /search/query populates from/subject/
   * receivedDateTime/isRead/hasAttachments/parentFolderId, but NOT toRecipients
   * or flag, regardless of the requested fields — so mixed-mode results carry no
   * recipient list or flag. Acceptable for a result list (sender/subject/date
   * are present); the property-only and free-text-only paths are unaffected.
   */
  async searchMessagesQuery(kql: string, limit: number = 50): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();
    const body = {
      requests: [
        {
          entityTypes: ['message'],
          query: { queryString: kql },
          from: 0,
          size: limit,
          fields: MESSAGE_SEARCH_SELECT.split(','),
        },
      ],
    };
    const response = (await client.api('/search/query').post(body)) as SearchQueryResponse;
    const hits = response.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
    // Guard against a hit without a resource so no undefined enters the mapper.
    const messages = hits.map((h) => h.resource).filter((r): r is MicrosoftGraph.Message => r != null);
    await this.upgradeMessageIdsToImmutable(messages);
    return messages;
  }

  /**
   * Translates Exchange item IDs between formats (U5b-3 / D2). Used to upgrade
   * the mutable REST IDs that `$search` returns into immutable REST entry IDs, so
   * search-minted self-encoding tokens (`em_`) survive a later move. Returns a map
   * of source ID → immutable target ID; IDs that Graph could not translate are
   * simply absent (partial success is normal and handled by the caller).
   */
  async translateExchangeIds(
    inputIds: string[],
    sourceIdType: string = 'restId',
    targetIdType: string = 'restImmutableEntryId'
  ): Promise<Map<string, string>> {
    const out = new Map<string, string>();
    if (inputIds.length === 0) {
      return out;
    }
    const client = await this.getClient();
    // Graph caps translateExchangeIds at 1000 input IDs per call; chunk to stay
    // under it even though search result sets are far smaller.
    for (let i = 0; i < inputIds.length; i += 1000) {
      const chunk = inputIds.slice(i, i + 1000);
      const response = (await client.api('/me/translateExchangeIds').post({
        inputIds: chunk,
        sourceIdType,
        targetIdType,
      })) as { value?: Array<{ sourceId?: string; targetId?: string }> };
      for (const pair of response.value ?? []) {
        if (pair.sourceId != null && pair.sourceId.length > 0 && pair.targetId != null && pair.targetId.length > 0) {
          out.set(pair.sourceId, pair.targetId);
        }
      }
    }
    return out;
  }

  /**
   * Rewrites the mutable `$search`-minted IDs on a message list to their
   * immutable equivalents in place (U5b-3 / D2), so the mapper mints durable
   * `em_` tokens. Graceful degradation: if translation fails wholesale (throw) or
   * for individual IDs (absent from the map), those messages keep their mutable
   * ID — the `em_` token still resolves this session, it just isn't move-durable.
   * Returns the number of IDs left un-upgraded (the caller's `degraded_ids` count).
   */
  private async upgradeMessageIdsToImmutable(messages: MicrosoftGraph.Message[]): Promise<number> {
    const mutableIds = messages
      .map((m) => m.id)
      .filter((id): id is string => typeof id === 'string' && id.length > 0);
    if (mutableIds.length === 0) {
      return 0;
    }

    let translated: Map<string, string>;
    try {
      translated = await this.translateExchangeIds(mutableIds);
    } catch {
      // Whole-batch failure (throttled/unavailable): leave every ID mutable.
      return mutableIds.length;
    }

    let degraded = 0;
    for (const message of messages) {
      if (typeof message.id !== 'string' || message.id.length === 0) {
        continue;
      }
      const immutable = translated.get(message.id);
      if (immutable != null) {
        message.id = immutable;
      } else {
        degraded += 1;
      }
    }
    return degraded;
  }

  /**
   * Lists messages in a conversation by conversationId.
   */
  async listConversationMessages(
    conversationId: string,
    limit: number = 50
  ): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();
    const response = await client
      .api('/me/messages')
      .filter(`conversationId eq '${conversationId}'`)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId,parentFolderId')
      .orderby('receivedDateTime asc')
      .top(limit)
      .get() as PageCollection;
    return response.value as MicrosoftGraph.Message[];
  }

  /**
   * Gets a specific message with full body.
   */
  async getMessage(messageId: string): Promise<MicrosoftGraph.Message | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/messages/${messageId}`)
        .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,body,bodyPreview,conversationId,internetMessageId,parentFolderId')
        .get() as MicrosoftGraph.Message;
    } catch {
      return null;
    }
  }

  /**
   * Gets message delta for incremental sync.
   */
  async getMessagesDelta(
    folderId: string,
    deltaLink?: string
  ): Promise<{ messages: MicrosoftGraph.Message[]; deltaLink: string }> {
    const client = await this.getClient();
    let page: PageCollection;

    if (deltaLink != null) {
      page = await client.api(deltaLink).get() as PageCollection;
    } else {
      page = await client
        .api(`/me/mailFolders/${folderId}/messages/delta`)
        .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId,parentFolderId')
        .top(50)
        .get() as PageCollection;
    }

    const messages: MicrosoftGraph.Message[] = [...((page.value as MicrosoftGraph.Message[] | undefined) ?? [])];

    // The deltaLink lands on the LAST page; paging through nextLinks and reading
    // it from the first response would yield '' whenever the delta spans >1 page.
    let nextLink = page['@odata.nextLink'];
    while (nextLink != null) {
      page = await client.api(nextLink).get() as PageCollection;
      messages.push(...((page.value as MicrosoftGraph.Message[] | undefined) ?? []));
      nextLink = page['@odata.nextLink'];
    }

    const newDeltaLink = page['@odata.deltaLink'] ?? '';
    return { messages, deltaLink: newDeltaLink };
  }

  /**
   * Gets calendar-view delta for incremental event sync (U12).
   *
   * The v1.0 event delta is served by `/me/calendarView/delta`, bounded by a
   * start/end window that is baked into the returned deltaLink — subsequent
   * rounds just follow that link. `$select` is unsupported here, so full event
   * objects come back; deletes arrive as `@removed` entries. The deltaLink lands
   * on the *last* page, so it is read after paging (not from the first response).
   */
  async getCalendarViewDelta(
    startDateTime: string,
    endDateTime: string,
    deltaLink?: string
  ): Promise<{ events: MicrosoftGraph.Event[]; deltaLink: string }> {
    const client = await this.getClient();
    let page: PageCollection;

    if (deltaLink != null) {
      page = await client.api(deltaLink).get() as PageCollection;
    } else {
      const start = encodeURIComponent(startDateTime);
      const end = encodeURIComponent(endDateTime);
      page = await client
        .api(`/me/calendarView/delta?startDateTime=${start}&endDateTime=${end}`)
        .header('Prefer', 'odata.maxpagesize=50')
        .get() as PageCollection;
    }

    const events: MicrosoftGraph.Event[] = [...((page.value as MicrosoftGraph.Event[] | undefined) ?? [])];

    let nextLink = page['@odata.nextLink'];
    while (nextLink != null) {
      page = await client.api(nextLink).get() as PageCollection;
      events.push(...((page.value as MicrosoftGraph.Event[] | undefined) ?? []));
      nextLink = page['@odata.nextLink'];
    }

    const newDeltaLink = page['@odata.deltaLink'] ?? '';
    return { events, deltaLink: newDeltaLink };
  }

  // ===========================================================================
  // Mail Rules
  // ===========================================================================

  /**
   * Lists all inbox mail rules.
   */
  async listMailRules(): Promise<MicrosoftGraph.MessageRule[]> {
    const client = await this.getClient();
    const response = await client
      .api('/me/mailFolders/inbox/messageRules')
      .get() as PageCollection;
    return response.value as MicrosoftGraph.MessageRule[];
  }

  /**
   * Creates a new inbox mail rule.
   */
  async createMailRule(rule: Record<string, unknown>): Promise<MicrosoftGraph.MessageRule> {
    const client = await this.getClient();
    const result = await client
      .api('/me/mailFolders/inbox/messageRules')
      .post(rule) as MicrosoftGraph.MessageRule;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes an inbox mail rule.
   */
  async deleteMailRule(ruleId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/mailFolders/inbox/messageRules/${ruleId}`)
      .delete();
    this.cache.clear();
  }

  // ===========================================================================
  // Automatic Replies (Out of Office)
  // ===========================================================================

  /**
   * Gets the automatic replies (OOF) settings.
   */
  async getAutomaticReplies(): Promise<Record<string, unknown>> {
    const client = await this.getClient();
    return await client.api('/me/mailboxSettings/automaticRepliesSetting').get() as Record<string, unknown>;
  }

  /**
   * Sets the automatic replies (OOF) settings.
   */
  async setAutomaticReplies(settings: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client.api('/me/mailboxSettings').patch({ automaticRepliesSetting: settings });
  }

  // ===========================================================================
  // Mailbox Settings
  // ===========================================================================

  /**
   * Gets the full mailbox settings for the current user.
   */
  async getMailboxSettings(): Promise<Record<string, unknown>> {
    const client = await this.getClient();
    return await client.api('/me/mailboxSettings').get() as Record<string, unknown>;
  }

  /**
   * Updates mailbox settings for the current user.
   */
  async updateMailboxSettings(settings: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client.api('/me/mailboxSettings').patch(settings);
  }

  // ===========================================================================
  // Calendars
  // ===========================================================================

  /**
   * Lists all calendars.
   */
  async listCalendars(): Promise<MicrosoftGraph.Calendar[]> {
    const cacheKey = createCacheKey('listCalendars');
    const cached = this.cache.get<MicrosoftGraph.Calendar[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();

    const response = await client
      .api('/me/calendars')
      .select('id,name,color,isDefaultCalendar,canEdit')
      .get() as PageCollection;

    const result = response.value as MicrosoftGraph.Calendar[];
    this.cache.set(cacheKey, result, CacheTTL.FOLDERS);
    return result;
  }

  // ===========================================================================
  // Events
  // ===========================================================================

  /**
   * Lists events with optional date range.
   */
  async listEvents(
    limit: number = 50,
    calendarId?: string,
    startDate?: Date,
    endDate?: Date
  ): Promise<MicrosoftGraph.Event[]> {
    const client = await this.getClient();

    // If date range provided, use calendarView
    if (startDate != null && endDate != null) {
      const baseUrl = calendarId != null
        ? `/me/calendars/${calendarId}/calendarView`
        : '/me/calendarView';

      const response = await client
        .api(baseUrl)
        .query({
          startDateTime: startDate.toISOString(),
          endDateTime: endDate.toISOString(),
        })
        .select('id,subject,start,end,location,isAllDay,organizer,attendees,body,recurrence,iCalUId')
        .orderby('start/dateTime')
        .top(limit)
        .get() as PageCollection;

      return response.value as MicrosoftGraph.Event[];
    }

    // Otherwise, get upcoming events
    const baseUrl = calendarId != null
      ? `/me/calendars/${calendarId}/events`
      : '/me/events';

    const response = await client
      .api(baseUrl)
      .select('id,subject,start,end,location,isAllDay,organizer,attendees,body,recurrence,iCalUId')
      .orderby('start/dateTime')
      .top(limit)
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Event[];
  }

  /**
   * Gets a specific event.
   */
  async getEvent(eventId: string): Promise<MicrosoftGraph.Event | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/events/${eventId}`)
        .select('id,subject,start,end,location,isAllDay,organizer,attendees,body,recurrence,iCalUId')
        .get() as MicrosoftGraph.Event;
    } catch {
      return null;
    }
  }

  /**
   * Lists instances of a recurring event within a date range.
   */
  async listEventInstances(
    eventId: string,
    startDateTime: string,
    endDateTime: string
  ): Promise<MicrosoftGraph.Event[]> {
    const client = await this.getClient();
    const response = await client
      .api(`/me/events/${eventId}/instances`)
      .query({ startDateTime, endDateTime })
      .select('id,subject,start,end,location,isAllDay,isCancelled,organizer,recurrence,bodyPreview')
      .get() as PageCollection;
    return response.value as MicrosoftGraph.Event[];
  }

  // ===========================================================================
  // Calendar Write Operations
  // ===========================================================================

  /**
   * Creates a new calendar event.
   */
  async createEvent(
    event: Record<string, unknown>,
    calendarId?: string
  ): Promise<MicrosoftGraph.Event> {
    const client = await this.getClient();
    const url = calendarId != null
      ? `/me/calendars/${calendarId}/events`
      : '/me/events';

    const result = await client
      .api(url)
      .post(event) as MicrosoftGraph.Event;
    this.cache.clear();
    return result;
  }

  /**
   * Updates an existing calendar event.
   */
  async updateEvent(eventId: string, updates: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/events/${eventId}`)
      .patch(updates);
    this.cache.clear();
  }

  /**
   * Deletes a calendar event.
   */
  async deleteEvent(eventId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/events/${eventId}`)
      .delete();
    this.cache.clear();
  }

  /**
   * Responds to a calendar event invitation.
   */
  async respondToEvent(
    eventId: string,
    response: 'accept' | 'decline' | 'tentative',
    sendResponse: boolean,
    comment?: string
  ): Promise<void> {
    const client = await this.getClient();
    const actionMap: Record<string, string> = {
      accept: 'accept',
      decline: 'decline',
      tentative: 'tentativelyAccept',
    };
    const action = actionMap[response];

    await client
      .api(`/me/events/${eventId}/${action}`)
      .post({ sendResponse, comment: comment ?? '' });
    this.cache.clear();
  }

  // ===========================================================================
  // Contacts
  // ===========================================================================

  /**
   * Lists contacts with pagination.
   */
  async listContacts(limit: number = 50, skip: number = 0): Promise<MicrosoftGraph.Contact[]> {
    const cacheKey = createCacheKey('listContacts', limit, skip);
    const cached = this.cache.get<MicrosoftGraph.Contact[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();

    const response = await client
      .api('/me/contacts')
      .select('id,displayName,givenName,surname,middleName,nickName,companyName,jobTitle,department,emailAddresses,homePhones,businessPhones,mobilePhone,homeAddress,businessAddress,personalNotes')
      .orderby('displayName')
      .top(limit)
      .skip(skip)
      .get() as PageCollection;

    const result = response.value as MicrosoftGraph.Contact[];
    this.cache.set(cacheKey, result, CacheTTL.CONTACTS);
    return result;
  }

  /**
   * Searches contacts by display name.
   */
  async searchContacts(query: string, limit: number = 50): Promise<MicrosoftGraph.Contact[]> {
    const client = await this.getClient();

    const response = await client
      .api('/me/contacts')
      .filter(`contains(displayName,'${query}')`)
      .select('id,displayName,givenName,surname,middleName,nickName,companyName,jobTitle,department,emailAddresses,homePhones,businessPhones,mobilePhone,homeAddress,businessAddress,personalNotes')
      .top(limit)
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Contact[];
  }

  /**
   * Gets a specific contact.
   */
  async getContact(contactId: string): Promise<MicrosoftGraph.Contact | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/contacts/${contactId}`)
        .select('id,displayName,givenName,surname,middleName,nickName,companyName,jobTitle,department,emailAddresses,homePhones,businessPhones,mobilePhone,homeAddress,businessAddress,personalNotes')
        .get() as MicrosoftGraph.Contact;
    } catch {
      return null;
    }
  }

  // ===========================================================================
  // Contact Write Operations
  // ===========================================================================

  /**
   * Creates a new contact.
   */
  async createContact(contact: Record<string, unknown>): Promise<MicrosoftGraph.Contact> {
    const client = await this.getClient();
    const result = await client
      .api('/me/contacts')
      .post(contact) as MicrosoftGraph.Contact;
    this.cache.clear();
    return result;
  }

  /**
   * Updates an existing contact.
   */
  async updateContact(contactId: string, updates: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/contacts/${contactId}`)
      .patch(updates);
    this.cache.clear();
  }

  /**
   * Deletes a contact.
   */
  async deleteContact(contactId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/contacts/${contactId}`)
      .delete();
    this.cache.clear();
  }

  // ===========================================================================
  // Contact Photos
  // ===========================================================================

  /**
   * Gets the photo for a contact as raw binary data.
   */
  async getContactPhoto(contactId: string): Promise<ArrayBuffer> {
    const client = await this.getClient();
    return await client
      .api(`/me/contacts/${contactId}/photo/$value`)
      .responseType(ResponseType.ARRAYBUFFER).get() as ArrayBuffer;
  }

  /**
   * Sets or updates the photo for a contact.
   */
  async setContactPhoto(contactId: string, photoData: Buffer, contentType: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/contacts/${contactId}/photo/$value`)
      .header('Content-Type', contentType)
      .put(photoData);
    this.cache.clear();
  }

  // ===========================================================================
  // Contact Folders
  // ===========================================================================

  /**
   * Lists all contact folders.
   */
  async listContactFolders(): Promise<MicrosoftGraph.ContactFolder[]> {
    const client = await this.getClient();
    const response = await client.api('/me/contactFolders').get() as PageCollection;
    return response.value as MicrosoftGraph.ContactFolder[];
  }

  /**
   * Creates a new contact folder.
   */
  async createContactFolder(displayName: string): Promise<MicrosoftGraph.ContactFolder> {
    const client = await this.getClient();
    const result = await client.api('/me/contactFolders').post({ displayName }) as MicrosoftGraph.ContactFolder;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes a contact folder.
   */
  async deleteContactFolder(folderId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/contactFolders/${folderId}`).delete();
    this.cache.clear();
  }

  /**
   * Lists contacts in a specific contact folder.
   */
  async listContactsInFolder(folderId: string, limit: number = 100): Promise<MicrosoftGraph.Contact[]> {
    const client = await this.getClient();
    const response = await client
      .api(`/me/contactFolders/${folderId}/contacts`)
      .select('id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,jobTitle,companyName')
      .top(limit)
      .get() as PageCollection;
    return response.value as MicrosoftGraph.Contact[];
  }

  // ===========================================================================
  // Tasks (Microsoft To Do)
  // ===========================================================================

  /**
   * Lists task lists.
   */
  async listTaskLists(): Promise<MicrosoftGraph.TodoTaskList[]> {
    const cacheKey = createCacheKey('listTaskLists');
    const cached = this.cache.get<MicrosoftGraph.TodoTaskList[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();

    const response = await client
      .api('/me/todo/lists')
      .select('id,displayName,wellknownListName')
      .get() as PageCollection;

    const result = response.value as MicrosoftGraph.TodoTaskList[];
    this.cache.set(cacheKey, result, CacheTTL.FOLDERS);
    return result;
  }

  /**
   * Lists tasks in a task list.
   */
  async listTasks(
    taskListId: string,
    limit: number = 50,
    skip: number = 0,
    includeCompleted: boolean = true
  ): Promise<MicrosoftGraph.TodoTask[]> {
    const client = await this.getClient();

    let api = client
      .api(`/me/todo/lists/${taskListId}/tasks`)
      .select('id,title,status,importance,dueDateTime,completedDateTime,body,createdDateTime,lastModifiedDateTime,isReminderOn,reminderDateTime')
      .top(limit)
      .skip(skip);

    if (!includeCompleted) {
      api = api.filter("status ne 'completed'");
    }

    const response = await api.get() as PageCollection;
    return response.value as MicrosoftGraph.TodoTask[];
  }

  /**
   * Lists all tasks across all task lists.
   */
  async listAllTasks(
    limit: number = 50,
    skip: number = 0,
    includeCompleted: boolean = true
  ): Promise<Array<MicrosoftGraph.TodoTask & { taskListId: string }>> {
    const taskLists = await this.listTaskLists();
    const allTasks: Array<MicrosoftGraph.TodoTask & { taskListId: string }> = [];

    for (const list of taskLists) {
      if (list.id == null) continue;

      const tasks = await this.listTasks(list.id, 100, 0, includeCompleted);

      for (const task of tasks) {
        allTasks.push({ ...task, taskListId: list.id });
      }
    }

    // Sort by due date, then slice for pagination
    allTasks.sort((a, b) => {
      if (a.dueDateTime == null && b.dueDateTime == null) return 0;
      if (a.dueDateTime == null) return 1;
      if (b.dueDateTime == null) return -1;
      return new Date(a.dueDateTime.dateTime ?? '').getTime() - new Date(b.dueDateTime.dateTime ?? '').getTime();
    });

    return allTasks.slice(skip, skip + limit);
  }

  /**
   * Gets a specific task.
   */
  async getTask(taskListId: string, taskId: string): Promise<MicrosoftGraph.TodoTask | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/todo/lists/${taskListId}/tasks/${taskId}`)
        .select('id,title,status,importance,dueDateTime,completedDateTime,body,createdDateTime,lastModifiedDateTime,isReminderOn,reminderDateTime')
        .get() as MicrosoftGraph.TodoTask;
    } catch {
      return null;
    }
  }

  // ===========================================================================
  // Write Operations
  // ===========================================================================

  /**
   * Moves a message to a different folder.
   */
  async moveMessage(messageId: string, destinationFolderId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: destinationFolderId });
    this.cache.clear(); // Invalidate cache after mutation
  }

  /**
   * Deletes a message (moves to Deleted Items).
   */
  async deleteMessage(messageId: string): Promise<void> {
    const client = await this.getClient();
    // Move to deletedItems well-known folder
    await client
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: 'deleteditems' });
    this.cache.clear();
  }

  /**
   * Archives a message (moves to Archive folder).
   */
  async archiveMessage(messageId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: 'archive' });
    this.cache.clear();
  }

  /**
   * Moves a message to the Junk folder.
   */
  async junkMessage(messageId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: 'junkemail' });
    this.cache.clear();
  }

  /**
   * Updates message properties (read status, flag, categories).
   */
  async updateMessage(messageId: string, updates: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}`)
      .patch(updates);
    this.cache.clear();
  }

  /**
   * Creates a new mail folder.
   */
  async createMailFolder(
    displayName: string,
    parentFolderId?: string
  ): Promise<MicrosoftGraph.MailFolder> {
    const client = await this.getClient();
    const url = parentFolderId != null
      ? `/me/mailFolders/${parentFolderId}/childFolders`
      : '/me/mailFolders';

    const result = await client
      .api(url)
      .post({ displayName }) as MicrosoftGraph.MailFolder;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes a mail folder.
   */
  async deleteMailFolder(folderId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/mailFolders/${folderId}`)
      .delete();
    this.cache.clear();
  }

  /**
   * Renames a mail folder.
   */
  async renameMailFolder(folderId: string, newName: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/mailFolders/${folderId}`)
      .patch({ displayName: newName });
    this.cache.clear();
  }

  /**
   * Moves a mail folder to a new parent.
   */
  async moveMailFolder(folderId: string, destinationParentId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/mailFolders/${folderId}/move`)
      .post({ destinationId: destinationParentId });
    this.cache.clear();
  }

  /**
   * Deletes all messages in a folder.
   */
  async emptyMailFolder(folderId: string): Promise<void> {
    const client = await this.getClient();
    // Get all messages in the folder
    let response = await client
      .api(`/me/mailFolders/${folderId}/messages`)
      .select('id')
      .top(100)
      .get() as PageCollection;

    // Delete each message
    for (const message of response.value as MicrosoftGraph.Message[]) {
      if (message.id != null) {
        await client
          .api(`/me/messages/${message.id}/move`)
          .post({ destinationId: 'deleteditems' });
      }
    }

    // Handle pagination
    while (response['@odata.nextLink'] != null) {
      response = await client.api(response['@odata.nextLink']).get() as PageCollection;
      for (const message of response.value as MicrosoftGraph.Message[]) {
        if (message.id != null) {
          await client
            .api(`/me/messages/${message.id}/move`)
            .post({ destinationId: 'deleteditems' });
        }
      }
    }

    this.cache.clear();
  }

  // ===========================================================================
  // Draft & Send Operations
  // ===========================================================================

  /**
   * Creates a new draft message.
   */
  async createDraft(message: {
    subject: string;
    body: MicrosoftGraph.ItemBody;
    toRecipients?: MicrosoftGraph.Recipient[];
    ccRecipients?: MicrosoftGraph.Recipient[];
    bccRecipients?: MicrosoftGraph.Recipient[];
    isDraft?: boolean;
  }): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api('/me/messages')
      .post(message) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  /**
   * Updates an existing draft message.
   */
  async updateDraft(messageId: string, updates: Record<string, unknown>): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/messages/${messageId}`)
      .patch(updates) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  /**
   * Sends an existing draft message.
   */
  async sendDraft(messageId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}/send`)
      .post(null);
    this.cache.clear();
  }

  /**
   * Sends a new email directly without creating a draft.
   */
  async sendMail(message: {
    subject: string;
    body: MicrosoftGraph.ItemBody;
    toRecipients: MicrosoftGraph.Recipient[];
    ccRecipients?: MicrosoftGraph.Recipient[];
    bccRecipients?: MicrosoftGraph.Recipient[];
  }): Promise<void> {
    const client = await this.getClient();
    await client
      .api('/me/sendMail')
      .post({ message });
    this.cache.clear();
  }

  /**
   * Replies to a message, or replies to all recipients.
   */
  async replyMessage(messageId: string, comment: string, replyAll: boolean): Promise<void> {
    const client = await this.getClient();
    const action = replyAll ? 'replyAll' : 'reply';
    await client
      .api(`/me/messages/${messageId}/${action}`)
      .post({ comment });
    this.cache.clear();
  }

  /**
   * Forwards a message to specified recipients.
   */
  async forwardMessage(
    messageId: string,
    toRecipients: MicrosoftGraph.Recipient[],
    comment?: string
  ): Promise<void> {
    const client = await this.getClient();
    const body: { toRecipients: MicrosoftGraph.Recipient[]; comment?: string } = { toRecipients };
    if (comment != null) {
      body.comment = comment;
    }
    await client
      .api(`/me/messages/${messageId}/forward`)
      .post(body);
    this.cache.clear();
  }

  /**
   * Creates a reply draft for a message.
   */
  async createReplyDraft(messageId: string, comment?: string, body?: { contentType: string; content: string }): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const postBody: Record<string, unknown> = {};
    if (comment != null) postBody.comment = comment;
    if (body != null) postBody.message = { body };
    const result = await client
      .api(`/me/messages/${messageId}/createReply`)
      .post(Object.keys(postBody).length > 0 ? postBody : null) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  /**
   * Creates a reply-all draft for a message.
   */
  async createReplyAllDraft(messageId: string, comment?: string, body?: { contentType: string; content: string }): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const postBody: Record<string, unknown> = {};
    if (comment != null) postBody.comment = comment;
    if (body != null) postBody.message = { body };
    const result = await client
      .api(`/me/messages/${messageId}/createReplyAll`)
      .post(Object.keys(postBody).length > 0 ? postBody : null) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  /**
   * Creates a forward draft for a message.
   */
  async createForwardDraft(messageId: string): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/messages/${messageId}/createForward`)
      .post(null) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  // ---------------------------------------------------------------------------
  // Calendar Scheduling
  // ---------------------------------------------------------------------------

  /**
   * Gets the free/busy schedule for one or more people.
   * POST /me/calendar/getSchedule
   */
  async getSchedule(params: {
    schedules: string[];
    startTime: { dateTime: string; timeZone: string };
    endTime: { dateTime: string; timeZone: string };
    availabilityViewInterval?: number;
  }): Promise<unknown[]> {
    const client = await this.getClient();
    const response = await client.api('/me/calendar/getSchedule').post(params) as { value: unknown[] };
    return response.value;
  }

  /**
   * Suggests meeting times for a set of attendees.
   * POST /me/findMeetingTimes
   */
  async findMeetingTimes(params: {
    attendees: Array<{ emailAddress: { address: string }; type: string }>;
    meetingDuration: string;
    timeConstraint?: {
      timeslots: Array<{
        start: { dateTime: string; timeZone: string };
        end: { dateTime: string; timeZone: string };
      }>;
    };
    maxCandidates?: number;
  }): Promise<unknown> {
    const client = await this.getClient();
    return (await client.api('/me/findMeetingTimes').post(params)) as unknown;
  }

  // ===========================================================================
  // Attachment Operations
  // ===========================================================================

  /**
   * Lists attachments on a message.
   */
  async listAttachments(messageId: string): Promise<MicrosoftGraph.Attachment[]> {
    const client = await this.getClient();

    const response = await client
      .api(`/me/messages/${messageId}/attachments`)
      .select('id,name,size,contentType,isInline')
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Attachment[];
  }

  /**
   * Gets a specific attachment with full content (including contentBytes).
   */
  async getAttachment(messageId: string, attachmentId: string): Promise<MicrosoftGraph.FileAttachment> {
    const client = await this.getClient();

    return await client
      .api(`/me/messages/${messageId}/attachments/${attachmentId}`)
      .get() as MicrosoftGraph.FileAttachment;
  }

  /**
   * Adds an inline base64 attachment to a message (<= 3MB).
   */
  async addAttachment(messageId: string, attachment: Record<string, unknown>): Promise<MicrosoftGraph.Attachment> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/messages/${messageId}/attachments`)
      .post(attachment) as MicrosoftGraph.Attachment;
    this.cache.clear();
    return result;
  }

  /**
   * Creates an upload session for large file attachments (> 3MB).
   */
  async createUploadSession(messageId: string, body: Record<string, unknown>): Promise<{ uploadUrl: string }> {
    const client = await this.getClient();
    return await client
      .api(`/me/messages/${messageId}/attachments/createUploadSession`)
      .post(body) as { uploadUrl: string };
  }

  // ===========================================================================
  // Tasks (Microsoft To Do) - continued
  // ===========================================================================

  /**
   * Searches tasks by title.
   */
  async searchTasks(query: string, limit: number = 50): Promise<Array<MicrosoftGraph.TodoTask & { taskListId: string }>> {
    const allTasks = await this.listAllTasks(1000, 0, true);

    const queryLower = query.toLowerCase();
    const matched = allTasks.filter(
      (task) => task.title?.toLowerCase().includes(queryLower) ?? false
    );

    return matched.slice(0, limit);
  }

  // ===========================================================================
  // Task Write Operations
  // ===========================================================================

  /**
   * Creates a new task in a task list.
   */
  async createTask(
    taskListId: string,
    task: Record<string, unknown>
  ): Promise<MicrosoftGraph.TodoTask> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/todo/lists/${taskListId}/tasks`)
      .post(task) as MicrosoftGraph.TodoTask;
    this.cache.clear();
    return result;
  }

  /**
   * Updates an existing task.
   */
  async updateTask(
    taskListId: string,
    taskId: string,
    updates: Record<string, unknown>
  ): Promise<MicrosoftGraph.TodoTask> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/todo/lists/${taskListId}/tasks/${taskId}`)
      .patch(updates) as MicrosoftGraph.TodoTask;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes a task.
   */
  async deleteTask(taskListId: string, taskId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/todo/lists/${taskListId}/tasks/${taskId}`)
      .delete();
    this.cache.clear();
  }

  /**
   * Creates a new task list.
   */
  async createTaskList(displayName: string): Promise<MicrosoftGraph.TodoTaskList> {
    const client = await this.getClient();
    const result = await client
      .api('/me/todo/lists')
      .post({ displayName }) as MicrosoftGraph.TodoTaskList;
    this.cache.clear();
    return result;
  }

  /**
   * Updates a task list (e.g. rename).
   */
  async updateTaskList(listId: string, updates: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/todo/lists/${listId}`).patch(updates);
    this.cache.clear();
  }

  /**
   * Deletes a task list.
   */
  async deleteTaskList(listId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/todo/lists/${listId}`).delete();
    this.cache.clear();
  }

  // ===========================================================================
  // Master Categories
  // ===========================================================================

  /**
   * Lists all master categories.
   */
  async listMasterCategories(): Promise<MicrosoftGraph.OutlookCategory[]> {
    const client = await this.getClient();
    const response = await client.api('/me/outlook/masterCategories').get() as PageCollection;
    return response.value as MicrosoftGraph.OutlookCategory[];
  }

  /**
   * Creates a new master category.
   */
  async createMasterCategory(displayName: string, color: string): Promise<MicrosoftGraph.OutlookCategory> {
    const client = await this.getClient();
    const result = await client.api('/me/outlook/masterCategories').post({ displayName, color }) as MicrosoftGraph.OutlookCategory;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes a master category.
   */
  async deleteMasterCategory(categoryId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/outlook/masterCategories/${categoryId}`).delete();
    this.cache.clear();
  }

  // ===========================================================================
  // Focused Inbox Overrides
  // ===========================================================================

  /**
   * Lists all focused inbox overrides.
   */
  async listFocusedOverrides(): Promise<MicrosoftGraph.InferenceClassificationOverride[]> {
    const client = await this.getClient();
    const response = await client.api('/me/inferenceClassification/overrides').get() as PageCollection;
    return response.value as MicrosoftGraph.InferenceClassificationOverride[];
  }

  /**
   * Creates a focused inbox override.
   */
  async createFocusedOverride(senderAddress: string, classifyAs: 'focused' | 'other'): Promise<MicrosoftGraph.InferenceClassificationOverride> {
    const client = await this.getClient();
    const result = await client.api('/me/inferenceClassification/overrides').post({
      classifyAs,
      senderEmailAddress: { address: senderAddress },
    }) as MicrosoftGraph.InferenceClassificationOverride;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes a focused inbox override.
   */
  async deleteFocusedOverride(overrideId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/inferenceClassification/overrides/${overrideId}`).delete();
    this.cache.clear();
  }

  // ===========================================================================
  // Message Headers & MIME
  // ===========================================================================

  /**
   * Gets internet message headers for a message.
   */
  async getMessageHeaders(messageId: string): Promise<Array<{ name: string; value: string }>> {
    const client = await this.getClient();
    const message = await client
      .api(`/me/messages/${messageId}`)
      .select('internetMessageHeaders')
      .get() as MicrosoftGraph.Message;
    return (message.internetMessageHeaders ?? []) as Array<{ name: string; value: string }>;
  }

  /**
   * Gets the MIME content of a message.
   */
  async getMessageMime(messageId: string): Promise<string> {
    const client = await this.getClient();
    return await client.api(`/me/messages/${messageId}/$value`).get() as string;
  }

  // ===========================================================================
  // Calendar Groups
  // ===========================================================================

  /**
   * Lists all calendar groups.
   */
  async listCalendarGroups(): Promise<MicrosoftGraph.CalendarGroup[]> {
    const client = await this.getClient();
    const response = await client.api('/me/calendarGroups').get() as PageCollection;
    return response.value as MicrosoftGraph.CalendarGroup[];
  }

  /**
   * Creates a new calendar group.
   */
  async createCalendarGroup(name: string): Promise<MicrosoftGraph.CalendarGroup> {
    const client = await this.getClient();
    const result = await client.api('/me/calendarGroups').post({ name }) as MicrosoftGraph.CalendarGroup;
    this.cache.clear();
    return result;
  }

  // ===========================================================================
  // Calendar Permissions
  // ===========================================================================

  /**
   * Lists all permissions for a calendar.
   */
  async listCalendarPermissions(calendarId: string): Promise<MicrosoftGraph.CalendarPermission[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/calendars/${calendarId}/calendarPermissions`).get() as PageCollection;
    return response.value as MicrosoftGraph.CalendarPermission[];
  }

  /**
   * Creates a calendar permission (shares a calendar).
   */
  async createCalendarPermission(calendarId: string, permission: Record<string, unknown>): Promise<MicrosoftGraph.CalendarPermission> {
    const client = await this.getClient();
    const result = await client.api(`/me/calendars/${calendarId}/calendarPermissions`).post(permission) as MicrosoftGraph.CalendarPermission;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes a calendar permission.
   */
  async deleteCalendarPermission(calendarId: string, permissionId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/calendars/${calendarId}/calendarPermissions/${permissionId}`).delete();
    this.cache.clear();
  }

  // ===========================================================================
  // Room Lists & Rooms
  // ===========================================================================

  /**
   * GET /me/findRoomLists
   */
  async listRoomLists(): Promise<MicrosoftGraph.EmailAddress[]> {
    const client = await this.getClient();
    const response = await client.api('/me/findRoomLists').get() as { value: MicrosoftGraph.EmailAddress[] };
    return response.value;
  }

  /**
   * GET /me/findRooms or /me/findRooms(RoomList='...')
   */
  async listRooms(roomListEmail?: string): Promise<MicrosoftGraph.EmailAddress[]> {
    const client = await this.getClient();
    const endpoint = roomListEmail != null
      ? `/me/findRooms(RoomList='${roomListEmail}')`
      : '/me/findRooms';
    const response = await client.api(endpoint).get() as { value: MicrosoftGraph.EmailAddress[] };
    return response.value;
  }

  // ===========================================================================
  // Mail Tips
  // ===========================================================================

  /**
   * Gets mail tips for the specified email addresses.
   */
  async getMailTips(emailAddresses: string[]): Promise<Record<string, unknown>[]> {
    const client = await this.getClient();
    const response = await client.api('/me/getMailTips').post({
      emailAddresses,
      mailTipsOptions: 'automaticReplies,mailboxFullStatus,maxMessageSize,deliveryRestriction,externalMemberCount',
    }) as { value: Record<string, unknown>[] };
    return response.value;
  }

  // ===========================================================================
  // Teams
  // ===========================================================================

  /**
   * Lists all teams the current user has joined.
   */
  async listJoinedTeams(): Promise<MicrosoftGraph.Team[]> {
    const client = await this.getClient();
    const response = await client.api('/me/joinedTeams').get() as PageCollection;
    return response.value as MicrosoftGraph.Team[];
  }

  /**
   * Lists all channels in a team.
   */
  async listChannels(teamId: string): Promise<MicrosoftGraph.Channel[]> {
    const client = await this.getClient();
    const response = await client.api(`/teams/${teamId}/channels`).get() as PageCollection;
    return response.value as MicrosoftGraph.Channel[];
  }

  /**
   * Gets a specific channel.
   */
  async getChannel(teamId: string, channelId: string): Promise<MicrosoftGraph.Channel> {
    const client = await this.getClient();
    return await client.api(`/teams/${teamId}/channels/${channelId}`).get() as MicrosoftGraph.Channel;
  }

  /**
   * Creates a new channel in a team.
   */
  async createChannel(teamId: string, displayName: string, description?: string): Promise<MicrosoftGraph.Channel> {
    const client = await this.getClient();
    const body: Record<string, unknown> = { displayName };
    if (description != null) body['description'] = description;
    return await client.api(`/teams/${teamId}/channels`).post(body) as MicrosoftGraph.Channel;
  }

  /**
   * Updates a channel's properties.
   */
  async updateChannel(teamId: string, channelId: string, updates: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client.api(`/teams/${teamId}/channels/${channelId}`).patch(updates);
  }

  /**
   * Deletes a channel.
   */
  async deleteChannel(teamId: string, channelId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/teams/${teamId}/channels/${channelId}`).delete();
  }

  /**
   * Lists members of a team.
   */
  async listTeamMembers(teamId: string): Promise<MicrosoftGraph.ConversationMember[]> {
    const client = await this.getClient();
    const response = await client.api(`/teams/${teamId}/members`).get() as PageCollection;
    return response.value as MicrosoftGraph.ConversationMember[];
  }

  // ===========================================================================
  // Channel Messages
  // ===========================================================================

  /**
   * Lists recent messages in a channel.
   */
  async listChannelMessages(teamId: string, channelId: string, top: number = 25): Promise<MicrosoftGraph.ChatMessage[]> {
    const client = await this.getClient();
    const response = await client.api(`/teams/${teamId}/channels/${channelId}/messages`).top(top).get() as PageCollection;
    return response.value as MicrosoftGraph.ChatMessage[];
  }

  /**
   * Gets a specific channel message.
   */
  async getChannelMessage(teamId: string, channelId: string, messageId: string): Promise<MicrosoftGraph.ChatMessage> {
    const client = await this.getClient();
    return await client.api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}`).get() as MicrosoftGraph.ChatMessage;
  }

  /**
   * Lists replies to a channel message.
   */
  async listChannelMessageReplies(teamId: string, channelId: string, messageId: string): Promise<MicrosoftGraph.ChatMessage[]> {
    const client = await this.getClient();
    const response = await client.api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`).get() as PageCollection;
    return response.value as MicrosoftGraph.ChatMessage[];
  }

  /**
   * Sends a new message to a channel.
   */
  async sendChannelMessage(teamId: string, channelId: string, body: string, contentType: string = 'html'): Promise<MicrosoftGraph.ChatMessage> {
    const client = await this.getClient();
    return await client.api(`/teams/${teamId}/channels/${channelId}/messages`).post({
      body: { contentType, content: body },
    }) as MicrosoftGraph.ChatMessage;
  }

  /**
   * Replies to a channel message.
   */
  async replyToChannelMessage(teamId: string, channelId: string, messageId: string, body: string, contentType: string = 'html'): Promise<MicrosoftGraph.ChatMessage> {
    const client = await this.getClient();
    return await client.api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`).post({
      body: { contentType, content: body },
    }) as MicrosoftGraph.ChatMessage;
  }

  // ===========================================================================
  // Chats
  // ===========================================================================

  /**
   * Lists the signed-in user's chats.
   *
   * @param top - Page size (default 25)
   * @param options.expandMembers - Also `$expand=members` (displayName + email inline)
   * @param options.chatType - Optional `$filter` on chatType
   * @param options.pageAll - Follow `@odata.nextLink` (capped) for full enumeration
   */
  async listChats(
    top: number = 25,
    options: {
      expandMembers?: boolean;
      chatType?: 'oneOnOne' | 'group';
      pageAll?: boolean;
    } = {},
  ): Promise<MicrosoftGraph.Chat[]> {
    const client = await this.getClient();
    const expand = options.expandMembers === true
      ? 'lastMessagePreview,members'
      : 'lastMessagePreview';

    let request = client.api('/me/chats')
      .top(top)
      .orderby('lastMessagePreview/createdDateTime desc')
      .expand(expand);
    if (options.chatType != null) {
      request = request.filter(`chatType eq '${options.chatType}'`);
    }

    let response = await request.get() as PageCollection;
    const result = [...(response.value as MicrosoftGraph.Chat[])];

    if (options.pageAll === true) {
      let pages = 1;
      while (response['@odata.nextLink'] != null && pages < 20) {
        response = await client.api(response['@odata.nextLink']).get() as PageCollection;
        result.push(...(response.value as MicrosoftGraph.Chat[]));
        pages += 1;
      }
    }

    return result;
  }

  /**
   * Creates a chat, or for oneOnOne returns the existing chat if one already
   * exists between the same two members (Graph get-or-create semantics).
   *
   * Every participant — including the signed-in user — must be listed. This
   * method always adds `/me` and the given member identifiers (email, UPN, or id).
   */
  async createChat(
    chatType: 'oneOnOne' | 'group',
    memberIdentifiers: string[],
    topic?: string,
  ): Promise<MicrosoftGraph.Chat> {
    const client = await this.getClient();
    const me = await client.api('/me').select('id').get() as { id: string };
    const seen = new Set<string>([me.id.toLowerCase()]);
    const identifiers = [me.id];
    for (const raw of memberIdentifiers) {
      const key = raw.trim().toLowerCase();
      if (key === '' || seen.has(key)) continue;
      seen.add(key);
      identifiers.push(raw.trim());
    }

    const members = identifiers.map((identifier) => {
      // OData string literals escape ' as ''; guest UPNs often contain #.
      const escaped = identifier.replace(/'/g, "''").replace(/#/g, '%23');
      return {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${escaped}')`,
      };
    });

    const body: Record<string, unknown> = { chatType, members };
    if (topic != null && chatType === 'group') {
      body['topic'] = topic;
    }

    return await client.api('/chats').post(body) as MicrosoftGraph.Chat;
  }

  async getChat(chatId: string): Promise<MicrosoftGraph.Chat> {
    const client = await this.getClient();
    return await client.api(`/me/chats/${chatId}`).get() as MicrosoftGraph.Chat;
  }

  async listChatMessages(chatId: string, top: number = 25): Promise<MicrosoftGraph.ChatMessage[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/chats/${chatId}/messages`).top(top).get() as PageCollection;
    return response.value as MicrosoftGraph.ChatMessage[];
  }

  async sendChatMessage(chatId: string, body: string, contentType: string = 'html'): Promise<MicrosoftGraph.ChatMessage> {
    const client = await this.getClient();
    return await client.api(`/me/chats/${chatId}/messages`).post({
      body: { contentType, content: body },
    }) as MicrosoftGraph.ChatMessage;
  }

  async getChatMessage(chatId: string, messageId: string): Promise<MicrosoftGraph.ChatMessage> {
    const client = await this.getClient();
    return await client.api(`/me/chats/${chatId}/messages/${messageId}`).get() as MicrosoftGraph.ChatMessage;
  }

  // Channel message reactions
  async setChannelMessageReaction(teamId: string, channelId: string, messageId: string, reactionType: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/setReaction`)
      .post({ reactionType });
  }

  async unsetChannelMessageReaction(teamId: string, channelId: string, messageId: string, reactionType: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/unsetReaction`)
      .post({ reactionType });
  }

  // Chat message reactions
  async setChatMessageReaction(chatId: string, messageId: string, reactionType: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/chats/${chatId}/messages/${messageId}/setReaction`)
      .post({ reactionType });
  }

  async unsetChatMessageReaction(chatId: string, messageId: string, reactionType: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/chats/${chatId}/messages/${messageId}/unsetReaction`)
      .post({ reactionType });
  }

  async listChatMembers(chatId: string): Promise<MicrosoftGraph.ConversationMember[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/chats/${chatId}/members`).get() as PageCollection;
    return response.value as MicrosoftGraph.ConversationMember[];
  }

  // ===========================================================================
  // Checklist Items
  // ===========================================================================

  /**
   * Lists checklist items on a task.
   */
  async listChecklistItems(taskListId: string, taskId: string): Promise<MicrosoftGraph.ChecklistItem[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/checklistItems`).get() as PageCollection;
    return response.value as MicrosoftGraph.ChecklistItem[];
  }

  /**
   * Creates a checklist item on a task.
   */
  async createChecklistItem(taskListId: string, taskId: string, displayName: string, isChecked: boolean = false): Promise<MicrosoftGraph.ChecklistItem> {
    const client = await this.getClient();
    return await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/checklistItems`).post({
      displayName,
      isChecked,
    }) as MicrosoftGraph.ChecklistItem;
  }

  /**
   * Updates a checklist item.
   */
  async updateChecklistItem(taskListId: string, taskId: string, checklistItemId: string, updates: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/checklistItems/${checklistItemId}`).patch(updates);
  }

  /**
   * Deletes a checklist item.
   */
  async deleteChecklistItem(taskListId: string, taskId: string, checklistItemId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/checklistItems/${checklistItemId}`).delete();
  }

  // ===========================================================================
  // Linked Resources
  // ===========================================================================

  async listLinkedResources(taskListId: string, taskId: string): Promise<MicrosoftGraph.LinkedResource[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/linkedResources`).get() as PageCollection;
    return response.value as MicrosoftGraph.LinkedResource[];
  }

  async createLinkedResource(taskListId: string, taskId: string, webUrl: string, applicationName: string, displayName?: string): Promise<MicrosoftGraph.LinkedResource> {
    const client = await this.getClient();
    const body: Record<string, unknown> = { webUrl, applicationName };
    if (displayName != null) body['displayName'] = displayName;
    return await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/linkedResources`).post(body) as MicrosoftGraph.LinkedResource;
  }

  async deleteLinkedResource(taskListId: string, taskId: string, linkedResourceId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/linkedResources/${linkedResourceId}`).delete();
  }

  // ===========================================================================
  // Task Attachments
  // ===========================================================================

  async listTaskAttachments(taskListId: string, taskId: string): Promise<MicrosoftGraph.AttachmentBase[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/attachments`).get() as PageCollection;
    return response.value as MicrosoftGraph.AttachmentBase[];
  }

  async createTaskAttachment(taskListId: string, taskId: string, name: string, contentBytes: string, contentType: string = 'application/octet-stream'): Promise<MicrosoftGraph.AttachmentBase> {
    const client = await this.getClient();
    return await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/attachments`).post({
      '@odata.type': '#microsoft.graph.taskFileAttachment',
      name,
      contentBytes,
      contentType,
    }) as MicrosoftGraph.AttachmentBase;
  }

  async deleteTaskAttachment(taskListId: string, taskId: string, attachmentId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/attachments/${attachmentId}`).delete();
  }

  // ===========================================================================
  // Planner
  // ===========================================================================

  async listPlans(): Promise<MicrosoftGraph.PlannerPlan[]> {
    const client = await this.getClient();
    const response = await client.api('/me/planner/plans').get() as PageCollection;
    return response.value as MicrosoftGraph.PlannerPlan[];
  }

  async getPlan(planId: string): Promise<MicrosoftGraph.PlannerPlan> {
    const client = await this.getClient();
    return await client.api(`/planner/plans/${planId}`).get() as MicrosoftGraph.PlannerPlan;
  }

  async createPlan(title: string, groupId: string): Promise<MicrosoftGraph.PlannerPlan> {
    const client = await this.getClient();
    return await client.api('/planner/plans').post({
      title,
      owner: groupId,
      container: { url: `https://graph.microsoft.com/v1.0/groups/${groupId}`, type: 'group' },
    }) as MicrosoftGraph.PlannerPlan;
  }

  async updatePlan(planId: string, updates: Record<string, unknown>, etag: string): Promise<MicrosoftGraph.PlannerPlan> {
    const client = await this.getClient();
    return await client.api(`/planner/plans/${planId}`).header('If-Match', etag).patch(updates) as MicrosoftGraph.PlannerPlan;
  }

  async listBuckets(planId: string): Promise<MicrosoftGraph.PlannerBucket[]> {
    const client = await this.getClient();
    const response = await client.api(`/planner/plans/${planId}/buckets`).get() as PageCollection;
    return response.value as MicrosoftGraph.PlannerBucket[];
  }

  async createBucket(planId: string, name: string): Promise<MicrosoftGraph.PlannerBucket> {
    const client = await this.getClient();
    return await client.api('/planner/buckets').post({ planId, name }) as MicrosoftGraph.PlannerBucket;
  }

  async getBucket(bucketId: string): Promise<MicrosoftGraph.PlannerBucket> {
    const client = await this.getClient();
    return await client.api(`/planner/buckets/${bucketId}`).get() as MicrosoftGraph.PlannerBucket;
  }

  async updateBucket(bucketId: string, updates: Record<string, unknown>, etag: string): Promise<MicrosoftGraph.PlannerBucket> {
    const client = await this.getClient();
    return await client.api(`/planner/buckets/${bucketId}`).header('If-Match', etag).patch(updates) as MicrosoftGraph.PlannerBucket;
  }

  async deleteBucket(bucketId: string, etag: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/planner/buckets/${bucketId}`).header('If-Match', etag).delete();
  }

  // ===========================================================================
  // Planner Tasks
  // ===========================================================================

  async listPlannerTasks(planId: string): Promise<MicrosoftGraph.PlannerTask[]> {
    const client = await this.getClient();
    const response = await client.api(`/planner/plans/${planId}/tasks`).get() as PageCollection;
    return response.value as MicrosoftGraph.PlannerTask[];
  }

  /** All Planner tasks assigned to the signed-in user, across every plan. */
  async listMyPlannerTasks(): Promise<MicrosoftGraph.PlannerTask[]> {
    const client = await this.getClient();
    const response = await client.api('/me/planner/tasks').get() as PageCollection;
    return response.value as MicrosoftGraph.PlannerTask[];
  }

  async getPlannerTask(taskId: string): Promise<MicrosoftGraph.PlannerTask> {
    const client = await this.getClient();
    return await client.api(`/planner/tasks/${taskId}`).get() as MicrosoftGraph.PlannerTask;
  }

  async createPlannerTask(task: Record<string, unknown>): Promise<MicrosoftGraph.PlannerTask> {
    const client = await this.getClient();
    return await client.api('/planner/tasks').post(task) as MicrosoftGraph.PlannerTask;
  }

  async updatePlannerTask(taskId: string, updates: Record<string, unknown>, etag: string): Promise<MicrosoftGraph.PlannerTask> {
    const client = await this.getClient();
    return await client.api(`/planner/tasks/${taskId}`).header('If-Match', etag).patch(updates) as MicrosoftGraph.PlannerTask;
  }

  async deletePlannerTask(taskId: string, etag: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/planner/tasks/${taskId}`).header('If-Match', etag).delete();
  }

  async getPlannerTaskDetails(taskId: string): Promise<MicrosoftGraph.PlannerTaskDetails> {
    const client = await this.getClient();
    return await client.api(`/planner/tasks/${taskId}/details`).get() as MicrosoftGraph.PlannerTaskDetails;
  }

  async updatePlannerTaskDetails(taskId: string, updates: Record<string, unknown>, etag: string): Promise<MicrosoftGraph.PlannerTaskDetails> {
    const client = await this.getClient();
    return await client.api(`/planner/tasks/${taskId}/details`).header('If-Match', etag).patch(updates) as MicrosoftGraph.PlannerTaskDetails;
  }

  // Planner Task Chat Messages (beta — delegated only)
  // https://learn.microsoft.com/en-us/graph/api/resources/plannertaskchatmessage

  async listPlannerTaskMessages(
    taskId: string,
    skipToken?: string,
  ): Promise<{ messages: GraphEntity[]; nextSkipToken?: string }> {
    const client = await this.getClient();
    let request = client.api(`/planner/tasks/${taskId}/messages`).version('beta');
    if (skipToken != null && skipToken.length > 0) {
      request = request.query({ $skipToken: skipToken });
    }
    const response = await request.get() as PageCollection & { '@odata.nextLink'?: string };
    const messages = (response.value ?? []) as GraphEntity[];
    let nextSkipToken: string | undefined;
    const nextLink = response['@odata.nextLink'];
    if (nextLink != null) {
      try {
        nextSkipToken = new URL(nextLink).searchParams.get('$skipToken') ?? undefined;
      } catch {
        nextSkipToken = undefined;
      }
    }
    return nextSkipToken != null
      ? { messages, nextSkipToken }
      : { messages };
  }

  async createPlannerTaskMessage(
    taskId: string,
    body: Record<string, unknown>,
  ): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client.api(`/planner/tasks/${taskId}/messages`).version('beta').post(body) as GraphEntity;
  }

  async updatePlannerTaskMessage(
    taskId: string,
    messageId: string,
    body: Record<string, unknown>,
  ): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client.api(`/planner/tasks/${taskId}/messages/${messageId}`).version('beta').patch(body) as GraphEntity;
  }

  async deletePlannerTaskMessage(taskId: string, messageId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/planner/tasks/${taskId}/messages/${messageId}`).version('beta').delete();
  }

  // ===========================================================================
  // People & Presence
  // ===========================================================================

  async listRelevantPeople(top: number = 25): Promise<MicrosoftGraph.Person[]> {
    const client = await this.getClient();
    const response = await client.api('/me/people').top(top).get() as PageCollection;
    return response.value as MicrosoftGraph.Person[];
  }

  async searchPeople(query: string, top: number = 25): Promise<MicrosoftGraph.Person[]> {
    const client = await this.getClient();
    const response = await client.api('/me/people').search('"' + query + '"').top(top).get() as PageCollection;
    return response.value as MicrosoftGraph.Person[];
  }

  async getManager(): Promise<MicrosoftGraph.DirectoryObject> {
    const client = await this.getClient();
    return await client.api('/me/manager').get() as MicrosoftGraph.DirectoryObject;
  }

  async getDirectReports(): Promise<MicrosoftGraph.DirectoryObject[]> {
    const client = await this.getClient();
    const response = await client.api('/me/directReports').get() as PageCollection;
    return response.value as MicrosoftGraph.DirectoryObject[];
  }

  async getUserProfile(identifier: string): Promise<MicrosoftGraph.User> {
    const client = await this.getClient();
    return await client.api(`/users/${identifier}`).get() as MicrosoftGraph.User;
  }

  async getUserPhoto(identifier: string): Promise<ArrayBuffer> {
    const client = await this.getClient();
    return await client.api(`/users/${identifier}/photo/$value`).responseType(ResponseType.ARRAYBUFFER).get() as ArrayBuffer;
  }

  async getUserPresence(identifier: string): Promise<MicrosoftGraph.Presence> {
    const client = await this.getClient();
    return await client.api(`/users/${identifier}/presence`).get() as MicrosoftGraph.Presence;
  }

  async getUsersPresence(userIds: string[]): Promise<MicrosoftGraph.Presence[]> {
    const client = await this.getClient();
    const response = await client.api('/communications/getPresencesByUserId').post({ ids: userIds }) as { value: MicrosoftGraph.Presence[] };
    return response.value;
  }

  // ===========================================================================
  // Shared Mailbox / Delegate Access (/users/{upn}/...) — read-only (#40)
  // ===========================================================================
  //
  // These mirror the `/me/...` read paths but target another user's mailbox,
  // calendar, or drive via delegate/shared access. Unlike the `/me` variants,
  // they deliberately DO NOT swallow errors: a 403 must surface (mapped to
  // GRAPH_PERMISSION_DENIED at the dispatch chokepoint) so the caller learns
  // they lack shared access rather than getting a silent empty/`null` result.
  // Results are returned with raw Graph ids (durable tokens are `/me`- and
  // account-scoped, so they cannot address another mailbox's items).

  async listSharedMailFolders(mailbox: string): Promise<MicrosoftGraph.MailFolder[]> {
    const client = await this.getClient();
    const response = await client
      .api(`/users/${encodeURIComponent(mailbox)}/mailFolders`)
      .select('id,displayName,parentFolderId,totalItemCount,unreadItemCount')
      .top(100)
      .get() as PageCollection;
    return response.value as MicrosoftGraph.MailFolder[];
  }

  async listSharedMessages(
    mailbox: string,
    folderId?: string,
    limit: number = 25,
    unreadOnly: boolean = false,
  ): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();
    const apiPath = folderId != null
      ? `/users/${encodeURIComponent(mailbox)}/mailFolders/${encodeURIComponent(folderId)}/messages`
      : `/users/${encodeURIComponent(mailbox)}/messages`;
    let request = client
      .api(apiPath)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId')
      .top(limit);
    if (unreadOnly) {
      // Graph rejects `$orderby` on a property not led by `$filter` (isRead),
      // so drop the sort when filtering unread and rely on the default order.
      request = request.filter('isRead eq false');
    } else {
      request = request.orderby('receivedDateTime desc');
    }
    const response = await request.get() as PageCollection;
    return response.value as MicrosoftGraph.Message[];
  }

  async getSharedMessage(
    mailbox: string,
    messageId: string,
    includeBody: boolean = false,
  ): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const select = includeBody
      ? 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,body,bodyPreview,conversationId,parentFolderId'
      : 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,parentFolderId';
    return await client
      .api(`/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(messageId)}`)
      .select(select)
      .get() as MicrosoftGraph.Message;
  }

  async searchSharedMessages(
    mailbox: string,
    query: string,
    limit: number = 25,
  ): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();
    // Mail $search cannot be combined with $orderby, so ordering is omitted.
    // Strip embedded double quotes so a stray `"` can't break out of the
    // quoted KQL `$search="..."` term.
    const safeQuery = query.replace(/"/g, ' ');
    const response = await client
      .api(`/users/${encodeURIComponent(mailbox)}/messages`)
      .query({ $search: `"${safeQuery}"` })
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId')
      .top(limit)
      .get() as PageCollection;
    return response.value as MicrosoftGraph.Message[];
  }

  async listSharedEvents(
    mailbox: string,
    limit: number = 25,
    startDate?: Date,
    endDate?: Date,
  ): Promise<MicrosoftGraph.Event[]> {
    const client = await this.getClient();
    if (startDate != null && endDate != null) {
      const response = await client
        .api(`/users/${encodeURIComponent(mailbox)}/calendarView`)
        .query({ startDateTime: startDate.toISOString(), endDateTime: endDate.toISOString() })
        .select('id,subject,start,end,location,isAllDay,organizer,attendees,bodyPreview')
        .orderby('start/dateTime')
        .top(limit)
        .get() as PageCollection;
      return response.value as MicrosoftGraph.Event[];
    }
    const response = await client
      .api(`/users/${encodeURIComponent(mailbox)}/events`)
      .select('id,subject,start,end,location,isAllDay,organizer,attendees,bodyPreview')
      .orderby('start/dateTime')
      .top(limit)
      .get() as PageCollection;
    return response.value as MicrosoftGraph.Event[];
  }

  async getSharedEvent(mailbox: string, eventId: string): Promise<MicrosoftGraph.Event> {
    const client = await this.getClient();
    return await client
      .api(`/users/${encodeURIComponent(mailbox)}/events/${encodeURIComponent(eventId)}`)
      .select('id,subject,start,end,location,isAllDay,organizer,attendees,body,recurrence,bodyPreview')
      .get() as MicrosoftGraph.Event;
  }

  async listSharedDriveItems(mailbox: string, itemId?: string): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const apiPath = itemId != null
      ? `/users/${encodeURIComponent(mailbox)}/drive/items/${encodeURIComponent(itemId)}/children`
      : `/users/${encodeURIComponent(mailbox)}/drive/root/children`;
    const response = await client.api(apiPath).get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  async searchSharedDriveItems(mailbox: string, query: string, limit: number = 25): Promise<GraphEntity[]> {
    const client = await this.getClient();
    // Escape for the OData string literal first (a lone `'` closes the literal;
    // OData escapes it by doubling), then URL-encode the whole term.
    const safeQuery = encodeURIComponent(query.replace(/'/g, "''"));
    const response = await client
      .api(`/users/${encodeURIComponent(mailbox)}/drive/root/search(q='${safeQuery}')`)
      .top(limit)
      .get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }


  /**
   * Sends multiple requests in a single $batch call to the Graph API.
   * Automatically splits into multiple batches if there are more than 20 requests.
   */
  async batchRequests(requests: BatchRequest[]): Promise<Map<string, BatchResponseItem>> {
    const client = await this.getClient();
    const batches = splitIntoBatches(requests);
    const allResults = new Map<string, BatchResponseItem>();

    for (const batch of batches) {
      const payload = buildBatchPayload(batch);
      const response = await client.api('/$batch').post(payload) as { responses: BatchResponseItem[] };
      const results = parseBatchResponse(response);
      for (const [id, result] of results) {
        allResults.set(id, result);
      }
    }

    return allResults;
  }

  // ===========================================================================
  // Online Meetings
  // ===========================================================================

  async listOnlineMeetings(limit: number = 20): Promise<GraphEntity[]> {
    const client = await this.getClient();
    // Graph rejects $top on /me/onlineMeetings, so limit client-side. Guard a
    // negative limit (slice(0, -n) would count from the end).
    const cap = Math.max(0, limit);
    const response = await client.api('/me/onlineMeetings')
      .orderby('startDateTime desc')
      .get() as GraphCollectionResponse<GraphEntity>;
    return response.value.slice(0, cap);
  }

  async getOnlineMeeting(meetingId: string): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client.api(`/me/onlineMeetings/${meetingId}`).get() as GraphEntity;
  }

  async listMeetingRecordings(meetingId: string): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/onlineMeetings/${meetingId}/recordings`).get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  async getMeetingRecordingContent(meetingId: string, recordingId: string): Promise<ArrayBuffer> {
    const client = await this.getClient();
    return await client.api(`/me/onlineMeetings/${meetingId}/recordings/${recordingId}/content`).responseType(ResponseType.ARRAYBUFFER).get() as ArrayBuffer;
  }

  async listMeetingTranscripts(meetingId: string): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/onlineMeetings/${meetingId}/transcripts`).get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  async getMeetingTranscriptContent(meetingId: string, transcriptId: string, format: string = 'text/vtt'): Promise<string> {
    const client = await this.getClient();
    return await client.api(`/me/onlineMeetings/${meetingId}/transcripts/${transcriptId}/content`)
      .header('Accept', format)
      .get() as string;
  }


  // ===========================================================================
  // Excel Online (Workbook)
  // ===========================================================================

  /**
   * Lists worksheets in an Excel workbook.
   */
  async listWorksheets(driveItemId: string): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/drive/items/${driveItemId}/workbook/worksheets`).get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  /**
   * Gets cell values for a specific range in a worksheet.
   */
  async getWorksheetRange(driveItemId: string, worksheetName: string, range: string): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client.api(`/me/drive/items/${driveItemId}/workbook/worksheets/${encodeURIComponent(worksheetName)}/range(address='${encodeURIComponent(range)}')`).get() as GraphEntity;
  }

  /**
   * Gets the used range (all data) for a worksheet.
   */
  async getUsedRange(driveItemId: string, worksheetName: string): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client.api(`/me/drive/items/${driveItemId}/workbook/worksheets/${encodeURIComponent(worksheetName)}/usedRange`).get() as GraphEntity;
  }

  /**
   * Updates cell values for a specific range in a worksheet.
   */
  async updateWorksheetRange(driveItemId: string, worksheetName: string, range: string, values: unknown[][]): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client.api(`/me/drive/items/${driveItemId}/workbook/worksheets/${encodeURIComponent(worksheetName)}/range(address='${encodeURIComponent(range)}')`).patch({ values }) as GraphEntity;
  }

  /**
   * Gets rows from a named table in an Excel workbook.
   */
  async getTableData(driveItemId: string, tableName: string): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/drive/items/${driveItemId}/workbook/tables/${encodeURIComponent(tableName)}/rows`).get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }
  // ===========================================================================
  // OneDrive
  // ===========================================================================

  async listDriveItems(itemId?: string): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const apiPath = itemId != null ? `/me/drive/items/${itemId}/children` : '/me/drive/root/children';
    const response = await client.api(apiPath).get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  async searchDriveItems(query: string, limit: number = 25): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/drive/root/search(q='${encodeURIComponent(query)}')`).top(limit).get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  async getDriveItem(itemId: string): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client.api(`/me/drive/items/${itemId}`).get() as GraphEntity;
  }

  async downloadDriveItem(itemId: string): Promise<ArrayBuffer> {
    const client = await this.getClient();
    return await client.api(`/me/drive/items/${itemId}/content`).responseType(ResponseType.ARRAYBUFFER).get() as ArrayBuffer;
  }

  async uploadDriveItem(parentPath: string, fileName: string, content: Buffer): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client.api(`/me/drive/root:/${parentPath}/${fileName}:/content`)
      .header('Content-Type', 'application/octet-stream')
      .put(content) as GraphEntity;
  }

  async listRecentDriveItems(): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const response = await client.api('/me/drive/recent').get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  async listSharedWithMe(): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const response = await client.api('/me/drive/sharedWithMe').get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  async createSharingLink(itemId: string, type: string, scope: string): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client.api(`/me/drive/items/${itemId}/createLink`).post({ type, scope }) as GraphEntity;
  }

  async deleteDriveItem(itemId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/drive/items/${itemId}`).delete();
  }

  // ===========================================================================
  // SharePoint Sites & Document Libraries
  // ===========================================================================

  /**
   * Lists sites the current user follows.
   */
  async listFollowedSites(): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const response = await client.api('/me/followedSites').get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  /**
   * Searches for SharePoint sites by keyword.
   */
  async searchSites(query: string): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const response = await client.api(`/sites?search=${encodeURIComponent(query)}`).get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  /**
   * Gets a specific SharePoint site by ID.
   */
  async getSite(siteId: string): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client.api(`/sites/${siteId}`).get() as GraphEntity;
  }

  /**
   * Lists document libraries (drives) for a SharePoint site.
   */
  async listDocumentLibraries(siteId: string): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const response = await client.api(`/sites/${siteId}/drives`).get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  /**
   * Lists items in a document library or folder.
   */
  async listLibraryItems(driveId: string, itemId?: string): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const apiPath = itemId != null ? `/drives/${driveId}/items/${itemId}/children` : `/drives/${driveId}/root/children`;
    const response = await client.api(apiPath).get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  /**
   * Downloads a file from a document library.
   */
  async downloadLibraryFile(driveId: string, itemId: string): Promise<ArrayBuffer> {
    const client = await this.getClient();
    return await client.api(`/drives/${driveId}/items/${itemId}/content`).responseType(ResponseType.ARRAYBUFFER).get() as ArrayBuffer;
  }

  /**
   * Creates a folder in a document library. Omit parentItemId to create at the
   * library root. conflictBehavior is `fail` (default) or `rename`.
   */
  async createLibraryFolder(driveId: string, parentItemId: string | undefined, folderName: string, conflictBehavior: string): Promise<GraphEntity> {
    const client = await this.getClient();
    const apiPath = parentItemId != null ? `/drives/${driveId}/items/${parentItemId}/children` : `/drives/${driveId}/root/children`;
    return await client.api(apiPath).post({
      name: folderName,
      folder: {},
      '@microsoft.graph.conflictBehavior': conflictBehavior,
    }) as GraphEntity;
  }

  /**
   * Uploads a file into a document library via a simple PUT. Omit parentItemId to
   * upload at the library root. conflictBehavior is `fail` (default), `replace`,
   * or `rename`. Simple upload is limited by Microsoft Graph to 4 MB.
   */
  async uploadLibraryFile(driveId: string, parentItemId: string | undefined, fileName: string, content: Buffer, conflictBehavior: string): Promise<GraphEntity> {
    const client = await this.getClient();
    const encodedName = encodeURIComponent(fileName);
    const apiPath = parentItemId != null
      ? `/drives/${driveId}/items/${parentItemId}:/${encodedName}:/content`
      : `/drives/${driveId}/root:/${encodedName}:/content`;
    return await client.api(apiPath)
      .query({ '@microsoft.graph.conflictBehavior': conflictBehavior })
      .header('Content-Type', 'application/octet-stream')
      .put(content) as GraphEntity;
  }

  // ===========================================================================
  // SharePoint Lists
  // ===========================================================================

  /**
   * Lists the SharePoint lists in a site.
   */
  async listSharePointLists(siteId: string): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const response = await client
      .api(`/sites/${siteId}/lists`)
      .select('id,name,displayName,description,webUrl,createdDateTime,lastModifiedDateTime')
      .get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  /**
   * Gets a specific SharePoint list.
   */
  async getSharePointList(siteId: string, listId: string): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client.api(`/sites/${siteId}/lists/${listId}`).get() as GraphEntity;
  }

  /**
   * Creates a SharePoint list in a site.
   */
  async createSharePointList(siteId: string, body: Record<string, unknown>): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client.api(`/sites/${siteId}/lists`).post(body) as GraphEntity;
  }

  /**
   * Lists the column definitions for a SharePoint list.
   */
  async listSharePointListColumns(siteId: string, listId: string): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const response = await client
      .api(`/sites/${siteId}/lists/${listId}/columns`)
      .get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  /**
   * Lists the items in a SharePoint list, expanding their field values.
   */
  async listSharePointListItems(siteId: string, listId: string, limit: number = 50): Promise<GraphEntity[]> {
    const client = await this.getClient();
    const response = await client
      .api(`/sites/${siteId}/lists/${listId}/items`)
      .expand('fields')
      .top(limit)
      .get() as GraphCollectionResponse<GraphEntity>;
    return response.value;
  }

  /**
   * Gets a specific SharePoint list item, expanding its field values.
   */
  async getSharePointListItem(siteId: string, listId: string, itemId: string): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client
      .api(`/sites/${siteId}/lists/${listId}/items/${itemId}`)
      .expand('fields')
      .get() as GraphEntity;
  }

  /**
   * Creates an item in a SharePoint list from a map of column → value.
   */
  async createSharePointListItem(siteId: string, listId: string, fields: Record<string, unknown>): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client.api(`/sites/${siteId}/lists/${listId}/items`).post({ fields }) as GraphEntity;
  }

  /**
   * Updates the field values of a SharePoint list item.
   */
  async updateSharePointListItem(siteId: string, listId: string, itemId: string, fields: Record<string, unknown>): Promise<GraphEntity> {
    const client = await this.getClient();
    return await client.api(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`).patch(fields) as GraphEntity;
  }

  /**
   * Deletes an item from a SharePoint list.
   */
  async deleteSharePointListItem(siteId: string, listId: string, itemId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/sites/${siteId}/lists/${listId}/items/${itemId}`).delete();
  }

  // ===========================================================================
  // OneNote
  // ===========================================================================

  /**
   * Lists all OneNote notebooks for the current user.
   */
  async listNotebooks(): Promise<MicrosoftGraph.Notebook[]> {
    const client = await this.getClient();
    const response = await client.api('/me/onenote/notebooks').get() as PageCollection;
    return response.value as MicrosoftGraph.Notebook[];
  }

  /**
   * Lists OneNote sections, optionally scoped to a notebook.
   */
  async listNoteSections(notebookGraphId?: string): Promise<MicrosoftGraph.OnenoteSection[]> {
    const client = await this.getClient();
    const apiPath = notebookGraphId != null
      ? `/me/onenote/notebooks/${notebookGraphId}/sections`
      : '/me/onenote/sections';
    const response = await client.api(apiPath).get() as PageCollection;
    return response.value as MicrosoftGraph.OnenoteSection[];
  }

  /**
   * Lists OneNote pages, optionally scoped to a section.
   */
  async listNotePages(sectionGraphId?: string): Promise<MicrosoftGraph.OnenotePage[]> {
    const client = await this.getClient();
    if (sectionGraphId != null) {
      const response = await client.api(`/me/onenote/sections/${sectionGraphId}/pages`).get() as PageCollection;
      return response.value as MicrosoftGraph.OnenotePage[];
    }
    const response = await client
      .api('/me/onenote/pages')
      .top(50)
      .orderby('lastModifiedDateTime desc')
      .get() as PageCollection;
    return response.value as MicrosoftGraph.OnenotePage[];
  }

  /**
   * Gets a OneNote page's metadata.
   */
  async getNotePage(pageGraphId: string): Promise<MicrosoftGraph.OnenotePage> {
    const client = await this.getClient();
    return await client.api(`/me/onenote/pages/${pageGraphId}`).get() as MicrosoftGraph.OnenotePage;
  }

  /**
   * Gets a OneNote page's HTML content.
   */
  async getNotePageContent(pageGraphId: string): Promise<string> {
    const client = await this.getClient();
    return await client
      .api(`/me/onenote/pages/${pageGraphId}/content`)
      .responseType(ResponseType.TEXT)
      .get() as string;
  }

  /**
   * Searches OneNote pages by keyword.
   */
  async searchNotePages(query: string): Promise<MicrosoftGraph.OnenotePage[]> {
    const client = await this.getClient();
    // Strip embedded double-quotes so they can't unbalance the quoted $search
    // expression (`$search="..."`) and 400 the request.
    const safeQuery = query.replace(/"/g, '');
    const response = await client
      .api('/me/onenote/pages')
      .search(`"${safeQuery}"`)
      .get() as PageCollection;
    return response.value as MicrosoftGraph.OnenotePage[];
  }

  /**
   * Creates a new OneNote page in a section from raw HTML.
   */
  async createNotePage(sectionGraphId: string, html: string): Promise<MicrosoftGraph.OnenotePage> {
    const client = await this.getClient();
    return await client
      .api(`/me/onenote/sections/${sectionGraphId}/pages`)
      .header('Content-Type', 'text/html')
      .post(html) as MicrosoftGraph.OnenotePage;
  }
}

/**
 * Creates a new Graph client instance.
 */
export function createGraphClient(deviceCodeCallback?: DeviceCodeCallback): GraphClient {
  return new GraphClient(deviceCodeCallback);
}
