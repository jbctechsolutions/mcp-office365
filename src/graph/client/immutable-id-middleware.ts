/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Immutable-ID preference middleware (U5b-3 / D2).
 *
 * Adds `Prefer: IdType="ImmutableId"` to Outlook item requests so Graph returns
 * immutable item IDs — IDs that survive a move between folders. The self-encoding
 * tokens (`em_ ev_ ct_`) carry the returned Graph ID verbatim, so minting them
 * from immutable IDs is what makes those tokens durable across a move rather than
 * breaking the moment an item is filed elsewhere.
 *
 * Scope is decided by ANCHORING on the Outlook user context: the collection
 * immediately after `/me/` (or after the id in `/users/{id}/`) must be an Outlook
 * collection. This is what keeps Teams and chat out — `/teams/{id}/channels/{id}/
 * messages` has no user anchor, and `/me/chats/{id}/messages` anchors on `chats`,
 * not an Outlook collection — even though both carry a `messages` segment. It also
 * keeps To Do (`/me/todo/.../attachments`) and Planner out for the same reason,
 * while covering nested item reads (`/me/mailFolders/{id}/messages`) and shared
 * mailboxes (`/users/{upn}/messages`, #40).
 *
 * The header is NOT applied to `$search` requests: Graph does not return immutable
 * IDs for `$search`, so search-minted IDs are upgraded separately via
 * `translateExchangeIds` (see graph-client.ts).
 *
 * Note: the header only shapes the RESPONSE id format — Graph accepts either the
 * default or immutable id form in the request URL regardless of the header
 * (verified live), so applying it unconditionally to Outlook item requests never
 * breaks resolution of an already-minted token.
 */

import type { Context, FetchOptions, Middleware } from '@microsoft/microsoft-graph-client';

const PREFER_HEADER = 'Prefer';
const IMMUTABLE_ID_PREFERENCE = 'IdType="ImmutableId"';

/**
 * Outlook collections that hang off a user root (`/me` or `/users/{id}`) and lead
 * to immutable-ID-bearing items. Container collections (mailFolders, calendars,
 * calendarGroups, contactFolders) are included so a nested item read like
 * `/me/mailFolders/{id}/messages` — which anchors on `mailfolders` — opts in; the
 * header is harmlessly ignored on a bare container GET (their IDs are already
 * constant).
 */
const OUTLOOK_ROOT_COLLECTIONS: ReadonlySet<string> = new Set([
  'messages',
  'mailfolders',
  'events',
  'calendar',
  'calendars',
  'calendarview',
  'calendargroups',
  'contacts',
  'contactfolders',
]);

/**
 * Matches the `$search` system query parameter in raw (`$search=`) or
 * percent-encoded (`%24search=`) form, anchored to a param boundary (`^` or `&`)
 * so it does not false-positive on a `$filter` VALUE that merely contains the
 * text "search" (that appears as `…%24search%27`, without the trailing `=`).
 */
const SEARCH_PARAM = /(^|&)(\$|%24)search=/;

/**
 * True when the request targets an Outlook item resource that supports immutable
 * IDs — anchored to the `/me` or `/users/{id}` user context, and excluding
 * `$search` requests (immutable IDs are unsupported for `$search`, and pairing the
 * two is avoided rather than relied upon).
 */
export function isOutlookImmutableIdResource(url: string): boolean {
  const queryIndex = url.indexOf('?');
  const path = (queryIndex === -1 ? url : url.slice(0, queryIndex)).toLowerCase();
  const query = queryIndex === -1 ? '' : url.slice(queryIndex + 1).toLowerCase();
  if (SEARCH_PARAM.test(query)) {
    return false;
  }
  const segments = path.split('/').filter((segment) => segment.length > 0);
  // Anchor on the Outlook user context: the collection right after `me`
  // (`/me/{collection}`) or after the id in `/users/{id}/{collection}`.
  let collection: string | undefined;
  const meIndex = segments.indexOf('me');
  if (meIndex !== -1) {
    collection = segments[meIndex + 1];
  } else {
    const usersIndex = segments.indexOf('users');
    if (usersIndex !== -1) {
      collection = segments[usersIndex + 2];
    }
  }
  return collection != null && OUTLOOK_ROOT_COLLECTIONS.has(collection);
}

/** True when a Prefer value already carries an IdType preference. */
function hasIdTypePreference(value: string): boolean {
  return value.toLowerCase().includes('idtype');
}

/** Appends the immutable-ID preference to an existing Prefer value. */
function appendPreference(existing: string): string {
  return existing.length > 0 ? `${existing}, ${IMMUTABLE_ID_PREFERENCE}` : IMMUTABLE_ID_PREFERENCE;
}

/**
 * Adds (or merges into) the `Prefer` header on the request options, preserving
 * any preference already set (e.g. `odata.maxpagesize`) and never doubling the
 * IdType token if it is somehow already present.
 */
function addImmutableIdPreference(options: FetchOptions): void {
  const headers = options.headers;

  if (headers instanceof Headers) {
    const existing = headers.get(PREFER_HEADER);
    if (existing == null) {
      headers.set(PREFER_HEADER, IMMUTABLE_ID_PREFERENCE);
    } else if (!hasIdTypePreference(existing)) {
      headers.set(PREFER_HEADER, appendPreference(existing));
    }
    return;
  }

  if (Array.isArray(headers)) {
    const entry = headers.find((pair) => (pair[0] ?? '').toLowerCase() === PREFER_HEADER.toLowerCase());
    if (entry == null) {
      headers.push([PREFER_HEADER, IMMUTABLE_ID_PREFERENCE]);
    } else {
      const existing = entry[1] ?? '';
      if (!hasIdTypePreference(existing)) {
        entry[1] = appendPreference(existing);
      }
    }
    return;
  }

  const record = (headers ?? {}) as Record<string, string>;
  const key = Object.keys(record).find((k) => k.toLowerCase() === PREFER_HEADER.toLowerCase());
  if (key == null) {
    record[PREFER_HEADER] = IMMUTABLE_ID_PREFERENCE;
  } else {
    const existing = record[key] ?? '';
    if (!hasIdTypePreference(existing)) {
      record[key] = appendPreference(existing);
    }
  }
  options.headers = record;
}

/**
 * Graph SDK middleware that opts Outlook item reads into immutable IDs. Sits
 * directly after the auth handler and before the retry handler, so the header is
 * set once and rides every retry attempt without being re-appended.
 */
export class ImmutableIdMiddleware implements Middleware {
  private nextMiddleware: Middleware | undefined;

  public async execute(context: Context): Promise<void> {
    // `context.request` is typed as RequestInfo, which resolves to `any` under
    // this repo's lib config; narrow to the shape we read to stay lint-clean.
    const request = context.request as string | { url?: unknown };
    const url = typeof request === 'string' ? request : String(request.url ?? '');
    if (isOutlookImmutableIdResource(url)) {
      context.options = context.options ?? {};
      addImmutableIdPreference(context.options);
    }
    if (this.nextMiddleware != null) {
      await this.nextMiddleware.execute(context);
    }
  }

  public setNext(next: Middleware): void {
    this.nextMiddleware = next;
  }
}
