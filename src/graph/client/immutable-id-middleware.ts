/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Immutable-ID preference middleware (U5b-3 / D2).
 *
 * Adds `Prefer: IdType="ImmutableId"` to Outlook/Exchange requests so Graph
 * returns immutable item IDs — IDs that survive a move between folders. The
 * self-encoding tokens (`em_ ev_ ct_ fd_`) carry the returned Graph ID directly,
 * so minting them from immutable IDs is what makes those tokens durable across a
 * move rather than breaking the moment an item is filed elsewhere.
 *
 * Scope: the header is applied only to Outlook resource paths (messages, mail
 * folders, events, calendars, contacts, contact folders). It is deliberately NOT
 * applied to `$search` requests: Graph does not return immutable IDs for
 * `$search`, so search-minted IDs are upgraded separately via
 * `translateExchangeIds` (see graph-client.ts).
 */

import type { Context, FetchOptions, Middleware } from '@microsoft/microsoft-graph-client';

const PREFER_HEADER = 'Prefer';
const IMMUTABLE_ID_PREFERENCE = 'IdType="ImmutableId"';

/**
 * Outlook resource path segments that support immutable IDs. Matching any of
 * these as a whole path segment (not a substring) opts the request in — this
 * covers nested collections too (e.g. `/me/mailFolders/{id}/messages`,
 * `/me/calendars/{id}/events`, `/me/contactFolders/{id}/contacts`).
 */
const OUTLOOK_SEGMENTS: ReadonlySet<string> = new Set([
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
 * True when the request targets an Outlook resource that supports immutable IDs
 * AND is not a `$search` request (immutable IDs are unsupported for `$search`,
 * and pairing the two is avoided rather than relied upon).
 */
export function isOutlookImmutableIdResource(url: string): boolean {
  const queryIndex = url.indexOf('?');
  const path = (queryIndex === -1 ? url : url.slice(0, queryIndex)).toLowerCase();
  const query = queryIndex === -1 ? '' : url.slice(queryIndex + 1).toLowerCase();
  if (query.includes('$search')) {
    return false;
  }
  for (const segment of path.split('/')) {
    if (OUTLOOK_SEGMENTS.has(segment)) {
      return true;
    }
  }
  return false;
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
 * Graph SDK middleware that opts Outlook reads into immutable IDs. Sits directly
 * after the auth handler and before the retry handler, so the header is set once
 * and rides every retry attempt without being re-appended.
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
