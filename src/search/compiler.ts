/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Structured email-search compiler (U7 / D9). Replaces raw KQL — the #1 observed
 * failure class — with typed params compiled server-side, so correct operator
 * syntax, quoting, and date formatting are code, not model guesswork.
 *
 * Graph makes `$filter` and `$search` mutually exclusive on messages, so:
 * - property-only criteria  → `$filter`
 * - free-text-only criteria → quoted `$search`
 * - mixed                   → server-built KQL for `POST /search/query`
 *   (the only single-request path for both at once).
 *
 * Field → mechanism classification below is the documented default; per D9's
 * execution note it is confirmed/refined by a live-mailbox spike before release
 * (Graph's `$filter` has no `contains` on message subject/body, so those are
 * treated as free-text here).
 */

import { ValidationError } from '../utils/errors.js';

/**
 * Structured, typed email-search criteria (replaces the raw `query` KQL). Fields
 * are `| undefined`-tolerant so a zod-parsed input object (which carries explicit
 * `undefined` for absent optionals under exactOptionalPropertyTypes) passes
 * directly; every field is guarded with `!= null` before use.
 */
export interface EmailSearchParams {
  /** Sender address (exact). Property → `$filter`. */
  from?: string | undefined;
  /** Recipient address (exact). Property → `$filter`. */
  to?: string | undefined;
  /** Received on/after this ISO date/datetime. Property → `$filter`. */
  received_after?: string | undefined;
  /** Received on/before this ISO date/datetime. Property → `$filter`. */
  received_before?: string | undefined;
  /** Only messages with attachments. Property → `$filter`. */
  has_attachments?: boolean | undefined;
  /** Only unread messages. Property → `$filter`. */
  is_unread?: boolean | undefined;
  /** Importance level. Property → `$filter`. */
  importance?: 'low' | 'normal' | 'high' | undefined;
  /** Subject contains (free-text — Graph `$filter` has no `contains` here). */
  subject_contains?: string | undefined;
  /** Body contains (free-text). */
  body_contains?: string | undefined;
  /** Free-text across the message. */
  text?: string | undefined;
}

/** The compiled query and which Graph mechanism executes it. */
export type CompiledSearch =
  | { mechanism: 'filter'; filter: string }
  | { mechanism: 'search'; search: string }
  | { mechanism: 'searchQuery'; kql: string };

const ISO_DATE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?)?$/;

/**
 * Compiles structured params into the correct Graph mechanism + query string.
 * Throws {@link ValidationError} when no criterion is given or a date is malformed.
 */
export function compileEmailSearch(params: EmailSearchParams): CompiledSearch {
  const filters = buildFilters(params);
  const terms = buildSearchTerms(params);

  if (filters.length === 0 && terms.length === 0) {
    throw new ValidationError(
      'Provide at least one search criterion (e.g. from, subject_contains, received_after, text).',
    );
  }

  if (terms.length === 0) {
    return { mechanism: 'filter', filter: filters.join(' and ') };
  }
  if (filters.length === 0) {
    return { mechanism: 'search', search: quoteSearch(terms) };
  }
  return { mechanism: 'searchQuery', kql: buildKql(params) };
}

/** Property criteria → OData `$filter` clauses (order is stable for snapshots). */
function buildFilters(p: EmailSearchParams): string[] {
  const out: string[] = [];
  if (p.from != null && p.from.length > 0) {
    out.push(`from/emailAddress/address eq '${escapeODataString(p.from)}'`);
  }
  if (p.to != null && p.to.length > 0) {
    out.push(`toRecipients/any(r:r/emailAddress/address eq '${escapeODataString(p.to)}')`);
  }
  if (p.received_after != null) {
    out.push(`receivedDateTime ge ${requireIsoDate(p.received_after, 'received_after')}`);
  }
  if (p.received_before != null) {
    out.push(`receivedDateTime le ${requireIsoDate(p.received_before, 'received_before')}`);
  }
  if (p.has_attachments != null) {
    out.push(`hasAttachments eq ${p.has_attachments ? 'true' : 'false'}`);
  }
  if (p.is_unread != null) {
    out.push(`isRead eq ${p.is_unread ? 'false' : 'true'}`);
  }
  if (p.importance != null) {
    out.push(`importance eq '${p.importance}'`);
  }
  return out;
}

/** Free-text criteria → search terms (subject/body `contains` isn't a filter). */
function buildSearchTerms(p: EmailSearchParams): string[] {
  const out: string[] = [];
  if (p.text != null && p.text.trim().length > 0) {
    out.push(p.text.trim());
  }
  if (p.subject_contains != null && p.subject_contains.trim().length > 0) {
    out.push(`subject:${p.subject_contains.trim()}`);
  }
  if (p.body_contains != null && p.body_contains.trim().length > 0) {
    out.push(`body:${p.body_contains.trim()}`);
  }
  return out;
}

/**
 * Builds the `$search` value: each term wrapped in a quoted phrase, AND-joined.
 * Embedded double quotes are STRIPPED (not backslash-escaped): Graph's $search
 * has no documented in-phrase escape, so a `"` in a value could otherwise close
 * the phrase early and inject a boolean/property clause.
 */
function quoteSearch(terms: string[]): string {
  return terms.map((t) => kqlPhrase(t)).join(' AND ');
}

/**
 * Builds a KQL string for `POST /search/query` combining property and free-text
 * criteria (the only single-request path when both are present). EVERY free
 * value — including from/to — is wrapped as a quoted phrase so a value with
 * whitespace or KQL operators (e.g. `alice@x.com AND body:secret`) cannot inject
 * extra clauses. Dates are day-granular here (KQL range syntax); the $filter
 * path preserves time-of-day, so mixed-mode queries are intentionally
 * day-granular on dates (documented on the tool in U7b).
 */
function buildKql(p: EmailSearchParams): string {
  const clauses: string[] = [];
  if (p.from != null && p.from.length > 0) clauses.push(`from:${kqlPhrase(p.from)}`);
  if (p.to != null && p.to.length > 0) clauses.push(`to:${kqlPhrase(p.to)}`);
  if (p.received_after != null) clauses.push(`received>=${toKqlDate(p.received_after, 'received_after')}`);
  if (p.received_before != null) clauses.push(`received<=${toKqlDate(p.received_before, 'received_before')}`);
  if (p.has_attachments != null) clauses.push(`hasattachment:${p.has_attachments ? 'true' : 'false'}`);
  if (p.is_unread === true) clauses.push('isread:false');
  else if (p.is_unread === false) clauses.push('isread:true');
  if (p.importance != null) clauses.push(`importance:${p.importance}`);
  if (p.subject_contains != null && p.subject_contains.trim().length > 0) {
    clauses.push(`subject:${kqlPhrase(p.subject_contains.trim())}`);
  }
  if (p.body_contains != null && p.body_contains.trim().length > 0) {
    clauses.push(`body:${kqlPhrase(p.body_contains.trim())}`);
  }
  if (p.text != null && p.text.trim().length > 0) {
    clauses.push(kqlPhrase(p.text.trim()));
  }
  return clauses.join(' AND ');
}

function requireIsoDate(value: string, field: string): string {
  if (!ISO_DATE.test(value) || Number.isNaN(Date.parse(value))) {
    throw invalidDate(field, value);
  }
  // Reject rollover-invalid calendar dates: Date.parse('2026-02-30') is NOT NaN
  // (it rolls to Mar 2), so re-serialize the date part and require an exact
  // match to catch day/month overflow.
  const datePart = value.slice(0, 10);
  if (new Date(`${datePart}T00:00:00Z`).toISOString().slice(0, 10) !== datePart) {
    throw invalidDate(field, value);
  }
  // A bare date becomes midnight UTC so `ge`/`le` compare against a full instant.
  return value.includes('T') ? value : `${value}T00:00:00Z`;
}

function invalidDate(field: string, value: string): ValidationError {
  return new ValidationError(
    `${field} must be a valid ISO date (YYYY-MM-DD) or datetime; got "${value}".`,
  );
}

function toKqlDate(value: string, field: string): string {
  requireIsoDate(value, field);
  return value.slice(0, 10);
}

/** Escapes single quotes for an OData string literal (doubled per OData rules). */
function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

/** Wraps a value as a quoted KQL phrase, stripping embedded double quotes. */
function kqlPhrase(value: string): string {
  return `"${stripQuotes(value)}"`;
}

function stripQuotes(value: string): string {
  return value.replace(/"/g, '');
}
