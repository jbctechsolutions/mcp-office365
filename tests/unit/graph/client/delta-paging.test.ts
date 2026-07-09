/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Regression tests for delta cursor paging (U12): the `@odata.deltaLink` lands
 * on the LAST page, so a multi-page delta must read it from the final response,
 * not the first. Reading it from the first page yields '' and pins the delta
 * mirror in a permanent re-baseline loop.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';

interface Page { value: unknown[]; '@odata.nextLink'?: string; '@odata.deltaLink'?: string }

/** Responses served in order per `.get()` call, regardless of URL. */
let pages: Page[];
const seenUrls: string[] = [];
const seenHeaders: Array<Record<string, string>> = [];

function builder(): unknown {
  const headers: Record<string, string> = {};
  const b: Record<string, unknown> = {
    select: () => b,
    top: () => b,
    header: (k: string, v: string) => { headers[k] = v; return b; },
    get: () => {
      seenHeaders.push({ ...headers });
      const next = pages.shift();
      return Promise.resolve(next ?? { value: [] });
    },
  };
  return b;
}

const mockApi = vi.fn((url: string) => { seenUrls.push(url); return builder(); });

vi.mock('@microsoft/microsoft-graph-client', () => ({
  Client: {
    init: vi.fn(() => ({ api: mockApi })),
    initWithMiddleware: vi.fn(() => ({ api: mockApi })),
  },
  AuthenticationHandler: class { setNext(): void {} },
  RetryHandler: class { setNext(): void {} },
  RetryHandlerOptions: class {},
  RedirectHandler: class { setNext(): void {} },
  HTTPMessageHandler: class { setNext(): void {} },
  ResponseType: { ARRAYBUFFER: 'arraybuffer', JSON: 'json' },
}));

vi.mock('../../../../src/graph/auth/index.js', () => ({
  getAccessToken: vi.fn().mockResolvedValue('test-access-token'),
}));

vi.mock('isomorphic-fetch', () => ({ default: vi.fn() }));

import { GraphClient } from '../../../../src/graph/client/graph-client.js';

beforeEach(() => {
  pages = [];
  seenUrls.length = 0;
  seenHeaders.length = 0;
});

describe('getMessagesDelta paging', () => {
  it('reads the deltaLink from the last page of a multi-page delta', async () => {
    pages = [
      { value: [{ id: 'a' }], '@odata.nextLink': 'https://graph/next' },
      { value: [{ id: 'b' }], '@odata.deltaLink': 'https://graph/delta-final' },
    ];
    const client = new GraphClient();

    const { messages, deltaLink } = await client.getMessagesDelta('inbox');

    expect(messages.map((m) => m.id)).toEqual(['a', 'b']);
    expect(deltaLink).toBe('https://graph/delta-final');
    expect(seenUrls[0]).toContain('/me/mailFolders/inbox/messages/delta');
    expect(seenUrls[1]).toBe('https://graph/next');
  });
});

describe('getCalendarViewDelta paging', () => {
  it('sends the window + Prefer header and reads the deltaLink from the last page', async () => {
    pages = [
      { value: [{ id: 'e1' }], '@odata.nextLink': 'https://graph/cal-next' },
      { value: [{ id: 'e2' }], '@odata.deltaLink': 'https://graph/cal-delta-final' },
    ];
    const client = new GraphClient();

    const { events, deltaLink } = await client.getCalendarViewDelta('2026-01-01T00:00:00Z', '2026-04-01T00:00:00Z');

    expect(events.map((e) => e.id)).toEqual(['e1', 'e2']);
    expect(deltaLink).toBe('https://graph/cal-delta-final');
    expect(seenUrls[0]).toContain('/me/calendarView/delta?startDateTime=');
    expect(seenHeaders[0]?.Prefer).toBe('odata.maxpagesize=50');
  });

  it('follows a stored deltaLink directly without re-sending the window', async () => {
    pages = [{ value: [{ id: 'e3' }], '@odata.deltaLink': 'https://graph/cal-delta-2' }];
    const client = new GraphClient();

    const { deltaLink } = await client.getCalendarViewDelta('x', 'y', 'https://graph/stored-cursor');

    expect(seenUrls[0]).toBe('https://graph/stored-cursor');
    expect(deltaLink).toBe('https://graph/cal-delta-2');
  });
});
