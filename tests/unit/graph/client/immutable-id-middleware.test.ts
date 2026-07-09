/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Behavioral tests for the immutable-ID preference middleware (U5b-3 / D2):
 * which requests get `Prefer: IdType="ImmutableId"`, how the header merges with
 * an existing Prefer value, and that the chain is always forwarded.
 */

import { describe, it, expect, vi } from 'vitest';
import type { Context, Middleware } from '@microsoft/microsoft-graph-client';
import {
  ImmutableIdMiddleware,
  isOutlookImmutableIdResource,
} from '../../../../src/graph/client/immutable-id-middleware.js';

const BASE = 'https://graph.microsoft.com/v1.0';

/** A tail middleware that records the context it was handed. */
function recordingNext(): { middleware: Middleware; last: () => Context | undefined } {
  let seen: Context | undefined;
  return {
    middleware: {
      execute: vi.fn((context: Context) => {
        seen = context;
        return Promise.resolve();
      }),
    },
    last: () => seen,
  };
}

/** Reads the Prefer header out of whatever HeadersInit shape options carries. */
function readPrefer(context: Context): string | undefined {
  const headers = context.options?.headers;
  if (headers == null) return undefined;
  if (headers instanceof Headers) return headers.get('Prefer') ?? undefined;
  if (Array.isArray(headers)) return headers.find(([k]) => k.toLowerCase() === 'prefer')?.[1];
  const rec = headers as Record<string, string>;
  const key = Object.keys(rec).find((k) => k.toLowerCase() === 'prefer');
  return key != null ? rec[key] : undefined;
}

async function run(request: string, options?: Context['options']): Promise<Context> {
  const mw = new ImmutableIdMiddleware();
  const next = recordingNext();
  mw.setNext(next.middleware);
  const context: Context = options != null ? { request, options } : { request };
  await mw.execute(context);
  expect(next.middleware.execute).toHaveBeenCalledOnce();
  return context;
}

describe('isOutlookImmutableIdResource', () => {
  it('matches Outlook resources and their nested collections', () => {
    for (const url of [
      `${BASE}/me/messages`,
      `${BASE}/me/messages/AAA`,
      `${BASE}/me/mailFolders/inbox/messages`,
      `${BASE}/me/events`,
      `${BASE}/me/calendars/cal-1/events`,
      `${BASE}/me/calendarView?startDateTime=x&endDateTime=y`,
      `${BASE}/me/calendarGroups`,
      `${BASE}/me/contacts`,
      `${BASE}/me/contactFolders/cf-1/contacts`,
    ]) {
      expect(isOutlookImmutableIdResource(url)).toBe(true);
    }
  });

  it('does not match non-Outlook resources', () => {
    for (const url of [
      `${BASE}/me/todo/lists/l-1/tasks`,
      `${BASE}/teams/t-1/channels`,
      `${BASE}/me/drive/items/i-1`,
      `${BASE}/search/query`,
      `${BASE}/me/translateExchangeIds`,
      `${BASE}/planner/tasks/p-1`,
    ]) {
      expect(isOutlookImmutableIdResource(url)).toBe(false);
    }
  });

  it('excludes $search requests (immutable IDs are unsupported for $search)', () => {
    expect(isOutlookImmutableIdResource(`${BASE}/me/messages?$search="report"`)).toBe(false);
    expect(isOutlookImmutableIdResource(`${BASE}/me/mailFolders/inbox/messages?$search="x"&$top=50`)).toBe(false);
  });
});

describe('ImmutableIdMiddleware.execute', () => {
  it('adds the Prefer header on an Outlook request with no options', async () => {
    const context = await run(`${BASE}/me/messages/AAA`);
    expect(readPrefer(context)).toBe('IdType="ImmutableId"');
  });

  it('does not add the header on a non-Outlook request', async () => {
    const context = await run(`${BASE}/me/drive/items/i-1`);
    expect(readPrefer(context)).toBeUndefined();
  });

  it('does not add the header on a $search request', async () => {
    const context = await run(`${BASE}/me/messages?$search="q"`);
    expect(readPrefer(context)).toBeUndefined();
  });

  it('merges with an existing Prefer value (record headers)', async () => {
    const context = await run(`${BASE}/me/calendarView`, {
      headers: { Prefer: 'odata.maxpagesize=50' },
    });
    expect(readPrefer(context)).toBe('odata.maxpagesize=50, IdType="ImmutableId"');
  });

  it('merges into a Headers instance', async () => {
    const headers = new Headers({ Prefer: 'outlook.timezone="UTC"' });
    const context = await run(`${BASE}/me/events`, { headers });
    expect(readPrefer(context)).toBe('outlook.timezone="UTC", IdType="ImmutableId"');
  });

  it('merges into array-style headers', async () => {
    const context = await run(`${BASE}/me/messages`, {
      headers: [['Prefer', 'odata.maxpagesize=10']] as [string, string][],
    });
    expect(readPrefer(context)).toBe('odata.maxpagesize=10, IdType="ImmutableId"');
  });

  it('does not double the IdType preference if already present', async () => {
    const context = await run(`${BASE}/me/messages`, {
      headers: { prefer: 'IdType="ImmutableId"' },
    });
    expect(readPrefer(context)).toBe('IdType="ImmutableId"');
  });

  it('forwards even when there is no next middleware', async () => {
    const mw = new ImmutableIdMiddleware();
    await expect(mw.execute({ request: `${BASE}/me/messages` })).resolves.toBeUndefined();
  });
});
