/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Behavioral tests for the immutable-ID preference middleware (U5b-3 / D2):
 * which requests get `Prefer: IdType="ImmutableId"` (anchored on the Outlook user
 * context so Teams/chat/To Do are excluded), how the header merges with an
 * existing Prefer value, the skip-marker opt-out, and chain forwarding.
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

/** Reads a header (case-insensitive) out of whatever HeadersInit shape options carries. */
function readHeader(context: Context, name: string): string | undefined {
  const headers = context.options?.headers;
  const lower = name.toLowerCase();
  if (headers == null) return undefined;
  if (headers instanceof Headers) return headers.get(name) ?? undefined;
  if (Array.isArray(headers)) return headers.find(([k]) => k.toLowerCase() === lower)?.[1];
  const rec = headers as Record<string, string>;
  const key = Object.keys(rec).find((k) => k.toLowerCase() === lower);
  return key != null ? rec[key] : undefined;
}

function readPrefer(context: Context): string | undefined {
  return readHeader(context, 'Prefer');
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
  it('matches Outlook item resources anchored on /me or /users', () => {
    for (const url of [
      `${BASE}/me/messages`,
      `${BASE}/me/messages/AAA`,
      `${BASE}/me/mailFolders/inbox/messages`, // nested — anchors on mailFolders
      `${BASE}/me/mailFolders/inbox`, // bare container — header harmlessly ignored
      `${BASE}/me/events`,
      `${BASE}/me/events/evt-1/instances`,
      `${BASE}/me/messages/AAA/attachments`,
      `${BASE}/me/calendars/cal-1/events`,
      `${BASE}/me/calendarView?startDateTime=x&endDateTime=y`,
      `${BASE}/me/calendarGroups`,
      `${BASE}/me/contacts`,
      `${BASE}/me/contactFolders/cf-1/contacts`,
      `${BASE}/users/bob@contoso.com/messages`, // shared mailbox (#40)
      `${BASE}/users/bob@contoso.com/mailFolders/inbox/messages`,
    ]) {
      expect(isOutlookImmutableIdResource(url)).toBe(true);
    }
  });

  it('excludes Teams and chat message endpoints (no Outlook user anchor)', () => {
    for (const url of [
      `${BASE}/teams/t-1/channels/c-1/messages`,
      `${BASE}/teams/t-1/channels/c-1/messages/m-1`,
      `${BASE}/teams/t-1/channels/c-1/messages/m-1/replies`,
      `${BASE}/chats/c-1/messages`,
      `${BASE}/chats/c-1/messages/m-1/setReaction`,
      `${BASE}/me/chats/c-1/messages`, // under /me/ but anchors on `chats`
      `${BASE}/me/chats/c-1/messages/m-1`,
    ]) {
      expect(isOutlookImmutableIdResource(url)).toBe(false);
    }
  });

  it('excludes To Do and other non-Outlook resources', () => {
    for (const url of [
      `${BASE}/me/todo/lists/l-1/tasks/t-1/attachments`, // anchors on `todo`, not attachments
      `${BASE}/me/todo/lists/l-1/tasks`,
      `${BASE}/me/drive/items/i-1`,
      `${BASE}/search/query`,
      `${BASE}/me/translateExchangeIds`,
      `${BASE}/planner/tasks/p-1`,
    ]) {
      expect(isOutlookImmutableIdResource(url)).toBe(false);
    }
  });

  it('excludes $search requests, raw and percent-encoded', () => {
    expect(isOutlookImmutableIdResource(`${BASE}/me/messages?$search="report"`)).toBe(false);
    expect(isOutlookImmutableIdResource(`${BASE}/me/messages?%24search=%22report%22`)).toBe(false);
    expect(isOutlookImmutableIdResource(`${BASE}/me/mailFolders/inbox/messages?$search="x"&$top=50`)).toBe(false);
  });

  it('does NOT treat a $filter value containing "search" text as a $search request', () => {
    // The `$` in the filter value is percent-encoded (%24) but is not a param —
    // no trailing `=`, not at a param boundary — so the header still applies.
    expect(isOutlookImmutableIdResource(`${BASE}/me/messages?$filter=contains(subject,'%24search')`)).toBe(true);
  });
});

describe('ImmutableIdMiddleware.execute', () => {
  it('adds the Prefer header on an Outlook item request with no options', async () => {
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

  it('does not add the header on a Teams channel messages request', async () => {
    const context = await run(`${BASE}/teams/t-1/channels/c-1/messages`);
    expect(readPrefer(context)).toBeUndefined();
  });

  it('does not add the header on a chat messages request', async () => {
    const context = await run(`${BASE}/me/chats/c-1/messages`);
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
