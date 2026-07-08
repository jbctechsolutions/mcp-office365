/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Proves the D5/U8 middleware is actually installed — not just that the pure
 * `shouldRetryGraphRequest` predicate is correct. The other graph-client test
 * files stub the handlers as inert no-ops, so a regression that dropped a
 * `setNext`, swapped the chain order, or forgot to pass `shouldRetry` into
 * `RetryHandlerOptions` would pass every one of them. This file records the
 * construction + wiring calls and asserts the chain: auth → retry → redirect → http.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';

// `mock`-prefixed so the hoisted vi.mock factory may reference it.
const mockState: {
  setNextCalls: Array<{ from: string; to: string }>;
  capturedShouldRetry: ((d: number, a: number, req: unknown, opts: unknown, res: unknown) => boolean) | null;
  initMiddlewareArg: { tag?: string } | null;
} = { setNextCalls: [], capturedShouldRetry: null, initMiddlewareArg: null };

vi.mock('@microsoft/microsoft-graph-client', () => {
  class TaggedHandler {
    constructor(public readonly tag: string) {}
    setNext(next: unknown): void {
      mockState.setNextCalls.push({ from: this.tag, to: (next as TaggedHandler).tag });
    }
  }
  return {
    Client: {
      initWithMiddleware: vi.fn((opts: { middleware: { tag?: string } }) => {
        mockState.initMiddlewareArg = opts.middleware;
        const builder: Record<string, unknown> = {};
        for (const m of ['select', 'top', 'skip', 'orderby', 'filter', 'search', 'query', 'header', 'responseType']) {
          builder[m] = vi.fn(() => builder);
        }
        builder.get = vi.fn().mockResolvedValue({ value: [] });
        return { api: vi.fn(() => builder) };
      }),
    },
    AuthenticationHandler: class extends TaggedHandler {
      constructor() {
        super('auth');
      }
    },
    RetryHandler: class extends TaggedHandler {
      constructor(public readonly options: unknown) {
        super('retry');
      }
    },
    RetryHandlerOptions: class {
      constructor(
        public readonly delay: number | undefined,
        public readonly maxRetries: number | undefined,
        public readonly shouldRetry: (d: number, a: number, req: unknown, opts: unknown, res: unknown) => boolean
      ) {
        mockState.capturedShouldRetry = shouldRetry;
      }
    },
    RedirectHandler: class extends TaggedHandler {
      constructor() {
        super('redirect');
      }
    },
    HTTPMessageHandler: class extends TaggedHandler {
      constructor() {
        super('http');
      }
    },
    ResponseType: { ARRAYBUFFER: 'arraybuffer' },
  };
});

vi.mock('../../../../src/graph/auth/index.js', () => ({
  getAccessToken: vi.fn().mockResolvedValue('test-access-token'),
}));

vi.mock('isomorphic-fetch', () => ({ default: vi.fn() }));

import { GraphClient } from '../../../../src/graph/client/graph-client.js';

describe('graph transport middleware wiring (U8/D5)', () => {
  beforeEach(async () => {
    mockState.setNextCalls.length = 0;
    mockState.capturedShouldRetry = null;
    mockState.initMiddlewareArg = null;
    // Trigger the private getClient() by making any Graph call.
    const client = new GraphClient();
    await client.listOnlineMeetings(1);
  });

  it('builds the chain in order auth → retry → redirect → http', () => {
    expect(mockState.setNextCalls).toEqual([
      { from: 'auth', to: 'retry' },
      { from: 'retry', to: 'redirect' },
      { from: 'redirect', to: 'http' },
    ]);
  });

  it('passes the head of the chain (auth) to initWithMiddleware', () => {
    expect(mockState.initMiddlewareArg?.tag).toBe('auth');
  });

  it('installs a shouldRetry callback into RetryHandlerOptions', () => {
    expect(mockState.capturedShouldRetry).toBeTypeOf('function');
  });

  it('the installed callback excludes writes and includes idempotent reads', () => {
    const fn = mockState.capturedShouldRetry;
    if (fn === null) throw new Error('shouldRetry was not captured');
    const res503 = { status: 503 };
    // GET 503 → retriable
    expect(fn(0, 0, 'https://graph.microsoft.com/v1.0/me/messages', { method: 'GET' }, res503)).toBe(true);
    // POST 503 → never (double-send guard)
    expect(fn(0, 0, 'https://graph.microsoft.com/v1.0/me/messages', { method: 'POST' }, res503)).toBe(false);
    // sendMail GET → never
    expect(fn(0, 0, 'https://graph.microsoft.com/v1.0/me/sendMail', { method: 'GET' }, res503)).toBe(false);
  });
});
