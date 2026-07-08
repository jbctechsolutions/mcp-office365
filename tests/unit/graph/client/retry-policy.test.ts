/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect } from 'vitest';
import { shouldRetryGraphRequest } from '../../../../src/graph/client/graph-client.js';

const GET = 'https://graph.microsoft.com/v1.0/me/messages';
const SEND = 'https://graph.microsoft.com/v1.0/me/sendMail';

describe('shouldRetryGraphRequest (D5 policy)', () => {
  it('retries transient statuses on idempotent GET reads', () => {
    for (const status of [429, 502, 503, 504]) {
      expect(shouldRetryGraphRequest('GET', GET, status)).toBe(true);
    }
  });

  it('retries on HEAD and OPTIONS too', () => {
    expect(shouldRetryGraphRequest('HEAD', GET, 503)).toBe(true);
    expect(shouldRetryGraphRequest('OPTIONS', GET, 503)).toBe(true);
  });

  it('is case-insensitive on the method', () => {
    expect(shouldRetryGraphRequest('get', GET, 429)).toBe(true);
  });

  it('does not retry non-transient statuses', () => {
    for (const status of [200, 400, 401, 403, 404, 500, 501]) {
      expect(shouldRetryGraphRequest('GET', GET, status)).toBe(false);
    }
  });

  it('never retries writes (POST/PATCH/PUT/DELETE) even on 429/5xx', () => {
    for (const method of ['POST', 'PATCH', 'PUT', 'DELETE']) {
      expect(shouldRetryGraphRequest(method, GET, 429)).toBe(false);
      expect(shouldRetryGraphRequest(method, GET, 503)).toBe(false);
    }
  });

  it('never retries a sendMail action, even if surfaced as a GET', () => {
    expect(shouldRetryGraphRequest('GET', SEND, 429)).toBe(false);
    expect(shouldRetryGraphRequest('GET', SEND, 503)).toBe(false);
  });

  it('never retries reply/forward/createReply/createForward actions', () => {
    const base = 'https://graph.microsoft.com/v1.0/me/messages/AAA';
    for (const action of ['reply', 'forward', 'createReply', 'createForward', 'send']) {
      expect(shouldRetryGraphRequest('GET', `${base}/${action}`, 503)).toBe(false);
    }
  });
});
