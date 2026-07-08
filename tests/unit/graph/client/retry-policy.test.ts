/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect } from 'vitest';
import { RetryHandler } from '@microsoft/microsoft-graph-client';
// Importing graph-client runs its module-load side effect that widens the SDK's
// retry-status set to include 502 (the SDK gates 502 out by default).
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
    for (const action of ['reply', 'forward', 'createReply', 'createForward']) {
      expect(shouldRetryGraphRequest('GET', `${base}/${action}`, 503)).toBe(false);
    }
  });

  it('still retries a GET whose path segment merely starts with an action word (e.g. a file named reply.docx)', () => {
    // The regex anchors the action to end-of-path/query, so a drive item named
    // "reply.docx" read via GET is NOT mistaken for an OData reply action.
    const url = 'https://graph.microsoft.com/v1.0/me/drive/root:/Reports/reply.docx:/content';
    expect(shouldRetryGraphRequest('GET', url, 503)).toBe(true);
  });

  it('actually widens the SDK retry-status set to include 502 (otherwise the policy is dead code)', () => {
    // The SDK's RetryHandler gates on its private static RETRY_STATUS_CODES
    // BEFORE calling shouldRetry, so 502 must be present there or it is never
    // retried at runtime regardless of the predicate.
    const codes = (RetryHandler as unknown as { RETRY_STATUS_CODES: number[] }).RETRY_STATUS_CODES;
    expect(codes).toContain(502);
    expect(codes).toContain(429);
    expect(codes).toContain(503);
    expect(codes).toContain(504);
  });
});
