import { describe, it, expect } from 'vitest';
import { buildBatchPayload, parseBatchResponse, splitIntoBatches, type BatchRequest, type BatchResponseItem } from '../../../src/graph/client/batch.js';

describe('buildBatchPayload', () => {
  it('builds a valid $batch request body', () => {
    const requests: BatchRequest[] = [
      { id: '1', method: 'GET', url: '/me/messages' },
      { id: '2', method: 'GET', url: '/me/events' },
    ];
    const payload = buildBatchPayload(requests);
    expect(payload.requests).toHaveLength(2);
    expect(payload.requests[0].id).toBe('1');
    expect(payload.requests[0].method).toBe('GET');
    expect(payload.requests[0].url).toBe('/me/messages');
  });

  it('includes headers and body when provided', () => {
    const requests: BatchRequest[] = [
      { id: '1', method: 'POST', url: '/me/messages', headers: { 'Content-Type': 'application/json' }, body: { subject: 'test' } },
    ];
    const payload = buildBatchPayload(requests);
    expect(payload.requests[0].headers).toEqual({ 'Content-Type': 'application/json' });
    expect(payload.requests[0].body).toEqual({ subject: 'test' });
  });

  it('handles empty request array', () => {
    const payload = buildBatchPayload([]);
    expect(payload.requests).toHaveLength(0);
  });
});

describe('splitIntoBatches', () => {
  it('returns single batch for 20 or fewer requests', () => {
    const requests = Array.from({ length: 15 }, (_, i) => ({
      id: String(i), method: 'GET' as const, url: `/me/item/${i}`,
    }));
    const batches = splitIntoBatches(requests);
    expect(batches).toHaveLength(1);
    expect(batches[0]).toHaveLength(15);
  });

  it('splits requests into chunks of 20', () => {
    const requests = Array.from({ length: 25 }, (_, i) => ({
      id: String(i), method: 'GET' as const, url: `/me/item/${i}`,
    }));
    const batches = splitIntoBatches(requests);
    expect(batches).toHaveLength(2);
    expect(batches[0]).toHaveLength(20);
    expect(batches[1]).toHaveLength(5);
  });

  it('handles exactly 20 requests as single batch', () => {
    const requests = Array.from({ length: 20 }, (_, i) => ({
      id: String(i), method: 'GET' as const, url: `/me/item/${i}`,
    }));
    const batches = splitIntoBatches(requests);
    expect(batches).toHaveLength(1);
  });

  it('handles empty array', () => {
    const batches = splitIntoBatches([]);
    expect(batches).toHaveLength(0);
  });

  it('supports custom batch size', () => {
    const requests = Array.from({ length: 10 }, (_, i) => ({
      id: String(i), method: 'GET' as const, url: `/me/item/${i}`,
    }));
    const batches = splitIntoBatches(requests, 3);
    expect(batches).toHaveLength(4);
    expect(batches[0]).toHaveLength(3);
    expect(batches[3]).toHaveLength(1);
  });
});

describe('parseBatchResponse', () => {
  it('maps response IDs to results', () => {
    const response = {
      responses: [
        { id: '1', status: 200, body: { value: [{ name: 'test' }] } },
        { id: '2', status: 200, body: { value: [] } },
      ] as BatchResponseItem[],
    };
    const results = parseBatchResponse(response);
    expect(results.get('1')?.status).toBe(200);
    expect(results.get('2')?.status).toBe(200);
    expect(results.size).toBe(2);
  });

  it('handles partial failures', () => {
    const response = {
      responses: [
        { id: '1', status: 200, body: { value: [] } },
        { id: '2', status: 404, body: { error: { code: 'NotFound', message: 'Resource not found' } } },
      ] as BatchResponseItem[],
    };
    const results = parseBatchResponse(response);
    expect(results.get('1')?.status).toBe(200);
    expect(results.get('2')?.status).toBe(404);
  });

  it('handles empty responses', () => {
    const response = { responses: [] as BatchResponseItem[] };
    const results = parseBatchResponse(response);
    expect(results.size).toBe(0);
  });

  it('preserves response headers', () => {
    const response = {
      responses: [
        { id: '1', status: 200, headers: { 'ETag': '"abc123"' }, body: { value: [] } },
      ] as BatchResponseItem[],
    };
    const results = parseBatchResponse(response);
    expect(results.get('1')?.headers?.ETag).toBe('"abc123"');
  });
});
