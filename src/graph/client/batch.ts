/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Microsoft Graph $batch request utilities.
 * Supports batching up to 20 requests per batch (Graph API limit).
 */

export interface BatchRequest {
  id: string;
  method: string;
  url: string;
  headers?: Record<string, string>;
  body?: unknown;
}

export interface BatchResponseItem {
  id: string;
  status: number;
  headers?: Record<string, string>;
  body: unknown;
}

/**
 * Builds a $batch request payload from an array of individual requests.
 */
export function buildBatchPayload(requests: BatchRequest[]): { requests: BatchRequest[] } {
  return { requests };
}

/**
 * Splits requests into batches of the given size (default 20, the Graph API limit).
 */
export function splitIntoBatches(requests: BatchRequest[], maxPerBatch: number = 20): BatchRequest[][] {
  if (requests.length === 0) return [];
  const batches: BatchRequest[][] = [];
  for (let i = 0; i < requests.length; i += maxPerBatch) {
    batches.push(requests.slice(i, i + maxPerBatch));
  }
  return batches;
}

/**
 * Parses a $batch response into a Map keyed by request ID.
 */
export function parseBatchResponse(response: { responses: BatchResponseItem[] }): Map<string, BatchResponseItem> {
  const map = new Map<string, BatchResponseItem>();
  for (const item of response.responses) {
    map.set(item.id, item);
  }
  return map;
}
