/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * U5 isolation guard: an authenticated remote request (remoteMode) with no OBO
 * credential must FAIL CLOSED on a tool call — never fall back to the device-code
 * / process-global identity (which would bind every user to one cached account).
 */

import { describe, expect, it } from 'vitest';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createServer } from '../../../src/index.js';

async function callFolders(remoteMode: boolean): Promise<string> {
  const server = createServer(remoteMode ? { remoteMode: true } : {});
  const client = new Client({ name: 't', version: '1' }, { capabilities: {} });
  const [c, s] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(c), server.connect(s)]);
  const result = await client.callTool({ name: 'list_folders', arguments: {} });
  await client.close();
  const content = (result.content ?? []) as Array<{ type: string; text?: string }>;
  return content.map((b) => b.text ?? '').join('\n');
}

describe('remoteMode fail-closed (U5)', () => {
  it('rejects a Graph tool call with a clear error when OBO is unprovisioned', async () => {
    const text = await callFolders(true);
    // A device-code prompt/identity must never be reached; the error names the
    // missing OBO credential and points at the runbook.
    expect(text).toMatch(/On-Behalf-Of credential not configured/i);
    expect(text).not.toMatch(/device.?code/i);
  });
});
