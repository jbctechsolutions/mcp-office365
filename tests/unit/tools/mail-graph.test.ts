/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for GraphMailTools batch-read isolation (U5b). getEmailAsync now THROWS
 * on an unresolvable id (NUMERIC_ID_UNSUPPORTED / ID_ENTITY_MISMATCH) instead of
 * returning undefined, so the get_emails batch must isolate per-id failures
 * rather than aborting the whole Promise.all.
 */

import { describe, it, expect, vi } from 'vitest';
import { GraphMailTools } from '../../../src/tools/mail-graph.js';
import { mintSelfEncoded } from '../../../src/ids/token.js';
import { NumericIdUnsupportedError } from '../../../src/utils/errors.js';
import type { GraphRepository } from '../../../src/graph/repository.js';
import type { GraphContentReaders } from '../../../src/graph/content-readers.js';

describe('GraphMailTools.getEmails — per-id isolation', () => {
  it('returns valid emails plus a per-id error when one id is unresolvable', async () => {
    const goodToken = mintSelfEncoded('message', 'msg-good');
    const repository = {
      getEmailAsync: vi.fn(async (id: string | number) => {
        if (typeof id === 'number') throw new NumericIdUnsupportedError(id);
        return { id, dataFilePath: null, subject: 'Hi', folderId: 0 } as never;
      }),
    } as unknown as GraphRepository;
    const contentReaders = {
      email: { readEmailBodyAsync: vi.fn(async () => null) },
    } as unknown as GraphContentReaders;

    const tools = new GraphMailTools(repository, contentReaders);
    const result = await tools.getEmails({ email_ids: [goodToken, 99999], include_body: false } as never);
    const parsed = JSON.parse(result.content[0].text) as { emails: Array<Record<string, unknown>> };

    expect(parsed.emails).toHaveLength(2);
    expect(parsed.emails[0].id).toBe(goodToken);
    // The bad (legacy numeric) id is isolated, not fatal to the batch.
    expect(parsed.emails[1].id).toBe(99999);
    expect(parsed.emails[1].error).toContain('Numeric ID');
  });
});
