# Reply/Forward as Draft Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add `reply_as_draft` and `forward_as_draft` tools that create editable draft messages via Graph API, integrating with the existing `update_draft`/`send_draft` flow.

**Architecture:** Three new GraphClient methods (`createReplyDraft`, `createReplyAllDraft`, `createForwardDraft`) call the native Graph API `createReply`/`createReplyAll`/`createForward` endpoints. Two new repository methods cache the returned draft IDs. Two new MailSendTools handler methods orchestrate the flow. Two new tool definitions + Zod schemas + handler cases wire it into the MCP server.

**Tech Stack:** TypeScript, Vitest, Zod, Microsoft Graph API v1.0, `@microsoft/microsoft-graph-types`

---

### Task 1: Add GraphClient draft creation methods

**Files:**
- Modify: `src/graph/client/graph-client.ts` (after `forwardMessage` method ~line 831)
- Test: `tests/unit/graph/client/api-calls.test.ts`

**Step 1: Write failing tests**

Add these tests after the existing `forwardMessage` tests in `tests/unit/graph/client/api-calls.test.ts`. Find the `describe('Draft & send operation endpoints')` or similar section and add a new describe block:

```typescript
describe('Reply/Forward as draft endpoints', () => {
  it('createReplyDraft POSTs to /me/messages/{id}/createReply', async () => {
    setupMock({ id: 'draft-reply-1', subject: 'Re: Hello', isDraft: true });

    await client.createReplyDraft('msg-1');

    expect(apiCalls).toHaveLength(1);
    expect(apiCalls[0].url).toBe('/me/messages/msg-1/createReply');
    expect(apiCalls[0].method).toBe('post');
  });

  it('createReplyAllDraft POSTs to /me/messages/{id}/createReplyAll', async () => {
    setupMock({ id: 'draft-reply-all-1', subject: 'Re: Hello', isDraft: true });

    await client.createReplyAllDraft('msg-1');

    expect(apiCalls).toHaveLength(1);
    expect(apiCalls[0].url).toBe('/me/messages/msg-1/createReplyAll');
    expect(apiCalls[0].method).toBe('post');
  });

  it('createForwardDraft POSTs to /me/messages/{id}/createForward', async () => {
    setupMock({ id: 'draft-fwd-1', subject: 'Fwd: Hello', isDraft: true });

    await client.createForwardDraft('msg-1');

    expect(apiCalls).toHaveLength(1);
    expect(apiCalls[0].url).toBe('/me/messages/msg-1/createForward');
    expect(apiCalls[0].method).toBe('post');
  });

  it('createReplyDraft clears cache', async () => {
    await client.listMessages('inbox', 10, 0);
    apiCalls.length = 0;

    setupMock({ id: 'draft-reply-cache', isDraft: true });
    await client.createReplyDraft('msg-1');
    apiCalls.length = 0;

    setupMock();
    await client.listMessages('inbox', 10, 0);

    const getCalls = apiCalls.filter(c => c.method === 'get');
    expect(getCalls.length).toBeGreaterThan(0);
  });

  it('createReplyAllDraft clears cache', async () => {
    await client.listMessages('inbox', 10, 0);
    apiCalls.length = 0;

    setupMock({ id: 'draft-ra-cache', isDraft: true });
    await client.createReplyAllDraft('msg-1');
    apiCalls.length = 0;

    setupMock();
    await client.listMessages('inbox', 10, 0);

    const getCalls = apiCalls.filter(c => c.method === 'get');
    expect(getCalls.length).toBeGreaterThan(0);
  });

  it('createForwardDraft clears cache', async () => {
    await client.listMessages('inbox', 10, 0);
    apiCalls.length = 0;

    setupMock({ id: 'draft-fwd-cache', isDraft: true });
    await client.createForwardDraft('msg-1');
    apiCalls.length = 0;

    setupMock();
    await client.listMessages('inbox', 10, 0);

    const getCalls = apiCalls.filter(c => c.method === 'get');
    expect(getCalls.length).toBeGreaterThan(0);
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/graph/client/api-calls.test.ts`
Expected: FAIL — `client.createReplyDraft is not a function`

**Step 3: Implement GraphClient methods**

Add after the `forwardMessage` method (~line 831) in `src/graph/client/graph-client.ts`:

```typescript
  async createReplyDraft(messageId: string): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/messages/${messageId}/createReply`)
      .post(null) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  async createReplyAllDraft(messageId: string): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/messages/${messageId}/createReplyAll`)
      .post(null) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  async createForwardDraft(messageId: string): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/messages/${messageId}/createForward`)
      .post(null) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }
```

**Step 4: Run tests to verify they pass**

Run: `npx vitest run tests/unit/graph/client/api-calls.test.ts`
Expected: ALL PASS

**Step 5: Commit**

```bash
git add src/graph/client/graph-client.ts tests/unit/graph/client/api-calls.test.ts
git commit -m "feat: Add GraphClient createReplyDraft, createReplyAllDraft, createForwardDraft methods"
```

---

### Task 2: Add GraphRepository draft creation methods

**Files:**
- Modify: `src/graph/repository.ts` (after `forwardMessageAsync` ~line 864)
- Test: `tests/unit/graph/repository.test.ts`

**Step 1: Write failing tests**

Add to the mock client in `tests/unit/graph/repository.test.ts` (around line 52, after `forwardMessage`):

```typescript
      createReplyDraft: vi.fn(),
      createReplyAllDraft: vi.fn(),
      createForwardDraft: vi.fn(),
```

Then add a new `describe` block after the existing Draft & Send Operations tests (find the `describe('Draft & Send Operations (Async)')` section):

```typescript
  describe('Reply/Forward as Draft (Async)', () => {
    describe('replyAsDraftAsync', () => {
      it('creates a reply draft and caches the result', async () => {
        // Populate message cache
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-orig', subject: 'Hello' },
        ]);
        await repository.searchEmailsAsync('Hello', 50);

        mockClient.createReplyDraft.mockResolvedValue({
          id: 'draft-reply-1',
          subject: 'Re: Hello',
          toRecipients: [{ emailAddress: { address: 'sender@example.com' } }],
        });

        const result = await repository.replyAsDraftAsync(hashStringToNumber('msg-orig'));

        expect(mockClient.createReplyDraft).toHaveBeenCalledWith('msg-orig');
        expect(result.numericId).toBe(hashStringToNumber('draft-reply-1'));
        expect(result.graphId).toBe('draft-reply-1');

        // Verify cached
        expect(repository.getGraphId('message', result.numericId)).toBe('draft-reply-1');
      });

      it('creates a reply-all draft when replyAll is true', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-orig2', subject: 'Team' },
        ]);
        await repository.searchEmailsAsync('Team', 50);

        mockClient.createReplyAllDraft.mockResolvedValue({
          id: 'draft-ra-1',
          subject: 'Re: Team',
          toRecipients: [{ emailAddress: { address: 'all@example.com' } }],
        });

        const result = await repository.replyAsDraftAsync(
          hashStringToNumber('msg-orig2'),
          true
        );

        expect(mockClient.createReplyAllDraft).toHaveBeenCalledWith('msg-orig2');
        expect(result.graphId).toBe('draft-ra-1');
      });

      it('updates draft body when comment is provided', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-comment', subject: 'FYI' },
        ]);
        await repository.searchEmailsAsync('FYI', 50);

        mockClient.createReplyDraft.mockResolvedValue({
          id: 'draft-comment-1',
          subject: 'Re: FYI',
          toRecipients: [],
        });
        mockClient.updateDraft.mockResolvedValue(undefined);

        await repository.replyAsDraftAsync(
          hashStringToNumber('msg-comment'),
          false,
          'Thanks for sharing!'
        );

        expect(mockClient.updateDraft).toHaveBeenCalledWith('draft-comment-1', {
          body: { contentType: 'text', content: 'Thanks for sharing!' },
        });
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.replyAsDraftAsync(99999)
        ).rejects.toThrow('Message ID 99999 not found in cache');
      });
    });

    describe('forwardAsDraftAsync', () => {
      it('creates a forward draft and caches the result', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-fwd', subject: 'Report' },
        ]);
        await repository.searchEmailsAsync('Report', 50);

        mockClient.createForwardDraft.mockResolvedValue({
          id: 'draft-fwd-1',
          subject: 'Fwd: Report',
          toRecipients: [],
        });

        const result = await repository.forwardAsDraftAsync(hashStringToNumber('msg-fwd'));

        expect(mockClient.createForwardDraft).toHaveBeenCalledWith('msg-fwd');
        expect(result.numericId).toBe(hashStringToNumber('draft-fwd-1'));
        expect(repository.getGraphId('message', result.numericId)).toBe('draft-fwd-1');
      });

      it('updates draft with recipients and comment when provided', async () => {
        mockClient.searchMessages.mockResolvedValue([
          { id: 'msg-fwd2', subject: 'Info' },
        ]);
        await repository.searchEmailsAsync('Info', 50);

        mockClient.createForwardDraft.mockResolvedValue({
          id: 'draft-fwd-2',
          subject: 'Fwd: Info',
          toRecipients: [],
        });
        mockClient.updateDraft.mockResolvedValue(undefined);

        await repository.forwardAsDraftAsync(
          hashStringToNumber('msg-fwd2'),
          ['alice@example.com', 'bob@example.com'],
          'Please review'
        );

        expect(mockClient.updateDraft).toHaveBeenCalledWith('draft-fwd-2', {
          toRecipients: [
            { emailAddress: { address: 'alice@example.com' } },
            { emailAddress: { address: 'bob@example.com' } },
          ],
          body: { contentType: 'text', content: 'Please review' },
        });
      });

      it('throws if message not in cache', async () => {
        await expect(
          repository.forwardAsDraftAsync(99999)
        ).rejects.toThrow('Message ID 99999 not found in cache');
      });
    });
  });
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/graph/repository.test.ts`
Expected: FAIL — `repository.replyAsDraftAsync is not a function`

**Step 3: Implement repository methods**

Add after `forwardMessageAsync` (~line 864) in `src/graph/repository.ts`:

```typescript
  async replyAsDraftAsync(
    messageId: number,
    replyAll = false,
    comment?: string,
  ): Promise<{ numericId: number; graphId: string }> {
    const graphMessageId = this.idCache.messages.get(messageId);
    if (graphMessageId == null) throw new Error(`Message ID ${messageId} not found in cache`);

    const draft = replyAll
      ? await this.client.createReplyAllDraft(graphMessageId)
      : await this.client.createReplyDraft(graphMessageId);

    const graphId = draft.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.messages.set(numericId, graphId);

    if (comment != null) {
      await this.client.updateDraft(graphId, {
        body: { contentType: 'text', content: comment },
      });
    }

    return { numericId, graphId };
  }

  async forwardAsDraftAsync(
    messageId: number,
    toRecipients?: string[],
    comment?: string,
  ): Promise<{ numericId: number; graphId: string }> {
    const graphMessageId = this.idCache.messages.get(messageId);
    if (graphMessageId == null) throw new Error(`Message ID ${messageId} not found in cache`);

    const draft = await this.client.createForwardDraft(graphMessageId);

    const graphId = draft.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.messages.set(numericId, graphId);

    const updates: Record<string, unknown> = {};
    if (toRecipients != null && toRecipients.length > 0) {
      updates.toRecipients = toRecipients.map(addr => ({
        emailAddress: { address: addr },
      }));
    }
    if (comment != null) {
      updates.body = { contentType: 'text', content: comment };
    }
    if (Object.keys(updates).length > 0) {
      await this.client.updateDraft(graphId, updates);
    }

    return { numericId, graphId };
  }
```

**Step 4: Run tests to verify they pass**

Run: `npx vitest run tests/unit/graph/repository.test.ts`
Expected: ALL PASS

**Step 5: Commit**

```bash
git add src/graph/repository.ts tests/unit/graph/repository.test.ts
git commit -m "feat: Add GraphRepository replyAsDraftAsync and forwardAsDraftAsync methods"
```

---

### Task 3: Add MailSendTools handler methods and Zod schemas

**Files:**
- Modify: `src/tools/mail-send.ts` (schemas after line ~161, methods in class)
- Test: `tests/unit/tools/mail-send.test.ts`

**Step 1: Write failing tests**

Add to the mock repository in `tests/unit/tools/mail-send.test.ts` (find the mock repo object and add):

```typescript
      replyAsDraftAsync: vi.fn(),
      forwardAsDraftAsync: vi.fn(),
```

Then add a new describe block:

```typescript
  describe('replyAsDraft', () => {
    it('creates a reply draft and returns draft info', async () => {
      (repo.replyAsDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({
        numericId: 101,
        graphId: 'draft-reply-graph-101',
      });

      const result = await tools.replyAsDraft({
        message_id: 42,
        reply_all: false,
      });

      expect(repo.replyAsDraftAsync).toHaveBeenCalledWith(42, false, undefined);
      expect(result).toEqual({
        success: true,
        draft_id: 101,
        message: 'Reply draft created. Use update_draft to edit, then send_draft or prepare_send_email to send.',
      });
    });

    it('passes comment and reply_all to repository', async () => {
      (repo.replyAsDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({
        numericId: 102,
        graphId: 'draft-ra-graph-102',
      });

      await tools.replyAsDraft({
        message_id: 42,
        comment: 'Thanks!',
        reply_all: true,
      });

      expect(repo.replyAsDraftAsync).toHaveBeenCalledWith(42, true, 'Thanks!');
    });
  });

  describe('forwardAsDraft', () => {
    it('creates a forward draft and returns draft info', async () => {
      (repo.forwardAsDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({
        numericId: 201,
        graphId: 'draft-fwd-graph-201',
      });

      const result = await tools.forwardAsDraft({
        message_id: 42,
      });

      expect(repo.forwardAsDraftAsync).toHaveBeenCalledWith(42, undefined, undefined);
      expect(result).toEqual({
        success: true,
        draft_id: 201,
        message: 'Forward draft created. Use update_draft to edit, then send_draft or prepare_send_email to send.',
      });
    });

    it('passes recipients and comment to repository', async () => {
      (repo.forwardAsDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({
        numericId: 202,
        graphId: 'draft-fwd-graph-202',
      });

      await tools.forwardAsDraft({
        message_id: 42,
        to_recipients: ['alice@example.com'],
        comment: 'FYI',
      });

      expect(repo.forwardAsDraftAsync).toHaveBeenCalledWith(
        42,
        ['alice@example.com'],
        'FYI'
      );
    });
  });
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/tools/mail-send.test.ts`
Expected: FAIL — `tools.replyAsDraft is not a function`

**Step 3: Add Zod schemas and interface methods**

In `src/tools/mail-send.ts`, add schemas after `ConfirmForwardEmailInput` (~line 165):

```typescript
// =============================================================================
// Input Schemas -- Draft Reply/Forward (Non-Destructive)
// =============================================================================

export const ReplyAsDraftInput = z.strictObject({
  message_id: z.number().int().positive().describe('The message ID to reply to'),
  comment: z.string().optional().describe('Initial reply body text'),
  reply_all: z.boolean().default(false).describe('Reply to all recipients (default: false)'),
});

export const ForwardAsDraftInput = z.strictObject({
  message_id: z.number().int().positive().describe('The message ID to forward'),
  to_recipients: z.array(z.string().email()).optional().describe('Forward recipients'),
  comment: z.string().optional().describe('Initial forward body text'),
});
```

Add to `IMailSendRepository` interface (after `forwardMessageAsync` ~line 73):

```typescript
  replyAsDraftAsync(messageId: number, replyAll?: boolean, comment?: string): Promise<CreateDraftResult>;
  forwardAsDraftAsync(messageId: number, toRecipients?: string[], comment?: string): Promise<CreateDraftResult>;
```

**Step 4: Implement handler methods**

Add to the `MailSendTools` class (after the `confirmForwardEmail` method):

```typescript
  async replyAsDraft(params: z.infer<typeof ReplyAsDraftInput>): Promise<{
    success: boolean;
    draft_id: number;
    message: string;
  }> {
    const { numericId } = await this.repository.replyAsDraftAsync(
      params.message_id,
      params.reply_all,
      params.comment,
    );
    return {
      success: true,
      draft_id: numericId,
      message: 'Reply draft created. Use update_draft to edit, then send_draft or prepare_send_email to send.',
    };
  }

  async forwardAsDraft(params: z.infer<typeof ForwardAsDraftInput>): Promise<{
    success: boolean;
    draft_id: number;
    message: string;
  }> {
    const { numericId } = await this.repository.forwardAsDraftAsync(
      params.message_id,
      params.to_recipients,
      params.comment,
    );
    return {
      success: true,
      draft_id: numericId,
      message: 'Forward draft created. Use update_draft to edit, then send_draft or prepare_send_email to send.',
    };
  }
```

**Step 5: Run tests to verify they pass**

Run: `npx vitest run tests/unit/tools/mail-send.test.ts`
Expected: ALL PASS

**Step 6: Commit**

```bash
git add src/tools/mail-send.ts tests/unit/tools/mail-send.test.ts
git commit -m "feat: Add MailSendTools replyAsDraft and forwardAsDraft handlers"
```

---

### Task 4: Wire tools into index.ts and update E2E

**Files:**
- Modify: `src/index.ts` (tool definitions, handler cases)
- Modify: `tests/e2e/mcp-client.test.ts` (tool count 72 → 74)

**Step 1: Add tool definitions**

Find the TOOLS array in `src/index.ts`. Add after the `confirm_forward_email` tool definition (around line 1534):

```typescript
    {
      name: 'reply_as_draft',
      description: 'Create a reply (or reply-all) as an editable draft. Returns a draft_id for use with update_draft and send_draft.',
      inputSchema: {
        type: 'object' as const,
        properties: {
          message_id: { type: 'number', description: 'The message ID to reply to' },
          comment: { type: 'string', description: 'Initial reply body text' },
          reply_all: { type: 'boolean', description: 'Reply to all recipients (default: false)' },
        },
        required: ['message_id'],
      },
    },
    {
      name: 'forward_as_draft',
      description: 'Create a forward as an editable draft. Returns a draft_id for use with update_draft and send_draft.',
      inputSchema: {
        type: 'object' as const,
        properties: {
          message_id: { type: 'number', description: 'The message ID to forward' },
          to_recipients: {
            type: 'array',
            items: { type: 'string' },
            description: 'Forward recipients (can also add later via update_draft)',
          },
          comment: { type: 'string', description: 'Initial forward body text' },
        },
        required: ['message_id'],
      },
    },
```

**Step 2: Add handler cases**

Find the handler cases for `prepare_forward_email` / `confirm_forward_email` in the `handleGraphToolCall` function and add after them:

```typescript
      case 'reply_as_draft': {
        const params = ReplyAsDraftInput.parse(args);
        return sendTools.replyAsDraft(params);
      }

      case 'forward_as_draft': {
        const params = ForwardAsDraftInput.parse(args);
        return sendTools.forwardAsDraft(params);
      }
```

**Step 3: Add imports**

Make sure `ReplyAsDraftInput` and `ForwardAsDraftInput` are imported from `./tools/mail-send.js` in `src/index.ts`. Find the existing import line that imports schemas like `CreateDraftInput`, `PrepareReplyEmailInput`, etc. and add the two new ones.

**Step 4: Update E2E tool count**

In `tests/e2e/mcp-client.test.ts`, change line 44:

```typescript
// From:
expect(result.tools.length).toBe(72);
// To:
expect(result.tools.length).toBe(74);
```

**Step 5: Run full test suite**

Run: `npx vitest run`
Expected: ALL PASS (with updated tool count)

**Step 6: Run TypeScript and ESLint checks**

Run: `npx tsc --noEmit && npx eslint src --ext .ts`
Expected: 0 errors

**Step 7: Commit**

```bash
git add src/index.ts tests/e2e/mcp-client.test.ts
git commit -m "feat: Wire reply_as_draft and forward_as_draft tools into MCP server"
```
