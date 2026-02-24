# Graph API Write Operations — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add 34 new Graph API tools across 5 phases: mail drafts/sending (12), attachments (2+integration), calendar writes (7), contact writes (5), and task writes (8).

**Architecture:** Each phase follows the same three-layer pattern: GraphClient methods → GraphRepository async wrappers → index.ts tool definitions + handler cases. Destructive operations use a two-phase prepare/confirm approval pattern via ApprovalTokenManager. A new MailSendTools class handles send approval logic.

**Tech Stack:** TypeScript, Microsoft Graph API v1.0 (`@microsoft/microsoft-graph-client`), Zod schemas, Vitest for testing.

---

## Phase 1: Mail Drafts & Sending

### Task 1: Extend approval types for send operations

**Files:**
- Modify: `src/approval/types.ts`
- Modify: `src/approval/hash.ts`
- Modify: `src/approval/index.ts`
- Test: `tests/unit/approval/token-manager.test.ts`

**Step 1: Write failing test for new operation types**

Add to `tests/unit/approval/token-manager.test.ts`:

```typescript
describe('send operation types', () => {
  it('should accept send_draft operation', () => {
    const token = manager.generateToken({
      operation: 'send_draft',
      targetType: 'email',
      targetId: 1,
      targetHash: 'abc123',
    });
    expect(token.operation).toBe('send_draft');
  });

  it('should accept send_email operation', () => {
    const token = manager.generateToken({
      operation: 'send_email',
      targetType: 'email',
      targetId: 0, // no target for direct send
      targetHash: 'abc123',
    });
    expect(token.operation).toBe('send_email');
  });

  it('should accept reply_email operation', () => {
    const token = manager.generateToken({
      operation: 'reply_email',
      targetType: 'email',
      targetId: 1,
      targetHash: 'abc123',
    });
    expect(token.operation).toBe('reply_email');
  });

  it('should accept forward_email operation', () => {
    const token = manager.generateToken({
      operation: 'forward_email',
      targetType: 'email',
      targetId: 1,
      targetHash: 'abc123',
    });
    expect(token.operation).toBe('forward_email');
  });
});
```

**Step 2: Run test to verify it fails**

Run: `npx vitest run tests/unit/approval/token-manager.test.ts --reporter=verbose`
Expected: TypeScript compilation error — `'send_draft'` not assignable to `OperationType`

**Step 3: Add send operation types**

In `src/approval/types.ts`, update the `OperationType` union:

```typescript
export type OperationType =
  | 'delete_email'
  | 'move_email'
  | 'archive_email'
  | 'junk_email'
  | 'delete_folder'
  | 'empty_folder'
  | 'batch_delete_emails'
  | 'batch_move_emails'
  | 'send_draft'
  | 'send_email'
  | 'reply_email'
  | 'forward_email'
  | 'delete_event'
  | 'delete_contact'
  | 'delete_task';
```

Also add to `TargetType`:

```typescript
export type TargetType = 'email' | 'folder' | 'event' | 'contact' | 'task';
```

**Step 4: Add hash functions for send operations**

In `src/approval/hash.ts`, add:

```typescript
export function hashDraftForSend(draft: {
  id: number;
  subject: string | null;
  recipientCount: number;
}): string {
  return createHash('sha256')
    .update(`${draft.id}:${draft.subject ?? ''}:${draft.recipientCount}`)
    .digest('hex')
    .slice(0, 16);
}

export function hashDirectSendForApproval(params: {
  subject: string;
  toCount: number;
  ccCount: number;
  bccCount: number;
}): string {
  return createHash('sha256')
    .update(`${params.subject}:${params.toCount}:${params.ccCount}:${params.bccCount}`)
    .digest('hex')
    .slice(0, 16);
}

export function hashReplyForApproval(params: {
  originalId: number;
  commentLength: number;
  replyAll: boolean;
}): string {
  return createHash('sha256')
    .update(`${params.originalId}:${params.commentLength}:${params.replyAll}`)
    .digest('hex')
    .slice(0, 16);
}

export function hashForwardForApproval(params: {
  originalId: number;
  recipientCount: number;
}): string {
  return createHash('sha256')
    .update(`${params.originalId}:${params.recipientCount}`)
    .digest('hex')
    .slice(0, 16);
}

export function hashEventForApproval(event: {
  id: number;
  subject: string | null;
  startDateTime: string | null;
}): string {
  return createHash('sha256')
    .update(`${event.id}:${event.subject ?? ''}:${event.startDateTime ?? ''}`)
    .digest('hex')
    .slice(0, 16);
}

export function hashContactForApproval(contact: {
  id: number;
  displayName: string | null;
  emailAddress: string | null;
}): string {
  return createHash('sha256')
    .update(`${contact.id}:${contact.displayName ?? ''}:${contact.emailAddress ?? ''}`)
    .digest('hex')
    .slice(0, 16);
}

export function hashTaskForApproval(task: {
  taskId: string;
  title: string | null;
  listId: string;
}): string {
  return createHash('sha256')
    .update(`${task.taskId}:${task.title ?? ''}:${task.listId}`)
    .digest('hex')
    .slice(0, 16);
}
```

Update `src/approval/index.ts` to re-export the new hash functions:

```typescript
export {
  hashEmailForApproval,
  hashFolderForApproval,
  hashDraftForSend,
  hashDirectSendForApproval,
  hashReplyForApproval,
  hashForwardForApproval,
  hashEventForApproval,
  hashContactForApproval,
  hashTaskForApproval,
} from './hash.js';
```

**Step 5: Run tests to verify they pass**

Run: `npx vitest run tests/unit/approval/token-manager.test.ts --reporter=verbose`
Expected: PASS

**Step 6: Commit**

```bash
git add src/approval/ tests/unit/approval/
git commit -m "feat: extend approval types for send, event, contact, and task operations"
```

---

### Task 2: Add GraphClient draft and send methods

**Files:**
- Modify: `src/graph/client/graph-client.ts`
- Test: `tests/unit/graph/client/api-calls.test.ts`

**Step 1: Write failing tests for new GraphClient methods**

Add to `tests/unit/graph/client/api-calls.test.ts` (follow existing test pattern with `createTrackingBuilder` and `mockApi`):

```typescript
describe('Draft operations', () => {
  it('createDraft sends POST /me/messages with isDraft:true', async () => {
    const { builder } = createTrackingBuilder({ id: 'draft-id-1', isDraft: true });
    mockApi.mockReturnValue(builder);

    await client.createDraft({
      subject: 'Test',
      body: 'Hello',
      bodyType: 'text',
      toRecipients: [{ emailAddress: { address: 'a@b.com' } }],
    });

    expect(apiCalls[0].url).toBe('/me/messages');
    expect(apiCalls[0].method).toBe('post');
    expect(apiCalls[0].body).toMatchObject({ isDraft: true, subject: 'Test' });
  });

  it('updateDraft sends PATCH /me/messages/{id}', async () => {
    const { builder } = createTrackingBuilder({});
    mockApi.mockReturnValue(builder);

    await client.updateDraft('draft-id-1', { subject: 'Updated' });

    expect(apiCalls[0].url).toBe('/me/messages/draft-id-1');
    expect(apiCalls[0].method).toBe('patch');
    expect(apiCalls[0].body).toMatchObject({ subject: 'Updated' });
  });

  it('sendDraft sends POST /me/messages/{id}/send', async () => {
    const { builder } = createTrackingBuilder(undefined);
    mockApi.mockReturnValue(builder);

    await client.sendDraft('draft-id-1');

    expect(apiCalls[0].url).toBe('/me/messages/draft-id-1/send');
    expect(apiCalls[0].method).toBe('post');
  });
});

describe('Send operations', () => {
  it('sendMail sends POST /me/sendMail', async () => {
    const { builder } = createTrackingBuilder(undefined);
    mockApi.mockReturnValue(builder);

    await client.sendMail({
      subject: 'Test',
      body: { contentType: 'text', content: 'Hello' },
      toRecipients: [{ emailAddress: { address: 'a@b.com' } }],
    });

    expect(apiCalls[0].url).toBe('/me/sendMail');
    expect(apiCalls[0].method).toBe('post');
    expect(apiCalls[0].body).toHaveProperty('message');
  });

  it('replyMessage sends POST /me/messages/{id}/reply', async () => {
    const { builder } = createTrackingBuilder(undefined);
    mockApi.mockReturnValue(builder);

    await client.replyMessage('msg-1', 'Thanks!', false);

    expect(apiCalls[0].url).toBe('/me/messages/msg-1/reply');
    expect(apiCalls[0].method).toBe('post');
    expect(apiCalls[0].body).toHaveProperty('comment', 'Thanks!');
  });

  it('replyAllMessage sends POST /me/messages/{id}/replyAll', async () => {
    const { builder } = createTrackingBuilder(undefined);
    mockApi.mockReturnValue(builder);

    await client.replyMessage('msg-1', 'Thanks!', true);

    expect(apiCalls[0].url).toBe('/me/messages/msg-1/replyAll');
    expect(apiCalls[0].method).toBe('post');
  });

  it('forwardMessage sends POST /me/messages/{id}/forward', async () => {
    const { builder } = createTrackingBuilder(undefined);
    mockApi.mockReturnValue(builder);

    await client.forwardMessage('msg-1', [{ emailAddress: { address: 'c@d.com' } }], 'FYI');

    expect(apiCalls[0].url).toBe('/me/messages/msg-1/forward');
    expect(apiCalls[0].method).toBe('post');
    expect(apiCalls[0].body).toHaveProperty('toRecipients');
    expect(apiCalls[0].body).toHaveProperty('comment', 'FYI');
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/graph/client/api-calls.test.ts --reporter=verbose`
Expected: FAIL — methods do not exist

**Step 3: Implement GraphClient methods**

Add to `src/graph/client/graph-client.ts` in the Write Operations section:

```typescript
  // ===========================================================================
  // Draft & Send Operations
  // ===========================================================================

  async createDraft(message: {
    subject: string;
    body: string;
    bodyType: 'text' | 'html';
    toRecipients?: MicrosoftGraph.Recipient[];
    ccRecipients?: MicrosoftGraph.Recipient[];
    bccRecipients?: MicrosoftGraph.Recipient[];
  }): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api('/me/messages')
      .post({
        isDraft: true,
        subject: message.subject,
        body: { contentType: message.bodyType, content: message.body },
        toRecipients: message.toRecipients ?? [],
        ccRecipients: message.ccRecipients ?? [],
        bccRecipients: message.bccRecipients ?? [],
      }) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  async updateDraft(
    messageId: string,
    updates: Record<string, unknown>
  ): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/messages/${messageId}`)
      .patch(updates) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  async sendDraft(messageId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}/send`)
      .post(null);
    this.cache.clear();
  }

  async sendMail(message: {
    subject: string;
    body: MicrosoftGraph.ItemBody;
    toRecipients: MicrosoftGraph.Recipient[];
    ccRecipients?: MicrosoftGraph.Recipient[];
    bccRecipients?: MicrosoftGraph.Recipient[];
  }): Promise<void> {
    const client = await this.getClient();
    await client
      .api('/me/sendMail')
      .post({
        message: {
          subject: message.subject,
          body: message.body,
          toRecipients: message.toRecipients,
          ccRecipients: message.ccRecipients ?? [],
          bccRecipients: message.bccRecipients ?? [],
        },
      });
    this.cache.clear();
  }

  async replyMessage(
    messageId: string,
    comment: string,
    replyAll: boolean
  ): Promise<void> {
    const client = await this.getClient();
    const endpoint = replyAll
      ? `/me/messages/${messageId}/replyAll`
      : `/me/messages/${messageId}/reply`;
    await client
      .api(endpoint)
      .post({ comment });
    this.cache.clear();
  }

  async forwardMessage(
    messageId: string,
    toRecipients: MicrosoftGraph.Recipient[],
    comment?: string
  ): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}/forward`)
      .post({
        toRecipients,
        comment: comment ?? '',
      });
    this.cache.clear();
  }
```

**Step 4: Run tests to verify they pass**

Run: `npx vitest run tests/unit/graph/client/api-calls.test.ts --reporter=verbose`
Expected: PASS

**Step 5: Commit**

```bash
git add src/graph/client/graph-client.ts tests/unit/graph/client/api-calls.test.ts
git commit -m "feat: add GraphClient draft and send methods"
```

---

### Task 3: Add GraphRepository draft and send methods

**Files:**
- Modify: `src/graph/repository.ts`
- Test: `tests/unit/graph/repository.test.ts`

**Step 1: Write failing tests**

Add to `tests/unit/graph/repository.test.ts`:

```typescript
describe('Draft operations', () => {
  it('createDraftAsync creates draft and adds to idCache', async () => {
    mockClient.createDraft.mockResolvedValue({ id: 'graph-draft-id', isDraft: true, subject: 'Test' });

    const result = await repository.createDraftAsync({
      subject: 'Test',
      body: 'Hello',
      bodyType: 'text',
      to: ['a@b.com'],
    });

    expect(result).toBeDefined();
    expect(mockClient.createDraft).toHaveBeenCalled();
    // Verify idCache was populated
    const graphId = repository.getGraphId('message', result);
    expect(graphId).toBe('graph-draft-id');
  });

  it('updateDraftAsync updates draft via graph ID', async () => {
    // Pre-populate cache
    mockClient.updateDraft.mockResolvedValue({});

    await repository.updateDraftAsync(numericDraftId, { subject: 'Updated' });
    expect(mockClient.updateDraft).toHaveBeenCalledWith('graph-draft-id', { subject: 'Updated' });
  });

  it('listDraftsAsync lists messages in drafts folder', async () => {
    mockClient.listMessages.mockResolvedValue([]);

    const result = await repository.listDraftsAsync(50, 0);
    expect(mockClient.listMessages).toHaveBeenCalledWith('drafts', 50, 0);
    expect(result).toEqual([]);
  });
});

describe('Send operations', () => {
  it('sendDraftAsync calls client.sendDraft', async () => {
    mockClient.sendDraft.mockResolvedValue(undefined);

    await repository.sendDraftAsync(numericDraftId);
    expect(mockClient.sendDraft).toHaveBeenCalledWith('graph-draft-id');
  });

  it('sendMailAsync calls client.sendMail', async () => {
    mockClient.sendMail.mockResolvedValue(undefined);

    await repository.sendMailAsync({
      subject: 'Test',
      body: 'Hello',
      bodyType: 'text',
      to: ['a@b.com'],
    });
    expect(mockClient.sendMail).toHaveBeenCalled();
  });

  it('replyMessageAsync calls client.replyMessage', async () => {
    mockClient.replyMessage.mockResolvedValue(undefined);

    await repository.replyMessageAsync(numericMsgId, 'Thanks!', true);
    expect(mockClient.replyMessage).toHaveBeenCalledWith('graph-msg-id', 'Thanks!', true);
  });

  it('forwardMessageAsync calls client.forwardMessage', async () => {
    mockClient.forwardMessage.mockResolvedValue(undefined);

    await repository.forwardMessageAsync(numericMsgId, ['c@d.com'], 'FYI');
    expect(mockClient.forwardMessage).toHaveBeenCalled();
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/graph/repository.test.ts --reporter=verbose`
Expected: FAIL — methods do not exist on GraphRepository

**Step 3: Implement GraphRepository methods**

Add to `src/graph/repository.ts` in the Write Operations section:

```typescript
  // ===========================================================================
  // Draft & Send Operations (Async)
  // ===========================================================================

  async createDraftAsync(params: {
    subject: string;
    body: string;
    bodyType: 'text' | 'html';
    to?: string[];
    cc?: string[];
    bcc?: string[];
  }): Promise<number> {
    const toRecipients = (params.to ?? []).map(addr => ({
      emailAddress: { address: addr },
    }));
    const ccRecipients = (params.cc ?? []).map(addr => ({
      emailAddress: { address: addr },
    }));
    const bccRecipients = (params.bcc ?? []).map(addr => ({
      emailAddress: { address: addr },
    }));

    const draft = await this.client.createDraft({
      subject: params.subject,
      body: params.body,
      bodyType: params.bodyType,
      toRecipients,
      ccRecipients,
      bccRecipients,
    });

    if (draft.id != null) {
      const numericId = hashStringToNumber(draft.id);
      this.idCache.messages.set(numericId, draft.id);
      return numericId;
    }
    throw new Error('Draft created but no ID returned');
  }

  async updateDraftAsync(
    draftId: number,
    updates: Record<string, unknown>
  ): Promise<void> {
    const graphId = this.idCache.messages.get(draftId);
    if (graphId == null) throw new Error(`Draft ID ${draftId} not found in cache`);
    await this.client.updateDraft(graphId, updates);
  }

  async listDraftsAsync(limit: number, offset: number): Promise<EmailRow[]> {
    return this.listEmailsWithGraphId('drafts', limit, offset);
  }

  async sendDraftAsync(draftId: number): Promise<void> {
    const graphId = this.idCache.messages.get(draftId);
    if (graphId == null) throw new Error(`Draft ID ${draftId} not found in cache`);
    await this.client.sendDraft(graphId);
  }

  async sendMailAsync(params: {
    subject: string;
    body: string;
    bodyType: 'text' | 'html';
    to: string[];
    cc?: string[];
    bcc?: string[];
  }): Promise<void> {
    const toRecipients = params.to.map(addr => ({
      emailAddress: { address: addr },
    }));
    const ccRecipients = (params.cc ?? []).map(addr => ({
      emailAddress: { address: addr },
    }));
    const bccRecipients = (params.bcc ?? []).map(addr => ({
      emailAddress: { address: addr },
    }));

    await this.client.sendMail({
      subject: params.subject,
      body: { contentType: params.bodyType, content: params.body },
      toRecipients,
      ccRecipients,
      bccRecipients,
    });
  }

  async replyMessageAsync(
    messageId: number,
    comment: string,
    replyAll: boolean
  ): Promise<void> {
    const graphId = this.idCache.messages.get(messageId);
    if (graphId == null) throw new Error(`Message ID ${messageId} not found in cache`);
    await this.client.replyMessage(graphId, comment, replyAll);
  }

  async forwardMessageAsync(
    messageId: number,
    toRecipients: string[],
    comment?: string
  ): Promise<void> {
    const graphId = this.idCache.messages.get(messageId);
    if (graphId == null) throw new Error(`Message ID ${messageId} not found in cache`);
    const recipients = toRecipients.map(addr => ({
      emailAddress: { address: addr },
    }));
    await this.client.forwardMessage(graphId, recipients, comment);
  }
```

**Step 4: Run tests to verify they pass**

Run: `npx vitest run tests/unit/graph/repository.test.ts --reporter=verbose`
Expected: PASS

**Step 5: Commit**

```bash
git add src/graph/repository.ts tests/unit/graph/repository.test.ts
git commit -m "feat: add GraphRepository draft and send methods"
```

---

### Task 4: Create MailSendTools class

**Files:**
- Create: `src/tools/mail-send.ts`
- Test: `tests/unit/tools/mail-send.test.ts`

**Step 1: Write failing tests for MailSendTools**

Create `tests/unit/tools/mail-send.test.ts`:

```typescript
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { MailSendTools } from '../../src/tools/mail-send.js';
import { ApprovalTokenManager } from '../../src/approval/index.js';

describe('MailSendTools', () => {
  let sendTools: MailSendTools;
  let tokenManager: ApprovalTokenManager;
  let mockRepository: any;

  beforeEach(() => {
    tokenManager = new ApprovalTokenManager();
    mockRepository = {
      getEmailAsync: vi.fn(),
      createDraftAsync: vi.fn(),
      updateDraftAsync: vi.fn(),
      listDraftsAsync: vi.fn(),
      sendDraftAsync: vi.fn(),
      sendMailAsync: vi.fn(),
      replyMessageAsync: vi.fn(),
      forwardMessageAsync: vi.fn(),
    };
    sendTools = new MailSendTools(mockRepository, tokenManager);
  });

  describe('create_draft', () => {
    it('creates a draft and returns numeric ID', async () => {
      mockRepository.createDraftAsync.mockResolvedValue(12345);

      const result = await sendTools.createDraft({
        to: ['a@b.com'],
        subject: 'Test',
        body: 'Hello',
        body_type: 'text',
      });

      expect(result.success).toBe(true);
      expect(result.draft_id).toBe(12345);
    });
  });

  describe('prepare_send_draft', () => {
    it('generates approval token with draft preview', async () => {
      mockRepository.getEmailAsync.mockResolvedValue({
        id: 1, subject: 'Test', sender: 'me',
        senderAddress: 'me@test.com', folderId: 10,
        timeReceived: null,
      });

      const result = await sendTools.prepareSendDraft({ draft_id: 1 });

      expect(result.token_id).toBeDefined();
      expect(result.expires_at).toBeDefined();
      expect(result.action).toContain('send');
    });
  });

  describe('confirm_send_draft', () => {
    it('sends the draft after token validation', async () => {
      mockRepository.getEmailAsync.mockResolvedValue({
        id: 1, subject: 'Test', sender: 'me',
        senderAddress: 'me@test.com', folderId: 10,
        timeReceived: null,
      });
      mockRepository.sendDraftAsync.mockResolvedValue(undefined);

      const prepared = await sendTools.prepareSendDraft({ draft_id: 1 });

      const result = await sendTools.confirmSendDraft({
        token_id: prepared.token_id,
        draft_id: 1,
      });

      expect(result.success).toBe(true);
      expect(mockRepository.sendDraftAsync).toHaveBeenCalledWith(1);
    });
  });

  describe('prepare_send_email', () => {
    it('generates approval token with send preview', async () => {
      const result = await sendTools.prepareSendEmail({
        to: ['a@b.com'],
        subject: 'Test',
        body: 'Hello',
        body_type: 'text',
      });

      expect(result.token_id).toBeDefined();
      expect(result.preview.subject).toBe('Test');
      expect(result.preview.to).toEqual(['a@b.com']);
    });
  });

  describe('prepare_reply_email', () => {
    it('generates approval token with reply preview', async () => {
      mockRepository.getEmailAsync.mockResolvedValue({
        id: 1, subject: 'Re: Original', sender: 'them',
        senderAddress: 'them@test.com', folderId: 10,
        timeReceived: 1700000000,
      });

      const result = await sendTools.prepareReplyEmail({
        message_id: 1,
        comment: 'Thanks!',
        reply_all: true,
      });

      expect(result.token_id).toBeDefined();
      expect(result.action).toContain('reply');
    });
  });

  describe('prepare_forward_email', () => {
    it('generates approval token with forward preview', async () => {
      mockRepository.getEmailAsync.mockResolvedValue({
        id: 1, subject: 'FW: Info', sender: 'them',
        senderAddress: 'them@test.com', folderId: 10,
        timeReceived: 1700000000,
      });

      const result = await sendTools.prepareForwardEmail({
        message_id: 1,
        to_recipients: ['c@d.com'],
        comment: 'FYI',
      });

      expect(result.token_id).toBeDefined();
      expect(result.preview.to_recipients).toEqual(['c@d.com']);
    });
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/tools/mail-send.test.ts --reporter=verbose`
Expected: FAIL — `MailSendTools` module does not exist

**Step 3: Implement MailSendTools**

Create `src/tools/mail-send.ts`. Follow the pattern of `mailbox-organization.ts`:

- Zod input schemas for all 12 tools (create_draft, update_draft, list_drafts, prepare/confirm send_draft, prepare/confirm send_email, prepare/confirm reply_email, prepare/confirm forward_email)
- `MailSendTools` class with methods for each tool
- Prepare methods generate tokens via `ApprovalTokenManager`
- Confirm methods validate + consume token, then call repository
- Uses hash functions from `src/approval/hash.ts`

The repository interface expected by MailSendTools:

```typescript
interface IMailSendRepository {
  getEmailAsync(id: number): Promise<EmailRow | undefined>;
  createDraftAsync(params: { subject: string; body: string; bodyType: string; to?: string[]; cc?: string[]; bcc?: string[] }): Promise<number>;
  updateDraftAsync(draftId: number, updates: Record<string, unknown>): Promise<void>;
  listDraftsAsync(limit: number, offset: number): Promise<EmailRow[]>;
  sendDraftAsync(draftId: number): Promise<void>;
  sendMailAsync(params: { subject: string; body: string; bodyType: string; to: string[]; cc?: string[]; bcc?: string[] }): Promise<void>;
  replyMessageAsync(messageId: number, comment: string, replyAll: boolean): Promise<void>;
  forwardMessageAsync(messageId: number, toRecipients: string[], comment?: string): Promise<void>;
}
```

**Step 4: Run tests to verify they pass**

Run: `npx vitest run tests/unit/tools/mail-send.test.ts --reporter=verbose`
Expected: PASS

**Step 5: Commit**

```bash
git add src/tools/mail-send.ts tests/unit/tools/mail-send.test.ts
git commit -m "feat: add MailSendTools class with draft and send approval"
```

---

### Task 5: Wire Phase 1 tools into index.ts

**Files:**
- Modify: `src/tools/index.ts` — re-export MailSendTools schemas
- Modify: `src/index.ts` — add tool definitions to TOOLS array, add handler cases, instantiate MailSendTools

**Step 1: Add tool definitions to TOOLS array**

In `src/index.ts`, add 12 new tool definitions to the `TOOLS` array. Follow existing patterns for `inputSchema`. The new tools are:

- `create_draft` — non-destructive, params: `to` (string[]), `cc` (string[]), `bcc` (string[]), `subject` (string), `body` (string), `body_type` (enum: text/html)
- `update_draft` — non-destructive, params: `draft_id` (number), same optional params as create_draft
- `list_drafts` — read, params: `limit` (number), `offset` (number)
- `prepare_send_draft` / `confirm_send_draft` — two-phase approval
- `prepare_send_email` / `confirm_send_email` — two-phase approval
- `prepare_reply_email` / `confirm_reply_email` — two-phase approval, includes `reply_all` boolean default true
- `prepare_forward_email` / `confirm_forward_email` — two-phase approval

**Step 2: Add handler cases in handleGraphToolCall**

Add cases for each new tool name in the switch statement, delegating to the `sendTools` instance.

**Step 3: Instantiate MailSendTools**

In the initialization section, create `MailSendTools` with `GraphRepository` and `ApprovalTokenManager`. Pass it to `handleGraphToolCall`.

Update `handleGraphToolCall` signature to accept `sendTools` parameter.

**Step 4: Update tools/index.ts exports**

Add re-exports for MailSendTools schemas.

**Step 5: Run full test suite**

Run: `npx vitest run --reporter=verbose`
Expected: All tests PASS

**Step 6: Commit**

```bash
git add src/index.ts src/tools/index.ts
git commit -m "feat: wire Phase 1 draft and send tools into MCP server"
```

---

## Phase 2: Attachment Support

### Task 6: Add GraphClient attachment methods

**Files:**
- Modify: `src/graph/client/graph-client.ts`
- Test: `tests/unit/graph/client/api-calls.test.ts`

**Step 1: Write failing tests**

```typescript
describe('Attachment operations', () => {
  it('listAttachments sends GET /me/messages/{id}/attachments', async () => {
    const { builder } = createTrackingBuilder({ value: [] });
    mockApi.mockReturnValue(builder);

    await client.listAttachments('msg-1');

    expect(apiCalls[0].url).toBe('/me/messages/msg-1/attachments');
    expect(apiCalls[0].method).toBe('get');
  });

  it('getAttachment sends GET /me/messages/{id}/attachments/{attachmentId}', async () => {
    const { builder } = createTrackingBuilder({ contentBytes: 'base64data' });
    mockApi.mockReturnValue(builder);

    await client.getAttachment('msg-1', 'att-1');

    expect(apiCalls[0].url).toBe('/me/messages/msg-1/attachments/att-1');
    expect(apiCalls[0].method).toBe('get');
  });

  it('addAttachment sends POST /me/messages/{id}/attachments', async () => {
    const { builder } = createTrackingBuilder({ id: 'att-new' });
    mockApi.mockReturnValue(builder);

    await client.addAttachment('msg-1', {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'file.txt',
      contentBytes: 'base64data',
    });

    expect(apiCalls[0].url).toBe('/me/messages/msg-1/attachments');
    expect(apiCalls[0].method).toBe('post');
  });

  it('createUploadSession sends POST for large attachments', async () => {
    const { builder } = createTrackingBuilder({ uploadUrl: 'https://upload.example.com' });
    mockApi.mockReturnValue(builder);

    await client.createUploadSession('msg-1', {
      AttachmentItem: {
        attachmentType: 'file',
        name: 'large.zip',
        size: 5000000,
      },
    });

    expect(apiCalls[0].url).toBe('/me/messages/msg-1/attachments/createUploadSession');
    expect(apiCalls[0].method).toBe('post');
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/graph/client/api-calls.test.ts --reporter=verbose`
Expected: FAIL

**Step 3: Implement GraphClient attachment methods**

```typescript
  // ===========================================================================
  // Attachment Operations
  // ===========================================================================

  async listAttachments(messageId: string): Promise<MicrosoftGraph.Attachment[]> {
    const client = await this.getClient();
    const response = await client
      .api(`/me/messages/${messageId}/attachments`)
      .select('id,name,size,contentType,isInline')
      .get() as PageCollection;
    return response.value as MicrosoftGraph.Attachment[];
  }

  async getAttachment(messageId: string, attachmentId: string): Promise<MicrosoftGraph.FileAttachment> {
    const client = await this.getClient();
    return await client
      .api(`/me/messages/${messageId}/attachments/${attachmentId}`)
      .get() as MicrosoftGraph.FileAttachment;
  }

  async addAttachment(
    messageId: string,
    attachment: Record<string, unknown>
  ): Promise<MicrosoftGraph.Attachment> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/messages/${messageId}/attachments`)
      .post(attachment) as MicrosoftGraph.Attachment;
    this.cache.clear();
    return result;
  }

  async createUploadSession(
    messageId: string,
    body: Record<string, unknown>
  ): Promise<{ uploadUrl: string }> {
    const client = await this.getClient();
    return await client
      .api(`/me/messages/${messageId}/attachments/createUploadSession`)
      .post(body) as { uploadUrl: string };
  }
```

**Step 4: Run tests, verify pass, commit**

```bash
git add src/graph/client/graph-client.ts tests/unit/graph/client/api-calls.test.ts
git commit -m "feat: add GraphClient attachment methods"
```

---

### Task 7: Add attachment upload helper and download logic

**Files:**
- Create: `src/graph/attachments.ts`
- Test: `tests/unit/graph/attachments.test.ts`

This file contains:
- `uploadAttachment(client, messageId, filePath, name?, contentType?)` — reads file, routes to inline (<= 3MB) or upload session (> 3MB)
- `downloadAttachment(client, messageId, attachmentId, downloadDir)` — fetches attachment, writes to disk, returns file path
- `sanitizeFilename(name)` — prevents path traversal
- `getDownloadDir()` — reads `MCP_OUTLOOK_DOWNLOAD_DIR` env var, falls back to `os.tmpdir()`

**Step 1: Write tests for upload routing and download**

Test the size routing (mock file reads), filename sanitization, and download directory resolution.

**Step 2: Implement the helper module**

**Step 3: Run tests, verify pass, commit**

```bash
git add src/graph/attachments.ts tests/unit/graph/attachments.test.ts
git commit -m "feat: add attachment upload/download helpers"
```

---

### Task 8: Wire attachment tools into index.ts and update stubs

**Files:**
- Modify: `src/graph/repository.ts` — add `listAttachmentsAsync`, `downloadAttachmentAsync`
- Modify: `src/index.ts` — replace attachment stubs with real implementations, add attachment idCache bucket
- Modify: `src/graph/repository.ts` — add `attachments` to IdCache interface

**Step 1: Add idCache bucket for attachments**

In `src/graph/repository.ts`, add to IdCache:

```typescript
attachments: Map<number, { messageId: string; attachmentId: string }>;
```

Initialize in constructor: `attachments: new Map()`

**Step 2: Implement repository methods**

```typescript
async listAttachmentsAsync(emailId: number): Promise<Array<{
  id: number;
  name: string;
  size: number;
  contentType: string;
  isInline: boolean;
}>> {
  const graphMessageId = this.idCache.messages.get(emailId);
  if (graphMessageId == null) throw new Error(`Message ID ${emailId} not found in cache`);

  const attachments = await this.client.listAttachments(graphMessageId);

  return attachments.map(att => {
    const numericId = hashStringToNumber(att.id ?? '');
    this.idCache.attachments.set(numericId, {
      messageId: graphMessageId,
      attachmentId: att.id ?? '',
    });
    return {
      id: numericId,
      name: att.name ?? 'unnamed',
      size: att.size ?? 0,
      contentType: (att as any).contentType ?? 'application/octet-stream',
      isInline: (att as any).isInline ?? false,
    };
  });
}
```

**Step 3: Replace stubs in handleGraphToolCall**

Replace the `list_attachments` and `download_attachment` stubs with real implementations.

**Step 4: Run full test suite, commit**

```bash
git add src/graph/repository.ts src/index.ts
git commit -m "feat: implement attachment list and download tools"
```

---

### Task 9: Integrate attachments with draft/send tools

**Files:**
- Modify: `src/tools/mail-send.ts` — add optional `attachments` param to create_draft and send_email schemas
- Modify: `src/index.ts` — update create_draft and send_email tool definitions with attachment params

**Step 1: Update Zod schemas**

Add to create_draft and send_email input schemas:

```typescript
attachments: z.array(z.strictObject({
  file_path: z.string().describe('Absolute path to the file to attach'),
  name: z.string().optional().describe('Override filename'),
  content_type: z.string().optional().describe('Override MIME type'),
})).optional().describe('File attachments'),
```

**Step 2: Implement attachment flow in MailSendTools**

For `send_email`: create hidden draft → attach files → send draft.
For `create_draft`: create draft → attach files → return draft ID.

**Step 3: Run tests, commit**

```bash
git add src/tools/mail-send.ts src/index.ts
git commit -m "feat: integrate attachments with draft and send tools"
```

---

## Phase 3: Calendar Write Operations

### Task 10: Add GraphClient calendar write methods

**Files:**
- Modify: `src/graph/client/graph-client.ts`
- Test: `tests/unit/graph/client/api-calls.test.ts`

**Step 1: Write failing tests**

```typescript
describe('Calendar write operations', () => {
  it('createEvent sends POST /me/events', async () => {
    const { builder } = createTrackingBuilder({ id: 'event-1' });
    mockApi.mockReturnValue(builder);

    await client.createEvent({
      subject: 'Meeting',
      start: { dateTime: '2026-03-01T10:00:00', timeZone: 'America/New_York' },
      end: { dateTime: '2026-03-01T11:00:00', timeZone: 'America/New_York' },
    });

    expect(apiCalls[0].url).toBe('/me/events');
    expect(apiCalls[0].method).toBe('post');
  });

  it('createEvent with calendarId sends POST /me/calendars/{id}/events', async () => {
    const { builder } = createTrackingBuilder({ id: 'event-1' });
    mockApi.mockReturnValue(builder);

    await client.createEvent({
      subject: 'Meeting',
      start: { dateTime: '2026-03-01T10:00:00', timeZone: 'UTC' },
      end: { dateTime: '2026-03-01T11:00:00', timeZone: 'UTC' },
    }, 'cal-1');

    expect(apiCalls[0].url).toBe('/me/calendars/cal-1/events');
  });

  it('updateEvent sends PATCH /me/events/{id}', async () => {
    const { builder } = createTrackingBuilder({});
    mockApi.mockReturnValue(builder);

    await client.updateEvent('event-1', { subject: 'Updated' });

    expect(apiCalls[0].url).toBe('/me/events/event-1');
    expect(apiCalls[0].method).toBe('patch');
  });

  it('deleteEvent sends DELETE /me/events/{id}', async () => {
    const { builder } = createTrackingBuilder(undefined);
    mockApi.mockReturnValue(builder);

    await client.deleteEvent('event-1');

    expect(apiCalls[0].url).toBe('/me/events/event-1');
    expect(apiCalls[0].method).toBe('delete');
  });

  it('respondToEvent accept sends POST /me/events/{id}/accept', async () => {
    const { builder } = createTrackingBuilder(undefined);
    mockApi.mockReturnValue(builder);

    await client.respondToEvent('event-1', 'accept', true, 'Will be there');

    expect(apiCalls[0].url).toBe('/me/events/event-1/accept');
    expect(apiCalls[0].method).toBe('post');
    expect(apiCalls[0].body).toMatchObject({ sendResponse: true, comment: 'Will be there' });
  });

  it('respondToEvent decline sends POST /me/events/{id}/decline', async () => {
    const { builder } = createTrackingBuilder(undefined);
    mockApi.mockReturnValue(builder);

    await client.respondToEvent('event-1', 'decline', false);

    expect(apiCalls[0].url).toBe('/me/events/event-1/decline');
  });

  it('respondToEvent tentative sends POST /me/events/{id}/tentativelyAccept', async () => {
    const { builder } = createTrackingBuilder(undefined);
    mockApi.mockReturnValue(builder);

    await client.respondToEvent('event-1', 'tentative', true);

    expect(apiCalls[0].url).toBe('/me/events/event-1/tentativelyAccept');
  });
});
```

**Step 2: Implement GraphClient methods**

```typescript
  // ===========================================================================
  // Calendar Write Operations
  // ===========================================================================

  async createEvent(
    event: Record<string, unknown>,
    calendarId?: string
  ): Promise<MicrosoftGraph.Event> {
    const client = await this.getClient();
    const url = calendarId != null
      ? `/me/calendars/${calendarId}/events`
      : '/me/events';
    const result = await client.api(url).post(event) as MicrosoftGraph.Event;
    this.cache.clear();
    return result;
  }

  async updateEvent(
    eventId: string,
    updates: Record<string, unknown>
  ): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/events/${eventId}`).patch(updates);
    this.cache.clear();
  }

  async deleteEvent(eventId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/events/${eventId}`).delete();
    this.cache.clear();
  }

  async respondToEvent(
    eventId: string,
    response: 'accept' | 'decline' | 'tentative',
    sendResponse: boolean,
    comment?: string
  ): Promise<void> {
    const client = await this.getClient();
    const endpointMap = {
      accept: 'accept',
      decline: 'decline',
      tentative: 'tentativelyAccept',
    };
    await client
      .api(`/me/events/${eventId}/${endpointMap[response]}`)
      .post({ sendResponse, comment: comment ?? '' });
    this.cache.clear();
  }
```

**Step 3: Run tests, verify pass, commit**

```bash
git add src/graph/client/graph-client.ts tests/unit/graph/client/api-calls.test.ts
git commit -m "feat: add GraphClient calendar write methods"
```

---

### Task 11: Add GraphRepository calendar write methods + wire into index.ts

**Files:**
- Modify: `src/graph/repository.ts`
- Modify: `src/index.ts` — replace create_event stub, add new tool definitions and handler cases
- Test: `tests/unit/graph/repository.test.ts`

Repository methods: `createEventAsync`, `updateEventAsync`, `deleteEventAsync`, `respondToEventAsync`

New tools in TOOLS array: `create_event` (update existing), `update_event`, `respond_to_event`, `prepare_delete_event`, `confirm_delete_event`

The `prepare_delete_event` / `confirm_delete_event` follow the same two-phase pattern using `hashEventForApproval`.

**Step 1: Write tests, Step 2: Implement, Step 3: Verify, Step 4: Commit**

```bash
git add src/graph/repository.ts src/index.ts tests/unit/graph/repository.test.ts
git commit -m "feat: wire calendar write tools into MCP server"
```

---

## Phase 4: Contact Write Operations

### Task 12: Add GraphClient contact write methods

**Files:**
- Modify: `src/graph/client/graph-client.ts`
- Test: `tests/unit/graph/client/api-calls.test.ts`

Methods: `createContact`, `updateContact`, `deleteContact`

Endpoints:
- `POST /me/contacts`
- `PATCH /me/contacts/{id}`
- `DELETE /me/contacts/{id}`

Field mapping (tool params → Graph API fields) happens at the repository layer.

**Step 1-4: Write tests, implement, verify, commit**

```bash
git add src/graph/client/graph-client.ts tests/unit/graph/client/api-calls.test.ts
git commit -m "feat: add GraphClient contact write methods"
```

---

### Task 13: Add GraphRepository contact write methods + wire into index.ts

**Files:**
- Modify: `src/graph/repository.ts`
- Modify: `src/index.ts`
- Test: `tests/unit/graph/repository.test.ts`

Repository methods: `createContactAsync` (handles field mapping), `updateContactAsync`, `deleteContactAsync`

New tools: `create_contact`, `update_contact`, `prepare_delete_contact`, `confirm_delete_contact`

Field mapping in `createContactAsync`:
```typescript
const graphContact = {
  givenName: params.given_name,
  surname: params.surname,
  emailAddresses: params.email ? [{ address: params.email }] : [],
  businessPhones: params.phone ? [params.phone] : [],
  mobilePhone: params.mobile_phone,
  companyName: params.company,
  jobTitle: params.job_title,
  businessAddress: {
    street: params.street_address,
    city: params.city,
    state: params.state,
    postalCode: params.postal_code,
    countryOrRegion: params.country,
  },
};
```

**Step 1-4: Write tests, implement, verify, commit**

```bash
git add src/graph/repository.ts src/index.ts tests/unit/graph/repository.test.ts
git commit -m "feat: wire contact write tools into MCP server"
```

---

## Phase 5: Task Write Operations

### Task 14: Add GraphClient task write methods

**Files:**
- Modify: `src/graph/client/graph-client.ts`
- Test: `tests/unit/graph/client/api-calls.test.ts`

Methods: `createTask`, `updateTask`, `deleteTask`, `createTaskList`

Endpoints:
- `POST /me/todo/lists/{listId}/tasks`
- `PATCH /me/todo/lists/{listId}/tasks/{taskId}`
- `DELETE /me/todo/lists/{listId}/tasks/{taskId}`
- `POST /me/todo/lists`

**Step 1-4: Write tests, implement, verify, commit**

```bash
git add src/graph/client/graph-client.ts tests/unit/graph/client/api-calls.test.ts
git commit -m "feat: add GraphClient task write methods"
```

---

### Task 15: Add GraphRepository task write methods + wire into index.ts

**Files:**
- Modify: `src/graph/repository.ts` — add `taskLists` to IdCache, add `createTaskAsync`, `updateTaskAsync`, `completeTaskAsync`, `deleteTaskAsync`, `createTaskListAsync`
- Modify: `src/index.ts` — add tool definitions and handler cases
- Test: `tests/unit/graph/repository.test.ts`

New IdCache bucket:
```typescript
taskLists: Map<number, string>  // numeric ID → Graph task list ID
```

Populate `taskLists` cache in `listTasksAsync` when iterating task lists.

New tools: `create_task`, `update_task`, `complete_task`, `create_task_list`, `prepare_delete_task`, `confirm_delete_task`

`complete_task` is a convenience that calls `updateTaskAsync` with:
```typescript
{ status: 'completed', completedDateTime: { dateTime: new Date().toISOString(), timeZone: 'UTC' } }
```

**Step 1-4: Write tests, implement, verify, commit**

```bash
git add src/graph/repository.ts src/index.ts tests/unit/graph/repository.test.ts
git commit -m "feat: wire task write tools into MCP server"
```

---

### Task 16: Final integration test and cleanup

**Files:**
- Run: full test suite
- Verify: TypeScript compilation with `npx tsc --noEmit`
- Verify: lint with `npx eslint src/`

**Step 1: Run full test suite**

Run: `npx vitest run --reporter=verbose`
Expected: All tests PASS

**Step 2: Verify TypeScript compilation**

Run: `npx tsc --noEmit`
Expected: No errors

**Step 3: Verify lint**

Run: `npx eslint src/ --max-warnings=0`
Expected: No errors or warnings

**Step 4: Final commit if any cleanup needed**

```bash
git commit -m "chore: final cleanup for Graph API write operations"
```

---

## Summary

| Phase | Tasks | New Tools | Key Files |
|-------|-------|-----------|-----------|
| 1: Mail Drafts/Sending | 1-5 | 12 | `mail-send.ts`, approval types, GraphClient, GraphRepository, index.ts |
| 2: Attachments | 6-9 | 2 + integration | `attachments.ts`, GraphClient, GraphRepository, index.ts |
| 3: Calendar Writes | 10-11 | 7 | GraphClient, GraphRepository, index.ts |
| 4: Contact Writes | 12-13 | 5 | GraphClient, GraphRepository, index.ts |
| 5: Task Writes | 14-16 | 8 | GraphClient, GraphRepository, index.ts |
| **Total** | **16** | **34** | |
