# Feature Enhancements Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add 13 new features across mail, tasks, contacts, and calendar domains.

**Architecture:** Each feature follows the same 4-layer pattern: Graph client method → Repository method → Tool schema/handler in index.ts → Tests. Features are independent and committed individually.

**Tech Stack:** TypeScript, Microsoft Graph API (`@microsoft/microsoft-graph-client`), Zod schemas, Vitest tests.

**Run tests:** `npx vitest run` (full suite), or `npx vitest run tests/unit/path/to/test.ts` (specific)

**Type check:** `npx tsc --noEmit`

---

## Phase 1: Mail Features

### Task 1: Email Importance (`set_email_importance`)

Simplest feature — follows exact same pattern as `setEmailFlagAsync`.

**Files:**
- Modify: `src/tools/mailbox-organization.ts` — add schema + method
- Modify: `src/graph/repository.ts` — add `setEmailImportanceAsync`
- Modify: `src/index.ts` — add tool definition + case handler
- Modify: `tests/unit/tools/mailbox-organization.test.ts` — add tests
- Modify: `tests/unit/graph/repository.test.ts` — add repo test
- Modify: `README.md` — add to tools list

**Step 1: Add schema and method to mailbox-organization.ts**

Find `SetEmailCategoriesInput` and add after it:

```typescript
export const SetEmailImportanceInput = z.strictObject({
  email_id: z.number().int().positive().describe('The email ID'),
  importance: z.enum(['low', 'normal', 'high']).describe('Email importance level'),
});
```

Find the `IMailboxOrganizationRepository` interface. Add:

```typescript
setEmailImportanceAsync(emailId: number, importance: string): Promise<void>;
```

Find the class method `setEmailCategories`. Add after it:

```typescript
async setEmailImportance(params: z.infer<typeof SetEmailImportanceInput>): Promise<{ success: boolean; message: string }> {
  await this.repository.setEmailImportanceAsync(params.email_id, params.importance);
  return { success: true, message: `Email importance set to ${params.importance}.` };
}
```

**Step 2: Add repository method**

In `src/graph/repository.ts`, find `setEmailCategoriesAsync`. Add after it:

```typescript
async setEmailImportanceAsync(emailId: number, importance: string): Promise<void> {
  const graphId = this.idCache.messages.get(emailId);
  if (graphId == null) throw new Error(`Message ID ${emailId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
  await this.client.updateMessage(graphId, { importance });
}
```

**Step 3: Add tool definition and handler in index.ts**

In the TOOLS array, find the `set_email_categories` tool entry. Add after it:

```typescript
{
  name: 'set_email_importance',
  description: 'Set email importance/priority level (Graph API)',
  inputSchema: {
    type: 'object',
    properties: {
      email_id: { type: 'number', description: 'The email ID' },
      importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Importance level' },
    },
    required: ['email_id', 'importance'],
  },
},
```

In `handleOrgToolCall` or the org tools handler, add the case:

```typescript
case 'set_email_importance': {
  const params = SetEmailImportanceInput.parse(args);
  const result = await orgTools.setEmailImportance(params);
  return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
}
```

**Step 4: Write tests**

In `tests/unit/tools/mailbox-organization.test.ts`, add a describe block:

```typescript
describe('setEmailImportance', () => {
  it('sets importance on an email', async () => {
    (repo.setEmailImportanceAsync as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);
    const result = await tools.setEmailImportance({ email_id: 1, importance: 'high' });
    expect(result).toEqual({ success: true, message: 'Email importance set to high.' });
    expect(repo.setEmailImportanceAsync).toHaveBeenCalledWith(1, 'high');
  });
});
```

In `tests/unit/graph/repository.test.ts`, add:

```typescript
describe('setEmailImportanceAsync', () => {
  it('updates message importance', async () => {
    // Populate cache
    mockClient.searchMessages.mockResolvedValue([{ id: 'msg-imp', subject: 'Test' }]);
    await repository.searchEmailsAsync('Test', 50);
    mockClient.updateMessage.mockResolvedValue(undefined);

    await repository.setEmailImportanceAsync(hashStringToNumber('msg-imp'), 'high');
    expect(mockClient.updateMessage).toHaveBeenCalledWith('msg-imp', { importance: 'high' });
  });

  it('throws when email not in cache', async () => {
    await expect(repository.setEmailImportanceAsync(99999, 'high'))
      .rejects.toThrow('Message ID 99999 not found in cache');
  });
});
```

**Step 5: Update README.md**

Add `- \`set_email_importance\` - Set email importance/priority level (low, normal, high)` to the Mailbox Organization section.

**Step 6: Run tests and type check, commit**

```
npx vitest run && npx tsc --noEmit
git add -A && git commit -m "feat: add set_email_importance tool"
```

---

### Task 2: Draft Attachments (`add_draft_attachment` / `add_draft_inline_image`)

**Files:**
- Modify: `src/tools/mail-send.ts` — add schemas + methods
- Modify: `src/index.ts` — add tool definitions + case handlers
- Modify: `tests/unit/tools/mail-send.test.ts` — add tests
- Modify: `README.md`

**Step 1: Add schemas to mail-send.ts**

After `UpdateDraftInput`, add:

```typescript
export const AddDraftAttachmentInput = z.strictObject({
  draft_id: z.number().int().positive().describe('The draft ID to add attachment to'),
  file_path: z.string().describe('Absolute path to the file to attach'),
  name: z.string().optional().describe('Override filename'),
  content_type: z.string().optional().describe('Override MIME type'),
});

export const AddDraftInlineImageInput = z.strictObject({
  draft_id: z.number().int().positive().describe('The draft ID to add inline image to'),
  file_path: z.string().describe('Absolute path to the image file'),
  content_id: z.string().describe('Content-ID for referencing in HTML (use in <img src="cid:content_id">)'),
});
```

**Step 2: Add methods to MailSendTools class**

The tools need to resolve the draft's Graph ID. Add a `getGraphIdForDraft` method to `IMailSendRepository`:

```typescript
getGraphIdForDraft(draftId: number): string | undefined;
```

Implement in GraphRepository by exposing the cache lookup:

```typescript
getGraphIdForDraft(draftId: number): string | undefined {
  return this.idCache.messages.get(draftId);
}
```

Then add methods to MailSendTools:

```typescript
async addDraftAttachment(params: z.infer<typeof AddDraftAttachmentInput>): Promise<{ success: boolean; message: string }> {
  const graphId = this.repository.getGraphIdForDraft(params.draft_id);
  if (graphId == null) throw new Error(`Draft ID ${params.draft_id} not found in cache. Try listing drafts first.`);
  const graphClient = this.repository.getGraphClient();
  await uploadAttachment(graphClient, graphId, params.file_path, params.name, params.content_type);
  return { success: true, message: 'Attachment added to draft.' };
}

async addDraftInlineImage(params: z.infer<typeof AddDraftInlineImageInput>): Promise<{ success: boolean; message: string }> {
  const graphId = this.repository.getGraphIdForDraft(params.draft_id);
  if (graphId == null) throw new Error(`Draft ID ${params.draft_id} not found in cache. Try listing drafts first.`);
  const graphClient = this.repository.getGraphClient();
  await uploadInlineAttachment(graphClient, graphId, params.file_path, params.content_id);
  return { success: true, message: 'Inline image added to draft.' };
}
```

**Step 3: Add tool definitions and handlers in index.ts**

Add two tool entries after the `update_draft` tool. Add two case handlers in `handleSendToolCall`.

**Step 4: Write tests**

Test both methods: success case (mock resolves), cache miss (throws), verify upload functions called with correct args.

**Step 5: Update README, run tests, commit**

```
git commit -m "feat: add add_draft_attachment and add_draft_inline_image tools"
```

---

### Task 3: Batch Read Emails (`get_emails`)

**Files:**
- Modify: `src/tools/mail.ts` — add schema
- Modify: `src/index.ts` — add tool definition + case handler
- Modify: `tests/unit/graph/repository.test.ts` — (no new repo method needed)
- Modify: `README.md`

**Step 1: Add schema to mail.ts**

```typescript
export const GetEmailsInput = z.strictObject({
  email_ids: z.array(z.number().int().positive()).min(1).max(25)
    .describe('Array of email IDs to fetch (max 25)'),
  include_body: z.boolean().default(false).describe('Include full email body'),
  strip_html: z.boolean().default(false).describe('Strip HTML from body'),
});
```

**Step 2: Add handler in index.ts**

```typescript
case 'get_emails': {
  const params = GetEmailsInput.parse(args);
  const results = await Promise.all(
    params.email_ids.map(async (id) => {
      const email = await repository.getEmailAsync(id);
      if (email == null) return { id, error: 'Not found' };
      let body: string | null = null;
      if (params.include_body) {
        body = await contentReaders.email.readEmailBodyAsync(email.dataFilePath);
        if (params.strip_html && body != null) body = stripHtml(body);
      }
      return { ...transformEmailRow(email), body };
    })
  );
  return { content: [{ type: 'text', text: JSON.stringify({ emails: results }, null, 2) }] };
}
```

**Step 3: Add tool definition, tests, README update, commit**

```
git commit -m "feat: add get_emails batch read tool"
```

---

### Task 4: Conversation/Thread View (`list_conversation`)

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add `listConversationMessages`
- Modify: `src/graph/repository.ts` — add `listConversationAsync`, add conversationId reverse cache
- Modify: `src/tools/mail.ts` — add schema
- Modify: `src/index.ts` — add tool + handler
- Modify: `tests/unit/graph/client/api-calls.test.ts`
- Modify: `tests/unit/graph/repository.test.ts`
- Modify: `README.md`

**Step 1: Add Graph client method**

In `graph-client.ts`, after `searchMessagesInFolder`, add:

```typescript
async listConversationMessages(
  conversationId: string,
  limit: number = 50
): Promise<MicrosoftGraph.Message[]> {
  const client = await this.getClient();
  const response = await client
    .api('/me/messages')
    .filter(`conversationId eq '${conversationId}'`)
    .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId,parentFolderId')
    .orderby('receivedDateTime asc')
    .top(limit)
    .get() as PageCollection;
  return response.value as MicrosoftGraph.Message[];
}
```

**Step 2: Add reverse conversationId cache to repository**

The challenge: we have `hashStringToNumber(conversationId)` → numeric, but need to go numeric → Graph string. Add a new cache map:

```typescript
// In IdCache interface:
conversations: Map<number, string>;  // numeric conversationId → Graph string conversationId
```

In `mapAndCacheMessages` (or wherever messages are cached), also cache the conversationId:

```typescript
if (msg.conversationId != null) {
  const numericConvId = hashStringToNumber(msg.conversationId);
  this.idCache.conversations.set(numericConvId, msg.conversationId);
}
```

**Step 3: Add repository method**

```typescript
async listConversationAsync(messageId: number, limit: number): Promise<EmailRow[]> {
  // First get the message to find its conversationId
  const email = await this.getEmailAsync(messageId);
  if (email == null) throw new Error(`Message ID ${messageId} not found`);
  if (email.conversationId == null) throw new Error(`Message has no conversation ID`);

  // Look up the Graph string conversationId from cache
  const graphConversationId = this.idCache.conversations.get(email.conversationId);
  if (graphConversationId == null) throw new Error(`Conversation ID not found in cache. Try fetching the email first.`);

  const messages = await this.client.listConversationMessages(graphConversationId, limit);
  // Cache all returned message IDs
  for (const msg of messages) {
    if (msg.id != null) {
      this.idCache.messages.set(hashStringToNumber(msg.id), msg.id);
    }
  }
  return messages.map((m) => mapMessageToEmailRow(m));
}
```

**Step 4: Add schema, tool definition, handler, tests, README, commit**

```typescript
export const ListConversationInput = z.strictObject({
  message_id: z.number().int().positive().describe('Any message ID from the conversation'),
  limit: z.number().int().min(1).max(100).default(25).describe('Max messages to return'),
});
```

```
git commit -m "feat: add list_conversation thread view tool"
```

---

### Task 5: KQL Advanced Search (`search_emails_advanced`)

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add `searchMessagesKql`, `searchMessagesKqlInFolder`
- Modify: `src/graph/repository.ts` — add `searchEmailsAdvancedAsync`, `searchEmailsAdvancedInFolderAsync`
- Modify: `src/tools/mail.ts` — add schema
- Modify: `src/index.ts` — add tool + handler
- Tests + README

**Step 1: Add Graph client methods**

The difference from existing `searchMessages`: pass the query directly without wrapping in quotes.

```typescript
async searchMessagesKql(query: string, limit: number = 50): Promise<MicrosoftGraph.Message[]> {
  const client = await this.getClient();
  const response = await client
    .api('/me/messages')
    .search(query)
    .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId,parentFolderId')
    .top(limit)
    .get() as PageCollection;
  return response.value as MicrosoftGraph.Message[];
}

async searchMessagesKqlInFolder(
  folderId: string,
  query: string,
  limit: number = 50
): Promise<MicrosoftGraph.Message[]> {
  const client = await this.getClient();
  const response = await client
    .api(`/me/mailFolders/${folderId}/messages`)
    .search(query)
    .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId,parentFolderId')
    .top(limit)
    .get() as PageCollection;
  return response.value as MicrosoftGraph.Message[];
}
```

**Step 2: Add repository methods**

Follow the same pattern as `searchEmailsAsync` but call the KQL variants.

**Step 3: Add schema**

```typescript
export const SearchEmailsAdvancedInput = z.strictObject({
  query: z.string().min(1).describe(
    'KQL search query. Examples: from:alice, subject:"quarterly report", hasAttachments:true, received>=2024-01-01. Combine with AND/OR.'
  ),
  folder_id: z.number().int().positive().optional().describe('Optional folder to search in'),
  limit: z.number().int().min(1).max(100).default(50).describe('Max results'),
});
```

**Step 4: Tool, handler, tests, README, commit**

```
git commit -m "feat: add search_emails_advanced KQL search tool"
```

---

### Task 6: Delta Sync (`check_new_emails`)

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add `getMessagesDelta`
- Modify: `src/graph/repository.ts` — add delta link storage + `checkNewEmailsAsync`
- Modify: `src/tools/mail.ts` — add schema
- Modify: `src/index.ts` — add tool + handler
- Tests + README

**Step 1: Add Graph client method**

```typescript
async getMessagesDelta(
  folderId: string,
  deltaLink?: string
): Promise<{ messages: MicrosoftGraph.Message[]; deltaLink: string }> {
  const client = await this.getClient();
  let response;

  if (deltaLink != null) {
    // Use the delta link from a previous call
    response = await client.api(deltaLink).get() as PageCollection;
  } else {
    // Initial delta call
    response = await client
      .api(`/me/mailFolders/${folderId}/messages/delta`)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId,parentFolderId')
      .top(50)
      .get() as PageCollection;
  }

  // Collect all messages across pages
  const messages: MicrosoftGraph.Message[] = [...(response.value ?? [])];

  // Follow @odata.nextLink pages
  let nextLink = response['@odata.nextLink'];
  while (nextLink != null) {
    const nextPage = await client.api(nextLink).get() as PageCollection;
    messages.push(...(nextPage.value ?? []));
    nextLink = nextPage['@odata.nextLink'];
  }

  // Get the delta link for next call
  const newDeltaLink = response['@odata.deltaLink'] ?? '';
  return { messages, deltaLink: newDeltaLink };
}
```

**Step 2: Add delta link storage to repository**

Add to the class fields (not IdCache — this is session state):

```typescript
private readonly deltaLinks: Map<number, string> = new Map();
```

Add repository method:

```typescript
async checkNewEmailsAsync(folderId: number): Promise<{ emails: EmailRow[]; isInitialSync: boolean }> {
  const graphFolderId = this.idCache.folders.get(folderId);
  if (graphFolderId == null) throw new Error(`Folder ID ${folderId} not found in cache. Try listing folders first.`);

  const existingDeltaLink = this.deltaLinks.get(folderId);
  const isInitialSync = existingDeltaLink == null;

  const { messages, deltaLink } = await this.client.getMessagesDelta(
    graphFolderId,
    existingDeltaLink
  );

  // Store new delta link
  if (deltaLink) {
    this.deltaLinks.set(folderId, deltaLink);
  }

  // Cache message IDs
  for (const msg of messages) {
    if (msg.id != null) {
      this.idCache.messages.set(hashStringToNumber(msg.id), msg.id);
    }
  }

  // Filter out deleted messages (they have @removed property)
  const activeMessages = messages.filter((m) => !(m as any)['@removed']);
  return {
    emails: activeMessages.map((m) => mapMessageToEmailRow(m)),
    isInitialSync,
  };
}
```

**Step 3: Schema, tool, handler**

```typescript
export const CheckNewEmailsInput = z.strictObject({
  folder_id: z.number().int().positive().describe('Folder ID to check for new emails'),
});
```

Handler returns `{ emails, is_initial_sync }`.

**Step 4: Tests, README, commit**

```
git commit -m "feat: add check_new_emails delta sync tool"
```

---

### Task 7: Mail Rules (`list_mail_rules` / `create_mail_rule` / `prepare_delete_mail_rule` / `confirm_delete_mail_rule`)

This is the most complex mail feature — new CRUD with two-phase delete.

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add 3 methods
- Modify: `src/graph/repository.ts` — add rule ID cache + 3 methods
- Create: `src/tools/mail-rules.ts` — new tool module with schemas, interface, class
- Modify: `src/index.ts` — add 4 tool definitions + case handlers + wire up
- Create: `tests/unit/tools/mail-rules.test.ts`
- Modify: `tests/unit/graph/repository.test.ts`
- Modify: `README.md`

**Step 1: Graph client methods**

```typescript
// After the mail section in graph-client.ts:

async listMailRules(): Promise<MicrosoftGraph.MessageRule[]> {
  const client = await this.getClient();
  const response = await client
    .api('/me/mailFolders/inbox/messageRules')
    .get() as PageCollection;
  return response.value as MicrosoftGraph.MessageRule[];
}

async createMailRule(rule: Record<string, unknown>): Promise<MicrosoftGraph.MessageRule> {
  const client = await this.getClient();
  const result = await client
    .api('/me/mailFolders/inbox/messageRules')
    .post(rule) as MicrosoftGraph.MessageRule;
  this.cache.clear();
  return result;
}

async deleteMailRule(ruleId: string): Promise<void> {
  const client = await this.getClient();
  await client
    .api(`/me/mailFolders/inbox/messageRules/${ruleId}`)
    .delete();
  this.cache.clear();
}
```

**Step 2: Add rule ID cache to repository**

Add `rules: Map<number, string>` to the IdCache interface and initialization.

Repository methods:

```typescript
async listMailRulesAsync(): Promise<Array<{ id: number; displayName: string; sequence: number; isEnabled: boolean; conditions: unknown; actions: unknown }>> {
  const rules = await this.client.listMailRules();
  return rules.map((rule) => {
    const graphId = rule.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.rules.set(numericId, graphId);
    return {
      id: numericId,
      displayName: rule.displayName ?? '',
      sequence: rule.sequence ?? 0,
      isEnabled: rule.isEnabled ?? true,
      conditions: rule.conditions ?? {},
      actions: rule.actions ?? {},
    };
  });
}

async createMailRuleAsync(rule: Record<string, unknown>): Promise<number> {
  const created = await this.client.createMailRule(rule);
  const graphId = created.id!;
  const numericId = hashStringToNumber(graphId);
  this.idCache.rules.set(numericId, graphId);
  return numericId;
}

async deleteMailRuleAsync(ruleId: number): Promise<void> {
  const graphId = this.idCache.rules.get(ruleId);
  if (graphId == null) throw new Error(`Rule ID ${ruleId} not found in cache. Try listing mail rules first.`);
  await this.client.deleteMailRule(graphId);
  this.idCache.rules.delete(ruleId);
}
```

**Step 3: Create mail-rules.ts tool module**

Follow the pattern of `mailbox-organization.ts`. Define:
- `IMailRulesRepository` interface
- Input schemas: `ListMailRulesInput`, `CreateMailRuleInput`, `PrepareDeleteMailRuleInput`, `ConfirmDeleteMailRuleInput`
- `MailRulesTools` class with list, create, prepareDelete, confirmDelete
- Two-phase delete follows the exact pattern from mailbox-organization (token manager)

`CreateMailRuleInput` schema:

```typescript
export const CreateMailRuleInput = z.strictObject({
  display_name: z.string().describe('Rule name'),
  sequence: z.number().int().min(1).optional().describe('Rule priority order'),
  is_enabled: z.boolean().default(true).describe('Whether rule is active'),
  conditions: z.strictObject({
    from_addresses: z.array(z.string().email()).optional().describe('Match sender addresses'),
    subject_contains: z.array(z.string()).optional().describe('Subject contains any of these strings'),
    body_contains: z.array(z.string()).optional().describe('Body contains any of these strings'),
    sender_contains: z.array(z.string()).optional().describe('Sender field contains these strings'),
    has_attachments: z.boolean().optional().describe('Has attachments'),
    importance: z.enum(['low', 'normal', 'high']).optional().describe('Match importance level'),
  }).describe('Conditions that trigger the rule'),
  actions: z.strictObject({
    move_to_folder: z.number().int().positive().optional().describe('Folder ID to move to'),
    mark_as_read: z.boolean().optional().describe('Mark as read'),
    mark_importance: z.enum(['low', 'normal', 'high']).optional().describe('Set importance'),
    forward_to: z.array(z.string().email()).optional().describe('Forward to these addresses'),
    delete: z.boolean().optional().describe('Delete the message'),
    stop_processing_rules: z.boolean().optional().describe('Stop processing more rules'),
  }).describe('Actions to perform'),
});
```

**Step 4: Wire up in index.ts, write tests, README, commit**

```
git commit -m "feat: add mail rules tools (list, create, delete)"
```

---

## Phase 2: Task Features

### Task 8: List Task Lists (`list_task_lists`)

**Files:**
- Modify: `src/graph/repository.ts` — add `listTaskListsAsync`
- Modify: `src/index.ts` — add tool definition + case handler
- Modify: `tests/unit/graph/repository.test.ts`
- Modify: `README.md`

**Step 1: Repository method**

```typescript
async listTaskListsAsync(): Promise<Array<{ id: number; name: string; isDefault: boolean }>> {
  const lists = await this.client.listTaskLists();
  return lists.map((list) => {
    const graphId = list.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.taskLists.set(numericId, graphId);
    return {
      id: numericId,
      name: list.displayName ?? '',
      isDefault: list.isOwner ?? false,
    };
  });
}
```

Note: Check the actual Graph API `TodoTaskList` properties. `wellknownListName` === 'defaultList' is how to detect the default list. Adjust:

```typescript
isDefault: list.wellknownListName === 'defaultList',
```

**Step 2: Tool definition + handler**

```typescript
{
  name: 'list_task_lists',
  description: 'List all task lists (Microsoft To Do) (Graph API)',
  inputSchema: { type: 'object', properties: {}, required: [] },
},
```

Handler:
```typescript
case 'list_task_lists': {
  const lists = await repository.listTaskListsAsync();
  return { content: [{ type: 'text', text: JSON.stringify({ task_lists: lists }, null, 2) }] };
}
```

**Step 3: Tests, README, commit**

```
git commit -m "feat: add list_task_lists tool"
```

---

### Task 9: Task List Management (`rename_task_list` / `prepare_delete_task_list` / `confirm_delete_task_list`)

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add `updateTaskList`, `deleteTaskList`
- Modify: `src/graph/repository.ts` — add `renameTaskListAsync`, `deleteTaskListAsync`
- Modify: `src/index.ts` — add 3 tool definitions + handlers
- Tests + README

**Step 1: Graph client methods**

After `createTaskList`:

```typescript
async updateTaskList(listId: string, updates: Record<string, unknown>): Promise<void> {
  const client = await this.getClient();
  await client.api(`/me/todo/lists/${listId}`).patch(updates);
  this.cache.clear();
}

async deleteTaskList(listId: string): Promise<void> {
  const client = await this.getClient();
  await client.api(`/me/todo/lists/${listId}`).delete();
  this.cache.clear();
}
```

**Step 2: Repository methods**

```typescript
async renameTaskListAsync(listId: number, name: string): Promise<void> {
  const graphId = this.idCache.taskLists.get(listId);
  if (graphId == null) throw new Error(`Task list ID ${listId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
  await this.client.updateTaskList(graphId, { displayName: name });
}

async deleteTaskListAsync(listId: number): Promise<void> {
  const graphId = this.idCache.taskLists.get(listId);
  if (graphId == null) throw new Error(`Task list ID ${listId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
  await this.client.deleteTaskList(graphId);
  this.idCache.taskLists.delete(listId);
}
```

**Step 3: Tools — rename is non-destructive, delete is two-phase**

Use the existing two-phase pattern from mailbox-organization. The delete needs prepare/confirm with approval tokens. Add schemas and handlers in index.ts. The `prepare_delete_task_list` generates a token, `confirm_delete_task_list` consumes it and calls `deleteTaskListAsync`.

**Step 4: Tests, README, commit**

```
git commit -m "feat: add rename_task_list and delete_task_list tools"
```

---

### Task 10: Task Recurrence

**Files:**
- Modify: `src/graph/repository.ts` — extend `createTaskAsync` and `updateTaskAsync` params
- Modify: `src/index.ts` — extend `CreateTaskGraphInput` and `UpdateTaskGraphInput` schemas, update handlers
- Modify: `tests/unit/graph/repository.test.ts`
- Modify: `README.md`

**Step 1: Extend repository createTaskAsync**

Add `recurrence` to the params type:

```typescript
recurrence?: {
  pattern: 'daily' | 'weekly' | 'monthly' | 'yearly';
  interval?: number;
  days_of_week?: string[];
  day_of_month?: number;
  range_type: 'endDate' | 'noEnd' | 'numbered';
  start_date: string;
  end_date?: string;
  occurrences?: number;
};
```

In the method body, after building `graphTask`, add:

```typescript
if (params.recurrence != null) {
  graphTask.recurrence = {
    pattern: {
      type: params.recurrence.pattern,
      interval: params.recurrence.interval ?? 1,
      ...(params.recurrence.days_of_week != null ? { daysOfWeek: params.recurrence.days_of_week } : {}),
      ...(params.recurrence.day_of_month != null ? { dayOfMonth: params.recurrence.day_of_month } : {}),
    },
    range: {
      type: params.recurrence.range_type,
      startDate: params.recurrence.start_date,
      ...(params.recurrence.end_date != null ? { endDate: params.recurrence.end_date } : {}),
      ...(params.recurrence.occurrences != null ? { numberOfOccurrences: params.recurrence.occurrences } : {}),
    },
  };
}
```

**Step 2: Extend the index.ts schemas and handlers**

Add `recurrence` properties to `CreateTaskGraphInput` and `UpdateTaskGraphInput`. Update the `create_task` and `update_task` case handlers to pass recurrence through.

**Step 3: Tests, README, commit**

```
git commit -m "feat: add task recurrence support to create_task and update_task"
```

---

## Phase 3: Contact Features

### Task 11: Contact Folders/Groups (`list_contact_folders` / `create_contact_folder` / `prepare_delete_contact_folder` / `confirm_delete_contact_folder`)

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add 3 methods
- Modify: `src/graph/repository.ts` — add contactFolders cache + 3 methods
- Modify: `src/index.ts` — add 4 tool definitions + handlers
- Modify: `src/tools/contacts.ts` — extend `ListContactsInput` with optional `folder_id`
- Tests + README

**Step 1: Graph client methods**

```typescript
async listContactFolders(): Promise<MicrosoftGraph.ContactFolder[]> {
  const client = await this.getClient();
  const response = await client.api('/me/contactFolders').get() as PageCollection;
  return response.value as MicrosoftGraph.ContactFolder[];
}

async createContactFolder(displayName: string): Promise<MicrosoftGraph.ContactFolder> {
  const client = await this.getClient();
  const result = await client.api('/me/contactFolders').post({ displayName }) as MicrosoftGraph.ContactFolder;
  this.cache.clear();
  return result;
}

async deleteContactFolder(folderId: string): Promise<void> {
  const client = await this.getClient();
  await client.api(`/me/contactFolders/${folderId}`).delete();
  this.cache.clear();
}
```

**Step 2: Add `contactFolders: Map<number, string>` to IdCache**

**Step 3: Repository methods + enhance list_contacts with folder filter**

Add `listContactFoldersAsync`, `createContactFolderAsync`, `deleteContactFolderAsync`.

For folder-filtered contacts, add a new Graph client method `listContactsInFolder(folderId)` and repository method `listContactsInFolderAsync`.

**Step 4: Tools, tests, README, commit**

```
git commit -m "feat: add contact folder tools and folder filtering"
```

---

### Task 12: Contact Photos (`get_contact_photo` / `set_contact_photo`)

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add 2 methods
- Modify: `src/graph/repository.ts` — add 2 methods
- Modify: `src/index.ts` — add 2 tool definitions + handlers
- Tests + README

**Step 1: Graph client methods**

```typescript
async getContactPhoto(contactId: string): Promise<ArrayBuffer> {
  const client = await this.getClient();
  return await client
    .api(`/me/contacts/${contactId}/photo/$value`)
    .get() as ArrayBuffer;
}

async setContactPhoto(contactId: string, photoData: Buffer, contentType: string): Promise<void> {
  const client = await this.getClient();
  await client
    .api(`/me/contacts/${contactId}/photo/$value`)
    .header('Content-Type', contentType)
    .put(photoData);
  this.cache.clear();
}
```

**Step 2: Repository methods**

```typescript
async getContactPhotoAsync(contactId: number): Promise<{ filePath: string; contentType: string }> {
  const graphId = this.idCache.contacts.get(contactId);
  if (graphId == null) throw new Error(`Contact ID ${contactId} not found in cache...`);

  const photoData = await this.client.getContactPhoto(graphId);
  // Save to downloads directory
  const downloadDir = getDownloadDir();
  const filePath = path.join(downloadDir, `contact-${contactId}-photo.jpg`);
  fs.writeFileSync(filePath, Buffer.from(photoData));
  return { filePath, contentType: 'image/jpeg' };
}

async setContactPhotoAsync(contactId: number, filePath: string): Promise<void> {
  const graphId = this.idCache.contacts.get(contactId);
  if (graphId == null) throw new Error(`Contact ID ${contactId} not found in cache...`);

  const photoData = fs.readFileSync(filePath);
  const ext = path.extname(filePath).toLowerCase();
  const contentType = ext === '.png' ? 'image/png' : 'image/jpeg';
  await this.client.setContactPhoto(graphId, photoData, contentType);
}
```

**Step 3: Tools, tests, README, commit**

```
git commit -m "feat: add contact photo get/set tools"
```

---

## Phase 4: Calendar Features

### Task 13: Recurring Event Instances (`list_event_instances` / `update_event_instance` / `prepare_delete_event_instance` / `confirm_delete_event_instance`)

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add `listEventInstances`
- Modify: `src/graph/repository.ts` — add `listEventInstancesAsync`
- Modify: `src/index.ts` — add 4 tool definitions + handlers
- Tests + README

**Step 1: Graph client method**

```typescript
async listEventInstances(
  eventId: string,
  startDateTime: string,
  endDateTime: string
): Promise<MicrosoftGraph.Event[]> {
  const client = await this.getClient();
  const response = await client
    .api(`/me/events/${eventId}/instances`)
    .query({ startDateTime, endDateTime })
    .select('id,subject,start,end,location,isAllDay,isCancelled,organizer,recurrence,bodyPreview')
    .get() as PageCollection;
  return response.value as MicrosoftGraph.Event[];
}
```

**Step 2: Repository method**

```typescript
async listEventInstancesAsync(
  eventId: number,
  startDate: string,
  endDate: string
): Promise<EventRow[]> {
  const graphId = this.idCache.events.get(eventId);
  if (graphId == null) throw new Error(`Event ID ${eventId} not found in cache...`);

  const instances = await this.client.listEventInstances(graphId, startDate, endDate);
  // Cache instance IDs (they're regular event IDs)
  for (const inst of instances) {
    if (inst.id != null) {
      this.idCache.events.set(hashStringToNumber(inst.id), inst.id);
    }
  }
  return instances.map((e) => mapEventToEventRow(e));
}
```

**Step 3: Tools**

- `list_event_instances` — input: `{ event_id, start_date, end_date }`
- `update_event_instance` — reuse existing `update_event` handler (instance IDs work the same as event IDs)
- `prepare_delete_event_instance` / `confirm_delete_event_instance` — reuse existing delete event two-phase pattern

Note: Since event instance IDs get cached as regular event IDs, the existing `update_event` and `delete_event` tools already work on instances. We may only need `list_event_instances` as a new tool, with documentation noting that `update_event` and `delete_event` work on instance IDs. This avoids unnecessary tool duplication.

**Step 4: Tests, README, commit**

```
git commit -m "feat: add list_event_instances and recurring event instance management"
```

---

## Final Step

After all 13 features are implemented:

1. Run full test suite: `npx vitest run`
2. Run type check: `npx tsc --noEmit`
3. Update the tool count in README.md header
4. Final commit for any documentation touchups
