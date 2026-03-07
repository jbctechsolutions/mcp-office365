# Phase 0 + Phase 1: Package Rename & Outlook Gaps Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Rename package from mcp-outlook-mac to mcp-office365-mac, add new OAuth scopes for M365 services, and implement 10 missing Outlook features (~22 new tools + 2 extended).

**Architecture:** Each feature follows the 4-layer pattern: Graph client method -> Repository method -> Tool schema/handler in index.ts -> Tests. Phase 0 is infrastructure (rename, scopes). Phase 1 adds Outlook gap features one commit each.

**Tech Stack:** TypeScript, Microsoft Graph API, Zod schemas, Vitest tests.

**Run tests:** `npx vitest run` (full suite), or `npx vitest run tests/unit/path/to/test.ts` (specific)

**Type check:** `npx tsc --noEmit`

---

## Phase 0: Package Rename + New OAuth Scopes

### Task 1: Package Rename

Rename all references from `mcp-outlook-mac` to `mcp-office365-mac` and `outlook` branding to `office365`/`M365`.

**Files:**
- Modify: `package.json`
- Modify: `README.md`
- Modify: `src/index.ts` (server name in MCP registration)

**Step 1: Update package.json**

Change these fields:
```json
{
  "name": "@jbctechsolutions/mcp-office365-mac",
  "description": "MCP server for Microsoft 365 with dual backend support (AppleScript for classic Outlook, Graph API for new Outlook/M365). Read/write access to mail, calendar, contacts, tasks, teams, planner, and more.",
  "bin": {
    "mcp-office365-mac": "./dist/index.js"
  }
}
```

Also update the `keywords` array — add `"office365"`, `"microsoft-365"`, `"teams"`, `"planner"`. Keep existing keywords.

**Step 2: Update README.md**

- Line 1: Change title to `# Office 365 MCP Server`
- Line 7: Update description to reference Microsoft 365
- Line 11: Keep tool count as-is for now
- Update Quick Start section: change `npx -y mcp-outlook-mac` to `npx -y mcp-office365-mac`
- Update `npx @jbctechsolutions/mcp-outlook-mac auth` to `npx @jbctechsolutions/mcp-office365-mac auth`

**Step 3: Update src/index.ts server name**

Find the MCP server instantiation (around line 2044):
```typescript
const server = new Server(
  {
    name: 'outlook-mcp-server',
```
Change to:
```typescript
const server = new Server(
  {
    name: 'office365-mcp-server',
```

**Step 4: Run tests, type check, commit**

```bash
npx vitest run && npx tsc --noEmit
git commit -m "chore: rename package from mcp-outlook-mac to mcp-office365-mac"
```

---

### Task 2: Add New OAuth Scopes

Add M365 scopes for Teams, People, Planner, and Presence.

**Files:**
- Modify: `src/graph/auth/config.ts`
- Modify: `tests/unit/graph/auth/config.test.ts`

**Step 1: Update GRAPH_SCOPES**

In `src/graph/auth/config.ts`, find `GRAPH_SCOPES` (line 33). Add new scopes:

```typescript
export const GRAPH_SCOPES = [
  // Outlook (existing)
  'Mail.ReadWrite',
  'Calendars.ReadWrite',
  'Contacts.ReadWrite',
  'Tasks.ReadWrite',
  'User.Read',
  'offline_access',
  // Teams
  'ChannelMessage.Read.All',
  'ChannelMessage.Send',
  'Channel.ReadBasic.All',
  'Team.ReadBasic.All',
  'Chat.ReadWrite',
  'ChatMessage.Send',
  // People & Presence
  'People.Read',
  'User.ReadBasic.All',
  'Presence.Read.All',
  // Planner
  'Group.Read.All',
] as const;
```

**Step 2: Update tests**

In `tests/unit/graph/auth/config.test.ts`, find any test that asserts on the scope count or specific scopes. Update to match the new list of 16 scopes.

**Step 3: Run tests, commit**

```bash
npx vitest run && npx tsc --noEmit
git commit -m "feat: add M365 OAuth scopes for Teams, People, Planner"
```

---

## Phase 1: Outlook Gap Features

### Task 3: Automatic Replies (OOF) — `get_automatic_replies` / `set_automatic_replies`

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add 2 methods
- Modify: `src/graph/repository.ts` — add 2 methods
- Modify: `src/index.ts` — add 2 tool definitions + handlers
- Modify: `tests/unit/graph/client/api-calls.test.ts`
- Modify: `tests/unit/graph/repository.test.ts`
- Modify: `README.md`
- Modify: `tests/e2e/mcp-client.test.ts`

**Step 1: Graph client methods**

Add after the Mail Rules section in `graph-client.ts`:

```typescript
// ===========================================================================
// Mailbox Settings
// ===========================================================================

async getAutomaticReplies(): Promise<Record<string, unknown>> {
  const client = await this.getClient();
  return await client
    .api('/me/mailboxSettings/automaticRepliesSetting')
    .get() as Record<string, unknown>;
}

async setAutomaticReplies(settings: Record<string, unknown>): Promise<void> {
  const client = await this.getClient();
  await client
    .api('/me/mailboxSettings')
    .patch({ automaticRepliesSetting: settings });
}
```

**Step 2: Repository methods**

```typescript
async getAutomaticRepliesAsync(): Promise<{
  status: string;
  externalAudience: string;
  internalReplyMessage: string;
  externalReplyMessage: string;
  scheduledStartDateTime: string | null;
  scheduledEndDateTime: string | null;
}> {
  const settings = await this.client.getAutomaticReplies();
  return {
    status: (settings as any).status ?? 'disabled',
    externalAudience: (settings as any).externalAudience ?? 'none',
    internalReplyMessage: (settings as any).internalReplyMessage ?? '',
    externalReplyMessage: (settings as any).externalReplyMessage ?? '',
    scheduledStartDateTime: (settings as any).scheduledStartDateTime?.dateTime ?? null,
    scheduledEndDateTime: (settings as any).scheduledEndDateTime?.dateTime ?? null,
  };
}

async setAutomaticRepliesAsync(params: {
  status: 'disabled' | 'alwaysEnabled' | 'scheduled';
  externalAudience?: 'none' | 'contactsOnly' | 'all';
  internalReplyMessage?: string;
  externalReplyMessage?: string;
  scheduledStartDateTime?: string;
  scheduledEndDateTime?: string;
}): Promise<void> {
  const settings: Record<string, unknown> = {
    status: params.status,
  };
  if (params.externalAudience != null) settings['externalAudience'] = params.externalAudience;
  if (params.internalReplyMessage != null) settings['internalReplyMessage'] = params.internalReplyMessage;
  if (params.externalReplyMessage != null) settings['externalReplyMessage'] = params.externalReplyMessage;
  if (params.scheduledStartDateTime != null) {
    settings['scheduledStartDateTime'] = { dateTime: params.scheduledStartDateTime, timeZone: 'UTC' };
  }
  if (params.scheduledEndDateTime != null) {
    settings['scheduledEndDateTime'] = { dateTime: params.scheduledEndDateTime, timeZone: 'UTC' };
  }
  await this.client.setAutomaticReplies(settings);
}
```

**Step 3: Zod schemas + tool definitions + handlers in index.ts**

```typescript
const SetAutomaticRepliesInput = z.strictObject({
  status: z.enum(['disabled', 'alwaysEnabled', 'scheduled']).describe('OOF status'),
  external_audience: z.enum(['none', 'contactsOnly', 'all']).optional().describe('Who sees external reply'),
  internal_reply_message: z.string().optional().describe('Reply message for internal senders (HTML)'),
  external_reply_message: z.string().optional().describe('Reply message for external senders (HTML)'),
  scheduled_start: z.string().optional().describe('Schedule start (ISO 8601)'),
  scheduled_end: z.string().optional().describe('Schedule end (ISO 8601)'),
});
```

Tool definitions:
```typescript
{
  name: 'get_automatic_replies',
  description: 'Get automatic replies (out-of-office) settings (Graph API)',
  inputSchema: { type: 'object', properties: {}, required: [] },
},
{
  name: 'set_automatic_replies',
  description: 'Set automatic replies (out-of-office) settings (Graph API)',
  inputSchema: zodToJsonSchema(SetAutomaticRepliesInput) as Tool['inputSchema'],
},
```

Add both to `GRAPH_ONLY_TOOL_NAMES`. Handler for `get_automatic_replies` calls `repo.getAutomaticRepliesAsync()`. Handler for `set_automatic_replies` parses input and calls `repo.setAutomaticRepliesAsync(...)`.

**Step 4: Tests, README, e2e count, commit**

```
git commit -m "feat: add automatic replies (OOF) tools"
```

---

### Task 4: Mailbox Settings — `get_mailbox_settings` / `update_mailbox_settings`

**Files:** Same pattern as Task 3.

**Step 1: Graph client methods**

```typescript
async getMailboxSettings(): Promise<Record<string, unknown>> {
  const client = await this.getClient();
  return await client.api('/me/mailboxSettings').get() as Record<string, unknown>;
}

async updateMailboxSettings(settings: Record<string, unknown>): Promise<void> {
  const client = await this.getClient();
  await client.api('/me/mailboxSettings').patch(settings);
}
```

**Step 2: Repository methods**

```typescript
async getMailboxSettingsAsync(): Promise<{
  language: string | null;
  timeZone: string | null;
  dateFormat: string | null;
  timeFormat: string | null;
  workingHours: unknown | null;
}> {
  const settings = await this.client.getMailboxSettings();
  return {
    language: (settings as any).language?.locale ?? null,
    timeZone: (settings as any).timeZone ?? null,
    dateFormat: (settings as any).dateFormat ?? null,
    timeFormat: (settings as any).timeFormat ?? null,
    workingHours: (settings as any).workingHours ?? null,
  };
}

async updateMailboxSettingsAsync(params: {
  language?: string;
  timeZone?: string;
  dateFormat?: string;
  timeFormat?: string;
}): Promise<void> {
  const settings: Record<string, unknown> = {};
  if (params.language != null) settings['language'] = { locale: params.language };
  if (params.timeZone != null) settings['timeZone'] = params.timeZone;
  if (params.dateFormat != null) settings['dateFormat'] = params.dateFormat;
  if (params.timeFormat != null) settings['timeFormat'] = params.timeFormat;
  await this.client.updateMailboxSettings(settings);
}
```

**Step 3: Schemas, tools, handlers in index.ts**

```typescript
const UpdateMailboxSettingsInput = z.strictObject({
  language: z.string().optional().describe('Locale code (e.g. en-US)'),
  time_zone: z.string().optional().describe('Time zone (e.g. America/New_York)'),
  date_format: z.string().optional().describe('Date format string'),
  time_format: z.string().optional().describe('Time format string'),
});
```

Add both to `GRAPH_ONLY_TOOL_NAMES`.

**Step 4: Tests, README, e2e count, commit**

```
git commit -m "feat: add mailbox settings tools"
```

---

### Task 5: Master Categories — `list_categories` / `create_category` / `prepare_delete_category` / `confirm_delete_category`

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add 3 methods
- Modify: `src/graph/repository.ts` — add category ID cache + 3 methods
- Modify: `src/approval/types.ts` — add `'delete_category'` + `'category'`
- Modify: `src/index.ts` — add 4 tool definitions + handlers
- Tests + README + e2e count

**Step 1: Graph client methods**

```typescript
// ===========================================================================
// Categories
// ===========================================================================

async listMasterCategories(): Promise<MicrosoftGraph.OutlookCategory[]> {
  const client = await this.getClient();
  const response = await client
    .api('/me/outlook/masterCategories')
    .get() as PageCollection;
  return response.value as MicrosoftGraph.OutlookCategory[];
}

async createMasterCategory(displayName: string, color: string): Promise<MicrosoftGraph.OutlookCategory> {
  const client = await this.getClient();
  const result = await client
    .api('/me/outlook/masterCategories')
    .post({ displayName, color }) as MicrosoftGraph.OutlookCategory;
  this.cache.clear();
  return result;
}

async deleteMasterCategory(categoryId: string): Promise<void> {
  const client = await this.getClient();
  await client.api(`/me/outlook/masterCategories/${categoryId}`).delete();
  this.cache.clear();
}
```

**Step 2: Add `categories: Map<number, string>` to IdCache**

**Step 3: Repository methods**

```typescript
async listCategoriesAsync(): Promise<Array<{ id: number; name: string; color: string }>> {
  const cats = await this.client.listMasterCategories();
  return cats.map((cat) => {
    const graphId = cat.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.categories.set(numericId, graphId);
    return {
      id: numericId,
      name: cat.displayName ?? '',
      color: cat.color ?? '',
    };
  });
}

async createCategoryAsync(displayName: string, color: string): Promise<number> {
  const created = await this.client.createMasterCategory(displayName, color);
  const graphId = created.id!;
  const numericId = hashStringToNumber(graphId);
  this.idCache.categories.set(numericId, graphId);
  return numericId;
}

async deleteCategoryAsync(categoryId: number): Promise<void> {
  const graphId = this.idCache.categories.get(categoryId);
  if (graphId == null) throw new Error(`Category ID ${categoryId} not found in cache. Try listing categories first.`);
  await this.client.deleteMasterCategory(graphId);
  this.idCache.categories.delete(categoryId);
}
```

**Step 4: Approval types**

Add `'delete_category'` to OperationType, `'category'` to TargetType.

**Step 5: Schemas + tools + handlers**

```typescript
const CreateCategoryInput = z.strictObject({
  name: z.string().min(1).describe('Category name'),
  color: z.enum([
    'preset0', 'preset1', 'preset2', 'preset3', 'preset4', 'preset5',
    'preset6', 'preset7', 'preset8', 'preset9', 'preset10', 'preset11',
    'preset12', 'preset13', 'preset14', 'preset15', 'preset16', 'preset17',
    'preset18', 'preset19', 'preset20', 'preset21', 'preset22', 'preset23', 'preset24',
    'none',
  ]).describe('Category color preset'),
});

const PrepareDeleteCategoryInput = z.strictObject({
  category_id: z.number().int().positive().describe('Category ID to delete'),
});

const ConfirmDeleteCategoryInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_category'),
});
```

4 tools, all in `GRAPH_ONLY_TOOL_NAMES`. Two-phase delete follows mail rules pattern.

**Step 6: Tests, README, e2e count, commit**

```
git commit -m "feat: add master category management tools"
```

---

### Task 6: Focused Inbox — `list_focused_overrides` / `create_focused_override` / `prepare_delete_focused_override` / `confirm_delete_focused_override`

**Files:** Same pattern as Task 5.

**Step 1: Graph client methods**

```typescript
// ===========================================================================
// Focused Inbox
// ===========================================================================

async listFocusedOverrides(): Promise<MicrosoftGraph.InferenceClassificationOverride[]> {
  const client = await this.getClient();
  const response = await client
    .api('/me/inferenceClassification/overrides')
    .get() as PageCollection;
  return response.value as MicrosoftGraph.InferenceClassificationOverride[];
}

async createFocusedOverride(
  senderAddress: string,
  classifyAs: 'focused' | 'other'
): Promise<MicrosoftGraph.InferenceClassificationOverride> {
  const client = await this.getClient();
  const result = await client
    .api('/me/inferenceClassification/overrides')
    .post({
      classifyAs,
      senderEmailAddress: { address: senderAddress },
    }) as MicrosoftGraph.InferenceClassificationOverride;
  return result;
}

async deleteFocusedOverride(overrideId: string): Promise<void> {
  const client = await this.getClient();
  await client
    .api(`/me/inferenceClassification/overrides/${overrideId}`)
    .delete();
}
```

**Step 2: Add `focusedOverrides: Map<number, string>` to IdCache**

**Step 3: Repository methods** — list (caches IDs), create (returns numeric ID), delete (resolves from cache)

**Step 4: Approval types** — add `'delete_focused_override'` to OperationType, `'focused_override'` to TargetType

**Step 5: Schemas, tools, handlers** — 4 tools, all `GRAPH_ONLY_TOOL_NAMES`

```typescript
const CreateFocusedOverrideInput = z.strictObject({
  sender_address: z.string().email().describe('Sender email address'),
  classify_as: z.enum(['focused', 'other']).describe('Classification'),
});
```

**Step 6: Tests, README, e2e count, commit**

```
git commit -m "feat: add focused inbox override tools"
```

---

### Task 7: Mail Tips — `get_mail_tips`

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add 1 method
- Modify: `src/graph/repository.ts` — add 1 method
- Modify: `src/index.ts` — add 1 tool definition + handler
- Tests + README + e2e count

**Step 1: Graph client method**

```typescript
async getMailTips(emailAddresses: string[]): Promise<Record<string, unknown>[]> {
  const client = await this.getClient();
  const response = await client
    .api('/me/getMailTips')
    .post({
      emailAddresses,
      mailTipsOptions: 'automaticReplies,mailboxFullStatus,maxMessageSize,deliveryRestriction,externalMemberCount',
    }) as { value: Record<string, unknown>[] };
  return response.value;
}
```

**Step 2: Repository method**

```typescript
async getMailTipsAsync(emailAddresses: string[]): Promise<Array<{
  emailAddress: string;
  automaticReplies: { message: string } | null;
  mailboxFull: boolean;
  deliveryRestricted: boolean;
  externalMemberCount: number;
  maxMessageSize: number;
}>> {
  const tips = await this.client.getMailTips(emailAddresses);
  return tips.map((tip: any) => ({
    emailAddress: tip.emailAddress?.address ?? '',
    automaticReplies: tip.automaticReplies?.message ? { message: tip.automaticReplies.message } : null,
    mailboxFull: tip.mailboxFull ?? false,
    deliveryRestricted: tip.deliveryRestricted ?? false,
    externalMemberCount: tip.externalMemberCount ?? 0,
    maxMessageSize: tip.maxMessageSize ?? 0,
  }));
}
```

**Step 3: Schema + tool + handler**

```typescript
const GetMailTipsInput = z.strictObject({
  email_addresses: z.array(z.string().email()).min(1).max(20).describe('Email addresses to check'),
});
```

1 tool in `GRAPH_ONLY_TOOL_NAMES`.

**Step 4: Tests, README, e2e count, commit**

```
git commit -m "feat: add get_mail_tips tool"
```

---

### Task 8: Message Headers / MIME — `get_message_headers` / `get_message_mime`

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add 2 methods
- Modify: `src/graph/repository.ts` — add 2 methods
- Modify: `src/index.ts` — add 2 tool definitions + handlers
- Tests + README + e2e count

**Step 1: Graph client methods**

```typescript
async getMessageHeaders(messageId: string): Promise<Array<{ name: string; value: string }>> {
  const client = await this.getClient();
  const message = await client
    .api(`/me/messages/${messageId}`)
    .select('internetMessageHeaders')
    .get() as MicrosoftGraph.Message;
  return (message.internetMessageHeaders ?? []) as Array<{ name: string; value: string }>;
}

async getMessageMime(messageId: string): Promise<string> {
  const client = await this.getClient();
  return await client
    .api(`/me/messages/${messageId}/$value`)
    .get() as string;
}
```

**Step 2: Repository methods**

```typescript
async getMessageHeadersAsync(emailId: number): Promise<Array<{ name: string; value: string }>> {
  const graphId = this.idCache.messages.get(emailId);
  if (graphId == null) throw new Error(`Email ID ${emailId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
  return await this.client.getMessageHeaders(graphId);
}

async getMessageMimeAsync(emailId: number): Promise<{ filePath: string }> {
  const graphId = this.idCache.messages.get(emailId);
  if (graphId == null) throw new Error(`Email ID ${emailId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
  const mime = await this.client.getMessageMime(graphId);
  const downloadDir = getDownloadDir();
  const filePath = path.join(downloadDir, `email-${emailId}.eml`);
  fs.writeFileSync(filePath, mime, 'utf-8');
  return { filePath };
}
```

Note: `getDownloadDir`, `fs`, `path` are already imported in repository.ts from Task 12 (contact photos).

**Step 3: Schemas + tools + handlers**

```typescript
const GetMessageHeadersInput = z.strictObject({
  email_id: z.number().int().positive().describe('Email ID'),
});

const GetMessageMimeInput = z.strictObject({
  email_id: z.number().int().positive().describe('Email ID'),
});
```

2 tools in `GRAPH_ONLY_TOOL_NAMES`.

**Step 4: Tests, README, e2e count, commit**

```
git commit -m "feat: add message headers and MIME content tools"
```

---

### Task 9: Calendar Groups — `list_calendar_groups` / `create_calendar_group`

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add 2 methods
- Modify: `src/graph/repository.ts` — add calendarGroups to IdCache + 2 methods
- Modify: `src/index.ts` — add 2 tool definitions + handlers
- Tests + README + e2e count

**Step 1: Graph client methods**

```typescript
// ===========================================================================
// Calendar Groups
// ===========================================================================

async listCalendarGroups(): Promise<MicrosoftGraph.CalendarGroup[]> {
  const client = await this.getClient();
  const response = await client.api('/me/calendarGroups').get() as PageCollection;
  return response.value as MicrosoftGraph.CalendarGroup[];
}

async createCalendarGroup(name: string): Promise<MicrosoftGraph.CalendarGroup> {
  const client = await this.getClient();
  const result = await client.api('/me/calendarGroups').post({ name }) as MicrosoftGraph.CalendarGroup;
  this.cache.clear();
  return result;
}
```

**Step 2: Add `calendarGroups: Map<number, string>` to IdCache**

**Step 3: Repository methods** — list (map + cache IDs), create (cache + return numeric ID)

**Step 4: Schemas, tools, handlers**

```typescript
const CreateCalendarGroupInput = z.strictObject({
  name: z.string().min(1).describe('Calendar group name'),
});
```

2 tools in `GRAPH_ONLY_TOOL_NAMES`.

**Step 5: Tests, README, e2e count, commit**

```
git commit -m "feat: add calendar group tools"
```

---

### Task 10: Calendar Permissions — `list_calendar_permissions` / `create_calendar_permission` / `prepare_delete_calendar_permission` / `confirm_delete_calendar_permission`

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add 3 methods
- Modify: `src/graph/repository.ts` — add calendarPermissions to IdCache + 3 methods
- Modify: `src/approval/types.ts` — add `'delete_calendar_permission'` + `'calendar_permission'`
- Modify: `src/index.ts` — add 4 tool definitions + handlers
- Tests + README + e2e count

**Step 1: Graph client methods**

```typescript
async listCalendarPermissions(calendarId: string): Promise<MicrosoftGraph.CalendarPermission[]> {
  const client = await this.getClient();
  const response = await client
    .api(`/me/calendars/${calendarId}/calendarPermissions`)
    .get() as PageCollection;
  return response.value as MicrosoftGraph.CalendarPermission[];
}

async createCalendarPermission(
  calendarId: string,
  permission: Record<string, unknown>
): Promise<MicrosoftGraph.CalendarPermission> {
  const client = await this.getClient();
  const result = await client
    .api(`/me/calendars/${calendarId}/calendarPermissions`)
    .post(permission) as MicrosoftGraph.CalendarPermission;
  this.cache.clear();
  return result;
}

async deleteCalendarPermission(calendarId: string, permissionId: string): Promise<void> {
  const client = await this.getClient();
  await client
    .api(`/me/calendars/${calendarId}/calendarPermissions/${permissionId}`)
    .delete();
  this.cache.clear();
}
```

**Step 2: Add `calendarPermissions: Map<number, { calendarId: string; permissionId: string }>` to IdCache**

**Step 3: Repository methods** — list (resolve calendar ID from idCache.events or a new calendar cache), create, delete

**Step 4: Schemas**

```typescript
const ListCalendarPermissionsInput = z.strictObject({
  calendar_id: z.number().int().positive().describe('Calendar ID'),
});

const CreateCalendarPermissionInput = z.strictObject({
  calendar_id: z.number().int().positive().describe('Calendar ID'),
  email_address: z.string().email().describe('Email of person to share with'),
  role: z.enum(['read', 'write', 'delegateWithoutPrivateEventAccess', 'delegateWithPrivateEventAccess']).describe('Permission level'),
});

const PrepareDeleteCalendarPermissionInput = z.strictObject({
  permission_id: z.number().int().positive().describe('Calendar permission ID'),
});

const ConfirmDeleteCalendarPermissionInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_calendar_permission'),
});
```

4 tools in `GRAPH_ONLY_TOOL_NAMES`. Two-phase delete pattern.

**Step 5: Tests, README, e2e count, commit**

```
git commit -m "feat: add calendar permission tools"
```

---

### Task 11: Room Lists — `list_room_lists` / `list_rooms`

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add 2 methods
- Modify: `src/graph/repository.ts` — add 2 methods
- Modify: `src/index.ts` — add 2 tool definitions + handlers
- Tests + README + e2e count

**Step 1: Graph client methods**

```typescript
async listRoomLists(): Promise<MicrosoftGraph.EmailAddress[]> {
  const client = await this.getClient();
  const response = await client
    .api('/me/findRoomLists')
    .get() as { value: MicrosoftGraph.EmailAddress[] };
  return response.value;
}

async listRooms(roomListEmail?: string): Promise<MicrosoftGraph.EmailAddress[]> {
  const client = await this.getClient();
  const endpoint = roomListEmail != null
    ? `/me/findRooms(RoomList='${roomListEmail}')`
    : '/me/findRooms';
  const response = await client.api(endpoint).get() as { value: MicrosoftGraph.EmailAddress[] };
  return response.value;
}
```

**Step 2: Repository methods** — thin wrappers, no ID caching needed (room lists use email addresses, not IDs)

**Step 3: Schemas**

```typescript
const ListRoomsInput = z.strictObject({
  room_list_email: z.string().email().optional().describe('Room list email to filter by (from list_room_lists)'),
});
```

2 tools in `GRAPH_ONLY_TOOL_NAMES`.

**Step 4: Tests, README, e2e count, commit**

```
git commit -m "feat: add room list and room finder tools"
```

---

### Task 12: Online Meetings — extend `create_event` / `update_event`

**Files:**
- Modify: `src/graph/repository.ts` — extend create/update event params
- Modify: `src/index.ts` — extend CreateEventGraphInput and UpdateEventGraphInput schemas
- Modify: `tests/unit/graph/repository.test.ts`
- Modify: `README.md`

**Step 1: Extend repository createEventAsync**

Find `createEventAsync` in repository.ts. Add online meeting support to the params type and the graph event object:

```typescript
if (params.is_online_meeting) {
  (graphEvent as any).isOnlineMeeting = true;
  (graphEvent as any).onlineMeetingProvider = params.online_meeting_provider ?? 'teamsForBusiness';
}
```

The response should include `onlineMeeting.joinUrl` in the returned event data.

**Step 2: Extend schemas in index.ts**

Add to both `CreateEventGraphInput` and `UpdateEventGraphInput`:

```typescript
is_online_meeting: z.boolean().optional().describe('Create as online Teams meeting'),
online_meeting_provider: z.enum(['teamsForBusiness', 'skypeForBusiness', 'skypeForConsumer']).optional().describe('Online meeting provider (default: teamsForBusiness)'),
```

**Step 3: Extend event mapping**

Ensure the event mapper includes `onlineMeeting.joinUrl` in the returned EventRow. Find `mapEventToEventRow` in `src/graph/mappers/` and add:

```typescript
onlineMeetingUrl: event.onlineMeeting?.joinUrl ?? null,
```

**Step 4: Tests, README update (update create_event/update_event descriptions), commit**

```
git commit -m "feat: add online meeting support to create_event and update_event"
```

---

## Summary

| Task | Feature | New Tools | Commit |
|------|---------|-----------|--------|
| 1 | Package rename | 0 | `chore: rename package from mcp-outlook-mac to mcp-office365-mac` |
| 2 | OAuth scopes | 0 | `feat: add M365 OAuth scopes for Teams, People, Planner` |
| 3 | Automatic replies | 2 | `feat: add automatic replies (OOF) tools` |
| 4 | Mailbox settings | 2 | `feat: add mailbox settings tools` |
| 5 | Master categories | 4 | `feat: add master category management tools` |
| 6 | Focused inbox | 4 | `feat: add focused inbox override tools` |
| 7 | Mail tips | 1 | `feat: add get_mail_tips tool` |
| 8 | Message headers/MIME | 2 | `feat: add message headers and MIME content tools` |
| 9 | Calendar groups | 2 | `feat: add calendar group tools` |
| 10 | Calendar permissions | 4 | `feat: add calendar permission tools` |
| 11 | Room lists | 2 | `feat: add room list and room finder tools` |
| 12 | Online meetings | 0 (2 extended) | `feat: add online meeting support to create_event and update_event` |
