# Graph API Write Operations Design

## Overview

Add write operations to the Graph API backend across five priority areas: mail drafts/sending, attachments, calendar writes, contact writes, and task writes. The AppleScript backend already supports some of these (send_email, create_event, respond_to_event); the Graph backend needs parity plus new capabilities that Graph makes possible (drafts, reply, forward).

## Design Decisions

- **All send operations use two-phase prepare/confirm approval** (send_email, send_draft, reply_email, forward_email) — sending is irreversible
- **All delete operations use two-phase prepare/confirm approval** (delete_event, delete_contact, delete_task) — consistent with existing email/folder deletes
- **Full attachment support** including upload sessions for files >3MB
- **Single `reply_email` tool** with `reply_all: boolean` defaulting to `true`
- **Configurable download directory** via `MCP_OUTLOOK_DOWNLOAD_DIR` env var, fallback to OS temp dir
- **New `MailSendTools` class** for send approval logic, parallel to `MailboxOrganizationTools`
- **Five implementation phases** matching priority order, each independently testable

## Phase 1: Mail Drafts & Sending

### New Tools (12)

**Non-destructive (no approval):**

| Tool | Endpoint | Description |
|------|----------|-------------|
| `create_draft` | `POST /me/messages` | Create draft with `isDraft: true`. Params: `to`, `cc`, `bcc`, `subject`, `body`, `body_type`. Returns draft numeric ID. |
| `update_draft` | `PATCH /me/messages/{id}` | Update existing draft. Same params, all optional. Validates message is still a draft. |
| `list_drafts` | `GET /me/mailFolders/drafts/messages` | List all drafts. Reuses `listMessages` pattern with well-known `drafts` folder. |

**Destructive (two-phase approval):**

| Prepare Tool | Confirm Tool | Endpoint | Description |
|-------------|-------------|----------|-------------|
| `prepare_send_draft` | `confirm_send_draft` | `POST /me/messages/{id}/send` | Send existing draft. Prepare shows subject, recipients, body preview. |
| `prepare_send_email` | `confirm_send_email` | `POST /me/sendMail` | Send immediately without draft. Prepare validates params, shows preview. |
| `prepare_reply_email` | `confirm_reply_email` | `POST /me/messages/{id}/reply` or `/replyAll` | Reply to message. Params: `message_id`, `comment`, `reply_all` (default `true`). |
| `prepare_forward_email` | `confirm_forward_email` | `POST /me/messages/{id}/forward` | Forward message. Params: `message_id`, `to_recipients`, `comment`. |

### Approval Token Hashing

| Operation | Hash Input |
|-----------|-----------|
| Draft send | `SHA256(id:subject:recipientCount)` |
| Direct send | `SHA256(subject:toCount:ccCount:bccCount)` |
| Reply | `SHA256(originalId:comment_length:replyAll)` |
| Forward | `SHA256(originalId:recipientCount)` |

### Layer Changes

1. **GraphClient** — new methods: `createDraft()`, `updateDraft()`, `sendDraft()`, `sendMail()`, `replyMessage()`, `forwardMessage()`. All call `this.cache.clear()` on mutation.
2. **GraphRepository** — async wrappers, idCache updates for created drafts.
3. **New `MailSendTools` class** — handles prepare/confirm approval for all send operations. Parallel to `MailboxOrganizationTools`.
4. **index.ts** — new tool definitions + `handleGraphToolCall` cases.

## Phase 2: Attachment Support

### New Tools (2 + integration)

| Tool | Endpoint | Description |
|------|----------|-------------|
| `list_attachments` | `GET /me/messages/{id}/attachments` | Returns `{ id, name, size, contentType, isInline }`. Adds to idCache. |
| `download_attachment` | `GET /me/messages/{id}/attachments/{attachmentId}` | Decodes base64 `contentBytes`, writes to configured dir. Returns `{ filePath, name, size, contentType }`. |

### Attachment Integration

`create_draft`, `update_draft`, and `send_email` accept optional `attachments` param:

```typescript
attachments: Array<{ filePath: string; name?: string; contentType?: string }>
```

**Size-based routing:**
- **<= 3MB**: Inline base64 via `POST /me/messages/{id}/attachments` with `contentBytes`
- **> 3MB**: Upload session via `POST /me/messages/{id}/attachments/createUploadSession`, chunked `PUT` uploads (~3.75MB chunks)

For `send_email` (no pre-existing draft): create hidden draft first, attach files, then send.

### New idCache Bucket

```typescript
attachments: Map<number, { messageId: string; attachmentId: string }>
```

### Download Directory

```typescript
const downloadDir = process.env.MCP_OUTLOOK_DOWNLOAD_DIR || os.tmpdir();
const attachmentDir = path.join(downloadDir, 'mcp-outlook-attachments');
```

Filenames sanitized to prevent path traversal. Duplicate names get numeric suffix.

## Phase 3: Calendar Write Operations

### New Tools (7)

**Non-destructive (no approval):**

| Tool | Endpoint | Description |
|------|----------|-------------|
| `create_event` | `POST /me/events` or `POST /me/calendars/{id}/events` | Params: `subject`, `start`, `end`, `timezone`, `location`, `body`, `body_type`, `attendees`, `is_all_day`, `recurrence`, `calendar_id`. |
| `update_event` | `PATCH /me/events/{id}` | Same params, all optional. |
| `respond_to_event` | `POST /me/events/{id}/accept\|decline\|tentativelyAccept` | Params: `event_id`, `response`, `send_response` (default true), `comment`. No approval — reversible. |

**Destructive (two-phase approval):**

| Prepare Tool | Confirm Tool | Endpoint | Description |
|-------------|-------------|----------|-------------|
| `prepare_delete_event` | `confirm_delete_event` | `DELETE /me/events/{id}` | Prepare shows subject, date/time, attendees. Hash: `SHA256(id:subject:startDateTime)`. |

### Recurrence Support

Passed through to Graph API as `PatternedRecurrence`:

```typescript
recurrence: {
  pattern: { type: 'daily'|'weekly'|'monthly'|'yearly', interval: number, daysOfWeek?: string[] },
  range: { type: 'endDate'|'noEnd'|'numbered', startDate: string, endDate?: string, numberOfOccurrences?: number }
}
```

## Phase 4: Contact Write Operations

### New Tools (5)

**Non-destructive (no approval):**

| Tool | Endpoint | Description |
|------|----------|-------------|
| `create_contact` | `POST /me/contacts` | Params: `given_name`, `surname`, `email`, `phone`, `mobile_phone`, `company`, `job_title`, address fields. |
| `update_contact` | `PATCH /me/contacts/{id}` | Same params, all optional. |

**Destructive (two-phase approval):**

| Prepare Tool | Confirm Tool | Endpoint | Description |
|-------------|-------------|----------|-------------|
| `prepare_delete_contact` | `confirm_delete_contact` | `DELETE /me/contacts/{id}` | Prepare shows name, email, company. Hash: `SHA256(id:displayName:emailAddress)`. |

### Field Mapping

| Tool Param | Graph API Field |
|-----------|----------------|
| `given_name` | `givenName` |
| `surname` | `surname` |
| `email` | `emailAddresses: [{ address, name }]` |
| `phone` | `businessPhones: [phone]` |
| `mobile_phone` | `mobilePhone` |
| `company` | `companyName` |
| `job_title` | `jobTitle` |
| Address fields | `businessAddress: { street, city, state, postalCode, countryOrRegion }` |

## Phase 5: Task Write Operations

### New Tools (8)

**Non-destructive (no approval):**

| Tool | Endpoint | Description |
|------|----------|-------------|
| `create_task` | `POST /me/todo/lists/{listId}/tasks` | Params: `list_id` (required), `title`, `body`, `due_date`, `importance`, `reminder_date`. |
| `update_task` | `PATCH /me/todo/lists/{listId}/tasks/{taskId}` | Same params + `status`. All optional. Uses `getTaskInfo()` for dual ID resolution. |
| `complete_task` | (wraps update_task) | Convenience: sets `status: 'completed'` + `completedDateTime`. Param: `task_id`. |
| `create_task_list` | `POST /me/todo/lists` | Param: `display_name`. Returns list numeric ID. |

**Destructive (two-phase approval):**

| Prepare Tool | Confirm Tool | Endpoint | Description |
|-------------|-------------|----------|-------------|
| `prepare_delete_task` | `confirm_delete_task` | `DELETE /me/todo/lists/{listId}/tasks/{taskId}` | Prepare shows title, status, due date, list name. Hash: `SHA256(taskId:title:listId)`. |

### New idCache Bucket

```typescript
taskLists: Map<number, string>  // numeric ID -> Graph task list ID
```

## Cross-Cutting Concerns

### Cache Invalidation
All mutations call `this.cache.clear()` — consistent with existing write operations.

### idCache Updates
All create operations add the new entity to the appropriate idCache bucket so subsequent operations can reference them by numeric ID.

### Error Handling
Follow existing patterns: custom error types, `wrapError()`, `{ code, message }` JSON with `isError: true`.

### Scopes
Already configured: `Mail.ReadWrite`, `Calendars.ReadWrite`, `Contacts.ReadWrite`, `Tasks.ReadWrite`. No auth changes needed.
