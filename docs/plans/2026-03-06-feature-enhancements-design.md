# Feature Enhancements Design

13 new features across mail, tasks, contacts, and calendar.

## Mail Features

### 1. Conversation/Thread View — `list_conversation`

Fetch all messages in an email thread. Uses the already-mapped `conversationId` field.

- **Graph client:** `listConversationMessages(conversationId: string, limit: number)` — `GET /me/messages?$filter=conversationId eq '{id}'&$orderby=receivedDateTime asc`
- **Repository:** `listConversationAsync(messageId: number, limit: number)` — looks up the message to get its `conversationId`, maps the Graph string conversationId from the numeric hash, then fetches all messages with that ID
- **Tool:** `list_conversation` — input: `{ message_id: number, limit?: number (default 25) }`
- **Returns:** Array of email rows ordered chronologically

### 2. Draft Attachments — `add_draft_attachment` / `add_draft_inline_image`

Add attachments to existing drafts. Reuses `uploadAttachment` and `uploadInlineAttachment` from `src/graph/attachments.ts`.

- **No new Graph client or repository methods** — tools resolve draft Graph ID from cache, call existing upload functions
- **Tools in `mail-send.ts`:**
  - `add_draft_attachment` — input: `{ draft_id: number, file_path: string, name?: string, content_type?: string }`
  - `add_draft_inline_image` — input: `{ draft_id: number, file_path: string, content_id: string }`
- **Returns:** `{ success: true, message: string }`

### 3. Batch Read — `get_emails`

Fetch multiple emails by ID in a single tool call.

- **No new Graph client or repository methods** — uses `Promise.all` over existing `getEmailAsync`
- **Tool:** `get_emails` — input: `{ email_ids: number[] }` (max 25)
- **Returns:** Array of full email details, with nulls for IDs not found

### 4. KQL Advanced Search — `search_emails_advanced`

Expose Microsoft Graph's KQL (Keyword Query Language) for power-user search queries.

- **Graph client:** `searchMessagesKql(query: string, limit: number)` — passes raw query to `$search` without wrapping in quotes. Also `searchMessagesKqlInFolder(folderId, query, limit)`.
- **Repository:** `searchEmailsAdvancedAsync(query: string, limit: number)` and `searchEmailsAdvancedInFolderAsync(folderId, query, limit)`
- **Tool:** `search_emails_advanced` — input: `{ query: string, folder_id?: number, limit?: number }`
- **Tool description documents common KQL operators:** `from:`, `to:`, `subject:`, `hasAttachments:true`, `received>=2024-01-01`, `AND`, `OR`

### 5. Delta Sync — `check_new_emails`

Incremental polling using Graph delta queries. Returns only new/changed messages since last check.

- **Graph client:** `getMessagesDelta(folderId: string, deltaLink?: string): Promise<{ messages: Message[], deltaLink: string }>` — uses `/me/mailFolders/{id}/messages/delta`
- **Repository:** Stores delta links in `Map<number, string>` (keyed by folder numeric ID). `checkNewEmailsAsync(folderId: number)` returns changed messages and updates the stored delta link.
- **Tool:** `check_new_emails` — input: `{ folder_id: number }`
- **First call:** Returns recent messages (initial sync), stores delta link
- **Subsequent calls:** Returns only new/changed messages since last check
- **Returns:** `{ emails: EmailRow[], is_initial_sync: boolean }`

### 6. Email Importance — `set_email_importance`

Set importance/priority on emails.

- **Repository:** `setEmailImportanceAsync(emailId: number, importance: string)` — calls `client.updateMessage(graphId, { importance })`
- **Tool:** `set_email_importance` — input: `{ email_id: number, importance: 'low' | 'normal' | 'high' }`
- **Returns:** `{ success: true, message: string }`

### 7. Mail Rules — `list_mail_rules` / `create_mail_rule` / `prepare_delete_mail_rule` / `confirm_delete_mail_rule`

Manage inbox rules via Graph API.

- **Graph client:**
  - `listMailRules(): Promise<MessageRule[]>` — `GET /me/mailFolders/inbox/messageRules`
  - `createMailRule(rule: Record<string, unknown>): Promise<MessageRule>` — `POST /me/mailFolders/inbox/messageRules`
  - `deleteMailRule(ruleId: string): Promise<void>` — `DELETE /me/mailFolders/inbox/messageRules/{id}`
- **Repository:** `listMailRulesAsync()`, `createMailRuleAsync(params)`, `deleteMailRuleAsync(ruleId: number)` — with ID cache for rules
- **Tools:**
  - `list_mail_rules` — no input, returns all rules
  - `create_mail_rule` — input: `{ display_name, sequence, is_enabled?, conditions, actions }` where conditions/actions follow Graph API schema
  - `prepare_delete_mail_rule` / `confirm_delete_mail_rule` — two-phase delete
- **Conditions** include: `from_addresses`, `subject_contains`, `sender_contains`, `has_attachments`, `importance`, `body_contains`
- **Actions** include: `move_to_folder`, `mark_as_read`, `mark_importance`, `forward_to`, `delete`, `stop_processing_rules`

## Task Features

### 8. List Task Lists — `list_task_lists`

Expose the existing `listTaskLists()` Graph client method as a tool.

- **Repository:** `listTaskListsAsync(): Promise<Array<{ id: number, name: string, isDefault: boolean }>>` — maps and caches task list IDs
- **Tool:** `list_task_lists` — no input
- **Returns:** Array of `{ id, name, is_default }`

### 9. Task Recurrence

Add recurrence support to `create_task` and `update_task`.

- **Schema change:** Add optional `recurrence` field to `CreateTaskInput` and `UpdateTaskInput`:
  ```
  recurrence?: {
    pattern: 'daily' | 'weekly' | 'monthly' | 'yearly',
    interval?: number (default 1),
    days_of_week?: ('monday'|'tuesday'|...)[] (for weekly),
    day_of_month?: number (for monthly),
    range_type: 'endDate' | 'noEnd' | 'numbered',
    start_date: string (ISO date),
    end_date?: string (ISO date, for endDate range),
    occurrences?: number (for numbered range)
  }
  ```
- **Graph API mapping:** Converts to `patternedRecurrence` object with `recurrencePattern` and `recurrenceRange`
- **No new tools** — extends existing create_task/update_task

### 10. Task List Management — `rename_task_list` / `prepare_delete_task_list` / `confirm_delete_task_list`

- **Graph client:** `updateTaskList(listId: string, name: string)`, `deleteTaskList(listId: string)`
- **Repository:** `renameTaskListAsync(listId: number, name: string)`, `deleteTaskListAsync(listId: number)`
- **Tools:**
  - `rename_task_list` — input: `{ task_list_id: number, name: string }`
  - `prepare_delete_task_list` / `confirm_delete_task_list` — two-phase delete, input: `{ task_list_id: number }`

## Contact Features

### 11. Contact Groups — `list_contact_folders` / `create_contact_folder` / `prepare_delete_contact_folder` / `confirm_delete_contact_folder`

Graph API uses "contact folders" as the grouping mechanism.

- **Graph client:** `listContactFolders()`, `createContactFolder(name)`, `deleteContactFolder(folderId)`
- **Repository:** With ID cache for contact folders
- **Tools:**
  - `list_contact_folders` — no input
  - `create_contact_folder` — input: `{ name: string }`
  - `prepare_delete_contact_folder` / `confirm_delete_contact_folder` — two-phase delete
- **Enhance existing `list_contacts`:** Add optional `folder_id` to filter contacts by folder

### 12. Contact Photos — `get_contact_photo` / `set_contact_photo`

- **Graph client:** `getContactPhoto(contactId: string): Promise<ArrayBuffer>`, `setContactPhoto(contactId: string, photoData: Buffer, contentType: string): Promise<void>`
- **Repository:** `getContactPhotoAsync(contactId: number)`, `setContactPhotoAsync(contactId: number, filePath: string)`
- **Tools:**
  - `get_contact_photo` — input: `{ contact_id: number }`, downloads to temp file, returns path
  - `set_contact_photo` — input: `{ contact_id: number, file_path: string }`, uploads from file

## Calendar Features

### 13. Recurring Event Instances — `list_event_instances` / `update_event_instance` / `prepare_delete_event_instance` / `confirm_delete_event_instance`

Manage individual occurrences of recurring events.

- **Graph client:**
  - `listEventInstances(eventId: string, startDateTime: string, endDateTime: string): Promise<Event[]>`
  - Uses `GET /me/events/{id}/instances?startDateTime=...&endDateTime=...`
  - `updateEvent` and `deleteEvent` already work on instance IDs
- **Repository:** `listEventInstancesAsync(eventId: number, startDate: string, endDate: string)` — caches instance IDs
- **Tools:**
  - `list_event_instances` — input: `{ event_id: number, start_date: string, end_date: string }`
  - `update_event_instance` — reuses existing `update_event` pattern with instance ID
  - `prepare_delete_event_instance` / `confirm_delete_event_instance` — two-phase delete

## Decisions

- **Skipped: Multi-account filtering** — Graph API authenticates a single user; multi-account would require managing multiple auth tokens, which is a larger architectural change.
- **Skipped: Webhooks/subscriptions** — MCP servers don't expose HTTP endpoints. Delta sync (feature 5) covers the monitoring use case via polling.
