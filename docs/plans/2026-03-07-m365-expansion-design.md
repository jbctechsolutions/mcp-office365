# Microsoft 365 Ecosystem Expansion Design

Expand from Outlook-only to a full M365 MCP server covering Teams, To Do, People, and Planner.

## Decisions

- **Package rename:** `mcp-outlook-mac` -> `mcp-office365-mac`
- **Scope strategy:** Single package, all scopes requested upfront
- **Architecture:** Domain-based modules with thin index.ts orchestrator
- **Migration:** New services get clean module structure; existing Outlook code stays in place initially

## New OAuth Scopes

Added to existing `Mail.ReadWrite`, `Calendars.ReadWrite`, `Contacts.ReadWrite`, `Tasks.ReadWrite`, `User.Read`, `offline_access`:

```
ChannelMessage.Read.All    # read Teams channel messages
ChannelMessage.Send        # send Teams channel messages
Channel.ReadBasic.All      # list channels
Team.ReadBasic.All         # list teams
Chat.ReadWrite             # read/send chat messages
ChatMessage.Send           # send chat messages
People.Read                # relevant people + people search
User.ReadBasic.All         # org chart, user profiles
Presence.Read.All          # user presence status
Group.Read.All             # planner plans (tied to M365 groups)
```

## Directory Structure

```
src/
  index.ts                    # thin orchestrator (~200 lines)
  server.ts                   # MCP server setup, tool registration
  domains/
    mail/
      client.ts               # mail Graph client methods
      repository.ts           # mail repository methods
      tools.ts                # tool definitions + handlers
      schemas.ts              # Zod schemas
    calendar/
    contacts/
    tasks/                    # To Do (existing + extended)
    teams/
    people/
    planner/
  graph/
    client/graph-client.ts    # base client (auth, caching, pagination)
    auth/
  approval/
  applescript/                # unchanged
```

New services get the clean module structure. Existing Outlook domains stay in current files; refactored into `src/domains/` later.

---

## Phase 1: Outlook Gaps (~22 new tools, 2 extended)

All use existing scopes. No new permissions required.

### 1. Automatic Replies (OOF) -- `get_automatic_replies` / `set_automatic_replies`

- `GET /me/mailboxSettings/automaticRepliesSetting`
- `PATCH /me/mailboxSettings` with `automaticRepliesSetting`
- Set internal/external messages, schedule start/end, enable/disable

### 2. Mailbox Settings -- `get_mailbox_settings` / `update_mailbox_settings`

- `GET /me/mailboxSettings`
- `PATCH /me/mailboxSettings`
- Timezone, language, date/time format, working hours

### 3. Master Categories -- `list_categories` / `create_category` / `prepare_delete_category` / `confirm_delete_category`

- `GET /me/outlook/masterCategories`
- `POST /me/outlook/masterCategories`
- `DELETE /me/outlook/masterCategories/{id}`
- Category names with preset colors (`preset0`-`preset24`)

### 4. Focused Inbox -- `list_focused_overrides` / `create_focused_override` / `prepare_delete_focused_override` / `confirm_delete_focused_override`

- `GET /me/inferenceClassification/overrides`
- `POST /me/inferenceClassification/overrides`
- `DELETE /me/inferenceClassification/overrides/{id}`
- Classify senders as `focused` or `other`

### 5. Mail Tips -- `get_mail_tips`

- `POST /me/getMailTips`
- Input: array of email addresses
- Returns: auto-reply status, mailbox full, delivery restrictions, external member count

### 6. Message Headers -- `get_message_headers` / `get_message_mime`

- `GET /me/messages/{id}?$select=internetMessageHeaders` for headers
- `GET /me/messages/{id}/$value` for raw RFC 822 MIME content
- `get_message_mime` saves to file (can be large), returns file path

### 7. Calendar Groups -- `list_calendar_groups` / `create_calendar_group`

- `GET /me/calendarGroups`
- `POST /me/calendarGroups`
- Organize calendars into named groups

### 8. Calendar Permissions -- `list_calendar_permissions` / `create_calendar_permission` / `prepare_delete_calendar_permission` / `confirm_delete_calendar_permission`

- `GET /me/calendars/{id}/calendarPermissions`
- `POST /me/calendars/{id}/calendarPermissions`
- `DELETE /me/calendars/{id}/calendarPermissions/{id}`
- Roles: `read`, `write`, `delegateWithPrivateEventAccess`

### 9. Room Lists -- `list_room_lists` / `list_rooms`

- `GET /me/findRoomLists`
- `GET /me/findRooms` or filtered by room list
- Returns room name, email address, building, floor

### 10. Online Meetings -- extend `create_event` / `update_event`

- Add optional `is_online_meeting: true` + `online_meeting_provider: 'teamsForBusiness'`
- Returns join URL in event response
- No new tools, extends existing schemas

---

## Phase 2: Microsoft Teams (~19 new tools)

### Channels

- `list_teams` -- `GET /me/joinedTeams`
- `list_channels` -- `GET /teams/{id}/channels`
- `get_channel` -- `GET /teams/{id}/channels/{id}`
- `create_channel` -- `POST /teams/{id}/channels`
- `update_channel` -- `PATCH /teams/{id}/channels/{id}`
- `prepare_delete_channel` / `confirm_delete_channel` -- `DELETE /teams/{id}/channels/{id}`

### Channel Messages

- `list_channel_messages` -- `GET /teams/{id}/channels/{id}/messages` (with pagination)
- `get_channel_message` -- `GET /teams/{id}/channels/{id}/messages/{id}` (with replies)
- `prepare_send_channel_message` / `confirm_send_channel_message` -- `POST /teams/{id}/channels/{id}/messages`
- `prepare_reply_channel_message` / `confirm_reply_channel_message` -- `POST /teams/{id}/channels/{id}/messages/{id}/replies`

### Chats (1:1 and group)

- `list_chats` -- `GET /me/chats`
- `get_chat` -- `GET /chats/{id}`
- `list_chat_messages` -- `GET /chats/{id}/messages`
- `prepare_send_chat_message` / `confirm_send_chat_message` -- `POST /chats/{id}/messages`

### Members

- `list_team_members` -- `GET /teams/{id}/members`
- `list_channel_members` -- `GET /teams/{id}/channels/{id}/members`

All send/reply operations use two-phase approval (visible to others, hard to retract).

---

## Phase 3: Microsoft To Do Extended (~12 new tools, 2 extended)

### Checklist Items (subtasks)

- `list_checklist_items` -- `GET /me/todo/lists/{id}/tasks/{id}/checklistItems`
- `create_checklist_item` -- `POST /me/todo/lists/{id}/tasks/{id}/checklistItems`
- `update_checklist_item` -- `PATCH /me/todo/lists/{id}/tasks/{id}/checklistItems/{id}`
- `prepare_delete_checklist_item` / `confirm_delete_checklist_item` -- `DELETE .../{id}`

### Linked Resources

- `list_linked_resources` -- `GET /me/todo/lists/{id}/tasks/{id}/linkedResources`
- `create_linked_resource` -- `POST /me/todo/lists/{id}/tasks/{id}/linkedResources`
- `prepare_delete_linked_resource` / `confirm_delete_linked_resource` -- `DELETE .../{id}`

### Task Attachments

- `list_task_attachments` -- `GET /me/todo/lists/{id}/tasks/{id}/attachments`
- `create_task_attachment` -- `POST /me/todo/lists/{id}/tasks/{id}/attachments`
- `prepare_delete_task_attachment` / `confirm_delete_task_attachment` -- `DELETE .../{id}`

### Categories on Tasks

- Extend existing `create_task` / `update_task` with optional `categories: string[]`
- No new tools

---

## Phase 4: People API (8 new tools)

- `list_relevant_people` -- `GET /me/people` (AI-ranked by communication patterns)
- `search_people` -- `GET /me/people?$search="{query}"`
- `get_manager` -- `GET /me/manager`
- `get_direct_reports` -- `GET /me/directReports`
- `get_user_profile` -- `GET /users/{id}` or `GET /users/{email}`
- `get_user_photo` -- `GET /users/{id}/photo/$value` (saves to file)
- `get_user_presence` -- `GET /users/{id}/presence`
- `get_users_presence` -- `POST /communications/getPresencesByUserId` (batch, max 650)

---

## Phase 5: Microsoft Planner (~14 new tools)

### Plans

- `list_plans` -- `GET /me/planner/plans`
- `get_plan` -- `GET /planner/plans/{id}`
- `create_plan` -- `POST /groups/{groupId}/planner/plans`
- `update_plan` -- `PATCH /planner/plans/{id}`

### Buckets

- `list_buckets` -- `GET /planner/plans/{id}/buckets`
- `create_bucket` -- `POST /planner/plans/{id}/buckets`
- `update_bucket` -- `PATCH /planner/buckets/{id}`
- `prepare_delete_bucket` / `confirm_delete_bucket` -- `DELETE /planner/buckets/{id}`

### Planner Tasks

- `list_planner_tasks` -- `GET /planner/plans/{id}/tasks`
- `get_planner_task` -- `GET /planner/tasks/{id}`
- `create_planner_task` -- `POST /planner/tasks`
- `update_planner_task` -- `PATCH /planner/tasks/{id}` (progress, priority, assignments, bucket, dates)
- `prepare_delete_planner_task` / `confirm_delete_planner_task` -- `DELETE /planner/tasks/{id}`

### Task Details

- `get_planner_task_details` -- `GET /planner/tasks/{id}/details`
- `update_planner_task_details` -- `PATCH /planner/tasks/{id}/details` (description, checklist, references)

Note: Planner uses ETags for concurrency. Every PATCH requires `If-Match` header. Repository caches ETags alongside IDs.

---

## Totals

| Phase | Domain | New Tools | Extended |
|-------|--------|-----------|----------|
| 0 | Package rename + restructure | 0 | -- |
| 1 | Outlook gaps | ~22 | 2 |
| 2 | Teams | ~19 | 0 |
| 3 | To Do extended | ~12 | 2 |
| 4 | People | 8 | 0 |
| 5 | Planner | ~14 | 0 |
| **Total** | | **~75** | **4** |

Current: 94 tools. After expansion: ~169 tools.

---

## Future: Phase 6 -- Planner Visualization

After all phases are complete, add tools to generate visual charts from Planner data:

- **Gantt charts** -- timeline view of planner tasks with dependencies, assignments, and progress
- **Kanban boards** -- bucket-based card view with task status and assignees
- **Burndown/burnup charts** -- progress tracking over time
- **Resource allocation** -- workload distribution across team members

Implementation: Generate SVG or HTML output from Planner task data. Could use a lightweight charting library or template-based SVG generation. Returns file path to the generated chart. This is a post-MVP enhancement -- build after all M365 CRUD tools are stable.
