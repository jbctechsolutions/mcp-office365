# @jbctechsolutions/mcp-office365

[![npm version](https://badge.fury.io/js/%40jbctechsolutions%2Fmcp-office365.svg)](https://www.npmjs.com/package/@jbctechsolutions/mcp-office365)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js Version](https://img.shields.io/node/v/@jbctechsolutions/mcp-office365)](https://nodejs.org)

MCP server for Microsoft 365 -- mail, calendar, contacts, tasks, teams, people, and planner.

A [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server that provides **181 tools** for full read/write access to Microsoft 365. Manage your emails, calendar events, contacts, tasks, notes, Teams channels and chats, people directory, and Planner boards directly through MCP.

## Features Overview

| Category | Tools | Description |
|----------|------:|-------------|
| Mail -- Reading | 9 | Folders, search, delta sync, conversations, unread counts |
| Mail -- Sending & Drafts | 16 | Send, draft, reply, forward with two-phase approval |
| Mail -- Signatures | 2 | Email signature management |
| Mail -- Organization | 24 | Read/unread, flags, categories, importance, move, delete, batch ops |
| Mail -- Rules | 4 | Inbox rule management |
| Mail -- Categories | 4 | Master category management |
| Mail -- Focused Inbox | 4 | Focused inbox override management |
| Mail -- Settings & Auto-Replies | 4 | Automatic replies, mailbox settings |
| Mail -- Tips & Headers | 3 | Mail tips, message headers, MIME export |
| Attachments | 2 | List and download email attachments |
| Calendar -- Events | 11 | List, search, create, update, delete, RSVP, recurring instances |
| Calendar -- Groups | 2 | Calendar group management |
| Calendar -- Permissions | 4 | Calendar sharing permissions |
| Calendar -- Rooms | 2 | Room lists and meeting rooms |
| Contacts & Folders | 13 | CRUD contacts, contact folders, photos |
| Tasks & Task Lists | 13 | To Do tasks with recurrence, task lists |
| Checklist Items | 5 | Subtasks on To Do tasks |
| Linked Resources | 4 | Linked resources on To Do tasks |
| Task Attachments | 4 | File attachments on To Do tasks |
| Notes (AppleScript only) | 3 | List, read, and search Outlook notes |
| Scheduling | 2 | Free/busy availability, meeting time suggestions |
| Teams -- Channels | 8 | Channel CRUD, team members |
| Teams -- Channel Messages | 6 | Read and send channel messages with replies |
| Teams -- Chats | 6 | 1:1 and group chats, send messages |
| People & Presence | 8 | People search, org chart, presence status |
| Planner | 17 | Plans, buckets, tasks, task details with ETag |
| Accounts | 1 | List configured Exchange accounts |
| **Total** | **181** | |

## Quick Start

### Install and run

```bash
npx -y @jbctechsolutions/mcp-office365
```

By default the server uses the **Microsoft Graph API** backend (cross-platform, full read/write access).

To use the **AppleScript backend** (classic Outlook for Mac only, limited features), set the environment variable:

```bash
USE_APPLESCRIPT=1
```

### Pre-authenticate (optional)

```bash
npx @jbctechsolutions/mcp-office365 auth
npx @jbctechsolutions/mcp-office365 auth --status
npx @jbctechsolutions/mcp-office365 auth --logout
```

### Claude Desktop configuration

Add to `~/Library/Application Support/Claude/claude_desktop_config.json`:

**Graph API backend (default):**
```json
{
  "mcpServers": {
    "office365": {
      "command": "npx",
      "args": ["-y", "@jbctechsolutions/mcp-office365"]
    }
  }
}
```

**AppleScript backend** (macOS + classic Outlook only):
```json
{
  "mcpServers": {
    "office365": {
      "command": "npx",
      "args": ["-y", "@jbctechsolutions/mcp-office365"],
      "env": {
        "USE_APPLESCRIPT": "1"
      }
    }
  }
}
```

### Claude Code configuration

**Option 1: Plugin (recommended)** -- add to `~/.claude/settings.json`:

```json
{
  "extraKnownMarketplaces": {
    "jbctechsolutions": {
      "source": {
        "source": "github",
        "repo": "jbctechsolutions/mcp-office365"
      }
    }
  },
  "enabledPlugins": {
    "office365@jbctechsolutions": true
  }
}
```

**Option 2: Manual** -- add `.mcp.json` to your project or `~/.claude/.mcp.json` globally:

```json
{
  "mcpServers": {
    "office365": {
      "command": "npx",
      "args": ["-y", "@jbctechsolutions/mcp-office365"]
    }
  }
}
```

## Authentication

### Device Code Flow (Graph API)

1. The server displays a device code and URL
2. Visit https://microsoft.com/devicelogin
3. Enter the code and sign in with your Microsoft account
4. Grant the requested permissions

Tokens are stored in `~/.outlook-mcp/tokens.json` and refreshed automatically.

### Custom Azure AD App Registration

For production, work accounts with conditional access, or full control over the app lifecycle:

1. **Register** in [Azure Portal](https://portal.azure.com) > Azure Active Directory > App registrations > New registration
   - Name: `Outlook MCP Server`
   - Supported account types: Multitenant + personal accounts
   - Redirect URI: leave blank
2. **Add API permissions** (Microsoft Graph > Delegated): see [Required Graph API Permissions](#required-graph-api-permissions) below
3. **Enable public client flows**: Authentication > Advanced settings > Allow public client flows > Yes
4. **Configure** with environment variables:

```json
{
  "mcpServers": {
    "office365": {
      "command": "npx",
      "args": ["-y", "@jbctechsolutions/mcp-office365"],
      "env": {
        "OUTLOOK_MCP_CLIENT_ID": "your-client-id-here",
        "OUTLOOK_MCP_TENANT_ID": "common"
      }
    }
  }
}
```

**Tenant options:** `common` (all accounts), `organizations` (work/school only), `consumers` (personal only), or a specific tenant ID.

## Tool Reference

All 181 tools listed below. Tools marked *(Graph API)* require `USE_GRAPH_API=1`. Tools marked *(AppleScript only)* are not available with Graph API.

<details>
<summary><strong>Accounts (1)</strong></summary>

| Tool | Description |
|------|-------------|
| `list_accounts` | List all Exchange accounts configured in Outlook |

</details>

<details>
<summary><strong>Mail -- Reading (9)</strong></summary>

| Tool | Description |
|------|-------------|
| `list_folders` | List all mail folders with message and unread counts |
| `list_emails` | List emails in a folder with pagination |
| `search_emails` | Search emails by subject, sender, or content |
| `search_emails_advanced` | Advanced email search using KQL (Keyword Query Language) *(Graph API)* |
| `check_new_emails` | Check for new/changed emails since last check via delta sync *(Graph API)* |
| `get_email` | Get full email details including body |
| `get_emails` | Get multiple emails by ID in a single call (max 25) |
| `list_conversation` | List all messages in an email conversation/thread *(Graph API)* |
| `get_unread_count` | Get unread email count |

</details>

<details>
<summary><strong>Mail -- Sending & Drafts (16)</strong></summary>

| Tool | Description |
|------|-------------|
| `send_email` | Send an email with optional CC, BCC, attachments, and HTML |
| `create_draft` | Create a draft email for later editing and sending |
| `update_draft` | Update an existing draft email |
| `add_draft_attachment` | Add a file attachment to a draft *(Graph API)* |
| `add_draft_inline_image` | Add an inline image to a draft for HTML body *(Graph API)* |
| `list_drafts` | List all draft emails |
| `prepare_send_draft` | Prepare to send a draft (two-phase approval) |
| `confirm_send_draft` | Confirm and send a draft |
| `prepare_send_email` | Prepare to send an email immediately (two-phase) |
| `confirm_send_email` | Confirm and send the email |
| `prepare_reply_email` | Prepare to reply to an email (two-phase) |
| `confirm_reply_email` | Confirm and send the reply |
| `prepare_forward_email` | Prepare to forward an email (two-phase) |
| `confirm_forward_email` | Confirm and forward the email |
| `reply_as_draft` | Create a reply (or reply-all) as an editable draft |
| `forward_as_draft` | Create a forward as an editable draft |

</details>

<details>
<summary><strong>Mail -- Organization (24)</strong></summary>

| Tool | Description |
|------|-------------|
| `mark_email_read` | Mark an email as read |
| `mark_email_unread` | Mark an email as unread |
| `set_email_flag` | Set a follow-up flag on an email |
| `clear_email_flag` | Clear the follow-up flag from an email |
| `set_email_categories` | Set categories on an email |
| `set_email_importance` | Set email importance level (low, normal, high) *(Graph API)* |
| `create_folder` | Create a new mail folder |
| `rename_folder` | Rename a mail folder |
| `move_folder` | Move a mail folder under a different parent |
| `prepare_delete_email` | Prepare to delete an email (two-phase) |
| `confirm_delete_email` | Confirm email deletion |
| `prepare_move_email` | Prepare to move an email to another folder (two-phase) |
| `confirm_move_email` | Confirm email move |
| `prepare_archive_email` | Prepare to archive an email (two-phase) |
| `confirm_archive_email` | Confirm email archive |
| `prepare_junk_email` | Prepare to mark an email as junk (two-phase) |
| `confirm_junk_email` | Confirm marking as junk |
| `prepare_delete_folder` | Prepare to delete a mail folder (two-phase) |
| `confirm_delete_folder` | Confirm folder deletion |
| `prepare_empty_folder` | Prepare to empty a folder (two-phase) |
| `confirm_empty_folder` | Confirm emptying folder |
| `prepare_batch_delete_emails` | Prepare to batch delete multiple emails (two-phase) |
| `prepare_batch_move_emails` | Prepare to batch move multiple emails (two-phase) |
| `confirm_batch_operation` | Confirm a batch delete or move operation |

</details>

<details>
<summary><strong>Mail -- Rules (4)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `list_mail_rules` | List all inbox mail rules |
| `create_mail_rule` | Create an inbox rule with conditions and actions |
| `prepare_delete_mail_rule` | Prepare to delete a mail rule (two-phase) |
| `confirm_delete_mail_rule` | Confirm mail rule deletion |

</details>

<details>
<summary><strong>Mail -- Categories (4)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `list_categories` | List all master categories |
| `create_category` | Create a new master category with a color preset |
| `prepare_delete_category` | Prepare to delete a master category (two-phase) |
| `confirm_delete_category` | Confirm category deletion |

</details>

<details>
<summary><strong>Mail -- Focused Inbox (4)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `list_focused_overrides` | List all focused inbox overrides |
| `create_focused_override` | Create a focused inbox override for a sender |
| `prepare_delete_focused_override` | Prepare to delete a focused inbox override (two-phase) |
| `confirm_delete_focused_override` | Confirm focused inbox override deletion |

</details>

<details>
<summary><strong>Mail -- Settings & Auto-Replies (4)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `get_automatic_replies` | Get automatic replies (out-of-office) settings |
| `set_automatic_replies` | Set automatic replies (out-of-office) settings |
| `get_mailbox_settings` | Get mailbox settings (language, time zone, formats, working hours) |
| `update_mailbox_settings` | Update mailbox settings (language, time zone, formats) |

</details>

<details>
<summary><strong>Mail -- Tips & Headers (3)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `get_mail_tips` | Get mail tips (auto-replies, mailbox full, restrictions) for addresses |
| `get_message_headers` | Get internet message headers (SPF, DKIM, routing) |
| `get_message_mime` | Download the full MIME content (.eml) of an email |

</details>

<details>
<summary><strong>Attachments (2)</strong></summary>

| Tool | Description |
|------|-------------|
| `list_attachments` | List attachment metadata (name, size, type) for an email |
| `download_attachment` | Download an email attachment to disk |

</details>

<details>
<summary><strong>Calendar -- Events (11)</strong></summary>

| Tool | Description |
|------|-------------|
| `list_calendars` | List all calendar folders |
| `list_events` | List calendar events with date range filtering |
| `get_event` | Get event details |
| `search_events` | Search events by title and/or date range |
| `create_event` | Create a calendar event (supports Teams online meetings) |
| `update_event` | Update event details (single instance or series) |
| `respond_to_event` | Accept, decline, or tentatively accept an invitation |
| `delete_event` | Delete a calendar event or recurring series |
| `prepare_delete_event` | Prepare to delete a calendar event (two-phase) *(Graph API)* |
| `confirm_delete_event` | Confirm calendar event deletion |
| `list_event_instances` | List instances of a recurring event in a date range *(Graph API)* |

</details>

<details>
<summary><strong>Calendar -- Groups (2)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `list_calendar_groups` | List all calendar groups |
| `create_calendar_group` | Create a new calendar group |

</details>

<details>
<summary><strong>Calendar -- Permissions (4)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `list_calendar_permissions` | List sharing permissions for a calendar |
| `create_calendar_permission` | Share a calendar by creating a permission |
| `prepare_delete_calendar_permission` | Prepare to delete a calendar permission (two-phase) |
| `confirm_delete_calendar_permission` | Confirm calendar permission deletion |

</details>

<details>
<summary><strong>Calendar -- Rooms (2)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `list_room_lists` | List all room lists (building/floor groupings) |
| `list_rooms` | List meeting rooms, optionally filtered by room list |

</details>

<details>
<summary><strong>Contacts & Folders (13)</strong></summary>

| Tool | Description |
|------|-------------|
| `list_contacts` | List contacts with pagination |
| `search_contacts` | Search contacts by name |
| `get_contact` | Get contact details |
| `create_contact` | Create a new contact *(Graph API)* |
| `update_contact` | Update contact details *(Graph API)* |
| `prepare_delete_contact` | Prepare to delete a contact (two-phase) *(Graph API)* |
| `confirm_delete_contact` | Confirm contact deletion |
| `get_contact_photo` | Download a contact's photo *(Graph API)* |
| `set_contact_photo` | Set or update a contact's photo *(Graph API)* |
| `list_contact_folders` | List all contact folders *(Graph API)* |
| `create_contact_folder` | Create a contact folder *(Graph API)* |
| `prepare_delete_contact_folder` | Prepare to delete a contact folder (two-phase) *(Graph API)* |
| `confirm_delete_contact_folder` | Confirm contact folder deletion |

</details>

<details>
<summary><strong>Tasks & Task Lists (13)</strong></summary>

| Tool | Description |
|------|-------------|
| `list_task_lists` | List all task lists (Microsoft To Do) *(Graph API)* |
| `list_tasks` | List tasks with pagination and filtering |
| `search_tasks` | Search tasks by name |
| `get_task` | Get task details |
| `create_task` | Create a task with optional recurrence *(Graph API)* |
| `update_task` | Update task details with optional recurrence *(Graph API)* |
| `complete_task` | Mark a task as completed *(Graph API)* |
| `create_task_list` | Create a new task list *(Graph API)* |
| `rename_task_list` | Rename a task list *(Graph API)* |
| `prepare_delete_task` | Prepare to delete a task (two-phase) *(Graph API)* |
| `confirm_delete_task` | Confirm task deletion |
| `prepare_delete_task_list` | Prepare to delete a task list (two-phase) *(Graph API)* |
| `confirm_delete_task_list` | Confirm task list deletion |

</details>

<details>
<summary><strong>Checklist Items (5)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `list_checklist_items` | List checklist items (subtasks) on a To Do task |
| `create_checklist_item` | Create a checklist item on a To Do task |
| `update_checklist_item` | Update a checklist item (toggle check, rename) |
| `prepare_delete_checklist_item` | Prepare to delete a checklist item (two-phase) |
| `confirm_delete_checklist_item` | Confirm checklist item deletion |

</details>

<details>
<summary><strong>Linked Resources (4)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `list_linked_resources` | List linked resources on a To Do task |
| `create_linked_resource` | Create a linked resource on a To Do task |
| `prepare_delete_linked_resource` | Prepare to delete a linked resource (two-phase) |
| `confirm_delete_linked_resource` | Confirm linked resource deletion |

</details>

<details>
<summary><strong>Task Attachments (4)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `list_task_attachments` | List attachments on a To Do task |
| `create_task_attachment` | Attach a file to a To Do task (base64 encoded) |
| `prepare_delete_task_attachment` | Prepare to delete a task attachment (two-phase) |
| `confirm_delete_task_attachment` | Confirm task attachment deletion |

</details>

<details>
<summary><strong>Notes (3)</strong> <em>(AppleScript only)</em></summary>

| Tool | Description |
|------|-------------|
| `list_notes` | List notes with pagination |
| `get_note` | Get note details |
| `search_notes` | Search notes by content |

> Notes are only available with the AppleScript backend. Microsoft Graph API does not provide access to Outlook Notes.

</details>

<details>
<summary><strong>Scheduling (2)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `check_availability` | Check free/busy availability for people in a time window |
| `find_meeting_times` | Find available meeting time slots for a group of attendees |

</details>

<details>
<summary><strong>Mail -- Signatures (2)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `set_signature` | Save an email signature auto-appended to outgoing emails |
| `get_signature` | Get the currently stored email signature |

</details>

<details>
<summary><strong>Teams -- Channels (8)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `list_teams` | List all Microsoft Teams the user has joined |
| `list_channels` | List all channels in a team |
| `get_channel` | Get details for a specific channel |
| `create_channel` | Create a new channel in a team |
| `update_channel` | Update a channel name or description |
| `prepare_delete_channel` | Prepare to delete a channel (two-phase) |
| `confirm_delete_channel` | Confirm channel deletion |
| `list_team_members` | List all members of a team |

</details>

<details>
<summary><strong>Teams -- Channel Messages (6)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `list_channel_messages` | List recent messages in a channel |
| `get_channel_message` | Get a specific channel message with its replies |
| `prepare_send_channel_message` | Prepare to send a message to a channel (two-phase) |
| `confirm_send_channel_message` | Confirm sending a channel message |
| `prepare_reply_channel_message` | Prepare to reply to a channel message (two-phase) |
| `confirm_reply_channel_message` | Confirm replying to a channel message |

</details>

<details>
<summary><strong>Teams -- Chats (6)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `list_chats` | List recent 1:1 and group chats |
| `get_chat` | Get details of a specific chat |
| `list_chat_messages` | List recent messages in a chat |
| `prepare_send_chat_message` | Prepare to send a message in a chat (two-phase) |
| `confirm_send_chat_message` | Confirm sending a chat message |
| `list_chat_members` | List members of a chat |

</details>

<details>
<summary><strong>People & Presence (8)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `list_relevant_people` | List AI-ranked relevant people for the current user |
| `search_people` | Search people by name or email |
| `get_manager` | Get the current user's manager |
| `get_direct_reports` | Get the current user's direct reports |
| `get_user_profile` | Get a user's profile information |
| `get_user_photo` | Get a user's profile photo |
| `get_user_presence` | Get a user's presence/availability status |
| `get_users_presence` | Batch get presence status for multiple users |

</details>

<details>
<summary><strong>Planner (17)</strong> <em>(Graph API)</em></summary>

| Tool | Description |
|------|-------------|
| `list_plans` | List all Planner plans the user has access to |
| `get_plan` | Get details for a specific Planner plan |
| `create_plan` | Create a new Planner plan in a Microsoft 365 group |
| `update_plan` | Update a Planner plan title |
| `list_buckets` | List all buckets in a Planner plan |
| `create_bucket` | Create a new bucket in a Planner plan |
| `update_bucket` | Update a Planner bucket name |
| `prepare_delete_bucket` | Prepare to delete a Planner bucket (two-phase) |
| `confirm_delete_bucket` | Confirm Planner bucket deletion |
| `list_planner_tasks` | List all tasks in a Planner plan |
| `get_planner_task` | Get details for a specific Planner task |
| `create_planner_task` | Create a new task in a Planner plan |
| `update_planner_task` | Update a Planner task |
| `prepare_delete_planner_task` | Prepare to delete a Planner task (two-phase) |
| `confirm_delete_planner_task` | Confirm Planner task deletion |
| `get_planner_task_details` | Get task details (description, checklist, references) |
| `update_planner_task_details` | Update task details (requires ETag from get_planner_task_details) |

</details>

## Architecture

### Dual Backend

The server supports two backends:

- **Microsoft Graph API (default)** -- connects to Microsoft 365 cloud services. Full read/write across all 181 tools. No Outlook installation required. Works on macOS, Windows, and Linux.
- **AppleScript** (`USE_APPLESCRIPT=1`) -- communicates with classic Outlook for Mac via `osascript`. Works offline, no Microsoft account needed. Limited to reading mail, calendar, contacts, tasks, and notes, plus calendar write operations and email sending.

### Two-Phase Approval

Destructive operations (delete, send, move, forward, etc.) use a prepare/confirm pattern. The `prepare_*` call returns a preview and a short-lived approval token. The `confirm_*` call executes the action only with a valid token. This prevents accidental data loss.

### ID Caching

The server maintains an internal ID mapping layer. AppleScript and Graph API use different ID formats; the caching layer assigns stable numeric IDs so tool callers do not need to track backend-specific identifiers.

### ETag Caching

Planner resources use ETag-based concurrency control. The server caches ETags from read operations and automatically includes them in update/delete requests to prevent conflicts.

### Source Layout

```
src/
  applescript/       AppleScript integration (legacy backend)
  graph/             Microsoft Graph API integration (default)
    auth/            MSAL authentication, device code flow, token cache
    client/          Graph client wrapper with response caching
    mappers/         Graph-to-internal type mappers
  tools/             MCP tool implementations (mail-send, mailbox-organization, etc.)
  types/             TypeScript type definitions
  utils/             Date parsing, error handling, utilities
```

## Required Graph API Permissions

These delegated permissions are requested when using the Graph API backend:

| Permission | Purpose |
|------------|---------|
| `Mail.ReadWrite` | Read, send, and manage mail |
| `Calendars.ReadWrite` | Read and manage calendar events |
| `Contacts.ReadWrite` | Read and manage contacts |
| `Tasks.ReadWrite` | Read and manage To Do tasks |
| `User.Read` | Read user profile |
| `offline_access` | Token refresh (maintain access) |
| `ChannelMessage.Read.All` | Read Teams channel messages |
| `ChannelMessage.Send` | Send Teams channel messages |
| `Channel.ReadBasic.All` | Read Teams channel metadata |
| `Team.ReadBasic.All` | Read Teams metadata |
| `Chat.ReadWrite` | Read and send Teams chat messages |
| `ChatMessage.Send` | Send Teams chat messages |
| `People.Read` | Read relevant people |
| `User.ReadBasic.All` | Read basic user profiles and photos |
| `Presence.Read.All` | Read user presence/availability |
| `Group.Read.All` | Read Microsoft 365 groups (for Planner) |

## Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `USE_APPLESCRIPT` | Set to `1` or `true` for legacy AppleScript backend | (unset -- uses Graph API) |
| `OUTLOOK_MCP_CLIENT_ID` | Override the embedded Azure AD client ID | (embedded) |
| `OUTLOOK_MCP_TENANT_ID` | Azure AD tenant ID | `common` |

## Known Limitations

**AppleScript backend:**
- Google accounts in Outlook are not accessible (macOS/Outlook limitation). Use IMAP configuration or the Graph API backend instead.
- Write operations limited to calendar events and email sending. All other writes require Graph API.

**Graph API backend:**
- Outlook Notes are not available (Graph API does not expose them). Use AppleScript backend for notes.

## Contributing

We welcome contributions. Please see our [Contributing Guide](CONTRIBUTING.md) for details.

**Support policy:** This is a part-time, hobby project. Response times are best-effort (typically 1-2 weeks for bug reports). See [SUPPORT.md](SUPPORT.md) for details.

This project adheres to our [Code of Conduct](CODE_OF_CONDUCT.md).

## Support This Project

- Star this repository
- [GitHub Sponsors](https://github.com/sponsors/jbctechsolutions)
- [Buy Me a Coffee](https://buymeacoffee.com/jbctechsolutions)
- [PayPal](https://paypal.me/jbctechsolutions)

## License

MIT License -- Copyright (c) 2026 JBC Tech Solutions, LLC. See [LICENSE](LICENSE) for details.
