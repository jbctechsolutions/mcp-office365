# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [v2.5.0] - 2026-03-09

### Added
- **Comprehensive README rewrite** with full tool reference for all tools across 27 categories
- **4 Planner Visualization tools** with 4 output formats (HTML, SVG, Markdown, Mermaid):
  - `generate_kanban_board` — Kanban view of tasks grouped by bucket with priority colors
  - `generate_gantt_chart` — Gantt timeline with start/due dates
  - `generate_plan_summary` — Overview stats, assignee workload, overdue items
  - `generate_burndown_chart` — Burndown/burnup of completed vs remaining tasks
- **6 Meeting Recordings & Transcripts tools** (Teams):
  - `list_online_meetings` / `get_online_meeting` — browse Teams meetings
  - `list_meeting_recordings` / `download_meeting_recording` — access recordings
  - `list_meeting_transcripts` / `get_meeting_transcript_content` — access transcripts (VTT/text)
- **4 Message Reaction tools** (Teams):
  - `list_message_reactions` — list reactions on channel/chat messages
  - `prepare_add_message_reaction` / `confirm_add_message_reaction` — two-phase reaction add
  - `remove_message_reaction` — remove own reaction
- **11 OneDrive tools**:
  - `list_drive_items` / `search_drive_items` / `get_drive_item` — browse and search files
  - `download_file` — download to local filesystem
  - `prepare_upload_file` / `confirm_upload_file` — two-phase upload
  - `list_recent_files` / `list_shared_with_me` — quick access to recent and shared files
  - `create_sharing_link` — create view/edit sharing links
  - `prepare_delete_drive_item` / `confirm_delete_drive_item` — two-phase delete
- **6 SharePoint tools**:
  - `list_sites` / `search_sites` / `get_site` — browse SharePoint sites
  - `list_document_libraries` / `list_library_items` — navigate document libraries
  - `download_library_file` — download SharePoint files
- **6 Excel Online tools**:
  - `list_worksheets` / `get_worksheet_range` / `get_used_range` — read spreadsheet data
  - `prepare_update_range` / `confirm_update_range` — two-phase cell updates
  - `get_table_data` — read named table rows
- **Graph $batch infrastructure** for future performance optimization

### Changed
- GitHub repository renamed from `mcp-outlook-mac` to `mcp-office365-mac`

## [v2.4.0] - 2026-03-07

### Added
- **17 Microsoft Planner tools** (Phase 5) bringing the total to ~189 tools
- **Plans** (4 tools):
  - `list_plans` / `get_plan` — browse planner plans
  - `create_plan` / `update_plan` — create and modify plans (with ETag concurrency)
- **Buckets** (5 tools):
  - `list_buckets` / `create_bucket` / `update_bucket` — manage plan buckets
  - `prepare_delete_bucket` / `confirm_delete_bucket` — two-phase bucket deletion
- **Planner Tasks** (6 tools):
  - `list_planner_tasks` / `get_planner_task` — browse tasks with assignments, priority, dates
  - `create_planner_task` / `update_planner_task` — create/update with bucket, assignments, priority
  - `prepare_delete_planner_task` / `confirm_delete_planner_task` — two-phase task deletion
- **Task Details** (2 tools):
  - `get_planner_task_details` / `update_planner_task_details` — rich description, checklist, and references
- **ETag caching** for all Planner write operations (plans, buckets, tasks, task details) with automatic `If-Match` header management

## [v2.3.0] - 2026-03-07

### Added
- **8 People API tools** (Phase 4) bringing the total to ~172 tools
- `list_relevant_people` — AI-ranked relevant contacts based on communication patterns
- `search_people` — search people by name or email
- `get_manager` / `get_direct_reports` — organizational chart navigation
- `get_user_profile` — detailed user profile by email or ID
- `get_user_photo` — download user profile photo to disk
- `get_user_presence` / `get_users_presence` — real-time presence status (single and batch)

## [v2.2.0] - 2026-03-07

### Added
- **13 Microsoft To Do extended tools + 2 extended** (Phase 3) bringing the total to ~164 tools
- **Checklist Items** (5 tools):
  - `list_checklist_items` / `create_checklist_item` / `update_checklist_item` — manage subtasks on To Do tasks
  - `prepare_delete_checklist_item` / `confirm_delete_checklist_item` — two-phase deletion
- **Linked Resources** (4 tools):
  - `list_linked_resources` / `create_linked_resource` — attach web links to tasks
  - `prepare_delete_linked_resource` / `confirm_delete_linked_resource` — two-phase deletion
- **Task Attachments** (4 tools):
  - `list_task_attachments` / `create_task_attachment` — attach files to tasks (base64)
  - `prepare_delete_task_attachment` / `confirm_delete_task_attachment` — two-phase deletion
- **Categories on Tasks**: `create_task` and `update_task` now accept optional `categories: string[]` parameter

## [v2.1.0] - 2026-03-07

### Added
- **20 Microsoft Teams tools** (Phase 2) bringing the total to ~151 tools
- **Teams & Channels** (8 tools):
  - `list_teams` — list all joined teams
  - `list_channels` / `get_channel` — browse and inspect channels
  - `create_channel` / `update_channel` — create and modify channels
  - `prepare_delete_channel` / `confirm_delete_channel` — two-phase channel deletion
  - `list_team_members` — list team membership with roles
- **Channel Messages** (6 tools):
  - `list_channel_messages` — list recent messages in a channel (with pagination)
  - `get_channel_message` — get a message with all its replies
  - `prepare_send_channel_message` / `confirm_send_channel_message` — two-phase message sending
  - `prepare_reply_channel_message` / `confirm_reply_channel_message` — two-phase reply sending
- **Chats** (6 tools):
  - `list_chats` — list recent 1:1, group, and meeting chats
  - `get_chat` — get chat details with web URL
  - `list_chat_messages` — list recent messages in a chat
  - `prepare_send_chat_message` / `confirm_send_chat_message` — two-phase chat message sending
  - `list_chat_members` — list chat participants with roles

## [v2.0.0] - 2026-03-07

### Breaking Changes
- **Package renamed** from `@jbctechsolutions/mcp-outlook-mac` to `@jbctechsolutions/mcp-office365-mac`
- **Binary renamed** from `mcp-outlook-mac` to `mcp-office365-mac`
- **OAuth scopes expanded** to include Teams, People, Planner, and Presence permissions (users will need to re-consent)

### Added
- **115 tools** (up from 74) across mail, calendar, contacts, tasks, and new M365 domains
- **13 new feature tools** (v1.5.0 -> v1.6.0 equivalent):
  - `set_email_importance` -- set email priority (low/normal/high)
  - `add_draft_attachment` / `add_draft_inline_image` -- add attachments to existing drafts
  - `get_emails` -- batch fetch up to 25 emails by ID
  - `list_conversation` -- thread/conversation view
  - `search_emails_advanced` -- KQL search with from:, subject:, hasAttachments:, date ranges
  - `check_new_emails` -- delta sync for incremental email polling
  - `list_mail_rules` / `create_mail_rule` / `prepare_delete_mail_rule` / `confirm_delete_mail_rule` -- inbox rule management
  - `list_task_lists` / `rename_task_list` / `prepare_delete_task_list` / `confirm_delete_task_list` -- task list management
  - `list_contact_folders` / `create_contact_folder` / `prepare_delete_contact_folder` / `confirm_delete_contact_folder` -- contact folder management
  - `get_contact_photo` / `set_contact_photo` -- contact photo management
  - `list_event_instances` -- recurring event instance management
  - Task recurrence support on `create_task` and `update_task`
- **10 Outlook gap features** (Phase 1):
  - `get_automatic_replies` / `set_automatic_replies` -- out-of-office settings
  - `get_mailbox_settings` / `update_mailbox_settings` -- timezone, language, date/time format
  - `list_categories` / `create_category` / `prepare_delete_category` / `confirm_delete_category` -- master category management
  - `list_focused_overrides` / `create_focused_override` / `prepare_delete_focused_override` / `confirm_delete_focused_override` -- focused inbox sender classification
  - `get_mail_tips` -- pre-send recipient checks (OOF, mailbox full, delivery restrictions)
  - `get_message_headers` / `get_message_mime` -- email header inspection and raw MIME export
  - `list_calendar_groups` / `create_calendar_group` -- calendar group organization
  - `list_calendar_permissions` / `create_calendar_permission` / `prepare_delete_calendar_permission` / `confirm_delete_calendar_permission` -- calendar sharing
  - `list_room_lists` / `list_rooms` -- meeting room discovery
  - Online meeting support on `create_event` and `update_event` (Teams/Skype links)
- **M365 OAuth scopes** for Teams, People, Planner, and Presence (preparing for Phase 2+)

## [v1.5.0] - 2026-02-26

### Added
- **`body_file` parameter** on `create_draft` and `prepare_send_email`: path to a file containing the email body so the server reads it from disk instead of receiving it in the MCP payload. Avoids transport size limits for large HTML (e.g. with embedded images). Either `body` or `body_file` is required.
- **`inline_images` on `create_draft`**: array of `{ file_path, content_id }` to attach images as inline parts; reference in HTML via `<img src="cid:content_id">` to avoid embedding base64 in the JSON body.
- **`uploadInlineAttachment()`** in Graph attachments helper: uploads a file as an inline attachment with `isInline` and `contentId`, MIME type from extension, 3MB max per image.

## [v1.3.0] - 2026-02-24

### Added
- **Graph API onboarding flow**: First-time users are now guided through authentication automatically
  - **Inline auth on first tool call**: When unauthenticated, the server triggers Microsoft's device code flow automatically instead of returning an error
  - **CLI `auth` subcommand**: `npx @jbctechsolutions/mcp-office365-mac auth` for standalone pre-authentication, `--status` to check auth state, `--logout` to sign out
  - **Auth mutex**: Concurrent tool calls during authentication safely coalesce into a single auth flow
- **Reply/forward as draft tools**: 2 new tools (72 → 74 total)
  - `reply_as_draft` — Create a reply (or reply-all) as an editable draft
  - `forward_as_draft` — Create a forward as an editable draft
  - Both return a `draft_id` for use with existing `update_draft` and `prepare_send_draft` tools

## [v1.2.1] - 2026-02-24

### Changed
- **Documentation overhaul**: Updated README to reflect v1.2.0 capabilities — 72 tools, full Graph API write operations, removed outdated "read-only" and "beta" references
- Updated Graph API permissions documentation (`Contacts.ReadWrite`, `Tasks.ReadWrite`)
- Updated security model documentation to reflect read/write access with two-phase approval
- Fixed CHANGELOG comparison links for all versions

## [v1.2.0] - 2026-02-23

### Added
- **Graph API write operations**: 34 new tools (38 → 72 total) across 5 categories:
  - **Mail send/draft tools**: Create, update, and send drafts; reply, reply-all, and forward messages with two-phase approval for send operations
  - **Attachment tools**: List, download, and upload attachments with automatic integration into draft creation and email sending workflows
  - **Calendar write tools**: Create, update, and delete events; respond to invitations (accept/decline/tentative) with two-phase approval for destructive operations
  - **Contact write tools**: Create, update, and delete contacts with full field mapping (name, email, phone, address) and two-phase delete approval
  - **Task write tools**: Create, update, complete, and delete tasks; create task lists with two-phase delete approval
- All write operations use the existing approval token pattern with TOCTOU hash protection for destructive operations
- 700+ new tests (1338 total) with Graph API test coverage improved to 88%

## [v1.1.8] - 2026-02-23

### Fixed
- **Graph API event/task times off by timezone offset**: `dateTimeTimeZoneToTimestamp()` ignored the `timeZone` field from Graph API `DateTimeTimeZone` objects. DateTime strings without a `Z` suffix (e.g. `"2026-02-23T16:00:00.0000000"`) were parsed as local time instead of UTC, shifting events by the local timezone offset (e.g. 5 hours in EST). Now appends `Z` when `timeZone` is `"UTC"`.

## [v1.1.7] - 2026-02-23

### Fixed
- **Graph API email/task timestamps**: Fixed `timeReceived`, `timeSent`, `dueDate`, and `startDate` showing dates 31 years in the future (2057 instead of 2026) due to Apple epoch offset being applied to Graph API Unix timestamps.
- **All Graph API timestamps now in local timezone**: Replaced UTC (`Z` suffix) with local timezone offset (e.g. `-05:00`) across events, emails, and tasks for human-readable dates.

### Added
- `unixTimestampToLocalIso()` utility function that converts Unix timestamps to ISO 8601 strings with local timezone offset.

## [v1.1.6] - 2026-02-23

### Fixed
- **Graph API event titles**: Events now display correct titles from `event.subject` instead of `null`. Added `subject` field to `EventRow` interface and populated it in all backends (Graph mapper, AppleScript adapter, SQL queries).
- **Graph API event dates**: Events now show correct dates (2026) instead of dates 31 years in the future (2057). Created `transformGraphEventRow()` with `unixTimestampToIso()` to avoid applying the Apple epoch offset to Graph API Unix timestamps.

### Added
- `unixTimestampToIso()` utility function for converting Unix timestamps to ISO strings without Apple epoch adjustment.
- 7 new unit tests covering event subject mapping and Unix timestamp conversion.

## [v1.1.5] - 2026-02-23

### Added
- **Graph API contract tests**: Added 46 tests verifying every API call uses correct URLs, HTTP methods, request bodies, query parameters, well-known folder names, and OData operators against the Microsoft Graph v1.0 documentation.

## [v1.1.4] - 2026-02-23

### Fixed
- **Graph API Event queries**: Fixed invalid `isRecurrence` property in `$select` strings causing `GRAPH_ERROR` when listing or fetching events. Changed to correct property name `recurrence`.

### Added
- **Graph API field validation tests**: Added 19 contract tests verifying all `$select` fields across every Graph API endpoint match valid Microsoft Graph v1.0 property names.

## [1.1.0] - 2026-02-08

### Added
- **Attachment Listing**: New `list_attachments` tool to retrieve attachment metadata (name, size, content type) for any email
- **Attachment Downloading**: New `download_attachment` tool to save email attachments to local disk
- **Inline Image Support**: `send_email` now supports `inline_images` parameter for embedding images in HTML email bodies via content IDs
- **Attachment Metadata in Emails**: `get_email` now includes `attachments` array with full metadata when the email has attachments
- **Attachment Error Handling**: New `AttachmentTooLargeError` and `AttachmentSaveError` error classes with 25 MB size limit enforcement

### Changed
- Enhanced AppleScript `getMessage` to return detailed attachment metadata (index, name, size, content type)
- Extended `Email` interface with optional `attachments: AttachmentInfo[]` field
- Updated `send_email` tool schema to accept `inline_images` array

## [1.0.3] - 2026-02-06

### Fixed
- **npx/Symlink Execution**: Fixed server not responding when run via npx or npm bin symlinks
  - Improved `isMainModule` check to handle symlinks, npx, and various execution contexts
  - Server now correctly initializes in all execution environments

## [1.0.2] - 2026-02-06

### Changed
- **CI/CD Improvements**
  - Enable automatic npm publishing via trusted publisher (OIDC)
  - Remove token-based authentication in favor of provenance-based publishing
  - Fix macOS compatibility in changelog extraction script

## [1.0.1] - 2026-02-06

### Fixed
- **npx Installation**: Fixed bin file not being executable when installed via npx
  - Added `chmod +x dist/index.js` to build script
  - Added executable verification step to CI/CD workflows
  - TypeScript compiler doesn't preserve executable permissions, now explicitly set after build

## [1.0.0] - 2026-02-03

### Added
- **Dual Backend Architecture**
  - AppleScript backend for classic Outlook for Mac
  - Microsoft Graph API backend for new Outlook

- **Mail Tools**
  - List mail folders
  - List, search, and get emails
  - Unread count by folder
  - Send email (AppleScript backend only)

- **Calendar Tools**
  - List calendars
  - List, search, and get events
  - Create, update, and delete events (AppleScript backend only)
  - RSVP to meeting invitations (AppleScript backend only)

- **Contact Tools**
  - List contacts
  - Search contacts
  - Get contact details

- **Task Tools**
  - List tasks
  - Search tasks
  - Get task details

- **Note Tools** (AppleScript backend only)
  - List notes
  - Search notes
  - Get note content

- **Account Management**
  - List all configured accounts
  - Get account details

- **Advanced Features**
  - Mailbox organization with two-phase approval system
  - Comprehensive error handling (24 specialized error classes)
  - Device code flow authentication for Graph API
  - Token caching and automatic refresh (`~/.outlook-mcp/tokens.json`)
  - Multi-account support

- **Development**
  - Comprehensive test suite (37 test files, 80% coverage threshold)
  - Strict TypeScript and ESLint configuration
  - Unit, integration, and E2E tests

### Known Limitations
- **Graph API Backend**
  - Notes not available (Microsoft Graph API limitation)
  - All other read/write operations fully implemented as of v1.2.0

- **Platform Limitations**
  - AppleScript backend does not support Google accounts (macOS limitation)
  - Graph API backend does not support Notes (Microsoft Graph limitation)
  - Requires macOS for AppleScript backend

- **Authentication**
  - Graph API requires Azure AD app registration
  - Shared public client ID provided for quick-start
  - Users can create custom Azure AD apps for production use

### Security
- Secure token storage with restrictive file permissions (0o600)
- Environment variable configuration for sensitive data
- No hardcoded credentials
- Two-phase approval system for destructive operations

---

[Unreleased]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v2.4.0...HEAD
[v2.5.0]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v2.4.0...v2.5.0
[v2.4.0]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v2.3.0...v2.4.0
[v2.3.0]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v2.2.0...v2.3.0
[v2.2.0]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v2.1.0...v2.2.0
[v2.1.0]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v2.0.0...v2.1.0
[v2.0.0]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v1.5.0...v2.0.0
[v1.3.0]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v1.2.1...v1.3.0
[v1.2.1]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v1.2.0...v1.2.1
[v1.2.0]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v1.1.8...v1.2.0
[v1.1.8]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v1.1.7...v1.1.8
[v1.1.7]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v1.1.6...v1.1.7
[v1.1.6]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v1.1.5...v1.1.6
[v1.1.5]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v1.1.4...v1.1.5
[v1.1.4]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v1.1.0...v1.1.4
[1.1.0]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v1.0.3...v1.1.0
[1.0.3]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v1.0.2...v1.0.3
[1.0.2]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/jbctechsolutions/mcp-office365-mac/releases/tag/v1.0.0
