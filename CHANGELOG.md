# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/jbctechsolutions/mcp-office365-mac/compare/v1.3.0...HEAD
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
