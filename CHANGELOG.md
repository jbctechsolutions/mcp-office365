# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
- **Graph API Backend (Beta)**
  - Event management write operations not yet implemented (create, update, delete, RSVP)
  - Email sending not yet implemented
  - Read operations are fully functional and stable

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

[Unreleased]: https://github.com/jbctechsolutions/mcp-outlook-mac/compare/v1.1.0...HEAD
[1.1.0]: https://github.com/jbctechsolutions/mcp-outlook-mac/compare/v1.0.3...v1.1.0
[1.0.3]: https://github.com/jbctechsolutions/mcp-outlook-mac/compare/v1.0.2...v1.0.3
[1.0.2]: https://github.com/jbctechsolutions/mcp-outlook-mac/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/jbctechsolutions/mcp-outlook-mac/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/jbctechsolutions/mcp-outlook-mac/releases/tag/v1.0.0
