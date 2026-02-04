# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/jbctechsolutions/mcp-outlook-mac/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/jbctechsolutions/mcp-outlook-mac/releases/tag/v1.0.0
