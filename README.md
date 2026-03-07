# Office 365 MCP Server

[![npm version](https://badge.fury.io/js/mcp-office365-mac.svg)](https://www.npmjs.com/package/mcp-office365-mac)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js Version](https://img.shields.io/node/v/mcp-office365-mac)](https://nodejs.org)

A [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server that provides full access to Microsoft 365. Read, write, and manage your emails, calendar events, contacts, tasks, and notes directly through MCP tools.

## Features

- **107 tools** - Full read/write access to mail, calendar, contacts, and tasks
- **Two backends** - AppleScript for classic Outlook, Microsoft Graph API for new Outlook
- **Two-phase approval** - Destructive operations (delete, send) require explicit confirmation
- **Works offline** - AppleScript backend requires no network (Graph API requires internet)
- **Fast and reliable** - Direct communication with Outlook or Microsoft servers

### Available Tools (78)

**Accounts (1)**
- `list_accounts` - List all configured Outlook accounts

**Mail - Reading (9)**
- `list_folders` - List all mail folders with unread counts
- `list_emails` - List emails in a folder with pagination
- `search_emails` - Search emails by subject, sender, or content
- `search_emails_advanced` - Advanced email search using KQL (Keyword Query Language) *(Graph API)*
- `check_new_emails` - Check for new/changed emails since last check (delta sync) *(Graph API)*
- `get_email` - Get full email details including body
- `get_emails` - Get multiple emails by ID in a single call (max 25)
- `list_conversation` - List all messages in an email conversation/thread
- `get_unread_count` - Get unread email count

**Mail - Sending & Drafts (16)** *(Graph API)*
- `send_email` - Send an email with attachments and HTML support
- `create_draft` - Create a new draft email
- `update_draft` - Update an existing draft
- `add_draft_attachment` - Add a file attachment to an existing draft *(Graph API)*
- `add_draft_inline_image` - Add an inline image to an existing draft *(Graph API)*
- `list_drafts` - List all draft emails
- `prepare_send_draft` / `confirm_send_draft` - Send a draft (two-phase)
- `prepare_send_email` / `confirm_send_email` - Compose and send (two-phase)
- `prepare_reply_email` / `confirm_reply_email` - Reply to a message (two-phase)
- `prepare_forward_email` / `confirm_forward_email` - Forward a message (two-phase)
- `reply_as_draft` - Create a reply (or reply-all) as an editable draft
- `forward_as_draft` - Create a forward as an editable draft

**Attachments (2)**
- `list_attachments` - List attachment metadata for an email
- `download_attachment` - Download an email attachment to disk

**Mailbox Organization (24)** *(Graph API)*
- `mark_email_read` / `mark_email_unread` - Toggle read status
- `set_email_flag` / `clear_email_flag` - Flag/unflag emails
- `set_email_categories` - Categorize emails
- `set_email_importance` - Set email importance/priority level (low, normal, high) *(Graph API)*
- `create_folder` / `rename_folder` / `move_folder` - Folder management
- `prepare_delete_email` / `confirm_delete_email` - Delete email (two-phase)
- `prepare_move_email` / `confirm_move_email` - Move email (two-phase)
- `prepare_archive_email` / `confirm_archive_email` - Archive email (two-phase)
- `prepare_junk_email` / `confirm_junk_email` - Mark as junk (two-phase)
- `prepare_delete_folder` / `confirm_delete_folder` - Delete folder (two-phase)
- `prepare_empty_folder` / `confirm_empty_folder` - Empty folder (two-phase)
- `prepare_batch_delete_emails` / `prepare_batch_move_emails` / `confirm_batch_operation` - Batch operations (two-phase)

**Mail Rules (4)** *(Graph API)*
- `list_mail_rules` - List all inbox mail rules
- `create_mail_rule` - Create an inbox mail rule with conditions and actions
- `prepare_delete_mail_rule` / `confirm_delete_mail_rule` - Delete a mail rule (two-phase)

**Master Categories (4)** *(Graph API)*
- `list_categories` - List all master categories
- `create_category` - Create a new master category with a color preset
- `prepare_delete_category` / `confirm_delete_category` - Delete a master category (two-phase)

**Automatic Replies (2)** *(Graph API)*
- `get_automatic_replies` - Get the current automatic replies (out-of-office) settings
- `set_automatic_replies` - Set automatic replies (out-of-office) settings

**Mailbox Settings (2)** *(Graph API)*
- `get_mailbox_settings` - Get the current mailbox settings (language, time zone, date/time formats, working hours)
- `update_mailbox_settings` - Update mailbox settings (language, time zone, date/time formats)

**Mail Tips (1)** *(Graph API)*
- `get_mail_tips` - Get mail tips (automatic replies, mailbox full, delivery restrictions, max message size) for email addresses

**Calendar - Reading (5)**
- `list_calendars` - List all calendars
- `list_events` - List events with date range filtering
- `get_event` - Get event details
- `search_events` - Search events by title
- `list_event_instances` - List instances of a recurring event within a date range *(Graph API)*

**Calendar - Writing (6)**
- `create_event` - Create a new calendar event
- `update_event` - Update event details (title, time, location, etc.); also works on instance IDs from `list_event_instances`
- `respond_to_event` - Accept, decline, or tentatively accept invitations
- `delete_event` - Delete a calendar event or recurring series; also works on instance IDs from `list_event_instances`
- `prepare_delete_event` / `confirm_delete_event` - Delete event with two-phase approval *(Graph API)*

**Contacts (9)**
- `list_contacts` - List all contacts with pagination
- `search_contacts` - Search contacts by name
- `get_contact` - Get contact details
- `create_contact` - Create a new contact *(Graph API)*
- `update_contact` - Update contact details *(Graph API)*
- `prepare_delete_contact` / `confirm_delete_contact` - Delete contact (two-phase) *(Graph API)*
- `get_contact_photo` - Download a contact's photo *(Graph API)*
- `set_contact_photo` - Set or update a contact's photo *(Graph API)*

**Contact Folders (4)**
- `list_contact_folders` - List all contact folders *(Graph API)*
- `create_contact_folder` - Create a contact folder *(Graph API)*
- `prepare_delete_contact_folder` / `confirm_delete_contact_folder` - Delete contact folder (two-phase) *(Graph API)*

**Tasks (13)**
- `list_task_lists` - List all task lists (Microsoft To Do) *(Graph API)*
- `list_tasks` - List tasks with completion filtering
- `get_task` - Get task details
- `search_tasks` - Search tasks by name
- `create_task` - Create a new task with optional recurrence *(Graph API)*
- `update_task` - Update task details with optional recurrence *(Graph API)*
- `complete_task` - Mark a task as complete *(Graph API)*
- `create_task_list` - Create a new task list *(Graph API)*
- `rename_task_list` - Rename a task list *(Graph API)*
- `prepare_delete_task_list` / `confirm_delete_task_list` - Delete task list (two-phase) *(Graph API)*
- `prepare_delete_task` / `confirm_delete_task` - Delete task (two-phase) *(Graph API)*

**Notes (3)** *(AppleScript only)*
- `list_notes` - List all notes
- `get_note` - Get note details
- `search_notes` - Search notes by content

> **Note**: Notes are only supported with the AppleScript backend. Microsoft Graph API does not provide access to Outlook Notes.

## 🚀 Quick Start

### Option 1: AppleScript Backend (Classic Outlook)
```bash
npx -y mcp-office365-mac
```
Requires: Classic Outlook for Mac running

### Option 2: Graph API Backend (New Outlook)
```bash
# Uses shared Azure AD app - works out of the box
npx -y mcp-office365-mac
```
Set `USE_GRAPH_API=1` in your MCP configuration.

**Pre-authenticate (optional):**
```bash
npx @jbctechsolutions/mcp-office365-mac auth
```

Or just configure the server — it will prompt for authentication on first use.

**For production or work accounts:** See [Custom Azure AD Setup](#custom-azure-ad-setup) below.

## Known Limitations

### AppleScript Backend

**Google Accounts Not Supported**

Google accounts configured in Outlook for Mac cannot be accessed via the AppleScript backend. This is a macOS/Outlook limitation - Google accounts use a proprietary OAuth integration that isn't exposed through AppleScript.

**Supported account types:**
- Exchange accounts
- IMAP accounts
- POP accounts

**Not supported:**
- Google accounts (native integration)

**Workarounds:**
1. Configure Google as an IMAP account instead of using the native Google integration
2. Use the Graph API backend (`USE_GRAPH_API=1`) which has different account handling

**Write Operations**

The AppleScript backend supports calendar event management (create, update, delete, RSVP) and email sending. All other write operations (drafts, mailbox organization, contacts, tasks) are only available via the Graph API backend.

### Microsoft Graph API Backend

**Notes Not Available**

Microsoft Graph API does not provide access to Outlook Notes. If you need access to notes, use the AppleScript backend.

## Backends

### AppleScript (Default)

The default backend uses AppleScript to communicate with Microsoft Outlook for Mac. This works best with classic Outlook and requires Outlook to be running.

### Microsoft Graph API

For "new Outlook" for Mac (cloud-based), use the Microsoft Graph API backend. This connects directly to Microsoft's servers and doesn't require Outlook to be running.

To enable the Graph API backend, set the environment variable:

```bash
USE_GRAPH_API=1
```

#### First-Time Authentication

When using the Graph API backend for the first time, you'll need to authenticate:

1. The server will display a device code and URL
2. Visit https://microsoft.com/devicelogin
3. Enter the code displayed in the terminal
4. Sign in with your Microsoft account
5. Grant the requested permissions

Your authentication tokens are stored securely in `~/.outlook-mcp/tokens.json` and will be refreshed automatically.

#### Required Permissions

The Graph API backend requests these Microsoft Graph permissions:
- `Mail.ReadWrite` - Read, send, and manage your mail
- `Calendars.ReadWrite` - Read and manage your calendars
- `Contacts.ReadWrite` - Read and manage your contacts
- `Tasks.ReadWrite` - Read and manage your tasks
- `User.Read` - Read your profile
- `offline_access` - Maintain access (for token refresh)

#### Security Model - Shared Azure AD App

This project provides a shared Azure AD application for quick-start convenience. **Here's what you should know:**

##### ✅ What the Shared App CAN Access

- **Only data you explicitly consent to** during the device code authentication flow
- **Only when you're actively using** the MCP server
- **Tokens are stored locally** on your machine (`~/.outlook-mcp/tokens.json`)
- **Read/write access** to mail, calendar, contacts, and tasks (with two-phase approval for destructive operations)

##### ❌ What the Shared App CANNOT Access

- **Your data when you're not using the server** - tokens are only used by your local MCP instance
- **Your password or credentials** - Microsoft handles authentication
- **Other users' data** - each user authenticates separately with their own account

##### 🔒 How It Works (Technical Details)

```
Your Device → Shared Azure App ID → Microsoft Authentication
                                           ↓
                                    Your Microsoft Account
                                           ↓
                                    Access Token (stored locally)
                                           ↓
                                    Microsoft Graph API → Your Data
```

**Key Security Points:**
- The Azure AD **client ID is public** (not a secret) - it's just an identifier
- **You authenticate with your own Microsoft account** - the app owner never sees your credentials
- **Access tokens are issued to you** and stored on your machine - the app owner cannot access them
- **Delegated permissions** mean the app can only act on your behalf when you're using it
- **Open source code** - you can audit exactly what the server does with your data

##### 🏢 For Production or Corporate Use

**We recommend creating your own Azure AD app if:**
- You're using a work/school account with conditional access policies
- Your organization requires internal app registrations only
- You want full control over the app lifecycle
- You need audit logs under your tenant

See [Custom Azure AD Setup](#custom-azure-ad-setup) below for instructions.

##### 🤝 Trust & Transparency

- ✅ **Open Source** - Full code available at [GitHub](https://github.com/jbctechsolutions/mcp-office365-mac)
- ✅ **Minimal Scopes** - Only requests necessary permissions
- ✅ **Standard Practice** - Same model used by Postman, Microsoft Graph Explorer, and many open-source tools
- ✅ **User Control** - You can revoke access anytime in your [Microsoft account settings](https://account.microsoft.com/privacy/app-access)
- ✅ **Override Option** - Use `OUTLOOK_MCP_CLIENT_ID` environment variable to use your own app

##### ⚠️ Shared App Considerations

**Potential limitations:**
- If the shared app is revoked or deleted, you'll need to use your own app
- All users share the app's rate limits (10,000 requests per 10 minutes - sufficient for typical use)
- Corporate policies may block external multi-tenant apps

**Risk to app owner (JBC Tech Solutions):**
- Microsoft could revoke the app if abuse is detected
- No access to your data or liability for your usage

**Cost:** Using the shared app is **free for everyone** - no charges to you or the app owner.

### Custom Azure AD Setup

The server includes a pre-configured shared Azure AD app for quick-start testing. For production use, custom deployments, or work/school accounts with conditional access, create your own:

#### 1. Register Azure AD Application

1. Go to [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. **Name:** `Outlook MCP Server`
3. **Supported account types:** "Accounts in any organizational directory and personal Microsoft accounts (Multitenant)"
4. **Redirect URI:** Leave blank
5. Click **Register**
6. Note the **Application (client) ID**

#### 2. Configure Permissions

1. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
2. Add these permissions:
   - `Mail.ReadWrite` - Read, send, and manage mail
   - `Calendars.ReadWrite` - Manage calendar events
   - `Contacts.ReadWrite` - Manage contacts
   - `Tasks.ReadWrite` - Manage tasks
   - `User.Read` - User profile
   - `offline_access` - Token refresh
3. Click **Add permissions**

#### 3. Enable Public Client Flow

1. Go to **Authentication**
2. **Advanced settings** → **Allow public client flows** → **Yes**
3. Click **Save**

#### 4. Configure Environment

```json
{
  "mcpServers": {
    "outlook-mac": {
      "command": "npx",
      "args": ["-y", "mcp-office365-mac"],
      "env": {
        "USE_GRAPH_API": "1",
        "OUTLOOK_MCP_CLIENT_ID": "your-client-id-here",
        "OUTLOOK_MCP_TENANT_ID": "common"
      }
    }
  }
}
```

#### Tenant Options
- `common` - Multi-tenant (personal, work, school accounts)
- `organizations` - Work/school accounts only
- `consumers` - Personal Microsoft accounts only
- `{tenant-id}` - Specific Azure AD tenant

#### Trade-offs

| Approach | Pros | Cons |
|----------|------|------|
| **Shared App** | Zero setup, works immediately | Shared with others, may not work with conditional access |
| **Custom App** | Full control, works with conditional access, audit logs | Requires Azure AD setup |

**Cost:** Both options are **free** - Azure AD Free tier is sufficient.

## Installation

### Using npx (recommended)

```bash
npx -y @jbctechsolutions/mcp-office365-mac
```

### Using npm

```bash
npm install -g @jbctechsolutions/mcp-office365-mac
```

## Configuration

### Claude Desktop

Add to your Claude Desktop configuration (`~/Library/Application Support/Claude/claude_desktop_config.json`):

**AppleScript backend (default):**
```json
{
  "mcpServers": {
    "outlook-mac": {
      "command": "npx",
      "args": ["-y", "@jbctechsolutions/mcp-office365-mac"]
    }
  }
}
```

**Graph API backend (for new Outlook):**
```json
{
  "mcpServers": {
    "outlook-mac": {
      "command": "npx",
      "args": ["-y", "@jbctechsolutions/mcp-office365-mac"],
      "env": {
        "USE_GRAPH_API": "1"
      }
    }
  }
}
```

### Claude Code

#### Option 1: Install as Plugin (Recommended)

Add the plugin marketplace to your `~/.claude/settings.json`:

```json
{
  "extraKnownMarketplaces": {
    "jbctechsolutions": {
      "source": {
        "source": "github",
        "repo": "jbctechsolutions/mcp-office365-mac"
      }
    }
  },
  "enabledPlugins": {
    "outlook-mac@jbctechsolutions": true
  }
}
```

#### Option 2: Manual Configuration

**Project-specific** - Add `.mcp.json` to your project:
```json
{
  "mcpServers": {
    "outlook-mac": {
      "command": "npx",
      "args": ["-y", "@jbctechsolutions/mcp-office365-mac"]
    }
  }
}
```

**Global** - Create `~/.claude/.mcp.json`:
```json
{
  "mcpServers": {
    "outlook-mac": {
      "command": "npx",
      "args": ["-y", "@jbctechsolutions/mcp-office365-mac"]
    }
  }
}
```

**Graph API backend (for new Outlook):**
```json
{
  "mcpServers": {
    "outlook-mac": {
      "command": "npx",
      "args": ["-y", "@jbctechsolutions/mcp-office365-mac"],
      "env": {
        "USE_GRAPH_API": "1"
      }
    }
  }
}
```

## Requirements

### AppleScript Backend
- macOS
- Microsoft Outlook for Mac (must be running when using tools)
- Node.js 18 or later
- Automation permission for Outlook (you'll be prompted on first use)

### Graph API Backend
- macOS, Windows, or Linux (no Outlook installation required)
- Microsoft account (personal or work/school)
- Node.js 18 or later
- Internet connection

## Permissions

### AppleScript Backend

The MCP server communicates with Outlook via AppleScript. On first use, you'll be prompted to grant automation permission. You can also configure this in:

**System Settings > Privacy & Security > Automation**

Make sure your terminal or Claude Desktop is allowed to control Microsoft Outlook.

### Graph API Backend

The Graph API backend requires you to sign in with your Microsoft account and grant permissions to read your mail, calendar, contacts, and tasks. See the [First-Time Authentication](#first-time-authentication) section above.

## Troubleshooting

### AppleScript Backend

#### Outlook not running

The server requires Outlook to be running. Start Microsoft Outlook before using the MCP tools.

#### Permission denied

If you see an automation permission error:

1. Open **System Settings > Privacy & Security > Automation**
2. Find your terminal app or Claude Desktop
3. Enable the toggle for Microsoft Outlook

#### Timeout errors

Large mailboxes may cause timeout errors. Try reducing the `limit` parameter in your queries.

### Graph API Backend

#### Pre-authentication

You can authenticate before configuring the MCP server:

```bash
# Authenticate
npx @jbctechsolutions/mcp-office365-mac auth

# Check status
npx @jbctechsolutions/mcp-office365-mac auth --status

# Sign out
npx @jbctechsolutions/mcp-office365-mac auth --logout
```

#### Authentication required

If you see "Microsoft Graph authentication required", you need to complete the device code flow:

1. Look for the device code in the terminal output
2. Visit https://microsoft.com/devicelogin
3. Enter the code and sign in

#### Rate limited

Microsoft Graph API has rate limits. If you see rate limit errors, wait a few moments before trying again.

#### Permission denied

If you see permission errors, you may need to re-authenticate or your admin may have restricted access. Sign out and sign in again:

```bash
# Delete token cache to force re-authentication
rm ~/.outlook-mcp/tokens.json
```

#### Notes not available

Outlook Notes are not supported by Microsoft Graph API. If you need access to notes, use the AppleScript backend instead.

## How It Works

This MCP server supports two backends:

### AppleScript Backend (Default)

Uses AppleScript to communicate with Microsoft Outlook for Mac:
- Works best with classic Outlook for Mac
- Requires Outlook to be running
- Works offline (no network required)
- Calendar write ops (create, update, delete, RSVP) and email sending
- Full support for Notes

### Graph API Backend

Uses Microsoft Graph API to access your data:
- Works with "new Outlook" for Mac (or any platform - no Outlook installation required)
- Connects directly to Microsoft's servers
- Full read/write operations: mail, drafts, calendar, contacts, tasks, mailbox organization
- Two-phase approval for destructive operations (delete, send)
- Supports personal and work/school accounts
- Does not support Notes (Graph API limitation)

## Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `USE_GRAPH_API` | Set to `1` or `true` to use Microsoft Graph API backend | (unset, uses AppleScript) |
| `OUTLOOK_MCP_CLIENT_ID` | Override the embedded Azure AD client ID | (embedded) |
| `OUTLOOK_MCP_TENANT_ID` | Azure AD tenant ID (`common` for multi-tenant) | `common` |

## Development

```bash
# Install dependencies
npm install

# Build
npm run build

# Run tests
npm test

# Run with coverage
npm run test:coverage

# Lint
npm run lint

# Type check
npm run typecheck
```

## Architecture

```
src/
├── applescript/        # AppleScript integration (default backend)
│   ├── executor.ts     # osascript wrapper
│   ├── scripts.ts      # AppleScript templates
│   ├── parser.ts       # Output parsing
│   ├── repository.ts   # IRepository implementation
│   └── content-readers.ts  # Content reader implementations
├── graph/              # Microsoft Graph API integration (optional backend)
│   ├── auth/           # Authentication (MSAL, device code flow)
│   │   ├── config.ts   # Azure AD configuration
│   │   ├── token-cache.ts  # Token persistence
│   │   └── device-code-flow.ts  # Authentication flow
│   ├── client/         # Graph client wrapper
│   │   ├── graph-client.ts  # API client
│   │   └── cache.ts    # Response caching
│   ├── mappers/        # Graph type to row type mappers
│   ├── repository.ts   # IRepository implementation
│   └── content-readers.ts  # Content reader implementations
├── tools/              # MCP tool implementations
│   ├── mail-send.ts   # Draft/send/reply/forward tools
│   └── mailbox-organization.ts  # Move, delete, flag, categorize tools
├── types/              # TypeScript type definitions
└── utils/              # Utilities (dates, errors, etc.)
```

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

**Support Policy:** This is a part-time, hobby project. Response times are best-effort (typically 1-2 weeks for bug reports). See [SUPPORT.md](SUPPORT.md) for details.

### Ways to Contribute
- 🐛 Report bugs
- ✨ Suggest features
- 📝 Improve documentation
- 🔀 Submit pull requests
- 💬 Help others in [Discussions](https://github.com/jbctechsolutions/mcp-office365-mac/discussions)

### Code of Conduct
This project adheres to our [Code of Conduct](CODE_OF_CONDUCT.md).

## 💝 Support This Project

If you find this project valuable, consider supporting its development:

- ⭐ **Star** this repository
- 💰 **Sponsor** via [GitHub Sponsors](https://github.com/sponsors/jbctechsolutions)
- ☕ **Buy Me a Coffee** at [buymeacoffee.com/jbctechsolutions](https://buymeacoffee.com/jbctechsolutions)
- 💵 **Donate** via [PayPal](https://paypal.me/jbctechsolutions)

Your support helps maintain and improve this project! 🙏

## 📄 License

MIT License

Copyright (c) 2026 JBC Tech Solutions, LLC

See [LICENSE](LICENSE) file for details.
