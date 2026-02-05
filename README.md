# Outlook MCP Server

A [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server that provides read-only access to Outlook for Mac. Access your emails, calendar events, contacts, tasks, and notes directly through MCP tools.

## Features

- **Mostly read-only** - Calendar event creation supported; all other data is read-only
- **Two backends** - AppleScript for classic Outlook, Microsoft Graph API for new Outlook
- **Works offline** - AppleScript backend requires no network (Graph API requires internet)
- **Fast and reliable** - Direct communication with Outlook or Microsoft servers

### Available Tools

**Accounts**
- `list_accounts` - List all configured Outlook accounts

**Mail**
- `list_folders` - List all mail folders with unread counts (supports `account_id` filtering)
- `list_emails` - List emails in a folder with pagination
- `search_emails` - Search emails by subject, sender, or content
- `get_email` - Get full email details including body
- `get_unread_count` - Get unread email count
- `send_email` - Send an email with attachments and HTML support (AppleScript backend only)

**Calendar**
- `list_calendars` - List all calendars
- `list_events` - List events with date range filtering
- `get_event` - Get event details
- `search_events` - Search events by title
- `create_event` - Create a new calendar event (AppleScript backend only)
- `respond_to_event` - Accept, decline, or tentatively accept event invitations (AppleScript backend only)
- `delete_event` - Delete a calendar event or recurring series (AppleScript backend only)
- `update_event` - Update event details (title, time, location, etc.) (AppleScript backend only)

**Contacts**
- `list_contacts` - List all contacts with pagination
- `search_contacts` - Search contacts by name
- `get_contact` - Get contact details

**Tasks**
- `list_tasks` - List tasks with completion filtering
- `get_task` - Get task details
- `search_tasks` - Search tasks by name

**Notes**
- `list_notes` - List all notes
- `get_note` - Get note details
- `search_notes` - Search notes by content

> **Note**: Notes are only supported with the AppleScript backend. Microsoft Graph API does not provide access to Outlook Notes.

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

Currently, write operations (event management, email sending) are only supported via the AppleScript backend. These features will be added to the Graph API backend in a future release:
- Event RSVP operations
- Event deletion
- Event updates
- Email sending

For these operations, use the AppleScript backend with classic Outlook for Mac.

### Graph API Backend

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
- `Mail.Read` - Read your mail
- `Calendars.Read` - Read your calendars (will require `Calendars.ReadWrite` when Graph API event creation is added)
- `Contacts.Read` - Read your contacts
- `Tasks.Read` - Read your tasks
- `User.Read` - Read your profile
- `offline_access` - Maintain access (for token refresh)

## Installation

### Using npx (recommended)

```bash
npx -y @jbctechsolutions/mcp-outlook-mac
```

### Using npm

```bash
npm install -g @jbctechsolutions/mcp-outlook-mac
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
      "args": ["-y", "@jbctechsolutions/mcp-outlook-mac"]
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
      "args": ["-y", "@jbctechsolutions/mcp-outlook-mac"],
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
        "repo": "jbctechsolutions/mcp-outlook-mac"
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
      "args": ["-y", "@jbctechsolutions/mcp-outlook-mac"]
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
      "args": ["-y", "@jbctechsolutions/mcp-outlook-mac"]
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
      "args": ["-y", "@jbctechsolutions/mcp-outlook-mac"],
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
- macOS, Windows, or Linux
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
- Full support for Notes

### Graph API Backend

Uses Microsoft Graph API to access your data:
- Works with "new Outlook" for Mac (cloud-based)
- Connects directly to Microsoft's servers
- Works without Outlook running
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
├── types/              # TypeScript type definitions
└── utils/              # Utilities (dates, errors, etc.)
```

## License

MIT
