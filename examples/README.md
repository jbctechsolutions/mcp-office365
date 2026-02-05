# Example Configurations

This directory contains example MCP server configurations for different use cases.

## Quick Links
- [AppleScript Backend (Classic Outlook)](#applescript-backend)
- [Graph API Backend (Quick Start)](#graph-api-quick-start)
- [Graph API Backend (Custom App)](#graph-api-custom-app)
- [Usage Examples](#usage-examples)

## AppleScript Backend

**File:** `claude-desktop-applescript.json`

For classic Outlook for Mac using local SQLite database.

**Requirements:**
- macOS
- Classic Outlook for Mac installed
- Outlook running
- Automation permissions granted

**Copy to:** `~/Library/Application Support/Claude/claude_desktop_config.json`

## Graph API Quick Start

**File:** `claude-desktop-graph-shared.json`

Uses shared Azure AD app for quick testing.

**Requirements:**
- Microsoft account (personal, work, or school)
- Internet connection
- Device code authentication on first run

**Copy to:** `~/Library/Application Support/Claude/claude_desktop_config.json`

## Graph API Custom App

**File:** `claude-desktop-graph-custom.json`

Uses your own Azure AD app registration.

**Requirements:**
- Azure AD app registration (see [README](../README.md#custom-azure-ad-setup))
- Your application client ID
- Microsoft account

**Setup:**
1. Create Azure AD app (see main README)
2. Replace `your-azure-app-client-id` with your actual client ID
3. Copy to `~/Library/Application Support/Claude/claude_desktop_config.json`

## Usage Examples

**File:** `usage-examples.md`

Common prompts and usage patterns for the Outlook MCP server.

---

## Troubleshooting

### Config File Not Loading
- Restart Claude Desktop after editing config
- Check JSON syntax (use a JSON validator)
- Check file path is correct

### AppleScript Backend Errors
- Ensure Outlook is running
- Grant automation permissions: System Settings > Privacy & Security > Automation
- Check Outlook profile name matches config

### Graph API Authentication Fails
- Check internet connection
- Complete device code flow within time limit
- Verify client ID is correct (if using custom app)
- Check tenant ID setting

For more help, see [Troubleshooting](../README.md#troubleshooting) in main README.
