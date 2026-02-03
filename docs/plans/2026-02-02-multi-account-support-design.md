# Multi-Account Support Design

**Date:** 2026-02-02
**Status:** Complete - With Known Limitations

## Requirements

Add support for querying multiple Outlook accounts in the MCP server.

### User Requirements
- List all available Outlook accounts
- Query specific account(s) or all accounts
- Maintain backward compatibility with existing single-account behavior

## Known Limitations

### Google Accounts Not Supported via AppleScript

**Issue:** Google accounts configured in Outlook for Mac are not accessible via AppleScript.

**Technical Details:**
- AppleScript exposes account types: `exchange account`, `imap account`, `pop account`
- Google accounts use a proprietary integration with OAuth authentication
- They are not exposed through any standard AppleScript account class
- Tested class names: `google account`, `gmail account`, `online account`, `cloud account` - none exist in the AppleScript dictionary

**Impact:** The multi-account feature only detects Exchange, IMAP, and POP accounts. Google accounts configured in Outlook will not appear in `list_accounts` results and cannot be queried.

**Workaround:** Users with Google accounts can use the Graph API backend (`USE_GRAPH_API=1`) which has different account handling, or access Google mail through IMAP configuration instead of the native Google account integration.

### Parameter Design
- `account_id?: number | number[] | "all"`
  - Omitted → query default account (backward compatible)
  - `account_id: 1` → query specific account
  - `account_id: [1, 2]` → query multiple specific accounts
  - `account_id: "all"` → query all accounts

### Response Format
- **Single account query:** Current format (backward compatible)
- **Multiple account query:** Grouped by account
  ```json
  {
    "accounts": [
      {
        "account_id": 1,
        "account_name": "Work",
        "emails": [...]
      }
    ]
  }
  ```

## Architecture

### Layer 1: Account Management
- **AccountRepository** - Queries Outlook for account information via AppleScript
  - `getDefaultAccount()` - Returns Outlook's default account
  - `listAccounts()` - Returns all configured accounts

### Layer 2: Account Resolution
- **AccountResolver** - Interprets `account_id` parameter
  - Input: `account_id?: number | number[] | "all"`
  - Output: Array of account IDs to query

### Layer 3: Data Repositories (Enhanced)
- Existing repositories accept `accountIds: number[]` parameter
- AppleScript queries filter by account

### Layer 4: Response Transformation
- **ResponseTransformer** - Formats responses
  - Single account → current format
  - Multiple accounts → grouped format

### Layer 5: MCP Tools (Enhanced)
- Add `account_id` to input schemas
- Use AccountResolver and ResponseTransformer
- Minimal changes to core logic

## AppleScript Implementation

### List Accounts
```applescript
tell application "Microsoft Outlook"
    set accountList to every exchange account
    set accountData to {}
    repeat with acc in accountList
        set accountInfo to {id: id of acc, name: name of acc, email: email address of acc}
        set end of accountData to accountInfo
    end repeat
    return accountData
end tell
```

### Get Default Account
```applescript
tell application "Microsoft Outlook"
    try
        set defaultAcc to default account
        return {id: id of defaultAcc}
    on error
        set firstAcc to first exchange account
        return {id: id of firstAcc}
    end try
end tell
```

### Enhanced Queries with Account Filter
Existing queries will iterate over specified account IDs and collect results.

## Implementation Status

- [x] Implement AccountRepository with AppleScript (`account-repository.ts`)
- [x] Create AppleScript templates for accounts (`account-scripts.ts`)
- [x] Add `list_accounts` MCP tool
- [x] Add `account_id` parameter to `list_folders` tool
- [x] Implement response transformation for grouped results
- [x] Test with real Outlook data
- [x] Document known limitations (Google accounts not supported)
