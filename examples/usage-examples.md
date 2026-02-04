# Usage Examples

Common prompts and tasks you can perform with the Outlook MCP server.

## Mail

### List Recent Emails
> "Show me my 10 most recent emails from my inbox"

### Search Emails
> "Search my emails for messages from john@example.com about the project proposal"

### Unread Count
> "How many unread emails do I have in my inbox?"

### Send Email (AppleScript only)
> "Send an email to jane@example.com with subject 'Meeting Notes' and body 'Thanks for the great meeting today!'"

## Calendar

### Today's Events
> "What's on my calendar today?"

### Search Events
> "Find all calendar events with 'standup' in the title for next week"

### Create Event (AppleScript only)
> "Create a calendar event for tomorrow at 2pm titled 'Team Sync' for 1 hour"

## Contacts

### Search Contacts
> "Find contacts with 'Smith' in their name"

### Get Contact Details
> "Get the full contact information for John Doe"

## Tasks

### List Tasks
> "Show me all my incomplete tasks"

### Search Tasks
> "Find tasks related to the Q1 project"

## Notes (AppleScript only)

### Search Notes
> "Find notes containing 'meeting agenda'"

### Recent Notes
> "Show me my 5 most recent notes"

## Accounts

### List Accounts
> "Which email accounts are configured in Outlook?"

## Advanced

### Multi-Step Workflows
> "Search my calendar for events this week, find the participants, and look up their contact information"

### Email Analysis
> "Analyze my emails from the last 7 days and summarize the main topics discussed"

### Task Planning
> "Show me all tasks due this week and group them by priority"

## Tips

1. **Be specific** - Include date ranges, folder names, or search terms
2. **Combine requests** - The MCP can handle multi-step workflows
3. **Natural language** - Write prompts conversationally
4. **Backend awareness** - Remember which features work with which backend

## Limitations

### Graph API Backend (Beta)
- ❌ Cannot create/update/delete events
- ❌ Cannot send emails
- ❌ Cannot access Notes
- ✅ All read operations work

### AppleScript Backend
- ❌ Requires Outlook to be running
- ❌ Doesn't work with Google accounts
- ✅ Full read/write support

See [Known Limitations](../README.md#known-limitations) in README for details.
