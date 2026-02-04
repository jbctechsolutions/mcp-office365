# Event Management & Email Sending - Design Document

**Date:** 2026-02-03
**Status:** Approved
**Implementation:** AppleScript backend first, Graph API backend later

## Overview

Add write operations to the Outlook MCP server:
1. **Event Management**: RSVP, delete, and update calendar events
2. **Email Sending**: Send rich emails with attachments

Both features will be implemented for the AppleScript backend first, following the existing dual-backend architecture pattern.

## Requirements

### Event Management (Priority 1)

**RSVP Operations**
- Accept, decline, or tentatively accept meeting invitations
- Send response to organizer (default: true)
- Optional comment to organizer
- Matches Outlook's native RSVP behavior

**Delete Events**
- Remove events from calendar
- Support for recurring events: delete single instance or entire series
- Default: single instance (safer)

**Update Events**
- Full update support: title, time, location, description, recurrence
- All fields optional (partial updates)
- Support for recurring events: update single instance or entire series
- Default: single instance (safer)

**Permission Handling**
- Let Outlook/AppleScript handle permission errors naturally
- No pre-validation of ownership or permissions
- AppleScript will return meaningful errors for unauthorized operations

### Email Sending (Priority 2)

**Rich Email Sending**
- To, CC, BCC recipients
- Subject and body
- HTML or plain text body (default: plain)
- Attachments (file paths)
- Reply-to address
- Multi-account support (select sending account)

**Out of Scope (Future)**
- Reply/Forward to existing threads
- Draft management
- Inline images

## Architecture

### Interface Design

We follow the existing pattern of separate interfaces for write operations:

**Existing:**
- `IRepository` - Read operations (both backends)
- `ICalendarWriter` - Event creation (AppleScript only)

**New:**
- `ICalendarManager` - Event management (RSVP, delete, update)
- `IMailSender` - Email sending

**Benefits:**
- Single responsibility principle
- Independent implementation (AppleScript first, Graph later)
- Consistent with existing codebase patterns
- Clean separation of concerns

### File Structure

```
src/applescript/
├── calendar-writer.ts       # Existing - create events
├── calendar-manager.ts      # NEW - RSVP, delete, update
├── mail-sender.ts           # NEW - send emails
└── scripts.ts               # Extend with new templates

src/tools/
├── calendar.ts              # Extend with new tools
└── email.ts                 # Extend with send_email tool

src/utils/
└── errors.ts                # Extend with new error types
```

## API Design

### Event Management APIs

#### `respond_to_event` Tool

```typescript
{
  event_id: number,
  response: "accept" | "decline" | "tentative",
  send_response?: boolean,  // default: true
  comment?: string          // optional message to organizer
}
```

Returns: `{ success: boolean, message?: string }`

#### `delete_event` Tool

```typescript
{
  event_id: number,
  apply_to?: "this_instance" | "all_in_series"  // default: "this_instance"
}
```

Returns: `{ success: boolean }`

#### `update_event` Tool

```typescript
{
  event_id: number,
  apply_to?: "this_instance" | "all_in_series",  // default: "this_instance"
  title?: string,
  start_date?: string,      // ISO 8601
  end_date?: string,        // ISO 8601
  location?: string,
  description?: string,
  is_all_day?: boolean,
  recurrence?: RecurrenceConfig  // same as create_event
}
```

Returns: `{ success: boolean, updated_fields: string[] }`

### Email Sending API

#### `send_email` Tool

```typescript
{
  to: string[],              // recipient email addresses
  subject: string,
  body: string,
  body_type?: "plain" | "html",  // default: "plain"
  cc?: string[],
  bcc?: string[],
  reply_to?: string,
  attachments?: Array<{
    path: string,            // file path on local system
    name?: string            // optional display name
  }>,
  account_id?: number        // which account to send from (if multiple)
}
```

Returns:
```typescript
{
  message_id: string,        // unique identifier for sent message
  sent_at: string            // ISO 8601 timestamp
}
```

### TypeScript Interfaces

#### ICalendarManager

```typescript
export interface ICalendarManager {
  respondToEvent(
    eventId: number,
    response: 'accept' | 'decline' | 'tentative',
    sendResponse: boolean,
    comment?: string
  ): void;

  deleteEvent(
    eventId: number,
    applyTo: 'this_instance' | 'all_in_series'
  ): void;

  updateEvent(
    eventId: number,
    updates: EventUpdates,
    applyTo: 'this_instance' | 'all_in_series'
  ): UpdatedEvent;
}

export interface EventUpdates {
  readonly title?: string;
  readonly startDate?: string;
  readonly endDate?: string;
  readonly location?: string;
  readonly description?: string;
  readonly isAllDay?: boolean;
  readonly recurrence?: RecurrenceConfig;
}

export interface UpdatedEvent {
  readonly id: number;
  readonly updatedFields: readonly string[];
}
```

#### IMailSender

```typescript
export interface IMailSender {
  sendEmail(params: SendEmailParams): SentEmail;
}

export interface SendEmailParams {
  readonly to: readonly string[];
  readonly subject: string;
  readonly body: string;
  readonly bodyType: 'plain' | 'html';
  readonly cc?: readonly string[];
  readonly bcc?: readonly string[];
  readonly replyTo?: string;
  readonly attachments?: readonly Attachment[];
  readonly accountId?: number;
}

export interface Attachment {
  readonly path: string;
  readonly name?: string;
}

export interface SentEmail {
  readonly messageId: string;
  readonly sentAt: string;
}
```

## Data Flow

```
1. MCP Client → Tool Handler
   User calls: respond_to_event(event_id: 123, response: "accept")

2. Tool Handler → Validation
   Zod schema validates input parameters

3. Tool Handler → ICalendarManager
   calendarManager.respondToEvent(123, "accept", true)

4. CalendarManager → Script Generator
   scripts.respondToEvent({ eventId: 123, response: "accept" })

5. Script Generator → AppleScript Template
   Generates AppleScript code for the operation

6. AppleScript → Executor
   executeAppleScriptOrThrow(script)

7. Executor → osascript
   Runs via child_process.spawn

8. osascript → Outlook
   Executes AppleScript commands in Outlook

9. Outlook → Response
   Returns structured output (JSON or delimited text)

10. Parser → Structured Data
    parseOperationResult(output)

11. Tool Handler → MCP Response
    Returns success/error to client
```

## Error Handling

### Error Types

We leverage the existing error system in `src/utils/errors.ts` and add new types:

**Existing Error Types (Reused):**
- `NotFoundError` - Event or account not found
- `ValidationError` - Invalid input parameters
- `AppleScriptPermissionDeniedError` - Permission denied by Outlook
- `OutlookNotRunningError` - Outlook not running

**New Error Types:**
```typescript
export const ErrorCode = {
  // ... existing codes
  ATTACHMENT_NOT_FOUND: 'ATTACHMENT_NOT_FOUND',
  MAIL_SEND_ERROR: 'MAIL_SEND_ERROR',
  RECURRING_EVENT_ERROR: 'RECURRING_EVENT_ERROR',
} as const;

export class AttachmentNotFoundError extends OutlookMcpError {
  readonly code = ErrorCode.ATTACHMENT_NOT_FOUND;
  constructor(path: string) {
    super(`Attachment file not found: ${path}`);
  }
}

export class MailSendError extends OutlookMcpError {
  readonly code = ErrorCode.MAIL_SEND_ERROR;
  constructor(reason: string) {
    super(`Failed to send email: ${reason}`);
  }
}

export class RecurringEventError extends OutlookMcpError {
  readonly code = ErrorCode.RECURRING_EVENT_ERROR;
  constructor(message: string) {
    super(message);
  }
}
```

### Error Handling Strategy

1. **Input Validation**: Zod schemas catch malformed requests before AppleScript execution
2. **AppleScript Errors**: Caught by `executeAppleScriptOrThrow` and converted to typed errors
3. **Operation-Specific Errors**: Get specific error codes for better debugging
4. **Helpful Messages**: All errors include guidance to solutions

## Testing Strategy

### Unit Tests

**`tests/unit/applescript/calendar-manager.test.ts`**
- Test AppleScript script generation for each operation
- Mock `executeAppleScriptOrThrow` to avoid requiring Outlook
- Verify correct parameters passed to scripts
- Test error handling (parsing failures, invalid inputs)
- Test recurring event handling logic

**`tests/unit/applescript/mail-sender.test.ts`**
- Test email script generation
- Test attachment path handling
- Test HTML vs plain text body
- Test CC/BCC array formatting
- Test multi-account selection

**`tests/unit/applescript/parser.test.ts`** (extend existing)
- Test parsing of operation results
- Test error message extraction
- Test edge cases (empty responses, malformed output)

### Integration Tests

**`tests/integration/tools/calendar.test.ts`** (extend existing)
- Test end-to-end: tool input → AppleScript → parsed output
- Use real Outlook if available, otherwise skip with `test.skip`
- Test RSVP, delete, update operations
- Test recurring event handling (this_instance vs all_in_series)

**`tests/integration/tools/email.test.ts`** (extend existing)
- Test email sending with various configurations
- Test attachment handling
- Test HTML email sending
- Verify sent email appears in Sent folder
- Test multi-account sending

### Test Data

- Create test fixtures for sample events
- Mock AppleScript responses for predictable testing
- Use small test attachments (text files) for email tests

### Coverage Goal

Maintain >85% coverage (matching existing codebase).

## Implementation Phases

### Phase 1: Event Management (AppleScript)

1. **Setup**
   - Create `src/applescript/calendar-manager.ts`
   - Add `ICalendarManager` interface
   - Extend `scripts.ts` with event management templates

2. **RSVP Operations**
   - Implement `respondToEvent` method
   - Add AppleScript template
   - Add parser for RSVP results
   - Add `respond_to_event` MCP tool
   - Write tests

3. **Delete Operations**
   - Implement `deleteEvent` method
   - Add AppleScript template with recurring event support
   - Add parser for delete results
   - Add `delete_event` MCP tool
   - Write tests

4. **Update Operations**
   - Implement `updateEvent` method
   - Add AppleScript template with full field support
   - Add parser for update results
   - Add `update_event` MCP tool
   - Write tests

### Phase 2: Email Sending (AppleScript)

1. **Setup**
   - Create `src/applescript/mail-sender.ts`
   - Add `IMailSender` interface
   - Extend `scripts.ts` with email sending templates

2. **Basic Email Sending**
   - Implement `sendEmail` method
   - Add AppleScript template for basic send
   - Add parser for send results
   - Add `send_email` MCP tool
   - Write tests

3. **Rich Features**
   - Add HTML body support
   - Add CC/BCC support
   - Add attachment support
   - Add reply-to support
   - Add multi-account support
   - Write tests

### Phase 3: Error Handling & Polish

1. **Error Types**
   - Add new error types to `src/utils/errors.ts`
   - Update error handling in all operations
   - Write error handling tests

2. **Documentation**
   - Update README with new tools
   - Add JSDoc comments
   - Update Graph API permissions note (for future)

3. **Integration Testing**
   - End-to-end tests with real Outlook
   - Test error scenarios
   - Test edge cases

### Future: Graph API Backend

Once AppleScript implementation is complete and tested, the same interfaces can be implemented for the Graph API backend:

1. Create `src/graph/calendar-manager.ts`
2. Create `src/graph/mail-sender.ts`
3. Use Microsoft Graph API endpoints:
   - POST `/me/events/{id}/accept`
   - POST `/me/events/{id}/decline`
   - POST `/me/events/{id}/tentativelyAccept`
   - DELETE `/me/events/{id}`
   - PATCH `/me/events/{id}`
   - POST `/me/sendMail`

## AppleScript Implementation Notes

### Event RSVP

Outlook for Mac uses `response status` property:
```applescript
tell application "Microsoft Outlook"
    set myEvent to calendar event id 123
    set response status of myEvent to accept
end tell
```

Values: `accept`, `decline`, `tentative accept`

### Event Deletion

```applescript
tell application "Microsoft Outlook"
    delete calendar event id 123
end tell
```

For recurring events, may need to specify occurrence.

### Event Update

```applescript
tell application "Microsoft Outlook"
    set myEvent to calendar event id 123
    set subject of myEvent to "New Title"
    set start time of myEvent to date "2026-02-05T10:00:00"
    -- etc.
end tell
```

### Email Sending

```applescript
tell application "Microsoft Outlook"
    set newMessage to make new outgoing message with properties {
        subject: "Hello",
        plain text content: "Body text"
    }
    make new recipient at newMessage with properties {
        email address: {address: "user@example.com"}
    }
    send newMessage
end tell
```

For HTML:
```applescript
set html content of newMessage to "<html>...</html>"
```

For attachments:
```applescript
make new attachment at newMessage with properties {
    file: (POSIX file "/path/to/file.pdf")
}
```

## Success Criteria

1. All three event management operations work correctly via MCP tools
2. Email sending works with all rich features (HTML, attachments, CC/BCC)
3. Recurring event handling works correctly (single instance vs entire series)
4. Error messages are clear and helpful
5. Test coverage >85%
6. No regressions in existing functionality
7. Documentation updated

## Future Considerations

- Graph API backend implementation
- Reply/Forward email support
- Draft management
- Batch operations
- Email templates
- Calendar sharing operations
