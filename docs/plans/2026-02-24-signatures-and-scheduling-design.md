# Email Signatures & Calendar Scheduling Design

## Problem

1. **Signatures**: Users have no way to include email signatures in drafts/sends. Microsoft Graph API has no dedicated signature endpoint — signatures are just part of the email body content.
2. **Scheduling**: Users cannot check other people's availability or find meeting times. The server only reads the authenticated user's own calendar.

## Solution

Two independent features adding 4 new tools (74 -> 78 total).

## Feature 1: Email Signatures

### Storage

Signatures stored as HTML at `~/.outlook-mcp/signature.html` (reuses existing config directory). One signature per user.

### Tools

**`set_signature`** — Saves an HTML signature to disk.
- Input: `{ content: string, content_type: 'html' | 'text' }`
- If `content_type` is `text`, wraps in `<pre>` tag for HTML storage
- Returns success confirmation

**`get_signature`** — Reads the stored signature.
- Input: none
- Returns the signature HTML content, or a message saying no signature is set

### Auto-Append Behavior

When creating/sending emails, the server appends `<br><br>` + signature HTML to the body when:
1. A signature file exists at `~/.outlook-mcp/signature.html`
2. The `include_signature` param is `true` (default) on the tool call

**Affected tools:**
- `create_draft`
- `send_email` (AppleScript)
- `prepare_send_email`
- `reply_as_draft`
- `forward_as_draft`

**NOT affected:**
- `update_draft` — user is manually editing, don't double-append

### Signature Append Helper

A shared helper function `appendSignature(body: string, bodyType: 'html' | 'text', includeSignature: boolean): string` that:
1. Checks if `includeSignature` is true
2. Reads `~/.outlook-mcp/signature.html`
3. If file exists, appends separator + signature to body
4. For HTML bodies: `<br><br>` separator
5. For text bodies: `\n\n--\n` separator + strip HTML from signature
6. Returns the modified body

## Feature 2: Calendar Scheduling

### Tools

**`check_availability`** — Check free/busy status for one or more people.
- Input: `{ email_addresses: string[], start_time: string (ISO 8601), end_time: string (ISO 8601), availability_view_interval?: number (minutes, default 30) }`
- Calls: `POST /me/calendar/getSchedule`
- Returns: For each person, their schedule items with start/end times and status (free, tentative, busy, oof, workingElsewhere, unknown)

**`find_meeting_times`** — Find available meeting slots for a group.
- Input: `{ attendees: string[] (email addresses), duration_minutes: number, start_time?: string (ISO 8601), end_time?: string (ISO 8601), max_candidates?: number (default 5) }`
- Calls: `POST /me/findMeetingTimes`
- Returns: Ranked list of suggested time slots with confidence scores and attendee availability

No new permissions needed — `Calendars.ReadWrite` already covers both endpoints.

## Files to Modify/Create

- Create: `src/signature.ts` — signature file I/O (read, write, append helper)
- Create: `src/tools/scheduling.ts` — scheduling tool handlers
- Modify: `src/graph/client/graph-client.ts` — 2 new methods (`getSchedule`, `findMeetingTimes`)
- Modify: `src/graph/repository.ts` — 2 new repository methods
- Modify: `src/tools/mail-send.ts` — `set_signature`/`get_signature` handlers + add `include_signature` to schemas
- Modify: `src/index.ts` — wire 4 new tools + handlers, add `include_signature` to existing tool inputSchemas

## Tests

- Unit tests for `src/signature.ts` (read, write, append logic, missing file handling)
- Unit tests for scheduling tool handlers (mock repository)
- Unit tests for GraphClient `getSchedule` and `findMeetingTimes` methods
- Unit tests for repository scheduling methods
- Update existing mail-send tests to verify signature auto-append
- E2E: tool count 74 -> 78
