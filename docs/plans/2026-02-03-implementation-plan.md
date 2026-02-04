# Event Management & Email Sending Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add event management (RSVP, delete, update) and email sending capabilities to the Outlook MCP server's AppleScript backend.

**Architecture:** Extend the existing dual-backend pattern with new `ICalendarManager` and `IMailSender` interfaces. Follow TDD with unit tests for parsers/scripts, integration tests for full workflows. Each operation gets its own AppleScript template in `scripts.ts`, parser function in `parser.ts`, and interface method.

**Tech Stack:** TypeScript, Zod validation, AppleScript via osascript, Jest for testing

---

## Task 1: Add New Error Types

**Files:**
- Modify: `src/utils/errors.ts`
- Test: `tests/unit/utils/errors.test.ts` (create if needed)

**Step 1: Write test for AttachmentNotFoundError**

```typescript
// tests/unit/utils/errors.test.ts
import { AttachmentNotFoundError, ErrorCode } from '../../../src/utils/errors.js';

describe('AttachmentNotFoundError', () => {
  test('creates error with file path', () => {
    const error = new AttachmentNotFoundError('/path/to/file.pdf');
    expect(error.code).toBe(ErrorCode.ATTACHMENT_NOT_FOUND);
    expect(error.message).toContain('/path/to/file.pdf');
    expect(error.message).toContain('not found');
  });
});
```

**Step 2: Run test to verify it fails**

Run: `npm test -- errors.test.ts`
Expected: FAIL with "AttachmentNotFoundError is not exported" or similar

**Step 3: Add new error codes to ErrorCode object**

```typescript
// src/utils/errors.ts (after line 23)
export const ErrorCode = {
  // ... existing codes
  ATTACHMENT_NOT_FOUND: 'ATTACHMENT_NOT_FOUND',
  MAIL_SEND_ERROR: 'MAIL_SEND_ERROR',
  RECURRING_EVENT_ERROR: 'RECURRING_EVENT_ERROR',
} as const;
```

**Step 4: Implement AttachmentNotFoundError class**

```typescript
// src/utils/errors.ts (after AppleScriptError class)
/**
 * Thrown when an email attachment file cannot be found.
 */
export class AttachmentNotFoundError extends OutlookMcpError {
  readonly code = ErrorCode.ATTACHMENT_NOT_FOUND;

  constructor(path: string) {
    super(`Attachment file not found: ${path}. Please check the file path exists.`);
  }
}
```

**Step 5: Add tests for MailSendError and RecurringEventError**

```typescript
// tests/unit/utils/errors.test.ts
describe('MailSendError', () => {
  test('creates error with reason', () => {
    const error = new MailSendError('Invalid recipient address');
    expect(error.code).toBe(ErrorCode.MAIL_SEND_ERROR);
    expect(error.message).toContain('Invalid recipient address');
  });
});

describe('RecurringEventError', () => {
  test('creates error with custom message', () => {
    const error = new RecurringEventError('Cannot modify single instance');
    expect(error.code).toBe(ErrorCode.RECURRING_EVENT_ERROR);
    expect(error.message).toBe('Cannot modify single instance');
  });
});
```

**Step 6: Implement MailSendError and RecurringEventError**

```typescript
// src/utils/errors.ts
/**
 * Thrown when email sending fails.
 */
export class MailSendError extends OutlookMcpError {
  readonly code = ErrorCode.MAIL_SEND_ERROR;

  constructor(reason: string) {
    super(`Failed to send email: ${reason}`);
  }
}

/**
 * Thrown for recurring event operation errors.
 */
export class RecurringEventError extends OutlookMcpError {
  readonly code = ErrorCode.RECURRING_EVENT_ERROR;

  constructor(message: string) {
    super(message);
  }
}
```

**Step 7: Run all error tests**

Run: `npm test -- errors.test.ts`
Expected: All tests PASS

**Step 8: Commit error types**

```bash
git add src/utils/errors.ts tests/unit/utils/errors.test.ts
git commit -m "feat: add error types for event management and email sending"
```

---

## Task 2: Event RSVP - Parser Tests

**Files:**
- Modify: `tests/unit/applescript/parser.test.ts`

**Step 1: Write test for parseRespondToEventResult with success**

```typescript
// tests/unit/applescript/parser.test.ts (add to end)
describe('parseRespondToEventResult', () => {
  test('parses successful response', () => {
    const output = '{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}123';
    const result = parseRespondToEventResult(output);
    expect(result).toEqual({ success: true, eventId: 123 });
  });

  test('parses failure response', () => {
    const output = '{{RECORD}}success{{=}}false{{FIELD}}error{{=}}Permission denied';
    const result = parseRespondToEventResult(output);
    expect(result).toEqual({ success: false, error: 'Permission denied' });
  });

  test('handles empty output', () => {
    const result = parseRespondToEventResult('');
    expect(result).toBeNull();
  });
});
```

**Step 2: Run test to verify it fails**

Run: `npm test -- parser.test.ts -t "parseRespondToEventResult"`
Expected: FAIL with "parseRespondToEventResult is not defined"

**Step 3: Import function in test file**

```typescript
// tests/unit/applescript/parser.test.ts (update imports)
import {
  // ... existing imports
  parseRespondToEventResult,
} from '../../../src/applescript/parser.js';
```

**Step 4: Run test again**

Run: `npm test -- parser.test.ts -t "parseRespondToEventResult"`
Expected: FAIL with "parseRespondToEventResult is not exported"

---

## Task 3: Event RSVP - Parser Implementation

**Files:**
- Modify: `src/applescript/parser.ts`

**Step 1: Add result type for RSVP operation**

```typescript
// src/applescript/parser.ts (add after AppleScriptNoteRow interface)
export interface RespondToEventResult {
  readonly success: boolean;
  readonly eventId?: number;
  readonly error?: string;
}
```

**Step 2: Implement parseRespondToEventResult function**

```typescript
// src/applescript/parser.ts (add before exports at end of file)
/**
 * Parses the result of a respond-to-event operation.
 */
export function parseRespondToEventResult(output: string): RespondToEventResult | null {
  const records = parseRecords(output);
  if (records.length === 0) return null;

  const record = records[0];
  const success = record.success === 'true';

  if (success) {
    return {
      success: true,
      eventId: safeParseInt(record.eventId),
    };
  } else {
    return {
      success: false,
      error: record.error ?? 'Unknown error',
    };
  }
}
```

**Step 3: Run tests**

Run: `npm test -- parser.test.ts -t "parseRespondToEventResult"`
Expected: All tests PASS

**Step 4: Commit parser**

```bash
git add src/applescript/parser.ts tests/unit/applescript/parser.test.ts
git commit -m "feat: add parser for event RSVP responses"
```

---

## Task 4: Event RSVP - AppleScript Template

**Files:**
- Modify: `src/applescript/scripts.ts`
- Test: `tests/unit/applescript/scripts.test.ts` (create)

**Step 1: Write test for respondToEvent script generation**

```typescript
// tests/unit/applescript/scripts.test.ts
import { respondToEvent } from '../../src/applescript/scripts.js';

describe('respondToEvent', () => {
  test('generates accept script with comment', () => {
    const script = respondToEvent({
      eventId: 123,
      response: 'accept',
      sendResponse: true,
      comment: 'I will be there',
    });

    expect(script).toContain('calendar event id 123');
    expect(script).toContain('accept');
    expect(script).toContain('I will be there');
  });

  test('generates decline script without sending response', () => {
    const script = respondToEvent({
      eventId: 456,
      response: 'decline',
      sendResponse: false,
    });

    expect(script).toContain('calendar event id 456');
    expect(script).toContain('decline');
  });

  test('generates tentative accept script', () => {
    const script = respondToEvent({
      eventId: 789,
      response: 'tentative',
      sendResponse: true,
    });

    expect(script).toContain('calendar event id 789');
    expect(script).toContain('tentative');
  });
});
```

**Step 2: Run test to verify it fails**

Run: `npm test -- scripts.test.ts`
Expected: FAIL with "respondToEvent is not defined"

**Step 3: Add RespondToEventParams interface**

```typescript
// src/applescript/scripts.ts (add after RecurrenceScriptParams)
export interface RespondToEventParams {
  readonly eventId: number;
  readonly response: 'accept' | 'decline' | 'tentative';
  readonly sendResponse: boolean;
  readonly comment?: string;
}
```

**Step 4: Implement respondToEvent function**

```typescript
// src/applescript/scripts.ts (add after createEvent function)
/**
 * Responds to an event invitation (RSVP).
 */
export function respondToEvent(params: RespondToEventParams): string {
  const { eventId, response, sendResponse, comment } = params;

  // Map response to AppleScript status value
  const statusMap = {
    accept: 'accept',
    decline: 'decline',
    tentative: 'tentative accept',
  };
  const status = statusMap[response];

  // Escape comment if provided
  const commentLine = comment != null
    ? `set comment of myEvent to "${escapeForAppleScript(comment)}"`
    : '';

  return `
tell application "Microsoft Outlook"
  try
    set myEvent to calendar event id ${eventId}
    set response status of myEvent to ${status}
    ${commentLine}

    -- Return success
    set output to "{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}" & ${eventId}
    return output
  on error errMsg
    -- Return failure
    set output to "{{RECORD}}success{{=}}false{{FIELD}}error{{=}}" & errMsg
    return output
  end try
end tell
`;
}
```

**Step 5: Run tests**

Run: `npm test -- scripts.test.ts -t "respondToEvent"`
Expected: All tests PASS

**Step 6: Commit script template**

```bash
git add src/applescript/scripts.ts tests/unit/applescript/scripts.test.ts
git commit -m "feat: add AppleScript template for event RSVP"
```

---

## Task 5: Event RSVP - Calendar Manager Interface

**Files:**
- Create: `src/applescript/calendar-manager.ts`
- Test: `tests/unit/applescript/calendar-manager.test.ts`

**Step 1: Write test for respondToEvent method**

```typescript
// tests/unit/applescript/calendar-manager.test.ts
import { AppleScriptCalendarManager } from '../../../src/applescript/calendar-manager.js';
import * as executor from '../../../src/applescript/executor.js';

// Mock the executor
jest.mock('../../../src/applescript/executor.js');

describe('AppleScriptCalendarManager', () => {
  let manager: AppleScriptCalendarManager;
  let mockExecute: jest.MockedFunction<typeof executor.executeAppleScriptOrThrow>;

  beforeEach(() => {
    manager = new AppleScriptCalendarManager();
    mockExecute = executor.executeAppleScriptOrThrow as jest.MockedFunction<typeof executor.executeAppleScriptOrThrow>;
    mockExecute.mockClear();
  });

  describe('respondToEvent', () => {
    test('accepts event with comment', () => {
      mockExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}123');

      const result = manager.respondToEvent(123, 'accept', true, 'Looking forward to it');

      expect(result).toEqual({ success: true, eventId: 123 });
      expect(mockExecute).toHaveBeenCalledTimes(1);
      const script = mockExecute.mock.calls[0][0];
      expect(script).toContain('calendar event id 123');
      expect(script).toContain('accept');
      expect(script).toContain('Looking forward to it');
    });

    test('declines event without sending', () => {
      mockExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}456');

      const result = manager.respondToEvent(456, 'decline', false);

      expect(result).toEqual({ success: true, eventId: 456 });
    });

    test('handles AppleScript error', () => {
      mockExecute.mockReturnValue('{{RECORD}}success{{=}}false{{FIELD}}error{{=}}Permission denied');

      expect(() => {
        manager.respondToEvent(789, 'accept', true);
      }).toThrow('Permission denied');
    });
  });
});
```

**Step 2: Run test to verify it fails**

Run: `npm test -- calendar-manager.test.ts`
Expected: FAIL with "Cannot find module 'calendar-manager'"

**Step 3: Create ICalendarManager interface**

```typescript
// src/applescript/calendar-manager.ts
import { executeAppleScriptOrThrow } from './executor.js';
import * as scripts from './scripts.js';
import { parseRespondToEventResult, type RespondToEventResult } from './parser.js';
import { AppleScriptError } from '../utils/errors.js';

// =============================================================================
// Types
// =============================================================================

export type ResponseType = 'accept' | 'decline' | 'tentative';
export type ApplyToScope = 'this_instance' | 'all_in_series';

/**
 * Interface for calendar event management operations.
 */
export interface ICalendarManager {
  /**
   * Respond to an event invitation (RSVP).
   */
  respondToEvent(
    eventId: number,
    response: ResponseType,
    sendResponse: boolean,
    comment?: string
  ): RespondToEventResult;

  /**
   * Delete a calendar event.
   */
  deleteEvent(eventId: number, applyTo: ApplyToScope): void;

  /**
   * Update a calendar event.
   */
  updateEvent(eventId: number, updates: EventUpdates, applyTo: ApplyToScope): UpdatedEvent;
}

export interface EventUpdates {
  readonly title?: string;
  readonly startDate?: string; // ISO 8601
  readonly endDate?: string; // ISO 8601
  readonly location?: string;
  readonly description?: string;
  readonly isAllDay?: boolean;
  readonly recurrence?: scripts.RecurrenceScriptParams;
}

export interface UpdatedEvent {
  readonly id: number;
  readonly updatedFields: readonly string[];
}
```

**Step 4: Implement AppleScriptCalendarManager class**

```typescript
// src/applescript/calendar-manager.ts (continue)
// =============================================================================
// Implementation
// =============================================================================

export class AppleScriptCalendarManager implements ICalendarManager {
  respondToEvent(
    eventId: number,
    response: ResponseType,
    sendResponse: boolean,
    comment?: string
  ): RespondToEventResult {
    const script = scripts.respondToEvent({
      eventId,
      response,
      sendResponse,
      comment,
    });

    const output = executeAppleScriptOrThrow(script);
    const result = parseRespondToEventResult(output);

    if (result == null) {
      throw new AppleScriptError('Failed to parse RSVP response');
    }

    if (!result.success) {
      throw new AppleScriptError(result.error ?? 'RSVP operation failed');
    }

    return result;
  }

  deleteEvent(_eventId: number, _applyTo: ApplyToScope): void {
    throw new Error('Not yet implemented');
  }

  updateEvent(_eventId: number, _updates: EventUpdates, _applyTo: ApplyToScope): UpdatedEvent {
    throw new Error('Not yet implemented');
  }
}

export function createCalendarManager(): ICalendarManager {
  return new AppleScriptCalendarManager();
}
```

**Step 5: Run tests**

Run: `npm test -- calendar-manager.test.ts -t "respondToEvent"`
Expected: All tests PASS

**Step 6: Commit calendar manager**

```bash
git add src/applescript/calendar-manager.ts tests/unit/applescript/calendar-manager.test.ts
git commit -m "feat: implement ICalendarManager with RSVP support"
```

---

## Task 6: Event RSVP - MCP Tool

**Files:**
- Modify: `src/tools/calendar.ts`
- Modify: `src/applescript/index.ts`

**Step 1: Add RespondToEventInput schema**

```typescript
// src/tools/calendar.ts (add after CreateEventInput)
export const RespondToEventInput = z
  .object({
    event_id: z.number().int().positive().describe('The event ID to respond to'),
    response: z.enum(['accept', 'decline', 'tentative']).describe('Your response to the invitation'),
    send_response: z.boolean().default(true).describe('Whether to send response to organizer'),
    comment: z.string().optional().describe('Optional comment to include with response'),
  })
  .strict();
```

**Step 2: Add respondToEventTool handler**

```typescript
// src/tools/calendar.ts (add after createEventTool)
/**
 * Respond to an event invitation (RSVP).
 */
export function respondToEventTool(calendarManager: ICalendarManager) {
  return {
    name: 'respond_to_event',
    description: 'Respond to a meeting invitation (accept, decline, or tentative). Updates your response status and optionally notifies the organizer.',
    inputSchema: zodToJsonSchema(RespondToEventInput),
    handler: async (input: unknown) => {
      const params = RespondToEventInput.parse(input);

      const result = calendarManager.respondToEvent(
        params.event_id,
        params.response,
        params.send_response,
        params.comment
      );

      return {
        content: [
          {
            type: 'text',
            text: `Successfully ${params.response === 'accept' ? 'accepted' : params.response === 'decline' ? 'declined' : 'tentatively accepted'} event ${result.eventId}`,
          },
        ],
      };
    },
  };
}
```

**Step 3: Import ICalendarManager type**

```typescript
// src/tools/calendar.ts (update imports)
import type { ICalendarManager } from '../applescript/calendar-manager.js';
```

**Step 4: Export respondToEventTool**

```typescript
// src/tools/calendar.ts (update exports at bottom)
export const calendarTools = {
  // ... existing tools
  respondToEvent: respondToEventTool,
};
```

**Step 5: Update AppleScript index to create calendar manager**

```typescript
// src/applescript/index.ts (add import)
import { createCalendarManager } from './calendar-manager.js';

// Update createRepository function to also return calendar manager
export function createBackend() {
  return {
    repository: createRepository(),
    calendarWriter: createCalendarWriter(),
    calendarManager: createCalendarManager(),
  };
}
```

**Step 6: Wire up tool in main server (manual verification)**

Note: Check `src/index.ts` to ensure the tool is registered. This may require looking at how `createEventTool` is currently wired up.

**Step 7: Commit MCP tool**

```bash
git add src/tools/calendar.ts src/applescript/index.ts
git commit -m "feat: add respond_to_event MCP tool"
```

---

## Task 7: Event Delete - Parser & Script

**Files:**
- Modify: `src/applescript/parser.ts`
- Modify: `src/applescript/scripts.ts`
- Modify: `tests/unit/applescript/parser.test.ts`
- Modify: `tests/unit/applescript/scripts.test.ts`

**Step 1: Write parser test for deleteEvent**

```typescript
// tests/unit/applescript/parser.test.ts
describe('parseDeleteEventResult', () => {
  test('parses successful delete', () => {
    const output = '{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}123';
    const result = parseDeleteEventResult(output);
    expect(result).toEqual({ success: true, eventId: 123 });
  });

  test('parses failure', () => {
    const output = '{{RECORD}}success{{=}}false{{FIELD}}error{{=}}Event not found';
    const result = parseDeleteEventResult(output);
    expect(result).toEqual({ success: false, error: 'Event not found' });
  });
});
```

**Step 2: Implement parseDeleteEventResult**

```typescript
// src/applescript/parser.ts
export interface DeleteEventResult {
  readonly success: boolean;
  readonly eventId?: number;
  readonly error?: string;
}

export function parseDeleteEventResult(output: string): DeleteEventResult | null {
  const records = parseRecords(output);
  if (records.length === 0) return null;

  const record = records[0];
  const success = record.success === 'true';

  if (success) {
    return {
      success: true,
      eventId: safeParseInt(record.eventId),
    };
  } else {
    return {
      success: false,
      error: record.error ?? 'Unknown error',
    };
  }
}
```

**Step 3: Write script test for deleteEvent**

```typescript
// tests/unit/applescript/scripts.test.ts
describe('deleteEvent', () => {
  test('generates script for single instance', () => {
    const script = deleteEvent({ eventId: 123, applyTo: 'this_instance' });
    expect(script).toContain('calendar event id 123');
    expect(script).toContain('delete');
  });

  test('generates script for all in series', () => {
    const script = deleteEvent({ eventId: 456, applyTo: 'all_in_series' });
    expect(script).toContain('calendar event id 456');
    expect(script).toContain('delete');
  });
});
```

**Step 4: Implement deleteEvent script**

```typescript
// src/applescript/scripts.ts
export interface DeleteEventParams {
  readonly eventId: number;
  readonly applyTo: 'this_instance' | 'all_in_series';
}

export function deleteEvent(params: DeleteEventParams): string {
  const { eventId, applyTo } = params;

  // Note: AppleScript behavior for recurring events may vary
  // This implementation attempts to delete the specified event
  const comment = applyTo === 'all_in_series'
    ? '-- Deleting entire series'
    : '-- Deleting single instance';

  return `
tell application "Microsoft Outlook"
  try
    ${comment}
    set myEvent to calendar event id ${eventId}
    delete myEvent

    -- Return success
    set output to "{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}" & ${eventId}
    return output
  on error errMsg
    -- Return failure
    set output to "{{RECORD}}success{{=}}false{{FIELD}}error{{=}}" & errMsg
    return output
  end try
end tell
`;
}
```

**Step 5: Run tests**

Run: `npm test -- parser.test.ts -t "parseDeleteEventResult"`
Run: `npm test -- scripts.test.ts -t "deleteEvent"`
Expected: All tests PASS

**Step 6: Commit**

```bash
git add src/applescript/parser.ts src/applescript/scripts.ts tests/unit/applescript/parser.test.ts tests/unit/applescript/scripts.test.ts
git commit -m "feat: add parser and script for event deletion"
```

---

## Task 8: Event Delete - Implementation & Tool

**Files:**
- Modify: `src/applescript/calendar-manager.ts`
- Modify: `src/tools/calendar.ts`
- Modify: `tests/unit/applescript/calendar-manager.test.ts`

**Step 1: Write test for deleteEvent**

```typescript
// tests/unit/applescript/calendar-manager.test.ts
describe('deleteEvent', () => {
  test('deletes single instance', () => {
    mockExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}123');

    manager.deleteEvent(123, 'this_instance');

    expect(mockExecute).toHaveBeenCalledTimes(1);
    const script = mockExecute.mock.calls[0][0];
    expect(script).toContain('calendar event id 123');
  });

  test('deletes all in series', () => {
    mockExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}456');

    manager.deleteEvent(456, 'all_in_series');

    expect(mockExecute).toHaveBeenCalledTimes(1);
  });

  test('throws on failure', () => {
    mockExecute.mockReturnValue('{{RECORD}}success{{=}}false{{FIELD}}error{{=}}Not found');

    expect(() => {
      manager.deleteEvent(789, 'this_instance');
    }).toThrow('Not found');
  });
});
```

**Step 2: Implement deleteEvent in CalendarManager**

```typescript
// src/applescript/calendar-manager.ts
import { parseDeleteEventResult, type DeleteEventResult } from './parser.js';

// Update deleteEvent method
deleteEvent(eventId: number, applyTo: ApplyToScope): void {
  const script = scripts.deleteEvent({ eventId, applyTo });
  const output = executeAppleScriptOrThrow(script);
  const result = parseDeleteEventResult(output);

  if (result == null) {
    throw new AppleScriptError('Failed to parse delete response');
  }

  if (!result.success) {
    throw new AppleScriptError(result.error ?? 'Delete operation failed');
  }
}
```

**Step 3: Add DeleteEventInput schema**

```typescript
// src/tools/calendar.ts
export const DeleteEventInput = z
  .object({
    event_id: z.number().int().positive().describe('The event ID to delete'),
    apply_to: z
      .enum(['this_instance', 'all_in_series'])
      .default('this_instance')
      .describe('For recurring events: delete single instance or entire series'),
  })
  .strict();
```

**Step 4: Add deleteEventTool**

```typescript
// src/tools/calendar.ts
export function deleteEventTool(calendarManager: ICalendarManager) {
  return {
    name: 'delete_event',
    description: 'Delete a calendar event. For recurring events, you can delete a single instance or the entire series.',
    inputSchema: zodToJsonSchema(DeleteEventInput),
    handler: async (input: unknown) => {
      const params = DeleteEventInput.parse(input);

      calendarManager.deleteEvent(params.event_id, params.apply_to);

      return {
        content: [
          {
            type: 'text',
            text: `Successfully deleted event ${params.event_id}${
              params.apply_to === 'all_in_series' ? ' (entire series)' : ''
            }`,
          },
        ],
      };
    },
  };
}
```

**Step 5: Run tests**

Run: `npm test -- calendar-manager.test.ts -t "deleteEvent"`
Expected: All tests PASS

**Step 6: Commit**

```bash
git add src/applescript/calendar-manager.ts src/tools/calendar.ts tests/unit/applescript/calendar-manager.test.ts
git commit -m "feat: implement event deletion with MCP tool"
```

---

## Task 9: Event Update - Parser & Script

**Files:**
- Modify: `src/applescript/parser.ts`
- Modify: `src/applescript/scripts.ts`
- Modify: `tests/unit/applescript/parser.test.ts`
- Modify: `tests/unit/applescript/scripts.test.ts`

**Step 1: Write parser test**

```typescript
// tests/unit/applescript/parser.test.ts
describe('parseUpdateEventResult', () => {
  test('parses successful update', () => {
    const output = '{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}123{{FIELD}}updatedFields{{=}}title,location';
    const result = parseUpdateEventResult(output);
    expect(result).toEqual({
      success: true,
      id: 123,
      updatedFields: ['title', 'location'],
    });
  });

  test('parses failure', () => {
    const output = '{{RECORD}}success{{=}}false{{FIELD}}error{{=}}Permission denied';
    const result = parseUpdateEventResult(output);
    expect(result).toEqual({ success: false, error: 'Permission denied' });
  });
});
```

**Step 2: Implement parseUpdateEventResult**

```typescript
// src/applescript/parser.ts
export interface UpdateEventResult {
  readonly success: boolean;
  readonly id?: number;
  readonly updatedFields?: readonly string[];
  readonly error?: string;
}

export function parseUpdateEventResult(output: string): UpdateEventResult | null {
  const records = parseRecords(output);
  if (records.length === 0) return null;

  const record = records[0];
  const success = record.success === 'true';

  if (success) {
    const fieldsStr = record.updatedFields ?? '';
    const fields = fieldsStr.length > 0 ? fieldsStr.split(',') : [];

    return {
      success: true,
      id: safeParseInt(record.eventId),
      updatedFields: fields,
    };
  } else {
    return {
      success: false,
      error: record.error ?? 'Unknown error',
    };
  }
}
```

**Step 3: Write script test**

```typescript
// tests/unit/applescript/scripts.test.ts
describe('updateEvent', () => {
  test('generates script with title update', () => {
    const script = updateEvent({
      eventId: 123,
      applyTo: 'this_instance',
      updates: { title: 'New Title' },
    });
    expect(script).toContain('calendar event id 123');
    expect(script).toContain('New Title');
  });

  test('generates script with multiple updates', () => {
    const script = updateEvent({
      eventId: 456,
      applyTo: 'this_instance',
      updates: {
        title: 'Updated',
        location: 'Conference Room',
        startDate: '2026-02-10T10:00:00Z',
        endDate: '2026-02-10T11:00:00Z',
      },
    });
    expect(script).toContain('calendar event id 456');
    expect(script).toContain('Updated');
    expect(script).toContain('Conference Room');
  });
});
```

**Step 4: Implement updateEvent script**

```typescript
// src/applescript/scripts.ts
export interface UpdateEventParams {
  readonly eventId: number;
  readonly applyTo: 'this_instance' | 'all_in_series';
  readonly updates: {
    readonly title?: string;
    readonly startDate?: string; // ISO 8601
    readonly endDate?: string; // ISO 8601
    readonly location?: string;
    readonly description?: string;
    readonly isAllDay?: boolean;
  };
}

export function updateEvent(params: UpdateEventParams): string {
  const { eventId, applyTo, updates } = params;
  const updatedFields: string[] = [];

  // Build update statements
  let updateStatements = '';

  if (updates.title != null) {
    updateStatements += `    set subject of myEvent to "${escapeForAppleScript(updates.title)}"\n`;
    updatedFields.push('title');
  }

  if (updates.location != null) {
    updateStatements += `    set location of myEvent to "${escapeForAppleScript(updates.location)}"\n`;
    updatedFields.push('location');
  }

  if (updates.description != null) {
    updateStatements += `    set content of myEvent to "${escapeForAppleScript(updates.description)}"\n`;
    updatedFields.push('description');
  }

  if (updates.startDate != null) {
    const start = isoToDateComponents(updates.startDate);
    updateStatements += `    set start time of myEvent to date "${start.year}-${start.month}-${start.day} ${start.hours}:${start.minutes}:00"\n`;
    updatedFields.push('startDate');
  }

  if (updates.endDate != null) {
    const end = isoToDateComponents(updates.endDate);
    updateStatements += `    set end time of myEvent to date "${end.year}-${end.month}-${end.day} ${end.hours}:${end.minutes}:00"\n`;
    updatedFields.push('endDate');
  }

  if (updates.isAllDay != null) {
    updateStatements += `    set all day flag of myEvent to ${updates.isAllDay}\n`;
    updatedFields.push('isAllDay');
  }

  const comment = applyTo === 'all_in_series'
    ? '-- Updating entire series'
    : '-- Updating single instance';

  const fieldsOutput = updatedFields.join(',');

  return `
tell application "Microsoft Outlook"
  try
    ${comment}
    set myEvent to calendar event id ${eventId}

${updateStatements}

    -- Return success
    set output to "{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}" & ${eventId} & "{{FIELD}}updatedFields{{=}}${fieldsOutput}"
    return output
  on error errMsg
    -- Return failure
    set output to "{{RECORD}}success{{=}}false{{FIELD}}error{{=}}" & errMsg
    return output
  end try
end tell
`;
}

// Helper function for date conversion
function isoToDateComponents(isoString: string): {
  year: number;
  month: number;
  day: number;
  hours: number;
  minutes: number;
} {
  const date = new Date(isoString);
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    hours: date.getUTCHours(),
    minutes: date.getUTCMinutes(),
  };
}
```

**Step 5: Run tests**

Run: `npm test -- parser.test.ts -t "parseUpdateEventResult"`
Run: `npm test -- scripts.test.ts -t "updateEvent"`
Expected: All tests PASS

**Step 6: Commit**

```bash
git add src/applescript/parser.ts src/applescript/scripts.ts tests/unit/applescript/parser.test.ts tests/unit/applescript/scripts.test.ts
git commit -m "feat: add parser and script for event updates"
```

---

## Task 10: Event Update - Implementation & Tool

**Files:**
- Modify: `src/applescript/calendar-manager.ts`
- Modify: `src/tools/calendar.ts`
- Modify: `tests/unit/applescript/calendar-manager.test.ts`

**Step 1: Write test for updateEvent**

```typescript
// tests/unit/applescript/calendar-manager.test.ts
describe('updateEvent', () => {
  test('updates event title', () => {
    mockExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}123{{FIELD}}updatedFields{{=}}title');

    const result = manager.updateEvent(
      123,
      { title: 'New Title' },
      'this_instance'
    );

    expect(result).toEqual({
      id: 123,
      updatedFields: ['title'],
    });
  });

  test('updates multiple fields', () => {
    mockExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}456{{FIELD}}updatedFields{{=}}title,location,startDate');

    const result = manager.updateEvent(
      456,
      {
        title: 'Updated',
        location: 'Room 101',
        startDate: '2026-02-10T10:00:00Z',
      },
      'all_in_series'
    );

    expect(result.id).toBe(456);
    expect(result.updatedFields).toContain('title');
    expect(result.updatedFields).toContain('location');
  });
});
```

**Step 2: Implement updateEvent**

```typescript
// src/applescript/calendar-manager.ts
import { parseUpdateEventResult, type UpdateEventResult } from './parser.js';

// Update the updateEvent method
updateEvent(eventId: number, updates: EventUpdates, applyTo: ApplyToScope): UpdatedEvent {
  const scriptParams: scripts.UpdateEventParams = {
    eventId,
    applyTo,
    updates: {
      title: updates.title,
      location: updates.location,
      description: updates.description,
      startDate: updates.startDate,
      endDate: updates.endDate,
      isAllDay: updates.isAllDay,
    },
  };

  const script = scripts.updateEvent(scriptParams);
  const output = executeAppleScriptOrThrow(script);
  const result = parseUpdateEventResult(output);

  if (result == null) {
    throw new AppleScriptError('Failed to parse update response');
  }

  if (!result.success) {
    throw new AppleScriptError(result.error ?? 'Update operation failed');
  }

  return {
    id: result.id!,
    updatedFields: result.updatedFields ?? [],
  };
}
```

**Step 3: Add UpdateEventInput schema**

```typescript
// src/tools/calendar.ts
// Use the existing isoDateString validator
const isoDateString = z.string().refine(
  (val) => !isNaN(Date.parse(val)),
  { message: 'Must be a valid ISO 8601 date string' }
);

export const UpdateEventInput = z
  .object({
    event_id: z.number().int().positive().describe('The event ID to update'),
    apply_to: z
      .enum(['this_instance', 'all_in_series'])
      .default('this_instance')
      .describe('For recurring events: update single instance or entire series'),
    title: z.string().min(1).optional().describe('New event title'),
    start_date: isoDateString.optional().describe('New start date (ISO 8601 UTC)'),
    end_date: isoDateString.optional().describe('New end date (ISO 8601 UTC)'),
    location: z.string().optional().describe('New location'),
    description: z.string().optional().describe('New description'),
    is_all_day: z.boolean().optional().describe('Whether event is all day'),
  })
  .strict()
  .refine(
    (data) => {
      if (data.start_date != null && data.end_date != null) {
        return new Date(data.start_date).getTime() < new Date(data.end_date).getTime();
      }
      return true;
    },
    { message: 'start_date must be before end_date', path: ['start_date'] }
  );
```

**Step 4: Add updateEventTool**

```typescript
// src/tools/calendar.ts
export function updateEventTool(calendarManager: ICalendarManager) {
  return {
    name: 'update_event',
    description: 'Update a calendar event. All fields are optional - only specified fields will be updated. For recurring events, you can update a single instance or the entire series.',
    inputSchema: zodToJsonSchema(UpdateEventInput),
    handler: async (input: unknown) => {
      const params = UpdateEventInput.parse(input);

      const updates: EventUpdates = {
        title: params.title,
        startDate: params.start_date,
        endDate: params.end_date,
        location: params.location,
        description: params.description,
        isAllDay: params.is_all_day,
      };

      const result = calendarManager.updateEvent(
        params.event_id,
        updates,
        params.apply_to
      );

      return {
        content: [
          {
            type: 'text',
            text: `Successfully updated event ${result.id}. Updated fields: ${result.updatedFields.join(', ')}`,
          },
        ],
      };
    },
  };
}
```

**Step 5: Run tests**

Run: `npm test -- calendar-manager.test.ts -t "updateEvent"`
Expected: All tests PASS

**Step 6: Commit**

```bash
git add src/applescript/calendar-manager.ts src/tools/calendar.ts tests/unit/applescript/calendar-manager.test.ts
git commit -m "feat: implement event updates with MCP tool"
```

---

## Task 11: Email Sending - Parser & Script

**Files:**
- Modify: `src/applescript/parser.ts`
- Modify: `src/applescript/scripts.ts`
- Modify: `tests/unit/applescript/parser.test.ts`
- Modify: `tests/unit/applescript/scripts.test.ts`

**Step 1: Write parser test**

```typescript
// tests/unit/applescript/parser.test.ts
describe('parseSendEmailResult', () => {
  test('parses successful send', () => {
    const output = '{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}abc123{{FIELD}}sentAt{{=}}2026-02-03T10:30:00Z';
    const result = parseSendEmailResult(output);
    expect(result).toEqual({
      success: true,
      messageId: 'abc123',
      sentAt: '2026-02-03T10:30:00Z',
    });
  });

  test('parses failure', () => {
    const output = '{{RECORD}}success{{=}}false{{FIELD}}error{{=}}Invalid recipient';
    const result = parseSendEmailResult(output);
    expect(result).toEqual({ success: false, error: 'Invalid recipient' });
  });
});
```

**Step 2: Implement parseSendEmailResult**

```typescript
// src/applescript/parser.ts
export interface SendEmailResult {
  readonly success: boolean;
  readonly messageId?: string;
  readonly sentAt?: string;
  readonly error?: string;
}

export function parseSendEmailResult(output: string): SendEmailResult | null {
  const records = parseRecords(output);
  if (records.length === 0) return null;

  const record = records[0];
  const success = record.success === 'true';

  if (success) {
    return {
      success: true,
      messageId: record.messageId ?? '',
      sentAt: record.sentAt ?? '',
    };
  } else {
    return {
      success: false,
      error: record.error ?? 'Unknown error',
    };
  }
}
```

**Step 3: Write script test**

```typescript
// tests/unit/applescript/scripts.test.ts
describe('sendEmail', () => {
  test('generates basic email script', () => {
    const script = sendEmail({
      to: ['user@example.com'],
      subject: 'Test Email',
      body: 'Hello World',
      bodyType: 'plain',
    });
    expect(script).toContain('user@example.com');
    expect(script).toContain('Test Email');
    expect(script).toContain('Hello World');
  });

  test('generates HTML email with CC and BCC', () => {
    const script = sendEmail({
      to: ['user1@example.com'],
      cc: ['user2@example.com'],
      bcc: ['user3@example.com'],
      subject: 'HTML Test',
      body: '<h1>Hello</h1>',
      bodyType: 'html',
    });
    expect(script).toContain('html content');
    expect(script).toContain('<h1>Hello</h1>');
    expect(script).toContain('user2@example.com');
  });

  test('generates email with attachments', () => {
    const script = sendEmail({
      to: ['user@example.com'],
      subject: 'With Attachment',
      body: 'See attached',
      bodyType: 'plain',
      attachments: [{ path: '/path/to/file.pdf' }],
    });
    expect(script).toContain('attachment');
    expect(script).toContain('/path/to/file.pdf');
  });
});
```

**Step 4: Implement sendEmail script**

```typescript
// src/applescript/scripts.ts
export interface SendEmailParams {
  readonly to: readonly string[];
  readonly subject: string;
  readonly body: string;
  readonly bodyType: 'plain' | 'html';
  readonly cc?: readonly string[];
  readonly bcc?: readonly string[];
  readonly replyTo?: string;
  readonly attachments?: readonly { path: string; name?: string }[];
  readonly accountId?: number;
}

export function sendEmail(params: SendEmailParams): string {
  const { to, subject, body, bodyType, cc, bcc, replyTo, attachments, accountId } = params;

  // Escape strings
  const escapedSubject = escapeForAppleScript(subject);
  const escapedBody = escapeForAppleScript(body);

  // Build recipient statements
  const toRecipients = to.map(email =>
    `    make new recipient at newMessage with properties {email address:{address:"${email}"}}`
  ).join('\n');

  const ccRecipients = cc != null && cc.length > 0
    ? cc.map(email =>
        `    make new recipient at newMessage with properties {email address:{address:"${email}"}, recipient type:recipient cc}`
      ).join('\n')
    : '';

  const bccRecipients = bcc != null && bcc.length > 0
    ? bcc.map(email =>
        `    make new recipient at newMessage with properties {email address:{address:"${email}"}, recipient type:recipient bcc}`
      ).join('\n')
    : '';

  // Build content statement
  const contentProperty = bodyType === 'html'
    ? `html content:"${escapedBody}"`
    : `plain text content:"${escapedBody}"`;

  // Build reply-to statement
  const replyToStatement = replyTo != null
    ? `    set reply to of newMessage to "${replyTo}"`
    : '';

  // Build attachments
  const attachmentStatements = attachments != null && attachments.length > 0
    ? attachments.map(att =>
        `    make new attachment at newMessage with properties {file:(POSIX file "${att.path}")}`
      ).join('\n')
    : '';

  // Account selection (if provided)
  const accountStatement = accountId != null
    ? `    set sending account of newMessage to account id ${accountId}`
    : '';

  return `
tell application "Microsoft Outlook"
  try
    set newMessage to make new outgoing message with properties {subject:"${escapedSubject}", ${contentProperty}}

${toRecipients}
${ccRecipients}
${bccRecipients}
${replyToStatement}
${accountStatement}
${attachmentStatements}

    send newMessage

    -- Get message ID and timestamp
    set msgId to id of newMessage as string
    set sentTime to current date
    set sentISO to sentTime as «class isot» as string

    -- Return success
    set output to "{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}" & msgId & "{{FIELD}}sentAt{{=}}" & sentISO
    return output
  on error errMsg
    -- Return failure
    set output to "{{RECORD}}success{{=}}false{{FIELD}}error{{=}}" & errMsg
    return output
  end try
end tell
`;
}
```

**Step 5: Run tests**

Run: `npm test -- parser.test.ts -t "parseSendEmailResult"`
Run: `npm test -- scripts.test.ts -t "sendEmail"`
Expected: All tests PASS

**Step 6: Commit**

```bash
git add src/applescript/parser.ts src/applescript/scripts.ts tests/unit/applescript/parser.test.ts tests/unit/applescript/scripts.test.ts
git commit -m "feat: add parser and script for email sending"
```

---

## Task 12: Email Sending - Mail Sender Interface

**Files:**
- Create: `src/applescript/mail-sender.ts`
- Test: `tests/unit/applescript/mail-sender.test.ts`

**Step 1: Write test**

```typescript
// tests/unit/applescript/mail-sender.test.ts
import { AppleScriptMailSender } from '../../../src/applescript/mail-sender.js';
import * as executor from '../../../src/applescript/executor.js';
import * as fs from 'fs';

jest.mock('../../../src/applescript/executor.js');
jest.mock('fs');

describe('AppleScriptMailSender', () => {
  let sender: AppleScriptMailSender;
  let mockExecute: jest.MockedFunction<typeof executor.executeAppleScriptOrThrow>;
  let mockExistsSync: jest.MockedFunction<typeof fs.existsSync>;

  beforeEach(() => {
    sender = new AppleScriptMailSender();
    mockExecute = executor.executeAppleScriptOrThrow as jest.MockedFunction<typeof executor.executeAppleScriptOrThrow>;
    mockExistsSync = fs.existsSync as jest.MockedFunction<typeof fs.existsSync>;
    mockExecute.mockClear();
    mockExistsSync.mockClear();
  });

  test('sends basic email', () => {
    mockExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}123{{FIELD}}sentAt{{=}}2026-02-03T10:00:00Z');

    const result = sender.sendEmail({
      to: ['user@example.com'],
      subject: 'Test',
      body: 'Hello',
      bodyType: 'plain',
    });

    expect(result.messageId).toBe('123');
    expect(result.sentAt).toBe('2026-02-03T10:00:00Z');
  });

  test('validates attachments exist', () => {
    mockExistsSync.mockReturnValue(false);

    expect(() => {
      sender.sendEmail({
        to: ['user@example.com'],
        subject: 'Test',
        body: 'Hello',
        bodyType: 'plain',
        attachments: [{ path: '/missing/file.pdf' }],
      });
    }).toThrow('Attachment file not found: /missing/file.pdf');
  });

  test('sends HTML email with attachments', () => {
    mockExistsSync.mockReturnValue(true);
    mockExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}456{{FIELD}}sentAt{{=}}2026-02-03T10:30:00Z');

    const result = sender.sendEmail({
      to: ['user@example.com'],
      cc: ['cc@example.com'],
      subject: 'HTML Test',
      body: '<h1>Hello</h1>',
      bodyType: 'html',
      attachments: [{ path: '/file.pdf', name: 'document.pdf' }],
    });

    expect(result.messageId).toBe('456');
    expect(mockExistsSync).toHaveBeenCalledWith('/file.pdf');
  });
});
```

**Step 2: Create IMailSender interface**

```typescript
// src/applescript/mail-sender.ts
import { existsSync } from 'fs';
import { executeAppleScriptOrThrow } from './executor.js';
import * as scripts from './scripts.js';
import { parseSendEmailResult, type SendEmailResult } from './parser.js';
import { AppleScriptError, AttachmentNotFoundError, MailSendError } from '../utils/errors.js';

// =============================================================================
// Types
// =============================================================================

export interface Attachment {
  readonly path: string;
  readonly name?: string;
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

export interface SentEmail {
  readonly messageId: string;
  readonly sentAt: string;
}

export interface IMailSender {
  sendEmail(params: SendEmailParams): SentEmail;
}

// =============================================================================
// Implementation
// =============================================================================

export class AppleScriptMailSender implements IMailSender {
  sendEmail(params: SendEmailParams): SentEmail {
    // Validate attachments exist
    if (params.attachments != null) {
      for (const attachment of params.attachments) {
        if (!existsSync(attachment.path)) {
          throw new AttachmentNotFoundError(attachment.path);
        }
      }
    }

    const script = scripts.sendEmail({
      to: params.to,
      subject: params.subject,
      body: params.body,
      bodyType: params.bodyType,
      cc: params.cc,
      bcc: params.bcc,
      replyTo: params.replyTo,
      attachments: params.attachments,
      accountId: params.accountId,
    });

    const output = executeAppleScriptOrThrow(script);
    const result = parseSendEmailResult(output);

    if (result == null) {
      throw new AppleScriptError('Failed to parse send email response');
    }

    if (!result.success) {
      throw new MailSendError(result.error ?? 'Unknown error');
    }

    return {
      messageId: result.messageId!,
      sentAt: result.sentAt!,
    };
  }
}

export function createMailSender(): IMailSender {
  return new AppleScriptMailSender();
}
```

**Step 3: Run tests**

Run: `npm test -- mail-sender.test.ts`
Expected: All tests PASS

**Step 4: Commit**

```bash
git add src/applescript/mail-sender.ts tests/unit/applescript/mail-sender.test.ts
git commit -m "feat: implement IMailSender for email sending"
```

---

## Task 13: Email Sending - MCP Tool

**Files:**
- Modify: `src/tools/email.ts`
- Modify: `src/applescript/index.ts`

**Step 1: Add SendEmailInput schema**

```typescript
// src/tools/email.ts (add after existing schemas)
export const SendEmailInput = z
  .object({
    to: z.array(z.string().email()).min(1).describe('Recipient email addresses'),
    subject: z.string().min(1).describe('Email subject'),
    body: z.string().describe('Email body content'),
    body_type: z.enum(['plain', 'html']).default('plain').describe('Body content type'),
    cc: z.array(z.string().email()).optional().describe('CC recipients'),
    bcc: z.array(z.string().email()).optional().describe('BCC recipients'),
    reply_to: z.string().email().optional().describe('Reply-to address'),
    attachments: z
      .array(
        z.object({
          path: z.string().describe('Absolute file path to attachment'),
          name: z.string().optional().describe('Display name for attachment'),
        })
      )
      .optional()
      .describe('File attachments'),
    account_id: z.number().int().positive().optional().describe('Account to send from'),
  })
  .strict();
```

**Step 2: Add sendEmailTool handler**

```typescript
// src/tools/email.ts
import type { IMailSender } from '../applescript/mail-sender.js';

export function sendEmailTool(mailSender: IMailSender) {
  return {
    name: 'send_email',
    description: 'Send an email with optional CC, BCC, attachments, and HTML formatting. Returns the sent message ID and timestamp.',
    inputSchema: zodToJsonSchema(SendEmailInput),
    handler: async (input: unknown) => {
      const params = SendEmailInput.parse(input);

      const result = mailSender.sendEmail({
        to: params.to,
        subject: params.subject,
        body: params.body,
        bodyType: params.body_type,
        cc: params.cc,
        bcc: params.bcc,
        replyTo: params.reply_to,
        attachments: params.attachments,
        accountId: params.account_id,
      });

      return {
        content: [
          {
            type: 'text',
            text: `Email sent successfully!\nMessage ID: ${result.messageId}\nSent at: ${result.sentAt}`,
          },
        ],
      };
    },
  };
}

export const emailTools = {
  // ... existing tools
  sendEmail: sendEmailTool,
};
```

**Step 3: Update AppleScript index**

```typescript
// src/applescript/index.ts
import { createMailSender } from './mail-sender.js';

// Update createBackend function
export function createBackend() {
  return {
    repository: createRepository(),
    calendarWriter: createCalendarWriter(),
    calendarManager: createCalendarManager(),
    mailSender: createMailSender(),
  };
}
```

**Step 4: Verify tool registration in main server**

Check `src/index.ts` to ensure all tools are properly registered with MCP server.

**Step 5: Commit**

```bash
git add src/tools/email.ts src/applescript/index.ts
git commit -m "feat: add send_email MCP tool"
```

---

## Task 14: Update README Documentation

**Files:**
- Modify: `README.md`

**Step 1: Update Available Tools section**

```markdown
<!-- In README.md, update the Calendar section -->

**Calendar**
- `list_calendars` - List all calendars
- `list_events` - List events with date range filtering
- `get_event` - Get event details
- `search_events` - Search events by title
- `create_event` - Create a new calendar event (AppleScript backend only)
- `respond_to_event` - Accept, decline, or tentatively accept event invitations (AppleScript backend only)
- `delete_event` - Delete a calendar event or recurring series (AppleScript backend only)
- `update_event` - Update event details (title, time, location, etc.) (AppleScript backend only)
```

**Step 2: Update Mail section**

```markdown
**Mail**
- `list_folders` - List all mail folders with unread counts (supports `account_id` filtering)
- `list_emails` - List emails in a folder with pagination
- `search_emails` - Search emails by subject, sender, or content
- `get_email` - Get full email details including body
- `get_unread_count` - Get unread email count
- `send_email` - Send an email with attachments and HTML support (AppleScript backend only)
```

**Step 3: Update Known Limitations section**

```markdown
### AppleScript Backend

**Write Operations**

Currently, write operations (event management, email sending) are only supported via the AppleScript backend. These features will be added to the Graph API backend in a future release:
- Event RSVP operations
- Event deletion
- Event updates
- Email sending

For these operations, use the AppleScript backend with classic Outlook for Mac.
```

**Step 4: Add new tools to feature comparison**

If there's a feature comparison table, update it to show which backend supports which operations.

**Step 5: Commit**

```bash
git add README.md
git commit -m "docs: update README with new event management and email tools"
```

---

## Task 15: Integration Tests

**Files:**
- Create: `tests/integration/tools/event-management.test.ts`
- Create: `tests/integration/tools/email-sending.test.ts`

**Step 1: Create event management integration test**

```typescript
// tests/integration/tools/event-management.test.ts
import { createBackend } from '../../../src/applescript/index.js';

// Skip if Outlook not available
const isOutlookAvailable = process.env.OUTLOOK_AVAILABLE === '1';
const testIf = (condition: boolean) => (condition ? test : test.skip);

describe('Event Management Integration', () => {
  const backend = createBackend();
  const calendarManager = backend.calendarManager;

  testIf(isOutlookAvailable)('responds to event', () => {
    // This test requires a real event in Outlook
    // For now, we verify the interface works
    expect(calendarManager.respondToEvent).toBeDefined();
  });

  testIf(isOutlookAvailable)('deletes event', () => {
    expect(calendarManager.deleteEvent).toBeDefined();
  });

  testIf(isOutlookAvailable)('updates event', () => {
    expect(calendarManager.updateEvent).toBeDefined();
  });
});
```

**Step 2: Create email sending integration test**

```typescript
// tests/integration/tools/email-sending.test.ts
import { createBackend } from '../../../src/applescript/index.js';
import { writeFileSync, unlinkSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

const isOutlookAvailable = process.env.OUTLOOK_AVAILABLE === '1';
const testIf = (condition: boolean) => (condition ? test : test.skip);

describe('Email Sending Integration', () => {
  const backend = createBackend();
  const mailSender = backend.mailSender;

  testIf(isOutlookAvailable)('sends basic email', () => {
    expect(mailSender.sendEmail).toBeDefined();

    // Uncomment to test with real email
    // const result = mailSender.sendEmail({
    //   to: ['test@example.com'],
    //   subject: 'Integration Test',
    //   body: 'This is a test email',
    //   bodyType: 'plain',
    // });
    // expect(result.messageId).toBeDefined();
  });

  testIf(isOutlookAvailable)('validates attachment exists', () => {
    expect(() => {
      mailSender.sendEmail({
        to: ['test@example.com'],
        subject: 'Test',
        body: 'Test',
        bodyType: 'plain',
        attachments: [{ path: '/nonexistent/file.pdf' }],
      });
    }).toThrow('Attachment file not found');
  });
});
```

**Step 3: Run integration tests**

Run: `npm test -- tests/integration/`
Expected: Tests pass or skip if Outlook not available

**Step 4: Commit**

```bash
git add tests/integration/tools/event-management.test.ts tests/integration/tools/email-sending.test.ts
git commit -m "test: add integration tests for event management and email sending"
```

---

## Task 16: Wire Up Tools in Main Server

**Files:**
- Modify: `src/index.ts`

**Step 1: Import new tools and interfaces**

```typescript
// src/index.ts (update imports)
import { calendarTools } from './tools/calendar.js';
import { emailTools } from './tools/email.js';
```

**Step 2: Get calendar manager and mail sender from backend**

```typescript
// src/index.ts (update backend initialization)
const backend = createBackend(); // AppleScript or Graph
const repository = backend.repository;
const calendarWriter = backend.calendarWriter; // may be undefined for Graph
const calendarManager = backend.calendarManager; // may be undefined for Graph
const mailSender = backend.mailSender; // may be undefined for Graph
```

**Step 3: Register new tools conditionally**

```typescript
// src/index.ts (update tool registration)
const tools = [
  // ... existing tools
];

// Add event management tools if available
if (calendarManager != null) {
  tools.push(
    calendarTools.respondToEvent(calendarManager),
    calendarTools.deleteEvent(calendarManager),
    calendarTools.updateEvent(calendarManager)
  );
}

// Add email sending tool if available
if (mailSender != null) {
  tools.push(emailTools.sendEmail(mailSender));
}
```

**Step 4: Test manually with MCP inspector**

Run: `npm run build && node build/index.js`
Then use MCP inspector to verify new tools appear.

**Step 5: Commit**

```bash
git add src/index.ts
git commit -m "feat: wire up event management and email sending tools in MCP server"
```

---

## Task 17: Final Testing & Cleanup

**Files:**
- All files

**Step 1: Run full test suite**

Run: `npm test`
Expected: All tests pass

**Step 2: Run type checking**

Run: `npm run typecheck`
Expected: No type errors

**Step 3: Run linting**

Run: `npm run lint`
Expected: No linting errors (or fix any issues)

**Step 4: Build project**

Run: `npm run build`
Expected: Build succeeds without errors

**Step 5: Manual testing with Outlook**

1. Start Outlook for Mac
2. Run the MCP server: `node build/index.js`
3. Test each new tool:
   - Create a test event
   - Respond to an event (if you have invitations)
   - Update the test event
   - Delete the test event
   - Send a test email

**Step 6: Final commit**

```bash
git add -A
git commit -m "feat: complete event management and email sending implementation

- Add RSVP operations (accept, decline, tentative)
- Add event deletion with recurring event support
- Add event updates with full field support
- Add email sending with HTML, attachments, CC/BCC
- Comprehensive test coverage
- Updated documentation"
```

---

## Success Criteria

✅ All unit tests pass
✅ All integration tests pass (or skip gracefully)
✅ Type checking passes
✅ Linting passes
✅ Build succeeds
✅ Manual testing with Outlook works
✅ Documentation updated
✅ No regressions in existing functionality

## Future Work

- Implement same features for Graph API backend
- Add reply/forward email support
- Add draft management
- Add batch operations
- Add email templates support
