/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for the AppleScript calendar backend's id guard (U5b). The shared
 * event_id schema is a string|number union for durable Graph tokens, but the
 * AppleScript write path interpolates the id into osascript, so a non-numeric
 * id must be rejected before it reaches the interpreter.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { AppleCalendarTools } from '../../../src/tools/calendar-apple.js';
import { ValidationError } from '../../../src/utils/errors.js';
import type { CalendarTools } from '../../../src/tools/calendar.js';
import type { ICalendarWriter, ICalendarManager } from '../../../src/applescript/index.js';

describe('AppleCalendarTools — event id guard', () => {
  let manager: ICalendarManager;
  let tools: AppleCalendarTools;

  beforeEach(() => {
    manager = {
      respondToEvent: vi.fn(() => ({ eventId: 1 })),
      deleteEvent: vi.fn(),
      updateEvent: vi.fn(() => ({})),
    } as unknown as ICalendarManager;
    tools = new AppleCalendarTools(
      {} as unknown as CalendarTools,
      {} as unknown as ICalendarWriter,
      manager,
    );
  });

  it('accepts a numeric event id and forwards it to the manager', () => {
    tools.deleteEvent({ event_id: 42, apply_to: 'this_instance' } as never);
    expect(manager.deleteEvent).toHaveBeenCalledWith(42, 'this_instance');
  });

  it('coerces a numeric-string event id', () => {
    tools.deleteEvent({ event_id: '42', apply_to: 'this_instance' } as never);
    expect(manager.deleteEvent).toHaveBeenCalledWith(42, 'this_instance');
  });

  it('rejects a durable ev_ token before it reaches osascript interpolation', () => {
    expect(() => tools.deleteEvent({ event_id: 'ev_QWxpY2U' } as never)).toThrow(ValidationError);
    expect(manager.deleteEvent).not.toHaveBeenCalled();
  });

  it('rejects a newline-bearing string (injection guard)', () => {
    expect(() =>
      tools.updateEvent({ event_id: '1\ndelete calendar 1' } as never),
    ).toThrow(ValidationError);
    expect(manager.updateEvent).not.toHaveBeenCalled();
  });

  it('rejects a non-positive id', () => {
    expect(() => tools.respondToEvent({ event_id: 0, response: 'accept', send_response: true } as never)).toThrow(ValidationError);
  });
});
