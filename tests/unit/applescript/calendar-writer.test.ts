import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('../../../src/applescript/executor.js', () => ({
  executeAppleScriptOrThrow: vi.fn(),
  escapeForAppleScript: (s: string) => s.replace(/\\/g, '\\\\').replace(/"/g, '\\"'),
}));

import { AppleScriptCalendarWriter } from '../../../src/applescript/calendar-writer.js';
import { executeAppleScriptOrThrow } from '../../../src/applescript/executor.js';

const mockedExecute = vi.mocked(executeAppleScriptOrThrow);

describe('AppleScriptCalendarWriter', () => {
  let writer: AppleScriptCalendarWriter;

  beforeEach(() => {
    vi.clearAllMocks();
    writer = new AppleScriptCalendarWriter();
  });

  describe('createEvent', () => {
    it('creates an event and returns the ID and calendar ID', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}12345{{FIELD}}calendarId{{=}}67');

      const result = writer.createEvent({
        title: 'Test Meeting',
        startDate: '2026-02-03T14:00:00Z',
        endDate: '2026-02-03T15:00:00Z',
      });

      expect(result.id).toBe(12345);
      expect(result.calendarId).toBe(67);
      expect(mockedExecute).toHaveBeenCalledOnce();
    });

    it('includes title in the generated script', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'Team Standup',
        startDate: '2026-02-03T09:00:00Z',
        endDate: '2026-02-03T09:30:00Z',
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('subject:"Team Standup"');
    });

    it('targets specific calendar when calendar_id is provided', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}200');

      writer.createEvent({
        title: 'Test',
        startDate: '2026-02-03T14:00:00Z',
        endDate: '2026-02-03T15:00:00Z',
        calendarId: 200,
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('at calendar id 200');
    });

    it('does not include calendar clause when no calendar_id', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'Test',
        startDate: '2026-02-03T14:00:00Z',
        endDate: '2026-02-03T15:00:00Z',
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).not.toContain('at calendar id');
    });

    it('includes location when provided', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'Test',
        startDate: '2026-02-03T14:00:00Z',
        endDate: '2026-02-03T15:00:00Z',
        location: 'Room 101',
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('location:"Room 101"');
    });

    it('includes description when provided', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'Test',
        startDate: '2026-02-03T14:00:00Z',
        endDate: '2026-02-03T15:00:00Z',
        description: 'Weekly sync meeting',
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('plain text content of newEvent to "Weekly sync meeting"');
    });

    it('sets all day flag when isAllDay is true', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'Day Off',
        startDate: '2026-02-03T00:00:00Z',
        endDate: '2026-02-04T00:00:00Z',
        isAllDay: true,
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('all day flag:true');
    });

    it('does not set all day flag when isAllDay is false', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'Test',
        startDate: '2026-02-03T14:00:00Z',
        endDate: '2026-02-03T15:00:00Z',
        isAllDay: false,
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).not.toContain('all day flag');
    });

    it('uses locale-safe date components in the script', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'Test',
        startDate: '2026-06-15T09:30:00Z',
        endDate: '2026-06-15T10:30:00Z',
      });

      const script = mockedExecute.mock.calls[0]![0];
      // Should use component-based date setting, not string-based
      expect(script).toContain('set year of theStartDate to');
      expect(script).toContain('set month of theStartDate to');
      expect(script).toContain('set day of theStartDate to');
      expect(script).toContain('set hours of theStartDate to');
      expect(script).toContain('set minutes of theStartDate to');
    });

    it('throws when parse returns null (empty output)', () => {
      mockedExecute.mockReturnValue('');

      expect(() => writer.createEvent({
        title: 'Test',
        startDate: '2026-02-03T14:00:00Z',
        endDate: '2026-02-03T15:00:00Z',
      })).toThrow('Failed to parse create event result');
    });

    it('returns null calendarId when not present in output', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      const result = writer.createEvent({
        title: 'Test',
        startDate: '2026-02-03T14:00:00Z',
        endDate: '2026-02-03T15:00:00Z',
      });

      expect(result.id).toBe(100);
      expect(result.calendarId).toBeNull();
    });

    it('escapes special characters in title', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'Meeting with "quotes" and \\backslash',
        startDate: '2026-02-03T14:00:00Z',
        endDate: '2026-02-03T15:00:00Z',
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('\\"quotes\\"');
      expect(script).toContain('\\\\backslash');
    });
  });

  describe('createEvent with recurrence', () => {
    it('generates daily recurrence script', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'Daily Standup',
        startDate: '2026-02-03T09:00:00Z',
        endDate: '2026-02-03T09:30:00Z',
        recurrence: { frequency: 'daily', interval: 1 },
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('set is recurring of newEvent to true');
      expect(script).toContain('set recurrence type of theRecurrence to daily recurrence');
      expect(script).toContain('set occurrence interval of theRecurrence to 1');
    });

    it('generates weekly recurrence with days of week', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'MWF Meeting',
        startDate: '2026-02-03T10:00:00Z',
        endDate: '2026-02-03T10:30:00Z',
        recurrence: {
          frequency: 'weekly',
          interval: 1,
          daysOfWeek: ['monday', 'wednesday', 'friday'],
        },
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('set recurrence type of theRecurrence to weekly recurrence');
      expect(script).toContain('set day of week mask of theRecurrence to {Monday, Wednesday, Friday}');
    });

    it('generates monthly by date recurrence', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'Monthly Review',
        startDate: '2026-02-15T14:00:00Z',
        endDate: '2026-02-15T15:00:00Z',
        recurrence: { frequency: 'monthly', interval: 1, dayOfMonth: 15 },
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('set recurrence type of theRecurrence to monthly recurrence');
      expect(script).toContain('set day of month of theRecurrence to 15');
    });

    it('generates monthly ordinal recurrence (3rd Thursday)', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'Board Meeting',
        startDate: '2026-02-19T14:00:00Z',
        endDate: '2026-02-19T16:00:00Z',
        recurrence: {
          frequency: 'monthly',
          interval: 1,
          weekOfMonth: 'third',
          dayOfWeekMonthly: 'thursday',
        },
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('set recurrence type of theRecurrence to month nth recurrence');
      expect(script).toContain('set day of week mask of theRecurrence to {Thursday}');
      expect(script).toContain('set instance of theRecurrence to 3');
    });

    it('generates end after count', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'Limited Series',
        startDate: '2026-02-03T09:00:00Z',
        endDate: '2026-02-03T09:30:00Z',
        recurrence: { frequency: 'daily', interval: 1, endAfterCount: 10 },
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('set occurrences of theRecurrence to 10');
    });

    it('generates end date with component-based date construction', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'Seasonal Event',
        startDate: '2026-02-03T09:00:00Z',
        endDate: '2026-02-03T09:30:00Z',
        recurrence: {
          frequency: 'weekly',
          interval: 1,
          daysOfWeek: ['tuesday'],
          endDate: '2026-06-30T00:00:00Z',
        },
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('set theEndRecurrenceDate to current date');
      expect(script).toContain('set year of theEndRecurrenceDate to');
      expect(script).toContain('set pattern end date of theRecurrence to theEndRecurrenceDate');
    });

    it('does not include recurrence block when recurrence is omitted', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'One-time Event',
        startDate: '2026-02-03T14:00:00Z',
        endDate: '2026-02-03T15:00:00Z',
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).not.toContain('set is recurring of newEvent to true');
      expect(script).not.toContain('theRecurrence');
    });

    it('uses correct interval value', () => {
      mockedExecute.mockReturnValue('{{RECORD}}id{{=}}100{{FIELD}}calendarId{{=}}');

      writer.createEvent({
        title: 'Bi-weekly Sync',
        startDate: '2026-02-03T10:00:00Z',
        endDate: '2026-02-03T10:30:00Z',
        recurrence: {
          frequency: 'weekly',
          interval: 2,
          daysOfWeek: ['monday'],
        },
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('set occurrence interval of theRecurrence to 2');
    });
  });
});
