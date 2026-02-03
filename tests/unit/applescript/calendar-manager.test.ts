import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('../../../src/applescript/executor.js', () => ({
  executeAppleScriptOrThrow: vi.fn(),
  escapeForAppleScript: (s: string) => s.replace(/\\/g, '\\\\').replace(/"/g, '\\"'),
}));

import { AppleScriptCalendarManager } from '../../../src/applescript/calendar-manager.js';
import { executeAppleScriptOrThrow } from '../../../src/applescript/executor.js';

const mockedExecute = vi.mocked(executeAppleScriptOrThrow);

describe('AppleScriptCalendarManager', () => {
  let manager: AppleScriptCalendarManager;

  beforeEach(() => {
    vi.clearAllMocks();
    manager = new AppleScriptCalendarManager();
  });

  describe('respondToEvent', () => {
    it('accepts event with comment', () => {
      mockedExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}123');

      const result = manager.respondToEvent(123, 'accept', true, 'Looking forward to it');

      expect(result).toEqual({ success: true, eventId: 123 });
      expect(mockedExecute).toHaveBeenCalledOnce();
      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('calendar event id 123');
      expect(script).toContain('accept');
      expect(script).toContain('Looking forward to it');
    });

    it('declines event without sending', () => {
      mockedExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}456');

      const result = manager.respondToEvent(456, 'decline', false);

      expect(result).toEqual({ success: true, eventId: 456 });
    });

    it('handles AppleScript error', () => {
      mockedExecute.mockReturnValue('{{RECORD}}success{{=}}false{{FIELD}}error{{=}}Permission denied');

      expect(() => {
        manager.respondToEvent(789, 'accept', true);
      }).toThrow('Permission denied');
    });

    it('throws error when parser returns null', () => {
      mockedExecute.mockReturnValue('invalid output');

      expect(() => {
        manager.respondToEvent(123, 'accept', true);
      }).toThrow('Failed to parse RSVP response');
    });
  });

  describe('deleteEvent', () => {
    it('deletes single instance', () => {
      mockedExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}123');

      manager.deleteEvent(123, 'this_instance');

      expect(mockedExecute).toHaveBeenCalledOnce();
      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('calendar event id 123');
      expect(script).toContain('delete');
    });

    it('deletes all in series', () => {
      mockedExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}456');

      manager.deleteEvent(456, 'all_in_series');

      expect(mockedExecute).toHaveBeenCalledOnce();
      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('calendar event id 456');
    });

    it('throws on failure', () => {
      mockedExecute.mockReturnValue('{{RECORD}}success{{=}}false{{FIELD}}error{{=}}Not found');

      expect(() => {
        manager.deleteEvent(789, 'this_instance');
      }).toThrow('Not found');
    });

    it('throws when parser returns null', () => {
      mockedExecute.mockReturnValue('invalid output');

      expect(() => {
        manager.deleteEvent(123, 'this_instance');
      }).toThrow('Failed to parse delete response');
    });
  });

  describe('updateEvent', () => {
    it('updates event title only', () => {
      mockedExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}123{{FIELD}}updatedFields{{=}}title');

      const result = manager.updateEvent(123, { title: 'Updated Meeting' }, 'this_instance');

      expect(result).toEqual({ id: 123, updatedFields: ['title'] });
      expect(mockedExecute).toHaveBeenCalledOnce();
      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('calendar event id 123');
      expect(script).toContain('set subject of myEvent to "Updated Meeting"');
    });

    it('updates multiple fields', () => {
      mockedExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}456{{FIELD}}updatedFields{{=}}title,location,description');

      const result = manager.updateEvent(
        456,
        {
          title: 'New Title',
          location: 'New Location',
          description: 'New Description',
        },
        'this_instance'
      );

      expect(result.id).toBe(456);
      expect(result.updatedFields).toEqual(['title', 'location', 'description']);
      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('set subject of myEvent to "New Title"');
      expect(script).toContain('set location of myEvent to "New Location"');
      expect(script).toContain('set content of myEvent to "New Description"');
    });

    it('updates dates', () => {
      mockedExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}789{{FIELD}}updatedFields{{=}}startDate,endDate');

      const result = manager.updateEvent(
        789,
        {
          startDate: '2026-02-03T10:00:00Z',
          endDate: '2026-02-03T11:00:00Z',
        },
        'this_instance'
      );

      expect(result.updatedFields).toEqual(['startDate', 'endDate']);
      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('set start time of myEvent to date');
      expect(script).toContain('set end time of myEvent to date');
    });

    it('updates all day flag', () => {
      mockedExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}100{{FIELD}}updatedFields{{=}}isAllDay');

      const result = manager.updateEvent(100, { isAllDay: true }, 'this_instance');

      expect(result.updatedFields).toEqual(['isAllDay']);
      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('set all day flag of myEvent to true');
    });

    it('updates entire series', () => {
      mockedExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}200{{FIELD}}updatedFields{{=}}title');

      manager.updateEvent(200, { title: 'Updated Series' }, 'all_in_series');

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('-- Updating entire series');
    });

    it('handles no fields updated', () => {
      mockedExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}300{{FIELD}}updatedFields{{=}}');

      const result = manager.updateEvent(300, {}, 'this_instance');

      expect(result.id).toBe(300);
      expect(result.updatedFields).toEqual([]);
    });

    it('throws on failure', () => {
      mockedExecute.mockReturnValue('{{RECORD}}success{{=}}false{{FIELD}}error{{=}}Event not found');

      expect(() => {
        manager.updateEvent(999, { title: 'Test' }, 'this_instance');
      }).toThrow('Event not found');
    });

    it('throws when parser returns null', () => {
      mockedExecute.mockReturnValue('invalid output');

      expect(() => {
        manager.updateEvent(123, { title: 'Test' }, 'this_instance');
      }).toThrow('Failed to parse update response');
    });

    it('escapes special characters', () => {
      mockedExecute.mockReturnValue('{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}123{{FIELD}}updatedFields{{=}}title');

      manager.updateEvent(123, { title: 'Test "quotes" and \\backslash' }, 'this_instance');

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('\\"quotes\\"');
      expect(script).toContain('\\\\backslash');
    });
  });
});
