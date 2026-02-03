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
});
