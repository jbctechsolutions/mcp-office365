/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Graph event mapper functions.
 */

import { describe, it, expect } from 'vitest';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { mapEventToEventRow } from '../../../../src/graph/mappers/event-mapper.js';
import { hashStringToNumber } from '../../../../src/graph/mappers/utils.js';

describe('graph/mappers/event-mapper', () => {
  describe('mapEventToEventRow', () => {
    it('maps event with all fields', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
        start: { dateTime: '2024-01-15T10:00:00', timeZone: 'UTC' },
        end: { dateTime: '2024-01-15T11:00:00', timeZone: 'UTC' },
        recurrence: {
          pattern: { type: 'daily', interval: 1 },
          range: { type: 'noEnd', startDate: '2024-01-15' },
        },
        isReminderOn: true,
        attendees: [
          { emailAddress: { address: 'a@example.com' } },
          { emailAddress: { address: 'b@example.com' } },
        ],
        iCalUId: 'ical-uid-123',
      };

      const result = mapEventToEventRow(event, 'calendar-456');

      expect(result.id).toBe(hashStringToNumber('event-123'));
      expect(result.folderId).toBe(hashStringToNumber('calendar-456'));
      expect(result.isRecurring).toBe(1);
      expect(result.hasReminder).toBe(1);
      expect(result.attendeeCount).toBe(2);
      expect(result.uid).toBe('ical-uid-123');
      expect(result.dataFilePath).toBe('graph-event:event-123');
    });

    it('handles event with null id', () => {
      const event: MicrosoftGraph.Event = {
        id: undefined,
      };

      const result = mapEventToEventRow(event);

      expect(result.id).toBe(hashStringToNumber(''));
      expect(result.dataFilePath).toBe('graph-event:');
    });

    it('handles event without calendarId', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
      };

      const result = mapEventToEventRow(event);

      expect(result.folderId).toBe(0);
    });

    it('handles non-recurring event', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
        recurrence: undefined,
      };

      const result = mapEventToEventRow(event);

      expect(result.isRecurring).toBe(0);
    });

    it('handles event with null recurrence', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
        recurrence: null,
      };

      const result = mapEventToEventRow(event);

      expect(result.isRecurring).toBe(0);
    });

    it('handles event without reminder', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
        isReminderOn: false,
      };

      const result = mapEventToEventRow(event);

      expect(result.hasReminder).toBe(0);
    });

    it('handles event with undefined isReminderOn', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
        isReminderOn: undefined,
      };

      const result = mapEventToEventRow(event);

      expect(result.hasReminder).toBe(0);
    });

    it('handles event without attendees', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
        attendees: undefined,
      };

      const result = mapEventToEventRow(event);

      expect(result.attendeeCount).toBe(0);
    });

    it('handles event with empty attendees array', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
        attendees: [],
      };

      const result = mapEventToEventRow(event);

      expect(result.attendeeCount).toBe(0);
    });

    it('handles event without iCalUId', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
        iCalUId: undefined,
      };

      const result = mapEventToEventRow(event);

      expect(result.uid).toBeNull();
    });

    it('handles event without start date', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
        start: undefined,
      };

      const result = mapEventToEventRow(event);

      expect(result.startDate).toBeNull();
    });

    it('handles event without end date', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
        end: undefined,
      };

      const result = mapEventToEventRow(event);

      expect(result.endDate).toBeNull();
    });

    it('parses start and end dates correctly', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
        start: { dateTime: '2024-01-15T10:00:00Z', timeZone: 'UTC' },
        end: { dateTime: '2024-01-15T11:00:00Z', timeZone: 'UTC' },
      };

      const result = mapEventToEventRow(event);

      expect(result.startDate).toBeTypeOf('number');
      expect(result.endDate).toBeTypeOf('number');
      expect(result.endDate!).toBeGreaterThan(result.startDate!);
    });

    it('sets masterRecordId to null', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
      };

      const result = mapEventToEventRow(event);

      expect(result.masterRecordId).toBeNull();
    });

    it('sets recurrenceId to null', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
      };

      const result = mapEventToEventRow(event);

      expect(result.recurrenceId).toBeNull();
    });

    it('includes subject from event', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
        subject: 'Team Standup',
      };

      const result = mapEventToEventRow(event);

      expect(result.subject).toBe('Team Standup');
    });

    it('sets subject to null when event has no subject', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
        subject: undefined,
      };

      const result = mapEventToEventRow(event);

      expect(result.subject).toBeNull();
    });

    it('extracts onlineMeetingUrl from onlineMeeting joinUrl', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-online',
        onlineMeeting: {
          joinUrl: 'https://teams.microsoft.com/l/meetup-join/abc123',
        },
      };

      const result = mapEventToEventRow(event);

      expect(result.onlineMeetingUrl).toBe('https://teams.microsoft.com/l/meetup-join/abc123');
    });

    it('sets onlineMeetingUrl to null when onlineMeeting is undefined', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-no-online',
        onlineMeeting: undefined,
      };

      const result = mapEventToEventRow(event);

      expect(result.onlineMeetingUrl).toBeNull();
    });

    it('sets onlineMeetingUrl to null when onlineMeeting has no joinUrl', () => {
      const event: MicrosoftGraph.Event = {
        id: 'event-no-joinurl',
        onlineMeeting: {},
      };

      const result = mapEventToEventRow(event);

      expect(result.onlineMeetingUrl).toBeNull();
    });

    it('stores startDate as Unix timestamp (not Apple timestamp)', () => {
      // Graph API returns dateTime without Z suffix, with separate timeZone
      // "2024-06-15T10:00:00.0000000" in UTC = Unix timestamp 1718442000
      const event: MicrosoftGraph.Event = {
        id: 'event-123',
        start: { dateTime: '2024-06-15T10:00:00.0000000', timeZone: 'UTC' },
      };

      const result = mapEventToEventRow(event);

      // dateTimeTimeZoneToTimestamp returns Unix seconds since 1970
      // The value should be a reasonable Unix timestamp (year 2024 range)
      // NOT an Apple timestamp (which would be ~31 years less)
      expect(result.startDate).toBeTypeOf('number');
      // Year 2024 Unix timestamps are around 1.7 billion
      expect(result.startDate!).toBeGreaterThan(1_700_000_000);
      expect(result.startDate!).toBeLessThan(1_800_000_000);
    });
  });
});
