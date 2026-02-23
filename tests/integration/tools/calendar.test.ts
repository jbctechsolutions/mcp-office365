/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { createTestDatabase, SAMPLE_COUNTS } from '../../fixtures/database.js';
import { createConnection, type IConnection } from '../../../src/database/connection.js';
import { createRepository, type IRepository } from '../../../src/database/repository.js';
import {
  CalendarTools,
  createCalendarTools,
  ListCalendarsInput,
  ListEventsInput,
  GetEventInput,
  SearchEventsInput,
  CreateEventInput,
  RecurrenceInput,
  type IEventContentReader,
  type EventDetails,
} from '../../../src/tools/calendar.js';

describe('CalendarTools', () => {
  let testDb: { path: string; cleanup: () => void };
  let connection: IConnection;
  let repository: IRepository;
  let calendarTools: CalendarTools;

  beforeEach(() => {
    testDb = createTestDatabase();
    connection = createConnection(testDb.path);
    repository = createRepository(connection);
    calendarTools = createCalendarTools(repository);
  });

  afterEach(() => {
    connection.close();
    testDb.cleanup();
  });

  // ---------------------------------------------------------------------------
  // Input Validation
  // ---------------------------------------------------------------------------

  describe('input validation', () => {
    it('validates ListCalendarsInput', () => {
      expect(() => ListCalendarsInput.parse({})).not.toThrow();
    });

    it('validates ListEventsInput with defaults', () => {
      const parsed = ListEventsInput.parse({});
      expect(parsed.limit).toBe(50);
      expect(parsed.calendar_id).toBeUndefined();
    });

    it('validates ListEventsInput with all options', () => {
      const input = {
        calendar_id: 1,
        start_date: '2024-01-01T00:00:00Z',
        end_date: '2024-12-31T23:59:59Z',
        limit: 25,
      };
      const parsed = ListEventsInput.parse(input);
      expect(parsed).toEqual(input);
    });

    it('validates GetEventInput', () => {
      const parsed = GetEventInput.parse({ event_id: 1 });
      expect(parsed.event_id).toBe(1);
    });

    it('validates SearchEventsInput', () => {
      const parsed = SearchEventsInput.parse({ query: 'meeting' });
      expect(parsed.query).toBe('meeting');
      expect(parsed.limit).toBe(50);
    });

    it('validates CreateEventInput with required fields', () => {
      const parsed = CreateEventInput.parse({
        title: 'Meeting',
        start_date: '2026-02-03T14:00:00Z',
        end_date: '2026-02-03T15:00:00Z',
      });
      expect(parsed.title).toBe('Meeting');
      expect(parsed.is_all_day).toBe(false);
    });

    it('rejects CreateEventInput missing title', () => {
      expect(() => CreateEventInput.parse({
        start_date: '2026-02-03T14:00:00Z',
        end_date: '2026-02-03T15:00:00Z',
      })).toThrow();
    });

    it('validates CreateEventInput with all optional fields', () => {
      const parsed = CreateEventInput.parse({
        title: 'Meeting',
        start_date: '2026-02-03T14:00:00Z',
        end_date: '2026-02-03T15:00:00Z',
        calendar_id: 123,
        location: 'Room A',
        description: 'Weekly sync',
        is_all_day: true,
      });
      expect(parsed.calendar_id).toBe(123);
      expect(parsed.location).toBe('Room A');
      expect(parsed.description).toBe('Weekly sync');
      expect(parsed.is_all_day).toBe(true);
    });

    it('rejects CreateEventInput with unknown fields', () => {
      expect(() => CreateEventInput.parse({
        title: 'Meeting',
        start_date: '2026-02-03T14:00:00Z',
        end_date: '2026-02-03T15:00:00Z',
        unknown_field: 'value',
      })).toThrow();
    });

    it('rejects CreateEventInput with empty title', () => {
      expect(() => CreateEventInput.parse({
        title: '',
        start_date: '2026-02-03T14:00:00Z',
        end_date: '2026-02-03T15:00:00Z',
      })).toThrow();
    });

    it('rejects CreateEventInput with invalid start_date format', () => {
      expect(() => CreateEventInput.parse({
        title: 'Meeting',
        start_date: 'not-a-date',
        end_date: '2026-02-03T15:00:00Z',
      })).toThrow();
    });

    it('rejects CreateEventInput with invalid end_date format', () => {
      expect(() => CreateEventInput.parse({
        title: 'Meeting',
        start_date: '2026-02-03T14:00:00Z',
        end_date: 'not-a-date',
      })).toThrow();
    });

    it('rejects CreateEventInput when start_date is after end_date', () => {
      expect(() => CreateEventInput.parse({
        title: 'Meeting',
        start_date: '2026-02-03T16:00:00Z',
        end_date: '2026-02-03T15:00:00Z',
      })).toThrow();
    });

    it('rejects CreateEventInput when start_date equals end_date', () => {
      expect(() => CreateEventInput.parse({
        title: 'Meeting',
        start_date: '2026-02-03T14:00:00Z',
        end_date: '2026-02-03T14:00:00Z',
      })).toThrow();
    });
  });

  // ---------------------------------------------------------------------------
  // Recurrence Input Validation
  // ---------------------------------------------------------------------------

  describe('recurrence input validation', () => {
    it('accepts daily recurrence with defaults', () => {
      const parsed = RecurrenceInput.parse({ frequency: 'daily' });
      expect(parsed.frequency).toBe('daily');
      expect(parsed.interval).toBe(1);
      expect(parsed.end).toEqual({ type: 'no_end' });
    });

    it('accepts weekly recurrence with days_of_week', () => {
      const parsed = RecurrenceInput.parse({
        frequency: 'weekly',
        days_of_week: ['monday', 'wednesday', 'friday'],
      });
      expect(parsed.frequency).toBe('weekly');
      expect(parsed.days_of_week).toEqual(['monday', 'wednesday', 'friday']);
    });

    it('rejects weekly recurrence without days_of_week', () => {
      expect(() => RecurrenceInput.parse({ frequency: 'weekly' })).toThrow();
    });

    it('accepts monthly recurrence with day_of_month', () => {
      const parsed = RecurrenceInput.parse({
        frequency: 'monthly',
        day_of_month: 15,
      });
      expect(parsed.day_of_month).toBe(15);
    });

    it('accepts monthly ordinal recurrence', () => {
      const parsed = RecurrenceInput.parse({
        frequency: 'monthly',
        week_of_month: 'third',
        day_of_week_monthly: 'thursday',
      });
      expect(parsed.week_of_month).toBe('third');
      expect(parsed.day_of_week_monthly).toBe('thursday');
    });

    it('rejects incomplete ordinal monthly (only week_of_month)', () => {
      expect(() => RecurrenceInput.parse({
        frequency: 'monthly',
        week_of_month: 'third',
      })).toThrow();
    });

    it('rejects incomplete ordinal monthly (only day_of_week_monthly)', () => {
      expect(() => RecurrenceInput.parse({
        frequency: 'monthly',
        day_of_week_monthly: 'thursday',
      })).toThrow();
    });

    it('accepts yearly recurrence with interval', () => {
      const parsed = RecurrenceInput.parse({ frequency: 'yearly', interval: 2 });
      expect(parsed.frequency).toBe('yearly');
      expect(parsed.interval).toBe(2);
    });

    it('accepts end_date end condition', () => {
      const parsed = RecurrenceInput.parse({
        frequency: 'daily',
        end: { type: 'end_date', date: '2026-12-31T00:00:00Z' },
      });
      expect(parsed.end).toEqual({ type: 'end_date', date: '2026-12-31T00:00:00Z' });
    });

    it('accepts end_after_count end condition', () => {
      const parsed = RecurrenceInput.parse({
        frequency: 'daily',
        end: { type: 'end_after_count', count: 10 },
      });
      expect(parsed.end).toEqual({ type: 'end_after_count', count: 10 });
    });

    it('rejects days_of_week on non-weekly frequency', () => {
      expect(() => RecurrenceInput.parse({
        frequency: 'daily',
        days_of_week: ['monday'],
      })).toThrow();
    });

    it('rejects monthly-specific fields on non-monthly frequency', () => {
      expect(() => RecurrenceInput.parse({
        frequency: 'weekly',
        days_of_week: ['monday'],
        day_of_month: 15,
      })).toThrow();
    });

    it('accepts recurrence on CreateEventInput', () => {
      const parsed = CreateEventInput.parse({
        title: 'Weekly Standup',
        start_date: '2026-02-03T09:00:00Z',
        end_date: '2026-02-03T09:30:00Z',
        recurrence: {
          frequency: 'weekly',
          days_of_week: ['monday', 'wednesday', 'friday'],
        },
      });
      expect(parsed.recurrence).toBeDefined();
      expect(parsed.recurrence!.frequency).toBe('weekly');
    });
  });

  // ---------------------------------------------------------------------------
  // listCalendars
  // ---------------------------------------------------------------------------

  describe('listCalendars', () => {
    it('returns calendar folders', () => {
      const calendars = calendarTools.listCalendars({});
      expect(calendars.length).toBe(1);
      expect(calendars[0]?.name).toBe('Calendar');
    });

    it('returns calendars with correct structure', () => {
      const calendars = calendarTools.listCalendars({});
      const cal = calendars[0];

      expect(cal).toHaveProperty('id');
      expect(cal).toHaveProperty('name');
      expect(cal).toHaveProperty('accountId');
    });
  });

  // ---------------------------------------------------------------------------
  // listEvents
  // ---------------------------------------------------------------------------

  describe('listEvents', () => {
    it('returns events', () => {
      const events = calendarTools.listEvents({ limit: 50 });
      expect(events.length).toBe(SAMPLE_COUNTS.events);
    });

    it('returns events with correct structure', () => {
      const events = calendarTools.listEvents({ limit: 1 });
      const event = events[0];

      expect(event).toHaveProperty('id');
      expect(event).toHaveProperty('folderId');
      expect(event).toHaveProperty('startDate');
      expect(event).toHaveProperty('endDate');
      expect(event).toHaveProperty('isRecurring');
      expect(event).toHaveProperty('hasReminder');
      expect(event).toHaveProperty('attendeeCount');
      expect(typeof event?.isRecurring).toBe('boolean');
    });

    it('respects limit parameter', () => {
      const events = calendarTools.listEvents({ limit: 1 });
      expect(events.length).toBe(1);
    });

    it('converts timestamps to ISO format', () => {
      const events = calendarTools.listEvents({ limit: 1 });
      const event = events[0];

      if (event?.startDate) {
        expect(event.startDate).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // getEvent
  // ---------------------------------------------------------------------------

  describe('getEvent', () => {
    it('returns event by ID', () => {
      const events = calendarTools.listEvents({ limit: 1 });
      const firstEvent = events[0];

      if (firstEvent) {
        const event = calendarTools.getEvent({ event_id: firstEvent.id });
        expect(event).not.toBeNull();
        expect(event?.id).toBe(firstEvent.id);
      }
    });

    it('returns null for non-existent ID', () => {
      const event = calendarTools.getEvent({ event_id: 99999 });
      expect(event).toBeNull();
    });

    it('includes additional fields in full event', () => {
      const events = calendarTools.listEvents({ limit: 1 });
      const firstEvent = events[0];

      if (firstEvent) {
        const event = calendarTools.getEvent({ event_id: firstEvent.id });
        expect(event).toHaveProperty('location');
        expect(event).toHaveProperty('description');
        expect(event).toHaveProperty('organizer');
        expect(event).toHaveProperty('attendees');
      }
    });
  });

  // ---------------------------------------------------------------------------
  // searchEvents
  // ---------------------------------------------------------------------------

  describe('searchEvents', () => {
    it('searches events by query', () => {
      const events = calendarTools.searchEvents({ query: 'meeting', limit: 50 });
      // searchEvents returns results based on repository search
      expect(Array.isArray(events)).toBe(true);
    });
  });

  // ---------------------------------------------------------------------------
  // Content Reader Integration
  // ---------------------------------------------------------------------------

  describe('content reader integration', () => {
    it('uses content reader for event details', () => {
      const mockDetails: EventDetails = {
        title: 'Team Meeting',
        location: 'Conference Room A',
        description: 'Weekly sync',
        organizer: 'boss@example.com',
        attendees: [
          { name: 'John', email: 'john@example.com', status: 'accepted' },
        ],
      };

      const mockContentReader: IEventContentReader = {
        readEventDetails: () => mockDetails,
      };

      const toolsWithReader = createCalendarTools(repository, mockContentReader);
      const events = toolsWithReader.listEvents({ limit: 1 });

      expect(events[0]?.title).toBe('Team Meeting');
    });

    it('can search events when content reader provides titles', () => {
      const mockContentReader: IEventContentReader = {
        readEventDetails: () => ({
          title: 'Team Meeting',
          location: null,
          description: null,
          organizer: null,
          attendees: [],
        }),
      };

      const toolsWithReader = createCalendarTools(repository, mockContentReader);
      const events = toolsWithReader.searchEvents({ query: 'Team', limit: 50 });

      expect(events.length).toBeGreaterThan(0);
      expect(events[0]?.title).toBe('Team Meeting');
    });
  });

  // ---------------------------------------------------------------------------
  // Factory Function
  // ---------------------------------------------------------------------------

  describe('createCalendarTools', () => {
    it('creates a CalendarTools instance', () => {
      const tools = createCalendarTools(repository);
      expect(tools).toBeInstanceOf(CalendarTools);
    });
  });
});
