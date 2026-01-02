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
    it('returns empty array when no content reader (titles come from content)', () => {
      const events = calendarTools.searchEvents({ query: 'meeting', limit: 50 });
      // Without a content reader, search can't find anything as titles are in content files
      expect(events.length).toBe(0);
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
