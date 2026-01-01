/**
 * Calendar-related MCP tools.
 *
 * Provides tools for listing calendars, events, and searching.
 */

import { z } from 'zod';
import type { IRepository, EventRow, FolderRow } from '../database/repository.js';
import type { CalendarFolder, EventSummary, Event, Attendee } from '../types/index.js';
import { appleTimestampToIso, isoToAppleTimestamp } from '../utils/dates.js';

// =============================================================================
// Input Schemas
// =============================================================================

export const ListCalendarsInput = z.object({}).strict();

export const ListEventsInput = z
  .object({
    calendar_id: z.number().int().positive().optional().describe('Optional calendar folder ID'),
    start_date: z.string().optional().describe('Start date filter (ISO 8601 format)'),
    end_date: z.string().optional().describe('End date filter (ISO 8601 format)'),
    limit: z
      .number()
      .int()
      .min(1)
      .max(100)
      .default(50)
      .describe('Maximum number of events to return (1-100)'),
  })
  .strict();

export const GetEventInput = z
  .object({
    event_id: z.number().int().positive().describe('The event ID to retrieve'),
  })
  .strict();

export const SearchEventsInput = z
  .object({
    query: z.string().min(1).describe('Search query for event titles'),
    limit: z
      .number()
      .int()
      .min(1)
      .max(100)
      .default(50)
      .describe('Maximum number of events to return (1-100)'),
  })
  .strict();

// =============================================================================
// Type Definitions
// =============================================================================

export type ListCalendarsParams = z.infer<typeof ListCalendarsInput>;
export type ListEventsParams = z.infer<typeof ListEventsInput>;
export type GetEventParams = z.infer<typeof GetEventInput>;
export type SearchEventsParams = z.infer<typeof SearchEventsInput>;

// =============================================================================
// Content Reader Interface
// =============================================================================

/**
 * Interface for reading event content from data files.
 */
export interface IEventContentReader {
  /**
   * Reads event details from the given data file path.
   */
  readEventDetails(dataFilePath: string | null): EventDetails | null;
}

/**
 * Event details from content file.
 */
export interface EventDetails {
  readonly title: string | null;
  readonly location: string | null;
  readonly description: string | null;
  readonly organizer: string | null;
  readonly attendees: readonly Attendee[];
}

/**
 * Default event content reader that returns null.
 */
export const nullEventContentReader: IEventContentReader = {
  readEventDetails: (): EventDetails | null => null,
};

// =============================================================================
// Transformers
// =============================================================================

/**
 * Transforms a database folder row to CalendarFolder.
 */
function transformCalendar(row: FolderRow): CalendarFolder {
  return {
    id: row.id,
    name: row.name ?? 'Unnamed',
    accountId: row.accountId,
  };
}

/**
 * Transforms a database event row to EventSummary.
 */
function transformEventSummary(row: EventRow, title: string | null = null): EventSummary {
  return {
    id: row.id,
    folderId: row.folderId,
    title: title,
    startDate: appleTimestampToIso(row.startDate),
    endDate: appleTimestampToIso(row.endDate),
    isRecurring: row.isRecurring === 1,
    hasReminder: row.hasReminder === 1,
    attendeeCount: row.attendeeCount,
    uid: row.uid,
  };
}

/**
 * Transforms a database event row to full Event.
 */
function transformEvent(row: EventRow, details: EventDetails | null): Event {
  const summary = transformEventSummary(row, details?.title ?? null);

  return {
    ...summary,
    location: details?.location ?? null,
    description: details?.description ?? null,
    organizer: details?.organizer ?? null,
    attendees: details?.attendees ?? [],
    masterRecordId: row.masterRecordId ?? null,
    recurrenceId: row.recurrenceId ?? null,
  };
}

// =============================================================================
// Calendar Tools Class
// =============================================================================

/**
 * Calendar tools implementation with dependency injection.
 */
export class CalendarTools {
  constructor(
    private readonly repository: IRepository,
    private readonly contentReader: IEventContentReader = nullEventContentReader
  ) {}

  /**
   * Lists all calendar folders.
   */
  listCalendars(_params: ListCalendarsParams): CalendarFolder[] {
    const rows = this.repository.listCalendars();
    return rows.map(transformCalendar);
  }

  /**
   * Lists events with optional filtering.
   */
  listEvents(params: ListEventsParams): EventSummary[] {
    const { calendar_id, start_date, end_date, limit } = params;

    let rows: EventRow[];

    if (start_date != null && end_date != null) {
      const startTimestamp = isoToAppleTimestamp(start_date);
      const endTimestamp = isoToAppleTimestamp(end_date);
      if (startTimestamp != null && endTimestamp != null) {
        rows = this.repository.listEventsByDateRange(startTimestamp, endTimestamp, limit);
      } else {
        rows = this.repository.listEvents(limit);
      }
    } else if (calendar_id != null) {
      rows = this.repository.listEventsByFolder(calendar_id, limit);
    } else {
      rows = this.repository.listEvents(limit);
    }

    return rows.map((row) => {
      const details = this.contentReader.readEventDetails(row.dataFilePath);
      return transformEventSummary(row, details?.title ?? null);
    });
  }

  /**
   * Gets a single event by ID.
   */
  getEvent(params: GetEventParams): Event | null {
    const { event_id } = params;

    const row = this.repository.getEvent(event_id);
    if (row == null) {
      return null;
    }

    const details = this.contentReader.readEventDetails(row.dataFilePath);
    return transformEvent(row, details);
  }

  /**
   * Searches events by title (requires content reader).
   * Note: Basic implementation - returns all events since title is in content files.
   */
  searchEvents(params: SearchEventsParams): EventSummary[] {
    const { query, limit } = params;
    const queryLower = query.toLowerCase();

    // Get all events and filter by title (from content reader)
    const rows = this.repository.listEvents(limit * 2); // Fetch more to filter
    const results: EventSummary[] = [];

    for (const row of rows) {
      if (results.length >= limit) break;

      const details = this.contentReader.readEventDetails(row.dataFilePath);
      const title = details?.title ?? '';

      if (title.toLowerCase().includes(queryLower)) {
        results.push(transformEventSummary(row, title));
      }
    }

    return results;
  }
}

/**
 * Creates calendar tools with the given repository.
 */
export function createCalendarTools(
  repository: IRepository,
  contentReader: IEventContentReader = nullEventContentReader
): CalendarTools {
  return new CalendarTools(repository, contentReader);
}
