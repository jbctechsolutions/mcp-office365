/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

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

export const RespondToEventInput = z
  .object({
    event_id: z.number().int().positive().describe('The event ID to respond to'),
    response: z.enum(['accept', 'decline', 'tentative']).describe('Your response to the invitation'),
    send_response: z.boolean().default(true).describe('Whether to send response to organizer'),
    comment: z.string().optional().describe('Optional comment to include with response'),
  })
  .strict();

const DayOfWeek = z.enum([
  'sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday',
]);

const RecurrenceEndInput = z.discriminatedUnion('type', [
  z.object({ type: z.literal('no_end') }).strict(),
  z.object({
    type: z.literal('end_date'),
    date: z.string().describe('End date in ISO 8601 format'),
  }).strict(),
  z.object({
    type: z.literal('end_after_count'),
    count: z.number().int().min(1).max(999).describe('Number of occurrences'),
  }).strict(),
]);

export const RecurrenceInput = z.object({
  frequency: z.enum(['daily', 'weekly', 'monthly', 'yearly']).describe('How often the event repeats'),
  interval: z.number().int().min(1).max(999).default(1).describe(
    'Number of frequency units between occurrences (e.g., 2 for every 2 weeks)'
  ),
  days_of_week: z.array(DayOfWeek).min(1).optional().describe(
    'Days of the week for weekly recurrence (e.g., ["monday", "wednesday"])'
  ),
  day_of_month: z.number().int().min(1).max(31).optional().describe(
    'Day of the month for monthly recurrence (e.g., 15 for the 15th)'
  ),
  week_of_month: z.enum(['first', 'second', 'third', 'fourth', 'last']).optional().describe(
    'Week of the month for ordinal monthly recurrence (e.g., "third" for 3rd Thursday)'
  ),
  day_of_week_monthly: DayOfWeek.optional().describe(
    'Day of week for ordinal monthly recurrence (used with week_of_month)'
  ),
  end: RecurrenceEndInput.default({ type: 'no_end' }).describe('When the recurrence ends'),
}).strict().superRefine((data, ctx) => {
  if (data.frequency === 'weekly' && (data.days_of_week == null || data.days_of_week.length === 0)) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: 'days_of_week is required for weekly recurrence',
      path: ['days_of_week'],
    });
  }
  if (data.frequency === 'monthly') {
    const hasOrdinal = data.week_of_month != null;
    const hasDayOfWeekMonthly = data.day_of_week_monthly != null;
    if (hasOrdinal !== hasDayOfWeekMonthly) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: 'week_of_month and day_of_week_monthly must both be provided for ordinal monthly recurrence',
        path: ['week_of_month'],
      });
    }
  }
  if (data.frequency !== 'monthly') {
    if (data.day_of_month != null || data.week_of_month != null || data.day_of_week_monthly != null) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: 'day_of_month, week_of_month, and day_of_week_monthly are only valid for monthly recurrence',
        path: ['frequency'],
      });
    }
  }
  if (data.frequency !== 'weekly' && data.days_of_week != null) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: 'days_of_week is only valid for weekly recurrence',
      path: ['days_of_week'],
    });
  }
});

const isoDateString = z
  .string()
  .refine((s) => !isNaN(Date.parse(s)), { message: 'Must be a valid ISO 8601 date string' });

export const CreateEventInput = z
  .object({
    title: z.string().min(1).describe('Event title/subject'),
    start_date: isoDateString.describe('Start date in ISO 8601 UTC format'),
    end_date: isoDateString.describe('End date in ISO 8601 UTC format'),
    calendar_id: z.number().int().positive().optional().describe('Optional calendar ID to create the event in'),
    location: z.string().optional().describe('Event location'),
    description: z.string().optional().describe('Event description/body text'),
    is_all_day: z.boolean().optional().default(false).describe('Whether this is an all-day event'),
    recurrence: RecurrenceInput.optional().describe('Recurrence pattern to make this a repeating event'),
  })
  .strict()
  .refine(
    (data) => new Date(data.start_date).getTime() < new Date(data.end_date).getTime(),
    { message: 'start_date must be before end_date', path: ['start_date'] }
  );

// =============================================================================
// Type Definitions
// =============================================================================

export type ListCalendarsParams = z.infer<typeof ListCalendarsInput>;
export type ListEventsParams = z.infer<typeof ListEventsInput>;
export type GetEventParams = z.infer<typeof GetEventInput>;
export type SearchEventsParams = z.infer<typeof SearchEventsInput>;
export type CreateEventParams = z.infer<typeof CreateEventInput>;
export type RecurrenceParams = z.infer<typeof RecurrenceInput>;
export type RespondToEventParams = z.infer<typeof RespondToEventInput>;

/**
 * Result of creating a calendar event.
 */
export interface CreateEventResult {
  readonly id: number;
  readonly title: string;
  readonly start_date: string;
  readonly end_date: string;
  readonly calendar_id: number | null;
  readonly location: string | null;
  readonly description: string | null;
  readonly is_all_day: boolean;
  readonly is_recurring: boolean;
}

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
