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
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset, requireAppleScriptToolset } from '../registry/context.js';
import type { ToolDefinition } from '../registry/types.js';
import type { GraphCalendarTools } from './calendar-graph.js';
import type { AppleCalendarTools } from './calendar-apple.js';

// Calendar is a dual-backend domain: the AppleScript backend serves it via
// AppleCalendarTools; the Graph backend serves it via GraphCalendarTools. The
// advertised (canonical) write schemas are Graph-shaped — Graph is the default
// backend and the AppleScript backend is frozen and adapts.
declare module '../registry/types.js' {
  interface GraphToolsets {
    calendarGraph: GraphCalendarTools;
  }
  interface AppleScriptToolsets {
    calendar: AppleCalendarTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const ListCalendarsInput = z.strictObject({});

export const ListEventsInput = z.strictObject({
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
});

export const GetEventInput = z.strictObject({
  event_id: z.number().int().positive().describe('The event ID to retrieve'),
});

export const SearchEventsInput = z.strictObject({
  query: z.string().min(1).optional().describe('Search query for event titles'),
  start_date: z.string().optional().describe('Start date filter (ISO 8601 format) - events starting on or after this date'),
  end_date: z.string().optional().describe('End date filter (ISO 8601 format) - events ending on or before this date'),
  limit: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(50)
    .describe('Maximum number of events to return (1-100)'),
}).refine(
  (data) => data.query != null || data.start_date != null || data.end_date != null,
  { message: 'At least one of query, start_date, or end_date must be provided' }
);

export const RespondToEventInput = z.strictObject({
  event_id: z.number().int().positive().describe('The event ID to respond to'),
  response: z.enum(['accept', 'decline', 'tentative']).describe('Your response to the invitation'),
  send_response: z.boolean().default(true).describe('Whether to send response to organizer'),
  comment: z.string().optional().describe('Optional comment to include with response'),
});

const DayOfWeek = z.enum([
  'sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday',
]);

const RecurrenceEndInput = z.discriminatedUnion('type', [
  z.strictObject({ type: z.literal('no_end') }),
  z.strictObject({
    type: z.literal('end_date'),
    date: z.string().describe('End date in ISO 8601 format'),
  }),
  z.strictObject({
    type: z.literal('end_after_count'),
    count: z.number().int().min(1).max(999).describe('Number of occurrences'),
  }),
]);

export const RecurrenceInput = z.strictObject({
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
}).superRefine((data, ctx) => {
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

export const CreateEventInput = z.strictObject({
  title: z.string().min(1).describe('Event title/subject'),
  start_date: isoDateString.describe('Start date in ISO 8601 UTC format'),
  end_date: isoDateString.describe('End date in ISO 8601 UTC format'),
  calendar_id: z.number().int().positive().optional().describe('Optional calendar ID to create the event in'),
  location: z.string().optional().describe('Event location'),
  description: z.string().optional().describe('Event description/body text'),
  is_all_day: z.boolean().optional().default(false).describe('Whether this is an all-day event'),
  recurrence: RecurrenceInput.optional().describe('Recurrence pattern to make this a repeating event'),
}).refine(
  (data) => new Date(data.start_date).getTime() < new Date(data.end_date).getTime(),
  { message: 'start_date must be before end_date', path: ['start_date'] }
);

// -----------------------------------------------------------------------------
// Canonical (advertised) write schemas — Graph-shaped.
//
// The two backends historically parsed different schemas for writes. The Graph
// schema is the canonical advertised input for each tool; the AppleScript
// backend receives the superset params and maps only the fields it supports.
// -----------------------------------------------------------------------------

const graphIsoDateString = z
  .string()
  .refine((s) => !isNaN(Date.parse(s)), { message: 'Must be a valid ISO 8601 date string' });

const GraphRecurrenceInput = z.object({
  pattern: z.object({
    type: z.enum(['daily', 'weekly', 'monthly', 'yearly']),
    interval: z.number().int().positive(),
    daysOfWeek: z.array(z.string()).optional(),
  }),
  range: z.object({
    type: z.enum(['endDate', 'noEnd', 'numbered']),
    startDate: z.string(),
    endDate: z.string().optional(),
    numberOfOccurrences: z.number().int().positive().optional(),
  }),
});

const GraphAttendeesInput = z.array(z.object({
  email: z.string().email(),
  name: z.string().optional(),
  type: z.enum(['required', 'optional']).optional(),
})).optional();

export const CreateEventGraphInput = z.strictObject({
  title: z.string().min(1),
  start_date: graphIsoDateString,
  end_date: graphIsoDateString,
  calendar_id: z.number().int().positive().optional(),
  location: z.string().optional(),
  description: z.string().optional(),
  is_all_day: z.boolean().optional().default(false),
  attendees: GraphAttendeesInput,
  recurrence: GraphRecurrenceInput.optional(),
  is_online_meeting: z.boolean().optional().describe('Create as online Teams meeting'),
  online_meeting_provider: z.enum(['teamsForBusiness', 'skypeForBusiness', 'skypeForConsumer']).optional().describe('Online meeting provider (default: teamsForBusiness)'),
}).refine(
  (data) => new Date(data.start_date).getTime() < new Date(data.end_date).getTime(),
  { message: 'start_date must be before end_date', path: ['start_date'] }
);

export const UpdateEventInput = z.strictObject({
  event_id: z.number().int().positive(),
  subject: z.string().optional(),
  start: z.string().optional(),
  end: z.string().optional(),
  timezone: z.string().optional(),
  location: z.string().optional(),
  body: z.string().optional(),
  body_type: z.enum(['text', 'html']).optional(),
  attendees: GraphAttendeesInput,
  is_all_day: z.boolean().optional(),
  recurrence: GraphRecurrenceInput.optional(),
  is_online_meeting: z.boolean().optional().describe('Create as online Teams meeting'),
  online_meeting_provider: z.enum(['teamsForBusiness', 'skypeForBusiness', 'skypeForConsumer']).optional().describe('Online meeting provider (default: teamsForBusiness)'),
  apply_to: z.enum(['this_instance', 'all_in_series']).optional().describe(
    'For recurring events (AppleScript backend): update single instance or entire series (default: this_instance). Ignored by the Graph backend.',
  ),
});

export const DeleteEventInput = z.strictObject({
  event_id: z.number().int().positive().describe('The event ID to delete'),
  apply_to: z.enum(['this_instance', 'all_in_series']).default('this_instance').describe(
    'For recurring events: delete single instance or entire series (default: this_instance)'
  ),
});

export const ListEventInstancesInput = z.strictObject({
  event_id: z.number().int().positive().describe('Recurring event ID'),
  start_date: z.string().describe('Start of date range (ISO 8601, e.g. 2024-01-01T00:00:00Z)'),
  end_date: z.string().describe('End of date range (ISO 8601, e.g. 2024-12-31T23:59:59Z)'),
});

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
export type CreateEventGraphParams = z.infer<typeof CreateEventGraphInput>;
export type UpdateEventParams = z.infer<typeof UpdateEventInput>;
export type DeleteEventParams = z.infer<typeof DeleteEventInput>;
export type ListEventInstancesParams = z.infer<typeof ListEventInstancesInput>;

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
   * Searches events by title and/or date range.
   */
  searchEvents(params: SearchEventsParams): EventSummary[] {
    const { query, start_date, end_date, limit } = params;

    const rows = this.repository.searchEvents(
      query ?? null,
      start_date ?? null,
      end_date ?? null,
      limit
    );

    return rows.map((row) => {
      const details = this.contentReader.readEventDetails(row.dataFilePath);
      return transformEventSummary(row, details?.title ?? null);
    });
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

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2 — dual backend)
// =============================================================================

/**
 * Registry tool definitions for the calendar domain. Each handler branches on
 * the active backend: Graph delegates to GraphCalendarTools; AppleScript
 * delegates to AppleCalendarTools. Both toolsets return MCP content directly.
 */
export function calendarToolDefinitions(): ToolDefinition[] {
  return [
    defineTool({
      name: 'list_calendars',
      description: 'List all calendar folders',
      input: ListCalendarsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph', 'applescript'],
      handler: (ctx) =>
        ctx.backend === 'graph'
          ? requireGraphToolset(ctx, 'calendarGraph').listCalendars()
          : requireAppleScriptToolset(ctx, 'calendar').listCalendars(),
    }),
    defineTool({
      name: 'list_events',
      description: 'List calendar events with optional date range filtering',
      input: ListEventsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph', 'applescript'],
      handler: (ctx, params) =>
        ctx.backend === 'graph'
          ? requireGraphToolset(ctx, 'calendarGraph').listEvents(params)
          : requireAppleScriptToolset(ctx, 'calendar').listEvents(params),
    }),
    defineTool({
      name: 'get_event',
      description: 'Get event details',
      input: GetEventInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph', 'applescript'],
      handler: (ctx, params) =>
        ctx.backend === 'graph'
          ? requireGraphToolset(ctx, 'calendarGraph').getEvent(params)
          : requireAppleScriptToolset(ctx, 'calendar').getEvent(params),
    }),
    defineTool({
      name: 'search_events',
      description: 'Search events by title and/or date range across all calendars',
      input: SearchEventsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph', 'applescript'],
      handler: (ctx, params) =>
        ctx.backend === 'graph'
          ? requireGraphToolset(ctx, 'calendarGraph').searchEvents(params)
          : requireAppleScriptToolset(ctx, 'calendar').searchEvents(params),
    }),
    defineTool({
      name: 'create_event',
      description: 'Create a new calendar event in Outlook. Supports online Teams meetings via is_online_meeting flag.',
      input: CreateEventGraphInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph', 'applescript'],
      handler: (ctx, params) =>
        ctx.backend === 'graph'
          ? requireGraphToolset(ctx, 'calendarGraph').createEvent(params)
          : requireAppleScriptToolset(ctx, 'calendar').createEvent(params),
    }),
    defineTool({
      name: 'respond_to_event',
      description: 'Respond to a meeting invitation (accept, decline, or tentative). Updates your response status and optionally notifies the organizer.',
      input: RespondToEventInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph', 'applescript'],
      handler: (ctx, params) =>
        ctx.backend === 'graph'
          ? requireGraphToolset(ctx, 'calendarGraph').respondToEvent(params)
          : requireAppleScriptToolset(ctx, 'calendar').respondToEvent(params),
    }),
    defineTool({
      name: 'delete_event',
      description: 'Delete a calendar event. For recurring events, you can delete a single instance or the entire series.',
      input: DeleteEventInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['calendar'],
      backends: ['graph', 'applescript'],
      handler: (ctx, params) =>
        ctx.backend === 'graph'
          ? requireGraphToolset(ctx, 'calendarGraph').deleteEvent(params)
          : requireAppleScriptToolset(ctx, 'calendar').deleteEvent(params),
    }),
    defineTool({
      name: 'update_event',
      description: 'Update a calendar event. All fields are optional - only specified fields will be updated. Supports online Teams meetings via is_online_meeting flag. For recurring events, you can update a single instance or the entire series.',
      input: UpdateEventInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph', 'applescript'],
      handler: (ctx, params) =>
        ctx.backend === 'graph'
          ? requireGraphToolset(ctx, 'calendarGraph').updateEvent(params)
          : requireAppleScriptToolset(ctx, 'calendar').updateEvent(params),
    }),
    defineTool({
      name: 'list_event_instances',
      description: 'List instances of a recurring event within a date range. Instance IDs can be used with update_event and delete_event. (Graph API)',
      input: ListEventInstancesInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'calendarGraph').listEventInstances(params),
    }),
  ];
}
