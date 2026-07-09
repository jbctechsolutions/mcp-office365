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
import type { Attendee } from '../types/index.js';
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolDefinition } from '../registry/types.js';
import type { GraphCalendarTools } from './calendar-graph.js';

// The advertised (canonical) write schemas below are Graph-shaped — Graph is
// the only backend.
declare module '../registry/types.js' {
  interface GraphToolsets {
    calendarGraph: GraphCalendarTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

// An event id accepts either a durable `ev_…` token (Graph backend, U5) or a
// numeric id (AppleScript/SQLite backend, D4). A numeric id on Graph is rejected
// with NUMERIC_ID_UNSUPPORTED by the resolver.
const EventIdSchema = z.union([z.string().min(1), z.number().int().positive()]);

export const ListCalendarsInput = z.strictObject({});

export const ListEventsInput = z.strictObject({
  calendar_id: z.string().min(1).optional().describe('Optional calendar folder ID'),
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
  event_id: EventIdSchema.describe('The event ID to retrieve'),
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
  event_id: EventIdSchema.describe('The event ID to respond to'),
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

// -----------------------------------------------------------------------------
// Canonical (advertised) write schemas — Graph-shaped.
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
  calendar_id: z.string().min(1).optional(),
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

// create_event/update_event advertise the Graph schema: field names follow
// Graph conventions (subject/start/end/body), and recurrence uses the Graph
// {pattern, range} shape.
export const UpdateEventInput = z.strictObject({
  event_id: EventIdSchema,
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
});

export const DeleteEventInput = z.strictObject({
  // To delete one occurrence of a recurring series, pass that occurrence's own
  // event_id (from list_event_instances); to delete the whole series, pass the
  // master event_id. Graph deletes exactly the id given — there is no separate
  // instance-vs-series flag.
  event_id: EventIdSchema.describe('The event ID to delete (an occurrence id deletes just that occurrence; the master id deletes the series)'),
});

export const ListEventInstancesInput = z.strictObject({
  event_id: EventIdSchema.describe('Recurring event ID'),
  start_date: z.string().describe('Start of date range (ISO 8601, e.g. 2024-01-01T00:00:00Z)'),
  end_date: z.string().describe('End of date range (ISO 8601, e.g. 2024-12-31T23:59:59Z)'),
});

export const PrepareDeleteEventInput = z.strictObject({
  event_id: EventIdSchema.describe('The event ID to delete'),
});

export const ConfirmDeleteEventInput = z.strictObject({
  token_id: z.uuid().describe('The approval token from prepare_delete_event'),
  event_id: EventIdSchema.describe('The event ID to delete'),
});

export const ListCalendarGroupsInput = z.strictObject({});

export const CreateCalendarGroupInput = z.strictObject({
  name: z.string().min(1).describe('Calendar group name'),
});

export const ListRoomListsInput = z.strictObject({});

export const ListRoomsInput = z.strictObject({
  room_list_email: z.string().email().optional().describe('Room list email to filter by (from list_room_lists)'),
});

// =============================================================================
// Type Definitions
// =============================================================================

export type ListCalendarsParams = z.infer<typeof ListCalendarsInput>;
export type ListEventsParams = z.infer<typeof ListEventsInput>;
export type GetEventParams = z.infer<typeof GetEventInput>;
export type SearchEventsParams = z.infer<typeof SearchEventsInput>;
export type RecurrenceParams = z.infer<typeof RecurrenceInput>;
export type RespondToEventParams = z.infer<typeof RespondToEventInput>;
export type CreateEventGraphParams = z.infer<typeof CreateEventGraphInput>;
export type UpdateEventParams = z.infer<typeof UpdateEventInput>;
export type DeleteEventParams = z.infer<typeof DeleteEventInput>;
export type ListEventInstancesParams = z.infer<typeof ListEventInstancesInput>;
export type PrepareDeleteEventParams = z.infer<typeof PrepareDeleteEventInput>;
export type ConfirmDeleteEventParams = z.infer<typeof ConfirmDeleteEventInput>;
export type CreateCalendarGroupParams = z.infer<typeof CreateCalendarGroupInput>;
export type ListRoomsParams = z.infer<typeof ListRoomsInput>;

/**
 * Result of creating a calendar event.
 */
export interface CreateEventResult {
  // Durable ev_ token on Graph (U5).
  readonly id: string | number;
  readonly title: string;
  readonly start_date: string;
  readonly end_date: string;
  readonly calendar_id: string | null;
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

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2 — dual backend)
// =============================================================================

/**
 * Registry tool definitions for the calendar domain. Each handler delegates to
 * GraphCalendarTools, which returns MCP content directly.
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
      backends: ['graph'],
      handler: (ctx) => requireGraphToolset(ctx, 'calendarGraph').listCalendars(),
    }),
    defineTool({
      name: 'list_events',
      description: 'List calendar events with optional date range filtering',
      input: ListEventsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'calendarGraph').listEvents(params),
    }),
    defineTool({
      name: 'get_event',
      description: 'Get event details',
      input: GetEventInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'calendarGraph').getEvent(params),
    }),
    defineTool({
      name: 'search_events',
      description: 'Search events by title and/or date range across all calendars',
      input: SearchEventsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'calendarGraph').searchEvents(params),
    }),
    defineTool({
      name: 'create_event',
      description: 'Create a new calendar event in Outlook. Supports online Teams meetings via is_online_meeting flag.',
      input: CreateEventGraphInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'calendarGraph').createEvent(params),
    }),
    defineTool({
      name: 'respond_to_event',
      description: 'Respond to a meeting invitation (accept, decline, or tentative). Updates your response status and optionally notifies the organizer.',
      input: RespondToEventInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'calendarGraph').respondToEvent(params),
    }),
    defineTool({
      name: 'delete_event',
      description: 'Delete a calendar event. For recurring events, you can delete a single instance or the entire series.',
      input: DeleteEventInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['calendar'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'calendarGraph').deleteEvent(params),
    }),
    defineTool({
      name: 'update_event',
      description: 'Update a calendar event. All fields are optional - only specified fields will be updated. Supports online Teams meetings via is_online_meeting flag. For recurring events, you can update a single instance or the entire series.',
      input: UpdateEventInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'calendarGraph').updateEvent(params),
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
    // ---- Graph-only destructive two-phase (delete event) ----
    defineTool({
      name: 'prepare_delete_event',
      description: 'Prepare to delete a calendar event. Returns a preview and approval token. Call confirm_delete_event to execute.',
      input: PrepareDeleteEventInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['calendar'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'calendarGraph').prepareDeleteEvent(params),
    }),
    defineTool({
      name: 'confirm_delete_event',
      description: 'Confirm deletion of a calendar event using a token from prepare_delete_event',
      input: ConfirmDeleteEventInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['calendar'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'calendarGraph').confirmDeleteEvent(params),
    }),
    // ---- Graph-only calendar groups ----
    defineTool({
      name: 'list_calendar_groups',
      description: 'List all calendar groups (Graph API)',
      input: ListCalendarGroupsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph'],
      handler: (ctx) => requireGraphToolset(ctx, 'calendarGraph').listCalendarGroups(),
    }),
    defineTool({
      name: 'create_calendar_group',
      description: 'Create a new calendar group (Graph API)',
      input: CreateCalendarGroupInput,
      annotations: { readOnlyHint: false, destructiveHint: false },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'calendarGraph').createCalendarGroup(params),
    }),
    // ---- Graph-only room lists & rooms ----
    defineTool({
      name: 'list_room_lists',
      description: 'List all room lists (building/floor groupings) in the organization (Graph API)',
      input: ListRoomListsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph'],
      handler: (ctx) => requireGraphToolset(ctx, 'calendarGraph').listRoomLists(),
    }),
    defineTool({
      name: 'list_rooms',
      description: 'List meeting rooms, optionally filtered by a room list email from list_room_lists (Graph API)',
      input: ListRoomsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'calendarGraph').listRooms(params),
    }),
  ];
}
