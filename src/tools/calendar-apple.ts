/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * AppleScript-backend calendar tools (v3 registry-driven architecture, U2 —
 * dual backend). Holds the calendar logic that previously lived inline in the
 * `handleAppleScriptToolCall` switch, so the registry handlers stay thin and
 * branch on `ctx.backend`.
 *
 * The advertised (canonical) schemas are Graph-shaped (Graph is the default
 * backend). This backend receives the superset params and maps only the fields
 * it supports to the AppleScript writer/manager, exactly as the pre-registry
 * dispatch did.
 */

import type { CalendarTools } from './calendar.js';
import type {
  CreateEventResult,
  ListEventsParams,
  GetEventParams,
  SearchEventsParams,
  CreateEventGraphParams,
  UpdateEventParams,
  DeleteEventParams,
  RespondToEventParams,
} from './calendar.js';
import type {
  ICalendarWriter,
  ICalendarManager,
  RecurrenceConfig,
  EventUpdates,
} from '../applescript/index.js';
import type { ToolResult } from '../registry/types.js';

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * AppleScript calendar tools. Each method mirrors the extracted inline
 * AppleScript case body and returns an MCP `ToolResult`, including the
 * write-capability null checks.
 */
export class AppleCalendarTools {
  constructor(
    private readonly calendarTools: CalendarTools,
    private readonly calendarWriter: ICalendarWriter | null,
    private readonly calendarManager: ICalendarManager | null
  ) {}

  listCalendars(): ToolResult {
    return jsonResult(this.calendarTools.listCalendars({}));
  }

  listEvents(params: ListEventsParams): ToolResult {
    return jsonResult(this.calendarTools.listEvents(params));
  }

  getEvent(params: GetEventParams): ToolResult {
    const result = this.calendarTools.getEvent(params);
    if (result == null) {
      return { content: [{ type: 'text', text: 'Event not found' }], isError: true };
    }
    return jsonResult(result);
  }

  searchEvents(params: SearchEventsParams): ToolResult {
    return jsonResult(this.calendarTools.searchEvents(params));
  }

  createEvent(params: CreateEventGraphParams): ToolResult {
    if (this.calendarWriter == null) {
      return {
        content: [{ type: 'text', text: 'Event creation is not available' }],
        isError: true,
      };
    }
    const writerParams: { title: string; startDate: string; endDate: string; calendarId?: number; location?: string; description?: string; isAllDay?: boolean; recurrence?: RecurrenceConfig } = {
      title: params.title,
      startDate: params.start_date,
      endDate: params.end_date,
    };
    if (params.calendar_id != null) writerParams.calendarId = params.calendar_id;
    if (params.location != null) writerParams.location = params.location;
    if (params.description != null) writerParams.description = params.description;
    if (params.is_all_day != null) writerParams.isAllDay = params.is_all_day;

    if (params.recurrence != null) {
      const rec = params.recurrence;
      const recConfig: RecurrenceConfig = {
        frequency: rec.pattern.type,
        interval: rec.pattern.interval,
      };
      const mut = recConfig as { -readonly [K in keyof RecurrenceConfig]: RecurrenceConfig[K] };
      if (rec.pattern.daysOfWeek != null) mut.daysOfWeek = rec.pattern.daysOfWeek;
      if (rec.range.type === 'endDate' && rec.range.endDate != null) mut.endDate = rec.range.endDate;
      if (rec.range.type === 'numbered' && rec.range.numberOfOccurrences != null) mut.endAfterCount = rec.range.numberOfOccurrences;
      writerParams.recurrence = recConfig;
    }

    const created = this.calendarWriter.createEvent(writerParams);

    const result: CreateEventResult = {
      id: created.id,
      title: params.title,
      start_date: params.start_date,
      end_date: params.end_date,
      calendar_id: created.calendarId,
      location: params.location ?? null,
      description: params.description ?? null,
      is_all_day: params.is_all_day,
      is_recurring: params.recurrence != null,
    };

    return jsonResult(result);
  }

  respondToEvent(params: RespondToEventParams): ToolResult {
    if (this.calendarManager == null) {
      return {
        content: [{ type: 'text', text: 'Event response is not available' }],
        isError: true,
      };
    }

    const result = this.calendarManager.respondToEvent(
      params.event_id,
      params.response,
      params.send_response,
      params.comment
    );

    const responseText = params.response === 'accept'
      ? 'accepted'
      : params.response === 'decline'
      ? 'declined'
      : 'tentatively accepted';

    return {
      content: [{
        type: 'text',
        text: `Successfully ${responseText} event ${result.eventId}`,
      }],
    };
  }

  deleteEvent(params: DeleteEventParams): ToolResult {
    if (this.calendarManager == null) {
      return {
        content: [{ type: 'text', text: 'Event deletion is not available' }],
        isError: true,
      };
    }
    const applyTo = params.apply_to ?? 'this_instance';

    this.calendarManager.deleteEvent(params.event_id, applyTo);

    const deleteText = applyTo === 'all_in_series' ? ' (entire series)' : '';
    return {
      content: [{
        type: 'text',
        text: `Successfully deleted event ${params.event_id}${deleteText}`,
      }],
    };
  }

  updateEvent(params: UpdateEventParams): ToolResult {
    if (this.calendarManager == null) {
      return {
        content: [{ type: 'text', text: 'Event update is not available' }],
        isError: true,
      };
    }

    // Validate date ordering if both dates are provided
    if (params.start != null && params.end != null) {
      if (new Date(params.start).getTime() >= new Date(params.end).getTime()) {
        return {
          content: [{ type: 'text', text: 'start_date must be before end_date' }],
          isError: true,
        };
      }
    }

    const updates: EventUpdates = {
      ...(params.subject != null && { title: params.subject }),
      ...(params.start != null && { startDate: params.start }),
      ...(params.end != null && { endDate: params.end }),
      ...(params.location != null && { location: params.location }),
      ...(params.body != null && { description: params.body }),
      ...(params.is_all_day != null && { isAllDay: params.is_all_day }),
    };

    const result = this.calendarManager.updateEvent(params.event_id, updates, params.apply_to ?? 'this_instance');

    return {
      content: [{
        type: 'text',
        text: `Successfully updated event ${result.id}. Updated fields: ${result.updatedFields.join(', ')}`,
      }],
    };
  }
}
