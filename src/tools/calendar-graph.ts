/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Graph-backend calendar tools (v3 registry-driven architecture, U2 — dual
 * backend). Holds the calendar logic that previously lived inline in the
 * `handleGraphToolCall` switch, so the registry handlers stay thin and branch
 * on `ctx.backend`.
 */

import type { GraphRepository } from '../graph/repository.js';
import type { GraphContentReaders } from '../graph/content-readers.js';
import type { EventRow, FolderRow } from '../database/repository.js';
import { unixTimestampToLocalIso } from '../graph/mappers/utils.js';
import { hashEventForApproval, type ApprovalTokenManager } from '../approval/index.js';
import type { ToolResult } from '../registry/types.js';
import type {
  CreateEventResult,
  ListEventsParams,
  GetEventParams,
  SearchEventsParams,
  CreateEventGraphParams,
  UpdateEventParams,
  DeleteEventParams,
  RespondToEventParams,
  ListEventInstancesParams,
  PrepareDeleteEventParams,
  ConfirmDeleteEventParams,
  CreateCalendarGroupParams,
  ListRoomsParams,
} from './calendar.js';

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Transforms a folder row into the calendar shape returned by the graph
 * backend's `list_calendars` tool.
 */
function transformFolderRow(row: FolderRow): {
  id: string;
  name: string;
  parentId: string | null;
  specialType: number;
  folderType: number;
  accountId: number;
  messageCount: number;
  unreadCount: number;
} {
  return {
    id: row.id,
    name: row.name ?? 'Unnamed',
    parentId: row.parentId,
    specialType: row.specialType,
    folderType: row.folderType,
    accountId: row.accountId,
    messageCount: row.messageCount,
    unreadCount: row.unreadCount,
  };
}

/**
 * Transforms an EventRow from the Graph backend.
 * Uses Unix timestamps (not Apple epoch) and includes subject from EventRow.
 */
function transformGraphEventRow(row: EventRow): {
  id: string | number;
  folderId: string;
  title: string | null;
  startDate: string | null;
  endDate: string | null;
  isRecurring: boolean;
  hasReminder: boolean;
  attendeeCount: number | null;
  onlineMeetingUrl: string | null;
} {
  return {
    id: row.id,
    folderId: row.folderId,
    title: row.subject ?? null,
    startDate: unixTimestampToLocalIso(row.startDate),
    endDate: unixTimestampToLocalIso(row.endDate),
    isRecurring: row.isRecurring === 1,
    hasReminder: row.hasReminder === 1,
    attendeeCount: row.attendeeCount,
    onlineMeetingUrl: row.onlineMeetingUrl ?? null,
  };
}

/**
 * Graph calendar tools. Each method mirrors the extracted inline graph case
 * body and returns an MCP `ToolResult`.
 */
export class GraphCalendarTools {
  constructor(
    private readonly repository: GraphRepository,
    private readonly contentReaders: GraphContentReaders,
    private readonly tokenManager: ApprovalTokenManager
  ) {}

  async listCalendars(): Promise<ToolResult> {
    const calendars = await this.repository.listCalendarsAsync();
    return jsonResult({ calendars: calendars.map(transformFolderRow) });
  }

  async listEvents(params: ListEventsParams): Promise<ToolResult> {
    let events;
    if (params.start_date != null && params.end_date != null) {
      const startTs = Math.floor(new Date(params.start_date).getTime() / 1000);
      const endTs = Math.floor(new Date(params.end_date).getTime() / 1000);
      events = await this.repository.listEventsByDateRangeAsync(startTs, endTs, params.limit);
    } else if (params.calendar_id != null) {
      events = await this.repository.listEventsByFolderAsync(params.calendar_id, params.limit);
    } else {
      events = await this.repository.listEventsAsync(params.limit);
    }
    return jsonResult({ events: events.map(transformGraphEventRow) });
  }

  async getEvent(params: GetEventParams): Promise<ToolResult> {
    const event = await this.repository.getEventAsync(params.event_id);
    if (event == null) {
      return { content: [{ type: 'text', text: 'Event not found' }], isError: true };
    }

    const details = await this.contentReaders.event.readEventDetailsAsync(event.dataFilePath);
    return jsonResult({ ...transformGraphEventRow(event), ...details });
  }

  async searchEvents(params: SearchEventsParams): Promise<ToolResult> {
    // Graph doesn't have direct event search, so we filter client-side
    const allEvents = await this.repository.listEventsAsync(1000);
    const events = allEvents.filter((e) => {
      const row = transformGraphEventRow(e);
      // Filter by title if query provided
      if (params.query != null) {
        const title = row.title?.toLowerCase() ?? '';
        if (!title.includes(params.query.toLowerCase())) return false;
      }
      // Filter by date range if provided
      if (params.start_date != null && row.startDate != null) {
        if (new Date(row.startDate) < new Date(params.start_date)) return false;
      }
      if (params.end_date != null && row.endDate != null) {
        if (new Date(row.endDate) > new Date(params.end_date)) return false;
      }
      return true;
    });
    return jsonResult({ events: events.slice(0, params.limit).map(transformGraphEventRow) });
  }

  async createEvent(params: CreateEventGraphParams): Promise<ToolResult> {
    const createParams: Parameters<typeof this.repository.createEventAsync>[0] = {
      subject: params.title,
      start: params.start_date,
      end: params.end_date,
    };
    if (params.location != null) createParams.location = params.location;
    if (params.description != null) createParams.body = params.description;
    createParams.bodyType = 'text';
    if (params.is_all_day) createParams.isAllDay = params.is_all_day;
    if (params.attendees != null) {
      createParams.attendees = params.attendees.map((a) => {
        const att: { email: string; name?: string; type?: 'required' | 'optional' } = { email: a.email };
        if (a.name != null) att.name = a.name;
        if (a.type != null) att.type = a.type;
        return att;
      });
    }
    if (params.recurrence != null) {
      const rec = params.recurrence;
      const pattern: { type: 'daily' | 'weekly' | 'monthly' | 'yearly'; interval: number; daysOfWeek?: string[] } = {
        type: rec.pattern.type,
        interval: rec.pattern.interval,
      };
      if (rec.pattern.daysOfWeek != null) pattern.daysOfWeek = rec.pattern.daysOfWeek;
      const range: { type: 'endDate' | 'noEnd' | 'numbered'; startDate: string; endDate?: string; numberOfOccurrences?: number } = {
        type: rec.range.type,
        startDate: rec.range.startDate,
      };
      if (rec.range.endDate != null) range.endDate = rec.range.endDate;
      if (rec.range.numberOfOccurrences != null) range.numberOfOccurrences = rec.range.numberOfOccurrences;
      createParams.recurrence = { pattern, range };
    }
    if (params.calendar_id != null) createParams.calendarId = params.calendar_id;
    if (params.is_online_meeting != null) createParams.is_online_meeting = params.is_online_meeting;
    if (params.online_meeting_provider != null) createParams.online_meeting_provider = params.online_meeting_provider;
    const numericId = await this.repository.createEventAsync(createParams);

    const result: CreateEventResult = {
      id: numericId,
      title: params.title,
      start_date: params.start_date,
      end_date: params.end_date,
      calendar_id: params.calendar_id ?? null,
      location: params.location ?? null,
      description: params.description ?? null,
      is_all_day: params.is_all_day,
      is_recurring: params.recurrence != null,
    };
    return jsonResult(result);
  }

  async updateEvent(params: UpdateEventParams): Promise<ToolResult> {
    const updates: Record<string, unknown> = {};
    const tz = params.timezone ?? Intl.DateTimeFormat().resolvedOptions().timeZone;

    if (params.subject != null) updates.subject = params.subject;
    if (params.start != null) updates.start = { dateTime: params.start, timeZone: tz };
    if (params.end != null) updates.end = { dateTime: params.end, timeZone: tz };
    if (params.location != null) updates.location = { displayName: params.location };
    if (params.body != null) {
      updates.body = {
        contentType: params.body_type ?? 'text',
        content: params.body,
      };
    }
    if (params.is_all_day != null) updates.isAllDay = params.is_all_day;
    if (params.attendees != null) {
      updates.attendees = params.attendees.map((a) => ({
        emailAddress: { address: a.email, name: a.name },
        type: a.type ?? 'required',
      }));
    }
    if (params.recurrence != null) updates.recurrence = params.recurrence;
    if (params.is_online_meeting != null) {
      updates.isOnlineMeeting = params.is_online_meeting;
      if (params.is_online_meeting) {
        updates.onlineMeetingProvider = params.online_meeting_provider ?? 'teamsForBusiness';
      }
    }

    await this.repository.updateEventAsync(params.event_id, updates);
    return { content: [{ type: 'text', text: `Successfully updated event ${params.event_id}` }] };
  }

  async respondToEvent(params: RespondToEventParams): Promise<ToolResult> {
    await this.repository.respondToEventAsync(
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
    return { content: [{ type: 'text', text: `Successfully ${responseText} event ${params.event_id}` }] };
  }

  async deleteEvent(params: DeleteEventParams): Promise<ToolResult> {
    // Deletes exactly the id given: an occurrence id deletes that occurrence, a
    // master id deletes the series (Graph has no instance-vs-series flag).
    await this.repository.deleteEventAsync(params.event_id);
    return { content: [{ type: 'text', text: `Successfully deleted event ${params.event_id}` }] };
  }

  async listEventInstances(params: ListEventInstancesParams): Promise<ToolResult> {
    const instances = await this.repository.listEventInstancesAsync(params.event_id, params.start_date, params.end_date);
    return jsonResult({ instances: instances.map(transformGraphEventRow), count: instances.length });
  }

  async prepareDeleteEvent(params: PrepareDeleteEventParams): Promise<ToolResult> {
    const event = await this.repository.getEventAsync(params.event_id);
    if (event == null) {
      return { content: [{ type: 'text', text: 'Event not found' }], isError: true };
    }

    const graphId = this.repository.getEventGraphId(params.event_id);
    const graphEvent = await this.repository.getClient().getEvent(graphId);
    const hash = hashEventForApproval({
      id: params.event_id,
      subject: graphEvent?.subject ?? null,
      startDateTime: graphEvent?.start?.dateTime ?? null,
    });

    const token = this.tokenManager.generateToken({
      operation: 'delete_event',
      targetType: 'event',
      targetId: params.event_id,
      targetHash: hash,
    });

    return jsonResult({
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      event: transformGraphEventRow(event),
      action: 'This event will be permanently deleted.',
    });
  }

  async confirmDeleteEvent(params: ConfirmDeleteEventParams): Promise<ToolResult> {
    // Re-fetch the event and compute fresh hash for comparison
    const graphId = this.repository.getEventGraphId(params.event_id);
    const graphEvent = await this.repository.getClient().getEvent(graphId);
    const currentHash = hashEventForApproval({
      id: params.event_id,
      subject: graphEvent?.subject ?? null,
      startDateTime: graphEvent?.start?.dateTime ?? null,
    });

    const validation = this.tokenManager.consumeToken(params.token_id, 'delete_event', params.event_id);
    if (!validation.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_event again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_event',
        TARGET_MISMATCH: 'Token was generated for a different event',
        ALREADY_CONSUMED: 'Token has already been used',
      };
      return {
        content: [{ type: 'text', text: errorMessages[validation.error ?? ''] ?? 'Invalid token' }],
        isError: true,
      };
    }

    // Check that the event hasn't changed since prepare
    if (validation.token!.targetHash !== currentHash) {
      return {
        content: [{ type: 'text', text: 'Event has changed since prepare was called. Please call prepare_delete_event again.' }],
        isError: true,
      };
    }

    await this.repository.deleteEventAsync(params.event_id);
    return { content: [{ type: 'text', text: `Successfully deleted event ${params.event_id}` }] };
  }

  async listCalendarGroups(): Promise<ToolResult> {
    const groups = await this.repository.listCalendarGroupsAsync();
    return jsonResult({ calendar_groups: groups });
  }

  async createCalendarGroup(params: CreateCalendarGroupParams): Promise<ToolResult> {
    const groupId = await this.repository.createCalendarGroupAsync(params.name);
    return jsonResult({ success: true, calendar_group_id: groupId, message: 'Calendar group created' });
  }

  async listRoomLists(): Promise<ToolResult> {
    const roomLists = await this.repository.listRoomListsAsync();
    return jsonResult({ room_lists: roomLists });
  }

  async listRooms(params: ListRoomsParams): Promise<ToolResult> {
    const rooms = await this.repository.listRoomsAsync(params.room_list_email);
    return jsonResult({ rooms });
  }
}
