/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Maps Microsoft Graph Event type to EventRow.
 */

import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import type { EventRow } from '../../database/repository.js';
import {
  hashStringToNumber,
  dateTimeTimeZoneToTimestamp,
  createGraphContentPath,
} from './utils.js';

/**
 * Maps a Graph Event to an EventRow.
 */
export function mapEventToEventRow(
  event: MicrosoftGraph.Event,
  calendarId?: string
): EventRow {
  const eventId = event.id ?? '';

  // Type assertions needed due to Graph API's NullableOption types
  // which are incompatible with exactOptionalPropertyTypes
  const start = event.start as { dateTime?: string; timeZone?: string } | null | undefined;
  const end = event.end as { dateTime?: string; timeZone?: string } | null | undefined;

  return {
    id: hashStringToNumber(eventId),
    folderId: calendarId != null ? hashStringToNumber(calendarId) : 0,
    subject: event.subject ?? null,
    startDate: dateTimeTimeZoneToTimestamp(start),
    endDate: dateTimeTimeZoneToTimestamp(end),
    isRecurring: event.recurrence != null ? 1 : 0,
    hasReminder: event.isReminderOn === true ? 1 : 0,
    attendeeCount: event.attendees?.length ?? 0,
    uid: event.iCalUId ?? null,
    masterRecordId: null, // Not directly available
    recurrenceId: null, // Not directly available
    dataFilePath: createGraphContentPath('event', eventId),
    onlineMeetingUrl: event.onlineMeeting?.joinUrl ?? null,
  };
}
