/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Maps Microsoft Graph Event type to EventRow.
 */

import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import type { EventRow } from '../../database/repository.js';
import { mintSelfEncoded } from '../../ids/token.js';
import {
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
    // Durable self-encoding ev_ token carrying the immutable Graph event id (U5).
    id: eventId.length > 0 ? mintSelfEncoded('event', eventId) : '',
    // Durable self-encoding fd_ token carrying the immutable Graph calendar id (U5).
    folderId: calendarId != null ? mintSelfEncoded('folder', calendarId) : '',
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
