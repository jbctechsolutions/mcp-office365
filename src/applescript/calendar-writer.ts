/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Calendar write operations using AppleScript.
 *
 * Follows the IAccountRepository pattern: a separate interface for
 * AppleScript-only capabilities, keeping IRepository read-only.
 */

import { executeAppleScriptOrThrow } from './executor.js';
import * as scripts from './scripts.js';
import type { RecurrenceScriptParams } from './scripts.js';
import { parseCreateEventResult } from './parser.js';

// =============================================================================
// Types
// =============================================================================

export interface RecurrenceConfig {
  readonly frequency: 'daily' | 'weekly' | 'monthly' | 'yearly';
  readonly interval: number;
  readonly daysOfWeek?: readonly string[];
  readonly dayOfMonth?: number;
  readonly weekOfMonth?: string;
  readonly dayOfWeekMonthly?: string;
  readonly endDate?: string; // ISO 8601
  readonly endAfterCount?: number;
}

export interface CreateEventParams {
  readonly title: string;
  readonly startDate: string; // ISO 8601
  readonly endDate: string; // ISO 8601
  readonly calendarId?: number;
  readonly location?: string;
  readonly description?: string;
  readonly isAllDay?: boolean;
  readonly recurrence?: RecurrenceConfig;
}

export interface CreatedEvent {
  readonly id: number;
  readonly calendarId: number | null;
}

export interface ICalendarWriter {
  createEvent(params: CreateEventParams): CreatedEvent;
}

// =============================================================================
// Date Conversion
// =============================================================================

/**
 * Parses an ISO 8601 date string into individual UTC components.
 * Uses UTC methods so that a Z-suffixed ISO string is interpreted
 * consistently regardless of the host machine's local timezone.
 */
function isoToDateComponents(isoString: string): {
  year: number;
  month: number;
  day: number;
  hours: number;
  minutes: number;
} {
  const date = new Date(isoString);
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1, // JS months are 0-indexed
    day: date.getUTCDate(),
    hours: date.getUTCHours(),
    minutes: date.getUTCMinutes(),
  };
}

// =============================================================================
// Implementation
// =============================================================================

export class AppleScriptCalendarWriter implements ICalendarWriter {
  createEvent(params: CreateEventParams): CreatedEvent {
    const start = isoToDateComponents(params.startDate);
    const end = isoToDateComponents(params.endDate);

    const scriptParams: Parameters<typeof scripts.createEvent>[0] = {
      title: params.title,
      startYear: start.year,
      startMonth: start.month,
      startDay: start.day,
      startHours: start.hours,
      startMinutes: start.minutes,
      endYear: end.year,
      endMonth: end.month,
      endDay: end.day,
      endHours: end.hours,
      endMinutes: end.minutes,
    };
    if (params.calendarId != null) scriptParams.calendarId = params.calendarId;
    if (params.location != null) scriptParams.location = params.location;
    if (params.description != null) scriptParams.description = params.description;
    if (params.isAllDay != null) scriptParams.isAllDay = params.isAllDay;

    if (params.recurrence != null) {
      const rec = params.recurrence;
      const recurrenceScript: RecurrenceScriptParams = {
        frequency: rec.frequency,
        interval: rec.interval,
      };
      // Use mutable alias to conditionally set optional properties
      const mut = recurrenceScript as { -readonly [K in keyof RecurrenceScriptParams]: RecurrenceScriptParams[K] };
      if (rec.daysOfWeek != null) mut.daysOfWeek = rec.daysOfWeek;
      if (rec.dayOfMonth != null) mut.dayOfMonth = rec.dayOfMonth;
      if (rec.weekOfMonth != null) mut.weekOfMonth = rec.weekOfMonth;
      if (rec.dayOfWeekMonthly != null) mut.dayOfWeekMonthly = rec.dayOfWeekMonthly;
      if (rec.endAfterCount != null) mut.endAfterCount = rec.endAfterCount;
      if (rec.endDate != null) mut.endDate = isoToDateComponents(rec.endDate);
      scriptParams.recurrence = recurrenceScript;
    }

    const script = scripts.createEvent(scriptParams);

    const output = executeAppleScriptOrThrow(script);
    const result = parseCreateEventResult(output);

    if (result == null) {
      throw new Error('Failed to parse create event result');
    }

    return result;
  }
}

export function createCalendarWriter(): ICalendarWriter {
  return new AppleScriptCalendarWriter();
}
