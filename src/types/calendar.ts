/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Calendar-related type definitions.
 */

/**
 * Calendar folder.
 */
export interface CalendarFolder {
  readonly id: number;
  readonly name: string;
  readonly accountId: number;
}

/**
 * Event summary for list views.
 */
export interface EventSummary {
  readonly id: number;
  readonly folderId: number;
  readonly title: string | null;
  readonly startDate: string | null;
  readonly endDate: string | null;
  readonly isRecurring: boolean;
  readonly hasReminder: boolean;
  readonly attendeeCount: number;
  readonly uid: string | null;
}

/**
 * Full event details including description and attendees.
 */
export interface Event extends EventSummary {
  readonly location: string | null;
  readonly description: string | null;
  readonly organizer: string | null;
  readonly attendees: readonly Attendee[];
  readonly masterRecordId: number | null;
  readonly recurrenceId: number | null;
}

/**
 * Event attendee information.
 */
export interface Attendee {
  readonly name: string | null;
  readonly email: string | null;
  readonly status: AttendeeStatus;
}

/**
 * Attendee response status.
 */
export const AttendeeStatus = {
  Unknown: 'unknown',
  Accepted: 'accepted',
  Declined: 'declined',
  Tentative: 'tentative',
  None: 'none',
} as const;

export type AttendeeStatus =
  (typeof AttendeeStatus)[keyof typeof AttendeeStatus];
