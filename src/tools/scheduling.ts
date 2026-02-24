/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Calendar scheduling MCP tools.
 *
 * Provides tools for checking free/busy availability and
 * finding optimal meeting times for groups of attendees.
 */

import { z } from 'zod';

// =============================================================================
// Repository Interface
// =============================================================================

export interface ISchedulingRepository {
  getScheduleAsync(params: {
    emailAddresses: string[];
    startTime: string;
    endTime: string;
    availabilityViewInterval?: number;
  }): Promise<unknown[]>;

  findMeetingTimesAsync(params: {
    attendees: string[];
    durationMinutes: number;
    startTime?: string;
    endTime?: string;
    maxCandidates?: number;
  }): Promise<unknown>;
}

// =============================================================================
// Zod Schemas
// =============================================================================

export const CheckAvailabilityInput = z.strictObject({
  email_addresses: z.array(z.string().email()).min(1).describe('Email addresses to check availability for'),
  start_time: z.string().describe('Start of time window (ISO 8601)'),
  end_time: z.string().describe('End of time window (ISO 8601)'),
  availability_view_interval: z.number().int().min(5).max(1440).default(30).describe('Time slot interval in minutes (default: 30)'),
});

export const FindMeetingTimesInput = z.strictObject({
  attendees: z.array(z.string().email()).min(1).describe('Attendee email addresses'),
  duration_minutes: z.number().int().min(1).describe('Meeting duration in minutes'),
  start_time: z.string().optional().describe('Start of search window (ISO 8601)'),
  end_time: z.string().optional().describe('End of search window (ISO 8601)'),
  max_candidates: z.number().int().min(1).max(25).default(5).describe('Max time suggestions to return (default: 5)'),
});

// =============================================================================
// Scheduling Tools Class
// =============================================================================

export class SchedulingTools {
  constructor(private readonly repository: ISchedulingRepository) {}

  async checkAvailability(params: z.input<typeof CheckAvailabilityInput>): Promise<{ schedules: unknown[] }> {
    const parsed = CheckAvailabilityInput.parse(params);
    const schedules = await this.repository.getScheduleAsync({
      emailAddresses: parsed.email_addresses,
      startTime: parsed.start_time,
      endTime: parsed.end_time,
      availabilityViewInterval: parsed.availability_view_interval,
    });
    return { schedules };
  }

  async findMeetingTimes(params: z.input<typeof FindMeetingTimesInput>): Promise<unknown> {
    const parsed = FindMeetingTimesInput.parse(params);
    return await this.repository.findMeetingTimesAsync({
      attendees: parsed.attendees,
      durationMinutes: parsed.duration_minutes,
      ...(parsed.start_time !== undefined && { startTime: parsed.start_time }),
      ...(parsed.end_time !== undefined && { endTime: parsed.end_time }),
      maxCandidates: parsed.max_candidates,
    });
  }
}

// =============================================================================
// Factory
// =============================================================================

export function createSchedulingTools(repository: ISchedulingRepository): SchedulingTools {
  return new SchedulingTools(repository);
}
