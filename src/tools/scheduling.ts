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
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolContext, ToolDefinition, ToolResult } from '../registry/types.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    scheduling: SchedulingTools;
  }
}

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

const iso8601DateTimeString = z
  .string()
  .refine((value) => !Number.isNaN(Date.parse(value)), {
    message: 'must be a valid ISO 8601 date-time string',
  });

export const CheckAvailabilityInput = z
  .strictObject({
    email_addresses: z.array(z.string().email()).min(1).describe('Email addresses to check availability for'),
    start_time: iso8601DateTimeString.describe('Start of time window (ISO 8601)'),
    end_time: iso8601DateTimeString.describe('End of time window (ISO 8601)'),
    availability_view_interval: z
      .number()
      .int()
      .min(5)
      .max(1440)
      .default(30)
      .describe('Time slot interval in minutes (default: 30)'),
  })
  .superRefine((data, ctx) => {
    const start = Date.parse(data.start_time);
    const end = Date.parse(data.end_time);
    if (!Number.isNaN(start) && !Number.isNaN(end) && start >= end) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ['end_time'],
        message: 'end_time must be after start_time',
      });
    }
  });

export const FindMeetingTimesInput = z
  .strictObject({
    attendees: z.array(z.string().email()).min(1).describe('Attendee email addresses'),
    duration_minutes: z.number().int().min(1).describe('Meeting duration in minutes'),
    start_time: iso8601DateTimeString.optional().describe('Start of search window (ISO 8601)'),
    end_time: iso8601DateTimeString.optional().describe('End of search window (ISO 8601)'),
    max_candidates: z
      .number()
      .int()
      .min(1)
      .max(25)
      .default(5)
      .describe('Max time suggestions to return (default: 5)'),
  })
  .superRefine((data, ctx) => {
    if (data.start_time != null && data.end_time != null) {
      const start = Date.parse(data.start_time);
      const end = Date.parse(data.end_time);
      if (!Number.isNaN(start) && !Number.isNaN(end) && start >= end) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ['end_time'],
          message: 'end_time must be after start_time when both are provided',
        });
      }
    }
    const hasStart = data.start_time !== undefined;
    const hasEnd = data.end_time !== undefined;
    if (hasStart !== hasEnd) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: hasStart ? ['end_time'] : ['start_time'],
        message: 'start_time and end_time must be provided together or omitted together',
      });
    }
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

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2)
// =============================================================================

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Registry tool definitions for the scheduling domain (Graph API only). The
 * SchedulingTools methods return raw objects, wrapped here to match the
 * pre-registry dispatch behavior exactly.
 */
export function schedulingToolDefinitions(): ToolDefinition[] {
  const tools = (ctx: ToolContext): SchedulingTools => requireGraphToolset(ctx, 'scheduling');

  return [
    defineTool({
      name: 'check_availability',
      description: 'Check free/busy availability for one or more people in a time window',
      input: CheckAvailabilityInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph'],
      handler: async (ctx, params) => jsonResult(await tools(ctx).checkAvailability(params)),
    }),
    defineTool({
      name: 'find_meeting_times',
      description: 'Find available meeting time slots for a group of attendees',
      input: FindMeetingTimesInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['calendar'],
      backends: ['graph'],
      handler: async (ctx, params) => jsonResult(await tools(ctx).findMeetingTimes(params)),
    }),
  ];
}
