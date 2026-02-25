/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { SchedulingTools, createSchedulingTools, type ISchedulingRepository } from '../../../src/tools/scheduling.js';

function createMockRepository(): ISchedulingRepository {
  return {
    getScheduleAsync: vi.fn(),
    findMeetingTimesAsync: vi.fn(),
  };
}

describe('scheduling tools', () => {
  let repo: ISchedulingRepository;
  let tools: SchedulingTools;

  beforeEach(() => {
    repo = createMockRepository();
    tools = new SchedulingTools(repo);
  });

  describe('checkAvailability', () => {
    it('returns schedule data for requested attendees', async () => {
      const mockSchedules = [
        {
          scheduleId: 'bob@example.com',
          availabilityView: '020120',
          scheduleItems: [
            { status: 'busy', start: { dateTime: '2026-02-24T10:00:00' }, end: { dateTime: '2026-02-24T11:00:00' } },
          ],
        },
      ];
      (repo.getScheduleAsync as ReturnType<typeof vi.fn>).mockResolvedValue(mockSchedules);

      const result = await tools.checkAvailability({
        email_addresses: ['bob@example.com'],
        start_time: '2026-02-24T08:00:00Z',
        end_time: '2026-02-24T18:00:00Z',
        availability_view_interval: 30,
      });

      expect(result).toEqual({ schedules: mockSchedules });
      expect(repo.getScheduleAsync).toHaveBeenCalledWith({
        emailAddresses: ['bob@example.com'],
        startTime: '2026-02-24T08:00:00Z',
        endTime: '2026-02-24T18:00:00Z',
        availabilityViewInterval: 30,
      });
    });

    it('uses default interval when not specified', async () => {
      (repo.getScheduleAsync as ReturnType<typeof vi.fn>).mockResolvedValue([]);

      await tools.checkAvailability({
        email_addresses: ['bob@example.com'],
        start_time: '2026-02-24T08:00:00Z',
        end_time: '2026-02-24T18:00:00Z',
      });

      expect(repo.getScheduleAsync).toHaveBeenCalledWith(
        expect.objectContaining({ availabilityViewInterval: 30 })
      );
    });
  });

  describe('findMeetingTimes', () => {
    it('returns meeting time suggestions', async () => {
      const mockResult = {
        meetingTimeSuggestions: [
          { confidence: 100, meetingTimeSlot: { start: {}, end: {} } },
        ],
        emptySuggestionsReason: '',
      };
      (repo.findMeetingTimesAsync as ReturnType<typeof vi.fn>).mockResolvedValue(mockResult);

      const result = await tools.findMeetingTimes({
        attendees: ['bob@example.com', 'alice@example.com'],
        duration_minutes: 60,
        start_time: '2026-02-24T08:00:00Z',
        end_time: '2026-02-24T18:00:00Z',
        max_candidates: 3,
      });

      expect(result).toEqual(mockResult);
      expect(repo.findMeetingTimesAsync).toHaveBeenCalledWith({
        attendees: ['bob@example.com', 'alice@example.com'],
        durationMinutes: 60,
        startTime: '2026-02-24T08:00:00Z',
        endTime: '2026-02-24T18:00:00Z',
        maxCandidates: 3,
      });
    });

    it('uses defaults for optional params', async () => {
      (repo.findMeetingTimesAsync as ReturnType<typeof vi.fn>).mockResolvedValue({ meetingTimeSuggestions: [] });

      await tools.findMeetingTimes({
        attendees: ['bob@example.com'],
        duration_minutes: 30,
      });

      expect(repo.findMeetingTimesAsync).toHaveBeenCalledWith({
        attendees: ['bob@example.com'],
        durationMinutes: 30,
        startTime: undefined,
        endTime: undefined,
        maxCandidates: 5,
      });
    });
  });

  describe('createSchedulingTools', () => {
    it('returns a SchedulingTools instance', () => {
      const result = createSchedulingTools(repo);
      expect(result).toBeInstanceOf(SchedulingTools);
    });
  });
});
