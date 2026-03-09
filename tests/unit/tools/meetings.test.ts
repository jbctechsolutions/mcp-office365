/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Microsoft Teams Online Meetings tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { MeetingsTools, type IMeetingsRepository } from '../../../src/tools/meetings.js';

describe('MeetingsTools', () => {
  let repo: IMeetingsRepository;
  let tools: MeetingsTools;

  beforeEach(() => {
    repo = {
      listOnlineMeetingsAsync: vi.fn(),
      getOnlineMeetingAsync: vi.fn(),
      listMeetingRecordingsAsync: vi.fn(),
      downloadMeetingRecordingAsync: vi.fn(),
      listMeetingTranscriptsAsync: vi.fn(),
      getMeetingTranscriptContentAsync: vi.fn(),
    };
    tools = new MeetingsTools(repo);
  });

  // ===========================================================================
  // Online Meetings
  // ===========================================================================

  describe('listOnlineMeetings', () => {
    it('returns meetings from the repository', async () => {
      const mockMeetings = [
        { id: 1, subject: 'Sprint Planning', startDateTime: '2026-03-01T10:00:00Z', endDateTime: '2026-03-01T11:00:00Z', joinUrl: 'https://teams.microsoft.com/l/meetup-join/1' },
        { id: 2, subject: 'Standup', startDateTime: '2026-03-02T09:00:00Z', endDateTime: '2026-03-02T09:15:00Z', joinUrl: 'https://teams.microsoft.com/l/meetup-join/2' },
      ];
      vi.mocked(repo.listOnlineMeetingsAsync).mockResolvedValue(mockMeetings);

      const result = await tools.listOnlineMeetings({});

      expect(result.content).toHaveLength(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.meetings).toEqual(mockMeetings);
    });

    it('passes limit to the repository', async () => {
      vi.mocked(repo.listOnlineMeetingsAsync).mockResolvedValue([]);

      await tools.listOnlineMeetings({ limit: 5 });

      expect(repo.listOnlineMeetingsAsync).toHaveBeenCalledWith(5);
    });

    it('passes undefined limit when not specified', async () => {
      vi.mocked(repo.listOnlineMeetingsAsync).mockResolvedValue([]);

      await tools.listOnlineMeetings({});

      expect(repo.listOnlineMeetingsAsync).toHaveBeenCalledWith(undefined);
    });
  });

  describe('getOnlineMeeting', () => {
    it('returns meeting details', async () => {
      const mockMeeting = {
        id: 1,
        subject: 'Sprint Planning',
        startDateTime: '2026-03-01T10:00:00Z',
        endDateTime: '2026-03-01T11:00:00Z',
        joinUrl: 'https://teams.microsoft.com/l/meetup-join/1',
        participants: { organizer: { identity: { user: { displayName: 'John' } } } },
      };
      vi.mocked(repo.getOnlineMeetingAsync).mockResolvedValue(mockMeeting);

      const result = await tools.getOnlineMeeting({ meeting_id: 1 });

      expect(repo.getOnlineMeetingAsync).toHaveBeenCalledWith(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.meeting).toEqual(mockMeeting);
    });

    it('returns error when meeting not found', async () => {
      vi.mocked(repo.getOnlineMeetingAsync).mockResolvedValue(undefined);

      const result = await tools.getOnlineMeeting({ meeting_id: 999 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBe('Meeting not found');
    });
  });

  // ===========================================================================
  // Recordings
  // ===========================================================================

  describe('listMeetingRecordings', () => {
    it('returns recordings for a meeting', async () => {
      const mockRecordings = [
        { id: 10, createdDateTime: '2026-03-01T11:05:00Z', recordingContentUrl: 'https://graph.microsoft.com/v1.0/...' },
        { id: 11, createdDateTime: '2026-03-01T11:10:00Z', recordingContentUrl: 'https://graph.microsoft.com/v1.0/...' },
      ];
      vi.mocked(repo.listMeetingRecordingsAsync).mockResolvedValue(mockRecordings);

      const result = await tools.listMeetingRecordings({ meeting_id: 1 });

      expect(repo.listMeetingRecordingsAsync).toHaveBeenCalledWith(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.recordings).toEqual(mockRecordings);
    });
  });

  describe('downloadMeetingRecording', () => {
    it('saves file and returns path', async () => {
      vi.mocked(repo.downloadMeetingRecordingAsync).mockResolvedValue('/tmp/recording.mp4');

      const result = await tools.downloadMeetingRecording({ recording_id: 10, output_path: '/tmp/recording.mp4' });

      expect(repo.downloadMeetingRecordingAsync).toHaveBeenCalledWith(10, '/tmp/recording.mp4');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.file_path).toBe('/tmp/recording.mp4');
      expect(parsed.message).toBe('Recording downloaded');
    });
  });

  // ===========================================================================
  // Transcripts
  // ===========================================================================

  describe('listMeetingTranscripts', () => {
    it('returns transcripts for a meeting', async () => {
      const mockTranscripts = [
        { id: 20, createdDateTime: '2026-03-01T11:05:00Z', contentUrl: 'https://graph.microsoft.com/v1.0/...' },
      ];
      vi.mocked(repo.listMeetingTranscriptsAsync).mockResolvedValue(mockTranscripts);

      const result = await tools.listMeetingTranscripts({ meeting_id: 1 });

      expect(repo.listMeetingTranscriptsAsync).toHaveBeenCalledWith(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.transcripts).toEqual(mockTranscripts);
    });
  });

  describe('getMeetingTranscriptContent', () => {
    it('returns transcript text', async () => {
      const vttContent = 'WEBVTT\n\n00:00:00.000 --> 00:00:05.000\nHello everyone, welcome to the meeting.';
      vi.mocked(repo.getMeetingTranscriptContentAsync).mockResolvedValue(vttContent);

      const result = await tools.getMeetingTranscriptContent({ transcript_id: 20 });

      expect(repo.getMeetingTranscriptContentAsync).toHaveBeenCalledWith(20, undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.transcript).toBe(vttContent);
    });

    it('passes format to the repository', async () => {
      vi.mocked(repo.getMeetingTranscriptContentAsync).mockResolvedValue('Hello everyone');

      await tools.getMeetingTranscriptContent({ transcript_id: 20, format: 'text/plain' });

      expect(repo.getMeetingTranscriptContentAsync).toHaveBeenCalledWith(20, 'text/plain');
    });
  });
});
