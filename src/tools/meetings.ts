/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Microsoft Teams Online Meetings MCP tools.
 *
 * Provides tools for accessing meeting recordings and transcripts.
 * All operations are read-only except for downloading recordings
 * to the local filesystem.
 */

import { z } from 'zod';

// =============================================================================
// Input Schemas
// =============================================================================

export const ListOnlineMeetingsInput = z.strictObject({
  limit: z.number().int().positive().max(100).optional().describe('Maximum number of meetings to return (default 20)'),
});

export const GetOnlineMeetingInput = z.strictObject({
  meeting_id: z.number().int().positive().describe('Meeting ID from list_online_meetings'),
});

export const ListMeetingRecordingsInput = z.strictObject({
  meeting_id: z.number().int().positive().describe('Meeting ID from list_online_meetings'),
});

export const DownloadMeetingRecordingInput = z.strictObject({
  recording_id: z.number().int().positive().describe('Recording ID from list_meeting_recordings'),
  output_path: z.string().min(1).describe('Local file path to save the recording'),
});

export const ListMeetingTranscriptsInput = z.strictObject({
  meeting_id: z.number().int().positive().describe('Meeting ID from list_online_meetings'),
});

export const GetMeetingTranscriptContentInput = z.strictObject({
  transcript_id: z.number().int().positive().describe('Transcript ID from list_meeting_transcripts'),
  format: z.enum(['text/vtt', 'text/plain']).optional().describe('Transcript format (default text/vtt)'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListOnlineMeetingsParams = z.infer<typeof ListOnlineMeetingsInput>;
export type GetOnlineMeetingParams = z.infer<typeof GetOnlineMeetingInput>;
export type ListMeetingRecordingsParams = z.infer<typeof ListMeetingRecordingsInput>;
export type DownloadMeetingRecordingParams = z.infer<typeof DownloadMeetingRecordingInput>;
export type ListMeetingTranscriptsParams = z.infer<typeof ListMeetingTranscriptsInput>;
export type GetMeetingTranscriptContentParams = z.infer<typeof GetMeetingTranscriptContentInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface IMeetingsRepository {
  listOnlineMeetingsAsync(limit?: number): Promise<Array<{
    id: number; subject: string; startDateTime: string; endDateTime: string; joinUrl: string;
  }>>;
  getOnlineMeetingAsync(meetingId: number): Promise<{
    id: number; subject: string; startDateTime: string; endDateTime: string; joinUrl: string;
    participants: unknown;
  } | undefined>;
  listMeetingRecordingsAsync(meetingId: number): Promise<Array<{
    id: number; createdDateTime: string; recordingContentUrl: string;
  }>>;
  downloadMeetingRecordingAsync(recordingId: number, outputPath: string): Promise<string>;
  listMeetingTranscriptsAsync(meetingId: number): Promise<Array<{
    id: number; createdDateTime: string; contentUrl: string;
  }>>;
  getMeetingTranscriptContentAsync(transcriptId: number, format?: string): Promise<string>;
}

// =============================================================================
// Meetings Tools
// =============================================================================

/**
 * Microsoft Teams Online Meetings tools for recordings and transcripts.
 */
export class MeetingsTools {
  constructor(
    private readonly repo: IMeetingsRepository,
  ) {}

  async listOnlineMeetings(params: ListOnlineMeetingsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const meetings = await this.repo.listOnlineMeetingsAsync(params.limit);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ meetings }, null, 2),
      }],
    };
  }

  async getOnlineMeeting(params: GetOnlineMeetingParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const meeting = await this.repo.getOnlineMeetingAsync(params.meeting_id);
    if (meeting == null) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ error: 'Meeting not found' }, null, 2),
        }],
      };
    }
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ meeting }, null, 2),
      }],
    };
  }

  async listMeetingRecordings(params: ListMeetingRecordingsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const recordings = await this.repo.listMeetingRecordingsAsync(params.meeting_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ recordings }, null, 2),
      }],
    };
  }

  async downloadMeetingRecording(params: DownloadMeetingRecordingParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const filePath = await this.repo.downloadMeetingRecordingAsync(params.recording_id, params.output_path);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, file_path: filePath, message: 'Recording downloaded' }, null, 2),
      }],
    };
  }

  async listMeetingTranscripts(params: ListMeetingTranscriptsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const transcripts = await this.repo.listMeetingTranscriptsAsync(params.meeting_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ transcripts }, null, 2),
      }],
    };
  }

  async getMeetingTranscriptContent(params: GetMeetingTranscriptContentParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const transcriptContent = await this.repo.getMeetingTranscriptContentAsync(params.transcript_id, params.format);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ transcript: transcriptContent }, null, 2),
      }],
    };
  }
}
