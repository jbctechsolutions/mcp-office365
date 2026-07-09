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
import { Id } from '../ids/schema.js';
import { nextActionFor } from '../ids/next-action.js';
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolContext, ToolDefinition } from '../registry/types.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    meetings: MeetingsTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const ListOnlineMeetingsInput = z.strictObject({
  limit: z.number().int().positive().max(100).optional().describe('Maximum number of meetings to return (default 20)'),
});

export const GetOnlineMeetingInput = z.strictObject({
  meeting_id: Id.onlineMeeting,
});

export const ListMeetingRecordingsInput = z.strictObject({
  meeting_id: Id.onlineMeeting,
});

export const DownloadMeetingRecordingInput = z.strictObject({
  recording_id: Id.recording,
  output_path: z.string().min(1).describe('Local file path to save the recording'),
});

export const ListMeetingTranscriptsInput = z.strictObject({
  meeting_id: Id.onlineMeeting,
});

export const GetMeetingTranscriptContentInput = z.strictObject({
  transcript_id: Id.transcript,
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
    id: string; subject: string; startDateTime: string; endDateTime: string; joinUrl: string;
  }>>;
  getOnlineMeetingAsync(meetingId: string | number): Promise<{
    id: string; subject: string; startDateTime: string; endDateTime: string; joinUrl: string;
    participants: unknown;
  } | undefined>;
  listMeetingRecordingsAsync(meetingId: string | number): Promise<Array<{
    id: string; createdDateTime: string; recordingContentUrl: string;
  }>>;
  downloadMeetingRecordingAsync(recordingId: string | number, outputPath: string): Promise<string>;
  listMeetingTranscriptsAsync(meetingId: string | number): Promise<Array<{
    id: string; createdDateTime: string; contentUrl: string;
  }>>;
  getMeetingTranscriptContentAsync(transcriptId: string | number, format?: string): Promise<string>;
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
        text: JSON.stringify({ meetings, next: nextActionFor('onlineMeeting') ?? undefined }, null, 2),
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

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2)
// =============================================================================

/**
 * Registry tool definitions for the meetings domain.
 */
export function meetingsToolDefinitions(): ToolDefinition[] {
  const tools = (ctx: ToolContext): MeetingsTools => requireGraphToolset(ctx, 'meetings');

  return [
    defineTool({
      name: 'list_online_meetings',
      description: 'List recent online meetings (Teams) for the current user (Graph API)',
      input: ListOnlineMeetingsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['meetings'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listOnlineMeetings(params),
    }),
    defineTool({
      name: 'get_online_meeting',
      description: 'Get details for a specific online meeting including participants (Graph API)',
      input: GetOnlineMeetingInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['meetings'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getOnlineMeeting(params),
    }),
    defineTool({
      name: 'list_meeting_recordings',
      description: 'List recordings for an online meeting (Graph API)',
      input: ListMeetingRecordingsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['meetings'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listMeetingRecordings(params),
    }),
    defineTool({
      name: 'download_meeting_recording',
      description: 'Download a meeting recording to a local file (Graph API)',
      input: DownloadMeetingRecordingInput,
      // Writes the recording to output_path on local disk — not read-only.
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['meetings'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).downloadMeetingRecording(params),
    }),
    defineTool({
      name: 'list_meeting_transcripts',
      description: 'List transcripts for an online meeting (Graph API)',
      input: ListMeetingTranscriptsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['meetings'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listMeetingTranscripts(params),
    }),
    defineTool({
      name: 'get_meeting_transcript_content',
      description: 'Get the content of a meeting transcript in VTT or plain text format (Graph API)',
      input: GetMeetingTranscriptContentInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['meetings'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getMeetingTranscriptContent(params),
    }),
  ];
}
