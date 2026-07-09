/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Next-action hints (U6).
 *
 * List/search/create tools that mint durable ids can attach a single short
 * `next` string to their JSON result, suggesting the natural follow-up tool(s)
 * and naming the token prefix the caller will feed forward — e.g. list_plans →
 * "Use get_plan or list_buckets with a returned `pl_` id." The wording and the
 * prefix are derived centrally (from the token tables) so they stay consistent
 * as the tool surface evolves.
 *
 * Intentionally minimal: one top-level `next` key per result (never per row),
 * added only where a follow-up is genuinely useful. Get and write results skip it.
 */

import { prefixForEntity, type EntityType } from './token.js';

/**
 * Follow-up tool names suggested for a returned id of each entity. Exported so a
 * contract test can cross-check every named tool against the registry.
 */
export const FOLLOWUP_TOOLS: Partial<Record<EntityType, string>> = {
  task: 'get_task, update_task, or complete_task',
  taskList: 'list_tasks or create_task',
  plan: 'get_plan or list_buckets',
  plannerBucket: 'list_planner_tasks or create_planner_task',
  plannerTask: 'get_planner_task or update_planner_task',
  message: 'get_email, list_attachments, or reply_as_draft',
  folder: 'list_emails or search_emails',
  contact: 'get_contact or update_contact',
  event: 'get_event or respond_to_event',
  driveItem: 'get_drive_item or download_file',
  site: 'list_document_libraries',
  documentLibrary: 'list_library_items',
  libraryDriveItem: 'download_library_file',
  sharePointList: 'list_list_items',
  sharePointListItem: 'get_list_item or update_list_item',
  team: 'list_channels',
  channel: 'list_channel_messages',
  chat: 'list_chat_messages',
  onlineMeeting: 'get_online_meeting, list_meeting_recordings, or list_meeting_transcripts',
  noteNotebook: 'list_note_sections',
  noteSection: 'list_note_pages',
  notePage: 'get_note_page',
};

/**
 * The next-action hint sentence for an entity, or null when none is defined.
 * The `_` suffix on the prefix mirrors the token shape callers see.
 */
export function nextActionFor(entityType: EntityType): string | null {
  const tools = FOLLOWUP_TOOLS[entityType];
  if (tools == null) {
    return null;
  }
  return `Use ${tools} with a returned \`${prefixForEntity(entityType)}_\` id.`;
}
