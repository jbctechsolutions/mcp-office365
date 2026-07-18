/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Canonical per-entity Zod id schemas (U6).
 *
 * Every tool that accepts a durable-ID param should reference the shared schema
 * for its entity (`Id.task`, `Id.plan`, …) instead of an ad-hoc
 * `z.string().min(1).describe(...)`. This gives one place to define:
 *
 * - **normalization** — `.trim()` so a copy/pasted id with stray whitespace
 *   still resolves; `.min(1)` so an empty/whitespace-only id fails fast.
 * - **a consistent, prefix-named description** — derived from the token tables
 *   (`prefixForEntity`), so the param doc always names the token shape (e.g.
 *   "a `td_` token") and where to get it.
 *
 * Deliberately NOT here: classification of the id (numeric-vs-token, wrong
 * entity, alias store lookup). That stays in {@link resolveId} at execution
 * time, which is the single source of truth and emits the specific typed codes
 * (`NUMERIC_ID_UNSUPPORTED`, `ID_ENTITY_MISMATCH`, `ID_UNKNOWN`) agents key on.
 * The schema is string-only (a JSON-number legacy id is a type error); a numeric
 * *string* still passes and reaches `resolveId` for the specific message.
 */

import { z } from 'zod';
import { prefixForEntity, type EntityType } from './token.js';

/** Human label + the tools that mint an entity's id, for the description. */
interface EntityMeta {
  readonly label: string;
  readonly from: string;
}

/**
 * Per-entity description metadata. Only entities actually exposed as tool id
 * params need an entry; others fall back to a generic description. Exported so a
 * contract test can cross-check every `from` tool name against the registry,
 * catching drift when a tool is renamed.
 */
export const ENTITY_META: Partial<Record<EntityType, EntityMeta>> = {
  message: { label: 'email message', from: 'list_emails / search_emails' },
  event: { label: 'calendar event', from: 'list_events / search_events' },
  contact: { label: 'contact', from: 'list_contacts / search_contacts' },
  folder: { label: 'mail folder', from: 'list_folders' },
  driveItem: { label: 'OneDrive item', from: 'list_drive_items / search_drive_items' },
  task: { label: 'To Do task', from: 'list_tasks / search_tasks' },
  taskList: { label: 'To Do task list', from: 'list_task_lists' },
  plan: { label: 'Planner plan', from: 'list_plans' },
  plannerBucket: { label: 'Planner bucket', from: 'list_buckets' },
  plannerTask: { label: 'Planner task', from: 'list_planner_tasks' },
  plannerTaskMessage: { label: 'Planner task comment', from: 'list_planner_task_messages' },
  attachment: { label: 'email attachment', from: 'list_attachments' },
  checklistItem: { label: 'task checklist item', from: 'list_checklist_items' },
  linkedResource: { label: 'task linked resource', from: 'list_linked_resources' },
  taskAttachment: { label: 'task attachment', from: 'list_task_attachments' },
  chat: { label: 'Teams chat', from: 'list_chats / find_chat' },
  team: { label: 'team', from: 'list_teams' },
  channel: { label: 'Teams channel', from: 'list_channels' },
  chatMessage: { label: 'chat message', from: 'list_chat_messages' },
  channelMessage: { label: 'channel message', from: 'list_channel_messages' },
  contactFolder: { label: 'contact folder', from: 'list_contact_folders' },
  mailRule: { label: 'inbox rule', from: 'list_mail_rules' },
  category: { label: 'master category', from: 'list_categories' },
  focusedOverride: { label: 'focused-inbox override', from: 'list_focused_overrides' },
  calendarPermission: { label: 'calendar permission', from: 'list_calendar_permissions' },
  onlineMeeting: { label: 'online meeting', from: 'list_online_meetings' },
  recording: { label: 'meeting recording', from: 'list_meeting_recordings' },
  transcript: { label: 'meeting transcript', from: 'list_meeting_transcripts' },
  site: { label: 'SharePoint site', from: 'list_sites / search_sites' },
  documentLibrary: { label: 'document library', from: 'list_document_libraries' },
  libraryDriveItem: { label: 'library item', from: 'list_library_items' },
  sharePointList: { label: 'SharePoint list', from: 'list_lists' },
  sharePointListItem: { label: 'SharePoint list item', from: 'list_list_items' },
  noteNotebook: { label: 'OneNote notebook', from: 'list_notebooks' },
  noteSection: { label: 'OneNote section', from: 'list_note_sections' },
  notePage: { label: 'OneNote page', from: 'list_note_pages / search_note_pages' },
};

/** Builds the canonical, prefix-named description for an entity id. */
export function describeId(entityType: EntityType): string {
  const prefix = prefixForEntity(entityType);
  const meta = ENTITY_META[entityType];
  const label = meta?.label ?? entityType;
  const source = meta != null ? ` from ${meta.from}` : '';
  return `Durable ${label} ID — a \`${prefix}_\` token${source}. A raw Graph id is also accepted.`;
}

/**
 * The canonical required id schema for an entity: a trimmed, non-empty string
 * with a prefix-named description. Classification stays in `resolveId`.
 */
export function idSchema(entityType: EntityType): z.ZodString {
  return z.string().trim().min(1).describe(describeId(entityType));
}

/** The canonical optional id schema for an entity. */
export function optionalIdSchema(entityType: EntityType): z.ZodOptional<z.ZodString> {
  return idSchema(entityType).optional();
}

/**
 * Canonical required id schemas keyed by entity. Tools reference these directly
 * (`Id.task`, `Id.plan`, …). A per-call `.describe()` may override the default
 * description where a param needs extra context.
 */
export const Id = {
  message: idSchema('message'),
  event: idSchema('event'),
  contact: idSchema('contact'),
  folder: idSchema('folder'),
  driveItem: idSchema('driveItem'),
  task: idSchema('task'),
  taskList: idSchema('taskList'),
  plan: idSchema('plan'),
  plannerBucket: idSchema('plannerBucket'),
  plannerTask: idSchema('plannerTask'),
  plannerTaskMessage: idSchema('plannerTaskMessage'),
  attachment: idSchema('attachment'),
  checklistItem: idSchema('checklistItem'),
  linkedResource: idSchema('linkedResource'),
  taskAttachment: idSchema('taskAttachment'),
  chat: idSchema('chat'),
  team: idSchema('team'),
  channel: idSchema('channel'),
  chatMessage: idSchema('chatMessage'),
  channelMessage: idSchema('channelMessage'),
  contactFolder: idSchema('contactFolder'),
  mailRule: idSchema('mailRule'),
  category: idSchema('category'),
  focusedOverride: idSchema('focusedOverride'),
  calendarPermission: idSchema('calendarPermission'),
  onlineMeeting: idSchema('onlineMeeting'),
  recording: idSchema('recording'),
  transcript: idSchema('transcript'),
  site: idSchema('site'),
  documentLibrary: idSchema('documentLibrary'),
  libraryDriveItem: idSchema('libraryDriveItem'),
  sharePointList: idSchema('sharePointList'),
  sharePointListItem: idSchema('sharePointListItem'),
  noteNotebook: idSchema('noteNotebook'),
  noteSection: idSchema('noteSection'),
  notePage: idSchema('notePage'),
} as const satisfies Record<string, z.ZodString>;
