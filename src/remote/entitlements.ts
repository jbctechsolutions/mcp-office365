/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Per-user tool entitlements (U6). A user's tool surface is resolved from a
 * config file keyed by Entra `oid`, defaulting to a PINNED, explicit tool list.
 *
 * The default is an explicit allow-list (not a preset expansion) on purpose:
 * because the surface is exactly these names, a server upgrade that adds new
 * tools cannot silently widen any user's access — the pinned list must be edited
 * (a reviewed jp-infrastructure change) for a new tool to become reachable. A
 * contract test asserts every pinned name still exists in the registry so
 * renames/removals surface as a failing build.
 *
 * Config is re-read on change (mtime), so an entitlement edit takes effect on a
 * user's next request without a restart.
 */

import { readFileSync, statSync } from 'node:fs';
import type { Backend, SurfaceOptions } from '../registry/index.js';

/**
 * PINNED v1 default tool surface (R7): mail, calendar, files/SharePoint, and
 * Planner — deliberately excluding shared-mailbox, mail-rules, and file-download
 * / photo tools. Generated from the registry; update via a reviewed change when
 * the default surface should shift (the contract test flags registry drift).
 */
export const DEFAULT_TOOL_SURFACE: readonly string[] = [
  'add_draft_attachment', 'add_draft_inline_image', 'check_availability', 'check_new_emails',
  'clear_email_flag', 'confirm_archive_email', 'confirm_batch_operation', 'confirm_delete_bucket',
  'confirm_delete_calendar_permission', 'confirm_delete_category', 'confirm_delete_drive_item',
  'confirm_delete_email', 'confirm_delete_event', 'confirm_delete_focused_override',
  'confirm_delete_folder', 'confirm_delete_list_item', 'confirm_delete_planner_task',
  'confirm_empty_folder', 'confirm_forward_email', 'confirm_junk_email', 'confirm_move_email',
  'confirm_reply_email', 'confirm_send_draft', 'confirm_send_email', 'confirm_upload_file',
  'confirm_upload_library_file', 'create_bucket', 'create_calendar_group',
  'create_calendar_permission', 'create_category', 'create_draft', 'create_event',
  'create_focused_override', 'create_folder', 'create_library_folder', 'create_list',
  'create_list_item', 'create_plan', 'create_planner_task', 'create_sharing_link', 'delete_event',
  'find_meeting_times', 'forward_as_draft', 'generate_burndown_chart', 'generate_gantt_chart',
  'generate_kanban_board', 'generate_plan_summary', 'get_automatic_replies', 'get_drive_item',
  'get_email', 'get_emails', 'get_event', 'get_list', 'get_list_item', 'get_mail_tips',
  'get_mailbox_settings', 'get_message_headers', 'get_message_mime', 'get_plan', 'get_planner_task',
  'get_planner_task_details', 'get_signature', 'get_site', 'get_unread_count', 'list_attachments',
  'list_buckets', 'list_calendar_groups', 'list_calendar_permissions', 'list_calendars',
  'list_categories', 'list_conversation', 'list_document_libraries', 'list_drafts',
  'list_drive_items', 'list_emails', 'list_event_instances', 'list_events', 'list_focused_overrides',
  'list_folders', 'list_library_items', 'list_list_columns', 'list_list_items', 'list_lists',
  'list_my_planner_tasks', 'list_planner_tasks', 'list_plans', 'list_recent_files', 'list_room_lists',
  'list_rooms', 'list_shared_with_me', 'list_sites', 'mark_email_read', 'mark_email_unread',
  'move_folder', 'prepare_archive_email', 'prepare_batch_delete_emails', 'prepare_batch_move_emails',
  'prepare_delete_bucket', 'prepare_delete_calendar_permission', 'prepare_delete_category',
  'prepare_delete_drive_item', 'prepare_delete_email', 'prepare_delete_event',
  'prepare_delete_focused_override', 'prepare_delete_folder', 'prepare_delete_list_item',
  'prepare_delete_planner_task', 'prepare_empty_folder', 'prepare_forward_email',
  'prepare_junk_email', 'prepare_move_email', 'prepare_reply_email', 'prepare_send_draft',
  'prepare_send_email', 'prepare_upload_file', 'prepare_upload_library_file', 'rename_folder',
  'reply_as_draft', 'reset_change_tracking', 'respond_to_event', 'search_drive_items',
  'search_emails', 'search_emails_advanced', 'search_events', 'search_sites', 'send_email',
  'set_automatic_replies', 'set_email_categories', 'set_email_flag', 'set_email_importance',
  'set_signature', 'update_bucket', 'update_draft', 'update_event', 'update_list_item',
  'update_mailbox_settings', 'update_plan', 'update_planner_task', 'update_planner_task_details',
  'what_changed',
];

/** A single user's entitlement entry. */
export interface UserEntitlement {
  /** Full surface (Joel's parity with local stdio) — ignores allow/default. */
  readonly fullAccess?: boolean;
  /** Explicit tool-name allow-list; overrides the pinned default. */
  readonly allow?: readonly string[];
  /** Tool names to remove (applies to full, default, or custom allow). */
  readonly exclude?: readonly string[];
}

/** Entitlement config file shape (keyed by Entra oid). */
export interface EntitlementConfig {
  readonly version?: number;
  readonly users?: Record<string, UserEntitlement>;
}

/** Resolves a user's tool-surface options. */
export interface EntitlementResolver {
  resolve(oid: string, backend: Backend): SurfaceOptions;
}

/**
 * Builds an entitlement resolver. When `configPath` is set, the file is read and
 * re-read on mtime change (edits take effect on the next request). When unset,
 * every user gets the pinned default surface.
 */
export function createEntitlementResolver(configPath?: string): EntitlementResolver {
  let cached: EntitlementConfig = {};
  let cachedMtimeMs = -1;

  function currentConfig(): EntitlementConfig {
    if (configPath == null) return {};
    let mtimeMs: number;
    try {
      mtimeMs = statSync(configPath).mtimeMs;
    } catch {
      // File missing/unreadable → fail safe to the pinned default for everyone.
      cached = {};
      cachedMtimeMs = -1;
      return cached;
    }
    if (mtimeMs !== cachedMtimeMs) {
      try {
        cached = parseConfig(readFileSync(configPath, 'utf-8'));
        cachedMtimeMs = mtimeMs;
      } catch (e) {
        // Malformed config must not widen access — keep the last-good (or default)
        // and warn; a bad edit fails safe rather than exposing the full surface.
        process.stderr.write(
          `[mcp-office365] entitlement config parse failed (${e instanceof Error ? e.message : String(e)}); ` +
            `keeping previous config.\n`,
        );
      }
    }
    return cached;
  }

  return {
    resolve(oid: string, backend: Backend): SurfaceOptions {
      const entry = currentConfig().users?.[oid];
      if (entry?.fullAccess === true) {
        return { backend, ...(entry.exclude != null ? { exclude: entry.exclude } : {}) };
      }
      return {
        backend,
        allow: entry?.allow ?? DEFAULT_TOOL_SURFACE,
        ...(entry?.exclude != null ? { exclude: entry.exclude } : {}),
      };
    },
  };
}

/** Parses + shallow-validates the entitlement config JSON. */
function parseConfig(raw: string): EntitlementConfig {
  const parsed: unknown = JSON.parse(raw);
  if (parsed == null || typeof parsed !== 'object') {
    throw new Error('entitlement config must be a JSON object');
  }
  const obj = parsed as { version?: unknown; users?: unknown };
  const users: Record<string, UserEntitlement> = {};
  if (obj.users != null) {
    if (typeof obj.users !== 'object') throw new Error('`users` must be an object keyed by oid');
    for (const [oid, value] of Object.entries(obj.users as Record<string, unknown>)) {
      const v = value as UserEntitlement;
      if (v.allow != null && !Array.isArray(v.allow)) throw new Error(`users.${oid}.allow must be an array`);
      if (v.exclude != null && !Array.isArray(v.exclude)) throw new Error(`users.${oid}.exclude must be an array`);
      users[oid] = {
        ...(v.fullAccess === true ? { fullAccess: true } : {}),
        ...(v.allow != null ? { allow: v.allow } : {}),
        ...(v.exclude != null ? { exclude: v.exclude } : {}),
      };
    }
  }
  return { ...(typeof obj.version === 'number' ? { version: obj.version } : {}), users };
}
