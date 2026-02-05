#!/usr/bin/env node
/**
 * Outlook MCP Server
 *
 * A Model Context Protocol server that provides read-only access to
 * Outlook for Mac via AppleScript or Microsoft Graph API.
 *
 * Backend selection:
 * - Set USE_GRAPH_API=1 to use Microsoft Graph API (required for new Outlook)
 * - Otherwise, AppleScript is used (works with classic Outlook)
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type Tool,
} from '@modelcontextprotocol/sdk/types.js';

import {
  createAppleScriptRepository,
  createAppleScriptContentReaders,
  createAccountRepository,
  createCalendarWriter,
  createCalendarManager,
  createMailSender,
  isOutlookRunning,
  type IAccountRepository,
  type ICalendarWriter,
  type ICalendarManager,
  type IMailSender,
  type RecurrenceConfig,
} from './applescript/index.js';
import {
  createGraphRepository,
  createGraphContentReadersWithClient,
  isAuthenticated,
  GraphMailboxAdapter,
  type GraphRepository,
  type GraphContentReaders,
} from './graph/index.js';
import { createMailTools } from './tools/mail.js';
import { createCalendarTools } from './tools/calendar.js';
import { createContactsTools } from './tools/contacts.js';
import { createTasksTools } from './tools/tasks.js';
import { createNotesTools } from './tools/notes.js';
import { createMailboxOrganizationTools } from './tools/mailbox-organization.js';
import {
  ListEmailsInput,
  SearchEmailsInput,
  GetEmailInput,
  GetUnreadCountInput,
  ListCalendarsInput,
  ListEventsInput,
  GetEventInput,
  SearchEventsInput,
  CreateEventInput,
  RespondToEventInput,
  ListContactsInput,
  SearchContactsInput,
  GetContactInput,
  ListTasksInput,
  SearchTasksInput,
  GetTaskInput,
  ListNotesInput,
  GetNoteInput,
  SearchNotesInput,
  PrepareDeleteEmailInput,
  ConfirmDeleteEmailInput,
  PrepareMoveEmailInput,
  ConfirmMoveEmailInput,
  PrepareArchiveEmailInput,
  ConfirmArchiveEmailInput,
  PrepareJunkEmailInput,
  ConfirmJunkEmailInput,
  PrepareDeleteFolderInput,
  ConfirmDeleteFolderInput,
  PrepareEmptyFolderInput,
  ConfirmEmptyFolderInput,
  PrepareBatchDeleteEmailsInput,
  PrepareBatchMoveEmailsInput,
  ConfirmBatchOperationInput,
  MarkEmailReadInput,
  MarkEmailUnreadInput,
  SetEmailFlagInput,
  ClearEmailFlagInput,
  SetEmailCategoriesInput,
  CreateFolderInput,
  RenameFolderInput,
  MoveFolderInput,
} from './tools/index.js';
import { ApprovalTokenManager } from './approval/index.js';
import type { CreateEventResult } from './tools/index.js';
import {
  wrapError,
  OutlookNotRunningError,
  GraphAuthRequiredError,
  GraphError,
} from './utils/errors.js';

// =============================================================================
// Backend Configuration
// =============================================================================

/**
 * Determines if we should use the Microsoft Graph API backend.
 */
function shouldUseGraphApi(): boolean {
  return process.env['USE_GRAPH_API'] === '1' || process.env['USE_GRAPH_API'] === 'true';
}

// =============================================================================
// Tool Definitions
// =============================================================================

const TOOLS: Tool[] = [
  // Account tools
  {
    name: 'list_accounts',
    description: 'List all Exchange accounts configured in Outlook with their details',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
  },
  // Mail tools
  {
    name: 'list_folders',
    description: 'List all mail folders with message and unread counts. Can filter by account.',
    inputSchema: {
      type: 'object',
      properties: {
        account_id: {
          oneOf: [
            { type: 'number', description: 'Specific account ID' },
            { type: 'array', items: { type: 'number' }, description: 'Multiple account IDs' },
            { type: 'string', enum: ['all'], description: 'All accounts' },
          ],
          description: 'Account filter: number (specific account), array (multiple accounts), "all" (all accounts), or omit for default account',
        },
      },
      required: [],
    },
  },
  {
    name: 'list_emails',
    description: 'List emails in a folder with pagination',
    inputSchema: {
      type: 'object',
      properties: {
        folder_id: {
          type: 'number',
          description: 'The folder ID to list emails from',
        },
        limit: {
          type: 'number',
          description: 'Maximum number of emails to return (1-100, default 50)',
          default: 50,
        },
        offset: {
          type: 'number',
          description: 'Number of emails to skip (default 0)',
          default: 0,
        },
        unread_only: {
          type: 'boolean',
          description: 'Only return unread emails (default false)',
          default: false,
        },
      },
      required: ['folder_id'],
    },
  },
  {
    name: 'search_emails',
    description: 'Search emails by subject, sender, or content',
    inputSchema: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'Search query',
        },
        folder_id: {
          type: 'number',
          description: 'Optional folder ID to limit search to',
        },
        limit: {
          type: 'number',
          description: 'Maximum number of emails to return (1-100, default 50)',
          default: 50,
        },
      },
      required: ['query'],
    },
  },
  {
    name: 'get_email',
    description: 'Get full email details including body',
    inputSchema: {
      type: 'object',
      properties: {
        email_id: {
          type: 'number',
          description: 'The email ID to retrieve',
        },
        include_body: {
          type: 'boolean',
          description: 'Include the email body (default true)',
          default: true,
        },
        strip_html: {
          type: 'boolean',
          description: 'Strip HTML from the body (default true)',
          default: true,
        },
      },
      required: ['email_id'],
    },
  },
  {
    name: 'get_unread_count',
    description: 'Get unread email count',
    inputSchema: {
      type: 'object',
      properties: {
        folder_id: {
          type: 'number',
          description: 'Optional folder ID to get unread count for',
        },
      },
      required: [],
    },
  },
  // Calendar tools
  {
    name: 'list_calendars',
    description: 'List all calendar folders',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
  },
  {
    name: 'list_events',
    description: 'List calendar events with optional date range filtering',
    inputSchema: {
      type: 'object',
      properties: {
        calendar_id: {
          type: 'number',
          description: 'Optional calendar folder ID',
        },
        start_date: {
          type: 'string',
          description: 'Start date filter (ISO 8601 format)',
        },
        end_date: {
          type: 'string',
          description: 'End date filter (ISO 8601 format)',
        },
        limit: {
          type: 'number',
          description: 'Maximum number of events to return (1-100, default 50)',
          default: 50,
        },
      },
      required: [],
    },
  },
  {
    name: 'get_event',
    description: 'Get event details',
    inputSchema: {
      type: 'object',
      properties: {
        event_id: {
          type: 'number',
          description: 'The event ID to retrieve',
        },
      },
      required: ['event_id'],
    },
  },
  {
    name: 'search_events',
    description: 'Search events by title',
    inputSchema: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'Search query',
        },
        limit: {
          type: 'number',
          description: 'Maximum number of events to return (1-100, default 50)',
          default: 50,
        },
      },
      required: ['query'],
    },
  },
  {
    name: 'create_event',
    description: 'Create a new calendar event in Outlook',
    inputSchema: {
      type: 'object',
      properties: {
        title: {
          type: 'string',
          description: 'Event title/subject',
        },
        start_date: {
          type: 'string',
          description: 'Start date in ISO 8601 UTC format (e.g., 2026-02-03T14:00:00Z). Times are interpreted as UTC.',
        },
        end_date: {
          type: 'string',
          description: 'End date in ISO 8601 UTC format (e.g., 2026-02-03T15:00:00Z). Times are interpreted as UTC.',
        },
        calendar_id: {
          type: 'number',
          description: 'Optional calendar ID to create the event in (defaults to primary calendar)',
        },
        location: {
          type: 'string',
          description: 'Event location',
        },
        description: {
          type: 'string',
          description: 'Event description/body text',
        },
        is_all_day: {
          type: 'boolean',
          description: 'Whether this is an all-day event (default false)',
          default: false,
        },
        recurrence: {
          type: 'object',
          description: 'Recurrence pattern to make this a repeating event',
          properties: {
            frequency: {
              type: 'string',
              enum: ['daily', 'weekly', 'monthly', 'yearly'],
              description: 'How often the event repeats',
            },
            interval: {
              type: 'number',
              description: 'Number of frequency units between occurrences (default 1)',
              default: 1,
            },
            days_of_week: {
              type: 'array',
              items: { type: 'string', enum: ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'] },
              description: 'Days of the week for weekly recurrence (e.g., ["monday", "wednesday"])',
            },
            day_of_month: {
              type: 'number',
              description: 'Day of the month for monthly recurrence (e.g., 15)',
            },
            week_of_month: {
              type: 'string',
              enum: ['first', 'second', 'third', 'fourth', 'last'],
              description: 'Week of the month for ordinal monthly recurrence (e.g., "third" for 3rd Thursday)',
            },
            day_of_week_monthly: {
              type: 'string',
              enum: ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'],
              description: 'Day of week for ordinal monthly recurrence (used with week_of_month)',
            },
            end: {
              type: 'object',
              description: 'When the recurrence ends (default: no end)',
              oneOf: [
                { properties: { type: { const: 'no_end' } }, required: ['type'] },
                { properties: { type: { const: 'end_date' }, date: { type: 'string', description: 'End date in ISO 8601 format' } }, required: ['type', 'date'] },
                { properties: { type: { const: 'end_after_count' }, count: { type: 'number', description: 'Number of occurrences' } }, required: ['type', 'count'] },
              ],
            },
          },
          required: ['frequency'],
        },
      },
      required: ['title', 'start_date', 'end_date'],
    },
  },
  {
    name: 'respond_to_event',
    description: 'Respond to a meeting invitation (accept, decline, or tentative). Updates your response status and optionally notifies the organizer.',
    inputSchema: {
      type: 'object',
      properties: {
        event_id: {
          type: 'number',
          description: 'The event ID to respond to',
        },
        response: {
          type: 'string',
          enum: ['accept', 'decline', 'tentative'],
          description: 'Your response to the invitation',
        },
        send_response: {
          type: 'boolean',
          description: 'Whether to send response to organizer (default true)',
          default: true,
        },
        comment: {
          type: 'string',
          description: 'Optional comment to include with response',
        },
      },
      required: ['event_id', 'response'],
    },
  },
  {
    name: 'delete_event',
    description: 'Delete a calendar event. For recurring events, you can delete a single instance or the entire series.',
    inputSchema: {
      type: 'object',
      properties: {
        event_id: {
          type: 'number',
          description: 'The event ID to delete',
        },
        apply_to: {
          type: 'string',
          enum: ['this_instance', 'all_in_series'],
          description: 'For recurring events: delete single instance or entire series (default: this_instance)',
          default: 'this_instance',
        },
      },
      required: ['event_id'],
    },
  },
  {
    name: 'update_event',
    description: 'Update a calendar event. All fields are optional - only specified fields will be updated. For recurring events, you can update a single instance or the entire series.',
    inputSchema: {
      type: 'object',
      properties: {
        event_id: {
          type: 'number',
          description: 'The event ID to update',
        },
        apply_to: {
          type: 'string',
          enum: ['this_instance', 'all_in_series'],
          description: 'For recurring events: update single instance or entire series (default: this_instance)',
          default: 'this_instance',
        },
        title: {
          type: 'string',
          description: 'New event title',
        },
        start_date: {
          type: 'string',
          description: 'New start date (ISO 8601 UTC format)',
        },
        end_date: {
          type: 'string',
          description: 'New end date (ISO 8601 UTC format)',
        },
        location: {
          type: 'string',
          description: 'New location',
        },
        description: {
          type: 'string',
          description: 'New description',
        },
        is_all_day: {
          type: 'boolean',
          description: 'Whether event is all day',
        },
      },
      required: ['event_id'],
    },
  },
  // Contact tools
  {
    name: 'list_contacts',
    description: 'List contacts with pagination',
    inputSchema: {
      type: 'object',
      properties: {
        limit: {
          type: 'number',
          description: 'Maximum number of contacts to return (1-100, default 50)',
          default: 50,
        },
        offset: {
          type: 'number',
          description: 'Number of contacts to skip (default 0)',
          default: 0,
        },
      },
      required: [],
    },
  },
  {
    name: 'search_contacts',
    description: 'Search contacts by name',
    inputSchema: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'Search query',
        },
        limit: {
          type: 'number',
          description: 'Maximum number of contacts to return (1-100, default 50)',
          default: 50,
        },
      },
      required: ['query'],
    },
  },
  {
    name: 'get_contact',
    description: 'Get contact details',
    inputSchema: {
      type: 'object',
      properties: {
        contact_id: {
          type: 'number',
          description: 'The contact ID to retrieve',
        },
      },
      required: ['contact_id'],
    },
  },
  // Task tools
  {
    name: 'list_tasks',
    description: 'List tasks with pagination and filtering',
    inputSchema: {
      type: 'object',
      properties: {
        limit: {
          type: 'number',
          description: 'Maximum number of tasks to return (1-100, default 50)',
          default: 50,
        },
        offset: {
          type: 'number',
          description: 'Number of tasks to skip (default 0)',
          default: 0,
        },
        include_completed: {
          type: 'boolean',
          description: 'Include completed tasks (default true)',
          default: true,
        },
      },
      required: [],
    },
  },
  {
    name: 'search_tasks',
    description: 'Search tasks by name',
    inputSchema: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'Search query',
        },
        limit: {
          type: 'number',
          description: 'Maximum number of tasks to return (1-100, default 50)',
          default: 50,
        },
      },
      required: ['query'],
    },
  },
  {
    name: 'get_task',
    description: 'Get task details',
    inputSchema: {
      type: 'object',
      properties: {
        task_id: {
          type: 'number',
          description: 'The task ID to retrieve',
        },
      },
      required: ['task_id'],
    },
  },
  // Note tools
  {
    name: 'list_notes',
    description: 'List notes with pagination',
    inputSchema: {
      type: 'object',
      properties: {
        limit: {
          type: 'number',
          description: 'Maximum number of notes to return (1-100, default 50)',
          default: 50,
        },
        offset: {
          type: 'number',
          description: 'Number of notes to skip (default 0)',
          default: 0,
        },
      },
      required: [],
    },
  },
  {
    name: 'get_note',
    description: 'Get note details',
    inputSchema: {
      type: 'object',
      properties: {
        note_id: {
          type: 'number',
          description: 'The note ID to retrieve',
        },
      },
      required: ['note_id'],
    },
  },
  {
    name: 'search_notes',
    description: 'Search notes by content',
    inputSchema: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'Search query',
        },
        limit: {
          type: 'number',
          description: 'Maximum number of notes to return (1-100, default 50)',
          default: 50,
        },
      },
      required: ['query'],
    },
  },
  // =========================================================================
  // Mailbox Organization — Destructive (Two-Phase Approval)
  // =========================================================================
  {
    name: 'prepare_delete_email',
    description: 'Prepare to delete an email (move to trash). Returns a preview and approval token. Call confirm_delete_email to execute.',
    inputSchema: {
      type: 'object',
      properties: {
        email_id: { type: 'number', description: 'The email ID to delete' },
      },
      required: ['email_id'],
    },
  },
  {
    name: 'confirm_delete_email',
    description: 'Confirm deletion of an email using a token from prepare_delete_email',
    inputSchema: {
      type: 'object',
      properties: {
        token_id: { type: 'string', description: 'The approval token from prepare_delete_email' },
        email_id: { type: 'number', description: 'The email ID to delete' },
      },
      required: ['token_id', 'email_id'],
    },
  },
  {
    name: 'prepare_move_email',
    description: 'Prepare to move an email to another folder. Returns a preview and approval token. Call confirm_move_email to execute.',
    inputSchema: {
      type: 'object',
      properties: {
        email_id: { type: 'number', description: 'The email ID to move' },
        destination_folder_id: { type: 'number', description: 'The destination folder ID' },
      },
      required: ['email_id', 'destination_folder_id'],
    },
  },
  {
    name: 'confirm_move_email',
    description: 'Confirm moving an email using a token from prepare_move_email',
    inputSchema: {
      type: 'object',
      properties: {
        token_id: { type: 'string', description: 'The approval token from prepare_move_email' },
        email_id: { type: 'number', description: 'The email ID to move' },
      },
      required: ['token_id', 'email_id'],
    },
  },
  {
    name: 'prepare_archive_email',
    description: 'Prepare to archive an email. Returns a preview and approval token. Call confirm_archive_email to execute.',
    inputSchema: {
      type: 'object',
      properties: {
        email_id: { type: 'number', description: 'The email ID to archive' },
      },
      required: ['email_id'],
    },
  },
  {
    name: 'confirm_archive_email',
    description: 'Confirm archiving an email using a token from prepare_archive_email',
    inputSchema: {
      type: 'object',
      properties: {
        token_id: { type: 'string', description: 'The approval token from prepare_archive_email' },
        email_id: { type: 'number', description: 'The email ID to archive' },
      },
      required: ['token_id', 'email_id'],
    },
  },
  {
    name: 'prepare_junk_email',
    description: 'Prepare to mark an email as junk. Returns a preview and approval token. Call confirm_junk_email to execute.',
    inputSchema: {
      type: 'object',
      properties: {
        email_id: { type: 'number', description: 'The email ID to mark as junk' },
      },
      required: ['email_id'],
    },
  },
  {
    name: 'confirm_junk_email',
    description: 'Confirm marking an email as junk using a token from prepare_junk_email',
    inputSchema: {
      type: 'object',
      properties: {
        token_id: { type: 'string', description: 'The approval token from prepare_junk_email' },
        email_id: { type: 'number', description: 'The email ID to mark as junk' },
      },
      required: ['token_id', 'email_id'],
    },
  },
  {
    name: 'prepare_delete_folder',
    description: 'Prepare to delete a mail folder. Returns a preview and approval token. Call confirm_delete_folder to execute.',
    inputSchema: {
      type: 'object',
      properties: {
        folder_id: { type: 'number', description: 'The folder ID to delete' },
      },
      required: ['folder_id'],
    },
  },
  {
    name: 'confirm_delete_folder',
    description: 'Confirm deletion of a folder using a token from prepare_delete_folder',
    inputSchema: {
      type: 'object',
      properties: {
        token_id: { type: 'string', description: 'The approval token from prepare_delete_folder' },
        folder_id: { type: 'number', description: 'The folder ID to delete' },
      },
      required: ['token_id', 'folder_id'],
    },
  },
  {
    name: 'prepare_empty_folder',
    description: 'Prepare to empty a mail folder (delete all messages). Returns a preview and approval token. Call confirm_empty_folder to execute.',
    inputSchema: {
      type: 'object',
      properties: {
        folder_id: { type: 'number', description: 'The folder ID to empty' },
      },
      required: ['folder_id'],
    },
  },
  {
    name: 'confirm_empty_folder',
    description: 'Confirm emptying a folder using a token from prepare_empty_folder',
    inputSchema: {
      type: 'object',
      properties: {
        token_id: { type: 'string', description: 'The approval token from prepare_empty_folder' },
        folder_id: { type: 'number', description: 'The folder ID to empty' },
      },
      required: ['token_id', 'folder_id'],
    },
  },
  {
    name: 'prepare_batch_delete_emails',
    description: 'Prepare to delete multiple emails. Returns individual tokens per email so you can selectively confirm. Call confirm_batch_operation to execute.',
    inputSchema: {
      type: 'object',
      properties: {
        email_ids: {
          type: 'array',
          items: { type: 'number' },
          description: 'The email IDs to delete (max 50)',
        },
      },
      required: ['email_ids'],
    },
  },
  {
    name: 'prepare_batch_move_emails',
    description: 'Prepare to move multiple emails. Returns individual tokens per email so you can selectively confirm. Call confirm_batch_operation to execute.',
    inputSchema: {
      type: 'object',
      properties: {
        email_ids: {
          type: 'array',
          items: { type: 'number' },
          description: 'The email IDs to move (max 50)',
        },
        destination_folder_id: { type: 'number', description: 'The destination folder ID' },
      },
      required: ['email_ids', 'destination_folder_id'],
    },
  },
  {
    name: 'confirm_batch_operation',
    description: 'Confirm a batch operation using tokens from prepare_batch_delete_emails or prepare_batch_move_emails. You may selectively confirm by omitting tokens.',
    inputSchema: {
      type: 'object',
      properties: {
        tokens: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              token_id: { type: 'string', description: 'The approval token' },
              email_id: { type: 'number', description: 'The email ID' },
            },
            required: ['token_id', 'email_id'],
          },
          description: 'Array of token/email pairs to confirm',
        },
      },
      required: ['tokens'],
    },
  },
  // =========================================================================
  // Mailbox Organization — Low-Risk (No Approval)
  // =========================================================================
  {
    name: 'mark_email_read',
    description: 'Mark an email as read',
    inputSchema: {
      type: 'object',
      properties: {
        email_id: { type: 'number', description: 'The email ID to mark as read' },
      },
      required: ['email_id'],
    },
  },
  {
    name: 'mark_email_unread',
    description: 'Mark an email as unread',
    inputSchema: {
      type: 'object',
      properties: {
        email_id: { type: 'number', description: 'The email ID to mark as unread' },
      },
      required: ['email_id'],
    },
  },
  {
    name: 'set_email_flag',
    description: 'Set a follow-up flag on an email',
    inputSchema: {
      type: 'object',
      properties: {
        email_id: { type: 'number', description: 'The email ID to flag' },
        flag_status: { type: 'number', description: 'Flag status: 0=not flagged, 1=flagged, 2=completed' },
      },
      required: ['email_id', 'flag_status'],
    },
  },
  {
    name: 'clear_email_flag',
    description: 'Clear the follow-up flag from an email',
    inputSchema: {
      type: 'object',
      properties: {
        email_id: { type: 'number', description: 'The email ID to clear the flag from' },
      },
      required: ['email_id'],
    },
  },
  {
    name: 'set_email_categories',
    description: 'Set categories on an email (replaces existing categories)',
    inputSchema: {
      type: 'object',
      properties: {
        email_id: { type: 'number', description: 'The email ID' },
        categories: {
          type: 'array',
          items: { type: 'string' },
          description: 'Categories to set. Use empty array to clear.',
        },
      },
      required: ['email_id', 'categories'],
    },
  },
  // =========================================================================
  // Mailbox Organization — Non-Destructive
  // =========================================================================
  {
    name: 'create_folder',
    description: 'Create a new mail folder',
    inputSchema: {
      type: 'object',
      properties: {
        name: { type: 'string', description: 'Name for the new folder' },
        parent_folder_id: { type: 'number', description: 'Optional parent folder ID (top-level if omitted)' },
      },
      required: ['name'],
    },
  },
  {
    name: 'rename_folder',
    description: 'Rename a mail folder',
    inputSchema: {
      type: 'object',
      properties: {
        folder_id: { type: 'number', description: 'The folder ID to rename' },
        new_name: { type: 'string', description: 'The new folder name' },
      },
      required: ['folder_id', 'new_name'],
    },
  },
  {
    name: 'move_folder',
    description: 'Move a mail folder under a different parent',
    inputSchema: {
      type: 'object',
      properties: {
        folder_id: { type: 'number', description: 'The folder ID to move' },
        destination_parent_id: { type: 'number', description: 'The destination parent folder ID' },
      },
      required: ['folder_id', 'destination_parent_id'],
    },
  },
  // Email sending tool
  {
    name: 'send_email',
    description: 'Send an email with optional CC, BCC, attachments, and HTML formatting. Returns the sent message ID and timestamp.',
    inputSchema: {
      type: 'object',
      properties: {
        to: {
          type: 'array',
          items: { type: 'string' },
          minItems: 1,
          description: 'Recipient email addresses',
        },
        subject: {
          type: 'string',
          minLength: 1,
          description: 'Email subject',
        },
        body: {
          type: 'string',
          description: 'Email body content',
        },
        body_type: {
          type: 'string',
          enum: ['plain', 'html'],
          default: 'plain',
          description: 'Body content type (default: plain)',
        },
        cc: {
          type: 'array',
          items: { type: 'string' },
          description: 'CC recipients',
        },
        bcc: {
          type: 'array',
          items: { type: 'string' },
          description: 'BCC recipients',
        },
        reply_to: {
          type: 'string',
          description: 'Reply-to address',
        },
        attachments: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              path: {
                type: 'string',
                description: 'Absolute file path to attachment',
              },
              name: {
                type: 'string',
                description: 'Display name for attachment',
              },
            },
            required: ['path'],
          },
          description: 'File attachments',
        },
        account_id: {
          type: 'number',
          description: 'Account to send from (optional)',
        },
      },
      required: ['to', 'subject', 'body'],
    },
  },
];

// =============================================================================
// Server Creation
// =============================================================================

/**
 * Creates and configures the MCP server.
 */
export function createServer(): Server {
  const server = new Server(
    {
      name: 'outlook-mcp',
      version: '0.1.0',
    },
    {
      capabilities: {
        tools: {},
      },
    }
  );

  // Determine which backend to use
  const useGraphApi = shouldUseGraphApi();

  // Shared state (used by both backends)
  const tokenManager = new ApprovalTokenManager();

  // Tools and backend state
  let initialized = false;
  let accountRepository: IAccountRepository | null = null;
  let mailTools: ReturnType<typeof createMailTools> | null = null;
  let calendarTools: ReturnType<typeof createCalendarTools> | null = null;
  let contactsTools: ReturnType<typeof createContactsTools> | null = null;
  let tasksTools: ReturnType<typeof createTasksTools> | null = null;
  let notesTools: ReturnType<typeof createNotesTools> | null = null;
  let orgTools: ReturnType<typeof createMailboxOrganizationTools> | null = null;
  let calendarWriter: ICalendarWriter | null = null;
  let calendarManager: ICalendarManager | null = null;
  let mailSender: IMailSender | null = null;

  // Graph-specific state
  let graphRepository: GraphRepository | null = null;
  let graphContentReaders: GraphContentReaders | null = null;

  /**
   * Initializes AppleScript backend.
   */
  function initializeAppleScriptBackend(): void {
    if (!isOutlookRunning()) {
      throw new OutlookNotRunningError();
    }

    const repository = createAppleScriptRepository();
    const contentReaders = createAppleScriptContentReaders();

    accountRepository = createAccountRepository();
    mailTools = createMailTools(repository, contentReaders.email);
    calendarTools = createCalendarTools(repository, contentReaders.event);
    contactsTools = createContactsTools(repository, contentReaders.contact);
    tasksTools = createTasksTools(repository, contentReaders.task);
    notesTools = createNotesTools(repository, contentReaders.note);
    orgTools = createMailboxOrganizationTools(repository, tokenManager);
    calendarWriter = createCalendarWriter();
    calendarManager = createCalendarManager();
    mailSender = createMailSender();

    initialized = true;
  }

  /**
   * Initializes Graph API backend.
   */
  async function initializeGraphBackend(): Promise<void> {
    // Check if already authenticated
    const authenticated = await isAuthenticated();
    if (!authenticated) {
      throw new GraphAuthRequiredError();
    }

    graphRepository = createGraphRepository();
    graphContentReaders = createGraphContentReadersWithClient(graphRepository.getClient());

    const adapter = new GraphMailboxAdapter(graphRepository);
    orgTools = createMailboxOrganizationTools(adapter, tokenManager);

    initialized = true;
  }

  /**
   * Ensures the backend is initialized.
   */
  async function ensureInitialized(): Promise<void> {
    if (initialized) return;

    if (useGraphApi) {
      await initializeGraphBackend();
    } else {
      initializeAppleScriptBackend();
    }
  }

  // Register tool list handler
  server.setRequestHandler(ListToolsRequestSchema, () => {
    return { tools: TOOLS };
  });

  // Register tool call handler (async for Graph API support)
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    try {
      await ensureInitialized();

      // Graph API mode - handle async operations directly
      if (useGraphApi && graphRepository != null) {
        return await handleGraphToolCall(name, args, graphRepository, graphContentReaders!, orgTools!);
      }

      // AppleScript mode - use sync tool interfaces
      return handleAppleScriptToolCall(
        name,
        args,
        accountRepository!,
        mailTools!,
        calendarTools!,
        contactsTools!,
        tasksTools!,
        notesTools!,
        orgTools!,
        calendarWriter,
        calendarManager,
        mailSender
      );
    } catch (error) {
      const wrappedError = wrapError(error, 'An error occurred');
      const message = `${wrappedError.code}: ${wrappedError.message}`;

      return {
        content: [{ type: 'text', text: message }],
        isError: true,
      };
    }
  });

  return server;
}

// =============================================================================
// Account Resolution Helper
// =============================================================================

/**
 * Resolves account_id parameter to an array of account IDs.
 * - undefined → [defaultAccountId]
 * - "all" → all account IDs
 * - number → [number]
 * - number[] → number[]
 */
function resolveAccountIds(
  accountId: number | number[] | 'all' | undefined,
  accountRepository: IAccountRepository
): number[] {
  // Case: undefined → use default account
  if (accountId === undefined) {
    const defaultId = accountRepository.getDefaultAccountId();
    return defaultId !== null ? [defaultId] : [];
  }

  // Case: "all" → use all accounts
  if (accountId === 'all') {
    const accounts = accountRepository.listAccounts();
    return accounts.map(acc => acc.id);
  }

  // Case: single number → return as array
  if (typeof accountId === 'number') {
    return [accountId];
  }

  // Case: array of numbers → return as-is
  if (Array.isArray(accountId)) {
    return accountId;
  }

  // Fallback: default account
  const defaultId = accountRepository.getDefaultAccountId();
  return defaultId !== null ? [defaultId] : [];
}

// =============================================================================
// Shared Mailbox Organization Handler
// =============================================================================

type ToolResult = { content: Array<{ type: string; text: string }>; isError?: boolean };

async function handleOrgToolCall(
  name: string,
  args: unknown,
  orgTools: ReturnType<typeof createMailboxOrganizationTools>
): Promise<ToolResult | null> {
  switch (name) {
    // Destructive (Two-Phase)
    case 'prepare_delete_email': {
      const params = PrepareDeleteEmailInput.parse(args);
      const result = await orgTools.prepareDeleteEmail(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'confirm_delete_email': {
      const params = ConfirmDeleteEmailInput.parse(args);
      const result = await orgTools.confirmDeleteEmail(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'prepare_move_email': {
      const params = PrepareMoveEmailInput.parse(args);
      const result = await orgTools.prepareMoveEmail(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'confirm_move_email': {
      const params = ConfirmMoveEmailInput.parse(args);
      const result = await orgTools.confirmMoveEmail(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'prepare_archive_email': {
      const params = PrepareArchiveEmailInput.parse(args);
      const result = await orgTools.prepareArchiveEmail(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'confirm_archive_email': {
      const params = ConfirmArchiveEmailInput.parse(args);
      const result = await orgTools.confirmArchiveEmail(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'prepare_junk_email': {
      const params = PrepareJunkEmailInput.parse(args);
      const result = await orgTools.prepareJunkEmail(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'confirm_junk_email': {
      const params = ConfirmJunkEmailInput.parse(args);
      const result = await orgTools.confirmJunkEmail(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'prepare_delete_folder': {
      const params = PrepareDeleteFolderInput.parse(args);
      const result = await orgTools.prepareDeleteFolder(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'confirm_delete_folder': {
      const params = ConfirmDeleteFolderInput.parse(args);
      const result = await orgTools.confirmDeleteFolder(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'prepare_empty_folder': {
      const params = PrepareEmptyFolderInput.parse(args);
      const result = await orgTools.prepareEmptyFolder(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'confirm_empty_folder': {
      const params = ConfirmEmptyFolderInput.parse(args);
      const result = await orgTools.confirmEmptyFolder(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'prepare_batch_delete_emails': {
      const params = PrepareBatchDeleteEmailsInput.parse(args);
      const result = await orgTools.prepareBatchDeleteEmails(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'prepare_batch_move_emails': {
      const params = PrepareBatchMoveEmailsInput.parse(args);
      const result = await orgTools.prepareBatchMoveEmails(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'confirm_batch_operation': {
      const params = ConfirmBatchOperationInput.parse(args);
      const result = await orgTools.confirmBatchOperation(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    // Low-Risk
    case 'mark_email_read': {
      const params = MarkEmailReadInput.parse(args);
      const result = await orgTools.markEmailRead(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'mark_email_unread': {
      const params = MarkEmailUnreadInput.parse(args);
      const result = await orgTools.markEmailUnread(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'set_email_flag': {
      const params = SetEmailFlagInput.parse(args);
      const result = await orgTools.setEmailFlag(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'clear_email_flag': {
      const params = ClearEmailFlagInput.parse(args);
      const result = await orgTools.clearEmailFlag(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'set_email_categories': {
      const params = SetEmailCategoriesInput.parse(args);
      const result = await orgTools.setEmailCategories(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    // Non-Destructive
    case 'create_folder': {
      const params = CreateFolderInput.parse(args);
      const result = await orgTools.createFolder(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'rename_folder': {
      const params = RenameFolderInput.parse(args);
      const result = await orgTools.renameFolder(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'move_folder': {
      const params = MoveFolderInput.parse(args);
      const result = await orgTools.moveFolder(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    default:
      return null;
  }
}

// =============================================================================
// AppleScript Tool Handler
// =============================================================================

async function handleAppleScriptToolCall(
  name: string,
  args: unknown,
  accountRepository: IAccountRepository,
  mailTools: ReturnType<typeof createMailTools>,
  calendarTools: ReturnType<typeof createCalendarTools>,
  contactsTools: ReturnType<typeof createContactsTools>,
  tasksTools: ReturnType<typeof createTasksTools>,
  notesTools: ReturnType<typeof createNotesTools>,
  orgTools: ReturnType<typeof createMailboxOrganizationTools>,
  calendarWriter: ICalendarWriter | null,
  calendarManager: ICalendarManager | null,
  mailSender: IMailSender | null
): Promise<ToolResult> {
  // Handle mailbox organization tools (shared between backends)
  const orgResult = await handleOrgToolCall(name, args, orgTools);
  if (orgResult != null) return orgResult;

  switch (name) {
    // Account tools
    case 'list_accounts': {
      const accounts = accountRepository.listAccounts();
      const result = {
        accounts: accounts.map(acc => ({
          id: acc.id,
          name: acc.name,
          email: acc.email,
          type: acc.type,
        })),
      };
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    // Mail tools
    case 'list_folders': {
      const params = args as { account_id?: number | number[] | 'all' } | undefined;
      const accountIds = resolveAccountIds(params?.account_id, accountRepository);

      // If querying multiple accounts, use grouped format
      if (accountIds.length > 1 || params?.account_id === 'all') {
        const foldersWithAccount = accountRepository.listMailFoldersByAccounts(accountIds);
        const accounts = accountRepository.listAccounts();

        // Group folders by account
        const groupedByAccount = accountIds.map(accountId => {
          const account = accounts.find(a => a.id === accountId);
          const folders = foldersWithAccount
            .filter(f => f.accountId === accountId)
            .map(f => ({
              id: f.id,
              name: f.name,
              unreadCount: f.unreadCount,
              messageCount: f.messageCount,
            }));

          return {
            account_id: accountId,
            account_name: account?.name ?? null,
            account_email: account?.email ?? null,
            folders,
          };
        });

        const result = { accounts: groupedByAccount };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      // Single account - use existing format for backward compatibility
      const result = mailTools.listFolders({});
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'list_emails': {
      const params = ListEmailsInput.parse(args);
      const result = mailTools.listEmails(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'search_emails': {
      const params = SearchEmailsInput.parse(args);
      const result = mailTools.searchEmails(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'get_email': {
      const params = GetEmailInput.parse(args);
      const result = mailTools.getEmail(params);
      if (result == null) {
        return { content: [{ type: 'text', text: 'Email not found' }], isError: true };
      }
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'get_unread_count': {
      const params = GetUnreadCountInput.parse(args ?? {});
      const result = mailTools.getUnreadCount(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    // Calendar tools
    case 'list_calendars': {
      const params = ListCalendarsInput.parse(args ?? {});
      const result = calendarTools.listCalendars(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'list_events': {
      const params = ListEventsInput.parse(args ?? {});
      const result = calendarTools.listEvents(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'get_event': {
      const params = GetEventInput.parse(args);
      const result = calendarTools.getEvent(params);
      if (result == null) {
        return { content: [{ type: 'text', text: 'Event not found' }], isError: true };
      }
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'search_events': {
      const params = SearchEventsInput.parse(args);
      const result = calendarTools.searchEvents(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'create_event': {
      if (calendarWriter == null) {
        return {
          content: [{ type: 'text', text: 'Event creation is not available' }],
          isError: true,
        };
      }
      const params = CreateEventInput.parse(args);
      const writerParams: { title: string; startDate: string; endDate: string; calendarId?: number; location?: string; description?: string; isAllDay?: boolean; recurrence?: RecurrenceConfig } = {
        title: params.title,
        startDate: params.start_date,
        endDate: params.end_date,
      };
      if (params.calendar_id != null) writerParams.calendarId = params.calendar_id;
      if (params.location != null) writerParams.location = params.location;
      if (params.description != null) writerParams.description = params.description;
      if (params.is_all_day != null) writerParams.isAllDay = params.is_all_day;

      if (params.recurrence != null) {
        const rec = params.recurrence;
        const recConfig: RecurrenceConfig = {
          frequency: rec.frequency,
          interval: rec.interval,
        };
        const mut = recConfig as { -readonly [K in keyof RecurrenceConfig]: RecurrenceConfig[K] };
        if (rec.days_of_week != null) mut.daysOfWeek = rec.days_of_week;
        if (rec.day_of_month != null) mut.dayOfMonth = rec.day_of_month;
        if (rec.week_of_month != null) mut.weekOfMonth = rec.week_of_month;
        if (rec.day_of_week_monthly != null) mut.dayOfWeekMonthly = rec.day_of_week_monthly;
        if (rec.end.type === 'end_date') mut.endDate = rec.end.date;
        if (rec.end.type === 'end_after_count') mut.endAfterCount = rec.end.count;
        writerParams.recurrence = recConfig;
      }

      const created = calendarWriter.createEvent(writerParams);

      const result: CreateEventResult = {
        id: created.id,
        title: params.title,
        start_date: params.start_date,
        end_date: params.end_date,
        calendar_id: created.calendarId,
        location: params.location ?? null,
        description: params.description ?? null,
        is_all_day: params.is_all_day,
        is_recurring: params.recurrence != null,
      };

      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'respond_to_event': {
      if (calendarManager == null) {
        return {
          content: [{ type: 'text', text: 'Event response is not available' }],
          isError: true,
        };
      }
      const params = RespondToEventInput.parse(args);

      const result = calendarManager.respondToEvent(
        params.event_id,
        params.response,
        params.send_response,
        params.comment
      );

      const responseText = params.response === 'accept'
        ? 'accepted'
        : params.response === 'decline'
        ? 'declined'
        : 'tentatively accepted';

      return {
        content: [{
          type: 'text',
          text: `Successfully ${responseText} event ${result.eventId}`,
        }],
      };
    }

    case 'delete_event': {
      if (calendarManager == null) {
        return {
          content: [{ type: 'text', text: 'Event deletion is not available' }],
          isError: true,
        };
      }
      const params = args as { event_id: number; apply_to?: 'this_instance' | 'all_in_series' };
      const applyTo = params.apply_to ?? 'this_instance';

      calendarManager.deleteEvent(params.event_id, applyTo);

      const deleteText = applyTo === 'all_in_series' ? ' (entire series)' : '';
      return {
        content: [{
          type: 'text',
          text: `Successfully deleted event ${params.event_id}${deleteText}`,
        }],
      };
    }

    case 'update_event': {
      if (calendarManager == null) {
        return {
          content: [{ type: 'text', text: 'Event update is not available' }],
          isError: true,
        };
      }
      const params = args as {
        event_id: number;
        apply_to?: 'this_instance' | 'all_in_series';
        title?: string;
        start_date?: string;
        end_date?: string;
        location?: string;
        description?: string;
        is_all_day?: boolean;
      };

      // Validate date ordering if both dates are provided
      if (params.start_date != null && params.end_date != null) {
        if (new Date(params.start_date).getTime() >= new Date(params.end_date).getTime()) {
          return {
            content: [{ type: 'text', text: 'start_date must be before end_date' }],
            isError: true,
          };
        }
      }

      const applyTo = params.apply_to ?? 'this_instance';
      const updates: import('./applescript/index.js').EventUpdates = {
        ...(params.title != null && { title: params.title }),
        ...(params.start_date != null && { startDate: params.start_date }),
        ...(params.end_date != null && { endDate: params.end_date }),
        ...(params.location != null && { location: params.location }),
        ...(params.description != null && { description: params.description }),
        ...(params.is_all_day != null && { isAllDay: params.is_all_day }),
      };

      const result = calendarManager.updateEvent(params.event_id, updates, applyTo);

      const updateText = applyTo === 'all_in_series' ? ' (entire series)' : '';
      return {
        content: [{
          type: 'text',
          text: `Successfully updated event ${result.id}${updateText}. Updated fields: ${result.updatedFields.join(', ')}`,
        }],
      };
    }

    // Contact tools
    case 'list_contacts': {
      const params = ListContactsInput.parse(args ?? {});
      const result = contactsTools.listContacts(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'search_contacts': {
      const params = SearchContactsInput.parse(args);
      const result = contactsTools.searchContacts(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'get_contact': {
      const params = GetContactInput.parse(args);
      const result = contactsTools.getContact(params);
      if (result == null) {
        return { content: [{ type: 'text', text: 'Contact not found' }], isError: true };
      }
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    // Task tools
    case 'list_tasks': {
      const params = ListTasksInput.parse(args ?? {});
      const result = tasksTools.listTasks(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'search_tasks': {
      const params = SearchTasksInput.parse(args);
      const result = tasksTools.searchTasks(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'get_task': {
      const params = GetTaskInput.parse(args);
      const result = tasksTools.getTask(params);
      if (result == null) {
        return { content: [{ type: 'text', text: 'Task not found' }], isError: true };
      }
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    // Note tools
    case 'list_notes': {
      const params = ListNotesInput.parse(args ?? {});
      const result = notesTools.listNotes(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'get_note': {
      const params = GetNoteInput.parse(args);
      const result = notesTools.getNote(params);
      if (result == null) {
        return { content: [{ type: 'text', text: 'Note not found' }], isError: true };
      }
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'search_notes': {
      const params = SearchNotesInput.parse(args);
      const result = notesTools.searchNotes(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    // Email sending tool
    case 'send_email': {
      if (mailSender == null) {
        return {
          content: [{ type: 'text', text: 'Email sending is not available' }],
          isError: true,
        };
      }

      const params = args as {
        to: string[];
        subject: string;
        body: string;
        body_type?: 'plain' | 'html';
        cc?: string[];
        bcc?: string[];
        reply_to?: string;
        attachments?: Array<{ path: string; name?: string }>;
        account_id?: number;
      };

      let sendParams: import('./applescript/index.js').MailSenderSendEmailParams = {
        to: params.to,
        subject: params.subject,
        body: params.body,
        bodyType: params.body_type ?? 'plain',
      };

      if (params.cc != null) sendParams = { ...sendParams, cc: params.cc };
      if (params.bcc != null) sendParams = { ...sendParams, bcc: params.bcc };
      if (params.reply_to != null) sendParams = { ...sendParams, replyTo: params.reply_to };
      if (params.attachments != null) sendParams = { ...sendParams, attachments: params.attachments };
      if (params.account_id != null) sendParams = { ...sendParams, accountId: params.account_id };

      const sent = mailSender.sendEmail(sendParams);

      const result = {
        message_id: sent.messageId,
        sent_at: sent.sentAt,
        status: 'sent',
      };

      return {
        content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
      };
    }

    default:
      return {
        content: [{ type: 'text', text: `Unknown tool: ${name}` }],
        isError: true,
      };
  }
}

// =============================================================================
// Graph API Tool Handler
// =============================================================================

async function handleGraphToolCall(
  name: string,
  args: unknown,
  repository: GraphRepository,
  contentReaders: GraphContentReaders,
  orgTools: ReturnType<typeof createMailboxOrganizationTools>
): Promise<ToolResult> {
  // Handle mailbox organization tools (shared between backends)
  const orgResult = await handleOrgToolCall(name, args, orgTools);
  if (orgResult != null) return orgResult;

  try {
    switch (name) {
      // Mail tools
      case 'list_folders': {
        const folders = await repository.listFoldersAsync();
        const result = { folders: folders.map(transformFolderRow) };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'list_emails': {
        const params = ListEmailsInput.parse(args);
        const emails = params.unread_only
          ? await repository.listUnreadEmailsAsync(params.folder_id, params.limit, params.offset)
          : await repository.listEmailsAsync(params.folder_id, params.limit, params.offset);
        const result = { emails: emails.map(transformEmailRow) };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'search_emails': {
        const params = SearchEmailsInput.parse(args);
        const emails = params.folder_id != null
          ? await repository.searchEmailsInFolderAsync(params.folder_id, params.query, params.limit)
          : await repository.searchEmailsAsync(params.query, params.limit);
        const result = { emails: emails.map(transformEmailRow) };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'get_email': {
        const params = GetEmailInput.parse(args);
        const email = await repository.getEmailAsync(params.email_id);
        if (email == null) {
          return { content: [{ type: 'text', text: 'Email not found' }], isError: true };
        }

        let body: string | null = null;
        if (params.include_body) {
          body = await contentReaders.email.readEmailBodyAsync(email.dataFilePath);
          if (params.strip_html && body != null) {
            body = stripHtml(body);
          }
        }

        const result = { ...transformEmailRow(email), body };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'get_unread_count': {
        const params = GetUnreadCountInput.parse(args ?? {});
        const count = params.folder_id != null
          ? await repository.getUnreadCountByFolderAsync(params.folder_id)
          : await repository.getUnreadCountAsync();
        const result = { total: count };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      // Calendar tools
      case 'list_calendars': {
        const calendars = await repository.listCalendarsAsync();
        const result = { calendars: calendars.map(transformFolderRow) };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'list_events': {
        const params = ListEventsInput.parse(args ?? {});
        let events;
        if (params.start_date != null && params.end_date != null) {
          const startTs = Math.floor(new Date(params.start_date).getTime() / 1000);
          const endTs = Math.floor(new Date(params.end_date).getTime() / 1000);
          events = await repository.listEventsByDateRangeAsync(startTs, endTs, params.limit);
        } else if (params.calendar_id != null) {
          events = await repository.listEventsByFolderAsync(params.calendar_id, params.limit);
        } else {
          events = await repository.listEventsAsync(params.limit);
        }
        const result = { events: events.map(transformEventRow) };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'get_event': {
        const params = GetEventInput.parse(args);
        const event = await repository.getEventAsync(params.event_id);
        if (event == null) {
          return { content: [{ type: 'text', text: 'Event not found' }], isError: true };
        }

        const details = await contentReaders.event.readEventDetailsAsync(event.dataFilePath);
        const result = { ...transformEventRow(event), ...details };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'search_events': {
        const params = SearchEventsInput.parse(args);
        // Graph doesn't have direct event search, so we filter client-side
        const allEvents = await repository.listEventsAsync(1000);
        const queryLower = params.query.toLowerCase();
        const events = allEvents.filter((e) =>
          transformEventRow(e).title?.toLowerCase().includes(queryLower)
        );
        const result = { events: events.slice(0, params.limit).map(transformEventRow) };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'create_event': {
        return {
          content: [{ type: 'text', text: 'Event creation is not yet supported via Microsoft Graph API' }],
          isError: true,
        };
      }

      // Contact tools
      case 'list_contacts': {
        const params = ListContactsInput.parse(args ?? {});
        const contacts = await repository.listContactsAsync(params.limit, params.offset);
        const result = { contacts: contacts.map(transformContactRow) };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'search_contacts': {
        const params = SearchContactsInput.parse(args);
        const contacts = await repository.searchContactsAsync(params.query, params.limit);
        const result = { contacts: contacts.map(transformContactRow) };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'get_contact': {
        const params = GetContactInput.parse(args);
        const contact = await repository.getContactAsync(params.contact_id);
        if (contact == null) {
          return { content: [{ type: 'text', text: 'Contact not found' }], isError: true };
        }

        const details = await contentReaders.contact.readContactDetailsAsync(contact.dataFilePath);
        const result = { ...transformContactRow(contact), ...details };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      // Task tools
      case 'list_tasks': {
        const params = ListTasksInput.parse(args ?? {});
        const tasks = params.include_completed
          ? await repository.listTasksAsync(params.limit, params.offset)
          : await repository.listIncompleteTasksAsync(params.limit, params.offset);
        const result = { tasks: tasks.map(transformTaskRow) };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'search_tasks': {
        const params = SearchTasksInput.parse(args);
        const tasks = await repository.searchTasksAsync(params.query, params.limit);
        const result = { tasks: tasks.map(transformTaskRow) };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'get_task': {
        const params = GetTaskInput.parse(args);
        const task = await repository.getTaskAsync(params.task_id);
        if (task == null) {
          return { content: [{ type: 'text', text: 'Task not found' }], isError: true };
        }

        const details = await contentReaders.task.readTaskDetailsAsync(task.dataFilePath);
        const result = { ...transformTaskRow(task), ...details };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      // Note tools - NOT SUPPORTED in Graph API
      case 'list_notes': {
        return {
          content: [{ type: 'text', text: JSON.stringify({ notes: [], message: 'Notes are not supported by Microsoft Graph API' }, null, 2) }],
        };
      }

      case 'get_note': {
        return {
          content: [{ type: 'text', text: 'Notes are not supported by Microsoft Graph API' }],
          isError: true,
        };
      }

      case 'search_notes': {
        return {
          content: [{ type: 'text', text: JSON.stringify({ notes: [], message: 'Notes are not supported by Microsoft Graph API' }, null, 2) }],
        };
      }

      default:
        return {
          content: [{ type: 'text', text: `Unknown tool: ${name}` }],
          isError: true,
        };
    }
  } catch (error) {
    throw new GraphError(
      error instanceof Error ? error.message : 'Graph API error',
      error instanceof Error ? error : undefined
    );
  }
}

// =============================================================================
// Transform Helpers for Graph Mode
// =============================================================================

import type { FolderRow, EmailRow, EventRow, ContactRow, TaskRow } from './database/repository.js';
import { appleTimestampToIso } from './utils/dates.js';

function transformFolderRow(row: FolderRow): {
  id: number;
  name: string;
  parentId: number | null;
  specialType: number;
  folderType: number;
  accountId: number;
  messageCount: number;
  unreadCount: number;
} {
  return {
    id: row.id,
    name: row.name ?? 'Unnamed',
    parentId: row.parentId,
    specialType: row.specialType,
    folderType: row.folderType,
    accountId: row.accountId,
    messageCount: row.messageCount,
    unreadCount: row.unreadCount,
  };
}

function transformEmailRow(row: EmailRow): {
  id: number;
  folderId: number | null;
  subject: string | null;
  sender: string | null;
  senderAddress: string | null;
  preview: string | null;
  isRead: boolean;
  timeReceived: string | null;
  timeSent: string | null;
  hasAttachment: boolean;
  priority: number | null;
  flagStatus: number | null;
  categories: string[];
} {
  return {
    id: row.id,
    folderId: row.folderId,
    subject: row.subject,
    sender: row.sender,
    senderAddress: row.senderAddress,
    preview: row.preview,
    isRead: row.isRead === 1,
    timeReceived: row.timeReceived != null ? appleTimestampToIso(row.timeReceived) : null,
    timeSent: row.timeSent != null ? appleTimestampToIso(row.timeSent) : null,
    hasAttachment: row.hasAttachment === 1,
    priority: row.priority,
    flagStatus: row.flagStatus,
    categories: parseEmailCategories(row.categories),
  };
}

function parseEmailCategories(buffer: Buffer | null): string[] {
  if (buffer == null || buffer.length === 0) return [];
  try {
    const text = buffer.toString('utf-8');
    return text.includes('\0')
      ? text.split('\0').filter(s => s.length > 0)
      : text.split(',').map(s => s.trim()).filter(s => s.length > 0);
  } catch {
    return [];
  }
}

function transformEventRow(row: EventRow) {
  return {
    id: row.id,
    folderId: row.folderId,
    title: null as string | null, // Will be filled from content reader
    startDate: row.startDate != null ? appleTimestampToIso(row.startDate) : null,
    endDate: row.endDate != null ? appleTimestampToIso(row.endDate) : null,
    isRecurring: row.isRecurring === 1,
    hasReminder: row.hasReminder === 1,
    attendeeCount: row.attendeeCount,
  };
}

function transformContactRow(row: ContactRow): {
  id: number;
  displayName: string | null;
  sortName: string | null;
} {
  return {
    id: row.id,
    displayName: row.displayName,
    sortName: row.sortName,
  };
}

function transformTaskRow(row: TaskRow): {
  id: number;
  folderId: number | null;
  name: string | null;
  isCompleted: boolean;
  dueDate: string | null;
  startDate: string | null;
  priority: number | null;
  hasReminder: boolean;
} {
  return {
    id: row.id,
    folderId: row.folderId,
    name: row.name,
    isCompleted: row.isCompleted === 1,
    dueDate: row.dueDate != null ? appleTimestampToIso(row.dueDate) : null,
    startDate: row.startDate != null ? appleTimestampToIso(row.startDate) : null,
    priority: row.priority,
    hasReminder: row.hasReminder === 1,
  };
}

function stripHtml(html: string): string {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

// =============================================================================
// Main Entry Point
// =============================================================================

async function main(): Promise<void> {
  const server = createServer();
  const transport = new StdioServerTransport();

  await server.connect(transport);
}

// Run if this is the main module (not imported for testing)
const isMainModule =
  import.meta.url === `file://${process.argv[1]}` ||
  (process.argv[1]?.endsWith('dist/index.js') === true);

if (isMainModule === true) {
  main().catch((error) => {
    console.error('Fatal error:', error);
    process.exit(1);
  });
}
