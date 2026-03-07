#!/usr/bin/env node
/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */
/**
 * Office 365 MCP Server
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
import { z } from 'zod';

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
  getAccessToken,
  GraphMailboxAdapter,
  type GraphRepository,
  type GraphContentReaders,
} from './graph/index.js';
import { parseCliCommand, handleAuthCommand, createAuthMutex } from './cli.js';
import { createMailTools } from './tools/mail.js';
import { createCalendarTools } from './tools/calendar.js';
import { createContactsTools } from './tools/contacts.js';
import { createTasksTools } from './tools/tasks.js';
import { createNotesTools } from './tools/notes.js';
import { createMailboxOrganizationTools } from './tools/mailbox-organization.js';
import {
  createMailSendTools,
  SetSignatureInput,
  GetSignatureInput,
} from './tools/mail-send.js';
import {
  createSchedulingTools,
  CheckAvailabilityInput,
  FindMeetingTimesInput,
} from './tools/scheduling.js';
import {
  MailRulesTools,
  CreateMailRuleInput,
  PrepareDeleteMailRuleInput,
  ConfirmDeleteMailRuleInput,
} from './tools/mail-rules.js';
import {
  CategoriesTools,
  CreateCategoryInput,
  PrepareDeleteCategoryInput,
  ConfirmDeleteCategoryInput,
} from './tools/categories.js';
import {
  CalendarPermissionsTools,
  ListCalendarPermissionsInput,
  CreateCalendarPermissionInput,
  PrepareDeleteCalendarPermissionInput,
  ConfirmDeleteCalendarPermissionInput,
} from './tools/calendar-permissions.js';
import {
  FocusedOverridesTools,
  CreateFocusedOverrideInput,
  PrepareDeleteFocusedOverrideInput,
  ConfirmDeleteFocusedOverrideInput,
} from './tools/focused-overrides.js';
import {
  ChecklistItemsTools,
  ListChecklistItemsInput,
  CreateChecklistItemInput,
  UpdateChecklistItemInput,
  PrepareDeleteChecklistItemInput,
  ConfirmDeleteChecklistItemInput,
} from './tools/checklist-items.js';
import {
  LinkedResourcesTools,
  ListLinkedResourcesInput,
  CreateLinkedResourceInput,
  PrepareDeleteLinkedResourceInput,
  ConfirmDeleteLinkedResourceInput,
} from './tools/linked-resources.js';
import {
  TaskAttachmentsTools,
  ListTaskAttachmentsInput,
  CreateTaskAttachmentInput,
  PrepareDeleteTaskAttachmentInput,
  ConfirmDeleteTaskAttachmentInput,
} from './tools/task-attachments.js';
import {
  TeamsTools,
  ListChannelsInput,
  GetChannelInput,
  CreateChannelInput,
  UpdateChannelInput,
  PrepareDeleteChannelInput,
  ConfirmDeleteChannelInput,
  ListTeamMembersInput,
  ListChannelMessagesInput,
  GetChannelMessageInput,
  PrepareSendChannelMessageInput,
  ConfirmSendChannelMessageInput,
  PrepareReplyChannelMessageInput,
  ConfirmReplyChannelMessageInput,
  ListChatsInput,
  GetChatInput,
  ListChatMessagesInput,
  PrepareSendChatMessageInput,
  ConfirmSendChatMessageInput,
  ListChatMembersInput,
} from './tools/teams.js';
import {
  ListEmailsInput,
  SearchEmailsInput,
  SearchEmailsAdvancedInput,
  GetEmailInput,
  GetEmailsInput,
  ListConversationInput,
  GetUnreadCountInput,
  ListAttachmentsInput,
  DownloadAttachmentInput,
  CheckNewEmailsInput,
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
  SetEmailImportanceInput,
  CreateFolderInput,
  RenameFolderInput,
  MoveFolderInput,
  CreateDraftInput,
  UpdateDraftInput,
  ListDraftsInput,
  PrepareSendDraftInput,
  ConfirmSendDraftInput,
  PrepareSendEmailInput,
  ConfirmSendEmailInput,
  PrepareReplyEmailInput,
  ConfirmReplyEmailInput,
  PrepareForwardEmailInput,
  ConfirmForwardEmailInput,
  ReplyAsDraftInput,
  ForwardAsDraftInput,
  AddDraftAttachmentInput,
  AddDraftInlineImageInput,
} from './tools/index.js';
import { ApprovalTokenManager, hashEventForApproval, hashContactForApproval, hashTaskForApproval } from './approval/index.js';
import type { CreateEventResult } from './tools/index.js';
import {
  wrapError,
  OutlookNotRunningError,
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
    name: 'search_emails_advanced',
    description: 'Search emails using KQL (Keyword Query Language) for advanced queries. Supports operators: from:, to:, subject:, hasAttachments:true, received>=2024-01-01, AND, OR. (Graph API)',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'KQL search query (e.g., from:alice AND subject:"report")' },
        folder_id: { type: 'number', description: 'Optional folder ID to search within' },
        limit: { type: 'number', description: 'Maximum results (default: 50)', default: 50 },
      },
      required: ['query'],
    },
  },
  {
    name: 'check_new_emails',
    description: 'Check for new or changed emails since last check using delta sync. First call returns recent messages (initial sync). Subsequent calls return only new/changed messages.',
    inputSchema: {
      type: 'object',
      properties: {
        folder_id: { type: 'number', description: 'Folder ID to check for new emails' },
      },
      required: ['folder_id'],
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
    name: 'get_emails',
    description: 'Get multiple emails by ID in a single call (max 25). Useful for batch operations or summarizing threads.',
    inputSchema: {
      type: 'object',
      properties: {
        email_ids: {
          type: 'array',
          items: { type: 'number' },
          description: 'Array of email IDs to fetch (max 25)',
        },
        include_body: {
          type: 'boolean',
          description: 'Include full email body (default: false)',
          default: false,
        },
        strip_html: {
          type: 'boolean',
          description: 'Strip HTML tags from body (default: false)',
          default: false,
        },
      },
      required: ['email_ids'],
    },
  },
  {
    name: 'list_conversation',
    description: 'List all messages in an email conversation/thread, ordered chronologically. Provide any message ID from the thread.',
    inputSchema: {
      type: 'object',
      properties: {
        message_id: { type: 'number', description: 'Any message ID from the conversation' },
        limit: { type: 'number', description: 'Maximum messages to return (default: 25)', default: 25 },
      },
      required: ['message_id'],
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
  // Attachment tools
  {
    name: 'list_attachments',
    description: 'List attachment metadata (name, size, type) for an email',
    inputSchema: {
      type: 'object',
      properties: {
        email_id: {
          type: 'number',
          description: 'The email ID to list attachments for',
        },
      },
      required: ['email_id'],
    },
  },
  {
    name: 'download_attachment',
    description: 'Download/save an email attachment to a file on disk. Returns the saved file path and size.',
    inputSchema: {
      type: 'object',
      properties: {
        email_id: {
          type: 'number',
          description: 'The email ID containing the attachment',
        },
        attachment_index: {
          type: 'number',
          description: 'The 1-based index of the attachment (from list_attachments)',
        },
        save_path: {
          type: 'string',
          description: 'Absolute file path where the attachment should be saved',
        },
      },
      required: ['email_id', 'attachment_index', 'save_path'],
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
    description: 'Search events by title and/or date range across all calendars',
    inputSchema: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'Search query for event titles',
        },
        start_date: {
          type: 'string',
          description: 'Start date filter in ISO 8601 format (events starting on or after this date)',
        },
        end_date: {
          type: 'string',
          description: 'End date filter in ISO 8601 format (events ending on or before this date)',
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
    name: 'create_event',
    description: 'Create a new calendar event in Outlook. Supports online Teams meetings via is_online_meeting flag.',
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
        is_online_meeting: {
          type: 'boolean',
          description: 'Create as online Teams meeting (default false)',
          default: false,
        },
        online_meeting_provider: {
          type: 'string',
          enum: ['teamsForBusiness', 'skypeForBusiness', 'skypeForConsumer'],
          description: 'Online meeting provider (default: teamsForBusiness)',
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
    description: 'Update a calendar event. All fields are optional - only specified fields will be updated. Supports online Teams meetings via is_online_meeting flag. For recurring events, you can update a single instance or the entire series.',
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
        is_online_meeting: {
          type: 'boolean',
          description: 'Set as online Teams meeting',
        },
        online_meeting_provider: {
          type: 'string',
          enum: ['teamsForBusiness', 'skypeForBusiness', 'skypeForConsumer'],
          description: 'Online meeting provider (default: teamsForBusiness)',
        },
      },
      required: ['event_id'],
    },
  },
  {
    name: 'prepare_delete_event',
    description: 'Prepare to delete a calendar event. Returns a preview and approval token. Call confirm_delete_event to execute.',
    inputSchema: {
      type: 'object',
      properties: {
        event_id: { type: 'number', description: 'The event ID to delete' },
      },
      required: ['event_id'],
    },
  },
  {
    name: 'confirm_delete_event',
    description: 'Confirm deletion of a calendar event using a token from prepare_delete_event',
    inputSchema: {
      type: 'object',
      properties: {
        token_id: { type: 'string', description: 'The approval token from prepare_delete_event' },
        event_id: { type: 'number', description: 'The event ID to delete' },
      },
      required: ['token_id', 'event_id'],
    },
  },
  {
    name: 'list_event_instances',
    description: 'List instances of a recurring event within a date range. Instance IDs can be used with update_event and delete_event. (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        event_id: { type: 'number', description: 'Recurring event ID' },
        start_date: { type: 'string', description: 'Start of date range (ISO 8601, e.g. 2024-01-01T00:00:00Z)' },
        end_date: { type: 'string', description: 'End of date range (ISO 8601, e.g. 2024-12-31T23:59:59Z)' },
      },
      required: ['event_id', 'start_date', 'end_date'],
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
        folder_id: {
          type: 'number',
          description: 'Filter contacts by contact folder ID (optional)',
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
  {
    name: 'create_contact',
    description: 'Create a new contact in Outlook. All fields are optional but at least one should be provided.',
    inputSchema: {
      type: 'object',
      properties: {
        given_name: { type: 'string', description: 'First name' },
        surname: { type: 'string', description: 'Last name' },
        email: { type: 'string', description: 'Email address' },
        phone: { type: 'string', description: 'Business phone number' },
        mobile_phone: { type: 'string', description: 'Mobile phone number' },
        company: { type: 'string', description: 'Company name' },
        job_title: { type: 'string', description: 'Job title' },
        street_address: { type: 'string', description: 'Street address' },
        city: { type: 'string', description: 'City' },
        state: { type: 'string', description: 'State or province' },
        postal_code: { type: 'string', description: 'Postal/ZIP code' },
        country: { type: 'string', description: 'Country or region' },
      },
      required: [],
    },
  },
  {
    name: 'update_contact',
    description: 'Update an existing contact. Only specified fields will be updated.',
    inputSchema: {
      type: 'object',
      properties: {
        contact_id: { type: 'number', description: 'The contact ID to update' },
        given_name: { type: 'string', description: 'First name' },
        surname: { type: 'string', description: 'Last name' },
        email: { type: 'string', description: 'Email address' },
        phone: { type: 'string', description: 'Business phone number' },
        mobile_phone: { type: 'string', description: 'Mobile phone number' },
        company: { type: 'string', description: 'Company name' },
        job_title: { type: 'string', description: 'Job title' },
        street_address: { type: 'string', description: 'Street address' },
        city: { type: 'string', description: 'City' },
        state: { type: 'string', description: 'State or province' },
        postal_code: { type: 'string', description: 'Postal/ZIP code' },
        country: { type: 'string', description: 'Country or region' },
      },
      required: ['contact_id'],
    },
  },
  {
    name: 'prepare_delete_contact',
    description: 'Prepare to delete a contact. Returns a preview and approval token. Call confirm_delete_contact to execute.',
    inputSchema: {
      type: 'object',
      properties: {
        contact_id: { type: 'number', description: 'The contact ID to delete' },
      },
      required: ['contact_id'],
    },
  },
  {
    name: 'confirm_delete_contact',
    description: 'Confirm deletion of a contact using a token from prepare_delete_contact',
    inputSchema: {
      type: 'object',
      properties: {
        token_id: { type: 'string', description: 'The approval token from prepare_delete_contact' },
        contact_id: { type: 'number', description: 'The contact ID to delete' },
      },
      required: ['token_id', 'contact_id'],
    },
  },
  // Task tools
  {
    name: 'list_task_lists',
    description: 'List all task lists (Microsoft To Do) (Graph API)',
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
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
  {
    name: 'create_task',
    description: 'Create a new task in a task list. Supports optional recurrence settings for repeating tasks.',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Task title' },
        task_list_id: { type: 'number', description: 'The task list ID to create the task in' },
        body: { type: 'string', description: 'Task body/notes' },
        body_type: { type: 'string', enum: ['text', 'html'], default: 'text', description: 'Body content type' },
        due_date: { type: 'string', description: 'Due date (ISO 8601 format)' },
        importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Task importance' },
        reminder_date: { type: 'string', description: 'Reminder date (ISO 8601 format)' },
        recurrence: {
          type: 'object',
          description: 'Task recurrence settings',
          properties: {
            pattern: { type: 'string', enum: ['daily', 'weekly', 'monthly', 'yearly'], description: 'Recurrence pattern type' },
            interval: { type: 'number', default: 1, description: 'Interval between occurrences' },
            days_of_week: { type: 'array', items: { type: 'string', enum: ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'] }, description: 'Days of week (for weekly pattern)' },
            day_of_month: { type: 'number', description: 'Day of month (for monthly pattern)' },
            range_type: { type: 'string', enum: ['endDate', 'noEnd', 'numbered'], description: 'How the recurrence ends' },
            start_date: { type: 'string', description: 'Start date (YYYY-MM-DD)' },
            end_date: { type: 'string', description: 'End date (YYYY-MM-DD, for endDate range)' },
            occurrences: { type: 'number', description: 'Number of occurrences (for numbered range)' },
          },
          required: ['pattern', 'range_type', 'start_date'],
        },
        categories: { type: 'array', items: { type: 'string' }, description: 'Category names to assign to the task' },
      },
      required: ['title', 'task_list_id'],
    },
  },
  {
    name: 'update_task',
    description: 'Update an existing task. Only specified fields will be updated. Supports optional recurrence settings for repeating tasks.',
    inputSchema: {
      type: 'object',
      properties: {
        task_id: { type: 'number', description: 'The task ID to update' },
        title: { type: 'string', description: 'New task title' },
        body: { type: 'string', description: 'New task body/notes' },
        body_type: { type: 'string', enum: ['text', 'html'], description: 'Body content type' },
        due_date: { type: 'string', description: 'New due date (ISO 8601 format)' },
        importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Task importance' },
        reminder_date: { type: 'string', description: 'Reminder date (ISO 8601 format)' },
        status: { type: 'string', enum: ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'], description: 'Task status' },
        recurrence: {
          type: 'object',
          description: 'Task recurrence settings',
          properties: {
            pattern: { type: 'string', enum: ['daily', 'weekly', 'monthly', 'yearly'], description: 'Recurrence pattern type' },
            interval: { type: 'number', default: 1, description: 'Interval between occurrences' },
            days_of_week: { type: 'array', items: { type: 'string', enum: ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'] }, description: 'Days of week (for weekly pattern)' },
            day_of_month: { type: 'number', description: 'Day of month (for monthly pattern)' },
            range_type: { type: 'string', enum: ['endDate', 'noEnd', 'numbered'], description: 'How the recurrence ends' },
            start_date: { type: 'string', description: 'Start date (YYYY-MM-DD)' },
            end_date: { type: 'string', description: 'End date (YYYY-MM-DD, for endDate range)' },
            occurrences: { type: 'number', description: 'Number of occurrences (for numbered range)' },
          },
          required: ['pattern', 'range_type', 'start_date'],
        },
        categories: { type: 'array', items: { type: 'string' }, description: 'Category names to assign to the task' },
      },
      required: ['task_id'],
    },
  },
  {
    name: 'complete_task',
    description: 'Mark a task as completed',
    inputSchema: {
      type: 'object',
      properties: {
        task_id: { type: 'number', description: 'The task ID to complete' },
      },
      required: ['task_id'],
    },
  },
  {
    name: 'create_task_list',
    description: 'Create a new task list',
    inputSchema: {
      type: 'object',
      properties: {
        display_name: { type: 'string', description: 'Name for the new task list' },
      },
      required: ['display_name'],
    },
  },
  {
    name: 'prepare_delete_task',
    description: 'Prepare to delete a task. Returns a preview and approval token. Call confirm_delete_task to execute.',
    inputSchema: {
      type: 'object',
      properties: {
        task_id: { type: 'number', description: 'The task ID to delete' },
      },
      required: ['task_id'],
    },
  },
  {
    name: 'confirm_delete_task',
    description: 'Confirm deletion of a task using a token from prepare_delete_task',
    inputSchema: {
      type: 'object',
      properties: {
        token_id: { type: 'string', description: 'The approval token from prepare_delete_task' },
        task_id: { type: 'number', description: 'The task ID to delete' },
      },
      required: ['token_id', 'task_id'],
    },
  },
  {
    name: 'rename_task_list',
    description: 'Rename a task list (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        task_list_id: { type: 'number', description: 'Task list ID' },
        name: { type: 'string', description: 'New name for the task list' },
      },
      required: ['task_list_id', 'name'],
    },
  },
  {
    name: 'prepare_delete_task_list',
    description: 'Prepare to delete a task list. Returns an approval token. Call confirm_delete_task_list to execute. (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        task_list_id: { type: 'number', description: 'Task list ID to delete' },
      },
      required: ['task_list_id'],
    },
  },
  {
    name: 'confirm_delete_task_list',
    description: 'Confirm task list deletion with approval token (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        token_id: { type: 'string', description: 'The approval token from prepare_delete_task_list' },
        task_list_id: { type: 'number', description: 'The task list ID to delete' },
      },
      required: ['token_id', 'task_list_id'],
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
  {
    name: 'set_email_importance',
    description: 'Set email importance/priority level (Graph API)',
    inputSchema: {
      type: 'object',
      properties: {
        email_id: { type: 'number', description: 'The email ID' },
        importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Importance level' },
      },
      required: ['email_id', 'importance'],
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
        inline_images: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              path: {
                type: 'string',
                description: 'Absolute file path to the image',
              },
              content_id: {
                type: 'string',
                description: 'Content ID for referencing in HTML body (use in <img src="cid:content_id">)',
              },
            },
            required: ['path', 'content_id'],
          },
          description: 'Inline images to embed in HTML body (reference via cid: in img tags)',
        },
        account_id: {
          type: 'number',
          description: 'Account to send from (optional)',
        },
      },
      required: ['to', 'subject', 'body'],
    },
  },
  // =========================================================================
  // Mail Send — Draft Management (Non-Destructive, Graph API only)
  // =========================================================================
  {
    name: 'create_draft',
    description: 'Create a draft email that can be edited and sent later',
    inputSchema: {
      type: 'object',
      properties: {
        to: {
          type: 'array',
          items: { type: 'string' },
          description: 'To recipients (email addresses)',
        },
        cc: {
          type: 'array',
          items: { type: 'string' },
          description: 'CC recipients (email addresses)',
        },
        bcc: {
          type: 'array',
          items: { type: 'string' },
          description: 'BCC recipients (email addresses)',
        },
        subject: {
          type: 'string',
          description: 'Email subject',
        },
        body: {
          type: 'string',
          description: 'Email body',
        },
        body_type: {
          type: 'string',
          enum: ['text', 'html'],
          default: 'text',
          description: 'Body content type (default: text)',
        },
        include_signature: {
          type: 'boolean',
          default: true,
          description: 'Include email signature (default: true)',
        },
        attachments: {
          type: 'array',
          description: 'File attachments',
          items: {
            type: 'object',
            properties: {
              file_path: { type: 'string', description: 'Absolute path to the file' },
              name: { type: 'string', description: 'Override filename' },
              content_type: { type: 'string', description: 'Override MIME type' },
            },
            required: ['file_path'],
          },
        },
        body_file: {
          type: 'string',
          description: 'Path to a file containing the email body (alternative to body; use to avoid large MCP payloads)',
        },
        inline_images: {
          type: 'array',
          description: 'Inline images for HTML body: reference in body via <img src="cid:content_id"> to avoid embedding base64 in the payload',
          items: {
            type: 'object',
            properties: {
              file_path: { type: 'string', description: 'Absolute path to the image file' },
              content_id: { type: 'string', description: 'Content-ID for HTML (e.g. "logo" for cid:logo)' },
            },
            required: ['file_path', 'content_id'],
          },
        },
      },
      required: ['subject'],
    },
  },
  {
    name: 'update_draft',
    description: 'Update an existing draft email',
    inputSchema: {
      type: 'object',
      properties: {
        draft_id: {
          type: 'number',
          description: 'The draft ID to update',
        },
        to: {
          type: 'array',
          items: { type: 'string' },
          description: 'To recipients (email addresses)',
        },
        cc: {
          type: 'array',
          items: { type: 'string' },
          description: 'CC recipients (email addresses)',
        },
        bcc: {
          type: 'array',
          items: { type: 'string' },
          description: 'BCC recipients (email addresses)',
        },
        subject: {
          type: 'string',
          description: 'Email subject',
        },
        body: {
          type: 'string',
          description: 'Email body',
        },
        body_type: {
          type: 'string',
          enum: ['text', 'html'],
          description: 'Body content type',
        },
      },
      required: ['draft_id'],
    },
  },
  {
    name: 'add_draft_attachment',
    description: 'Add a file attachment to an existing draft (Graph API)',
    inputSchema: {
      type: 'object',
      properties: {
        draft_id: { type: 'number', description: 'The draft ID' },
        file_path: { type: 'string', description: 'Absolute path to the file' },
        name: { type: 'string', description: 'Override filename (optional)' },
        content_type: { type: 'string', description: 'Override MIME type (optional)' },
      },
      required: ['draft_id', 'file_path'],
    },
  },
  {
    name: 'add_draft_inline_image',
    description: 'Add an inline image to an existing draft for use in HTML body (Graph API)',
    inputSchema: {
      type: 'object',
      properties: {
        draft_id: { type: 'number', description: 'The draft ID' },
        file_path: { type: 'string', description: 'Absolute path to the image file' },
        content_id: { type: 'string', description: 'Content-ID (reference in HTML as <img src="cid:content_id">)' },
      },
      required: ['draft_id', 'file_path', 'content_id'],
    },
  },
  {
    name: 'list_drafts',
    description: 'List all draft emails',
    inputSchema: {
      type: 'object',
      properties: {
        limit: {
          type: 'number',
          description: 'Maximum number of drafts to return (1-100, default 50)',
          default: 50,
        },
        offset: {
          type: 'number',
          description: 'Number of drafts to skip (default 0)',
          default: 0,
        },
      },
      required: [],
    },
  },
  // =========================================================================
  // Mail Send — Two-Phase Approval (Graph API only)
  // =========================================================================
  {
    name: 'prepare_send_draft',
    description: 'Prepare to send a draft email. Returns a preview and approval token.',
    inputSchema: {
      type: 'object',
      properties: {
        draft_id: { type: 'number', description: 'The draft ID to send' },
      },
      required: ['draft_id'],
    },
  },
  {
    name: 'confirm_send_draft',
    description: 'Confirm and send a draft email using the approval token.',
    inputSchema: {
      type: 'object',
      properties: {
        token_id: { type: 'string', description: 'Approval token from prepare_send_draft' },
        draft_id: { type: 'number', description: 'The draft ID to send' },
      },
      required: ['token_id', 'draft_id'],
    },
  },
  {
    name: 'prepare_send_email',
    description: 'Prepare to send an email immediately. Returns a preview and approval token.',
    inputSchema: {
      type: 'object',
      properties: {
        to: {
          type: 'array',
          items: { type: 'string' },
          minItems: 1,
          description: 'To recipients (email addresses)',
        },
        cc: {
          type: 'array',
          items: { type: 'string' },
          description: 'CC recipients (email addresses)',
        },
        bcc: {
          type: 'array',
          items: { type: 'string' },
          description: 'BCC recipients (email addresses)',
        },
        subject: {
          type: 'string',
          description: 'Email subject',
        },
        body: {
          type: 'string',
          description: 'Email body (omit when using body_file)',
        },
        body_file: {
          type: 'string',
          description: 'Path to a file containing the email body (alternative to body; use to avoid large MCP payloads)',
        },
        body_type: {
          type: 'string',
          enum: ['text', 'html'],
          default: 'text',
          description: 'Body content type (default: text)',
        },
        include_signature: {
          type: 'boolean',
          default: true,
          description: 'Include email signature (default: true)',
        },
        attachments: {
          type: 'array',
          description: 'File attachments',
          items: {
            type: 'object',
            properties: {
              file_path: { type: 'string', description: 'Absolute path to the file' },
              name: { type: 'string', description: 'Override filename' },
              content_type: { type: 'string', description: 'Override MIME type' },
            },
            required: ['file_path'],
          },
        },
      },
      required: ['to', 'subject'],
    },
  },
  {
    name: 'confirm_send_email',
    description: 'Confirm and send an email using the approval token.',
    inputSchema: {
      type: 'object',
      properties: {
        token_id: { type: 'string', description: 'Approval token from prepare_send_email' },
      },
      required: ['token_id'],
    },
  },
  {
    name: 'prepare_reply_email',
    description: 'Prepare to reply to an email. Returns a preview and approval token.',
    inputSchema: {
      type: 'object',
      properties: {
        message_id: { type: 'number', description: 'The message ID to reply to' },
        comment: { type: 'string', description: 'Reply body' },
        reply_all: {
          type: 'boolean',
          default: true,
          description: 'Reply to all recipients (default true)',
        },
      },
      required: ['message_id', 'comment'],
    },
  },
  {
    name: 'confirm_reply_email',
    description: 'Confirm and reply to an email using the approval token.',
    inputSchema: {
      type: 'object',
      properties: {
        token_id: { type: 'string', description: 'Approval token from prepare_reply_email' },
        message_id: { type: 'number', description: 'The message ID being replied to' },
      },
      required: ['token_id', 'message_id'],
    },
  },
  {
    name: 'prepare_forward_email',
    description: 'Prepare to forward an email. Returns a preview and approval token.',
    inputSchema: {
      type: 'object',
      properties: {
        message_id: { type: 'number', description: 'The message ID to forward' },
        to_recipients: {
          type: 'array',
          items: { type: 'string' },
          minItems: 1,
          description: 'Forward to recipients (email addresses)',
        },
        comment: { type: 'string', description: 'Optional comment to include' },
      },
      required: ['message_id', 'to_recipients'],
    },
  },
  {
    name: 'confirm_forward_email',
    description: 'Confirm and forward an email using the approval token.',
    inputSchema: {
      type: 'object',
      properties: {
        token_id: { type: 'string', description: 'Approval token from prepare_forward_email' },
        message_id: { type: 'number', description: 'The message ID being forwarded' },
      },
      required: ['token_id', 'message_id'],
    },
  },
  {
    name: 'reply_as_draft',
    description: 'Create a reply (or reply-all) as an editable draft. Returns a draft_id for use with update_draft and prepare_send_draft.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        message_id: { type: 'number', description: 'The message ID to reply to' },
        comment: { type: 'string', description: 'Initial reply body text' },
        reply_all: { type: 'boolean', default: false, description: 'Reply to all recipients (default: false)' },
        include_signature: {
          type: 'boolean',
          default: true,
          description: 'Include email signature (default: true)',
        },
      },
      required: ['message_id'],
    },
  },
  {
    name: 'forward_as_draft',
    description: 'Create a forward as an editable draft. Returns a draft_id for use with update_draft and prepare_send_draft.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        message_id: { type: 'number', description: 'The message ID to forward' },
        to_recipients: {
          type: 'array',
          items: { type: 'string' },
          description: 'Forward recipients (can also add later via update_draft)',
        },
        comment: { type: 'string', description: 'Initial forward body text' },
        include_signature: {
          type: 'boolean',
          default: true,
          description: 'Include email signature (default: true)',
        },
      },
      required: ['message_id'],
    },
  },
  // Signature management tools
  {
    name: 'set_signature',
    description: 'Save an email signature that will be auto-appended to outgoing emails',
    inputSchema: {
      type: 'object' as const,
      properties: {
        content: { type: 'string', description: 'Signature content (HTML or plain text)' },
        content_type: {
          type: 'string',
          enum: ['html', 'text'],
          default: 'html',
          description: 'Content type of the signature (default: html)',
        },
      },
      required: ['content'],
    },
  },
  {
    name: 'get_signature',
    description: 'Get the currently stored email signature',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
  // Calendar scheduling tools
  {
    name: 'check_availability',
    description: 'Check free/busy availability for one or more people in a time window',
    inputSchema: {
      type: 'object' as const,
      properties: {
        email_addresses: {
          type: 'array',
          items: { type: 'string' },
          minItems: 1,
          description: 'Email addresses to check availability for',
        },
        start_time: { type: 'string', description: 'Start of time window (ISO 8601)' },
        end_time: { type: 'string', description: 'End of time window (ISO 8601)' },
        availability_view_interval: {
          type: 'number',
          default: 30,
          description: 'Time slot interval in minutes (default: 30)',
        },
      },
      required: ['email_addresses', 'start_time', 'end_time'],
    },
  },
  {
    name: 'find_meeting_times',
    description: 'Find available meeting time slots for a group of attendees',
    inputSchema: {
      type: 'object' as const,
      properties: {
        attendees: {
          type: 'array',
          items: { type: 'string' },
          minItems: 1,
          description: 'Attendee email addresses',
        },
        duration_minutes: { type: 'number', description: 'Meeting duration in minutes' },
        start_time: { type: 'string', description: 'Start of search window (ISO 8601, optional)' },
        end_time: { type: 'string', description: 'End of search window (ISO 8601, optional)' },
        max_candidates: {
          type: 'number',
          default: 5,
          description: 'Maximum number of time suggestions (default: 5)',
        },
      },
      required: ['attendees', 'duration_minutes'],
    },
  },
  // Mail rules tools
  {
    name: 'list_mail_rules',
    description: 'List all inbox mail rules (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
  {
    name: 'create_mail_rule',
    description: 'Create an inbox mail rule with conditions and actions (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        display_name: { type: 'string', description: 'Rule name' },
        sequence: { type: 'number', description: 'Rule priority order' },
        is_enabled: { type: 'boolean', default: true, description: 'Whether rule is active' },
        conditions: {
          type: 'object',
          description: 'Conditions that trigger the rule',
          properties: {
            from_addresses: { type: 'array', items: { type: 'string' }, description: 'Match sender email addresses' },
            subject_contains: { type: 'array', items: { type: 'string' }, description: 'Subject contains any of these strings' },
            body_contains: { type: 'array', items: { type: 'string' }, description: 'Body contains any of these strings' },
            sender_contains: { type: 'array', items: { type: 'string' }, description: 'Sender field contains these strings' },
            has_attachments: { type: 'boolean', description: 'Has attachments' },
            importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Match importance level' },
          },
        },
        actions: {
          type: 'object',
          description: 'Actions to perform',
          properties: {
            move_to_folder: { type: 'number', description: 'Folder ID to move to' },
            mark_as_read: { type: 'boolean', description: 'Mark as read' },
            mark_importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Set importance' },
            forward_to: { type: 'array', items: { type: 'string' }, description: 'Forward to these email addresses' },
            delete: { type: 'boolean', description: 'Delete the message' },
            stop_processing_rules: { type: 'boolean', description: 'Stop processing more rules' },
          },
        },
      },
      required: ['display_name', 'conditions', 'actions'],
    },
  },
  {
    name: 'prepare_delete_mail_rule',
    description: 'Prepare to delete a mail rule. Returns a preview and approval token. Call confirm_delete_mail_rule to execute. (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        rule_id: { type: 'number', description: 'The rule ID to delete' },
      },
      required: ['rule_id'],
    },
  },
  {
    name: 'confirm_delete_mail_rule',
    description: 'Confirm mail rule deletion with approval token (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        token_id: { type: 'string', description: 'The approval token from prepare_delete_mail_rule' },
        rule_id: { type: 'number', description: 'The rule ID to delete' },
      },
      required: ['token_id', 'rule_id'],
    },
  },
  // Master categories tools
  {
    name: 'list_categories',
    description: 'List all master categories (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
  {
    name: 'create_category',
    description: 'Create a new master category (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        name: { type: 'string', description: 'Category name' },
        color: { type: 'string', enum: ['preset0','preset1','preset2','preset3','preset4','preset5','preset6','preset7','preset8','preset9','preset10','preset11','preset12','preset13','preset14','preset15','preset16','preset17','preset18','preset19','preset20','preset21','preset22','preset23','preset24','none'], description: 'Category color preset' },
      },
      required: ['name', 'color'],
    },
  },
  {
    name: 'prepare_delete_category',
    description: 'Prepare to delete a master category. Returns a preview and approval token. Call confirm_delete_category to execute. (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        category_id: { type: 'number', description: 'Category ID to delete' },
      },
      required: ['category_id'],
    },
  },
  {
    name: 'confirm_delete_category',
    description: 'Confirm category deletion with approval token (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        approval_token: { type: 'string', description: 'The approval token from prepare_delete_category' },
      },
      required: ['approval_token'],
    },
  },
  // Focused inbox override tools
  {
    name: 'list_focused_overrides',
    description: 'List all focused inbox overrides (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
  {
    name: 'create_focused_override',
    description: 'Create a focused inbox override for a sender (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        sender_address: { type: 'string', description: 'Sender email address' },
        classify_as: { type: 'string', enum: ['focused', 'other'], description: 'Classification' },
      },
      required: ['sender_address', 'classify_as'],
    },
  },
  {
    name: 'prepare_delete_focused_override',
    description: 'Prepare to delete a focused inbox override. Returns a preview and approval token. Call confirm_delete_focused_override to execute. (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        override_id: { type: 'number', description: 'Override ID to delete' },
      },
      required: ['override_id'],
    },
  },
  {
    name: 'confirm_delete_focused_override',
    description: 'Confirm focused inbox override deletion with approval token (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        approval_token: { type: 'string', description: 'The approval token from prepare_delete_focused_override' },
      },
      required: ['approval_token'],
    },
  },
  // Automatic replies (OOF) tools
  {
    name: 'get_automatic_replies',
    description: 'Get the current automatic replies (out-of-office) settings',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
  {
    name: 'set_automatic_replies',
    description: 'Set automatic replies (out-of-office) settings',
    inputSchema: {
      type: 'object' as const,
      properties: {
        status: {
          type: 'string',
          enum: ['disabled', 'alwaysEnabled', 'scheduled'],
          description: 'OOF status',
        },
        external_audience: {
          type: 'string',
          enum: ['none', 'contactsOnly', 'all'],
          description: 'Who sees external reply',
        },
        internal_reply_message: {
          type: 'string',
          description: 'Reply for internal senders (HTML)',
        },
        external_reply_message: {
          type: 'string',
          description: 'Reply for external senders (HTML)',
        },
        scheduled_start: {
          type: 'string',
          description: 'Schedule start (ISO 8601)',
        },
        scheduled_end: {
          type: 'string',
          description: 'Schedule end (ISO 8601)',
        },
      },
      required: ['status'],
    },
  },
  // Mailbox settings tools
  {
    name: 'get_mailbox_settings',
    description: 'Get the current mailbox settings (language, time zone, date/time formats, working hours)',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
  {
    name: 'update_mailbox_settings',
    description: 'Update mailbox settings (language, time zone, date/time formats)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        language: {
          type: 'string',
          description: 'Locale code (e.g. en-US)',
        },
        time_zone: {
          type: 'string',
          description: 'Time zone (e.g. America/New_York)',
        },
        date_format: {
          type: 'string',
          description: 'Date format string',
        },
        time_format: {
          type: 'string',
          description: 'Time format string',
        },
      },
      required: [],
    },
  },
  // Contact folder tools
  {
    name: 'list_contact_folders',
    description: 'List all contact folders (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
  {
    name: 'create_contact_folder',
    description: 'Create a contact folder (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        name: { type: 'string', description: 'Contact folder name' },
      },
      required: ['name'],
    },
  },
  {
    name: 'prepare_delete_contact_folder',
    description: 'Prepare to delete a contact folder. Returns an approval token. Call confirm_delete_contact_folder to execute. (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        folder_id: { type: 'number', description: 'Contact folder ID to delete' },
      },
      required: ['folder_id'],
    },
  },
  {
    name: 'confirm_delete_contact_folder',
    description: 'Confirm contact folder deletion with approval token (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        token_id: { type: 'string', description: 'The approval token from prepare_delete_contact_folder' },
        folder_id: { type: 'number', description: 'The contact folder ID to delete' },
      },
      required: ['token_id', 'folder_id'],
    },
  },
  // Contact photo tools
  {
    name: 'get_contact_photo',
    description: 'Download a contact\'s photo to a local file (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        contact_id: { type: 'number', description: 'Contact ID' },
      },
      required: ['contact_id'],
    },
  },
  {
    name: 'set_contact_photo',
    description: 'Set or update a contact\'s photo from a local file (JPEG or PNG) (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        contact_id: { type: 'number', description: 'Contact ID' },
        file_path: { type: 'string', description: 'Path to the photo file (JPEG or PNG)' },
      },
      required: ['contact_id', 'file_path'],
    },
  },
  // Mail tips tool
  {
    name: 'get_mail_tips',
    description: 'Get mail tips (automatic replies, mailbox full, delivery restrictions, max message size) for email addresses (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        email_addresses: {
          type: 'array',
          items: { type: 'string' },
          description: 'Email addresses to check (1-20)',
          minItems: 1,
          maxItems: 20,
        },
      },
      required: ['email_addresses'],
    },
  },
  // Message headers & MIME tools
  {
    name: 'get_message_headers',
    description: 'Get internet message headers (SPF, DKIM, routing, etc.) for an email (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        email_id: { type: 'number', description: 'Email ID' },
      },
      required: ['email_id'],
    },
  },
  {
    name: 'get_message_mime',
    description: 'Download the full MIME content (.eml) of an email to a local file (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        email_id: { type: 'number', description: 'Email ID' },
      },
      required: ['email_id'],
    },
  },
  // Calendar Group tools
  {
    name: 'list_calendar_groups',
    description: 'List all calendar groups (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
  {
    name: 'create_calendar_group',
    description: 'Create a new calendar group (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        name: { type: 'string', description: 'Calendar group name' },
      },
      required: ['name'],
    },
  },
  // Calendar Permission tools
  {
    name: 'list_calendar_permissions',
    description: 'List all sharing permissions for a calendar (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        calendar_id: { type: 'number', description: 'Calendar ID' },
      },
      required: ['calendar_id'],
    },
  },
  {
    name: 'create_calendar_permission',
    description: 'Share a calendar with someone by creating a permission (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        calendar_id: { type: 'number', description: 'Calendar ID' },
        email_address: { type: 'string', description: 'Email of person to share with' },
        role: { type: 'string', enum: ['read', 'write', 'delegateWithoutPrivateEventAccess', 'delegateWithPrivateEventAccess'], description: 'Permission level' },
      },
      required: ['calendar_id', 'email_address', 'role'],
    },
  },
  {
    name: 'prepare_delete_calendar_permission',
    description: 'Prepare to delete a calendar sharing permission. Returns a preview and approval token. Call confirm_delete_calendar_permission to execute. (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        permission_id: { type: 'number', description: 'Calendar permission ID to delete' },
      },
      required: ['permission_id'],
    },
  },
  {
    name: 'confirm_delete_calendar_permission',
    description: 'Confirm calendar permission deletion with approval token (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        approval_token: { type: 'string', description: 'The approval token from prepare_delete_calendar_permission' },
      },
      required: ['approval_token'],
    },
  },
  // Room lists & rooms tools
  {
    name: 'list_room_lists',
    description: 'List all room lists (building/floor groupings) in the organization (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
  {
    name: 'list_rooms',
    description: 'List meeting rooms, optionally filtered by a room list email from list_room_lists (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        room_list_email: {
          type: 'string',
          description: 'Room list email to filter by (from list_room_lists)',
        },
      },
      required: [],
    },
  },
  // Teams tools
  {
    name: 'list_teams',
    description: 'List all Microsoft Teams the user has joined (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
  {
    name: 'list_channels',
    description: 'List all channels in a team (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        team_id: { type: 'number', description: 'Team ID from list_teams' },
      },
      required: ['team_id'],
    },
  },
  {
    name: 'get_channel',
    description: 'Get details for a specific channel (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        channel_id: { type: 'number', description: 'Channel ID from list_channels' },
      },
      required: ['channel_id'],
    },
  },
  {
    name: 'create_channel',
    description: 'Create a new channel in a team (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        team_id: { type: 'number', description: 'Team ID from list_teams' },
        name: { type: 'string', description: 'Channel name' },
        description: { type: 'string', description: 'Channel description' },
      },
      required: ['team_id', 'name'],
    },
  },
  {
    name: 'update_channel',
    description: 'Update a channel name or description (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        channel_id: { type: 'number', description: 'Channel ID from list_channels' },
        name: { type: 'string', description: 'New channel name' },
        description: { type: 'string', description: 'New channel description' },
      },
      required: ['channel_id'],
    },
  },
  {
    name: 'prepare_delete_channel',
    description: 'Prepare to delete a channel. Returns an approval token. Call confirm_delete_channel to execute. (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        channel_id: { type: 'number', description: 'Channel ID to delete' },
      },
      required: ['channel_id'],
    },
  },
  {
    name: 'confirm_delete_channel',
    description: 'Confirm channel deletion with approval token (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        approval_token: { type: 'string', description: 'The approval token from prepare_delete_channel' },
      },
      required: ['approval_token'],
    },
  },
  {
    name: 'list_team_members',
    description: 'List all members of a team (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        team_id: { type: 'number', description: 'Team ID from list_teams' },
      },
      required: ['team_id'],
    },
  },
  {
    name: 'list_channel_messages',
    description: 'List recent messages in a channel (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        channel_id: { type: 'number', description: 'Channel ID from list_channels' },
        limit: { type: 'number', description: 'Max messages to return (default 25, max 50)' },
      },
      required: ['channel_id'],
    },
  },
  {
    name: 'get_channel_message',
    description: 'Get a specific channel message with its replies (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        message_id: { type: 'number', description: 'Message ID from list_channel_messages' },
      },
      required: ['message_id'],
    },
  },
  {
    name: 'prepare_send_channel_message',
    description: 'Prepare to send a message to a channel. Returns an approval token. Call confirm_send_channel_message to execute. (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        channel_id: { type: 'number', description: 'Channel ID to send message to' },
        body: { type: 'string', description: 'Message body' },
        content_type: { type: 'string', enum: ['text', 'html'], description: 'Content type (default: html)' },
      },
      required: ['channel_id', 'body'],
    },
  },
  {
    name: 'confirm_send_channel_message',
    description: 'Confirm sending a channel message with approval token (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        approval_token: { type: 'string', description: 'The approval token from prepare_send_channel_message' },
      },
      required: ['approval_token'],
    },
  },
  {
    name: 'prepare_reply_channel_message',
    description: 'Prepare to reply to a channel message. Returns an approval token. Call confirm_reply_channel_message to execute. (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        message_id: { type: 'number', description: 'Message ID to reply to' },
        body: { type: 'string', description: 'Reply body' },
        content_type: { type: 'string', enum: ['text', 'html'], description: 'Content type (default: html)' },
      },
      required: ['message_id', 'body'],
    },
  },
  {
    name: 'confirm_reply_channel_message',
    description: 'Confirm replying to a channel message with approval token (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        approval_token: { type: 'string', description: 'The approval token from prepare_reply_channel_message' },
      },
      required: ['approval_token'],
    },
  },
  {
    name: 'list_chats',
    description: 'List recent 1:1 and group chats (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        limit: { type: 'number', description: 'Max chats to return (default 25, max 50)' },
      },
    },
  },
  {
    name: 'get_chat',
    description: 'Get details of a specific chat (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        chat_id: { type: 'number', description: 'Chat ID from list_chats' },
      },
      required: ['chat_id'],
    },
  },
  {
    name: 'list_chat_messages',
    description: 'List recent messages in a chat (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        chat_id: { type: 'number', description: 'Chat ID from list_chats' },
        limit: { type: 'number', description: 'Max messages to return (default 25, max 50)' },
      },
      required: ['chat_id'],
    },
  },
  {
    name: 'prepare_send_chat_message',
    description: 'Prepare to send a message in a chat. Returns an approval token. Call confirm_send_chat_message to execute. (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        chat_id: { type: 'number', description: 'Chat ID to send message to' },
        body: { type: 'string', description: 'Message body' },
        content_type: { type: 'string', enum: ['text', 'html'], description: 'Content type (default: html)' },
      },
      required: ['chat_id', 'body'],
    },
  },
  {
    name: 'confirm_send_chat_message',
    description: 'Confirm sending a chat message with approval token (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        approval_token: { type: 'string', description: 'The approval token from prepare_send_chat_message' },
      },
      required: ['approval_token'],
    },
  },
  {
    name: 'list_chat_members',
    description: 'List members of a chat (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        chat_id: { type: 'number', description: 'Chat ID from list_chats' },
      },
      required: ['chat_id'],
    },
  },
  // Checklist Items tools
  {
    name: 'list_checklist_items',
    description: 'List checklist items (subtasks) on a To Do task (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        task_id: { type: 'number', description: 'Task ID from list_tasks or search_tasks' },
      },
      required: ['task_id'],
    },
  },
  {
    name: 'create_checklist_item',
    description: 'Create a checklist item (subtask) on a To Do task (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        task_id: { type: 'number', description: 'Task ID' },
        display_name: { type: 'string', description: 'Checklist item text' },
        is_checked: { type: 'boolean', description: 'Whether the item is checked (default: false)' },
      },
      required: ['task_id', 'display_name'],
    },
  },
  {
    name: 'update_checklist_item',
    description: 'Update a checklist item (toggle check, rename) on a To Do task (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        checklist_item_id: { type: 'number', description: 'Checklist item ID' },
        display_name: { type: 'string', description: 'New text' },
        is_checked: { type: 'boolean', description: 'Toggle checked state' },
      },
      required: ['checklist_item_id'],
    },
  },
  {
    name: 'prepare_delete_checklist_item',
    description: 'Prepare to delete a checklist item. Returns an approval token. Call confirm_delete_checklist_item to execute. (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        checklist_item_id: { type: 'number', description: 'Checklist item ID to delete' },
      },
      required: ['checklist_item_id'],
    },
  },
  {
    name: 'confirm_delete_checklist_item',
    description: 'Confirm deletion of a checklist item with approval token (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        approval_token: { type: 'string', description: 'The approval token from prepare_delete_checklist_item' },
      },
      required: ['approval_token'],
    },
  },
  // Linked Resources tools
  {
    name: 'list_linked_resources',
    description: 'List linked resources on a To Do task (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        task_id: { type: 'number', description: 'Task ID from list_tasks or search_tasks' },
      },
      required: ['task_id'],
    },
  },
  {
    name: 'create_linked_resource',
    description: 'Create a linked resource on a To Do task (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        task_id: { type: 'number', description: 'Task ID' },
        web_url: { type: 'string', description: 'URL of the linked resource' },
        application_name: { type: 'string', description: 'Name of the application' },
        display_name: { type: 'string', description: 'Display name of the linked resource' },
      },
      required: ['task_id', 'web_url', 'application_name'],
    },
  },
  {
    name: 'prepare_delete_linked_resource',
    description: 'Prepare to delete a linked resource. Returns an approval token. Call confirm_delete_linked_resource to execute. (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        linked_resource_id: { type: 'number', description: 'Linked resource ID to delete' },
      },
      required: ['linked_resource_id'],
    },
  },
  {
    name: 'confirm_delete_linked_resource',
    description: 'Confirm deletion of a linked resource with approval token (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        approval_token: { type: 'string', description: 'The approval token from prepare_delete_linked_resource' },
      },
      required: ['approval_token'],
    },
  },
  // Task Attachments tools
  {
    name: 'list_task_attachments',
    description: 'List attachments on a To Do task (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        task_id: { type: 'number', description: 'Task ID from list_tasks or search_tasks' },
      },
      required: ['task_id'],
    },
  },
  {
    name: 'create_task_attachment',
    description: 'Attach a file to a To Do task (small files only, base64 encoded) (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        task_id: { type: 'number', description: 'Task ID' },
        name: { type: 'string', description: 'File name of the attachment' },
        content_bytes: { type: 'string', description: 'Base64-encoded file content' },
        content_type: { type: 'string', description: 'MIME type (default: application/octet-stream)' },
      },
      required: ['task_id', 'name', 'content_bytes'],
    },
  },
  {
    name: 'prepare_delete_task_attachment',
    description: 'Prepare to delete a task attachment. Returns an approval token. Call confirm_delete_task_attachment to execute. (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        task_attachment_id: { type: 'number', description: 'Task attachment ID to delete' },
      },
      required: ['task_attachment_id'],
    },
  },
  {
    name: 'confirm_delete_task_attachment',
    description: 'Confirm deletion of a task attachment with approval token (Graph API)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        approval_token: { type: 'string', description: 'The approval token from prepare_delete_task_attachment' },
      },
      required: ['approval_token'],
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
      name: 'office365-mcp',
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
  let sendTools: ReturnType<typeof createMailSendTools> | null = null;
  let schedulingTools: ReturnType<typeof createSchedulingTools> | null = null;
  let rulesTools: MailRulesTools | null = null;
  let categoriesTools: CategoriesTools | null = null;
  let calendarPermissionsTools: CalendarPermissionsTools | null = null;
  let focusedOverridesTools: FocusedOverridesTools | null = null;
  let teamsTools: TeamsTools | null = null;
  let checklistItemsTools: ChecklistItemsTools | null = null;
  let linkedResourcesTools: LinkedResourcesTools | null = null;
  let taskAttachmentsTools: TaskAttachmentsTools | null = null;
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
    mailTools = createMailTools(repository, contentReaders.email, contentReaders.attachment);
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
   * If not authenticated, triggers the device code flow inline.
   */
  const initializeGraphBackend = createAuthMutex(async (): Promise<void> => {
    // Try to authenticate if needed (triggers device code flow for first-time users)
    const authenticated = await isAuthenticated();
    if (!authenticated) {
      await getAccessToken();
    }

    graphRepository = createGraphRepository();
    graphContentReaders = createGraphContentReadersWithClient(graphRepository.getClient());

    const adapter = new GraphMailboxAdapter(graphRepository);
    orgTools = createMailboxOrganizationTools(adapter, tokenManager);
    sendTools = createMailSendTools(graphRepository, tokenManager);
    schedulingTools = createSchedulingTools(graphRepository);
    rulesTools = new MailRulesTools(graphRepository, tokenManager);
    categoriesTools = new CategoriesTools(graphRepository, tokenManager);
    calendarPermissionsTools = new CalendarPermissionsTools(graphRepository, tokenManager);
    focusedOverridesTools = new FocusedOverridesTools(graphRepository, tokenManager);
    teamsTools = new TeamsTools(graphRepository, tokenManager);
    checklistItemsTools = new ChecklistItemsTools(graphRepository, tokenManager);
    linkedResourcesTools = new LinkedResourcesTools(graphRepository, tokenManager);
    taskAttachmentsTools = new TaskAttachmentsTools(graphRepository, tokenManager);

    initialized = true;
  });

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

  // Tools that only exist when using Graph API (signature + scheduling)
  const GRAPH_ONLY_TOOL_NAMES = new Set([
    'set_signature',
    'get_signature',
    'check_availability',
    'find_meeting_times',
    'list_conversation',
    'search_emails_advanced',
    'check_new_emails',
    'list_mail_rules',
    'create_mail_rule',
    'prepare_delete_mail_rule',
    'confirm_delete_mail_rule',
    'list_categories',
    'create_category',
    'prepare_delete_category',
    'confirm_delete_category',
    'list_focused_overrides',
    'create_focused_override',
    'prepare_delete_focused_override',
    'confirm_delete_focused_override',
    'list_task_lists',
    'rename_task_list',
    'prepare_delete_task_list',
    'confirm_delete_task_list',
    'get_automatic_replies',
    'set_automatic_replies',
    'get_mailbox_settings',
    'update_mailbox_settings',
    'list_contact_folders',
    'create_contact_folder',
    'prepare_delete_contact_folder',
    'confirm_delete_contact_folder',
    'get_contact_photo',
    'set_contact_photo',
    'list_event_instances',
    'get_mail_tips',
    'get_message_headers',
    'get_message_mime',
    'list_calendar_groups',
    'create_calendar_group',
    'list_calendar_permissions',
    'create_calendar_permission',
    'prepare_delete_calendar_permission',
    'confirm_delete_calendar_permission',
    'list_room_lists',
    'list_rooms',
    'list_teams',
    'list_channels',
    'get_channel',
    'create_channel',
    'update_channel',
    'prepare_delete_channel',
    'confirm_delete_channel',
    'list_team_members',
    'list_channel_messages',
    'get_channel_message',
    'prepare_send_channel_message',
    'confirm_send_channel_message',
    'prepare_reply_channel_message',
    'confirm_reply_channel_message',
    'list_chats',
    'get_chat',
    'list_chat_messages',
    'prepare_send_chat_message',
    'confirm_send_chat_message',
    'list_chat_members',
    'list_checklist_items',
    'create_checklist_item',
    'update_checklist_item',
    'prepare_delete_checklist_item',
    'confirm_delete_checklist_item',
    'list_linked_resources',
    'create_linked_resource',
    'prepare_delete_linked_resource',
    'confirm_delete_linked_resource',
    'list_task_attachments',
    'create_task_attachment',
    'prepare_delete_task_attachment',
    'confirm_delete_task_attachment',
  ]);

  // Register tool list handler
  server.setRequestHandler(ListToolsRequestSchema, () => {
    const tools = useGraphApi ? TOOLS : TOOLS.filter((t) => !GRAPH_ONLY_TOOL_NAMES.has(t.name));
    return { tools };
  });

  // Register tool call handler (async for Graph API support)
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    try {
      await ensureInitialized();

      // Graph API mode - handle async operations directly
      if (useGraphApi && graphRepository != null) {
        return await handleGraphToolCall(name, args, graphRepository, graphContentReaders!, orgTools!, sendTools!, schedulingTools!, rulesTools!, categoriesTools!, calendarPermissionsTools!, focusedOverridesTools!, teamsTools!, checklistItemsTools!, linkedResourcesTools!, taskAttachmentsTools!, tokenManager);
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
    case 'set_email_importance': {
      const params = SetEmailImportanceInput.parse(args);
      const result = await orgTools.setEmailImportance(params);
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
// Mail Send Tool Handler (Graph API only)
// =============================================================================

async function handleSendToolCall(
  name: string,
  args: unknown,
  sendTools: ReturnType<typeof createMailSendTools>
): Promise<ToolResult | null> {
  switch (name) {
    // Non-Destructive — Draft Management
    case 'create_draft': {
      const params = CreateDraftInput.parse(args);
      const result = await sendTools.createDraft(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'update_draft': {
      const params = UpdateDraftInput.parse(args);
      const result = await sendTools.updateDraft(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'add_draft_attachment': {
      const params = AddDraftAttachmentInput.parse(args);
      const result = await sendTools.addDraftAttachment(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'add_draft_inline_image': {
      const params = AddDraftInlineImageInput.parse(args);
      const result = await sendTools.addDraftInlineImage(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'list_drafts': {
      const params = ListDraftsInput.parse(args ?? {});
      const result = await sendTools.listDrafts(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    // Two-Phase — Send Draft
    case 'prepare_send_draft': {
      const params = PrepareSendDraftInput.parse(args);
      const result = await sendTools.prepareSendDraft(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'confirm_send_draft': {
      const params = ConfirmSendDraftInput.parse(args);
      const result = await sendTools.confirmSendDraft(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    // Two-Phase — Send Email (Direct)
    case 'prepare_send_email': {
      const params = PrepareSendEmailInput.parse(args);
      const result = sendTools.prepareSendEmail(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'confirm_send_email': {
      const params = ConfirmSendEmailInput.parse(args);
      const result = await sendTools.confirmSendEmail(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    // Two-Phase — Reply Email
    case 'prepare_reply_email': {
      const params = PrepareReplyEmailInput.parse(args);
      const result = await sendTools.prepareReplyEmail(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'confirm_reply_email': {
      const params = ConfirmReplyEmailInput.parse(args);
      const result = await sendTools.confirmReplyEmail(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    // Two-Phase — Forward Email
    case 'prepare_forward_email': {
      const params = PrepareForwardEmailInput.parse(args);
      const result = await sendTools.prepareForwardEmail(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'confirm_forward_email': {
      const params = ConfirmForwardEmailInput.parse(args);
      const result = await sendTools.confirmForwardEmail(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    // Draft Reply/Forward
    case 'reply_as_draft': {
      const params = ReplyAsDraftInput.parse(args);
      const result = await sendTools.replyAsDraft(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'forward_as_draft': {
      const params = ForwardAsDraftInput.parse(args);
      const result = await sendTools.forwardAsDraft(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'set_signature': {
      const params = SetSignatureInput.parse(args);
      const result = sendTools.setSignature(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'get_signature': {
      GetSignatureInput.parse(args ?? {});
      const result = sendTools.getSignature();
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    default:
      return null;
  }
}

// =============================================================================
// Scheduling Tool Handler (Graph API only)
// =============================================================================

async function handleSchedulingToolCall(
  name: string,
  args: unknown,
  schedulingTools: ReturnType<typeof createSchedulingTools>
): Promise<ToolResult | null> {
  switch (name) {
    case 'check_availability': {
      const params = CheckAvailabilityInput.parse(args);
      const result = await schedulingTools.checkAvailability(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'find_meeting_times': {
      const params = FindMeetingTimesInput.parse(args);
      const result = await schedulingTools.findMeetingTimes(params);
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

    case 'get_emails': {
      const params = GetEmailsInput.parse(args);
      const results = params.email_ids.map((id) => {
        const email = mailTools.getEmail({ email_id: id, include_body: params.include_body, strip_html: params.strip_html });
        if (email == null) return { id, error: 'Not found' };
        return email;
      });
      return { content: [{ type: 'text', text: JSON.stringify({ emails: results }, null, 2) }] };
    }

    case 'get_unread_count': {
      const params = GetUnreadCountInput.parse(args ?? {});
      const result = mailTools.getUnreadCount(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    // Attachment tools
    case 'list_attachments': {
      const params = ListAttachmentsInput.parse(args);
      const result = mailTools.listAttachments(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }

    case 'download_attachment': {
      const params = DownloadAttachmentInput.parse(args);
      const result = mailTools.downloadAttachment(params);
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
        inline_images?: Array<{ path: string; content_id: string }>;
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
      if (params.inline_images != null) {
        sendParams = {
          ...sendParams,
          inlineImages: params.inline_images.map(img => ({
            path: img.path,
            contentId: img.content_id,
          })),
        };
      }
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
// Calendar Write — Zod Schemas (Graph API)
// =============================================================================

const GraphCreateEventInput = z.strictObject({
  title: z.string().min(1),
  start_date: z.string().refine((s) => !isNaN(Date.parse(s)), { message: 'Must be a valid ISO 8601 date string' }),
  end_date: z.string().refine((s) => !isNaN(Date.parse(s)), { message: 'Must be a valid ISO 8601 date string' }),
  calendar_id: z.number().int().positive().optional(),
  location: z.string().optional(),
  description: z.string().optional(),
  is_all_day: z.boolean().optional().default(false),
  attendees: z.array(z.object({
    email: z.string().email(),
    name: z.string().optional(),
    type: z.enum(['required', 'optional']).optional(),
  })).optional(),
  recurrence: z.object({
    pattern: z.object({
      type: z.enum(['daily', 'weekly', 'monthly', 'yearly']),
      interval: z.number().int().positive(),
      daysOfWeek: z.array(z.string()).optional(),
    }),
    range: z.object({
      type: z.enum(['endDate', 'noEnd', 'numbered']),
      startDate: z.string(),
      endDate: z.string().optional(),
      numberOfOccurrences: z.number().int().positive().optional(),
    }),
  }).optional(),
  is_online_meeting: z.boolean().optional().describe('Create as online Teams meeting'),
  online_meeting_provider: z.enum(['teamsForBusiness', 'skypeForBusiness', 'skypeForConsumer']).optional().describe('Online meeting provider (default: teamsForBusiness)'),
}).refine(
  (data) => new Date(data.start_date).getTime() < new Date(data.end_date).getTime(),
  { message: 'start_date must be before end_date', path: ['start_date'] }
);

const UpdateEventInput = z.strictObject({
  event_id: z.number().int().positive(),
  subject: z.string().optional(),
  start: z.string().optional(),
  end: z.string().optional(),
  timezone: z.string().optional(),
  location: z.string().optional(),
  body: z.string().optional(),
  body_type: z.enum(['text', 'html']).optional(),
  attendees: z.array(z.object({
    email: z.string().email(),
    name: z.string().optional(),
    type: z.enum(['required', 'optional']).optional(),
  })).optional(),
  is_all_day: z.boolean().optional(),
  recurrence: z.object({
    pattern: z.object({
      type: z.enum(['daily', 'weekly', 'monthly', 'yearly']),
      interval: z.number().int().positive(),
      daysOfWeek: z.array(z.string()).optional(),
    }),
    range: z.object({
      type: z.enum(['endDate', 'noEnd', 'numbered']),
      startDate: z.string(),
      endDate: z.string().optional(),
      numberOfOccurrences: z.number().int().positive().optional(),
    }),
  }).optional(),
  is_online_meeting: z.boolean().optional().describe('Create as online Teams meeting'),
  online_meeting_provider: z.enum(['teamsForBusiness', 'skypeForBusiness', 'skypeForConsumer']).optional().describe('Online meeting provider (default: teamsForBusiness)'),
});

const RespondToEventGraphInput = z.strictObject({
  event_id: z.number().int().positive(),
  response: z.enum(['accept', 'decline', 'tentative']),
  send_response: z.boolean().default(true),
  comment: z.string().optional(),
});

const PrepareDeleteEventInput = z.strictObject({
  event_id: z.number().int().positive(),
});

const ConfirmDeleteEventInput = z.strictObject({
  token_id: z.uuid(),
  event_id: z.number().int().positive(),
});

const ListEventInstancesInput = z.strictObject({
  event_id: z.number().int().positive().describe('Recurring event ID'),
  start_date: z.string().describe('Start of date range (ISO 8601, e.g. 2024-01-01T00:00:00Z)'),
  end_date: z.string().describe('End of date range (ISO 8601, e.g. 2024-12-31T23:59:59Z)'),
});

// =============================================================================
// Contact Write — Zod Schemas (Graph API)
// =============================================================================

const CreateContactGraphInput = z.strictObject({
  given_name: z.string().optional(),
  surname: z.string().optional(),
  email: z.string().email().optional(),
  phone: z.string().optional(),
  mobile_phone: z.string().optional(),
  company: z.string().optional(),
  job_title: z.string().optional(),
  street_address: z.string().optional(),
  city: z.string().optional(),
  state: z.string().optional(),
  postal_code: z.string().optional(),
  country: z.string().optional(),
});

const UpdateContactGraphInput = z.strictObject({
  contact_id: z.number().int().positive(),
  given_name: z.string().optional(),
  surname: z.string().optional(),
  email: z.string().email().optional(),
  phone: z.string().optional(),
  mobile_phone: z.string().optional(),
  company: z.string().optional(),
  job_title: z.string().optional(),
  street_address: z.string().optional(),
  city: z.string().optional(),
  state: z.string().optional(),
  postal_code: z.string().optional(),
  country: z.string().optional(),
});

const PrepareDeleteContactInput = z.strictObject({
  contact_id: z.number().int().positive(),
});

const ConfirmDeleteContactInput = z.strictObject({
  token_id: z.uuid(),
  contact_id: z.number().int().positive(),
});

// =============================================================================
// Task Write — Zod Schemas (Graph API)
// =============================================================================

const RecurrenceSchema = z.strictObject({
  pattern: z.enum(['daily', 'weekly', 'monthly', 'yearly']).describe('Recurrence pattern type'),
  interval: z.number().int().min(1).default(1).describe('Interval between occurrences'),
  days_of_week: z.array(z.enum(['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'])).optional().describe('Days of week (for weekly pattern)'),
  day_of_month: z.number().int().min(1).max(31).optional().describe('Day of month (for monthly pattern)'),
  range_type: z.enum(['endDate', 'noEnd', 'numbered']).describe('How the recurrence ends'),
  start_date: z.string().describe('Start date (YYYY-MM-DD)'),
  end_date: z.string().optional().describe('End date (YYYY-MM-DD, for endDate range)'),
  occurrences: z.number().int().min(1).optional().describe('Number of occurrences (for numbered range)'),
}).optional().describe('Task recurrence settings');

const CreateTaskGraphInput = z.strictObject({
  title: z.string().min(1),
  task_list_id: z.number().int().positive(),
  body: z.string().optional(),
  body_type: z.enum(['text', 'html']).optional(),
  due_date: z.string().optional(),
  importance: z.enum(['low', 'normal', 'high']).optional(),
  reminder_date: z.string().optional(),
  recurrence: RecurrenceSchema,
  categories: z.array(z.string()).optional(),
});

const UpdateTaskGraphInput = z.strictObject({
  task_id: z.number().int().positive(),
  title: z.string().optional(),
  body: z.string().optional(),
  body_type: z.enum(['text', 'html']).optional(),
  due_date: z.string().optional(),
  importance: z.enum(['low', 'normal', 'high']).optional(),
  reminder_date: z.string().optional(),
  status: z.enum(['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred']).optional(),
  recurrence: RecurrenceSchema,
  categories: z.array(z.string()).optional(),
});

const CompleteTaskGraphInput = z.strictObject({
  task_id: z.number().int().positive(),
});

const CreateTaskListGraphInput = z.strictObject({
  display_name: z.string().min(1),
});

const PrepareDeleteTaskInput = z.strictObject({
  task_id: z.number().int().positive(),
});

const ConfirmDeleteTaskInput = z.strictObject({
  token_id: z.uuid(),
  task_id: z.number().int().positive(),
});

const RenameTaskListInput = z.strictObject({
  task_list_id: z.number().int().positive().describe('Task list ID'),
  name: z.string().min(1).describe('New name for the task list'),
});

const PrepareDeleteTaskListInput = z.strictObject({
  task_list_id: z.number().int().positive().describe('Task list ID to delete'),
});

const ConfirmDeleteTaskListInput = z.strictObject({
  token_id: z.string().uuid().describe('Approval token from prepare_delete_task_list'),
  task_list_id: z.number().int().positive().describe('The task list ID to delete'),
});

const GetAutomaticRepliesInput = z.strictObject({});

const SetAutomaticRepliesInput = z.strictObject({
  status: z.enum(['disabled', 'alwaysEnabled', 'scheduled']).describe('OOF status'),
  external_audience: z.enum(['none', 'contactsOnly', 'all']).optional().describe('Who sees external reply'),
  internal_reply_message: z.string().optional().describe('Reply for internal senders (HTML)'),
  external_reply_message: z.string().optional().describe('Reply for external senders (HTML)'),
  scheduled_start: z.string().optional().describe('Schedule start (ISO 8601)'),
  scheduled_end: z.string().optional().describe('Schedule end (ISO 8601)'),
});

const GetMailboxSettingsInput = z.strictObject({});

const UpdateMailboxSettingsInput = z.strictObject({
  language: z.string().optional().describe('Locale code (e.g. en-US)'),
  time_zone: z.string().optional().describe('Time zone (e.g. America/New_York)'),
  date_format: z.string().optional().describe('Date format string'),
  time_format: z.string().optional().describe('Time format string'),
});

const GetMailTipsInput = z.strictObject({
  email_addresses: z.array(z.string().email()).min(1).max(20).describe('Email addresses to check'),
});

const GetMessageHeadersInput = z.strictObject({
  email_id: z.number().int().positive().describe('Email ID'),
});

const GetMessageMimeInput = z.strictObject({
  email_id: z.number().int().positive().describe('Email ID'),
});

const CreateCalendarGroupInput = z.strictObject({
  name: z.string().min(1).describe('Calendar group name'),
});

const CreateContactFolderInput = z.strictObject({
  name: z.string().min(1).describe('Contact folder name'),
});

const PrepareDeleteContactFolderInput = z.strictObject({
  folder_id: z.number().int().positive().describe('Contact folder ID to delete'),
});

const ConfirmDeleteContactFolderInput = z.strictObject({
  token_id: z.string().uuid().describe('Approval token from prepare_delete_contact_folder'),
  folder_id: z.number().int().positive().describe('The contact folder ID to delete'),
});

const GetContactPhotoInput = z.strictObject({
  contact_id: z.number().int().positive().describe('Contact ID'),
});

const SetContactPhotoInput = z.strictObject({
  contact_id: z.number().int().positive().describe('Contact ID'),
  file_path: z.string().describe('Path to the photo file (JPEG or PNG)'),
});

const ListRoomsInput = z.strictObject({
  room_list_email: z.string().email().optional().describe('Room list email to filter by (from list_room_lists)'),
});

// =============================================================================
// Graph API Tool Handler
// =============================================================================

async function handleGraphToolCall(
  name: string,
  args: unknown,
  repository: GraphRepository,
  contentReaders: GraphContentReaders,
  orgTools: ReturnType<typeof createMailboxOrganizationTools>,
  sendTools: ReturnType<typeof createMailSendTools>,
  schedulingTools: ReturnType<typeof createSchedulingTools>,
  rulesTools: MailRulesTools,
  categoriesTools: CategoriesTools,
  calendarPermissionsTools: CalendarPermissionsTools,
  focusedOverridesTools: FocusedOverridesTools,
  teamsTools: TeamsTools,
  checklistItemsTools: ChecklistItemsTools,
  linkedResourcesTools: LinkedResourcesTools,
  taskAttachmentsTools: TaskAttachmentsTools,
  tokenManager: ApprovalTokenManager
): Promise<ToolResult> {
  // Handle mailbox organization tools (shared between backends)
  const orgResult = await handleOrgToolCall(name, args, orgTools);
  if (orgResult != null) return orgResult;

  // Handle mail send tools (Graph API only)
  const sendResult = await handleSendToolCall(name, args, sendTools);
  if (sendResult != null) return sendResult;

  // Handle scheduling tools (Graph API only)
  const schedulingResult = await handleSchedulingToolCall(name, args, schedulingTools);
  if (schedulingResult != null) return schedulingResult;

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

      case 'search_emails_advanced': {
        const params = SearchEmailsAdvancedInput.parse(args);
        const emails = params.folder_id != null
          ? await repository.searchEmailsAdvancedInFolderAsync(params.folder_id, params.query, params.limit)
          : await repository.searchEmailsAdvancedAsync(params.query, params.limit);
        const result = { emails: emails.map(transformEmailRow) };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'check_new_emails': {
        const params = CheckNewEmailsInput.parse(args);
        const deltaResult = await repository.checkNewEmailsAsync(params.folder_id);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              emails: deltaResult.emails.map(transformEmailRow),
              is_initial_sync: deltaResult.isInitialSync,
              count: deltaResult.emails.length,
            }, null, 2),
          }],
        };
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

      case 'get_emails': {
        const params = GetEmailsInput.parse(args);
        const results = await Promise.all(
          params.email_ids.map(async (id) => {
            const email = await repository.getEmailAsync(id);
            if (email == null) return { id, error: 'Not found' };
            let body: string | null = null;
            if (params.include_body) {
              body = await contentReaders.email.readEmailBodyAsync(email.dataFilePath);
              if (params.strip_html && body != null) body = stripHtml(body);
            }
            return { ...transformEmailRow(email), body };
          })
        );
        return { content: [{ type: 'text', text: JSON.stringify({ emails: results }, null, 2) }] };
      }

      case 'list_conversation': {
        const params = ListConversationInput.parse(args);
        const emails = await repository.listConversationAsync(params.message_id, params.limit);
        const result = { emails: emails.map(transformEmailRow) };
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

      // Attachment tools
      case 'list_attachments': {
        const params = ListAttachmentsInput.parse(args);
        const attachments = await repository.listAttachmentsAsync(params.email_id);
        const result = { attachments };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'download_attachment': {
        const params = DownloadAttachmentInput.parse(args);
        const result = await repository.downloadAttachmentAsync(params.attachment_index);
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
        const result = { events: events.map(transformGraphEventRow) };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'get_event': {
        const params = GetEventInput.parse(args);
        const event = await repository.getEventAsync(params.event_id);
        if (event == null) {
          return { content: [{ type: 'text', text: 'Event not found' }], isError: true };
        }

        const details = await contentReaders.event.readEventDetailsAsync(event.dataFilePath);
        const result = { ...transformGraphEventRow(event), ...details };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'search_events': {
        const params = SearchEventsInput.parse(args);
        // Graph doesn't have direct event search, so we filter client-side
        const allEvents = await repository.listEventsAsync(1000);
        const events = allEvents.filter((e) => {
          const row = transformGraphEventRow(e);
          // Filter by title if query provided
          if (params.query != null) {
            const title = row.title?.toLowerCase() ?? '';
            if (!title.includes(params.query.toLowerCase())) return false;
          }
          // Filter by date range if provided
          if (params.start_date != null && row.startDate != null) {
            if (new Date(row.startDate) < new Date(params.start_date)) return false;
          }
          if (params.end_date != null && row.endDate != null) {
            if (new Date(row.endDate) > new Date(params.end_date)) return false;
          }
          return true;
        });
        const result = { events: events.slice(0, params.limit).map(transformGraphEventRow) };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'create_event': {
        const params = GraphCreateEventInput.parse(args);
        const createParams: Parameters<typeof repository.createEventAsync>[0] = {
          subject: params.title,
          start: params.start_date,
          end: params.end_date,
        };
        if (params.location != null) createParams.location = params.location;
        if (params.description != null) createParams.body = params.description;
        createParams.bodyType = 'text';
        if (params.is_all_day) createParams.isAllDay = params.is_all_day;
        if (params.attendees != null) {
          createParams.attendees = params.attendees.map((a) => {
            const att: { email: string; name?: string; type?: 'required' | 'optional' } = { email: a.email };
            if (a.name != null) att.name = a.name;
            if (a.type != null) att.type = a.type;
            return att;
          });
        }
        if (params.recurrence != null) {
          const rec = params.recurrence;
          const pattern: { type: 'daily' | 'weekly' | 'monthly' | 'yearly'; interval: number; daysOfWeek?: string[] } = {
            type: rec.pattern.type,
            interval: rec.pattern.interval,
          };
          if (rec.pattern.daysOfWeek != null) pattern.daysOfWeek = rec.pattern.daysOfWeek;
          const range: { type: 'endDate' | 'noEnd' | 'numbered'; startDate: string; endDate?: string; numberOfOccurrences?: number } = {
            type: rec.range.type,
            startDate: rec.range.startDate,
          };
          if (rec.range.endDate != null) range.endDate = rec.range.endDate;
          if (rec.range.numberOfOccurrences != null) range.numberOfOccurrences = rec.range.numberOfOccurrences;
          createParams.recurrence = { pattern, range };
        }
        if (params.calendar_id != null) createParams.calendarId = params.calendar_id;
        if (params.is_online_meeting != null) createParams.is_online_meeting = params.is_online_meeting;
        if (params.online_meeting_provider != null) createParams.online_meeting_provider = params.online_meeting_provider;
        const numericId = await repository.createEventAsync(createParams);

        const result: CreateEventResult = {
          id: numericId,
          title: params.title,
          start_date: params.start_date,
          end_date: params.end_date,
          calendar_id: params.calendar_id ?? null,
          location: params.location ?? null,
          description: params.description ?? null,
          is_all_day: params.is_all_day,
          is_recurring: params.recurrence != null,
        };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'update_event': {
        const params = UpdateEventInput.parse(args);
        const updates: Record<string, unknown> = {};
        const tz = params.timezone ?? Intl.DateTimeFormat().resolvedOptions().timeZone;

        if (params.subject != null) updates.subject = params.subject;
        if (params.start != null) updates.start = { dateTime: params.start, timeZone: tz };
        if (params.end != null) updates.end = { dateTime: params.end, timeZone: tz };
        if (params.location != null) updates.location = { displayName: params.location };
        if (params.body != null) {
          updates.body = {
            contentType: params.body_type ?? 'text',
            content: params.body,
          };
        }
        if (params.is_all_day != null) updates.isAllDay = params.is_all_day;
        if (params.attendees != null) {
          updates.attendees = params.attendees.map((a) => ({
            emailAddress: { address: a.email, name: a.name },
            type: a.type ?? 'required',
          }));
        }
        if (params.recurrence != null) updates.recurrence = params.recurrence;
        if (params.is_online_meeting != null) {
          updates.isOnlineMeeting = params.is_online_meeting;
          if (params.is_online_meeting) {
            updates.onlineMeetingProvider = params.online_meeting_provider ?? 'teamsForBusiness';
          }
        }

        await repository.updateEventAsync(params.event_id, updates);
        return {
          content: [{ type: 'text', text: `Successfully updated event ${params.event_id}` }],
        };
      }

      case 'respond_to_event': {
        const params = RespondToEventGraphInput.parse(args);
        await repository.respondToEventAsync(
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
          content: [{ type: 'text', text: `Successfully ${responseText} event ${params.event_id}` }],
        };
      }

      case 'delete_event': {
        // For Graph API, direct delete_event is also supported (for AppleScript compatibility)
        const params = PrepareDeleteEventInput.parse(args);
        await repository.deleteEventAsync(params.event_id);
        return {
          content: [{ type: 'text', text: `Successfully deleted event ${params.event_id}` }],
        };
      }

      case 'list_event_instances': {
        const params = ListEventInstancesInput.parse(args);
        const instances = await repository.listEventInstancesAsync(params.event_id, params.start_date, params.end_date);
        return { content: [{ type: 'text', text: JSON.stringify({ instances: instances.map(transformGraphEventRow), count: instances.length }, null, 2) }] };
      }

      case 'prepare_delete_event': {
        const params = PrepareDeleteEventInput.parse(args);
        const event = await repository.getEventAsync(params.event_id);
        if (event == null) {
          return { content: [{ type: 'text', text: 'Event not found' }], isError: true };
        }

        const graphId = repository.getGraphId('event', params.event_id);
        const graphEvent = graphId != null ? await repository.getClient().getEvent(graphId) : null;
        const hash = hashEventForApproval({
          id: params.event_id,
          subject: graphEvent?.subject ?? null,
          startDateTime: graphEvent?.start?.dateTime ?? null,
        });

        const token = tokenManager.generateToken({
          operation: 'delete_event',
          targetType: 'event',
          targetId: params.event_id,
          targetHash: hash,
        });

        const result = {
          token_id: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          event: transformGraphEventRow(event),
          action: 'This event will be permanently deleted.',
        };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'confirm_delete_event': {
        const params = ConfirmDeleteEventInput.parse(args);

        // Re-fetch the event and compute fresh hash for comparison
        const graphId = repository.getGraphId('event', params.event_id);
        const graphEvent = graphId != null ? await repository.getClient().getEvent(graphId) : null;
        const currentHash = hashEventForApproval({
          id: params.event_id,
          subject: graphEvent?.subject ?? null,
          startDateTime: graphEvent?.start?.dateTime ?? null,
        });

        const validation = tokenManager.consumeToken(params.token_id, 'delete_event', params.event_id);
        if (!validation.valid) {
          const errorMessages: Record<string, string> = {
            NOT_FOUND: 'Token not found or already used',
            EXPIRED: 'Token has expired. Please call prepare_delete_event again.',
            OPERATION_MISMATCH: 'Token was not generated for delete_event',
            TARGET_MISMATCH: 'Token was generated for a different event',
            ALREADY_CONSUMED: 'Token has already been used',
          };
          return {
            content: [{ type: 'text', text: errorMessages[validation.error ?? ''] ?? 'Invalid token' }],
            isError: true,
          };
        }

        // Check that the event hasn't changed since prepare
        if (validation.token!.targetHash !== currentHash) {
          return {
            content: [{ type: 'text', text: 'Event has changed since prepare was called. Please call prepare_delete_event again.' }],
            isError: true,
          };
        }

        await repository.deleteEventAsync(params.event_id);
        return {
          content: [{ type: 'text', text: `Successfully deleted event ${params.event_id}` }],
        };
      }

      // Contact tools
      case 'list_contacts': {
        const params = ListContactsInput.parse(args ?? {});
        const contacts = params.folder_id != null
          ? await repository.listContactsInFolderAsync(params.folder_id, params.limit)
          : await repository.listContactsAsync(params.limit, params.offset);
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

      case 'create_contact': {
        const params = CreateContactGraphInput.parse(args);
        const numericId = await repository.createContactAsync({
          ...(params.given_name != null ? { given_name: params.given_name } : {}),
          ...(params.surname != null ? { surname: params.surname } : {}),
          ...(params.email != null ? { email: params.email } : {}),
          ...(params.phone != null ? { phone: params.phone } : {}),
          ...(params.mobile_phone != null ? { mobile_phone: params.mobile_phone } : {}),
          ...(params.company != null ? { company: params.company } : {}),
          ...(params.job_title != null ? { job_title: params.job_title } : {}),
          ...(params.street_address != null ? { street_address: params.street_address } : {}),
          ...(params.city != null ? { city: params.city } : {}),
          ...(params.state != null ? { state: params.state } : {}),
          ...(params.postal_code != null ? { postal_code: params.postal_code } : {}),
          ...(params.country != null ? { country: params.country } : {}),
        });
        const result = {
          id: numericId,
          given_name: params.given_name ?? null,
          surname: params.surname ?? null,
          email: params.email ?? null,
          status: 'created',
        };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'update_contact': {
        const params = UpdateContactGraphInput.parse(args);
        const updates: Record<string, unknown> = {};
        if (params.given_name != null) updates.givenName = params.given_name;
        if (params.surname != null) updates.surname = params.surname;
        if (params.email != null) updates.emailAddresses = [{ address: params.email }];
        if (params.phone != null) updates.businessPhones = [params.phone];
        if (params.mobile_phone != null) updates.mobilePhone = params.mobile_phone;
        if (params.company != null) updates.companyName = params.company;
        if (params.job_title != null) updates.jobTitle = params.job_title;
        if (params.street_address != null || params.city != null || params.state != null || params.postal_code != null || params.country != null) {
          const address: Record<string, string> = {};
          if (params.street_address != null) address.street = params.street_address;
          if (params.city != null) address.city = params.city;
          if (params.state != null) address.state = params.state;
          if (params.postal_code != null) address.postalCode = params.postal_code;
          if (params.country != null) address.countryOrRegion = params.country;
          updates.businessAddress = address;
        }
        await repository.updateContactAsync(params.contact_id, updates);
        return {
          content: [{ type: 'text', text: `Successfully updated contact ${params.contact_id}` }],
        };
      }

      case 'prepare_delete_contact': {
        const params = PrepareDeleteContactInput.parse(args);
        const contact = await repository.getContactAsync(params.contact_id);
        if (contact == null) {
          return { content: [{ type: 'text', text: 'Contact not found' }], isError: true };
        }

        const graphId = repository.getGraphId('contact', params.contact_id);
        const graphContact = graphId != null ? await repository.getClient().getContact(graphId) : null;
        const hash = hashContactForApproval({
          id: params.contact_id,
          displayName: graphContact?.displayName ?? null,
          emailAddress: graphContact?.emailAddresses?.[0]?.address ?? null,
        });

        const token = tokenManager.generateToken({
          operation: 'delete_contact',
          targetType: 'contact',
          targetId: params.contact_id,
          targetHash: hash,
        });

        const result = {
          token_id: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          contact: transformContactRow(contact),
          action: 'This contact will be permanently deleted.',
        };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'confirm_delete_contact': {
        const params = ConfirmDeleteContactInput.parse(args);

        const graphId = repository.getGraphId('contact', params.contact_id);
        const graphContact = graphId != null ? await repository.getClient().getContact(graphId) : null;
        const currentHash = hashContactForApproval({
          id: params.contact_id,
          displayName: graphContact?.displayName ?? null,
          emailAddress: graphContact?.emailAddresses?.[0]?.address ?? null,
        });

        const validation = tokenManager.consumeToken(params.token_id, 'delete_contact', params.contact_id);
        if (!validation.valid) {
          const errorMessages: Record<string, string> = {
            NOT_FOUND: 'Token not found or already used',
            EXPIRED: 'Token has expired. Please call prepare_delete_contact again.',
            OPERATION_MISMATCH: 'Token was not generated for delete_contact',
            TARGET_MISMATCH: 'Token was generated for a different contact',
            ALREADY_CONSUMED: 'Token has already been used',
          };
          return {
            content: [{ type: 'text', text: errorMessages[validation.error ?? ''] ?? 'Invalid token' }],
            isError: true,
          };
        }

        if (validation.token!.targetHash !== currentHash) {
          return {
            content: [{ type: 'text', text: 'Contact has changed since prepare was called. Please call prepare_delete_contact again.' }],
            isError: true,
          };
        }

        await repository.deleteContactAsync(params.contact_id);
        return {
          content: [{ type: 'text', text: `Successfully deleted contact ${params.contact_id}` }],
        };
      }

      // Task tools
      case 'list_task_lists': {
        const lists = await repository.listTaskListsAsync();
        return { content: [{ type: 'text', text: JSON.stringify({ task_lists: lists }, null, 2) }] };
      }

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

      case 'create_task': {
        const params = CreateTaskGraphInput.parse(args);
        const numericId = await repository.createTaskAsync({
          title: params.title,
          task_list_id: params.task_list_id,
          ...(params.body != null ? { body: params.body } : {}),
          ...(params.body_type != null ? { body_type: params.body_type } : {}),
          ...(params.due_date != null ? { due_date: params.due_date } : {}),
          ...(params.importance != null ? { importance: params.importance } : {}),
          ...(params.reminder_date != null ? { reminder_date: params.reminder_date } : {}),
          ...(params.recurrence != null ? { recurrence: params.recurrence } : {}),
          ...(params.categories != null ? { categories: params.categories } : {}),
        });
        const result = {
          id: numericId,
          title: params.title,
          task_list_id: params.task_list_id,
          status: 'created',
        };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'update_task': {
        const params = UpdateTaskGraphInput.parse(args);
        const updates: Record<string, unknown> = {};
        if (params.title != null) updates.title = params.title;
        if (params.body != null) {
          updates.body = {
            contentType: params.body_type ?? 'text',
            content: params.body,
          };
        }
        if (params.due_date != null) {
          updates.dueDateTime = {
            dateTime: params.due_date,
            timeZone: 'UTC',
          };
        }
        if (params.importance != null) updates.importance = params.importance;
        if (params.reminder_date != null) {
          updates.isReminderOn = true;
          updates.reminderDateTime = {
            dateTime: params.reminder_date,
            timeZone: 'UTC',
          };
        }
        if (params.status != null) updates.status = params.status;
        if (params.recurrence != null) {
          updates.recurrence = {
            pattern: {
              type: params.recurrence.pattern,
              interval: params.recurrence.interval ?? 1,
              ...(params.recurrence.days_of_week != null ? { daysOfWeek: params.recurrence.days_of_week } : {}),
              ...(params.recurrence.day_of_month != null ? { dayOfMonth: params.recurrence.day_of_month } : {}),
            },
            range: {
              type: params.recurrence.range_type,
              startDate: params.recurrence.start_date,
              ...(params.recurrence.end_date != null ? { endDate: params.recurrence.end_date } : {}),
              ...(params.recurrence.occurrences != null ? { numberOfOccurrences: params.recurrence.occurrences } : {}),
            },
          };
        }
        if (params.categories != null) updates.categories = params.categories;
        await repository.updateTaskAsync(params.task_id, updates);
        return {
          content: [{ type: 'text', text: `Successfully updated task ${params.task_id}` }],
        };
      }

      case 'complete_task': {
        const params = CompleteTaskGraphInput.parse(args);
        await repository.completeTaskAsync(params.task_id);
        return {
          content: [{ type: 'text', text: `Successfully completed task ${params.task_id}` }],
        };
      }

      case 'create_task_list': {
        const params = CreateTaskListGraphInput.parse(args);
        const numericId = await repository.createTaskListAsync(params.display_name);
        const result = {
          id: numericId,
          display_name: params.display_name,
          status: 'created',
        };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'rename_task_list': {
        const params = RenameTaskListInput.parse(args);
        await repository.renameTaskListAsync(params.task_list_id, params.name);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Task list renamed' }, null, 2) }] };
      }

      case 'prepare_delete_task_list': {
        const params = PrepareDeleteTaskListInput.parse(args);
        const token = tokenManager.generateToken({
          operation: 'delete_task_list',
          targetType: 'task_list',
          targetId: params.task_list_id,
          targetHash: String(params.task_list_id),
        });

        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify({
              token_id: token.tokenId,
              expires_at: new Date(token.expiresAt).toISOString(),
              task_list_id: params.task_list_id,
              action: `To confirm deleting task list ${params.task_list_id}, call confirm_delete_task_list with the token_id and task_list_id.`,
            }, null, 2),
          }],
        };
      }

      case 'confirm_delete_task_list': {
        const params = ConfirmDeleteTaskListInput.parse(args);
        const validation = tokenManager.consumeToken(params.token_id, 'delete_task_list', params.task_list_id);
        if (!validation.valid) {
          const errorMessages: Record<string, string> = {
            NOT_FOUND: 'Token not found or already used',
            EXPIRED: 'Token has expired. Please call prepare_delete_task_list again.',
            OPERATION_MISMATCH: 'Token was not generated for delete_task_list',
            TARGET_MISMATCH: 'Token was generated for a different task list',
            ALREADY_CONSUMED: 'Token has already been used',
          };
          return {
            content: [{
              type: 'text' as const,
              text: JSON.stringify({
                success: false,
                error: errorMessages[validation.error ?? ''] ?? 'Invalid token',
              }, null, 2),
            }],
          };
        }

        await repository.deleteTaskListAsync(params.task_list_id);
        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify({ success: true, message: 'Task list deleted' }, null, 2),
          }],
        };
      }

      case 'prepare_delete_task': {
        const params = PrepareDeleteTaskInput.parse(args);
        const task = await repository.getTaskAsync(params.task_id);
        if (task == null) {
          return { content: [{ type: 'text', text: 'Task not found' }], isError: true };
        }

        const taskInfo = repository.getTaskInfo(params.task_id);
        const graphTask = taskInfo != null
          ? await repository.getClient().getTask(taskInfo.taskListId, taskInfo.taskId)
          : null;
        const hash = hashTaskForApproval({
          taskId: taskInfo?.taskId ?? '',
          title: graphTask?.title ?? null,
          listId: taskInfo?.taskListId ?? '',
        });

        const token = tokenManager.generateToken({
          operation: 'delete_task',
          targetType: 'task',
          targetId: params.task_id,
          targetHash: hash,
        });

        const result = {
          token_id: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          task: transformTaskRow(task),
          action: 'This task will be permanently deleted.',
        };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'confirm_delete_task': {
        const params = ConfirmDeleteTaskInput.parse(args);

        // Re-fetch the task and compute fresh hash for comparison
        const taskInfo = repository.getTaskInfo(params.task_id);
        const graphTask = taskInfo != null
          ? await repository.getClient().getTask(taskInfo.taskListId, taskInfo.taskId)
          : null;
        const currentHash = hashTaskForApproval({
          taskId: taskInfo?.taskId ?? '',
          title: graphTask?.title ?? null,
          listId: taskInfo?.taskListId ?? '',
        });

        const validation = tokenManager.consumeToken(params.token_id, 'delete_task', params.task_id);
        if (!validation.valid) {
          const errorMessages: Record<string, string> = {
            NOT_FOUND: 'Token not found or already used',
            EXPIRED: 'Token has expired. Please call prepare_delete_task again.',
            OPERATION_MISMATCH: 'Token was not generated for delete_task',
            TARGET_MISMATCH: 'Token was generated for a different task',
            ALREADY_CONSUMED: 'Token has already been used',
          };
          return {
            content: [{ type: 'text', text: errorMessages[validation.error ?? ''] ?? 'Invalid token' }],
            isError: true,
          };
        }

        // Check that the task hasn't changed since prepare
        if (validation.token!.targetHash !== currentHash) {
          return {
            content: [{ type: 'text', text: 'Task has changed since prepare was called. Please call prepare_delete_task again.' }],
            isError: true,
          };
        }

        await repository.deleteTaskAsync(params.task_id);
        return {
          content: [{ type: 'text', text: `Successfully deleted task ${params.task_id}` }],
        };
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

      // Mail rules tools
      case 'list_mail_rules':
        return await rulesTools.listMailRules();

      case 'create_mail_rule': {
        const params = CreateMailRuleInput.parse(args);
        return await rulesTools.createMailRule(params);
      }

      case 'prepare_delete_mail_rule': {
        const params = PrepareDeleteMailRuleInput.parse(args);
        return rulesTools.prepareDeleteMailRule(params);
      }

      case 'confirm_delete_mail_rule': {
        const params = ConfirmDeleteMailRuleInput.parse(args);
        return await rulesTools.confirmDeleteMailRule(params);
      }

      // Master categories tools
      case 'list_categories':
        return await categoriesTools.listCategories();

      case 'create_category': {
        const params = CreateCategoryInput.parse(args);
        return await categoriesTools.createCategory(params);
      }

      case 'prepare_delete_category': {
        const params = PrepareDeleteCategoryInput.parse(args);
        return categoriesTools.prepareDeleteCategory(params);
      }

      case 'confirm_delete_category': {
        const params = ConfirmDeleteCategoryInput.parse(args);
        return await categoriesTools.confirmDeleteCategory(params);
      }

      // Focused inbox override tools
      case 'list_focused_overrides':
        return await focusedOverridesTools.listFocusedOverrides();

      case 'create_focused_override': {
        const params = CreateFocusedOverrideInput.parse(args);
        return await focusedOverridesTools.createFocusedOverride(params);
      }

      case 'prepare_delete_focused_override': {
        const params = PrepareDeleteFocusedOverrideInput.parse(args);
        return focusedOverridesTools.prepareDeleteFocusedOverride(params);
      }

      case 'confirm_delete_focused_override': {
        const params = ConfirmDeleteFocusedOverrideInput.parse(args);
        return await focusedOverridesTools.confirmDeleteFocusedOverride(params);
      }

      // Automatic replies (OOF) tools
      case 'get_automatic_replies': {
        GetAutomaticRepliesInput.parse(args ?? {});
        const result = await repository.getAutomaticRepliesAsync();
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'set_automatic_replies': {
        const params = SetAutomaticRepliesInput.parse(args);
        const replyParams: Parameters<typeof repository.setAutomaticRepliesAsync>[0] = {
          status: params.status,
        };
        if (params.external_audience != null) replyParams.externalAudience = params.external_audience;
        if (params.internal_reply_message != null) replyParams.internalReplyMessage = params.internal_reply_message;
        if (params.external_reply_message != null) replyParams.externalReplyMessage = params.external_reply_message;
        if (params.scheduled_start != null) replyParams.scheduledStartDateTime = params.scheduled_start;
        if (params.scheduled_end != null) replyParams.scheduledEndDateTime = params.scheduled_end;
        await repository.setAutomaticRepliesAsync(replyParams);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, status: params.status }, null, 2) }] };
      }

      // Mailbox settings tools
      case 'get_mailbox_settings': {
        GetMailboxSettingsInput.parse(args ?? {});
        const result = await repository.getMailboxSettingsAsync();
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'update_mailbox_settings': {
        const params = UpdateMailboxSettingsInput.parse(args);
        const settingsParams: Parameters<typeof repository.updateMailboxSettingsAsync>[0] = {};
        if (params.language != null) settingsParams.language = params.language;
        if (params.time_zone != null) settingsParams.timeZone = params.time_zone;
        if (params.date_format != null) settingsParams.dateFormat = params.date_format;
        if (params.time_format != null) settingsParams.timeFormat = params.time_format;
        await repository.updateMailboxSettingsAsync(settingsParams);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true }, null, 2) }] };
      }

      // Contact folder tools
      case 'list_contact_folders': {
        const folders = await repository.listContactFoldersAsync();
        return { content: [{ type: 'text', text: JSON.stringify({ contact_folders: folders }, null, 2) }] };
      }

      case 'create_contact_folder': {
        const params = CreateContactFolderInput.parse(args);
        const folderId = await repository.createContactFolderAsync(params.name);
        const result = {
          id: folderId,
          name: params.name,
          status: 'created',
        };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
      }

      case 'prepare_delete_contact_folder': {
        const params = PrepareDeleteContactFolderInput.parse(args);
        const token = tokenManager.generateToken({
          operation: 'delete_contact_folder',
          targetType: 'contact_folder',
          targetId: params.folder_id,
          targetHash: String(params.folder_id),
        });

        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify({
              token_id: token.tokenId,
              expires_at: new Date(token.expiresAt).toISOString(),
              folder_id: params.folder_id,
              action: `To confirm deleting contact folder ${params.folder_id}, call confirm_delete_contact_folder with the token_id and folder_id.`,
            }, null, 2),
          }],
        };
      }

      case 'confirm_delete_contact_folder': {
        const params = ConfirmDeleteContactFolderInput.parse(args);
        const validation = tokenManager.consumeToken(params.token_id, 'delete_contact_folder', params.folder_id);
        if (!validation.valid) {
          const errorMessages: Record<string, string> = {
            NOT_FOUND: 'Token not found or already used',
            EXPIRED: 'Token has expired. Please call prepare_delete_contact_folder again.',
            OPERATION_MISMATCH: 'Token was not generated for delete_contact_folder',
            TARGET_MISMATCH: 'Token was generated for a different contact folder',
            ALREADY_CONSUMED: 'Token has already been used',
          };
          return {
            content: [{
              type: 'text' as const,
              text: JSON.stringify({
                success: false,
                error: errorMessages[validation.error ?? ''] ?? 'Invalid token',
              }, null, 2),
            }],
          };
        }

        await repository.deleteContactFolderAsync(params.folder_id);
        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify({ success: true, message: 'Contact folder deleted' }, null, 2),
          }],
        };
      }

      case 'get_contact_photo': {
        const params = GetContactPhotoInput.parse(args);
        const result = await repository.getContactPhotoAsync(params.contact_id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, file_path: result.filePath, content_type: result.contentType }, null, 2) }] };
      }

      case 'set_contact_photo': {
        const params = SetContactPhotoInput.parse(args);
        await repository.setContactPhotoAsync(params.contact_id, params.file_path);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, message: 'Contact photo updated' }, null, 2) }] };
      }

      case 'get_mail_tips': {
        const params = GetMailTipsInput.parse(args);
        const tips = await repository.getMailTipsAsync(params.email_addresses);
        return { content: [{ type: 'text', text: JSON.stringify({ mail_tips: tips }, null, 2) }] };
      }

      case 'get_message_headers': {
        const params = GetMessageHeadersInput.parse(args);
        const headers = await repository.getMessageHeadersAsync(params.email_id);
        return { content: [{ type: 'text', text: JSON.stringify({ headers }, null, 2) }] };
      }

      case 'get_message_mime': {
        const params = GetMessageMimeInput.parse(args);
        const result = await repository.getMessageMimeAsync(params.email_id);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, file_path: result.filePath }, null, 2) }] };
      }

      // Calendar group tools
      case 'list_calendar_groups': {
        const groups = await repository.listCalendarGroupsAsync();
        return { content: [{ type: 'text', text: JSON.stringify({ calendar_groups: groups }, null, 2) }] };
      }

      case 'create_calendar_group': {
        const params = CreateCalendarGroupInput.parse(args);
        const groupId = await repository.createCalendarGroupAsync(params.name);
        return { content: [{ type: 'text', text: JSON.stringify({ success: true, calendar_group_id: groupId, message: 'Calendar group created' }, null, 2) }] };
      }

      // Calendar permission tools
      case 'list_calendar_permissions': {
        const params = ListCalendarPermissionsInput.parse(args);
        return await calendarPermissionsTools!.listCalendarPermissions(params);
      }

      case 'create_calendar_permission': {
        const params = CreateCalendarPermissionInput.parse(args);
        return await calendarPermissionsTools!.createCalendarPermission(params);
      }

      case 'prepare_delete_calendar_permission': {
        const params = PrepareDeleteCalendarPermissionInput.parse(args);
        return calendarPermissionsTools!.prepareDeleteCalendarPermission(params);
      }

      case 'confirm_delete_calendar_permission': {
        const params = ConfirmDeleteCalendarPermissionInput.parse(args);
        return await calendarPermissionsTools!.confirmDeleteCalendarPermission(params);
      }

      // Room lists & rooms tools
      case 'list_room_lists': {
        const roomLists = await repository.listRoomListsAsync();
        return { content: [{ type: 'text', text: JSON.stringify({ room_lists: roomLists }, null, 2) }] };
      }

      case 'list_rooms': {
        const params = ListRoomsInput.parse(args);
        const rooms = await repository.listRoomsAsync(params.room_list_email);
        return { content: [{ type: 'text', text: JSON.stringify({ rooms }, null, 2) }] };
      }

      // Teams tools
      case 'list_teams': {
        return await teamsTools.listTeams();
      }

      case 'list_channels': {
        const params = ListChannelsInput.parse(args);
        return await teamsTools.listChannels(params);
      }

      case 'get_channel': {
        const params = GetChannelInput.parse(args);
        return await teamsTools.getChannel(params);
      }

      case 'create_channel': {
        const params = CreateChannelInput.parse(args);
        return await teamsTools.createChannel(params);
      }

      case 'update_channel': {
        const params = UpdateChannelInput.parse(args);
        return await teamsTools.updateChannel(params);
      }

      case 'prepare_delete_channel': {
        const params = PrepareDeleteChannelInput.parse(args);
        return teamsTools.prepareDeleteChannel(params);
      }

      case 'confirm_delete_channel': {
        const params = ConfirmDeleteChannelInput.parse(args);
        return await teamsTools.confirmDeleteChannel(params);
      }

      case 'list_team_members': {
        const params = ListTeamMembersInput.parse(args);
        return await teamsTools.listTeamMembers(params);
      }

      case 'list_channel_messages': {
        const params = ListChannelMessagesInput.parse(args);
        return await teamsTools.listChannelMessages(params);
      }

      case 'get_channel_message': {
        const params = GetChannelMessageInput.parse(args);
        return await teamsTools.getChannelMessage(params);
      }

      case 'prepare_send_channel_message': {
        const params = PrepareSendChannelMessageInput.parse(args);
        return teamsTools.prepareSendChannelMessage(params);
      }

      case 'confirm_send_channel_message': {
        const params = ConfirmSendChannelMessageInput.parse(args);
        return await teamsTools.confirmSendChannelMessage(params);
      }

      case 'prepare_reply_channel_message': {
        const params = PrepareReplyChannelMessageInput.parse(args);
        return teamsTools.prepareReplyChannelMessage(params);
      }

      case 'confirm_reply_channel_message': {
        const params = ConfirmReplyChannelMessageInput.parse(args);
        return await teamsTools.confirmReplyChannelMessage(params);
      }

      case 'list_chats': {
        const params = ListChatsInput.parse(args);
        return await teamsTools.listChats(params);
      }

      case 'get_chat': {
        const params = GetChatInput.parse(args);
        return await teamsTools.getChat(params);
      }

      case 'list_chat_messages': {
        const params = ListChatMessagesInput.parse(args);
        return await teamsTools.listChatMessages(params);
      }

      case 'prepare_send_chat_message': {
        const params = PrepareSendChatMessageInput.parse(args);
        return teamsTools.prepareSendChatMessage(params);
      }

      case 'confirm_send_chat_message': {
        const params = ConfirmSendChatMessageInput.parse(args);
        return await teamsTools.confirmSendChatMessage(params);
      }

      case 'list_chat_members': {
        const params = ListChatMembersInput.parse(args);
        return await teamsTools.listChatMembers(params);
      }

      // Checklist Items tools
      case 'list_checklist_items': {
        const params = ListChecklistItemsInput.parse(args);
        return await checklistItemsTools.listChecklistItems(params);
      }

      case 'create_checklist_item': {
        const params = CreateChecklistItemInput.parse(args);
        return await checklistItemsTools.createChecklistItem(params);
      }

      case 'update_checklist_item': {
        const params = UpdateChecklistItemInput.parse(args);
        return await checklistItemsTools.updateChecklistItem(params);
      }

      case 'prepare_delete_checklist_item': {
        const params = PrepareDeleteChecklistItemInput.parse(args);
        return checklistItemsTools.prepareDeleteChecklistItem(params);
      }

      case 'confirm_delete_checklist_item': {
        const params = ConfirmDeleteChecklistItemInput.parse(args);
        return await checklistItemsTools.confirmDeleteChecklistItem(params);
      }

      // Linked Resources tools
      case 'list_linked_resources': {
        const params = ListLinkedResourcesInput.parse(args);
        return await linkedResourcesTools.listLinkedResources(params);
      }

      case 'create_linked_resource': {
        const params = CreateLinkedResourceInput.parse(args);
        return await linkedResourcesTools.createLinkedResource(params);
      }

      case 'prepare_delete_linked_resource': {
        const params = PrepareDeleteLinkedResourceInput.parse(args);
        return linkedResourcesTools.prepareDeleteLinkedResource(params);
      }

      case 'confirm_delete_linked_resource': {
        const params = ConfirmDeleteLinkedResourceInput.parse(args);
        return await linkedResourcesTools.confirmDeleteLinkedResource(params);
      }

      // Task Attachments tools
      case 'list_task_attachments': {
        const params = ListTaskAttachmentsInput.parse(args);
        return await taskAttachmentsTools.listTaskAttachments(params);
      }

      case 'create_task_attachment': {
        const params = CreateTaskAttachmentInput.parse(args);
        return await taskAttachmentsTools.createTaskAttachment(params);
      }

      case 'prepare_delete_task_attachment': {
        const params = PrepareDeleteTaskAttachmentInput.parse(args);
        return taskAttachmentsTools.prepareDeleteTaskAttachment(params);
      }

      case 'confirm_delete_task_attachment': {
        const params = ConfirmDeleteTaskAttachmentInput.parse(args);
        return await taskAttachmentsTools.confirmDeleteTaskAttachment(params);
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
import { unixTimestampToLocalIso } from './graph/mappers/utils.js';

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
  categories: readonly string[];
} {
  return {
    id: row.id,
    folderId: row.folderId,
    subject: row.subject,
    sender: row.sender,
    senderAddress: row.senderAddress,
    preview: row.preview,
    isRead: row.isRead === 1,
    timeReceived: unixTimestampToLocalIso(row.timeReceived),
    timeSent: unixTimestampToLocalIso(row.timeSent),
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

/**
 * Transforms an EventRow from the Graph backend.
 * Uses Unix timestamps (not Apple epoch) and includes subject from EventRow.
 */
function transformGraphEventRow(row: EventRow): {
  id: number;
  folderId: number | null;
  title: string | null;
  startDate: string | null;
  endDate: string | null;
  isRecurring: boolean;
  hasReminder: boolean;
  attendeeCount: number | null;
  onlineMeetingUrl: string | null;
} {
  return {
    id: row.id,
    folderId: row.folderId,
    title: row.subject ?? null,
    startDate: unixTimestampToLocalIso(row.startDate),
    endDate: unixTimestampToLocalIso(row.endDate),
    isRecurring: row.isRecurring === 1,
    hasReminder: row.hasReminder === 1,
    attendeeCount: row.attendeeCount,
    onlineMeetingUrl: row.onlineMeetingUrl ?? null,
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
    dueDate: unixTimestampToLocalIso(row.dueDate),
    startDate: unixTimestampToLocalIso(row.startDate),
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
  // Check for CLI subcommands before starting MCP server
  const cliCommand = parseCliCommand(process.argv.slice(2));
  if (cliCommand != null) {
    const exitCode = await handleAuthCommand(cliCommand.flags);
    process.exit(exitCode);
  }

  const server = createServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

// Run if this is the main module (not imported for testing)
// Check multiple conditions to handle direct execution, symlinks, and npx
const isMainModule =
  import.meta.url === `file://${process.argv[1]}` ||
  process.argv[1]?.endsWith('dist/index.js') === true ||
  process.argv[1]?.includes('mcp-office365-mac') === true ||
  // When run via npx or bin, process.argv[1] might be undefined or a symlink
  process.argv[1] === undefined ||
  import.meta.url.endsWith('/dist/index.js');

if (isMainModule) {
  main().catch((error) => {
    console.error('Fatal error:', error);
    process.exit(1);
  });
}
