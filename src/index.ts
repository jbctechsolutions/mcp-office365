#!/usr/bin/env node
/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */
/**
 * Office 365 MCP Server
 *
 * A Model Context Protocol server that provides full read/write access to
 * Microsoft 365 via Microsoft Graph API or legacy AppleScript.
 *
 * Backend selection:
 * - Graph API is the default (full-featured, cross-platform)
 * - Set USE_APPLESCRIPT=1 to use legacy AppleScript backend (macOS + classic Outlook only)
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
import { createRequire } from 'node:module';
import { ToolRegistry } from './registry/index.js';
import type { ToolContext, SurfaceOptions } from './registry/index.js';
import { allToolDefinitions } from './registry/all-tools.js';
import { parseCliCommand, handleAuthCommand, createAuthMutex } from './cli.js';

const pkg = createRequire(import.meta.url)('../package.json') as { version: string };
import { createMailTools } from './tools/mail.js';
import { GraphMailTools } from './tools/mail-graph.js';
import { AppleMailTools } from './tools/mail-apple.js';
import { createCalendarTools } from './tools/calendar.js';
import { GraphCalendarTools } from './tools/calendar-graph.js';
import { AppleCalendarTools } from './tools/calendar-apple.js';
import { createContactsTools } from './tools/contacts.js';
import { GraphContactsTools, transformContactRow } from './tools/contacts-graph.js';
import { createTasksTools } from './tools/tasks.js';
import { GraphTasksTools, transformTaskRow } from './tools/tasks-graph.js';
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
import { MailRulesTools } from './tools/mail-rules.js';
import { CategoriesTools } from './tools/categories.js';
import { CalendarPermissionsTools } from './tools/calendar-permissions.js';
import { FocusedOverridesTools } from './tools/focused-overrides.js';
import { ChecklistItemsTools } from './tools/checklist-items.js';
import { LinkedResourcesTools } from './tools/linked-resources.js';
import { TaskAttachmentsTools } from './tools/task-attachments.js';
import { TeamsTools } from './tools/teams.js';
import { PeopleTools } from './tools/people.js';
import { MeetingsTools } from './tools/meetings.js';
import { ExcelTools } from './tools/excel.js';
import { OneDriveTools } from './tools/onedrive.js';
import { PlannerTools } from './tools/planner.js';
import { PlannerVisualizationTools } from './tools/planner-visualization.js';
import { SharePointTools } from './tools/sharepoint.js';
import {
  SearchEmailsAdvancedInput,
  ListConversationInput,
  CheckNewEmailsInput,
  PrepareBatchDeleteEmailsInput,
  PrepareBatchMoveEmailsInput,
  ConfirmBatchOperationInput,
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
 * Graph API is the default. Set USE_APPLESCRIPT=1 to use the legacy AppleScript backend.
 * USE_GRAPH_API is still supported for backwards compatibility but is now the default.
 */
function shouldUseGraphApi(): boolean {
  const useAppleScript = process.env['USE_APPLESCRIPT'] === '1' || process.env['USE_APPLESCRIPT'] === 'true';
  if (useAppleScript) {
    return false;
  }
  return true;
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
  // Calendar tools
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
  // Contact tools
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
  // =========================================================================
  // Mailbox Organization — Destructive (Two-Phase Approval)
  // =========================================================================
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
  // Teams tools are served by the tool registry (v3, U2).
  // People API tools are served by the tool registry (v3, U2).
  // Planner tools, Planner Visualization tools, and Online Meetings tools are
  // served by the tool registry (v3, U2).
  // Excel Online tools and OneDrive tools are served by the tool registry
  // (v3, U2).
  // SharePoint Document Libraries are served by the tool registry (v3, U2).
];

// =============================================================================
// Server Creation
// =============================================================================

/**
 * Creates and configures the MCP server.
 */
/** Options controlling the exposed tool surface (preset / read-only filters). */
export interface ServerOptions {
  readonly presets?: SurfaceOptions['presets'];
  readonly readOnly?: boolean;
}

export function createServer(options: ServerOptions = {}): Server {
  const server = new Server(
    {
      name: 'office365-mcp',
      version: pkg.version,
    },
    {
      capabilities: {
        tools: {},
      },
    }
  );

  // Determine which backend to use
  const useGraphApi = shouldUseGraphApi();

  // Surface options resolved once for this server instance.
  const surface: SurfaceOptions = {
    backend: useGraphApi ? 'graph' : 'applescript',
    ...(options.presets != null ? { presets: options.presets } : {}),
    ...(options.readOnly != null ? { readOnly: options.readOnly } : {}),
  };

  // Registry-driven tool surface (v3, U1). Static metadata registers eagerly
  // so ListTools works before the backend is initialized; handlers bind to
  // live instances lazily via ToolContext at call time. Domains not yet
  // migrated fall through to the legacy TOOLS array + dispatch switch below.
  const registry = new ToolRegistry();
  registry.register(allToolDefinitions());
  const registeredNames = new Set(registry.names());

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
  let peopleTools: PeopleTools | null = null;
  let plannerTools: PlannerTools | null = null;
  let plannerVisualizationTools: PlannerVisualizationTools | null = null;
  let meetingsTools: MeetingsTools | null = null;
  let oneDriveTools: OneDriveTools | null = null;
  let sharePointTools: SharePointTools | null = null;
  let excelTools: ExcelTools | null = null;
  let checklistItemsTools: ChecklistItemsTools | null = null;
  let linkedResourcesTools: LinkedResourcesTools | null = null;
  let taskAttachmentsTools: TaskAttachmentsTools | null = null;
  let calendarWriter: ICalendarWriter | null = null;
  let calendarManager: ICalendarManager | null = null;
  let mailSender: IMailSender | null = null;

  // Graph-specific state
  let graphRepository: GraphRepository | null = null;
  let graphContentReaders: GraphContentReaders | null = null;
  let graphContactsTools: GraphContactsTools | null = null;
  let graphTasksTools: GraphTasksTools | null = null;
  let graphCalendarTools: GraphCalendarTools | null = null;
  let appleCalendarTools: AppleCalendarTools | null = null;
  let graphMailTools: GraphMailTools | null = null;
  let appleMailTools: AppleMailTools | null = null;

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
    appleCalendarTools = new AppleCalendarTools(calendarTools, calendarWriter, calendarManager);
    appleMailTools = new AppleMailTools(mailTools, accountRepository);

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
    graphContactsTools = new GraphContactsTools(graphRepository, graphContentReaders);
    graphTasksTools = new GraphTasksTools(graphRepository, graphContentReaders);
    graphCalendarTools = new GraphCalendarTools(graphRepository, graphContentReaders);
    graphMailTools = new GraphMailTools(graphRepository, graphContentReaders);

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
    peopleTools = new PeopleTools(graphRepository.getClient());
    plannerTools = new PlannerTools(graphRepository, tokenManager);
    plannerVisualizationTools = new PlannerVisualizationTools(graphRepository);
    meetingsTools = new MeetingsTools(graphRepository);
    oneDriveTools = new OneDriveTools(graphRepository, tokenManager);
    sharePointTools = new SharePointTools(graphRepository);
    excelTools = new ExcelTools(graphRepository, tokenManager);

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
    'get_mail_tips',
    'get_message_headers',
    'get_message_mime',
    'list_calendar_groups',
    'create_calendar_group',
    'list_room_lists',
    'list_rooms',
  ]);

  /** Builds the runtime context for registry handlers (post-initialization). */
  function buildToolContext(): ToolContext {
    return {
      backend: surface.backend,
      tokenManager,
      graph:
        useGraphApi
        && rulesTools != null
        && categoriesTools != null
        && focusedOverridesTools != null
        && calendarPermissionsTools != null
        && checklistItemsTools != null
        && linkedResourcesTools != null
        && taskAttachmentsTools != null
        && peopleTools != null
        && plannerVisualizationTools != null
        && meetingsTools != null
        && sharePointTools != null
        && teamsTools != null
        && plannerTools != null
        && oneDriveTools != null
        && excelTools != null
        && graphContactsTools != null
        && graphTasksTools != null
        && graphCalendarTools != null
        && graphMailTools != null
        && orgTools != null
          ? {
              rules: rulesTools,
              categories: categoriesTools,
              focusedOverrides: focusedOverridesTools,
              calendarPermissions: calendarPermissionsTools,
              checklistItems: checklistItemsTools,
              linkedResources: linkedResourcesTools,
              taskAttachments: taskAttachmentsTools,
              people: peopleTools,
              plannerVisualization: plannerVisualizationTools,
              meetings: meetingsTools,
              sharePoint: sharePointTools,
              teams: teamsTools,
              planner: plannerTools,
              oneDrive: oneDriveTools,
              excel: excelTools,
              contactsGraph: graphContactsTools,
              tasksGraph: graphTasksTools,
              calendarGraph: graphCalendarTools,
              mailGraph: graphMailTools,
              mailboxOrg: orgTools,
            }
          : null,
      applescript:
        !useGraphApi
        && notesTools != null
        && contactsTools != null
        && tasksTools != null
        && appleCalendarTools != null
        && appleMailTools != null
        && orgTools != null
          ? { notes: notesTools, contacts: contactsTools, tasks: tasksTools, calendar: appleCalendarTools, mail: appleMailTools, mailboxOrg: orgTools }
          : null,
    };
  }

  // Register tool list handler: registry tools first, then legacy TOOLS not
  // yet migrated (with the graph-only filter still applied in AppleScript mode).
  server.setRequestHandler(ListToolsRequestSchema, () => {
    const registryTools = registry.listTools(surface);
    const legacyTools = TOOLS.filter((t) => !registeredNames.has(t.name)).filter(
      (t) => useGraphApi || !GRAPH_ONLY_TOOL_NAMES.has(t.name),
    );
    return { tools: [...registryTools, ...legacyTools] };
  });

  // Register tool call handler (async for Graph API support)
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    try {
      await ensureInitialized();

      // Registry dispatch (v3): returns undefined for names not yet migrated,
      // which fall through to the legacy dispatch below.
      const registryResult = await registry.dispatch(name, args, buildToolContext(), surface);
      if (registryResult !== undefined) {
        return registryResult;
      }

      // Graph API mode - handle async operations directly
      if (useGraphApi && graphRepository != null) {
        return await handleGraphToolCall(name, args, graphRepository, graphContentReaders!, orgTools!, sendTools!, schedulingTools!, rulesTools!, categoriesTools!, calendarPermissionsTools!, focusedOverridesTools!, teamsTools!, checklistItemsTools!, linkedResourcesTools!, taskAttachmentsTools!, peopleTools!, plannerTools!, plannerVisualizationTools!, meetingsTools!, oneDriveTools!, sharePointTools!, excelTools!, tokenManager);
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
// Shared Mailbox Organization Handler
// =============================================================================

type ToolResult = { content: Array<{ type: string; text: string }>; isError?: boolean };

async function handleOrgToolCall(
  name: string,
  args: unknown,
  orgTools: ReturnType<typeof createMailboxOrganizationTools>
): Promise<ToolResult | null> {
  switch (name) {
    // Batch operations remain in the legacy dispatch: their prepare_ halves
    // pair with confirm_batch_operation (not a 1:1 confirm_batch_* tool), which
    // the registry's prepare/confirm invariant does not permit. All other
    // mailbox-organization tools are served by the registry.
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

    // Contact tools
    // Note tools
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

const PrepareDeleteEventInput = z.strictObject({
  event_id: z.number().int().positive(),
});

const ConfirmDeleteEventInput = z.strictObject({
  token_id: z.uuid(),
  event_id: z.number().int().positive(),
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
  peopleTools: PeopleTools,
  plannerTools: PlannerTools,
  plannerVisualizationTools: PlannerVisualizationTools,
  meetingsTools: MeetingsTools,
  oneDriveTools: OneDriveTools,
  sharePointTools: SharePointTools,
  excelTools: ExcelTools,
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

      case 'list_conversation': {
        const params = ListConversationInput.parse(args);
        const emails = await repository.listConversationAsync(params.message_id, params.limit);
        const result = { emails: emails.map(transformEmailRow) };
        return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
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
      // Mail rules, master categories, and focused inbox overrides are served
      // by the tool registry (v3, U2).

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

      // Calendar permissions are served by the tool registry (v3, U2).

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

      // Teams, checklist items, linked resources, task attachments, people,
      // planner, planner visualization, online meetings, Excel, OneDrive, and
      // SharePoint are served by the tool registry (v3, U2).

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

import type { EmailRow, EventRow } from './database/repository.js';
import { unixTimestampToLocalIso } from './graph/mappers/utils.js';

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
  process.argv[1]?.includes('mcp-office365') === true ||
  // When run via npx or bin, process.argv[1] might be undefined or a symlink
  process.argv[1] === undefined ||
  import.meta.url.endsWith('/dist/index.js');

if (isMainModule) {
  main().catch((error) => {
    console.error('Fatal error:', error);
    process.exit(1);
  });
}
