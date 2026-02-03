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
  isOutlookRunning,
  type IAccountRepository,
  type ICalendarWriter,
  type RecurrenceConfig,
} from './applescript/index.js';
import {
  createGraphRepository,
  createGraphContentReadersWithClient,
  isAuthenticated,
  type GraphRepository,
  type GraphContentReaders,
} from './graph/index.js';
import { createMailTools } from './tools/mail.js';
import { createCalendarTools } from './tools/calendar.js';
import { createContactsTools } from './tools/contacts.js';
import { createTasksTools } from './tools/tasks.js';
import { createNotesTools } from './tools/notes.js';
import {
  ListFoldersInput,
  ListEmailsInput,
  SearchEmailsInput,
  GetEmailInput,
  GetUnreadCountInput,
  ListCalendarsInput,
  ListEventsInput,
  GetEventInput,
  SearchEventsInput,
  CreateEventInput,
  ListContactsInput,
  SearchContactsInput,
  GetContactInput,
  ListTasksInput,
  SearchTasksInput,
  GetTaskInput,
  ListNotesInput,
  GetNoteInput,
  SearchNotesInput,
} from './tools/index.js';
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

  // Tools and backend state
  let initialized = false;
  let accountRepository: IAccountRepository | null = null;
  let mailTools: ReturnType<typeof createMailTools> | null = null;
  let calendarTools: ReturnType<typeof createCalendarTools> | null = null;
  let contactsTools: ReturnType<typeof createContactsTools> | null = null;
  let tasksTools: ReturnType<typeof createTasksTools> | null = null;
  let notesTools: ReturnType<typeof createNotesTools> | null = null;
  let calendarWriter: ICalendarWriter | null = null;

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
    calendarWriter = createCalendarWriter();

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

    // Note: We can't use the sync tool interfaces with async Graph API
    // So we won't initialize the *Tools objects for Graph mode
    // Instead, we'll handle Graph calls directly in the request handler

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
        return await handleGraphToolCall(name, args, graphRepository, graphContentReaders!);
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
        calendarWriter
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
// AppleScript Tool Handler
// =============================================================================

function handleAppleScriptToolCall(
  name: string,
  args: unknown,
  accountRepository: IAccountRepository,
  mailTools: ReturnType<typeof createMailTools>,
  calendarTools: ReturnType<typeof createCalendarTools>,
  contactsTools: ReturnType<typeof createContactsTools>,
  tasksTools: ReturnType<typeof createTasksTools>,
  notesTools: ReturnType<typeof createNotesTools>,
  calendarWriter: ICalendarWriter | null
): { content: Array<{ type: string; text: string }>; isError?: boolean } {
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
  contentReaders: GraphContentReaders
): Promise<{ content: Array<{ type: string; text: string }>; isError?: boolean }> {
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

function transformFolderRow(row: FolderRow) {
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

function transformEmailRow(row: EmailRow) {
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
  };
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

function transformContactRow(row: ContactRow) {
  return {
    id: row.id,
    displayName: row.displayName,
    sortName: row.sortName,
  };
}

function transformTaskRow(row: TaskRow) {
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
