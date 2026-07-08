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
import { GraphContactsTools } from './tools/contacts-graph.js';
import { GraphContactFoldersTools } from './tools/contact-folders.js';
import { createTasksTools } from './tools/tasks.js';
import { GraphTasksTools } from './tools/tasks-graph.js';
import { GraphTaskListsTools } from './tools/task-lists.js';
import { GraphMailboxSettingsTools } from './tools/mailbox-settings.js';
import { AccountsTools } from './tools/accounts.js';
import { createNotesTools } from './tools/notes.js';
import { createMailboxOrganizationTools } from './tools/mailbox-organization.js';
import { createMailSendTools } from './tools/mail-send.js';
import { createSchedulingTools } from './tools/scheduling.js';
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
import { ApprovalTokenManager } from './approval/index.js';
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
  let graphContactFoldersTools: GraphContactFoldersTools | null = null;
  let graphTasksTools: GraphTasksTools | null = null;
  let graphTaskListsTools: GraphTaskListsTools | null = null;
  let graphCalendarTools: GraphCalendarTools | null = null;
  let appleCalendarTools: AppleCalendarTools | null = null;
  let graphMailTools: GraphMailTools | null = null;
  let appleMailTools: AppleMailTools | null = null;
  let graphMailboxSettingsTools: GraphMailboxSettingsTools | null = null;
  let accountsTools: AccountsTools | null = null;

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
    accountsTools = new AccountsTools(accountRepository);
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
    graphContactsTools = new GraphContactsTools(graphRepository, graphContentReaders, tokenManager);
    graphContactFoldersTools = new GraphContactFoldersTools(graphRepository, tokenManager);
    graphTasksTools = new GraphTasksTools(graphRepository, graphContentReaders, tokenManager);
    graphTaskListsTools = new GraphTaskListsTools(graphRepository, tokenManager);
    graphCalendarTools = new GraphCalendarTools(graphRepository, graphContentReaders, tokenManager);
    graphMailTools = new GraphMailTools(graphRepository, graphContentReaders);
    graphMailboxSettingsTools = new GraphMailboxSettingsTools(graphRepository);
    accountRepository = createAccountRepository();
    accountsTools = new AccountsTools(accountRepository);

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

  // Tools that only exist when using Graph API but are still served by the
  // legacy TOOLS array. All previously graph-only legacy tools have migrated to
  // the tool registry (which applies its own per-backend filter), so this set
  // is now empty.
  const GRAPH_ONLY_TOOL_NAMES = new Set<string>([]);

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
        && graphContactFoldersTools != null
        && graphTasksTools != null
        && graphTaskListsTools != null
        && graphCalendarTools != null
        && graphMailTools != null
        && graphMailboxSettingsTools != null
        && accountsTools != null
        && sendTools != null
        && schedulingTools != null
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
              contactFolders: graphContactFoldersTools,
              tasksGraph: graphTasksTools,
              taskLists: graphTaskListsTools,
              calendarGraph: graphCalendarTools,
              mailGraph: graphMailTools,
              mailSend: sendTools,
              scheduling: schedulingTools,
              mailboxOrg: orgTools,
              mailboxSettings: graphMailboxSettingsTools,
              accounts: accountsTools,
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
        && accountsTools != null
          ? { notes: notesTools, contacts: contactsTools, tasks: tasksTools, calendar: appleCalendarTools, mail: appleMailTools, mailboxOrg: orgTools, accounts: accountsTools }
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
        return await handleGraphToolCall(name, args, graphRepository, graphContentReaders!, orgTools!, rulesTools!, categoriesTools!, calendarPermissionsTools!, focusedOverridesTools!, teamsTools!, checklistItemsTools!, linkedResourcesTools!, taskAttachmentsTools!, peopleTools!, plannerTools!, plannerVisualizationTools!, meetingsTools!, oneDriveTools!, sharePointTools!, excelTools!, tokenManager);
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

function handleOrgToolCall(
  name: string,
  _args: unknown,
  _orgTools: ReturnType<typeof createMailboxOrganizationTools>
): ToolResult | null {
  // All mailbox-organization tools (including the batch operations) are served
  // by the tool registry (v3, U2). This legacy hook is retained as a no-op
  // fallback and always returns null.
  switch (name) {
    default:
      return null;
  }
}

// =============================================================================
// AppleScript Tool Handler
// =============================================================================

// eslint-disable-next-line @typescript-eslint/require-await
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
  const orgResult = handleOrgToolCall(name, args, orgTools);
  if (orgResult != null) return orgResult;

  switch (name) {
    // Account tools (list_accounts) are served by the tool registry (v3, U2).
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
// Graph API Tool Handler
// =============================================================================

// eslint-disable-next-line @typescript-eslint/require-await
async function handleGraphToolCall(
  name: string,
  args: unknown,
  _repository: GraphRepository,
  _contentReaders: GraphContentReaders,
  orgTools: ReturnType<typeof createMailboxOrganizationTools>,
  _rulesTools: MailRulesTools,
  _categoriesTools: CategoriesTools,
  _calendarPermissionsTools: CalendarPermissionsTools,
  _focusedOverridesTools: FocusedOverridesTools,
  _teamsTools: TeamsTools,
  _checklistItemsTools: ChecklistItemsTools,
  _linkedResourcesTools: LinkedResourcesTools,
  _taskAttachmentsTools: TaskAttachmentsTools,
  _peopleTools: PeopleTools,
  _plannerTools: PlannerTools,
  _plannerVisualizationTools: PlannerVisualizationTools,
  _meetingsTools: MeetingsTools,
  _oneDriveTools: OneDriveTools,
  _sharePointTools: SharePointTools,
  _excelTools: ExcelTools,
  _tokenManager: ApprovalTokenManager
): Promise<ToolResult> {
  // Handle mailbox organization tools (shared between backends)
  const orgResult = handleOrgToolCall(name, args, orgTools);
  if (orgResult != null) return orgResult;

  // All Graph-backend tools are served by the tool registry (v3, U2); this
  // legacy switch retains only the unknown-tool fallback.
  try {
    switch (name) {
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
