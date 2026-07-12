#!/usr/bin/env node
/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */
/**
 * Office 365 MCP Server
 *
 * A Model Context Protocol server that provides full read/write access to
 * Microsoft 365 via the Microsoft Graph API.
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type CallToolResult,
} from '@modelcontextprotocol/sdk/types.js';

import {
  createGraphRepository,
  createGraphContentReadersWithClient,
  isAuthenticated,
  getAccessToken,
  resolveAccountId,
  currentAccountId,
  DEFAULT_ACCOUNT_ID,
  GraphMailboxAdapter,
  type GraphRepository,
  type GraphContentReaders,
} from './graph/index.js';
import { createRequire } from 'node:module';
import { ToolRegistry } from './registry/index.js';
import type { ToolContext, SurfaceOptions, ConfirmMode, Elicitor } from './registry/index.js';
import { createServerElicitor } from './registry/elicitor.js';
import { allToolDefinitions } from './registry/all-tools.js';
import {
  parseCliCommand,
  parseServerOptions,
  parseServeOptions,
  handleAuthCommand,
  createAuthMutex,
} from './cli.js';

// Prefer the build-time stamp (sits next to the compiled index.js) so a stale
// dist reports the version it was built from; package.json is the dev/test
// fallback where no stamp exists.
const requireFromHere = createRequire(import.meta.url);
function loadVersion(): string {
  try {
    const stamped = (requireFromHere('./build-info.json') as { version?: unknown }).version;
    if (typeof stamped === 'string' && stamped !== '') return stamped;
  } catch {
    /* no stamp — running from src (dev/tests) */
  }
  return (requireFromHere('../package.json') as { version: string }).version;
}
const pkg = { version: loadVersion() };
import { GraphMailTools } from './tools/mail-graph.js';
import { GraphCalendarTools } from './tools/calendar-graph.js';
import { GraphContactsTools } from './tools/contacts-graph.js';
import { GraphContactFoldersTools } from './tools/contact-folders.js';
import { GraphTasksTools } from './tools/tasks-graph.js';
import { GraphTaskListsTools } from './tools/task-lists.js';
import { GraphMailboxSettingsTools } from './tools/mailbox-settings.js';
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
import { SharedMailboxTools } from './tools/shared-mailbox.js';
import { MeetingsTools } from './tools/meetings.js';
import { OneNoteTools } from './tools/onenote.js';
import { ExcelTools } from './tools/excel.js';
import { OneDriveTools } from './tools/onedrive.js';
import { PlannerTools } from './tools/planner.js';
import { PlannerVisualizationTools } from './tools/planner-visualization.js';
import { SharePointTools } from './tools/sharepoint.js';
import { DeltaTools } from './tools/what-changed.js';
import { SharePointListsTools } from './tools/sharepoint-lists.js';
import { ApprovalTokenManager } from './approval/index.js';
import { StateStore } from './state/store.js';
import {
  toErrorEnvelope,
  ensureErrorEnvelopeText,
  ErrorCode,
  GraphAuthRequiredError,
  type ErrorEnvelope,
} from './utils/errors.js';

// =============================================================================
// Server Creation
// =============================================================================

/**
 * D10: normalize a tool result whose handler returned an error directly (rather
 * than throwing) so its text carries the stable envelope shape. Success results
 * and results already carrying an envelope pass through unchanged.
 */
function normalizeToolResult(result: CallToolResult): CallToolResult {
  if (result.isError !== true) {
    return result;
  }
  const text = (result.content ?? [])
    .filter((block): block is { type: 'text'; text: string } => block.type === 'text')
    .map((block) => block.text)
    .join('\n');
  const normalized = ensureErrorEnvelopeText(text);
  if (normalized === text) {
    return result;
  }
  return { ...result, content: [{ type: 'text', text: normalized }] };
}

/**
 * Creates and configures the MCP server.
 */
/** Options controlling the exposed tool surface (preset / read-only filters). */
export interface ServerOptions {
  readonly presets?: SurfaceOptions['presets'];
  readonly readOnly?: boolean;
  /** Destructive-confirm mode (U11): 'token' (default) or 'elicit'. */
  readonly confirmMode?: ConfirmMode;
  /**
   * Durable state store to back approval tokens and durable-ID aliases. When
   * omitted (the stdio path), each server opens its own via `StateStore.open()`.
   * Remote mode (U3) injects one process-scoped store so a per-request server is
   * cheap to build and does not re-open SQLite / re-run migrations per request.
   */
  readonly stateStore?: StateStore;
  /**
   * Whether an unauthenticated first tool call may trigger the interactive
   * device-code flow. True (default) for stdio at a terminal. Remote/serve mode
   * sets it false: there is no device-code channel over HTTP, so an unauthed
   * call must fail fast with a typed error rather than hang until MSAL times out
   * (and, with per-request servers, spawn concurrent device-code flows).
   */
  readonly interactiveAuth?: boolean;
}

/**
 * Derives the server options for remote/serve mode from the parsed CLI options.
 * Remote mode has no elicitation channel (force `token` confirm) and no
 * device-code channel (fail fast on an unauthenticated call rather than hang) —
 * these overrides are load-bearing and must hold regardless of user flags.
 */
export function serveServerOptions(base: ServerOptions): ServerOptions {
  return { ...base, confirmMode: 'token', interactiveAuth: false };
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

  // Surface options resolved once for this server instance. Graph is the only
  // backend.
  const surface: SurfaceOptions = {
    backend: 'graph',
    ...(options.presets != null ? { presets: options.presets } : {}),
    ...(options.readOnly != null ? { readOnly: options.readOnly } : {}),
  };

  // Confirmation mode (U11). In 'elicit' mode a destructive prepare asks the
  // user inline (capability-gated, 60s) and degrades to the token flow; 'token'
  // (default) is the plain two-phase behavior and needs no elicitor.
  const confirmMode: ConfirmMode = options.confirmMode ?? 'token';
  const elicit: Elicitor | undefined =
    confirmMode === 'elicit' ? createServerElicitor(server) : undefined;

  // Registry-driven tool surface (v3). The registry is the single source of
  // truth for every tool: static metadata registers eagerly so ListTools works
  // before the backend is initialized, and handlers bind to live instances
  // lazily via ToolContext at call time.
  const registry = new ToolRegistry();
  registry.register(allToolDefinitions());

  // The durable state store backs approval tokens (U9b) so a two-phase
  // approval survives a restart / a second window; a corrupt/locked db
  // degrades to in-memory (StateStore.open handles it). Remote mode injects a
  // shared, process-scoped store (U3); the stdio path opens its own here.
  const stateStore = options.stateStore ?? StateStore.open();
  // accountId is a thunk: the signed-in account (homeAccountId) is only known
  // after auth, later than this construction. resolveAccountId() populates it in
  // initializeGraphBackend; currentAccountId() reads the memo at each token op.
  const tokenManager = new ApprovalTokenManager({
    store: stateStore,
    accountId: currentAccountId,
  });

  // Tools and backend state
  let initialized = false;
  let orgTools: ReturnType<typeof createMailboxOrganizationTools> | null = null;
  let sendTools: ReturnType<typeof createMailSendTools> | null = null;
  let schedulingTools: ReturnType<typeof createSchedulingTools> | null = null;
  let rulesTools: MailRulesTools | null = null;
  let categoriesTools: CategoriesTools | null = null;
  let calendarPermissionsTools: CalendarPermissionsTools | null = null;
  let focusedOverridesTools: FocusedOverridesTools | null = null;
  let teamsTools: TeamsTools | null = null;
  let peopleTools: PeopleTools | null = null;
  let sharedMailboxTools: SharedMailboxTools | null = null;
  let plannerTools: PlannerTools | null = null;
  let plannerVisualizationTools: PlannerVisualizationTools | null = null;
  let meetingsTools: MeetingsTools | null = null;
  let onenoteTools: OneNoteTools | null = null;
  let oneDriveTools: OneDriveTools | null = null;
  let sharePointTools: SharePointTools | null = null;
  let sharePointListsTools: SharePointListsTools | null = null;
  let excelTools: ExcelTools | null = null;
  let checklistItemsTools: ChecklistItemsTools | null = null;
  let linkedResourcesTools: LinkedResourcesTools | null = null;
  let taskAttachmentsTools: TaskAttachmentsTools | null = null;
  let deltaTools: DeltaTools | null = null;

  // Graph-specific state
  let graphRepository: GraphRepository | null = null;
  let graphContentReaders: GraphContentReaders | null = null;
  let graphContactsTools: GraphContactsTools | null = null;
  let graphContactFoldersTools: GraphContactFoldersTools | null = null;
  let graphTasksTools: GraphTasksTools | null = null;
  let graphTaskListsTools: GraphTaskListsTools | null = null;
  let graphCalendarTools: GraphCalendarTools | null = null;
  let graphMailTools: GraphMailTools | null = null;
  let graphMailboxSettingsTools: GraphMailboxSettingsTools | null = null;

  /**
   * Initializes the Graph API backend.
   * If not authenticated, triggers the device code flow inline.
   */
  const initializeGraphBackend = createAuthMutex(async (): Promise<void> => {
    // Try to authenticate if needed (triggers device code flow for first-time users)
    const authenticated = await isAuthenticated();
    if (!authenticated) {
      // Remote/serve mode has no device-code channel: fail fast instead of
      // hanging on an interactive prompt no one can answer over HTTP.
      if (options.interactiveAuth === false) {
        throw new GraphAuthRequiredError('not_authenticated');
      }
      await getAccessToken();
    }

    // Capture the signed-in account (homeAccountId) so approval tokens (D8) and
    // the durable-ID alias table (D3) scope to this user, not the 'default'
    // fallback. Best-effort — an unresolved account leaves the fallback in place.
    await resolveAccountId();

    graphRepository = createGraphRepository(undefined, stateStore, currentAccountId);
    graphContentReaders = createGraphContentReadersWithClient(graphRepository.getClient());
    graphContactsTools = new GraphContactsTools(graphRepository, graphContentReaders, tokenManager);
    graphContactFoldersTools = new GraphContactFoldersTools(graphRepository, tokenManager);
    graphTasksTools = new GraphTasksTools(graphRepository, graphContentReaders, tokenManager);
    graphTaskListsTools = new GraphTaskListsTools(graphRepository, tokenManager);
    graphCalendarTools = new GraphCalendarTools(graphRepository, graphContentReaders, tokenManager);
    graphMailTools = new GraphMailTools(graphRepository, graphContentReaders);
    graphMailboxSettingsTools = new GraphMailboxSettingsTools(graphRepository);

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
    sharedMailboxTools = new SharedMailboxTools(graphRepository.getClient());
    plannerTools = new PlannerTools(graphRepository, tokenManager);
    plannerVisualizationTools = new PlannerVisualizationTools(graphRepository);
    meetingsTools = new MeetingsTools(graphRepository);
    onenoteTools = new OneNoteTools(graphRepository);
    oneDriveTools = new OneDriveTools(graphRepository, tokenManager);
    sharePointTools = new SharePointTools(graphRepository, tokenManager);
    sharePointListsTools = new SharePointListsTools(graphRepository, tokenManager);
    excelTools = new ExcelTools(graphRepository, tokenManager);
    deltaTools = new DeltaTools(graphRepository.getClient(), stateStore, currentAccountId);

    initialized = true;
  });

  /**
   * Ensures the backend is initialized.
   */
  async function ensureInitialized(): Promise<void> {
    if (initialized) return;
    await initializeGraphBackend();
  }

  // Tools that only exist when using Graph API but are still served by the
  // legacy TOOLS array. All previously graph-only legacy tools have migrated to
  // the tool registry (which applies its own per-backend filter), so this set
  // is now empty.

  /** Builds the runtime context for registry handlers (post-initialization). */
  function buildToolContext(): ToolContext {
    return {
      backend: surface.backend,
      tokenManager,
      confirmMode,
      ...(elicit != null ? { elicit } : {}),
      graph:
        rulesTools != null
        && categoriesTools != null
        && focusedOverridesTools != null
        && calendarPermissionsTools != null
        && checklistItemsTools != null
        && linkedResourcesTools != null
        && taskAttachmentsTools != null
        && peopleTools != null
        && sharedMailboxTools != null
        && plannerVisualizationTools != null
        && meetingsTools != null
        && onenoteTools != null
        && sharePointTools != null
        && sharePointListsTools != null
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
        && sendTools != null
        && schedulingTools != null
        && orgTools != null
        && deltaTools != null
          ? {
              rules: rulesTools,
              categories: categoriesTools,
              focusedOverrides: focusedOverridesTools,
              calendarPermissions: calendarPermissionsTools,
              checklistItems: checklistItemsTools,
              linkedResources: linkedResourcesTools,
              taskAttachments: taskAttachmentsTools,
              people: peopleTools,
              sharedMailbox: sharedMailboxTools,
              plannerVisualization: plannerVisualizationTools,
              meetings: meetingsTools,
              onenote: onenoteTools,
              sharePoint: sharePointTools,
              sharePointLists: sharePointListsTools,
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
              delta: deltaTools,
            }
          : null,
    };
  }

  // Register tool list handler: registry tools first, then legacy TOOLS not
  // yet migrated.
  server.setRequestHandler(ListToolsRequestSchema, () => {
    return { tools: registry.listTools(surface) };
  });

  // Register tool call handler. Every tool is served by the registry (v3);
  // an unregistered name (or one filtered out of the current surface) yields
  // an unknown-tool error.
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    // An unregistered tool name is a request-validation failure, not a Graph
    // backend failure — emit VALIDATION_ERROR directly rather than routing
    // through toErrorEnvelope (which would label it GRAPH_ERROR). Registered
    // tools filtered out of the current surface (read-only/preset) are NOT
    // handled here — dispatch() throws their mode-specific envelope instead.
    const unknownTool = (): CallToolResult => {
      const envelope: ErrorEnvelope = {
        code: ErrorCode.VALIDATION_ERROR,
        message: `Unknown tool: ${name}`,
        retriable: false,
        suggestion: 'Call ListTools to inspect the available tools for the current backend.',
      };
      return {
        content: [{ type: 'text', text: JSON.stringify(envelope, null, 2) }],
        isError: true,
      };
    };

    // Reject an unknown tool before initializing (and authenticating) the Graph
    // backend — no unknown tool name should trigger a device-code sign-in.
    if (!registry.has(name)) {
      return unknownTool();
    }

    try {
      await ensureInitialized();

      // Self-heal the account identity: initializeGraphBackend resolves it once,
      // but if getAccount() transiently returned null there (it swallows errors)
      // the fallback would otherwise be pinned for the process lifetime — a token
      // minted under 'default' is then NOT_FOUND in a sibling window/restart that
      // resolves the real id. Retry until the real homeAccountId is memoized;
      // once resolved this is a cheap sync no-op.
      if (currentAccountId() === DEFAULT_ACCOUNT_ID) {
        await resolveAccountId();
      }

      const registryResult = await registry.dispatch(name, args, buildToolContext(), surface);
      if (registryResult !== undefined) {
        // D10: handlers that return an error result directly (not-found,
        // approval-token mismatches, …) are normalized to the envelope shape
        // here so every failure path — thrown or returned — has one contract.
        return normalizeToolResult(registryResult as CallToolResult);
      }

      // Defensive: has() gated this path, so dispatch should not return undefined.
      return unknownTool();
    } catch (error) {
      // D10: every thrown failure surfaces as a stable typed envelope mapped at
      // this single point. Guard against a pathological error whose own
      // getters throw so the chokepoint itself can never reject.
      let envelope: ErrorEnvelope;
      try {
        envelope = toErrorEnvelope(error);
      } catch {
        envelope = { code: ErrorCode.GRAPH_ERROR, message: 'An unknown error occurred.', retriable: false };
      }

      return {
        content: [{ type: 'text', text: JSON.stringify(envelope, null, 2) }],
        isError: true,
      } satisfies CallToolResult;
    }
  });

  return server;
}


// =============================================================================
// Main Entry Point
// =============================================================================

async function main(): Promise<void> {
  const argv = process.argv.slice(2);

  // Check for CLI subcommands before starting MCP server
  const cliCommand = parseCliCommand(argv);
  if (cliCommand?.command === 'auth') {
    const exitCode = await handleAuthCommand(cliCommand.flags);
    process.exit(exitCode);
  }

  // Server-mode flags: --preset <names>, --read-only (U10). Under `serve`, these
  // apply as the process-wide outer bound; per-user surfaces layer inside it (U6).
  const serverFlags = cliCommand?.command === 'serve' ? cliCommand.flags : argv;
  let options: ServerOptions;
  try {
    const parsed = parseServerOptions(serverFlags);
    options = {
      readOnly: parsed.readOnly,
      confirmMode: parsed.confirmMode,
      ...(parsed.presets != null ? { presets: parsed.presets } : {}),
    };
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exit(1);
  }

  // `serve`: remote connector mode over stateless Streamable HTTP (U3).
  if (cliCommand?.command === 'serve') {
    try {
      const { host, port } = parseServeOptions(cliCommand.flags);
      const { startHttpServer } = await import('./remote/http-server.js');
      const stateDir = process.env.OUTLOOK_MCP_STATE_DIR;
      const stateStore = StateStore.open(stateDir != null ? { dir: stateDir } : {});
      const httpServer = await startHttpServer({
        host,
        port,
        serverOptions: serveServerOptions(options),
        stateStore,
      });
      // Drain in-flight requests on shutdown (Container Apps sends SIGTERM).
      const shutdown = (): void => {
        httpServer.close(() => process.exit(0));
      };
      process.on('SIGTERM', shutdown);
      process.on('SIGINT', shutdown);
      process.stderr.write(`mcp-office365 serve listening on http://${host}:${port}/mcp\n`);
    } catch (error) {
      process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
      process.exit(1);
    }
    return;
  }

  const server = createServer(options);
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
