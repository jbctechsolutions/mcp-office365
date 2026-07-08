/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Task-related MCP tools.
 *
 * Provides tools for listing, searching, and getting tasks.
 */

import { z } from 'zod';
import type { IRepository, TaskRow } from '../database/repository.js';
import type { TaskSummary, Task, PriorityValue } from '../types/index.js';
import { appleTimestampToIso } from '../utils/dates.js';
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset, requireAppleScriptToolset } from '../registry/context.js';
import type { ToolContext, ToolDefinition, ToolResult } from '../registry/types.js';
import type { GraphTasksTools } from './tasks-graph.js';

// Tasks are a dual-backend domain: the AppleScript backend serves them via
// TasksTools; the Graph backend serves them via GraphTasksTools.
declare module '../registry/types.js' {
  interface GraphToolsets {
    tasksGraph: GraphTasksTools;
  }
  interface AppleScriptToolsets {
    tasks: TasksTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const ListTasksInput = z.strictObject({
  limit: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(50)
    .describe('Maximum number of tasks to return (1-100)'),
  offset: z.number().int().min(0).default(0).describe('Number of tasks to skip'),
  include_completed: z.boolean().default(true).describe('Include completed tasks'),
});

export const SearchTasksInput = z.strictObject({
  query: z.string().min(1).describe('Search query for task names'),
  limit: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(50)
    .describe('Maximum number of tasks to return (1-100)'),
});

export const GetTaskInput = z.strictObject({
  task_id: z.number().int().positive().describe('The task ID to retrieve'),
});

// Task write schemas (Graph API)
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

export const CreateTaskInput = z.strictObject({
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

export const UpdateTaskInput = z.strictObject({
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

export const CompleteTaskInput = z.strictObject({
  task_id: z.number().int().positive(),
});

export const PrepareDeleteTaskInput = z.strictObject({
  task_id: z.number().int().positive(),
});

export const ConfirmDeleteTaskInput = z.strictObject({
  token_id: z.uuid(),
  task_id: z.number().int().positive(),
});

// =============================================================================
// Type Definitions
// =============================================================================

export type ListTasksParams = z.infer<typeof ListTasksInput>;
export type SearchTasksParams = z.infer<typeof SearchTasksInput>;
export type GetTaskParams = z.infer<typeof GetTaskInput>;
export type CreateTaskParams = z.infer<typeof CreateTaskInput>;
export type UpdateTaskParams = z.infer<typeof UpdateTaskInput>;
export type CompleteTaskParams = z.infer<typeof CompleteTaskInput>;
export type PrepareDeleteTaskParams = z.infer<typeof PrepareDeleteTaskInput>;
export type ConfirmDeleteTaskParams = z.infer<typeof ConfirmDeleteTaskInput>;

// =============================================================================
// Content Reader Interface
// =============================================================================

/**
 * Interface for reading task content from data files.
 */
export interface ITaskContentReader {
  /**
   * Reads task details from the given data file path.
   */
  readTaskDetails(dataFilePath: string | null): TaskDetails | null;
}

/**
 * Task details from content file.
 */
export interface TaskDetails {
  readonly body: string | null;
  readonly completedDate: string | null;
  readonly reminderDate: string | null;
  readonly categories: readonly string[];
}

/**
 * Default task content reader that returns null.
 */
export const nullTaskContentReader: ITaskContentReader = {
  readTaskDetails: (): TaskDetails | null => null,
};

// =============================================================================
// Transformers
// =============================================================================

/**
 * Transforms a database task row to TaskSummary.
 */
function transformTaskSummary(row: TaskRow): TaskSummary {
  return {
    id: row.id,
    folderId: row.folderId,
    name: row.name,
    isCompleted: row.isCompleted === 1,
    dueDate: appleTimestampToIso(row.dueDate),
    priority: row.priority as PriorityValue,
  };
}

/**
 * Transforms a database task row to full Task.
 */
function transformTask(row: TaskRow, details: TaskDetails | null): Task {
  const summary = transformTaskSummary(row);

  return {
    ...summary,
    startDate: appleTimestampToIso(row.startDate),
    completedDate: details?.completedDate ?? null,
    hasReminder: row.hasReminder === 1,
    reminderDate: details?.reminderDate ?? null,
    body: details?.body ?? null,
    categories: details?.categories ?? [],
  };
}

// =============================================================================
// Tasks Tools Class
// =============================================================================

/**
 * Tasks tools implementation with dependency injection.
 */
export class TasksTools {
  constructor(
    private readonly repository: IRepository,
    private readonly contentReader: ITaskContentReader = nullTaskContentReader
  ) {}

  /**
   * Lists tasks with pagination and filtering.
   */
  listTasks(params: ListTasksParams): TaskSummary[] {
    const { limit, offset, include_completed } = params;

    const rows = include_completed
      ? this.repository.listTasks(limit, offset)
      : this.repository.listIncompleteTasks(limit, offset);

    return rows.map(transformTaskSummary);
  }

  /**
   * Searches tasks by name.
   */
  searchTasks(params: SearchTasksParams): TaskSummary[] {
    const { query, limit } = params;
    const rows = this.repository.searchTasks(query, limit);
    return rows.map(transformTaskSummary);
  }

  /**
   * Gets a single task by ID.
   */
  getTask(params: GetTaskParams): Task | null {
    const { task_id } = params;

    const row = this.repository.getTask(task_id);
    if (row == null) {
      return null;
    }

    const details = this.contentReader.readTaskDetails(row.dataFilePath);
    return transformTask(row, details);
  }
}

/**
 * Creates tasks tools with the given repository.
 */
export function createTasksTools(
  repository: IRepository,
  contentReader: ITaskContentReader = nullTaskContentReader
): TasksTools {
  return new TasksTools(repository, contentReader);
}

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2 — dual backend)
// =============================================================================

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Registry tool definitions for the tasks domain. Each handler branches on the
 * active backend: Graph delegates to GraphTasksTools (which returns MCP content
 * directly); AppleScript delegates to TasksTools (which returns raw objects,
 * wrapped here to match the pre-registry dispatch behavior exactly).
 */
export function tasksToolDefinitions(): ToolDefinition[] {
  return [
    defineTool({
      name: 'list_tasks',
      description: 'List tasks with pagination and filtering',
      input: ListTasksInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['tasks'],
      backends: ['graph', 'applescript'],
      handler: (ctx, params) =>
        ctx.backend === 'graph'
          ? requireGraphToolset(ctx, 'tasksGraph').listTasks(params)
          : jsonResult(requireAppleScriptToolset(ctx, 'tasks').listTasks(params)),
    }),
    defineTool({
      name: 'search_tasks',
      description: 'Search tasks by name',
      input: SearchTasksInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['tasks'],
      backends: ['graph', 'applescript'],
      handler: (ctx, params) =>
        ctx.backend === 'graph'
          ? requireGraphToolset(ctx, 'tasksGraph').searchTasks(params)
          : jsonResult(requireAppleScriptToolset(ctx, 'tasks').searchTasks(params)),
    }),
    defineTool({
      name: 'get_task',
      description: 'Get task details',
      input: GetTaskInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['tasks'],
      backends: ['graph', 'applescript'],
      handler: (ctx: ToolContext, params): Promise<ToolResult> | ToolResult => {
        if (ctx.backend === 'graph') {
          return requireGraphToolset(ctx, 'tasksGraph').getTask(params);
        }
        const result = requireAppleScriptToolset(ctx, 'tasks').getTask(params);
        if (result == null) {
          return { content: [{ type: 'text', text: 'Task not found' }], isError: true };
        }
        return jsonResult(result);
      },
    }),
    defineTool({
      name: 'create_task',
      description: 'Create a new task in a task list. Supports optional recurrence settings for repeating tasks.',
      input: CreateTaskInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'tasksGraph').createTask(params),
    }),
    defineTool({
      name: 'update_task',
      description: 'Update an existing task. Only specified fields will be updated. Supports optional recurrence settings for repeating tasks.',
      input: UpdateTaskInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'tasksGraph').updateTask(params),
    }),
    defineTool({
      name: 'complete_task',
      description: 'Mark a task as completed',
      input: CompleteTaskInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'tasksGraph').completeTask(params),
    }),
    defineTool({
      name: 'prepare_delete_task',
      description: 'Prepare to delete a task. Returns a preview and approval token. Call confirm_delete_task to execute.',
      input: PrepareDeleteTaskInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'tasksGraph').prepareDeleteTask(params),
    }),
    defineTool({
      name: 'confirm_delete_task',
      description: 'Confirm deletion of a task using a token from prepare_delete_task',
      input: ConfirmDeleteTaskInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'tasksGraph').confirmDeleteTask(params),
    }),
  ];
}
