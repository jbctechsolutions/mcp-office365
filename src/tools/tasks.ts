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

// =============================================================================
// Type Definitions
// =============================================================================

export type ListTasksParams = z.infer<typeof ListTasksInput>;
export type SearchTasksParams = z.infer<typeof SearchTasksInput>;
export type GetTaskParams = z.infer<typeof GetTaskInput>;

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
  ];
}
