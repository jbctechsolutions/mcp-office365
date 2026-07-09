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
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolDefinition } from '../registry/types.js';
import type { GraphTasksTools } from './tasks-graph.js';

// Tasks are served by GraphTasksTools.
declare module '../registry/types.js' {
  interface GraphToolsets {
    tasksGraph: GraphTasksTools;
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
  task_id: z.string().min(1).describe('The task ID (td_ token) to retrieve'),
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
  task_list_id: z.string().min(1).describe('The task list ID (tl_ token)'),
  body: z.string().optional(),
  body_type: z.enum(['text', 'html']).optional(),
  due_date: z.string().optional(),
  importance: z.enum(['low', 'normal', 'high']).optional(),
  reminder_date: z.string().optional(),
  recurrence: RecurrenceSchema,
  categories: z.array(z.string()).optional(),
});

export const UpdateTaskInput = z.strictObject({
  task_id: z.string().min(1).describe('The task ID (td_ token)'),
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
  task_id: z.string().min(1).describe('The task ID (td_ token)'),
});

export const PrepareDeleteTaskInput = z.strictObject({
  task_id: z.string().min(1).describe('The task ID (td_ token)'),
});

export const ConfirmDeleteTaskInput = z.strictObject({
  token_id: z.uuid(),
  task_id: z.string().min(1).describe('The task ID (td_ token)'),
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

// =============================================================================
// Registry Definitions (v3 registry-driven architecture)
// =============================================================================

/**
 * Registry tool definitions for the tasks domain. Each handler delegates to
 * GraphTasksTools, which returns MCP content directly.
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
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'tasksGraph').listTasks(params),
    }),
    defineTool({
      name: 'search_tasks',
      description: 'Search tasks by name',
      input: SearchTasksInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'tasksGraph').searchTasks(params),
    }),
    defineTool({
      name: 'get_task',
      description: 'Get task details',
      input: GetTaskInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['tasks'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'tasksGraph').getTask(params),
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
