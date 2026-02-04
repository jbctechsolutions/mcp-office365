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

// =============================================================================
// Input Schemas
// =============================================================================

export const ListTasksInput = z
  .object({
    limit: z
      .number()
      .int()
      .min(1)
      .max(100)
      .default(50)
      .describe('Maximum number of tasks to return (1-100)'),
    offset: z.number().int().min(0).default(0).describe('Number of tasks to skip'),
    include_completed: z.boolean().default(true).describe('Include completed tasks'),
  })
  .strict();

export const SearchTasksInput = z
  .object({
    query: z.string().min(1).describe('Search query for task names'),
    limit: z
      .number()
      .int()
      .min(1)
      .max(100)
      .default(50)
      .describe('Maximum number of tasks to return (1-100)'),
  })
  .strict();

export const GetTaskInput = z
  .object({
    task_id: z.number().int().positive().describe('The task ID to retrieve'),
  })
  .strict();

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
