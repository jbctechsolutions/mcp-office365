/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Graph-backend task tools (v3 registry-driven architecture, U2 — dual
 * backend). Holds the task logic that previously lived inline in the
 * `handleGraphToolCall` switch, so the registry handlers stay thin and branch
 * on `ctx.backend`.
 */

import type { GraphRepository } from '../graph/repository.js';
import type { GraphContentReaders } from '../graph/content-readers.js';
import type { TaskRow } from '../database/repository.js';
import type { ApprovalTokenManager } from '../approval/index.js';
import { hashTaskForApproval } from '../approval/index.js';
import { unixTimestampToLocalIso } from '../graph/mappers/utils.js';
import type { ToolResult } from '../registry/types.js';
import type {
  ListTasksParams,
  SearchTasksParams,
  GetTaskParams,
  CreateTaskParams,
  UpdateTaskParams,
  CompleteTaskParams,
  PrepareDeleteTaskParams,
  ConfirmDeleteTaskParams,
} from './tasks.js';

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Transforms a Graph task row to the summary shape returned by the graph
 * backend's task tools.
 */
export function transformTaskRow(row: TaskRow): {
  id: string | number;
  folderId: string | number;
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

/**
 * Graph task tools. Each method mirrors the extracted inline graph case body
 * and returns an MCP `ToolResult`.
 */
export class GraphTasksTools {
  constructor(
    private readonly repository: GraphRepository,
    private readonly contentReaders: GraphContentReaders,
    private readonly tokenManager: ApprovalTokenManager
  ) {}

  async listTasks(params: ListTasksParams): Promise<ToolResult> {
    const tasks = params.include_completed
      ? await this.repository.listTasksAsync(params.limit, params.offset)
      : await this.repository.listIncompleteTasksAsync(params.limit, params.offset);
    return jsonResult({ tasks: tasks.map(transformTaskRow) });
  }

  async searchTasks(params: SearchTasksParams): Promise<ToolResult> {
    const tasks = await this.repository.searchTasksAsync(params.query, params.limit);
    return jsonResult({ tasks: tasks.map(transformTaskRow) });
  }

  async getTask(params: GetTaskParams): Promise<ToolResult> {
    const task = await this.repository.getTaskAsync(params.task_id);
    if (task == null) {
      return { content: [{ type: 'text', text: 'Task not found' }], isError: true };
    }

    const details = await this.contentReaders.task.readTaskDetailsAsync(task.dataFilePath);
    return jsonResult({ ...transformTaskRow(task), ...details });
  }

  async createTask(params: CreateTaskParams): Promise<ToolResult> {
    const taskId = await this.repository.createTaskAsync({
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
    return jsonResult({
      id: taskId,
      title: params.title,
      task_list_id: params.task_list_id,
      status: 'created',
    });
  }

  async updateTask(params: UpdateTaskParams): Promise<ToolResult> {
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
    await this.repository.updateTaskAsync(params.task_id, updates);
    return { content: [{ type: 'text', text: `Successfully updated task ${params.task_id}` }] };
  }

  async completeTask(params: CompleteTaskParams): Promise<ToolResult> {
    await this.repository.completeTaskAsync(params.task_id);
    return { content: [{ type: 'text', text: `Successfully completed task ${params.task_id}` }] };
  }

  async prepareDeleteTask(params: PrepareDeleteTaskParams): Promise<ToolResult> {
    const task = await this.repository.getTaskAsync(params.task_id);
    if (task == null) {
      return { content: [{ type: 'text', text: 'Task not found' }], isError: true };
    }

    const taskInfo = this.repository.getTaskInfo(params.task_id);
    const graphTask = taskInfo != null
      ? await this.repository.getClient().getTask(taskInfo.taskListId, taskInfo.taskId)
      : null;
    const hash = hashTaskForApproval({
      taskId: taskInfo?.taskId ?? '',
      title: graphTask?.title ?? null,
      listId: taskInfo?.taskListId ?? '',
    });

    const token = this.tokenManager.generateToken({
      operation: 'delete_task',
      targetType: 'task',
      targetId: params.task_id,
      targetHash: hash,
    });

    return jsonResult({
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      task: transformTaskRow(task),
      action: 'This task will be permanently deleted.',
    });
  }

  async confirmDeleteTask(params: ConfirmDeleteTaskParams): Promise<ToolResult> {
    // Re-fetch the task and compute fresh hash for comparison
    const taskInfo = this.repository.getTaskInfo(params.task_id);
    const graphTask = taskInfo != null
      ? await this.repository.getClient().getTask(taskInfo.taskListId, taskInfo.taskId)
      : null;
    const currentHash = hashTaskForApproval({
      taskId: taskInfo?.taskId ?? '',
      title: graphTask?.title ?? null,
      listId: taskInfo?.taskListId ?? '',
    });

    const validation = this.tokenManager.consumeToken(params.token_id, 'delete_task', params.task_id);
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

    await this.repository.deleteTaskAsync(params.task_id);
    return { content: [{ type: 'text', text: `Successfully deleted task ${params.task_id}` }] };
  }
}
