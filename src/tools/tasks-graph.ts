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
import { unixTimestampToLocalIso } from '../graph/mappers/utils.js';
import type { ToolResult } from '../registry/types.js';
import type { ListTasksParams, SearchTasksParams, GetTaskParams } from './tasks.js';

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Transforms a Graph task row to the summary shape returned by the graph
 * backend's task tools.
 */
export function transformTaskRow(row: TaskRow): {
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

/**
 * Graph task tools. Each method mirrors the extracted inline graph case body
 * and returns an MCP `ToolResult`.
 */
export class GraphTasksTools {
  constructor(
    private readonly repository: GraphRepository,
    private readonly contentReaders: GraphContentReaders
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
}
