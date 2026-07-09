/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Maps Microsoft Graph TodoTask type to TaskRow.
 */

import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import type { TaskRow } from '../../database/repository.js';
import {
  dateTimeTimeZoneToTimestamp,
  importanceToPriority,
  createGraphContentPath,
} from './utils.js';

/**
 * Extended TodoTask with taskListId for reference.
 */
export interface TodoTaskWithList extends MicrosoftGraph.TodoTask {
  taskListId?: string;
}

/**
 * Maps a Graph TodoTask to a TaskRow. `id`/`folderId` are pre-minted durable
 * tokens (`td_`/`tl_`) — mappers are pure and cannot mint alias tokens
 * themselves, so the repository computes them and passes them in.
 */
export function mapTaskToTaskRow(task: TodoTaskWithList, id: string | number, folderId: string | number): TaskRow {
  const taskId = task.id ?? '';

  // Type assertions needed due to Graph API's NullableOption types
  // which are incompatible with exactOptionalPropertyTypes
  const dueDateTime = task.dueDateTime as { dateTime?: string; timeZone?: string } | null | undefined;
  const startDateTime = task.startDateTime as { dateTime?: string; timeZone?: string } | null | undefined;

  return {
    id,
    folderId,
    name: task.title ?? null,
    isCompleted: task.status === 'completed' ? 1 : 0,
    dueDate: dateTimeTimeZoneToTimestamp(dueDateTime),
    startDate: dateTimeTimeZoneToTimestamp(startDateTime),
    priority: importanceToPriority(task.importance),
    hasReminder: task.isReminderOn === true ? 1 : 0,
    dataFilePath: createGraphContentPath('task', `${task.taskListId ?? 'default'}:${taskId}`),
  };
}
