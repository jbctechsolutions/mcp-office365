/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Task-related type definitions.
 */

import type { PriorityValue } from './mail.js';

/**
 * Task summary for list views.
 */
export interface TaskSummary {
  readonly id: number;
  readonly folderId: number;
  readonly name: string | null;
  readonly isCompleted: boolean;
  readonly dueDate: string | null;
  readonly priority: PriorityValue;
}

/**
 * Full task details including body content.
 */
export interface Task extends TaskSummary {
  readonly startDate: string | null;
  readonly completedDate: string | null;
  readonly hasReminder: boolean;
  readonly reminderDate: string | null;
  readonly body: string | null;
  readonly categories: readonly string[];
}
