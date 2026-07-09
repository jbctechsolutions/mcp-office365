/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Shared types for Planner visualization renderers.
 */

export interface PlanVisualizationData {
  plan: { id: string; title: string };
  buckets: Array<{ id: string; name: string; orderHint: string }>;
  tasks: Array<{
    id: string;
    title: string;
    bucketId: string;
    percentComplete: number;
    priority: number;
    startDateTime?: string | null;
    dueDateTime?: string | null;
    assignments: string[];
    completedDateTime?: string | null;
  }>;
}

export type VisualizationFormat = 'html' | 'svg' | 'markdown' | 'mermaid';

/** Priority value to human-readable label mapping. */
export const PRIORITY_LABELS: Record<number, string> = {
  1: 'Urgent',
  3: 'Important',
  5: 'Medium',
  9: 'Low',
};

/** Priority value to color mapping. */
export const PRIORITY_COLORS: Record<number, string> = {
  1: '#e74c3c', // red - urgent
  3: '#e67e22', // orange - important
  5: '#f1c40f', // yellow - medium
  9: '#2ecc71', // green - low
};

/**
 * Categorize a task's status based on its percentComplete and dates.
 */
export function getTaskStatus(task: {
  percentComplete: number;
  dueDateTime?: string | null;
  completedDateTime?: string | null;
}): 'completed' | 'in-progress' | 'not-started' | 'overdue' {
  if (task.percentComplete === 100 || task.completedDateTime != null) return 'completed';
  if (
    task.dueDateTime != null &&
    task.completedDateTime == null &&
    new Date(task.dueDateTime) < new Date()
  ) {
    return 'overdue';
  }
  if (task.percentComplete > 0) return 'in-progress';
  return 'not-started';
}
