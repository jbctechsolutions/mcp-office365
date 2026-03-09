/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Mermaid diagram renderer for Planner data.
 */

import { type PlanVisualizationData, getTaskStatus } from './types.js';

/**
 * Render a Gantt chart in pure Mermaid gantt syntax.
 */
export function renderGanttMermaid(data: PlanVisualizationData): string {
  const lines: string[] = [
    'gantt',
    '    dateFormat YYYY-MM-DD',
    `    title ${data.plan.title}`,
  ];

  for (const bucket of data.buckets) {
    const bucketTasks = data.tasks.filter((t) => t.bucketId === bucket.id);
    if (bucketTasks.length === 0) continue;
    lines.push(`    section ${bucket.name}`);
    for (const task of bucketTasks) {
      const start = task.startDateTime
        ? new Date(task.startDateTime).toISOString().split('T')[0]
        : null;
      const end = task.dueDateTime
        ? new Date(task.dueDateTime).toISOString().split('T')[0]
        : null;
      const status =
        task.percentComplete === 100
          ? 'done, '
          : task.percentComplete > 0
            ? 'active, '
            : '';
      if (start && end) {
        lines.push(`    ${task.title} :${status}${start}, ${end}`);
      } else if (start) {
        lines.push(`    ${task.title} :${status}${start}, 7d`);
      } else if (end) {
        const estimatedStart = new Date(end);
        estimatedStart.setDate(estimatedStart.getDate() - 7);
        const startStr = estimatedStart.toISOString().split('T')[0];
        lines.push(`    ${task.title} :${status}${startStr}, ${end}`);
      } else {
        lines.push(`    ${task.title} :${status}2026-01-01, 7d`);
      }
    }
  }

  return lines.join('\n');
}

/**
 * Render a Kanban board using Mermaid block-beta diagram.
 */
export function renderKanbanMermaid(data: PlanVisualizationData): string {
  const lines: string[] = ['block-beta'];
  const colCount = Math.max(data.buckets.length, 1);
  lines.push(`    columns ${colCount}`);

  // Header row
  for (const bucket of data.buckets) {
    const safeId = bucket.name.replace(/[^a-zA-Z0-9]/g, '_');
    lines.push(`    ${safeId}["${bucket.name}"]`);
  }

  // Task rows - each bucket column gets its tasks
  // We need to find the max number of tasks across buckets
  const tasksByBucket = data.buckets.map((b) =>
    data.tasks.filter((t) => t.bucketId === b.id)
  );
  const maxTasks = Math.max(...tasksByBucket.map((ts) => ts.length), 0);

  for (let row = 0; row < maxTasks; row++) {
    for (let col = 0; col < data.buckets.length; col++) {
      const tasks = tasksByBucket[col];
      if (tasks && row < tasks.length) {
        const task = tasks[row]!;
        const safeId = `task_${task.id}`;
        lines.push(`    ${safeId}["${task.title} (${task.percentComplete}%)"]`);
      } else {
        lines.push(`    space`);
      }
    }
  }

  return lines.join('\n');
}

/**
 * Render a summary pie chart in Mermaid syntax.
 */
export function renderSummaryMermaid(data: PlanVisualizationData): string {
  let notStarted = 0;
  let inProgress = 0;
  let completed = 0;
  let overdue = 0;

  for (const task of data.tasks) {
    const status = getTaskStatus(task);
    if (status === 'completed') completed++;
    else if (status === 'overdue') overdue++;
    else if (status === 'in-progress') inProgress++;
    else notStarted++;
  }

  const lines: string[] = [
    'pie',
    `    title ${data.plan.title} - Task Status`,
  ];

  if (notStarted > 0) lines.push(`    "Not Started" : ${notStarted}`);
  if (inProgress > 0) lines.push(`    "In Progress" : ${inProgress}`);
  if (completed > 0) lines.push(`    "Completed" : ${completed}`);
  if (overdue > 0) lines.push(`    "Overdue" : ${overdue}`);

  // If all zero, show a placeholder
  if (data.tasks.length === 0) {
    lines.push('    "No Tasks" : 1');
  }

  return lines.join('\n');
}

/**
 * Render a burndown chart using Mermaid xychart-beta.
 */
export function renderBurndownMermaid(data: PlanVisualizationData): string {
  const tasksWithDates = data.tasks.filter(
    (t) => t.dueDateTime || t.completedDateTime || t.startDateTime
  );

  if (tasksWithDates.length === 0) {
    return ['xychart-beta', '    title No date data available', '    x-axis ["N/A"]', '    y-axis "Tasks" 0 --> 1', '    line [0]'].join('\n');
  }

  const allDates = new Set<string>();
  for (const task of data.tasks) {
    if (task.startDateTime) allDates.add(new Date(task.startDateTime).toISOString().split('T')[0]!);
    if (task.dueDateTime) allDates.add(new Date(task.dueDateTime).toISOString().split('T')[0]!);
    if (task.completedDateTime) allDates.add(new Date(task.completedDateTime).toISOString().split('T')[0]!);
  }

  const sortedDates = [...allDates].sort();
  const total = data.tasks.length;
  const remaining: number[] = [];

  for (const date of sortedDates) {
    const completedByDate = data.tasks.filter(
      (t) =>
        t.completedDateTime &&
        new Date(t.completedDateTime).toISOString().split('T')[0]! <= date
    ).length;
    remaining.push(total - completedByDate);
  }

  const dateLabels = sortedDates.map((d) => `"${d}"`).join(', ');
  const dataPoints = remaining.join(', ');

  const lines: string[] = [
    'xychart-beta',
    `    title ${data.plan.title} - Burndown`,
    `    x-axis [${dateLabels}]`,
    `    y-axis "Remaining Tasks" 0 --> ${total}`,
    `    line [${dataPoints}]`,
  ];

  return lines.join('\n');
}
