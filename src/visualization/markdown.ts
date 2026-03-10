/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Markdown visualization renderer for Planner data.
 */

import {
  type PlanVisualizationData,
  PRIORITY_LABELS,
  getTaskStatus,
} from './types.js';

/**
 * Render a Kanban board as markdown tables grouped by bucket.
 */
export function renderKanbanMarkdown(data: PlanVisualizationData): string {
  const lines: string[] = [`# ${data.plan.title} - Kanban Board`, ''];

  for (const bucket of data.buckets) {
    const bucketTasks = data.tasks.filter((t) => t.bucketId === bucket.id);
    lines.push(`## ${bucket.name}`, '');
    lines.push('| Title | Priority | Assignees | % Complete | Due Date |');
    lines.push('|-------|----------|-----------|------------|----------|');

    if (bucketTasks.length === 0) {
      lines.push('| _No tasks_ | | | | |');
    } else {
      for (const task of bucketTasks) {
        const priority = PRIORITY_LABELS[task.priority] ?? `P${task.priority}`;
        const assignees = task.assignments.length > 0 ? task.assignments.join(', ') : '-';
        const due = task.dueDateTime != null
          ? new Date(task.dueDateTime).toISOString().split('T')[0]
          : '-';
        lines.push(
          `| ${task.title} | ${priority} | ${assignees} | ${task.percentComplete}% | ${due} |`
        );
      }
    }
    lines.push('');
  }

  // Handle tasks not assigned to any known bucket
  const knownBucketIds = new Set(data.buckets.map((b) => b.id));
  const orphanTasks = data.tasks.filter((t) => !knownBucketIds.has(t.bucketId));
  if (orphanTasks.length > 0) {
    lines.push('## Unassigned Bucket', '');
    lines.push('| Title | Priority | Assignees | % Complete | Due Date |');
    lines.push('|-------|----------|-----------|------------|----------|');
    for (const task of orphanTasks) {
      const priority = PRIORITY_LABELS[task.priority] ?? `P${task.priority}`;
      const assignees = task.assignments.length > 0 ? task.assignments.join(', ') : '-';
      const due = task.dueDateTime != null
        ? new Date(task.dueDateTime).toISOString().split('T')[0]
        : '-';
      lines.push(
        `| ${task.title} | ${priority} | ${assignees} | ${task.percentComplete}% | ${due} |`
      );
    }
    lines.push('');
  }

  return lines.join('\n');
}

/**
 * Render a Gantt chart as a Mermaid code block inside markdown.
 */
export function renderGanttMarkdown(data: PlanVisualizationData): string {
  const lines: string[] = [
    `# ${data.plan.title} - Gantt Chart`,
    '',
    '```mermaid',
    'gantt',
    '    dateFormat YYYY-MM-DD',
    `    title ${data.plan.title}`,
  ];

  for (const bucket of data.buckets) {
    const bucketTasks = data.tasks.filter((t) => t.bucketId === bucket.id);
    if (bucketTasks.length === 0) continue;
    lines.push(`    section ${bucket.name}`);
    for (const task of bucketTasks) {
      const start = task.startDateTime != null
        ? new Date(task.startDateTime).toISOString().split('T')[0]
        : null;
      const end = task.dueDateTime != null
        ? new Date(task.dueDateTime).toISOString().split('T')[0]
        : null;
      const status = task.percentComplete === 100 ? 'done, ' : task.percentComplete > 0 ? 'active, ' : '';
      if (start != null && end != null) {
        lines.push(`    ${task.title} :${status}${start}, ${end}`);
      } else if (start != null) {
        lines.push(`    ${task.title} :${status}${start}, 7d`);
      } else if (end != null) {
        // Estimate start as 7 days before due
        const estimatedStart = new Date(end);
        estimatedStart.setDate(estimatedStart.getDate() - 7);
        const startStr = estimatedStart.toISOString().split('T')[0];
        lines.push(`    ${task.title} :${status}${startStr}, ${end}`);
      } else {
        lines.push(`    ${task.title} :${status}2026-01-01, 7d`);
      }
    }
  }

  lines.push('```', '');
  return lines.join('\n');
}

/**
 * Render a summary dashboard as markdown with stats and assignee workload tables.
 */
export function renderSummaryMarkdown(data: PlanVisualizationData): string {
  const total = data.tasks.length;
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
    `# ${data.plan.title} - Summary`,
    '',
    '## Task Statistics',
    '',
    '| Metric | Count |',
    '|--------|-------|',
    `| Total Tasks | ${total} |`,
    `| Not Started | ${notStarted} |`,
    `| In Progress | ${inProgress} |`,
    `| Completed | ${completed} |`,
    `| Overdue | ${overdue} |`,
    '',
  ];

  // Assignee workload
  const assigneeMap = new Map<string, number>();
  for (const task of data.tasks) {
    for (const assignee of task.assignments) {
      assigneeMap.set(assignee, (assigneeMap.get(assignee) ?? 0) + 1);
    }
  }

  if (assigneeMap.size > 0) {
    lines.push('## Assignee Workload', '');
    lines.push('| Assignee | Task Count |');
    lines.push('|----------|------------|');
    for (const [assignee, count] of [...assigneeMap.entries()].sort((a, b) => b[1] - a[1])) {
      lines.push(`| ${assignee} | ${count} |`);
    }
    lines.push('');
  }

  return lines.join('\n');
}

/**
 * Render a burndown chart as a date-indexed markdown table.
 */
export function renderBurndownMarkdown(data: PlanVisualizationData): string {
  const lines: string[] = [`# ${data.plan.title} - Burndown`, ''];

  // Collect relevant dates
  const tasksWithDue = data.tasks.filter((t) => t.dueDateTime);
  if (tasksWithDue.length === 0) {
    lines.push('_No tasks with due dates to generate burndown chart._', '');
    return lines.join('\n');
  }

  const allDates = new Set<string>();
  for (const task of data.tasks) {
    if (task.startDateTime != null) allDates.add(new Date(task.startDateTime).toISOString().split('T')[0]!);
    if (task.dueDateTime != null) allDates.add(new Date(task.dueDateTime).toISOString().split('T')[0]!);
    if (task.completedDateTime != null) allDates.add(new Date(task.completedDateTime).toISOString().split('T')[0]!);
  }

  const sortedDates = [...allDates].sort();
  if (sortedDates.length === 0) {
    lines.push('_No date data available._', '');
    return lines.join('\n');
  }

  lines.push('| Date | Remaining Tasks | Completed |');
  lines.push('|------|-----------------|-----------|');

  const total = data.tasks.length;
  for (const date of sortedDates) {
    const completedByDate = data.tasks.filter(
      (t) => t.completedDateTime != null && new Date(t.completedDateTime).toISOString().split('T')[0]! <= date
    ).length;
    lines.push(`| ${date} | ${total - completedByDate} | ${completedByDate} |`);
  }

  lines.push('');
  return lines.join('\n');
}
