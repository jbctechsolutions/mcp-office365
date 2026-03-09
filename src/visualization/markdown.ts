/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Markdown renderers for Planner visualizations.
 */

import type { PlanVisualizationData } from './types.js';

/**
 * Renders a Kanban board as Markdown with columns for each bucket.
 */
export function renderKanbanMarkdown(data: PlanVisualizationData): string {
  const lines: string[] = [];
  lines.push(`# ${data.plan.title} — Kanban Board`);
  lines.push('');

  const bucketMap = new Map(data.buckets.map(b => [b.id, b.name]));

  for (const bucket of data.buckets) {
    lines.push(`## ${bucket.name}`);
    lines.push('');
    const bucketTasks = data.tasks.filter(t => t.bucketId === bucket.id);
    if (bucketTasks.length === 0) {
      lines.push('_No tasks_');
    } else {
      for (const task of bucketTasks) {
        const progress = task.percentComplete === 100 ? '[x]' : '[ ]';
        lines.push(`- ${progress} **${task.title}** (${task.percentComplete}%)`);
      }
    }
    lines.push('');
  }

  // Unassigned bucket tasks
  const unbucketed = data.tasks.filter(t => t.bucketId == null || !bucketMap.has(t.bucketId));
  if (unbucketed.length > 0) {
    lines.push('## Unassigned');
    lines.push('');
    for (const task of unbucketed) {
      const progress = task.percentComplete === 100 ? '[x]' : '[ ]';
      lines.push(`- ${progress} **${task.title}** (${task.percentComplete}%)`);
    }
    lines.push('');
  }

  return lines.join('\n');
}

/**
 * Renders a Gantt chart as Markdown table.
 */
export function renderGanttMarkdown(data: PlanVisualizationData): string {
  const lines: string[] = [];
  lines.push(`# ${data.plan.title} — Gantt Chart`);
  lines.push('');
  lines.push('| Task | Start | Due | Progress |');
  lines.push('|------|-------|-----|----------|');

  for (const task of data.tasks) {
    const start = task.startDateTime || '—';
    const due = task.dueDateTime || '—';
    lines.push(`| ${task.title} | ${start} | ${due} | ${task.percentComplete}% |`);
  }

  lines.push('');
  return lines.join('\n');
}

/**
 * Renders a plan summary as Markdown.
 */
export function renderSummaryMarkdown(data: PlanVisualizationData): string {
  const lines: string[] = [];
  const total = data.tasks.length;
  const completed = data.tasks.filter(t => t.percentComplete === 100).length;
  const inProgress = data.tasks.filter(t => t.percentComplete > 0 && t.percentComplete < 100).length;
  const notStarted = data.tasks.filter(t => t.percentComplete === 0).length;
  const avgProgress = total > 0 ? Math.round(data.tasks.reduce((sum, t) => sum + t.percentComplete, 0) / total) : 0;

  lines.push(`# ${data.plan.title} — Summary`);
  lines.push('');
  lines.push(`- **Total tasks:** ${total}`);
  lines.push(`- **Completed:** ${completed}`);
  lines.push(`- **In progress:** ${inProgress}`);
  lines.push(`- **Not started:** ${notStarted}`);
  lines.push(`- **Average progress:** ${avgProgress}%`);
  lines.push(`- **Buckets:** ${data.buckets.length}`);
  lines.push('');

  return lines.join('\n');
}

/**
 * Renders a burndown chart as Markdown table.
 */
export function renderBurndownMarkdown(data: PlanVisualizationData): string {
  const lines: string[] = [];
  const total = data.tasks.length;
  const completed = data.tasks.filter(t => t.percentComplete === 100).length;
  const remaining = total - completed;

  lines.push(`# ${data.plan.title} — Burndown`);
  lines.push('');
  lines.push(`- **Total tasks:** ${total}`);
  lines.push(`- **Completed:** ${completed}`);
  lines.push(`- **Remaining:** ${remaining}`);
  lines.push('');
  lines.push('| Status | Count |');
  lines.push('|--------|-------|');
  lines.push(`| Completed | ${completed} |`);
  lines.push(`| Remaining | ${remaining} |`);
  lines.push('');

  return lines.join('\n');
}
