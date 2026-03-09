/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Mermaid diagram renderers for Planner visualizations.
 */

import type { PlanVisualizationData } from './types.js';

/**
 * Renders a Kanban board as a Mermaid block diagram.
 */
export function renderKanbanMermaid(data: PlanVisualizationData): string {
  const lines: string[] = [];
  lines.push('```mermaid');
  lines.push('block-beta');

  const bucketMap = new Map(data.buckets.map(b => [b.id, b.name]));

  for (const bucket of data.buckets) {
    const safeName = bucket.name.replace(/[^a-zA-Z0-9]/g, '_');
    lines.push(`  columns 1`);
    lines.push(`  ${safeName}["${bucket.name}"]`);
    const bucketTasks = data.tasks.filter(t => t.bucketId === bucket.id);
    for (const task of bucketTasks) {
      const safeTask = `task_${task.id}`;
      const icon = task.percentComplete === 100 ? '✓' : '○';
      lines.push(`  ${safeTask}["${icon} ${task.title} (${task.percentComplete}%)"]`);
    }
  }

  lines.push('```');
  return lines.join('\n');
}

/**
 * Renders a Gantt chart as Mermaid gantt diagram.
 */
export function renderGanttMermaid(data: PlanVisualizationData): string {
  const lines: string[] = [];
  lines.push('```mermaid');
  lines.push('gantt');
  lines.push(`  title ${data.plan.title}`);
  lines.push('  dateFormat YYYY-MM-DD');
  lines.push('');

  const bucketMap = new Map(data.buckets.map(b => [b.id, b.name]));

  for (const bucket of data.buckets) {
    lines.push(`  section ${bucket.name}`);
    const bucketTasks = data.tasks.filter(t => t.bucketId === bucket.id);
    for (const task of bucketTasks) {
      const start = task.startDateTime ? task.startDateTime.slice(0, 10) : new Date().toISOString().slice(0, 10);
      const due = task.dueDateTime ? task.dueDateTime.slice(0, 10) : start;
      const status = task.percentComplete === 100 ? 'done,' : task.percentComplete > 0 ? 'active,' : '';
      lines.push(`  ${task.title} :${status} ${start}, ${due}`);
    }
  }

  lines.push('```');
  return lines.join('\n');
}

/**
 * Renders a summary as a Mermaid pie chart.
 */
export function renderSummaryMermaid(data: PlanVisualizationData): string {
  const completed = data.tasks.filter(t => t.percentComplete === 100).length;
  const inProgress = data.tasks.filter(t => t.percentComplete > 0 && t.percentComplete < 100).length;
  const notStarted = data.tasks.filter(t => t.percentComplete === 0).length;

  const lines: string[] = [];
  lines.push('```mermaid');
  lines.push('pie');
  lines.push(`  title ${data.plan.title} — Task Status`);
  if (completed > 0) lines.push(`  "Completed" : ${completed}`);
  if (inProgress > 0) lines.push(`  "In Progress" : ${inProgress}`);
  if (notStarted > 0) lines.push(`  "Not Started" : ${notStarted}`);
  lines.push('```');
  return lines.join('\n');
}

/**
 * Renders a burndown as a Mermaid XY chart.
 */
export function renderBurndownMermaid(data: PlanVisualizationData): string {
  const total = data.tasks.length;
  const completed = data.tasks.filter(t => t.percentComplete === 100).length;
  const remaining = total - completed;

  const lines: string[] = [];
  lines.push('```mermaid');
  lines.push('xychart-beta');
  lines.push(`  title "${data.plan.title} — Burndown"`);
  lines.push(`  x-axis ["Completed", "Remaining"]`);
  lines.push(`  y-axis "Tasks" 0 --> ${total}`);
  lines.push(`  bar [${completed}, ${remaining}]`);
  lines.push('```');
  return lines.join('\n');
}
