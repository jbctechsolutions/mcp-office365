/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * HTML renderers for Planner visualizations.
 */

import type { PlanVisualizationData } from './types.js';

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * Renders a Kanban board as an HTML document.
 */
export function renderKanbanHtml(data: PlanVisualizationData): string {
  const bucketMap = new Map(data.buckets.map(b => [b.id, b.name]));

  let columns = '';
  for (const bucket of data.buckets) {
    const bucketTasks = data.tasks.filter(t => t.bucketId === bucket.id);
    let cards = '';
    for (const task of bucketTasks) {
      const bg = task.percentComplete === 100 ? '#d4edda' : task.percentComplete > 0 ? '#fff3cd' : '#f8f9fa';
      cards += `<div style="background:${bg};padding:8px;margin:4px 0;border-radius:4px;border:1px solid #dee2e6;">
        <strong>${escapeHtml(task.title)}</strong><br/>
        <small>Progress: ${task.percentComplete}%</small>
      </div>`;
    }
    columns += `<div style="flex:1;min-width:200px;margin:0 8px;">
      <h3 style="background:#007bff;color:white;padding:8px;border-radius:4px;text-align:center;">${escapeHtml(bucket.name)}</h3>
      ${cards || '<p style="color:#6c757d;text-align:center;">No tasks</p>'}
    </div>`;
  }

  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>${escapeHtml(data.plan.title)} — Kanban</title></head>
<body style="font-family:sans-serif;padding:20px;">
<h1>${escapeHtml(data.plan.title)} — Kanban Board</h1>
<div style="display:flex;overflow-x:auto;">${columns}</div>
</body></html>`;
}

/**
 * Renders a Gantt chart as an HTML table.
 */
export function renderGanttHtml(data: PlanVisualizationData): string {
  let rows = '';
  for (const task of data.tasks) {
    const start = task.startDateTime || '—';
    const due = task.dueDateTime || '—';
    rows += `<tr>
      <td>${escapeHtml(task.title)}</td>
      <td>${escapeHtml(start)}</td>
      <td>${escapeHtml(due)}</td>
      <td><div style="background:#e9ecef;border-radius:4px;overflow:hidden;"><div style="background:#28a745;height:20px;width:${task.percentComplete}%;"></div></div></td>
    </tr>`;
  }

  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>${escapeHtml(data.plan.title)} — Gantt</title></head>
<body style="font-family:sans-serif;padding:20px;">
<h1>${escapeHtml(data.plan.title)} — Gantt Chart</h1>
<table style="width:100%;border-collapse:collapse;">
<thead><tr><th style="text-align:left;border-bottom:2px solid #dee2e6;padding:8px;">Task</th><th style="border-bottom:2px solid #dee2e6;padding:8px;">Start</th><th style="border-bottom:2px solid #dee2e6;padding:8px;">Due</th><th style="border-bottom:2px solid #dee2e6;padding:8px;min-width:150px;">Progress</th></tr></thead>
<tbody>${rows}</tbody>
</table>
</body></html>`;
}

/**
 * Renders a plan summary as HTML.
 */
export function renderSummaryHtml(data: PlanVisualizationData): string {
  const total = data.tasks.length;
  const completed = data.tasks.filter(t => t.percentComplete === 100).length;
  const inProgress = data.tasks.filter(t => t.percentComplete > 0 && t.percentComplete < 100).length;
  const notStarted = data.tasks.filter(t => t.percentComplete === 0).length;
  const avgProgress = total > 0 ? Math.round(data.tasks.reduce((sum, t) => sum + t.percentComplete, 0) / total) : 0;

  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>${escapeHtml(data.plan.title)} — Summary</title></head>
<body style="font-family:sans-serif;padding:20px;">
<h1>${escapeHtml(data.plan.title)} — Summary</h1>
<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:16px;">
  <div style="background:#007bff;color:white;padding:16px;border-radius:8px;text-align:center;"><h2>${total}</h2><p>Total Tasks</p></div>
  <div style="background:#28a745;color:white;padding:16px;border-radius:8px;text-align:center;"><h2>${completed}</h2><p>Completed</p></div>
  <div style="background:#ffc107;color:black;padding:16px;border-radius:8px;text-align:center;"><h2>${inProgress}</h2><p>In Progress</p></div>
  <div style="background:#6c757d;color:white;padding:16px;border-radius:8px;text-align:center;"><h2>${notStarted}</h2><p>Not Started</p></div>
  <div style="background:#17a2b8;color:white;padding:16px;border-radius:8px;text-align:center;"><h2>${avgProgress}%</h2><p>Avg Progress</p></div>
</div>
</body></html>`;
}

/**
 * Renders a burndown chart as HTML.
 */
export function renderBurndownHtml(data: PlanVisualizationData): string {
  const total = data.tasks.length;
  const completed = data.tasks.filter(t => t.percentComplete === 100).length;
  const remaining = total - completed;
  const completedPct = total > 0 ? Math.round((completed / total) * 100) : 0;

  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>${escapeHtml(data.plan.title)} — Burndown</title></head>
<body style="font-family:sans-serif;padding:20px;">
<h1>${escapeHtml(data.plan.title)} — Burndown</h1>
<div style="max-width:400px;">
  <div style="display:flex;align-items:center;margin:8px 0;">
    <span style="width:100px;">Completed</span>
    <div style="flex:1;background:#e9ecef;border-radius:4px;overflow:hidden;height:30px;">
      <div style="background:#28a745;height:100%;width:${completedPct}%;display:flex;align-items:center;justify-content:center;color:white;font-weight:bold;">${completed}</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;margin:8px 0;">
    <span style="width:100px;">Remaining</span>
    <div style="flex:1;background:#e9ecef;border-radius:4px;overflow:hidden;height:30px;">
      <div style="background:#dc3545;height:100%;width:${total > 0 ? Math.round((remaining / total) * 100) : 0}%;display:flex;align-items:center;justify-content:center;color:white;font-weight:bold;">${remaining}</div>
    </div>
  </div>
</div>
</body></html>`;
}
