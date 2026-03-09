/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Self-contained HTML visualization renderer for Planner data.
 * All styles are inline -- no external dependencies.
 */

import {
  type PlanVisualizationData,
  PRIORITY_LABELS,
  PRIORITY_COLORS,
  getTaskStatus,
} from './types.js';

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function priorityColor(priority: number): string {
  return PRIORITY_COLORS[priority] ?? '#95a5a6';
}

/**
 * Render a Kanban board as self-contained HTML with CSS Grid.
 */
export function renderKanbanHtml(data: PlanVisualizationData): string {
  const colCount = Math.max(data.buckets.length, 1);

  let cardsHtml = '';
  for (const bucket of data.buckets) {
    const bucketTasks = data.tasks.filter((t) => t.bucketId === bucket.id);
    let tasksHtml = '';
    for (const task of bucketTasks) {
      const pColor = priorityColor(task.priority);
      const pLabel = PRIORITY_LABELS[task.priority] ?? `P${task.priority}`;
      const assignees = task.assignments.length > 0 ? task.assignments.join(', ') : 'Unassigned';
      const due = task.dueDateTime
        ? new Date(task.dueDateTime).toISOString().split('T')[0]
        : 'No due date';
      tasksHtml += `
        <div class="kanban-card" style="border-left:4px solid ${pColor};background:#fff;border-radius:6px;padding:10px;margin-bottom:8px;box-shadow:0 1px 3px rgba(0,0,0,0.12);cursor:default;" title="Assignees: ${escapeHtml(assignees)}&#10;Due: ${escapeHtml(due)}&#10;Priority: ${escapeHtml(pLabel)}">
          <div style="font-weight:600;margin-bottom:4px;">${escapeHtml(task.title)}</div>
          <div style="font-size:12px;color:#666;">
            <span style="display:inline-block;background:${pColor};color:#fff;padding:1px 6px;border-radius:3px;font-size:11px;">${escapeHtml(pLabel)}</span>
            <span style="margin-left:6px;">${task.percentComplete}%</span>
          </div>
        </div>`;
    }
    if (bucketTasks.length === 0) {
      tasksHtml = '<div style="color:#999;font-style:italic;padding:10px;">No tasks</div>';
    }
    cardsHtml += `
      <div class="kanban-column" style="background:#f4f5f7;border-radius:8px;padding:12px;min-width:220px;">
        <div style="font-weight:700;font-size:14px;margin-bottom:10px;padding-bottom:8px;border-bottom:2px solid #ddd;">${escapeHtml(bucket.name)} <span style="color:#999;font-weight:400;">(${bucketTasks.length})</span></div>
        ${tasksHtml}
      </div>`;
  }

  return `<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(data.plan.title)} - Kanban</title></head>
<body style="margin:0;padding:20px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#e8eaed;">
  <h1 style="margin:0 0 16px 0;font-size:22px;">${escapeHtml(data.plan.title)} - Kanban Board</h1>
  <div style="display:grid;grid-template-columns:repeat(${colCount},minmax(240px,1fr));gap:16px;overflow-x:auto;">
    ${cardsHtml}
  </div>
</body>
</html>`;
}

/**
 * Render a Gantt chart as self-contained HTML with CSS positioning.
 */
export function renderGanttHtml(data: PlanVisualizationData): string {
  // Collect all tasks that have at least one date
  const tasksWithDates = data.tasks.filter((t) => t.startDateTime || t.dueDateTime);
  const allTasks = tasksWithDates.length > 0 ? tasksWithDates : data.tasks;

  // Determine date range
  const dates: Date[] = [];
  for (const task of allTasks) {
    if (task.startDateTime) dates.push(new Date(task.startDateTime));
    if (task.dueDateTime) dates.push(new Date(task.dueDateTime));
  }

  if (dates.length === 0) {
    return `<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><title>${escapeHtml(data.plan.title)} - Gantt</title></head>
<body style="margin:20px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;">
  <h1>${escapeHtml(data.plan.title)} - Gantt Chart</h1>
  <p>No tasks with dates to display.</p>
</body>
</html>`;
  }

  const minDate = new Date(Math.min(...dates.map((d) => d.getTime())));
  const maxDate = new Date(Math.max(...dates.map((d) => d.getTime())));
  // Add some padding
  minDate.setDate(minDate.getDate() - 1);
  maxDate.setDate(maxDate.getDate() + 1);
  const totalDays = Math.max(Math.ceil((maxDate.getTime() - minDate.getTime()) / 86400000), 1);

  const rowHeight = 36;
  const labelWidth = 200;
  const chartWidth = 800;
  const headerHeight = 40;

  // Date axis labels (show ~8 evenly spaced dates)
  const dateAxisCount = Math.min(totalDays, 8);
  let dateAxisHtml = '';
  for (let i = 0; i <= dateAxisCount; i++) {
    const d = new Date(minDate.getTime() + (i / dateAxisCount) * (maxDate.getTime() - minDate.getTime()));
    const x = (i / dateAxisCount) * chartWidth;
    dateAxisHtml += `<div style="position:absolute;left:${labelWidth + x}px;top:10px;font-size:11px;color:#666;transform:translateX(-50%);">${d.toISOString().split('T')[0]}</div>`;
  }

  let barsHtml = '';
  let rowIndex = 0;
  for (const task of allTasks) {
    const start = task.startDateTime ? new Date(task.startDateTime) : task.dueDateTime ? new Date(new Date(task.dueDateTime).getTime() - 7 * 86400000) : minDate;
    const end = task.dueDateTime ? new Date(task.dueDateTime) : new Date(start.getTime() + 7 * 86400000);
    const startOffset = Math.max(0, (start.getTime() - minDate.getTime()) / (maxDate.getTime() - minDate.getTime())) * chartWidth;
    const barWidth = Math.max(10, ((end.getTime() - start.getTime()) / (maxDate.getTime() - minDate.getTime())) * chartWidth);
    const y = headerHeight + rowIndex * rowHeight;
    const pColor = priorityColor(task.priority);
    const progressWidth = (task.percentComplete / 100) * barWidth;

    barsHtml += `
      <div style="position:absolute;left:0;top:${y}px;width:${labelWidth - 10}px;height:${rowHeight}px;display:flex;align-items:center;padding-left:8px;font-size:13px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;">${escapeHtml(task.title)}</div>
      <div style="position:absolute;left:${labelWidth + startOffset}px;top:${y + 6}px;width:${barWidth}px;height:${rowHeight - 14}px;background:#ddd;border-radius:4px;overflow:hidden;">
        <div style="width:${progressWidth}px;height:100%;background:${pColor};border-radius:4px 0 0 4px;"></div>
      </div>`;
    rowIndex++;
  }

  const totalHeight = headerHeight + rowIndex * rowHeight + 20;

  return `<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(data.plan.title)} - Gantt</title></head>
<body style="margin:0;padding:20px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#fafafa;">
  <h1 style="font-size:22px;margin:0 0 16px 0;">${escapeHtml(data.plan.title)} - Gantt Chart</h1>
  <div style="position:relative;width:${labelWidth + chartWidth}px;height:${totalHeight}px;background:#fff;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;padding:0;">
    ${dateAxisHtml}
    ${barsHtml}
  </div>
</body>
</html>`;
}

/**
 * Render a summary dashboard as self-contained HTML.
 */
export function renderSummaryHtml(data: PlanVisualizationData): string {
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

  const total = data.tasks.length;

  // Assignee workload
  const assigneeMap = new Map<string, number>();
  for (const task of data.tasks) {
    for (const assignee of task.assignments) {
      assigneeMap.set(assignee, (assigneeMap.get(assignee) ?? 0) + 1);
    }
  }

  let assigneeRows = '';
  for (const [assignee, count] of [...assigneeMap.entries()].sort((a, b) => b[1] - a[1])) {
    assigneeRows += `<tr><td style="padding:6px 12px;border-bottom:1px solid #eee;">${escapeHtml(assignee)}</td><td style="padding:6px 12px;border-bottom:1px solid #eee;text-align:center;">${count}</td></tr>`;
  }

  const statCard = (label: string, value: number, color: string) =>
    `<div style="background:#fff;border-radius:8px;padding:16px 20px;box-shadow:0 1px 3px rgba(0,0,0,0.1);text-align:center;min-width:120px;">
       <div style="font-size:28px;font-weight:700;color:${color};">${value}</div>
       <div style="font-size:13px;color:#666;margin-top:4px;">${label}</div>
     </div>`;

  // Priority distribution
  const priorityCounts: Record<number, number> = {};
  for (const task of data.tasks) {
    priorityCounts[task.priority] = (priorityCounts[task.priority] ?? 0) + 1;
  }

  let priorityBars = '';
  for (const [p, count] of Object.entries(priorityCounts).sort(([a], [b]) => Number(a) - Number(b))) {
    const pNum = Number(p);
    const pct = total > 0 ? (count / total) * 100 : 0;
    const color = priorityColor(pNum);
    const label = PRIORITY_LABELS[pNum] ?? `P${p}`;
    priorityBars += `
      <div style="display:flex;align-items:center;margin-bottom:6px;">
        <div style="width:80px;font-size:13px;">${escapeHtml(label)}</div>
        <div style="flex:1;background:#eee;border-radius:4px;height:20px;overflow:hidden;">
          <div style="width:${pct}%;height:100%;background:${color};border-radius:4px;"></div>
        </div>
        <div style="width:40px;text-align:right;font-size:13px;margin-left:8px;">${count}</div>
      </div>`;
  }

  return `<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(data.plan.title)} - Summary</title></head>
<body style="margin:0;padding:20px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#f0f2f5;">
  <h1 style="font-size:22px;margin:0 0 20px 0;">${escapeHtml(data.plan.title)} - Summary Dashboard</h1>
  <div style="display:flex;gap:16px;flex-wrap:wrap;margin-bottom:24px;">
    ${statCard('Total Tasks', total, '#2c3e50')}
    ${statCard('Not Started', notStarted, '#95a5a6')}
    ${statCard('In Progress', inProgress, '#3498db')}
    ${statCard('Completed', completed, '#2ecc71')}
    ${statCard('Overdue', overdue, '#e74c3c')}
  </div>
  ${priorityBars ? `<div style="background:#fff;border-radius:8px;padding:16px 20px;box-shadow:0 1px 3px rgba(0,0,0,0.1);margin-bottom:24px;max-width:500px;">
    <h2 style="font-size:16px;margin:0 0 12px 0;">Priority Distribution</h2>
    ${priorityBars}
  </div>` : ''}
  ${assigneeRows ? `<div style="background:#fff;border-radius:8px;padding:16px 20px;box-shadow:0 1px 3px rgba(0,0,0,0.1);max-width:400px;">
    <h2 style="font-size:16px;margin:0 0 12px 0;">Assignee Workload</h2>
    <table style="width:100%;border-collapse:collapse;">
      <thead><tr><th style="text-align:left;padding:6px 12px;border-bottom:2px solid #ddd;">Assignee</th><th style="text-align:center;padding:6px 12px;border-bottom:2px solid #ddd;">Tasks</th></tr></thead>
      <tbody>${assigneeRows}</tbody>
    </table>
  </div>` : ''}
</body>
</html>`;
}

/**
 * Render a burndown chart as self-contained HTML with an inline SVG line chart.
 */
export function renderBurndownHtml(data: PlanVisualizationData): string {
  const allDates = new Set<string>();
  for (const task of data.tasks) {
    if (task.startDateTime) allDates.add(new Date(task.startDateTime).toISOString().split('T')[0]!);
    if (task.dueDateTime) allDates.add(new Date(task.dueDateTime).toISOString().split('T')[0]!);
    if (task.completedDateTime) allDates.add(new Date(task.completedDateTime).toISOString().split('T')[0]!);
  }

  const sortedDates = [...allDates].sort();
  if (sortedDates.length === 0) {
    return `<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><title>${escapeHtml(data.plan.title)} - Burndown</title></head>
<body style="margin:20px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;">
  <h1>${escapeHtml(data.plan.title)} - Burndown Chart</h1>
  <p>No date data available for burndown chart.</p>
</body></html>`;
  }

  const total = data.tasks.length;
  const remaining: number[] = [];
  for (const date of sortedDates) {
    const completedByDate = data.tasks.filter(
      (t) => t.completedDateTime && new Date(t.completedDateTime).toISOString().split('T')[0]! <= date
    ).length;
    remaining.push(total - completedByDate);
  }

  // SVG chart dimensions
  const svgW = 700;
  const svgH = 350;
  const pad = { top: 30, right: 30, bottom: 60, left: 50 };
  const chartW = svgW - pad.left - pad.right;
  const chartH = svgH - pad.top - pad.bottom;
  const n = sortedDates.length;
  const yMax = Math.max(total, 1);

  // Actual line points
  const points = remaining.map((r, i) => {
    const x = pad.left + (n > 1 ? (i / (n - 1)) * chartW : chartW / 2);
    const y = pad.top + chartH - (r / yMax) * chartH;
    return `${x},${y}`;
  });

  // Ideal burndown line (straight from total to 0)
  const idealStart = `${pad.left},${pad.top}`;
  const idealEnd = `${pad.left + chartW},${pad.top + chartH}`;

  // X-axis labels
  let xLabels = '';
  const labelCount = Math.min(n, 8);
  for (let i = 0; i < labelCount; i++) {
    const idx = Math.floor(i * (n - 1) / Math.max(labelCount - 1, 1));
    const x = pad.left + (n > 1 ? (idx / (n - 1)) * chartW : chartW / 2);
    xLabels += `<text x="${x}" y="${svgH - 15}" text-anchor="middle" font-size="11" fill="#666" transform="rotate(-30,${x},${svgH - 15})">${sortedDates[idx]}</text>`;
  }

  // Y-axis labels
  let yLabels = '';
  const yTicks = 5;
  for (let i = 0; i <= yTicks; i++) {
    const val = Math.round((i / yTicks) * yMax);
    const y = pad.top + chartH - (i / yTicks) * chartH;
    yLabels += `<text x="${pad.left - 8}" y="${y + 4}" text-anchor="end" font-size="11" fill="#666">${val}</text>`;
    yLabels += `<line x1="${pad.left}" y1="${y}" x2="${pad.left + chartW}" y2="${y}" stroke="#eee" stroke-width="1"/>`;
  }

  // Data point circles
  const circles = remaining.map((r, i) => {
    const x = pad.left + (n > 1 ? (i / (n - 1)) * chartW : chartW / 2);
    const y = pad.top + chartH - (r / yMax) * chartH;
    return `<circle cx="${x}" cy="${y}" r="4" fill="#3498db" stroke="#fff" stroke-width="2"/>`;
  }).join('');

  return `<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(data.plan.title)} - Burndown</title></head>
<body style="margin:0;padding:20px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#fafafa;">
  <h1 style="font-size:22px;margin:0 0 16px 0;">${escapeHtml(data.plan.title)} - Burndown Chart</h1>
  <svg width="${svgW}" height="${svgH}" style="background:#fff;border:1px solid #e0e0e0;border-radius:8px;">
    ${yLabels}
    <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + chartH}" stroke="#ccc" stroke-width="1"/>
    <line x1="${pad.left}" y1="${pad.top + chartH}" x2="${pad.left + chartW}" y2="${pad.top + chartH}" stroke="#ccc" stroke-width="1"/>
    <line x1="${idealStart.split(',')[0]}" y1="${idealStart.split(',')[1]}" x2="${idealEnd.split(',')[0]}" y2="${idealEnd.split(',')[1]}" stroke="#e74c3c" stroke-width="2" stroke-dasharray="6,4" opacity="0.5"/>
    <polyline points="${points.join(' ')}" fill="none" stroke="#3498db" stroke-width="2.5"/>
    ${circles}
    ${xLabels}
    <text x="${svgW / 2}" y="${svgH - 2}" text-anchor="middle" font-size="12" fill="#333">Date</text>
    <text x="14" y="${svgH / 2}" text-anchor="middle" font-size="12" fill="#333" transform="rotate(-90,14,${svgH / 2})">Remaining</text>
  </svg>
</body>
</html>`;
}
