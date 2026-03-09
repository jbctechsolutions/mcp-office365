/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Pure SVG visualization renderer for Planner data.
 */

import {
  type PlanVisualizationData,
  PRIORITY_LABELS,
  PRIORITY_COLORS,
  getTaskStatus,
} from './types.js';

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function priorityColor(priority: number): string {
  return PRIORITY_COLORS[priority] ?? '#95a5a6';
}

/**
 * Render a Kanban board as pure SVG.
 */
export function renderKanbanSvg(data: PlanVisualizationData): string {
  const colWidth = 240;
  const cardHeight = 60;
  const cardGap = 8;
  const headerHeight = 40;
  const padding = 16;
  const colCount = Math.max(data.buckets.length, 1);
  const svgWidth = colCount * (colWidth + padding) + padding;

  const tasksByBucket = data.buckets.map((b) =>
    data.tasks.filter((t) => t.bucketId === b.id)
  );
  const maxCards = Math.max(...tasksByBucket.map((ts) => ts.length), 1);
  const svgHeight = headerHeight + padding + maxCards * (cardHeight + cardGap) + padding;

  let elements = '';

  for (let col = 0; col < data.buckets.length; col++) {
    const bucket = data.buckets[col]!;
    const tasks = tasksByBucket[col]!;
    const x = padding + col * (colWidth + padding);

    // Column background
    elements += `<rect x="${x}" y="0" width="${colWidth}" height="${svgHeight}" rx="8" fill="#f4f5f7"/>`;

    // Bucket label
    elements += `<text x="${x + colWidth / 2}" y="26" text-anchor="middle" font-size="14" font-weight="bold" fill="#333">${escapeXml(bucket.name)}</text>`;

    // Task cards
    for (let i = 0; i < tasks.length; i++) {
      const task = tasks[i]!;
      const cy = headerHeight + padding + i * (cardHeight + cardGap);
      const pColor = priorityColor(task.priority);
      const pLabel = PRIORITY_LABELS[task.priority] ?? `P${task.priority}`;

      // Card background
      elements += `<rect x="${x + 8}" y="${cy}" width="${colWidth - 16}" height="${cardHeight}" rx="6" fill="#fff" stroke="#e0e0e0" stroke-width="1"/>`;
      // Priority indicator bar
      elements += `<rect x="${x + 8}" y="${cy}" width="4" height="${cardHeight}" rx="2" fill="${pColor}"/>`;
      // Task title (truncated)
      const displayTitle = task.title.length > 28 ? task.title.slice(0, 26) + '...' : task.title;
      elements += `<text x="${x + 20}" y="${cy + 22}" font-size="12" font-weight="600" fill="#333">${escapeXml(displayTitle)}</text>`;
      // Priority and percent
      elements += `<text x="${x + 20}" y="${cy + 42}" font-size="10" fill="#888">${escapeXml(pLabel)} | ${task.percentComplete}%</text>`;
    }
  }

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${svgHeight}" style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;">
  ${elements}
</svg>`;
}

/**
 * Render a Gantt chart as pure SVG.
 */
export function renderGanttSvg(data: PlanVisualizationData): string {
  const labelWidth = 180;
  const chartWidth = 600;
  const rowHeight = 32;
  const headerHeight = 40;
  const padding = 10;
  const svgWidth = labelWidth + chartWidth + padding * 2;

  const allTasks = data.tasks.filter((t) => t.startDateTime || t.dueDateTime);
  if (allTasks.length === 0 && data.tasks.length > 0) {
    // Use all tasks with default dates if none have dates
    allTasks.push(...data.tasks);
  }

  const dates: number[] = [];
  for (const task of allTasks) {
    if (task.startDateTime) dates.push(new Date(task.startDateTime).getTime());
    if (task.dueDateTime) dates.push(new Date(task.dueDateTime).getTime());
  }

  if (dates.length === 0) {
    return `<svg xmlns="http://www.w3.org/2000/svg" width="400" height="60" style="font-family:sans-serif;">
  <text x="200" y="35" text-anchor="middle" font-size="14" fill="#666">No tasks with dates to display</text>
</svg>`;
  }

  const minTime = Math.min(...dates) - 86400000;
  const maxTime = Math.max(...dates) + 86400000;
  const totalRange = Math.max(maxTime - minTime, 1);
  const svgHeight = headerHeight + allTasks.length * rowHeight + padding;

  let elements = '';

  // Grid lines and date axis
  const dateCount = Math.min(8, allTasks.length + 2);
  for (let i = 0; i <= dateCount; i++) {
    const x = labelWidth + (i / dateCount) * chartWidth;
    const dateVal = new Date(minTime + (i / dateCount) * totalRange);
    const dateStr = dateVal.toISOString().split('T')[0]!;
    elements += `<line x1="${x}" y1="${headerHeight}" x2="${x}" y2="${svgHeight}" stroke="#eee" stroke-width="1"/>`;
    elements += `<text x="${x}" y="20" text-anchor="middle" font-size="10" fill="#666">${dateStr}</text>`;
  }

  // Task bars
  for (let i = 0; i < allTasks.length; i++) {
    const task = allTasks[i]!;
    const y = headerHeight + i * rowHeight;
    const pColor = priorityColor(task.priority);

    const startTime = task.startDateTime
      ? new Date(task.startDateTime).getTime()
      : task.dueDateTime
        ? new Date(task.dueDateTime).getTime() - 7 * 86400000
        : minTime;
    const endTime = task.dueDateTime
      ? new Date(task.dueDateTime).getTime()
      : startTime + 7 * 86400000;

    const barX = labelWidth + ((startTime - minTime) / totalRange) * chartWidth;
    const barW = Math.max(8, ((endTime - startTime) / totalRange) * chartWidth);
    const progressW = (task.percentComplete / 100) * barW;

    // Label
    const displayTitle = task.title.length > 22 ? task.title.slice(0, 20) + '...' : task.title;
    elements += `<text x="${labelWidth - 8}" y="${y + rowHeight / 2 + 4}" text-anchor="end" font-size="12" fill="#333">${escapeXml(displayTitle)}</text>`;
    // Bar background
    elements += `<rect x="${barX}" y="${y + 6}" width="${barW}" height="${rowHeight - 14}" rx="3" fill="#ddd"/>`;
    // Progress fill
    elements += `<rect x="${barX}" y="${y + 6}" width="${progressW}" height="${rowHeight - 14}" rx="3" fill="${pColor}"/>`;
  }

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${svgHeight}" style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;">
  <text x="${svgWidth / 2}" y="36" text-anchor="middle" font-size="11" fill="#999">Date Axis</text>
  ${elements}
</svg>`;
}

/**
 * Render a summary as SVG pie chart with stats.
 */
export function renderSummarySvg(data: PlanVisualizationData): string {
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
  const svgWidth = 500;
  const svgHeight = 350;
  const cx = 160;
  const cy = 170;
  const r = 100;

  const segments: Array<{ label: string; value: number; color: string }> = [
    { label: 'Not Started', value: notStarted, color: '#95a5a6' },
    { label: 'In Progress', value: inProgress, color: '#3498db' },
    { label: 'Completed', value: completed, color: '#2ecc71' },
    { label: 'Overdue', value: overdue, color: '#e74c3c' },
  ].filter((s) => s.value > 0);

  let paths = '';
  if (segments.length === 0) {
    // Empty state - full circle
    paths = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="#eee"/>`;
    paths += `<text x="${cx}" y="${cy + 4}" text-anchor="middle" font-size="14" fill="#999">No Tasks</text>`;
  } else if (segments.length === 1) {
    const seg = segments[0]!;
    paths = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${seg.color}"/>`;
  } else {
    let startAngle = -Math.PI / 2;
    for (const seg of segments) {
      const sliceAngle = (seg.value / total) * 2 * Math.PI;
      const endAngle = startAngle + sliceAngle;
      const largeArc = sliceAngle > Math.PI ? 1 : 0;
      const x1 = cx + r * Math.cos(startAngle);
      const y1 = cy + r * Math.sin(startAngle);
      const x2 = cx + r * Math.cos(endAngle);
      const y2 = cy + r * Math.sin(endAngle);
      paths += `<path d="M${cx},${cy} L${x1},${y1} A${r},${r} 0 ${largeArc},1 ${x2},${y2} Z" fill="${seg.color}"/>`;
      startAngle = endAngle;
    }
  }

  // Legend
  let legend = '';
  const legendX = 300;
  let legendY = 80;
  for (const seg of segments) {
    legend += `<rect x="${legendX}" y="${legendY - 10}" width="14" height="14" rx="2" fill="${seg.color}"/>`;
    legend += `<text x="${legendX + 20}" y="${legendY + 1}" font-size="13" fill="#333">${escapeXml(seg.label)}: ${seg.value}</text>`;
    legendY += 24;
  }

  // Stats text
  const statsY = legendY + 20;
  legend += `<text x="${legendX}" y="${statsY}" font-size="14" font-weight="bold" fill="#333">Total: ${total}</text>`;

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${svgHeight}" style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;">
  <text x="${svgWidth / 2}" y="30" text-anchor="middle" font-size="18" font-weight="bold" fill="#333">${escapeXml(data.plan.title)} - Summary</text>
  ${paths}
  ${legend}
</svg>`;
}

/**
 * Render a burndown chart as pure SVG line chart.
 */
export function renderBurndownSvg(data: PlanVisualizationData): string {
  const allDates = new Set<string>();
  for (const task of data.tasks) {
    if (task.startDateTime) allDates.add(new Date(task.startDateTime).toISOString().split('T')[0]!);
    if (task.dueDateTime) allDates.add(new Date(task.dueDateTime).toISOString().split('T')[0]!);
    if (task.completedDateTime) allDates.add(new Date(task.completedDateTime).toISOString().split('T')[0]!);
  }

  const sortedDates = [...allDates].sort();
  if (sortedDates.length === 0) {
    return `<svg xmlns="http://www.w3.org/2000/svg" width="400" height="60" style="font-family:sans-serif;">
  <text x="200" y="35" text-anchor="middle" font-size="14" fill="#666">No date data available</text>
</svg>`;
  }

  const total = data.tasks.length;
  const remaining: number[] = [];
  for (const date of sortedDates) {
    const completedByDate = data.tasks.filter(
      (t) => t.completedDateTime && new Date(t.completedDateTime).toISOString().split('T')[0]! <= date
    ).length;
    remaining.push(total - completedByDate);
  }

  const svgW = 700;
  const svgH = 380;
  const pad = { top: 50, right: 30, bottom: 70, left: 55 };
  const chartW = svgW - pad.left - pad.right;
  const chartH = svgH - pad.top - pad.bottom;
  const n = sortedDates.length;
  const yMax = Math.max(total, 1);

  let elements = '';

  // Grid and Y-axis
  const yTicks = 5;
  for (let i = 0; i <= yTicks; i++) {
    const val = Math.round((i / yTicks) * yMax);
    const y = pad.top + chartH - (i / yTicks) * chartH;
    elements += `<line x1="${pad.left}" y1="${y}" x2="${pad.left + chartW}" y2="${y}" stroke="#eee" stroke-width="1"/>`;
    elements += `<text x="${pad.left - 8}" y="${y + 4}" text-anchor="end" font-size="11" fill="#666">${val}</text>`;
  }

  // Axes
  elements += `<line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + chartH}" stroke="#ccc" stroke-width="1"/>`;
  elements += `<line x1="${pad.left}" y1="${pad.top + chartH}" x2="${pad.left + chartW}" y2="${pad.top + chartH}" stroke="#ccc" stroke-width="1"/>`;

  // X-axis labels
  const labelCount = Math.min(n, 8);
  for (let i = 0; i < labelCount; i++) {
    const idx = Math.floor(i * (n - 1) / Math.max(labelCount - 1, 1));
    const x = pad.left + (n > 1 ? (idx / (n - 1)) * chartW : chartW / 2);
    elements += `<text x="${x}" y="${pad.top + chartH + 20}" text-anchor="middle" font-size="10" fill="#666" transform="rotate(-30,${x},${pad.top + chartH + 20})">${sortedDates[idx]}</text>`;
  }

  // Ideal burndown line
  elements += `<line x1="${pad.left}" y1="${pad.top}" x2="${pad.left + chartW}" y2="${pad.top + chartH}" stroke="#e74c3c" stroke-width="2" stroke-dasharray="6,4" opacity="0.5"/>`;

  // Actual burndown line
  const points = remaining.map((r, i) => {
    const x = pad.left + (n > 1 ? (i / (n - 1)) * chartW : chartW / 2);
    const y = pad.top + chartH - (r / yMax) * chartH;
    return `${x},${y}`;
  });
  elements += `<polyline points="${points.join(' ')}" fill="none" stroke="#3498db" stroke-width="2.5"/>`;

  // Data point circles
  for (let i = 0; i < remaining.length; i++) {
    const x = pad.left + (n > 1 ? (i / (n - 1)) * chartW : chartW / 2);
    const y = pad.top + chartH - (remaining[i]! / yMax) * chartH;
    elements += `<circle cx="${x}" cy="${y}" r="4" fill="#3498db" stroke="#fff" stroke-width="2"/>`;
  }

  // Axis labels
  elements += `<text x="${svgW / 2}" y="${svgH - 5}" text-anchor="middle" font-size="12" fill="#333">Date</text>`;
  elements += `<text x="14" y="${svgH / 2}" text-anchor="middle" font-size="12" fill="#333" transform="rotate(-90,14,${svgH / 2})">Remaining Tasks</text>`;

  // Title
  elements += `<text x="${svgW / 2}" y="28" text-anchor="middle" font-size="18" font-weight="bold" fill="#333">${escapeXml(data.plan.title)} - Burndown</text>`;

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${svgW}" height="${svgH}" style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;">
  ${elements}
</svg>`;
}
