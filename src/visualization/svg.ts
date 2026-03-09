/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * SVG renderers for Planner visualizations.
 */

import type { PlanVisualizationData } from './types.js';

function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * Renders a Kanban board as SVG.
 */
export function renderKanbanSvg(data: PlanVisualizationData): string {
  const colWidth = 220;
  const padding = 10;
  const headerHeight = 40;
  const cardHeight = 50;
  const cardGap = 5;

  const maxTasksPerBucket = Math.max(1, ...data.buckets.map(b =>
    data.tasks.filter(t => t.bucketId === b.id).length
  ));
  const svgWidth = data.buckets.length * (colWidth + padding) + padding;
  const svgHeight = headerHeight + maxTasksPerBucket * (cardHeight + cardGap) + padding * 2;

  let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${svgHeight}" viewBox="0 0 ${svgWidth} ${svgHeight}">`;
  svg += `<style>text{font-family:sans-serif;font-size:12px;}</style>`;

  data.buckets.forEach((bucket, colIdx) => {
    const x = padding + colIdx * (colWidth + padding);
    svg += `<rect x="${x}" y="${padding}" width="${colWidth}" height="${headerHeight}" rx="4" fill="#007bff"/>`;
    svg += `<text x="${x + colWidth / 2}" y="${padding + 25}" fill="white" text-anchor="middle" font-weight="bold">${escapeXml(bucket.name)}</text>`;

    const tasks = data.tasks.filter(t => t.bucketId === bucket.id);
    tasks.forEach((task, taskIdx) => {
      const ty = padding + headerHeight + cardGap + taskIdx * (cardHeight + cardGap);
      const fill = task.percentComplete === 100 ? '#d4edda' : task.percentComplete > 0 ? '#fff3cd' : '#f8f9fa';
      svg += `<rect x="${x}" y="${ty}" width="${colWidth}" height="${cardHeight}" rx="4" fill="${fill}" stroke="#dee2e6"/>`;
      svg += `<text x="${x + 8}" y="${ty + 20}">${escapeXml(task.title)}</text>`;
      svg += `<text x="${x + 8}" y="${ty + 38}" fill="#6c757d">${task.percentComplete}%</text>`;
    });
  });

  svg += '</svg>';
  return svg;
}

/**
 * Renders a Gantt chart as SVG.
 */
export function renderGanttSvg(data: PlanVisualizationData): string {
  const rowHeight = 30;
  const labelWidth = 200;
  const chartWidth = 400;
  const svgWidth = labelWidth + chartWidth + 20;
  const svgHeight = (data.tasks.length + 1) * rowHeight + 20;

  let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${svgHeight}" viewBox="0 0 ${svgWidth} ${svgHeight}">`;
  svg += `<style>text{font-family:sans-serif;font-size:11px;}</style>`;

  // Header
  svg += `<text x="10" y="20" font-weight="bold">Task</text>`;
  svg += `<text x="${labelWidth + 10}" y="20" font-weight="bold">Progress</text>`;

  data.tasks.forEach((task, idx) => {
    const y = (idx + 1) * rowHeight + 10;
    svg += `<text x="10" y="${y + 18}">${escapeXml(task.title)}</text>`;
    svg += `<rect x="${labelWidth}" y="${y}" width="${chartWidth}" height="${rowHeight - 4}" rx="2" fill="#e9ecef"/>`;
    const barWidth = Math.round((task.percentComplete / 100) * chartWidth);
    svg += `<rect x="${labelWidth}" y="${y}" width="${barWidth}" height="${rowHeight - 4}" rx="2" fill="#28a745"/>`;
    svg += `<text x="${labelWidth + 5}" y="${y + 18}" fill="${task.percentComplete > 10 ? 'white' : '#333'}">${task.percentComplete}%</text>`;
  });

  svg += '</svg>';
  return svg;
}

/**
 * Renders a plan summary as SVG.
 */
export function renderSummarySvg(data: PlanVisualizationData): string {
  const total = data.tasks.length;
  const completed = data.tasks.filter(t => t.percentComplete === 100).length;
  const inProgress = data.tasks.filter(t => t.percentComplete > 0 && t.percentComplete < 100).length;
  const notStarted = data.tasks.filter(t => t.percentComplete === 0).length;

  const svgWidth = 400;
  const svgHeight = 200;

  let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${svgHeight}" viewBox="0 0 ${svgWidth} ${svgHeight}">`;
  svg += `<style>text{font-family:sans-serif;font-size:14px;}</style>`;
  svg += `<text x="200" y="30" text-anchor="middle" font-size="18" font-weight="bold">${escapeXml(data.plan.title)}</text>`;

  const items = [
    { label: 'Total', value: total, color: '#007bff' },
    { label: 'Completed', value: completed, color: '#28a745' },
    { label: 'In Progress', value: inProgress, color: '#ffc107' },
    { label: 'Not Started', value: notStarted, color: '#6c757d' },
  ];

  items.forEach((item, idx) => {
    const x = 20 + idx * 95;
    const y = 60;
    svg += `<rect x="${x}" y="${y}" width="80" height="80" rx="8" fill="${item.color}"/>`;
    svg += `<text x="${x + 40}" y="${y + 40}" text-anchor="middle" fill="white" font-size="24" font-weight="bold">${item.value}</text>`;
    svg += `<text x="${x + 40}" y="${y + 60}" text-anchor="middle" fill="white" font-size="11">${item.label}</text>`;
  });

  svg += '</svg>';
  return svg;
}

/**
 * Renders a burndown chart as SVG.
 */
export function renderBurndownSvg(data: PlanVisualizationData): string {
  const total = data.tasks.length;
  const completed = data.tasks.filter(t => t.percentComplete === 100).length;
  const remaining = total - completed;

  const svgWidth = 300;
  const svgHeight = 200;
  const barMaxHeight = 140;

  let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${svgWidth}" height="${svgHeight}" viewBox="0 0 ${svgWidth} ${svgHeight}">`;
  svg += `<style>text{font-family:sans-serif;font-size:12px;}</style>`;
  svg += `<text x="150" y="20" text-anchor="middle" font-size="14" font-weight="bold">Burndown</text>`;

  const completedHeight = total > 0 ? Math.round((completed / total) * barMaxHeight) : 0;
  const remainingHeight = total > 0 ? Math.round((remaining / total) * barMaxHeight) : 0;
  const baseY = 170;

  svg += `<rect x="60" y="${baseY - completedHeight}" width="60" height="${completedHeight}" fill="#28a745" rx="4"/>`;
  svg += `<text x="90" y="${baseY + 15}" text-anchor="middle">Completed</text>`;
  svg += `<text x="90" y="${baseY - completedHeight - 5}" text-anchor="middle" font-weight="bold">${completed}</text>`;

  svg += `<rect x="180" y="${baseY - remainingHeight}" width="60" height="${remainingHeight}" fill="#dc3545" rx="4"/>`;
  svg += `<text x="210" y="${baseY + 15}" text-anchor="middle">Remaining</text>`;
  svg += `<text x="210" y="${baseY - remainingHeight - 5}" text-anchor="middle" font-weight="bold">${remaining}</text>`;

  svg += '</svg>';
  return svg;
}
