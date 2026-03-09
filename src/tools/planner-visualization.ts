/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Planner visualization MCP tools.
 *
 * Provides tools for generating visual representations of Planner plans:
 * Kanban boards, Gantt charts, plan summaries, and burndown charts.
 */

import { z } from 'zod';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import type { PlanVisualizationData, VisualizationFormat } from '../visualization/types.js';
import { renderKanbanMarkdown, renderGanttMarkdown, renderSummaryMarkdown, renderBurndownMarkdown } from '../visualization/markdown.js';
import { renderKanbanMermaid, renderGanttMermaid, renderSummaryMermaid, renderBurndownMermaid } from '../visualization/mermaid.js';
import { renderKanbanHtml, renderGanttHtml, renderSummaryHtml, renderBurndownHtml } from '../visualization/html.js';
import { renderKanbanSvg, renderGanttSvg, renderSummarySvg, renderBurndownSvg } from '../visualization/svg.js';

// =============================================================================
// Input Schemas
// =============================================================================

export const GenerateKanbanBoardInput = z.strictObject({
  plan_id: z.number().int().positive().describe('Plan ID from list_plans'),
  format: z.enum(['html', 'svg', 'markdown', 'mermaid']).default('html').describe('Output format'),
  output_path: z.string().optional().describe('Custom file path for output (default: temp directory)'),
});

export const GenerateGanttChartInput = z.strictObject({
  plan_id: z.number().int().positive().describe('Plan ID from list_plans'),
  format: z.enum(['html', 'svg', 'markdown', 'mermaid']).default('html').describe('Output format'),
  output_path: z.string().optional().describe('Custom file path for output (default: temp directory)'),
});

export const GeneratePlanSummaryInput = z.strictObject({
  plan_id: z.number().int().positive().describe('Plan ID from list_plans'),
  format: z.enum(['html', 'svg', 'markdown', 'mermaid']).default('html').describe('Output format'),
  output_path: z.string().optional().describe('Custom file path for output (default: temp directory)'),
});

export const GenerateBurndownChartInput = z.strictObject({
  plan_id: z.number().int().positive().describe('Plan ID from list_plans'),
  format: z.enum(['html', 'svg', 'markdown', 'mermaid']).default('html').describe('Output format'),
  output_path: z.string().optional().describe('Custom file path for output (default: temp directory)'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type GenerateKanbanBoardParams = z.infer<typeof GenerateKanbanBoardInput>;
export type GenerateGanttChartParams = z.infer<typeof GenerateGanttChartInput>;
export type GeneratePlanSummaryParams = z.infer<typeof GeneratePlanSummaryInput>;
export type GenerateBurndownChartParams = z.infer<typeof GenerateBurndownChartInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface IPlannerVisualizationRepository {
  getPlanVisualizationDataAsync(planId: number): Promise<PlanVisualizationData>;
}

// =============================================================================
// Helpers
// =============================================================================

const FORMAT_EXTENSIONS: Record<VisualizationFormat, string> = {
  html: '.html',
  svg: '.svg',
  markdown: '.md',
  mermaid: '.md',
};

function writeOutput(content: string, format: VisualizationFormat, baseName: string, outputPath?: string): string {
  const ext = FORMAT_EXTENSIONS[format];
  const filePath = outputPath ?? path.join(os.tmpdir(), `${baseName}-${Date.now()}${ext}`);
  fs.writeFileSync(filePath, content, 'utf-8');
  return filePath;
}

function getPreview(content: string, maxLength: number = 500): string {
  return content.length > maxLength ? content.slice(0, maxLength) + '...' : content;
}

// =============================================================================
// Planner Visualization Tools
// =============================================================================

/**
 * Tools for generating visual representations of Planner plans.
 */
export class PlannerVisualizationTools {
  constructor(
    private readonly repo: IPlannerVisualizationRepository,
  ) {}

  async generateKanbanBoard(params: GenerateKanbanBoardParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const data = await this.repo.getPlanVisualizationDataAsync(params.plan_id);
    const format = params.format;

    let content: string;
    switch (format) {
      case 'html': content = renderKanbanHtml(data); break;
      case 'svg': content = renderKanbanSvg(data); break;
      case 'markdown': content = renderKanbanMarkdown(data); break;
      case 'mermaid': content = renderKanbanMermaid(data); break;
    }

    const filePath = writeOutput(content, format, 'kanban', params.output_path);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          file_path: filePath,
          format,
          preview: getPreview(content),
        }, null, 2),
      }],
    };
  }

  async generateGanttChart(params: GenerateGanttChartParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const data = await this.repo.getPlanVisualizationDataAsync(params.plan_id);
    const format = params.format;

    let content: string;
    switch (format) {
      case 'html': content = renderGanttHtml(data); break;
      case 'svg': content = renderGanttSvg(data); break;
      case 'markdown': content = renderGanttMarkdown(data); break;
      case 'mermaid': content = renderGanttMermaid(data); break;
    }

    const filePath = writeOutput(content, format, 'gantt', params.output_path);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          file_path: filePath,
          format,
          preview: getPreview(content),
        }, null, 2),
      }],
    };
  }

  async generatePlanSummary(params: GeneratePlanSummaryParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const data = await this.repo.getPlanVisualizationDataAsync(params.plan_id);
    const format = params.format;

    let content: string;
    switch (format) {
      case 'html': content = renderSummaryHtml(data); break;
      case 'svg': content = renderSummarySvg(data); break;
      case 'markdown': content = renderSummaryMarkdown(data); break;
      case 'mermaid': content = renderSummaryMermaid(data); break;
    }

    const filePath = writeOutput(content, format, 'summary', params.output_path);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          file_path: filePath,
          format,
          preview: getPreview(content),
        }, null, 2),
      }],
    };
  }

  async generateBurndownChart(params: GenerateBurndownChartParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const data = await this.repo.getPlanVisualizationDataAsync(params.plan_id);
    const format = params.format;

    let content: string;
    switch (format) {
      case 'html': content = renderBurndownHtml(data); break;
      case 'svg': content = renderBurndownSvg(data); break;
      case 'markdown': content = renderBurndownMarkdown(data); break;
      case 'mermaid': content = renderBurndownMermaid(data); break;
    }

    const filePath = writeOutput(content, format, 'burndown', params.output_path);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          file_path: filePath,
          format,
          preview: getPreview(content),
        }, null, 2),
      }],
    };
  }
}
