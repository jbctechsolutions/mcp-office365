/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Planner visualization tools.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import {
  PlannerVisualizationTools,
  type IPlannerVisualizationRepository,
} from '../../../src/tools/planner-visualization.js';
import type { PlanVisualizationData } from '../../../src/visualization/types.js';

const MOCK_DATA: PlanVisualizationData = {
  plan: {
    id: 1,
    title: 'Sprint 42',
    owner: 'group-abc',
    createdDateTime: '2026-01-01T00:00:00Z',
  },
  buckets: [
    { id: 10, name: 'To Do' },
    { id: 20, name: 'In Progress' },
    { id: 30, name: 'Done' },
  ],
  tasks: [
    {
      id: 100, title: 'Design UI', bucketId: 10, assignees: ['user1'],
      percentComplete: 0, priority: 5, startDateTime: '2026-03-01', dueDateTime: '2026-03-05', createdDateTime: '2026-02-28T00:00:00Z',
    },
    {
      id: 101, title: 'Build API', bucketId: 20, assignees: ['user2'],
      percentComplete: 50, priority: 3, startDateTime: '2026-03-01', dueDateTime: '2026-03-10', createdDateTime: '2026-02-28T00:00:00Z',
    },
    {
      id: 102, title: 'Write docs', bucketId: 30, assignees: ['user1'],
      percentComplete: 100, priority: 5, startDateTime: '2026-02-20', dueDateTime: '2026-02-25', createdDateTime: '2026-02-15T00:00:00Z',
    },
  ],
};

describe('PlannerVisualizationTools', () => {
  let repo: IPlannerVisualizationRepository;
  let tools: PlannerVisualizationTools;
  const writtenFiles: string[] = [];

  beforeEach(() => {
    repo = {
      getPlanVisualizationDataAsync: vi.fn().mockResolvedValue(MOCK_DATA),
    };
    tools = new PlannerVisualizationTools(repo);
  });

  afterEach(() => {
    // Clean up temp files
    for (const f of writtenFiles) {
      try { fs.unlinkSync(f); } catch { /* ignore */ }
    }
    writtenFiles.length = 0;
  });

  function parseResult(result: { content: Array<{ type: string; text: string }> }): {
    file_path: string; format: string; preview: string;
  } {
    expect(result.content).toHaveLength(1);
    const parsed = JSON.parse(result.content[0].text);
    if (parsed.file_path) writtenFiles.push(parsed.file_path);
    return parsed;
  }

  // ===========================================================================
  // generateKanbanBoard
  // ===========================================================================

  describe('generateKanbanBoard', () => {
    it('returns html format by default', async () => {
      const result = await tools.generateKanbanBoard({ plan_id: 1, format: 'html' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('html');
      expect(parsed.file_path).toMatch(/kanban.*\.html$/);
      expect(parsed.preview).toBeTruthy();
      expect(repo.getPlanVisualizationDataAsync).toHaveBeenCalledWith(1);
    });

    it('returns markdown format when requested', async () => {
      const result = await tools.generateKanbanBoard({ plan_id: 1, format: 'markdown' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('markdown');
      expect(parsed.file_path).toMatch(/kanban.*\.md$/);
    });

    it('returns svg format when requested', async () => {
      const result = await tools.generateKanbanBoard({ plan_id: 1, format: 'svg' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('svg');
      expect(parsed.file_path).toMatch(/kanban.*\.svg$/);
    });

    it('returns mermaid format when requested', async () => {
      const result = await tools.generateKanbanBoard({ plan_id: 1, format: 'mermaid' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('mermaid');
      expect(parsed.file_path).toMatch(/kanban.*\.md$/);
    });
  });

  // ===========================================================================
  // generateGanttChart
  // ===========================================================================

  describe('generateGanttChart', () => {
    it('returns html format by default', async () => {
      const result = await tools.generateGanttChart({ plan_id: 1, format: 'html' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('html');
      expect(parsed.file_path).toMatch(/gantt.*\.html$/);
      expect(parsed.preview).toBeTruthy();
      expect(repo.getPlanVisualizationDataAsync).toHaveBeenCalledWith(1);
    });

    it('returns markdown format when requested', async () => {
      const result = await tools.generateGanttChart({ plan_id: 1, format: 'markdown' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('markdown');
      expect(parsed.file_path).toMatch(/gantt.*\.md$/);
    });

    it('returns svg format when requested', async () => {
      const result = await tools.generateGanttChart({ plan_id: 1, format: 'svg' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('svg');
      expect(parsed.file_path).toMatch(/gantt.*\.svg$/);
    });

    it('returns mermaid format when requested', async () => {
      const result = await tools.generateGanttChart({ plan_id: 1, format: 'mermaid' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('mermaid');
    });
  });

  // ===========================================================================
  // generatePlanSummary
  // ===========================================================================

  describe('generatePlanSummary', () => {
    it('returns html format by default', async () => {
      const result = await tools.generatePlanSummary({ plan_id: 1, format: 'html' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('html');
      expect(parsed.file_path).toMatch(/summary.*\.html$/);
      expect(parsed.preview).toBeTruthy();
      expect(repo.getPlanVisualizationDataAsync).toHaveBeenCalledWith(1);
    });

    it('returns markdown format when requested', async () => {
      const result = await tools.generatePlanSummary({ plan_id: 1, format: 'markdown' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('markdown');
      expect(parsed.file_path).toMatch(/summary.*\.md$/);
    });

    it('returns svg format when requested', async () => {
      const result = await tools.generatePlanSummary({ plan_id: 1, format: 'svg' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('svg');
      expect(parsed.file_path).toMatch(/summary.*\.svg$/);
    });

    it('returns mermaid format when requested', async () => {
      const result = await tools.generatePlanSummary({ plan_id: 1, format: 'mermaid' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('mermaid');
    });
  });

  // ===========================================================================
  // generateBurndownChart
  // ===========================================================================

  describe('generateBurndownChart', () => {
    it('returns html format by default', async () => {
      const result = await tools.generateBurndownChart({ plan_id: 1, format: 'html' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('html');
      expect(parsed.file_path).toMatch(/burndown.*\.html$/);
      expect(parsed.preview).toBeTruthy();
      expect(repo.getPlanVisualizationDataAsync).toHaveBeenCalledWith(1);
    });

    it('returns markdown format when requested', async () => {
      const result = await tools.generateBurndownChart({ plan_id: 1, format: 'markdown' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('markdown');
      expect(parsed.file_path).toMatch(/burndown.*\.md$/);
    });

    it('returns svg format when requested', async () => {
      const result = await tools.generateBurndownChart({ plan_id: 1, format: 'svg' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('svg');
      expect(parsed.file_path).toMatch(/burndown.*\.svg$/);
    });

    it('returns mermaid format when requested', async () => {
      const result = await tools.generateBurndownChart({ plan_id: 1, format: 'mermaid' });
      const parsed = parseResult(result);
      expect(parsed.format).toBe('mermaid');
    });
  });
});
