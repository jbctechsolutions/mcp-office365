/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Mermaid visualization renderer.
 */

import { describe, it, expect } from 'vitest';
import type { PlanVisualizationData } from '../../../src/visualization/types.js';
import {
  renderGanttMermaid,
  renderKanbanMermaid,
  renderSummaryMermaid,
  renderBurndownMermaid,
} from '../../../src/visualization/mermaid.js';

function makeSampleData(): PlanVisualizationData {
  return {
    plan: { id: 1, title: 'Sprint 1' },
    buckets: [
      { id: 10, name: 'To Do', orderHint: '1' },
      { id: 20, name: 'In Progress', orderHint: '2' },
    ],
    tasks: [
      {
        id: 101,
        title: 'Setup project',
        bucketId: 10,
        percentComplete: 0,
        priority: 5,
        startDateTime: '2026-01-01T00:00:00Z',
        dueDateTime: '2026-01-10T00:00:00Z',
        assignments: ['Alice'],
        completedDateTime: null,
      },
      {
        id: 102,
        title: 'Build API',
        bucketId: 20,
        percentComplete: 50,
        priority: 1,
        startDateTime: '2026-01-05T00:00:00Z',
        dueDateTime: '2026-01-15T00:00:00Z',
        assignments: ['Bob'],
        completedDateTime: null,
      },
      {
        id: 103,
        title: 'Write docs',
        bucketId: 10,
        percentComplete: 100,
        priority: 9,
        startDateTime: null,
        dueDateTime: '2026-01-08T00:00:00Z',
        assignments: [],
        completedDateTime: '2026-01-07T00:00:00Z',
      },
    ],
  };
}

function makeEmptyData(): PlanVisualizationData {
  return {
    plan: { id: 1, title: 'Empty Plan' },
    buckets: [{ id: 10, name: 'Backlog', orderHint: '1' }],
    tasks: [],
  };
}

function makeSingleBucketData(): PlanVisualizationData {
  return {
    plan: { id: 1, title: 'Single Bucket' },
    buckets: [{ id: 10, name: 'All Tasks', orderHint: '1' }],
    tasks: [
      {
        id: 201,
        title: 'Only task',
        bucketId: 10,
        percentComplete: 25,
        priority: 3,
        startDateTime: '2026-02-01T00:00:00Z',
        dueDateTime: '2026-02-10T00:00:00Z',
        assignments: ['Dave'],
        completedDateTime: null,
      },
    ],
  };
}

describe('Mermaid Renderer', () => {
  describe('renderGanttMermaid', () => {
    it('produces valid gantt syntax', () => {
      const result = renderGanttMermaid(makeSampleData());
      expect(result).toMatch(/^gantt/);
      expect(result).toContain('dateFormat YYYY-MM-DD');
      expect(result).toContain('title Sprint 1');
    });

    it('creates sections per bucket', () => {
      const result = renderGanttMermaid(makeSampleData());
      expect(result).toContain('section To Do');
      expect(result).toContain('section In Progress');
    });

    it('marks completed and active tasks', () => {
      const result = renderGanttMermaid(makeSampleData());
      expect(result).toContain('Write docs :done,');
      expect(result).toContain('Build API :active,');
    });

    it('includes date ranges for tasks', () => {
      const result = renderGanttMermaid(makeSampleData());
      expect(result).toContain('2026-01-01');
      expect(result).toContain('2026-01-10');
    });
  });

  describe('renderKanbanMermaid', () => {
    it('uses block-beta syntax', () => {
      const result = renderKanbanMermaid(makeSampleData());
      expect(result).toMatch(/^block-beta/);
      expect(result).toContain('columns 2');
    });

    it('includes bucket names', () => {
      const result = renderKanbanMermaid(makeSampleData());
      expect(result).toContain('To_Do');
      expect(result).toContain('In_Progress');
    });

    it('includes task titles with percentages', () => {
      const result = renderKanbanMermaid(makeSampleData());
      expect(result).toContain('Setup project (0%)');
      expect(result).toContain('Build API (50%)');
    });

    it('handles single bucket', () => {
      const result = renderKanbanMermaid(makeSingleBucketData());
      expect(result).toContain('columns 1');
      expect(result).toContain('Only task (25%)');
    });
  });

  describe('renderSummaryMermaid', () => {
    it('produces valid pie chart syntax', () => {
      const result = renderSummaryMermaid(makeSampleData());
      expect(result).toMatch(/^pie/);
      expect(result).toContain('title Sprint 1');
    });

    it('includes task status categories', () => {
      const result = renderSummaryMermaid(makeSampleData());
      expect(result).toContain('"Completed"');
      // Tasks with past due dates and not completed are classified as Overdue
      expect(result).toContain('"Overdue"');
    });

    it('handles empty tasks with placeholder', () => {
      const result = renderSummaryMermaid(makeEmptyData());
      expect(result).toContain('"No Tasks"');
    });
  });

  describe('renderBurndownMermaid', () => {
    it('produces valid xychart-beta syntax', () => {
      const result = renderBurndownMermaid(makeSampleData());
      expect(result).toMatch(/^xychart-beta/);
      expect(result).toContain('title Sprint 1 - Burndown');
    });

    it('includes x-axis date labels', () => {
      const result = renderBurndownMermaid(makeSampleData());
      expect(result).toContain('x-axis');
      expect(result).toContain('2026-01');
    });

    it('includes y-axis and line data', () => {
      const result = renderBurndownMermaid(makeSampleData());
      expect(result).toContain('y-axis');
      expect(result).toContain('line [');
    });

    it('handles tasks with no dates', () => {
      const noDateData: PlanVisualizationData = {
        plan: { id: 1, title: 'No Dates' },
        buckets: [{ id: 10, name: 'Backlog', orderHint: '1' }],
        tasks: [
          {
            id: 301,
            title: 'No date task',
            bucketId: 10,
            percentComplete: 0,
            priority: 5,
            assignments: [],
          },
        ],
      };
      const result = renderBurndownMermaid(noDateData);
      expect(result).toContain('No date data available');
    });
  });
});
