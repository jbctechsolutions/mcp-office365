/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for SVG visualization renderer.
 */

import { describe, it, expect } from 'vitest';
import type { PlanVisualizationData } from '../../../src/visualization/types.js';
import {
  renderKanbanSvg,
  renderGanttSvg,
  renderSummarySvg,
  renderBurndownSvg,
} from '../../../src/visualization/svg.js';

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
        startDateTime: '2026-01-01T00:00:00Z',
        dueDateTime: '2026-01-08T00:00:00Z',
        assignments: ['Alice'],
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

function makeNoDatesData(): PlanVisualizationData {
  return {
    plan: { id: 1, title: 'No Dates' },
    buckets: [{ id: 10, name: 'Backlog', orderHint: '1' }],
    tasks: [
      {
        id: 201,
        title: 'Undated task',
        bucketId: 10,
        percentComplete: 0,
        priority: 5,
        assignments: [],
      },
    ],
  };
}

function makeSingleBucketData(): PlanVisualizationData {
  return {
    plan: { id: 1, title: 'Single Bucket' },
    buckets: [{ id: 10, name: 'All Tasks', orderHint: '1' }],
    tasks: [
      {
        id: 301,
        title: 'Only task',
        bucketId: 10,
        percentComplete: 75,
        priority: 3,
        startDateTime: '2026-02-01T00:00:00Z',
        dueDateTime: '2026-02-10T00:00:00Z',
        assignments: ['Dave'],
        completedDateTime: null,
      },
    ],
  };
}

describe('SVG Renderer', () => {
  describe('renderKanbanSvg', () => {
    it('produces valid SVG element', () => {
      const result = renderKanbanSvg(makeSampleData());
      expect(result).toContain('<svg');
      expect(result).toContain('xmlns="http://www.w3.org/2000/svg"');
      expect(result).toContain('</svg>');
    });

    it('renders bucket column backgrounds', () => {
      const result = renderKanbanSvg(makeSampleData());
      // Column backgrounds are <rect> with fill
      expect(result).toContain('<rect');
      expect(result).toContain('fill="#f4f5f7"');
    });

    it('renders bucket labels', () => {
      const result = renderKanbanSvg(makeSampleData());
      expect(result).toContain('To Do');
      expect(result).toContain('In Progress');
    });

    it('renders task cards with rounded corners', () => {
      const result = renderKanbanSvg(makeSampleData());
      expect(result).toContain('rx="6"');
    });

    it('uses priority colors for card indicators', () => {
      const result = renderKanbanSvg(makeSampleData());
      // Urgent red
      expect(result).toContain('#e74c3c');
      // Medium yellow
      expect(result).toContain('#f1c40f');
      // Low green
      expect(result).toContain('#2ecc71');
    });

    it('includes task titles', () => {
      const result = renderKanbanSvg(makeSampleData());
      expect(result).toContain('Setup project');
      expect(result).toContain('Build API');
      expect(result).toContain('Write docs');
    });

    it('handles single bucket', () => {
      const result = renderKanbanSvg(makeSingleBucketData());
      expect(result).toContain('All Tasks');
      expect(result).toContain('Only task');
    });

    it('handles empty tasks', () => {
      const result = renderKanbanSvg(makeEmptyData());
      expect(result).toContain('<svg');
      expect(result).toContain('Backlog');
    });
  });

  describe('renderGanttSvg', () => {
    it('produces valid SVG with task bars', () => {
      const result = renderGanttSvg(makeSampleData());
      expect(result).toContain('<svg');
      expect(result).toContain('<rect');
      expect(result).toContain('</svg>');
    });

    it('renders task labels', () => {
      const result = renderGanttSvg(makeSampleData());
      expect(result).toContain('Setup project');
      expect(result).toContain('Build API');
    });

    it('renders grid lines', () => {
      const result = renderGanttSvg(makeSampleData());
      expect(result).toContain('<line');
    });

    it('renders date axis labels', () => {
      const result = renderGanttSvg(makeSampleData());
      expect(result).toContain('2026-01');
    });

    it('handles tasks with no dates', () => {
      const result = renderGanttSvg(makeNoDatesData());
      expect(result).toContain('No tasks with dates');
    });
  });

  describe('renderSummarySvg', () => {
    it('produces valid SVG with pie chart', () => {
      const result = renderSummarySvg(makeSampleData());
      expect(result).toContain('<svg');
      expect(result).toContain('</svg>');
    });

    it('includes pie chart arcs or circle', () => {
      const result = renderSummarySvg(makeSampleData());
      // Should have path elements for pie arcs
      expect(result).toContain('<path');
    });

    it('includes legend with status labels', () => {
      const result = renderSummarySvg(makeSampleData());
      // Tasks with past due dates are classified as Overdue
      expect(result).toContain('Overdue');
      expect(result).toContain('Completed');
    });

    it('shows total count', () => {
      const result = renderSummarySvg(makeSampleData());
      expect(result).toContain('Total: 3');
    });

    it('handles empty tasks with placeholder', () => {
      const result = renderSummarySvg(makeEmptyData());
      expect(result).toContain('No Tasks');
    });
  });

  describe('renderBurndownSvg', () => {
    it('produces valid SVG with polyline and circles', () => {
      const result = renderBurndownSvg(makeSampleData());
      expect(result).toContain('<svg');
      expect(result).toContain('<polyline');
      expect(result).toContain('<circle');
      expect(result).toContain('</svg>');
    });

    it('includes ideal burndown line', () => {
      const result = renderBurndownSvg(makeSampleData());
      expect(result).toContain('stroke-dasharray');
      expect(result).toContain('#e74c3c');
    });

    it('includes axis lines', () => {
      const result = renderBurndownSvg(makeSampleData());
      expect(result).toContain('<line');
    });

    it('includes date labels', () => {
      const result = renderBurndownSvg(makeSampleData());
      expect(result).toContain('2026-01');
    });

    it('includes axis labels', () => {
      const result = renderBurndownSvg(makeSampleData());
      expect(result).toContain('Date');
      expect(result).toContain('Remaining Tasks');
    });

    it('includes title', () => {
      const result = renderBurndownSvg(makeSampleData());
      expect(result).toContain('Sprint 1 - Burndown');
    });

    it('handles no date data', () => {
      const result = renderBurndownSvg(makeNoDatesData());
      expect(result).toContain('No date data available');
    });
  });
});
