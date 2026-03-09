/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for HTML visualization renderer.
 */

import { describe, it, expect } from 'vitest';
import type { PlanVisualizationData } from '../../../src/visualization/types.js';
import {
  renderKanbanHtml,
  renderGanttHtml,
  renderSummaryHtml,
  renderBurndownHtml,
} from '../../../src/visualization/html.js';

function makeSampleData(): PlanVisualizationData {
  return {
    plan: { id: 1, title: 'Sprint 1' },
    buckets: [
      { id: 10, name: 'To Do', orderHint: '1' },
      { id: 20, name: 'In Progress', orderHint: '2' },
      { id: 30, name: 'Done', orderHint: '3' },
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
        assignments: ['Bob', 'Carol'],
        completedDateTime: null,
      },
      {
        id: 103,
        title: 'Write docs',
        bucketId: 30,
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
    buckets: [{ id: 10, name: 'To Do', orderHint: '1' }],
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

describe('HTML Renderer', () => {
  describe('renderKanbanHtml', () => {
    it('produces valid HTML document', () => {
      const result = renderKanbanHtml(makeSampleData());
      expect(result).toContain('<!DOCTYPE html>');
      expect(result).toContain('<html');
      expect(result).toContain('</html>');
    });

    it('contains CSS grid layout', () => {
      const result = renderKanbanHtml(makeSampleData());
      expect(result).toContain('display:grid');
      expect(result).toContain('grid-template-columns');
    });

    it('renders bucket columns', () => {
      const result = renderKanbanHtml(makeSampleData());
      expect(result).toContain('To Do');
      expect(result).toContain('In Progress');
      expect(result).toContain('Done');
    });

    it('renders task cards with priority colors', () => {
      const result = renderKanbanHtml(makeSampleData());
      // Urgent = red (#e74c3c)
      expect(result).toContain('#e74c3c');
      // Medium = yellow (#f1c40f)
      expect(result).toContain('#f1c40f');
      // Low = green (#2ecc71)
      expect(result).toContain('#2ecc71');
    });

    it('includes hover tooltips with assignee info', () => {
      const result = renderKanbanHtml(makeSampleData());
      expect(result).toContain('title=');
      expect(result).toContain('Bob, Carol');
    });

    it('handles empty tasks', () => {
      const result = renderKanbanHtml(makeEmptyData());
      expect(result).toContain('No tasks');
    });
  });

  describe('renderGanttHtml', () => {
    it('produces valid HTML with positioned bars', () => {
      const result = renderGanttHtml(makeSampleData());
      expect(result).toContain('<!DOCTYPE html>');
      expect(result).toContain('position:absolute');
    });

    it('renders task labels', () => {
      const result = renderGanttHtml(makeSampleData());
      expect(result).toContain('Setup project');
      expect(result).toContain('Build API');
    });

    it('renders date axis', () => {
      const result = renderGanttHtml(makeSampleData());
      expect(result).toContain('2026-01');
    });

    it('handles tasks with no dates', () => {
      const result = renderGanttHtml(makeNoDatesData());
      expect(result).toContain('No tasks with dates');
    });
  });

  describe('renderSummaryHtml', () => {
    it('produces valid HTML with stat cards', () => {
      const result = renderSummaryHtml(makeSampleData());
      expect(result).toContain('<!DOCTYPE html>');
      expect(result).toContain('Total Tasks');
      expect(result).toContain('Not Started');
      expect(result).toContain('In Progress');
      expect(result).toContain('Completed');
      expect(result).toContain('Overdue');
    });

    it('includes priority distribution bars', () => {
      const result = renderSummaryHtml(makeSampleData());
      expect(result).toContain('Priority Distribution');
      expect(result).toContain('Urgent');
      expect(result).toContain('Medium');
    });

    it('includes assignee workload table', () => {
      const result = renderSummaryHtml(makeSampleData());
      expect(result).toContain('Assignee Workload');
      expect(result).toContain('Alice');
      expect(result).toContain('Bob');
    });

    it('handles empty tasks', () => {
      const result = renderSummaryHtml(makeEmptyData());
      expect(result).toContain('Total Tasks');
    });
  });

  describe('renderBurndownHtml', () => {
    it('produces HTML with inline SVG chart', () => {
      const result = renderBurndownHtml(makeSampleData());
      expect(result).toContain('<!DOCTYPE html>');
      expect(result).toContain('<svg');
      expect(result).toContain('<polyline');
      expect(result).toContain('<circle');
    });

    it('includes ideal burndown line', () => {
      const result = renderBurndownHtml(makeSampleData());
      expect(result).toContain('stroke-dasharray');
      expect(result).toContain('#e74c3c');
    });

    it('includes axis labels', () => {
      const result = renderBurndownHtml(makeSampleData());
      expect(result).toContain('Date');
      expect(result).toContain('Remaining');
    });

    it('handles no date data', () => {
      const result = renderBurndownHtml(makeNoDatesData());
      expect(result).toContain('No date data available');
    });
  });
});
