/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Markdown visualization renderer.
 */

import { describe, it, expect } from 'vitest';
import type { PlanVisualizationData } from '../../../src/visualization/types.js';
import {
  renderKanbanMarkdown,
  renderGanttMarkdown,
  renderSummaryMarkdown,
  renderBurndownMarkdown,
} from '../../../src/visualization/markdown.js';

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
    plan: { id: 1, title: 'No Dates Plan' },
    buckets: [{ id: 10, name: 'Backlog', orderHint: '1' }],
    tasks: [
      {
        id: 201,
        title: 'Unscheduled task',
        bucketId: 10,
        percentComplete: 0,
        priority: 5,
        assignments: [],
      },
    ],
  };
}

describe('Markdown Renderer', () => {
  describe('renderKanbanMarkdown', () => {
    it('groups tasks by bucket with table headers', () => {
      const result = renderKanbanMarkdown(makeSampleData());
      expect(result).toContain('# Sprint 1 - Kanban Board');
      expect(result).toContain('## To Do');
      expect(result).toContain('## In Progress');
      expect(result).toContain('## Done');
      expect(result).toContain('| Title | Priority | Assignees | % Complete | Due Date |');
    });

    it('includes task details in table rows', () => {
      const result = renderKanbanMarkdown(makeSampleData());
      expect(result).toContain('Setup project');
      expect(result).toContain('Medium');
      expect(result).toContain('Alice');
      expect(result).toContain('0%');
      expect(result).toContain('2026-01-10');
    });

    it('shows correct priority labels', () => {
      const result = renderKanbanMarkdown(makeSampleData());
      expect(result).toContain('Urgent');
      expect(result).toContain('Medium');
      expect(result).toContain('Low');
    });

    it('handles empty tasks array', () => {
      const result = renderKanbanMarkdown(makeEmptyData());
      expect(result).toContain('_No tasks_');
    });

    it('handles tasks with no assignees', () => {
      const result = renderKanbanMarkdown(makeNoDatesData());
      expect(result).toContain('| Unscheduled task');
      // Should show '-' for no assignees
      expect(result).toMatch(/\|\s*-\s*\|/);
    });
  });

  describe('renderGanttMarkdown', () => {
    it('wraps mermaid gantt in code block', () => {
      const result = renderGanttMarkdown(makeSampleData());
      expect(result).toContain('```mermaid');
      expect(result).toContain('gantt');
      expect(result).toContain('dateFormat YYYY-MM-DD');
      expect(result).toContain('```');
    });

    it('creates sections per bucket', () => {
      const result = renderGanttMarkdown(makeSampleData());
      expect(result).toContain('section To Do');
      expect(result).toContain('section In Progress');
      expect(result).toContain('section Done');
    });

    it('marks completed tasks as done', () => {
      const result = renderGanttMarkdown(makeSampleData());
      expect(result).toContain('Write docs :done,');
    });

    it('marks in-progress tasks as active', () => {
      const result = renderGanttMarkdown(makeSampleData());
      expect(result).toContain('Build API :active,');
    });
  });

  describe('renderSummaryMarkdown', () => {
    it('includes statistics table', () => {
      const result = renderSummaryMarkdown(makeSampleData());
      expect(result).toContain('## Task Statistics');
      expect(result).toContain('| Total Tasks | 3 |');
      expect(result).toContain('| Completed | 1 |');
    });

    it('includes assignee workload', () => {
      const result = renderSummaryMarkdown(makeSampleData());
      expect(result).toContain('## Assignee Workload');
      expect(result).toContain('Alice');
      expect(result).toContain('Bob');
    });

    it('handles empty tasks', () => {
      const result = renderSummaryMarkdown(makeEmptyData());
      expect(result).toContain('| Total Tasks | 0 |');
    });
  });

  describe('renderBurndownMarkdown', () => {
    it('includes date-indexed table', () => {
      const result = renderBurndownMarkdown(makeSampleData());
      expect(result).toContain('| Date | Remaining Tasks | Completed |');
      expect(result).toContain('2026-01-07');
    });

    it('handles tasks with no due dates', () => {
      const result = renderBurndownMarkdown(makeNoDatesData());
      expect(result).toContain('No tasks with due dates');
    });
  });
});
