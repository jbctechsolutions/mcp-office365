/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { createTestDatabase, SAMPLE_COUNTS } from '../../fixtures/database.js';
import { createConnection, type IConnection } from '../../../src/database/connection.js';
import { createRepository, type IRepository } from '../../../src/database/repository.js';
import {
  TasksTools,
  createTasksTools,
  ListTasksInput,
  SearchTasksInput,
  GetTaskInput,
  type ITaskContentReader,
  type TaskDetails,
} from '../../../src/tools/tasks.js';

describe('TasksTools', () => {
  let testDb: { path: string; cleanup: () => void };
  let connection: IConnection;
  let repository: IRepository;
  let tasksTools: TasksTools;

  beforeEach(() => {
    testDb = createTestDatabase();
    connection = createConnection(testDb.path);
    repository = createRepository(connection);
    tasksTools = createTasksTools(repository);
  });

  afterEach(() => {
    connection.close();
    testDb.cleanup();
  });

  // ---------------------------------------------------------------------------
  // Input Validation
  // ---------------------------------------------------------------------------

  describe('input validation', () => {
    it('validates ListTasksInput with defaults', () => {
      const parsed = ListTasksInput.parse({});
      expect(parsed.limit).toBe(50);
      expect(parsed.offset).toBe(0);
      expect(parsed.include_completed).toBe(true);
    });

    it('validates ListTasksInput with options', () => {
      const input = { limit: 25, offset: 10, include_completed: false };
      const parsed = ListTasksInput.parse(input);
      expect(parsed).toEqual(input);
    });

    it('validates SearchTasksInput', () => {
      const parsed = SearchTasksInput.parse({ query: 'report' });
      expect(parsed.query).toBe('report');
      expect(parsed.limit).toBe(50);
    });

    it('validates GetTaskInput', () => {
      const parsed = GetTaskInput.parse({ task_id: 1 });
      expect(parsed.task_id).toBe(1);
    });
  });

  // ---------------------------------------------------------------------------
  // listTasks
  // ---------------------------------------------------------------------------

  describe('listTasks', () => {
    it('returns all tasks', () => {
      const tasks = tasksTools.listTasks({ limit: 50, offset: 0, include_completed: true });
      expect(tasks.length).toBe(SAMPLE_COUNTS.tasks);
    });

    it('returns tasks with correct structure', () => {
      const tasks = tasksTools.listTasks({ limit: 1, offset: 0, include_completed: true });
      const task = tasks[0];

      expect(task).toHaveProperty('id');
      expect(task).toHaveProperty('folderId');
      expect(task).toHaveProperty('name');
      expect(task).toHaveProperty('isCompleted');
      expect(task).toHaveProperty('dueDate');
      expect(task).toHaveProperty('priority');
      expect(typeof task?.isCompleted).toBe('boolean');
    });

    it('respects limit parameter', () => {
      const tasks = tasksTools.listTasks({ limit: 1, offset: 0, include_completed: true });
      expect(tasks.length).toBe(1);
    });

    it('filters incomplete tasks when include_completed is false', () => {
      const tasks = tasksTools.listTasks({ limit: 50, offset: 0, include_completed: false });
      expect(tasks.length).toBe(SAMPLE_COUNTS.incompleteTasks);
      expect(tasks.every((t) => t.isCompleted === false)).toBe(true);
    });

    it('converts timestamps to ISO format', () => {
      const tasks = tasksTools.listTasks({ limit: 50, offset: 0, include_completed: true });
      const taskWithDate = tasks.find((t) => t.dueDate != null);

      if (taskWithDate?.dueDate) {
        expect(taskWithDate.dueDate).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // searchTasks
  // ---------------------------------------------------------------------------

  describe('searchTasks', () => {
    it('finds tasks by name', () => {
      const tasks = tasksTools.searchTasks({ query: 'report', limit: 50 });
      expect(tasks.length).toBeGreaterThan(0);
    });

    it('returns empty array for no matches', () => {
      const tasks = tasksTools.searchTasks({ query: 'xyznonexistent', limit: 50 });
      expect(tasks.length).toBe(0);
    });
  });

  // ---------------------------------------------------------------------------
  // getTask
  // ---------------------------------------------------------------------------

  describe('getTask', () => {
    it('returns task by ID', () => {
      const tasks = tasksTools.listTasks({ limit: 1, offset: 0, include_completed: true });
      const firstTask = tasks[0];

      if (firstTask) {
        const task = tasksTools.getTask({ task_id: firstTask.id });
        expect(task).not.toBeNull();
        expect(task?.id).toBe(firstTask.id);
      }
    });

    it('returns null for non-existent ID', () => {
      const task = tasksTools.getTask({ task_id: 99999 });
      expect(task).toBeNull();
    });

    it('includes additional fields in full task', () => {
      const tasks = tasksTools.listTasks({ limit: 1, offset: 0, include_completed: true });
      const firstTask = tasks[0];

      if (firstTask) {
        const task = tasksTools.getTask({ task_id: firstTask.id });
        expect(task).toHaveProperty('startDate');
        expect(task).toHaveProperty('completedDate');
        expect(task).toHaveProperty('hasReminder');
        expect(task).toHaveProperty('body');
        expect(task).toHaveProperty('categories');
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Content Reader Integration
  // ---------------------------------------------------------------------------

  describe('content reader integration', () => {
    it('uses content reader for task details', () => {
      const mockDetails: TaskDetails = {
        body: 'Task description here',
        completedDate: '2024-01-15T10:00:00.000Z',
        reminderDate: '2024-01-14T09:00:00.000Z',
        categories: ['Work', 'Important'],
      };

      const mockContentReader: ITaskContentReader = {
        readTaskDetails: () => mockDetails,
      };

      const toolsWithReader = createTasksTools(repository, mockContentReader);
      const tasks = toolsWithReader.listTasks({ limit: 1, offset: 0, include_completed: true });

      if (tasks[0]) {
        const task = toolsWithReader.getTask({ task_id: tasks[0].id });
        expect(task?.body).toBe('Task description here');
        expect(task?.completedDate).toBe('2024-01-15T10:00:00.000Z');
        expect(task?.categories).toEqual(['Work', 'Important']);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Factory Function
  // ---------------------------------------------------------------------------

  describe('createTasksTools', () => {
    it('creates a TasksTools instance', () => {
      const tools = createTasksTools(repository);
      expect(tools).toBeInstanceOf(TasksTools);
    });
  });
});
