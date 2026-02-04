/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Graph task mapper functions.
 */

import { describe, it, expect } from 'vitest';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { mapTaskToTaskRow, type TodoTaskWithList } from '../../../../src/graph/mappers/task-mapper.js';
import { hashStringToNumber } from '../../../../src/graph/mappers/utils.js';

describe('graph/mappers/task-mapper', () => {
  describe('mapTaskToTaskRow', () => {
    it('maps task with all fields', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        taskListId: 'list-456',
        title: 'Test Task',
        status: 'completed',
        dueDateTime: { dateTime: '2024-01-15T17:00:00', timeZone: 'UTC' },
        startDateTime: { dateTime: '2024-01-10T09:00:00', timeZone: 'UTC' },
        importance: 'high',
        isReminderOn: true,
      };

      const result = mapTaskToTaskRow(task);

      expect(result.id).toBe(hashStringToNumber('task-123'));
      expect(result.folderId).toBe(hashStringToNumber('list-456'));
      expect(result.name).toBe('Test Task');
      expect(result.isCompleted).toBe(1);
      expect(result.priority).toBe(1); // high
      expect(result.hasReminder).toBe(1);
      expect(result.dataFilePath).toBe('graph-task:list-456:task-123');
    });

    it('handles task with null id', () => {
      const task: TodoTaskWithList = {
        id: undefined,
        title: 'Test',
      };

      const result = mapTaskToTaskRow(task);

      expect(result.id).toBe(hashStringToNumber(''));
    });

    it('handles task without taskListId', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        taskListId: undefined,
      };

      const result = mapTaskToTaskRow(task);

      expect(result.folderId).toBe(0);
      expect(result.dataFilePath).toBe('graph-task:default:task-123');
    });

    it('handles task with null title', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        title: undefined,
      };

      const result = mapTaskToTaskRow(task);

      expect(result.name).toBeNull();
    });

    it('handles incomplete task', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        status: 'notStarted',
      };

      const result = mapTaskToTaskRow(task);

      expect(result.isCompleted).toBe(0);
    });

    it('handles in-progress task', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        status: 'inProgress',
      };

      const result = mapTaskToTaskRow(task);

      expect(result.isCompleted).toBe(0);
    });

    it('handles task without status', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        status: undefined,
      };

      const result = mapTaskToTaskRow(task);

      expect(result.isCompleted).toBe(0);
    });

    it('handles task without due date', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        dueDateTime: undefined,
      };

      const result = mapTaskToTaskRow(task);

      expect(result.dueDate).toBeNull();
    });

    it('handles task without start date', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        startDateTime: undefined,
      };

      const result = mapTaskToTaskRow(task);

      expect(result.startDate).toBeNull();
    });

    it('handles low importance', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        importance: 'low',
      };

      const result = mapTaskToTaskRow(task);

      expect(result.priority).toBe(-1);
    });

    it('handles normal importance', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        importance: 'normal',
      };

      const result = mapTaskToTaskRow(task);

      expect(result.priority).toBe(0);
    });

    it('handles task without importance', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        importance: undefined,
      };

      const result = mapTaskToTaskRow(task);

      expect(result.priority).toBe(0);
    });

    it('handles task without reminder', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        isReminderOn: false,
      };

      const result = mapTaskToTaskRow(task);

      expect(result.hasReminder).toBe(0);
    });

    it('handles task with undefined isReminderOn', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        isReminderOn: undefined,
      };

      const result = mapTaskToTaskRow(task);

      expect(result.hasReminder).toBe(0);
    });

    it('parses due date correctly', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        dueDateTime: { dateTime: '2024-01-15T17:00:00Z', timeZone: 'UTC' },
      };

      const result = mapTaskToTaskRow(task);

      expect(result.dueDate).toBeTypeOf('number');
      expect(result.dueDate).toBeGreaterThan(0);
    });

    it('parses start date correctly', () => {
      const task: TodoTaskWithList = {
        id: 'task-123',
        startDateTime: { dateTime: '2024-01-10T09:00:00Z', timeZone: 'UTC' },
      };

      const result = mapTaskToTaskRow(task);

      expect(result.startDate).toBeTypeOf('number');
      expect(result.startDate).toBeGreaterThan(0);
    });
  });
});
