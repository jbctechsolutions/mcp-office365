/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for batch-optimized repository methods.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { GraphRepository, createGraphRepository } from '../../../src/graph/repository.js';
import { hashStringToNumber } from '../../../src/graph/mappers/utils.js';
import type { BatchResponseItem } from '../../../src/graph/client/batch.js';

// Mock the GraphClient with planner + batch methods
vi.mock('../../../src/graph/client/index.js', () => ({
  GraphClient: vi.fn().mockImplementation(function() {
    return {
      listPlannerTasks: vi.fn(),
      getPlannerTask: vi.fn(),
      getPlannerTaskDetails: vi.fn(),
      listPlans: vi.fn(),
      batchRequests: vi.fn(),
    };
  }),
}));

// Mock the downloadAttachment helper and getDownloadDir
vi.mock('../../../src/graph/attachments.js', () => ({
  downloadAttachment: vi.fn(),
  getDownloadDir: vi.fn().mockReturnValue('/tmp/mcp-outlook-attachments'),
}));

// Mock fs and path (required by repository)
vi.mock('fs', () => ({
  writeFileSync: vi.fn(),
  readFileSync: vi.fn().mockReturnValue(Buffer.from('fake')),
  mkdirSync: vi.fn(),
  existsSync: vi.fn().mockReturnValue(false),
}));

vi.mock('path', () => ({
  join: vi.fn().mockImplementation((...args: string[]) => args.join('/')),
  extname: vi.fn().mockImplementation((p: string) => {
    const dot = p.lastIndexOf('.');
    return dot >= 0 ? p.substring(dot) : '';
  }),
}));

function makePlannerTask(id: string, title: string) {
  return {
    id,
    title,
    bucketId: 'bucket-abc',
    assignments: { 'user1@example.com': {} },
    percentComplete: 0,
    priority: 5,
    startDateTime: '2026-01-01T00:00:00Z',
    dueDateTime: '2026-01-15T00:00:00Z',
    createdDateTime: '2025-12-01T00:00:00Z',
    '@odata.etag': `"etag-${id}"`,
  };
}

describe('batch-optimized repository methods', () => {
  let repository: GraphRepository;
  let mockClient: any;
  const planGraphId = 'plan-graph-id-123';
  let planNumericId: number;

  beforeEach(async () => {
    vi.clearAllMocks();
    repository = createGraphRepository();
    mockClient = (repository as any).client;

    // Seed the plan ID in the cache so planner methods work
    planNumericId = hashStringToNumber(planGraphId);
    (repository as any).idCache.plans.set(planNumericId, {
      planId: planGraphId,
      etag: '"plan-etag"',
    });
  });

  describe('listPlannerTasksWithDetailsAsync', () => {
    it('returns tasks with details when all batch requests succeed', async () => {
      const task1 = makePlannerTask('task-id-1', 'Task One');
      const task2 = makePlannerTask('task-id-2', 'Task Two');
      mockClient.listPlannerTasks.mockResolvedValue([task1, task2]);

      const task1NumericId = hashStringToNumber('task-id-1');
      const task2NumericId = hashStringToNumber('task-id-2');

      const batchResults = new Map<string, BatchResponseItem>();
      batchResults.set(String(task1NumericId), {
        id: String(task1NumericId),
        status: 200,
        headers: { ETag: '"detail-etag-1"' },
        body: {
          description: 'Description for task 1',
          checklist: { 'item1': { title: 'Check item', isChecked: false } },
          references: {},
        },
      });
      batchResults.set(String(task2NumericId), {
        id: String(task2NumericId),
        status: 200,
        headers: { ETag: '"detail-etag-2"' },
        body: {
          description: 'Description for task 2',
          checklist: {},
          references: { 'ref1': { alias: 'link' } },
        },
      });
      mockClient.batchRequests.mockResolvedValue(batchResults);

      const result = await repository.listPlannerTasksWithDetailsAsync(planNumericId);

      expect(result).toHaveLength(2);
      expect(result[0].title).toBe('Task One');
      expect(result[0].details).toBeDefined();
      expect(result[0].details!.description).toBe('Description for task 1');
      expect(result[0].details!.checklist).toHaveProperty('item1');
      expect(result[0].details!.etag).toBe('"detail-etag-1"');

      expect(result[1].title).toBe('Task Two');
      expect(result[1].details).toBeDefined();
      expect(result[1].details!.description).toBe('Description for task 2');
      expect(result[1].details!.references).toHaveProperty('ref1');

      // Verify batchRequests was called with correct URLs
      const batchCallArgs = mockClient.batchRequests.mock.calls[0][0];
      expect(batchCallArgs).toHaveLength(2);
      expect(batchCallArgs[0].method).toBe('GET');
      expect(batchCallArgs[0].url).toContain('/planner/tasks/task-id-1/details');
      expect(batchCallArgs[1].url).toContain('/planner/tasks/task-id-2/details');
    });

    it('handles partial failures gracefully', async () => {
      const task1 = makePlannerTask('task-id-1', 'Task One');
      const task2 = makePlannerTask('task-id-2', 'Task Two');
      mockClient.listPlannerTasks.mockResolvedValue([task1, task2]);

      const task1NumericId = hashStringToNumber('task-id-1');
      const task2NumericId = hashStringToNumber('task-id-2');

      const batchResults = new Map<string, BatchResponseItem>();
      batchResults.set(String(task1NumericId), {
        id: String(task1NumericId),
        status: 200,
        headers: { ETag: '"detail-etag-1"' },
        body: {
          description: 'Description for task 1',
          checklist: {},
          references: {},
        },
      });
      // Task 2 fails with 404
      batchResults.set(String(task2NumericId), {
        id: String(task2NumericId),
        status: 404,
        body: { error: { code: 'NotFound', message: 'Not found' } },
      });
      mockClient.batchRequests.mockResolvedValue(batchResults);

      const result = await repository.listPlannerTasksWithDetailsAsync(planNumericId);

      expect(result).toHaveLength(2);
      expect(result[0].details).toBeDefined();
      expect(result[0].details!.description).toBe('Description for task 1');
      // Task 2 should have undefined details due to failure
      expect(result[1].title).toBe('Task Two');
      expect(result[1].details).toBeUndefined();
    });

    it('returns empty array when plan has no tasks', async () => {
      mockClient.listPlannerTasks.mockResolvedValue([]);

      const result = await repository.listPlannerTasksWithDetailsAsync(planNumericId);

      expect(result).toHaveLength(0);
      // batchRequests should not be called when there are no tasks
      expect(mockClient.batchRequests).not.toHaveBeenCalled();
    });

    it('handles missing batch result for a task', async () => {
      const task1 = makePlannerTask('task-id-1', 'Task One');
      mockClient.listPlannerTasks.mockResolvedValue([task1]);

      // Return an empty batch result map (no results for any task)
      mockClient.batchRequests.mockResolvedValue(new Map());

      const result = await repository.listPlannerTasksWithDetailsAsync(planNumericId);

      expect(result).toHaveLength(1);
      expect(result[0].title).toBe('Task One');
      expect(result[0].details).toBeUndefined();
    });

    it('builds correct batch requests for many tasks (>20 handled by client)', async () => {
      // Create 25 tasks to verify that the method passes all requests to batchRequests
      // (the client's batchRequests method handles splitting into batches of 20)
      const tasks = Array.from({ length: 25 }, (_, i) =>
        makePlannerTask(`task-id-${i}`, `Task ${i}`),
      );
      mockClient.listPlannerTasks.mockResolvedValue(tasks);

      const batchResults = new Map<string, BatchResponseItem>();
      for (let i = 0; i < 25; i++) {
        const numericId = hashStringToNumber(`task-id-${i}`);
        batchResults.set(String(numericId), {
          id: String(numericId),
          status: 200,
          headers: { ETag: `"detail-etag-${i}"` },
          body: {
            description: `Description ${i}`,
            checklist: {},
            references: {},
          },
        });
      }
      mockClient.batchRequests.mockResolvedValue(batchResults);

      const result = await repository.listPlannerTasksWithDetailsAsync(planNumericId);

      expect(result).toHaveLength(25);
      // Verify all tasks have details
      for (let i = 0; i < 25; i++) {
        expect(result[i].details).toBeDefined();
        expect(result[i].details!.description).toBe(`Description ${i}`);
      }

      // Verify batchRequests was called with all 25 requests
      const batchCallArgs = mockClient.batchRequests.mock.calls[0][0];
      expect(batchCallArgs).toHaveLength(25);
    });

    it('caches task detail ETags for later updates', async () => {
      const task1 = makePlannerTask('task-id-1', 'Task One');
      mockClient.listPlannerTasks.mockResolvedValue([task1]);

      const task1NumericId = hashStringToNumber('task-id-1');

      const batchResults = new Map<string, BatchResponseItem>();
      batchResults.set(String(task1NumericId), {
        id: String(task1NumericId),
        status: 200,
        headers: { ETag: '"detail-etag-cached"' },
        body: {
          description: 'Cached description',
          checklist: {},
          references: {},
        },
      });
      mockClient.batchRequests.mockResolvedValue(batchResults);

      await repository.listPlannerTasksWithDetailsAsync(planNumericId);

      // Verify the detail ETag was cached
      const cachedDetail = (repository as any).idCache.plannerTaskDetails.get(task1NumericId);
      expect(cachedDetail).toBeDefined();
      expect(cachedDetail.etag).toBe('"detail-etag-cached"');
      expect(cachedDetail.taskId).toBe('task-id-1');
    });

    it('throws when plan ID is not in cache', async () => {
      mockClient.listPlans.mockResolvedValue([]);
      await expect(
        repository.listPlannerTasksWithDetailsAsync(999999),
      ).rejects.toThrow('not found');
    });
  });
});
