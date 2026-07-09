/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for batch-optimized repository methods.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { GraphRepository, createGraphRepository } from '../../../src/graph/repository.js';
import { StateStore } from '../../../src/state/store.js';
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
  let planTok: string;

  beforeEach(async () => {
    vi.clearAllMocks();
    // Alias-backed entities (plans, planner tasks, …) need a store to resolve
    // their tokens. fs is mocked in this suite, so StateStore.open degrades to
    // an in-memory sqlite db — still a fully-functional alias table for the run.
    const store = StateStore.open({ dir: '/tmp/mcp-o365-batch-opt-test', warn: () => {} });
    repository = createGraphRepository(undefined, store);
    mockClient = (repository as any).client;

    // Mint a durable pl_ token for the plan so plan-scoped Planner methods resolve.
    mockClient.listPlans.mockResolvedValue([
      { id: planGraphId, title: 'Plan', owner: 'group-1', createdDateTime: '' },
    ]);
    const plans = await repository.listPlansAsync();
    planTok = plans[0].id;
  });

  describe('listPlannerTasksWithDetailsAsync', () => {
    it('returns tasks with details when all batch requests succeed', async () => {
      const task1 = makePlannerTask('task-id-1', 'Task One');
      const task2 = makePlannerTask('task-id-2', 'Task Two');
      mockClient.listPlannerTasks.mockResolvedValue([task1, task2]);

      // Learn the durable pt_ tokens the repository mints for these Graph ids
      // (mintAlias is idempotent — minting twice for the same id + account
      // returns the same token, so this matches what the call under test mints).
      const listed = await repository.listPlannerTasksAsync(planTok);
      const [task1Tok, task2Tok] = listed.map((t) => t.id);

      const batchResults = new Map<string, BatchResponseItem>();
      batchResults.set(task1Tok, {
        id: task1Tok,
        status: 200,
        headers: { ETag: '"detail-etag-1"' },
        body: {
          description: 'Description for task 1',
          checklist: { 'item1': { title: 'Check item', isChecked: false } },
          references: {},
        },
      });
      batchResults.set(task2Tok, {
        id: task2Tok,
        status: 200,
        headers: { ETag: '"detail-etag-2"' },
        body: {
          description: 'Description for task 2',
          checklist: {},
          references: { 'ref1': { alias: 'link' } },
        },
      });
      mockClient.batchRequests.mockResolvedValue(batchResults);

      const result = await repository.listPlannerTasksWithDetailsAsync(planTok);

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

      const listed = await repository.listPlannerTasksAsync(planTok);
      const [task1Tok, task2Tok] = listed.map((t) => t.id);

      const batchResults = new Map<string, BatchResponseItem>();
      batchResults.set(task1Tok, {
        id: task1Tok,
        status: 200,
        headers: { ETag: '"detail-etag-1"' },
        body: {
          description: 'Description for task 1',
          checklist: {},
          references: {},
        },
      });
      // Task 2 fails with 404
      batchResults.set(task2Tok, {
        id: task2Tok,
        status: 404,
        body: { error: { code: 'NotFound', message: 'Not found' } },
      });
      mockClient.batchRequests.mockResolvedValue(batchResults);

      const result = await repository.listPlannerTasksWithDetailsAsync(planTok);

      expect(result).toHaveLength(2);
      expect(result[0].details).toBeDefined();
      expect(result[0].details!.description).toBe('Description for task 1');
      // Task 2 should have undefined details due to failure
      expect(result[1].title).toBe('Task Two');
      expect(result[1].details).toBeUndefined();
    });

    it('returns empty array when plan has no tasks', async () => {
      mockClient.listPlannerTasks.mockResolvedValue([]);

      const result = await repository.listPlannerTasksWithDetailsAsync(planTok);

      expect(result).toHaveLength(0);
      // batchRequests should not be called when there are no tasks
      expect(mockClient.batchRequests).not.toHaveBeenCalled();
    });

    it('handles missing batch result for a task', async () => {
      const task1 = makePlannerTask('task-id-1', 'Task One');
      mockClient.listPlannerTasks.mockResolvedValue([task1]);

      // Return an empty batch result map (no results for any task)
      mockClient.batchRequests.mockResolvedValue(new Map());

      const result = await repository.listPlannerTasksWithDetailsAsync(planTok);

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

      const listed = await repository.listPlannerTasksAsync(planTok);

      const batchResults = new Map<string, BatchResponseItem>();
      for (let i = 0; i < 25; i++) {
        const tok = listed[i].id;
        batchResults.set(tok, {
          id: tok,
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

      const result = await repository.listPlannerTasksWithDetailsAsync(planTok);

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

    it('returns the freshly-fetched detail etag (U5b-5 — no cross-call etag cache)', async () => {
      const task1 = makePlannerTask('task-id-1', 'Task One');
      mockClient.listPlannerTasks.mockResolvedValue([task1]);

      const listed = await repository.listPlannerTasksAsync(planTok);
      const task1Tok = listed[0].id;

      const batchResults = new Map<string, BatchResponseItem>();
      batchResults.set(task1Tok, {
        id: task1Tok,
        status: 200,
        headers: { ETag: '"detail-etag-fresh"' },
        body: {
          description: 'Fresh description',
          checklist: {},
          references: {},
        },
      });
      mockClient.batchRequests.mockResolvedValue(batchResults);

      const result = await repository.listPlannerTasksWithDetailsAsync(planTok);

      expect(result[0].details!.etag).toBe('"detail-etag-fresh"');
    });

    it('rejects an unresolvable plan token', async () => {
      mockClient.listPlans.mockResolvedValue([]);
      await expect(
        repository.listPlannerTasksWithDetailsAsync('pl_totallybogus000'),
      ).rejects.toThrow('Unknown or unresolvable');
    });
  });
});
