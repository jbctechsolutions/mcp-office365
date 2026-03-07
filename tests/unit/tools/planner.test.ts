/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Microsoft Planner Plans and Buckets tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { PlannerTools, type IPlannerRepository } from '../../../src/tools/planner.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';

describe('PlannerTools', () => {
  let repo: IPlannerRepository;
  let tokenManager: ApprovalTokenManager;
  let tools: PlannerTools;

  beforeEach(() => {
    repo = {
      listPlansAsync: vi.fn(),
      getPlanAsync: vi.fn(),
      createPlanAsync: vi.fn(),
      updatePlanAsync: vi.fn(),
      listBucketsAsync: vi.fn(),
      createBucketAsync: vi.fn(),
      updateBucketAsync: vi.fn(),
      deleteBucketAsync: vi.fn(),
      listPlannerTasksAsync: vi.fn(),
      getPlannerTaskAsync: vi.fn(),
      createPlannerTaskAsync: vi.fn(),
      updatePlannerTaskAsync: vi.fn(),
      deletePlannerTaskAsync: vi.fn(),
    };
    tokenManager = new ApprovalTokenManager();
    tools = new PlannerTools(repo, tokenManager);
  });

  // ===========================================================================
  // Plans
  // ===========================================================================

  describe('listPlans', () => {
    it('returns plans from the repository', async () => {
      const mockPlans = [
        { id: 1, title: 'Sprint Plan', owner: 'group-abc', createdDateTime: '2026-01-01T00:00:00Z' },
        { id: 2, title: 'Product Roadmap', owner: 'group-def', createdDateTime: '2026-02-01T00:00:00Z' },
      ];
      vi.mocked(repo.listPlansAsync).mockResolvedValue(mockPlans);

      const result = await tools.listPlans();

      expect(result.content).toHaveLength(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.plans).toEqual(mockPlans);
    });
  });

  describe('getPlan', () => {
    it('returns plan details including etag', async () => {
      const mockPlan = {
        id: 1, title: 'Sprint Plan', owner: 'group-abc', createdDateTime: '2026-01-01T00:00:00Z', etag: 'W/"abc123"',
      };
      vi.mocked(repo.getPlanAsync).mockResolvedValue(mockPlan);

      const result = await tools.getPlan({ plan_id: 1 });

      expect(repo.getPlanAsync).toHaveBeenCalledWith(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.plan).toEqual(mockPlan);
      expect(parsed.plan.etag).toBe('W/"abc123"');
    });
  });

  describe('createPlan', () => {
    it('creates a plan and returns the ID', async () => {
      vi.mocked(repo.createPlanAsync).mockResolvedValue(42);

      const result = await tools.createPlan({ title: 'New Plan', group_id: 'group-xyz' });

      expect(repo.createPlanAsync).toHaveBeenCalledWith('New Plan', 'group-xyz');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.plan_id).toBe(42);
      expect(parsed.message).toBe('Plan created');
    });
  });

  describe('updatePlan', () => {
    it('updates a plan title', async () => {
      vi.mocked(repo.updatePlanAsync).mockResolvedValue(undefined);

      const result = await tools.updatePlan({ plan_id: 1, title: 'Renamed Plan' });

      expect(repo.updatePlanAsync).toHaveBeenCalledWith(1, { title: 'Renamed Plan' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Plan updated');
    });

    it('calls update with empty updates when no title provided', async () => {
      vi.mocked(repo.updatePlanAsync).mockResolvedValue(undefined);

      await tools.updatePlan({ plan_id: 1 });

      expect(repo.updatePlanAsync).toHaveBeenCalledWith(1, {});
    });
  });

  // ===========================================================================
  // Buckets
  // ===========================================================================

  describe('listBuckets', () => {
    it('returns buckets for a plan', async () => {
      const mockBuckets = [
        { id: 10, name: 'To Do', planId: 1, orderHint: '1' },
        { id: 11, name: 'In Progress', planId: 1, orderHint: '2' },
      ];
      vi.mocked(repo.listBucketsAsync).mockResolvedValue(mockBuckets);

      const result = await tools.listBuckets({ plan_id: 1 });

      expect(repo.listBucketsAsync).toHaveBeenCalledWith(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.buckets).toEqual(mockBuckets);
    });
  });

  describe('createBucket', () => {
    it('creates a bucket and returns the ID', async () => {
      vi.mocked(repo.createBucketAsync).mockResolvedValue(99);

      const result = await tools.createBucket({ plan_id: 1, name: 'Done' });

      expect(repo.createBucketAsync).toHaveBeenCalledWith(1, 'Done');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.bucket_id).toBe(99);
      expect(parsed.message).toBe('Bucket created');
    });
  });

  describe('updateBucket', () => {
    it('updates a bucket name', async () => {
      vi.mocked(repo.updateBucketAsync).mockResolvedValue(undefined);

      const result = await tools.updateBucket({ bucket_id: 10, name: 'Renamed' });

      expect(repo.updateBucketAsync).toHaveBeenCalledWith(10, { name: 'Renamed' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Bucket updated');
    });

    it('calls update with empty updates when no name provided', async () => {
      vi.mocked(repo.updateBucketAsync).mockResolvedValue(undefined);

      await tools.updateBucket({ bucket_id: 10 });

      expect(repo.updateBucketAsync).toHaveBeenCalledWith(10, {});
    });
  });

  describe('prepareDeleteBucket', () => {
    it('generates an approval token', () => {
      const result = tools.prepareDeleteBucket({ bucket_id: 42 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.bucket_id).toBe(42);
      expect(parsed.action).toContain('confirm_delete_bucket');
    });
  });

  describe('confirmDeleteBucket', () => {
    it('deletes the bucket with a valid token', async () => {
      vi.mocked(repo.deleteBucketAsync).mockResolvedValue(undefined);

      // Generate a token first
      const prepareResult = tools.prepareDeleteBucket({ bucket_id: 42 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeleteBucket({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Bucket deleted');
      expect(repo.deleteBucketAsync).toHaveBeenCalledWith(42);
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmDeleteBucket({ approval_token: 'invalid-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error when token is reused', async () => {
      vi.mocked(repo.deleteBucketAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeleteBucket({ bucket_id: 42 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      // First use should succeed
      await tools.confirmDeleteBucket({ approval_token });

      // Second use should fail
      const result = await tools.confirmDeleteBucket({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });

  // ===========================================================================
  // Planner Tasks
  // ===========================================================================

  describe('listPlannerTasks', () => {
    it('returns tasks for a plan', async () => {
      const mockTasks = [
        { id: 100, title: 'Task A', bucketId: 10, assignees: ['user1'], percentComplete: 0, priority: 5, startDateTime: '', dueDateTime: '2026-04-01T00:00:00Z', createdDateTime: '2026-03-01T00:00:00Z' },
        { id: 101, title: 'Task B', bucketId: null, assignees: [], percentComplete: 50, priority: 3, startDateTime: '2026-03-01T00:00:00Z', dueDateTime: '', createdDateTime: '2026-03-02T00:00:00Z' },
      ];
      vi.mocked(repo.listPlannerTasksAsync).mockResolvedValue(mockTasks);

      const result = await tools.listPlannerTasks({ plan_id: 1 });

      expect(repo.listPlannerTasksAsync).toHaveBeenCalledWith(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.tasks).toEqual(mockTasks);
    });
  });

  describe('getPlannerTask', () => {
    it('returns task details including etag', async () => {
      const mockTask = {
        id: 100, title: 'Task A', bucketId: 10, assignees: ['user1'],
        percentComplete: 0, priority: 5, startDateTime: '', dueDateTime: '2026-04-01T00:00:00Z',
        createdDateTime: '2026-03-01T00:00:00Z', conversationThreadId: 'thread-1',
        orderHint: '1', etag: 'W/"task-etag"',
      };
      vi.mocked(repo.getPlannerTaskAsync).mockResolvedValue(mockTask);

      const result = await tools.getPlannerTask({ task_id: 100 });

      expect(repo.getPlannerTaskAsync).toHaveBeenCalledWith(100);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.task).toEqual(mockTask);
      expect(parsed.task.etag).toBe('W/"task-etag"');
    });
  });

  describe('createPlannerTask', () => {
    it('creates a task and returns the ID', async () => {
      vi.mocked(repo.createPlannerTaskAsync).mockResolvedValue(200);

      const result = await tools.createPlannerTask({ plan_id: 1, title: 'New Task' });

      expect(repo.createPlannerTaskAsync).toHaveBeenCalledWith(1, 'New Task', undefined, undefined, undefined, undefined, undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.task_id).toBe(200);
      expect(parsed.message).toBe('Planner task created');
    });

    it('passes all optional parameters', async () => {
      vi.mocked(repo.createPlannerTaskAsync).mockResolvedValue(201);
      const assignments = { 'user-1': { '@odata.type': '#microsoft.graph.plannerAssignment', 'orderHint': ' !' } };

      await tools.createPlannerTask({
        plan_id: 1, title: 'Full Task', bucket_id: 10,
        assignments, priority: 3,
        start_date: '2026-03-01T00:00:00Z', due_date: '2026-04-01T00:00:00Z',
      });

      expect(repo.createPlannerTaskAsync).toHaveBeenCalledWith(
        1, 'Full Task', 10, assignments, 3, '2026-03-01T00:00:00Z', '2026-04-01T00:00:00Z',
      );
    });
  });

  describe('updatePlannerTask', () => {
    it('updates a task title', async () => {
      vi.mocked(repo.updatePlannerTaskAsync).mockResolvedValue(undefined);

      const result = await tools.updatePlannerTask({ task_id: 100, title: 'Renamed Task' });

      expect(repo.updatePlannerTaskAsync).toHaveBeenCalledWith(100, { title: 'Renamed Task' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Planner task updated');
    });

    it('updates multiple fields', async () => {
      vi.mocked(repo.updatePlannerTaskAsync).mockResolvedValue(undefined);

      await tools.updatePlannerTask({
        task_id: 100, percent_complete: 75, priority: 1, due_date: '2026-05-01T00:00:00Z',
      });

      expect(repo.updatePlannerTaskAsync).toHaveBeenCalledWith(100, {
        percentComplete: 75, priority: 1, dueDate: '2026-05-01T00:00:00Z',
      });
    });

    it('calls update with empty updates when no fields provided', async () => {
      vi.mocked(repo.updatePlannerTaskAsync).mockResolvedValue(undefined);

      await tools.updatePlannerTask({ task_id: 100 });

      expect(repo.updatePlannerTaskAsync).toHaveBeenCalledWith(100, {});
    });
  });

  describe('prepareDeletePlannerTask', () => {
    it('generates an approval token', () => {
      const result = tools.prepareDeletePlannerTask({ task_id: 100 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.task_id).toBe(100);
      expect(parsed.action).toContain('confirm_delete_planner_task');
    });
  });

  describe('confirmDeletePlannerTask', () => {
    it('deletes the task with a valid token', async () => {
      vi.mocked(repo.deletePlannerTaskAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeletePlannerTask({ task_id: 100 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeletePlannerTask({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Planner task deleted');
      expect(repo.deletePlannerTaskAsync).toHaveBeenCalledWith(100);
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmDeletePlannerTask({ approval_token: 'invalid-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error when token is reused', async () => {
      vi.mocked(repo.deletePlannerTaskAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeletePlannerTask({ task_id: 100 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      // First use should succeed
      await tools.confirmDeletePlannerTask({ approval_token });

      // Second use should fail
      const result = await tools.confirmDeletePlannerTask({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });
});
