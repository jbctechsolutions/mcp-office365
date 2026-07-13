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
      listMyPlannerTasksAsync: vi.fn(),
      getPlannerTaskAsync: vi.fn(),
      createPlannerTaskAsync: vi.fn(),
      updatePlannerTaskAsync: vi.fn(),
      deletePlannerTaskAsync: vi.fn(),
      getPlannerTaskDetailsAsync: vi.fn(),
      updatePlannerTaskDetailsAsync: vi.fn(),
      listPlannerTaskMessagesAsync: vi.fn(),
      createPlannerTaskMessageAsync: vi.fn(),
      updatePlannerTaskMessageAsync: vi.fn(),
      deletePlannerTaskMessageAsync: vi.fn(),
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
        { id: 'pl_a1', title: 'Sprint Plan', owner: 'group-abc', createdDateTime: '2026-01-01T00:00:00Z' },
        { id: 'pl_b2', title: 'Product Roadmap', owner: 'group-def', createdDateTime: '2026-02-01T00:00:00Z' },
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
        id: 'pl_a1', title: 'Sprint Plan', owner: 'group-abc', createdDateTime: '2026-01-01T00:00:00Z', etag: 'W/"abc123"',
      };
      vi.mocked(repo.getPlanAsync).mockResolvedValue(mockPlan);

      const result = await tools.getPlan({ plan_id: 'pl_a1' });

      expect(repo.getPlanAsync).toHaveBeenCalledWith('pl_a1');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.plan).toEqual(mockPlan);
      expect(parsed.plan.etag).toBe('W/"abc123"');
    });
  });

  describe('createPlan', () => {
    it('creates a plan and returns the ID', async () => {
      vi.mocked(repo.createPlanAsync).mockResolvedValue('pl_new42');

      const result = await tools.createPlan({ title: 'New Plan', group_id: 'group-xyz' });

      expect(repo.createPlanAsync).toHaveBeenCalledWith('New Plan', 'group-xyz');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.plan_id).toBe('pl_new42');
      expect(parsed.message).toBe('Plan created');
    });
  });

  describe('updatePlan', () => {
    it('updates a plan title', async () => {
      vi.mocked(repo.updatePlanAsync).mockResolvedValue(undefined);

      const result = await tools.updatePlan({ plan_id: 'pl_a1', title: 'Renamed Plan' });

      expect(repo.updatePlanAsync).toHaveBeenCalledWith('pl_a1', { title: 'Renamed Plan' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Plan updated');
    });

    it('calls update with empty updates when no title provided', async () => {
      vi.mocked(repo.updatePlanAsync).mockResolvedValue(undefined);

      await tools.updatePlan({ plan_id: 'pl_a1' });

      expect(repo.updatePlanAsync).toHaveBeenCalledWith('pl_a1', {});
    });
  });

  // ===========================================================================
  // Buckets
  // ===========================================================================

  describe('listBuckets', () => {
    it('returns buckets for a plan', async () => {
      const mockBuckets = [
        { id: 'pb_10', name: 'To Do', planId: 'pl_a1', orderHint: '1' },
        { id: 'pb_11', name: 'In Progress', planId: 'pl_a1', orderHint: '2' },
      ];
      vi.mocked(repo.listBucketsAsync).mockResolvedValue(mockBuckets);

      const result = await tools.listBuckets({ plan_id: 'pl_a1' });

      expect(repo.listBucketsAsync).toHaveBeenCalledWith('pl_a1');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.buckets).toEqual(mockBuckets);
    });
  });

  describe('createBucket', () => {
    it('creates a bucket and returns the ID', async () => {
      vi.mocked(repo.createBucketAsync).mockResolvedValue('pb_99');

      const result = await tools.createBucket({ plan_id: 'pl_a1', name: 'Done' });

      expect(repo.createBucketAsync).toHaveBeenCalledWith('pl_a1', 'Done');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.bucket_id).toBe('pb_99');
      expect(parsed.message).toBe('Bucket created');
    });
  });

  describe('updateBucket', () => {
    it('updates a bucket name', async () => {
      vi.mocked(repo.updateBucketAsync).mockResolvedValue(undefined);

      const result = await tools.updateBucket({ bucket_id: 'pb_10', name: 'Renamed' });

      expect(repo.updateBucketAsync).toHaveBeenCalledWith('pb_10', { name: 'Renamed' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Bucket updated');
    });

    it('calls update with empty updates when no name provided', async () => {
      vi.mocked(repo.updateBucketAsync).mockResolvedValue(undefined);

      await tools.updateBucket({ bucket_id: 'pb_10' });

      expect(repo.updateBucketAsync).toHaveBeenCalledWith('pb_10', {});
    });
  });

  describe('prepareDeleteBucket', () => {
    it('generates an approval token', () => {
      const result = tools.prepareDeleteBucket({ bucket_id: 'pb_42' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.bucket_id).toBe('pb_42');
      expect(parsed.action).toContain('confirm_delete_bucket');
    });
  });

  describe('confirmDeleteBucket', () => {
    it('deletes the bucket with a valid token', async () => {
      vi.mocked(repo.deleteBucketAsync).mockResolvedValue(undefined);

      // Generate a token first
      const prepareResult = tools.prepareDeleteBucket({ bucket_id: 'pb_42' });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeleteBucket({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Bucket deleted');
      expect(repo.deleteBucketAsync).toHaveBeenCalledWith('pb_42');
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmDeleteBucket({ approval_token: 'invalid-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error when token is reused', async () => {
      vi.mocked(repo.deleteBucketAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeleteBucket({ bucket_id: 'pb_42' });
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
        { id: 'pt_100', title: 'Task A', bucketId: 'pb_10', assignees: ['user1'], percentComplete: 0, priority: 5, startDateTime: '', dueDateTime: '2026-04-01T00:00:00Z', createdDateTime: '2026-03-01T00:00:00Z' },
        { id: 'pt_101', title: 'Task B', bucketId: null, assignees: [], percentComplete: 50, priority: 3, startDateTime: '2026-03-01T00:00:00Z', dueDateTime: '', createdDateTime: '2026-03-02T00:00:00Z' },
      ];
      vi.mocked(repo.listPlannerTasksAsync).mockResolvedValue(mockTasks);

      const result = await tools.listPlannerTasks({ plan_id: 'pl_a1' });

      expect(repo.listPlannerTasksAsync).toHaveBeenCalledWith('pl_a1');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.tasks).toEqual(mockTasks);
    });
  });

  describe('listMyPlannerTasks', () => {
    it('returns the signed-in user tasks across plans (with planId per task)', async () => {
      const mockTasks = [
        { id: 'pt_100', title: 'Task A', planId: 'pl_900', bucketId: 'pb_10', assignees: ['user1'], percentComplete: 0, priority: 5, startDateTime: '', dueDateTime: '2026-04-01T00:00:00Z', createdDateTime: '2026-03-01T00:00:00Z' },
        { id: 'pt_101', title: 'Task B', planId: 'pl_901', bucketId: null, assignees: ['user1'], percentComplete: 50, priority: 3, startDateTime: '', dueDateTime: '', createdDateTime: '2026-03-02T00:00:00Z' },
      ];
      vi.mocked(repo.listMyPlannerTasksAsync).mockResolvedValue(mockTasks);

      const result = await tools.listMyPlannerTasks({});

      expect(repo.listMyPlannerTasksAsync).toHaveBeenCalledWith();
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.tasks).toEqual(mockTasks);
    });
  });

  describe('getPlannerTask', () => {
    it('returns task details including etag', async () => {
      const mockTask = {
        id: 'pt_100', title: 'Task A', bucketId: 'pb_10', assignees: ['user1'],
        percentComplete: 0, priority: 5, startDateTime: '', dueDateTime: '2026-04-01T00:00:00Z',
        createdDateTime: '2026-03-01T00:00:00Z', conversationThreadId: 'thread-1',
        orderHint: '1', etag: 'W/"task-etag"',
      };
      vi.mocked(repo.getPlannerTaskAsync).mockResolvedValue(mockTask);

      const result = await tools.getPlannerTask({ task_id: 'pt_100' });

      expect(repo.getPlannerTaskAsync).toHaveBeenCalledWith('pt_100');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.task).toEqual(mockTask);
      expect(parsed.task.etag).toBe('W/"task-etag"');
    });
  });

  describe('createPlannerTask', () => {
    it('creates a task and returns the ID', async () => {
      vi.mocked(repo.createPlannerTaskAsync).mockResolvedValue('pt_200');

      const result = await tools.createPlannerTask({ plan_id: 'pl_a1', title: 'New Task' });

      expect(repo.createPlannerTaskAsync).toHaveBeenCalledWith('pl_a1', 'New Task', undefined, undefined, undefined, undefined, undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.task_id).toBe('pt_200');
      expect(parsed.message).toBe('Planner task created');
    });

    it('passes all optional parameters', async () => {
      vi.mocked(repo.createPlannerTaskAsync).mockResolvedValue('pt_201');
      const assignments = { 'user-1': { '@odata.type': '#microsoft.graph.plannerAssignment', 'orderHint': ' !' } };

      await tools.createPlannerTask({
        plan_id: 'pl_a1', title: 'Full Task', bucket_id: 'pb_10',
        assignments, priority: 3,
        start_date: '2026-03-01T00:00:00Z', due_date: '2026-04-01T00:00:00Z',
      });

      expect(repo.createPlannerTaskAsync).toHaveBeenCalledWith(
        'pl_a1', 'Full Task', 'pb_10', assignments, 3, '2026-03-01T00:00:00Z', '2026-04-01T00:00:00Z',
      );
    });
  });

  describe('updatePlannerTask', () => {
    it('updates a task title', async () => {
      vi.mocked(repo.updatePlannerTaskAsync).mockResolvedValue(undefined);

      const result = await tools.updatePlannerTask({ task_id: 'pt_100', title: 'Renamed Task' });

      expect(repo.updatePlannerTaskAsync).toHaveBeenCalledWith('pt_100', { title: 'Renamed Task' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Planner task updated');
    });

    it('updates multiple fields', async () => {
      vi.mocked(repo.updatePlannerTaskAsync).mockResolvedValue(undefined);

      await tools.updatePlannerTask({
        task_id: 'pt_100', percent_complete: 75, priority: 1, due_date: '2026-05-01T00:00:00Z',
      });

      expect(repo.updatePlannerTaskAsync).toHaveBeenCalledWith('pt_100', {
        percentComplete: 75, priority: 1, dueDate: '2026-05-01T00:00:00Z',
      });
    });

    it('calls update with empty updates when no fields provided', async () => {
      vi.mocked(repo.updatePlannerTaskAsync).mockResolvedValue(undefined);

      await tools.updatePlannerTask({ task_id: 'pt_100' });

      expect(repo.updatePlannerTaskAsync).toHaveBeenCalledWith('pt_100', {});
    });
  });

  describe('prepareDeletePlannerTask', () => {
    it('generates an approval token', () => {
      const result = tools.prepareDeletePlannerTask({ task_id: 'pt_100' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.task_id).toBe('pt_100');
      expect(parsed.action).toContain('confirm_delete_planner_task');
    });
  });

  describe('confirmDeletePlannerTask', () => {
    it('deletes the task with a valid token', async () => {
      vi.mocked(repo.deletePlannerTaskAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeletePlannerTask({ task_id: 'pt_100' });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeletePlannerTask({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Planner task deleted');
      expect(repo.deletePlannerTaskAsync).toHaveBeenCalledWith('pt_100');
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmDeletePlannerTask({ approval_token: 'invalid-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error when token is reused', async () => {
      vi.mocked(repo.deletePlannerTaskAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeletePlannerTask({ task_id: 'pt_100' });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      // First use should succeed
      await tools.confirmDeletePlannerTask({ approval_token });

      // Second use should fail
      const result = await tools.confirmDeletePlannerTask({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });

  // ===========================================================================
  // Planner Task Details
  // ===========================================================================

  describe('getPlannerTaskDetails', () => {
    it('returns task details from the repository', async () => {
      const mockDetails = {
        id: 'pt_1',
        description: 'Task notes here',
        checklist: {
          'guid-1': { title: 'Step 1', isChecked: false },
          'guid-2': { title: 'Step 2', isChecked: true },
        },
        references: {
          'https%3A//example.com': { alias: 'Example', type: 'Url' },
        },
        etag: 'W/"details-etag-123"',
      };
      vi.mocked(repo.getPlannerTaskDetailsAsync).mockResolvedValue(mockDetails);

      const result = await tools.getPlannerTaskDetails({ task_id: 'pt_1' });

      expect(repo.getPlannerTaskDetailsAsync).toHaveBeenCalledWith('pt_1');
      expect(result.content).toHaveLength(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.details).toEqual(mockDetails);
      expect(parsed.details.description).toBe('Task notes here');
      expect(parsed.details.etag).toBe('W/"details-etag-123"');
    });
  });

  describe('updatePlannerTaskDetails', () => {
    it('updates task details with description', async () => {
      vi.mocked(repo.updatePlannerTaskDetailsAsync).mockResolvedValue();

      const result = await tools.updatePlannerTaskDetails({
        task_id: 'pt_1',
        description: 'Updated notes',
      });

      expect(repo.updatePlannerTaskDetailsAsync).toHaveBeenCalledWith('pt_1', {
        description: 'Updated notes',
      });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Planner task details updated');
    });

    it('updates task details with checklist and references', async () => {
      vi.mocked(repo.updatePlannerTaskDetailsAsync).mockResolvedValue();

      const checklist = {
        'guid-abc': { title: 'New item', isChecked: false },
      };
      const references = {
        'https%3A//docs.example.com': { alias: 'Docs', type: 'Url' },
      };

      const result = await tools.updatePlannerTaskDetails({
        task_id: 'pt_2',
        checklist,
        references,
      });

      expect(repo.updatePlannerTaskDetailsAsync).toHaveBeenCalledWith('pt_2', {
        checklist,
        references,
      });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
    });

    it('only includes provided fields in updates', async () => {
      vi.mocked(repo.updatePlannerTaskDetailsAsync).mockResolvedValue();

      await tools.updatePlannerTaskDetails({ task_id: 'pt_3' });

      expect(repo.updatePlannerTaskDetailsAsync).toHaveBeenCalledWith('pt_3', {});
    });
  });

  // ===========================================================================
  // Planner Task Messages
  // ===========================================================================

  describe('listPlannerTaskMessages', () => {
    it('returns messages and optional paging token', async () => {
      const mockMessages = [{
        id: 'pm_abc',
        content: 'Looks good',
        messageType: 'richTextHtml',
        createdDateTime: '2026-07-13T00:00:00Z',
        editedTime: null,
        deletedTime: null,
        createdByUserId: 'user-1',
        mentions: [],
      }];
      vi.mocked(repo.listPlannerTaskMessagesAsync).mockResolvedValue({
        messages: mockMessages,
        nextSkipToken: 'token-xyz',
      });

      const result = await tools.listPlannerTaskMessages({ task_id: 'pt_1', skip_token: 'prev' });

      expect(repo.listPlannerTaskMessagesAsync).toHaveBeenCalledWith('pt_1', 'prev');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.messages).toEqual(mockMessages);
      expect(parsed.next_skip_token).toBe('token-xyz');
    });
  });

  describe('createPlannerTaskMessage', () => {
    it('posts a comment and returns the message id', async () => {
      vi.mocked(repo.createPlannerTaskMessageAsync).mockResolvedValue('pm_new');

      const result = await tools.createPlannerTaskMessage({
        task_id: 'pt_1',
        content: 'Preview ready: https://example.com',
        mention_user_ids: ['ben@example.com', 'user-2'],
      });

      expect(repo.createPlannerTaskMessageAsync).toHaveBeenCalledWith(
        'pt_1',
        'Preview ready: https://example.com',
        ['ben@example.com', 'user-2'],
      );
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message_id).toBe('pm_new');
    });
  });

  describe('updatePlannerTaskMessage', () => {
    it('updates a comment', async () => {
      vi.mocked(repo.updatePlannerTaskMessageAsync).mockResolvedValue();

      const result = await tools.updatePlannerTaskMessage({
        message_id: 'pm_abc',
        content: 'Updated text',
      });

      expect(repo.updatePlannerTaskMessageAsync).toHaveBeenCalledWith('pm_abc', 'Updated text', undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
    });
  });

  describe('prepareDeletePlannerTaskMessage', () => {
    it('generates an approval token', () => {
      const result = tools.prepareDeletePlannerTaskMessage({ message_id: 'pm_abc' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(parsed.message_id).toBe('pm_abc');
      expect(parsed.action).toContain('confirm_delete_planner_task_message');
    });
  });

  describe('confirmDeletePlannerTaskMessage', () => {
    it('deletes the comment with a valid token', async () => {
      vi.mocked(repo.deletePlannerTaskMessageAsync).mockResolvedValue();

      const prepareResult = tools.prepareDeletePlannerTaskMessage({ message_id: 'pm_abc' });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeletePlannerTaskMessage({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(repo.deletePlannerTaskMessageAsync).toHaveBeenCalledWith('pm_abc');
    });
  });
});
