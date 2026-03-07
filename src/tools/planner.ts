/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Microsoft Planner Plans and Buckets MCP tools.
 *
 * Provides tools for managing Planner plans and buckets with ETag caching
 * for optimistic concurrency control, and two-phase approval for destructive
 * delete operations.
 */

import { z } from 'zod';
import type { ApprovalTokenManager } from '../approval/index.js';

// =============================================================================
// Input Schemas
// =============================================================================

export const ListPlansInput = z.strictObject({});

export const GetPlanInput = z.strictObject({
  plan_id: z.number().int().positive().describe('Plan ID from list_plans'),
});

export const CreatePlanInput = z.strictObject({
  title: z.string().min(1).describe('Plan title'),
  group_id: z.string().min(1).describe('M365 group ID that owns the plan'),
});

export const UpdatePlanInput = z.strictObject({
  plan_id: z.number().int().positive().describe('Plan ID from list_plans'),
  title: z.string().min(1).optional().describe('New plan title'),
});

export const ListBucketsInput = z.strictObject({
  plan_id: z.number().int().positive().describe('Plan ID from list_plans'),
});

export const CreateBucketInput = z.strictObject({
  plan_id: z.number().int().positive().describe('Plan ID from list_plans'),
  name: z.string().min(1).describe('Bucket name'),
});

export const UpdateBucketInput = z.strictObject({
  bucket_id: z.number().int().positive().describe('Bucket ID from list_buckets'),
  name: z.string().min(1).optional().describe('New bucket name'),
});

export const PrepareDeleteBucketInput = z.strictObject({
  bucket_id: z.number().int().positive().describe('Bucket ID from list_buckets'),
});

export const ConfirmDeleteBucketInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_bucket'),
});

export const ListPlannerTasksInput = z.strictObject({
  plan_id: z.number().int().positive().describe('Plan ID from list_plans'),
});

export const GetPlannerTaskInput = z.strictObject({
  task_id: z.number().int().positive().describe('Task ID from list_planner_tasks'),
});

export const CreatePlannerTaskInput = z.strictObject({
  plan_id: z.number().int().positive().describe('Plan ID from list_plans'),
  title: z.string().min(1).describe('Task title'),
  bucket_id: z.number().int().positive().optional().describe('Bucket ID from list_buckets'),
  assignments: z.record(z.string(), z.object({}).passthrough()).optional().describe('User assignments. Keys are user IDs, values should be { "@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": " !" }'),
  priority: z.number().int().min(0).max(10).optional().describe('Priority (0-10)'),
  start_date: z.string().optional().describe('Start date in ISO format'),
  due_date: z.string().optional().describe('Due date in ISO format'),
});

export const UpdatePlannerTaskInput = z.strictObject({
  task_id: z.number().int().positive().describe('Task ID from list_planner_tasks'),
  title: z.string().min(1).optional().describe('New task title'),
  bucket_id: z.number().int().positive().optional().describe('New bucket ID from list_buckets'),
  percent_complete: z.number().int().min(0).max(100).optional().describe('Percent complete (0-100)'),
  priority: z.number().int().min(0).max(10).optional().describe('Priority (0-10)'),
  start_date: z.string().optional().describe('Start date in ISO format'),
  due_date: z.string().optional().describe('Due date in ISO format'),
  assignments: z.record(z.string(), z.object({}).passthrough()).optional().describe('User assignments. Keys are user IDs, values should be { "@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": " !" }'),
});

export const PrepareDeletePlannerTaskInput = z.strictObject({
  task_id: z.number().int().positive().describe('Task ID from list_planner_tasks'),
});

export const ConfirmDeletePlannerTaskInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_planner_task'),
});

export const GetPlannerTaskDetailsInput = z.strictObject({
  task_id: z.number().int().positive().describe('Planner task ID'),
});

export const UpdatePlannerTaskDetailsInput = z.strictObject({
  task_id: z.number().int().positive().describe('Planner task ID'),
  description: z.string().optional().describe('Task description/notes'),
  checklist: z.record(z.string(), z.object({}).passthrough()).optional().describe('Checklist items. Keys are GUIDs, values have title (string) and isChecked (boolean)'),
  references: z.record(z.string(), z.object({}).passthrough()).optional().describe('Reference links. Keys are encoded URLs, values have alias (string) and type (string)'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListPlansParams = z.infer<typeof ListPlansInput>;
export type GetPlanParams = z.infer<typeof GetPlanInput>;
export type CreatePlanParams = z.infer<typeof CreatePlanInput>;
export type UpdatePlanParams = z.infer<typeof UpdatePlanInput>;
export type ListBucketsParams = z.infer<typeof ListBucketsInput>;
export type CreateBucketParams = z.infer<typeof CreateBucketInput>;
export type UpdateBucketParams = z.infer<typeof UpdateBucketInput>;
export type PrepareDeleteBucketParams = z.infer<typeof PrepareDeleteBucketInput>;
export type ConfirmDeleteBucketParams = z.infer<typeof ConfirmDeleteBucketInput>;
export type ListPlannerTasksParams = z.infer<typeof ListPlannerTasksInput>;
export type GetPlannerTaskParams = z.infer<typeof GetPlannerTaskInput>;
export type CreatePlannerTaskParams = z.infer<typeof CreatePlannerTaskInput>;
export type UpdatePlannerTaskParams = z.infer<typeof UpdatePlannerTaskInput>;
export type PrepareDeletePlannerTaskParams = z.infer<typeof PrepareDeletePlannerTaskInput>;
export type ConfirmDeletePlannerTaskParams = z.infer<typeof ConfirmDeletePlannerTaskInput>;
export type GetPlannerTaskDetailsParams = z.infer<typeof GetPlannerTaskDetailsInput>;
export type UpdatePlannerTaskDetailsParams = z.infer<typeof UpdatePlannerTaskDetailsInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface IPlannerRepository {
  listPlansAsync(): Promise<Array<{ id: number; title: string; owner: string; createdDateTime: string }>>;
  getPlanAsync(planId: number): Promise<{ id: number; title: string; owner: string; createdDateTime: string; etag: string }>;
  createPlanAsync(title: string, groupId: string): Promise<number>;
  updatePlanAsync(planId: number, updates: { title?: string }): Promise<void>;
  listBucketsAsync(planId: number): Promise<Array<{ id: number; name: string; planId: number; orderHint: string }>>;
  createBucketAsync(planId: number, name: string): Promise<number>;
  updateBucketAsync(bucketId: number, updates: { name?: string }): Promise<void>;
  deleteBucketAsync(bucketId: number): Promise<void>;
  listPlannerTasksAsync(planId: number): Promise<Array<{
    id: number; title: string; bucketId: number | null; assignees: string[];
    percentComplete: number; priority: number; startDateTime: string;
    dueDateTime: string; createdDateTime: string;
  }>>;
  getPlannerTaskAsync(taskId: number): Promise<{
    id: number; title: string; bucketId: number | null; assignees: string[];
    percentComplete: number; priority: number; startDateTime: string;
    dueDateTime: string; createdDateTime: string; conversationThreadId: string;
    orderHint: string; etag: string;
  }>;
  createPlannerTaskAsync(
    planId: number, title: string, bucketId?: number,
    assignments?: Record<string, object>, priority?: number,
    startDate?: string, dueDate?: string,
  ): Promise<number>;
  updatePlannerTaskAsync(taskId: number, updates: {
    title?: string; bucketId?: number; percentComplete?: number;
    priority?: number; startDate?: string; dueDate?: string;
    assignments?: Record<string, object>;
  }): Promise<void>;
  deletePlannerTaskAsync(taskId: number): Promise<void>;
  getPlannerTaskDetailsAsync(taskId: number): Promise<{
    id: number; description: string; checklist: Record<string, unknown>;
    references: Record<string, unknown>; etag: string;
  }>;
  updatePlannerTaskDetailsAsync(taskId: number, updates: {
    description?: string; checklist?: Record<string, object>;
    references?: Record<string, object>;
  }): Promise<void>;
}

// =============================================================================
// Planner Tools
// =============================================================================

/**
 * Microsoft Planner tools with ETag caching and two-phase delete approval.
 */
export class PlannerTools {
  constructor(
    private readonly repo: IPlannerRepository,
    private readonly tokenManager: ApprovalTokenManager,
  ) {}

  async listPlans(): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const plans = await this.repo.listPlansAsync();
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ plans }, null, 2),
      }],
    };
  }

  async getPlan(params: GetPlanParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const plan = await this.repo.getPlanAsync(params.plan_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ plan }, null, 2),
      }],
    };
  }

  async createPlan(params: CreatePlanParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const planId = await this.repo.createPlanAsync(params.title, params.group_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, plan_id: planId, message: 'Plan created' }, null, 2),
      }],
    };
  }

  async updatePlan(params: UpdatePlanParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const updates: { title?: string } = {};
    if (params.title != null) updates.title = params.title;
    await this.repo.updatePlanAsync(params.plan_id, updates);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Plan updated' }, null, 2),
      }],
    };
  }

  async listBuckets(params: ListBucketsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const buckets = await this.repo.listBucketsAsync(params.plan_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ buckets }, null, 2),
      }],
    };
  }

  async createBucket(params: CreateBucketParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const bucketId = await this.repo.createBucketAsync(params.plan_id, params.name);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, bucket_id: bucketId, message: 'Bucket created' }, null, 2),
      }],
    };
  }

  async updateBucket(params: UpdateBucketParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const updates: { name?: string } = {};
    if (params.name != null) updates.name = params.name;
    await this.repo.updateBucketAsync(params.bucket_id, updates);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Bucket updated' }, null, 2),
      }],
    };
  }

  prepareDeleteBucket(params: PrepareDeleteBucketParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'delete_bucket',
      targetType: 'bucket',
      targetId: params.bucket_id,
      targetHash: String(params.bucket_id),
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          bucket_id: params.bucket_id,
          action: `To confirm deleting bucket ${params.bucket_id}, call confirm_delete_bucket with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmDeleteBucket(params: ConfirmDeleteBucketParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const token = this.tokenManager.lookupToken(params.approval_token);
    if (token == null) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({
            success: false,
            error: 'Token not found or already used',
          }, null, 2),
        }],
      };
    }

    const result = this.tokenManager.consumeToken(params.approval_token, 'delete_bucket', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_bucket again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_bucket',
        TARGET_MISMATCH: 'Token was generated for a different bucket',
        ALREADY_CONSUMED: 'Token has already been used',
      };
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({
            success: false,
            error: errorMessages[result.error ?? ''] ?? 'Invalid token',
          }, null, 2),
        }],
      };
    }

    await this.repo.deleteBucketAsync(result.token!.targetId);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Bucket deleted' }, null, 2),
      }],
    };
  }

  // ===========================================================================
  // Planner Tasks
  // ===========================================================================

  async listPlannerTasks(params: ListPlannerTasksParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const tasks = await this.repo.listPlannerTasksAsync(params.plan_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ tasks }, null, 2),
      }],
    };
  }

  async getPlannerTask(params: GetPlannerTaskParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const task = await this.repo.getPlannerTaskAsync(params.task_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ task }, null, 2),
      }],
    };
  }

  async createPlannerTask(params: CreatePlannerTaskParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const taskId = await this.repo.createPlannerTaskAsync(
      params.plan_id,
      params.title,
      params.bucket_id,
      params.assignments,
      params.priority,
      params.start_date,
      params.due_date,
    );
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, task_id: taskId, message: 'Planner task created' }, null, 2),
      }],
    };
  }

  async updatePlannerTask(params: UpdatePlannerTaskParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const updates: {
      title?: string; bucketId?: number; percentComplete?: number;
      priority?: number; startDate?: string; dueDate?: string;
      assignments?: Record<string, object>;
    } = {};
    if (params.title != null) updates.title = params.title;
    if (params.bucket_id != null) updates.bucketId = params.bucket_id;
    if (params.percent_complete != null) updates.percentComplete = params.percent_complete;
    if (params.priority != null) updates.priority = params.priority;
    if (params.start_date != null) updates.startDate = params.start_date;
    if (params.due_date != null) updates.dueDate = params.due_date;
    if (params.assignments != null) updates.assignments = params.assignments;
    await this.repo.updatePlannerTaskAsync(params.task_id, updates);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Planner task updated' }, null, 2),
      }],
    };
  }

  prepareDeletePlannerTask(params: PrepareDeletePlannerTaskParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'delete_planner_task',
      targetType: 'planner_task',
      targetId: params.task_id,
      targetHash: String(params.task_id),
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          task_id: params.task_id,
          action: `To confirm deleting planner task ${params.task_id}, call confirm_delete_planner_task with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmDeletePlannerTask(params: ConfirmDeletePlannerTaskParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const token = this.tokenManager.lookupToken(params.approval_token);
    if (token == null) {
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({
            success: false,
            error: 'Token not found or already used',
          }, null, 2),
        }],
      };
    }

    const result = this.tokenManager.consumeToken(params.approval_token, 'delete_planner_task', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_planner_task again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_planner_task',
        TARGET_MISMATCH: 'Token was generated for a different task',
        ALREADY_CONSUMED: 'Token has already been used',
      };
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({
            success: false,
            error: errorMessages[result.error ?? ''] ?? 'Invalid token',
          }, null, 2),
        }],
      };
    }

    await this.repo.deletePlannerTaskAsync(result.token!.targetId);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Planner task deleted' }, null, 2),
      }],
    };
  }

  // ===========================================================================
  // Planner Task Details
  // ===========================================================================

  async getPlannerTaskDetails(params: GetPlannerTaskDetailsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const details = await this.repo.getPlannerTaskDetailsAsync(params.task_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ details }, null, 2),
      }],
    };
  }

  async updatePlannerTaskDetails(params: UpdatePlannerTaskDetailsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const updates: {
      description?: string;
      checklist?: Record<string, object>;
      references?: Record<string, object>;
    } = {};
    if (params.description != null) updates.description = params.description;
    if (params.checklist != null) updates.checklist = params.checklist;
    if (params.references != null) updates.references = params.references;
    await this.repo.updatePlannerTaskDetailsAsync(params.task_id, updates);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Planner task details updated' }, null, 2),
      }],
    };
  }
}
