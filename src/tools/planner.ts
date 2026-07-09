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
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset } from '../registry/context.js';
import { Id } from '../ids/schema.js';
import { nextActionFor } from '../ids/next-action.js';
import type { ToolContext, ToolDefinition } from '../registry/types.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    planner: PlannerTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const ListPlansInput = z.strictObject({});

export const GetPlanInput = z.strictObject({
  plan_id: Id.plan,
});

export const CreatePlanInput = z.strictObject({
  title: z.string().min(1).describe('Plan title'),
  group_id: z.string().min(1).describe('M365 group ID that owns the plan'),
});

export const UpdatePlanInput = z.strictObject({
  plan_id: Id.plan,
  title: z.string().min(1).optional().describe('New plan title'),
});

export const ListBucketsInput = z.strictObject({
  plan_id: Id.plan,
});

export const CreateBucketInput = z.strictObject({
  plan_id: Id.plan,
  name: z.string().min(1).describe('Bucket name'),
});

export const UpdateBucketInput = z.strictObject({
  bucket_id: Id.plannerBucket,
  name: z.string().min(1).optional().describe('New bucket name'),
});

export const PrepareDeleteBucketInput = z.strictObject({
  bucket_id: Id.plannerBucket,
});

export const ConfirmDeleteBucketInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_bucket'),
});

export const ListPlannerTasksInput = z.strictObject({
  plan_id: Id.plan,
});

export const ListMyPlannerTasksInput = z.strictObject({});

export const GetPlannerTaskInput = z.strictObject({
  task_id: Id.plannerTask,
});

export const CreatePlannerTaskInput = z.strictObject({
  plan_id: Id.plan,
  title: z.string().min(1).describe('Task title'),
  bucket_id: Id.plannerBucket.optional(),
  assignments: z.record(z.string(), z.object({}).passthrough()).optional().describe('User assignments. Keys are user IDs, values should be { "@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": " !" }'),
  priority: z.number().int().min(0).max(10).optional().describe('Priority (0-10)'),
  start_date: z.string().optional().describe('Start date in ISO format'),
  due_date: z.string().optional().describe('Due date in ISO format'),
});

export const UpdatePlannerTaskInput = z.strictObject({
  task_id: Id.plannerTask,
  title: z.string().min(1).optional().describe('New task title'),
  bucket_id: Id.plannerBucket.optional().describe('New bucket ID — a `pb_` token from list_buckets.'),
  percent_complete: z.number().int().min(0).max(100).optional().describe('Percent complete (0-100)'),
  priority: z.number().int().min(0).max(10).optional().describe('Priority (0-10)'),
  start_date: z.string().optional().describe('Start date in ISO format'),
  due_date: z.string().optional().describe('Due date in ISO format'),
  assignments: z.record(z.string(), z.object({}).passthrough()).optional().describe('User assignments. Keys are user IDs, values should be { "@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": " !" }'),
});

export const PrepareDeletePlannerTaskInput = z.strictObject({
  task_id: Id.plannerTask,
});

export const ConfirmDeletePlannerTaskInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_planner_task'),
});

export const GetPlannerTaskDetailsInput = z.strictObject({
  task_id: Id.plannerTask,
});

export const UpdatePlannerTaskDetailsInput = z.strictObject({
  task_id: Id.plannerTask,
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
export type ListMyPlannerTasksParams = z.infer<typeof ListMyPlannerTasksInput>;
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
  listPlansAsync(): Promise<Array<{ id: string; title: string; owner: string; createdDateTime: string }>>;
  getPlanAsync(planId: string | number): Promise<{ id: string; title: string; owner: string; createdDateTime: string; etag: string }>;
  createPlanAsync(title: string, groupId: string): Promise<string>;
  updatePlanAsync(planId: string | number, updates: { title?: string }): Promise<void>;
  listBucketsAsync(planId: string | number): Promise<Array<{ id: string; name: string; planId: string; orderHint: string }>>;
  createBucketAsync(planId: string | number, name: string): Promise<string>;
  updateBucketAsync(bucketId: string | number, updates: { name?: string }): Promise<void>;
  deleteBucketAsync(bucketId: string | number): Promise<void>;
  listPlannerTasksAsync(planId: string | number): Promise<Array<{
    id: string; title: string; bucketId: string | null; assignees: string[];
    percentComplete: number; priority: number; startDateTime: string;
    dueDateTime: string; createdDateTime: string;
  }>>;
  listMyPlannerTasksAsync(): Promise<Array<{
    id: string; title: string; planId: string; bucketId: string | null;
    assignees: string[]; percentComplete: number; priority: number;
    startDateTime: string; dueDateTime: string; createdDateTime: string;
  }>>;
  getPlannerTaskAsync(taskId: string | number): Promise<{
    id: string; title: string; bucketId: string | null; assignees: string[];
    percentComplete: number; priority: number; startDateTime: string;
    dueDateTime: string; createdDateTime: string; conversationThreadId: string;
    orderHint: string; etag: string;
  }>;
  createPlannerTaskAsync(
    planId: string | number, title: string, bucketId?: string | number,
    assignments?: Record<string, object>, priority?: number,
    startDate?: string, dueDate?: string,
  ): Promise<string>;
  updatePlannerTaskAsync(taskId: string | number, updates: {
    title?: string; bucketId?: string | number; percentComplete?: number;
    priority?: number; startDate?: string; dueDate?: string;
    assignments?: Record<string, object>;
  }): Promise<void>;
  deletePlannerTaskAsync(taskId: string | number): Promise<void>;
  getPlannerTaskDetailsAsync(taskId: string | number): Promise<{
    id: string; description: string; checklist: Record<string, unknown>;
    references: Record<string, unknown>; etag: string;
  }>;
  updatePlannerTaskDetailsAsync(taskId: string | number, updates: {
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
        text: JSON.stringify({ plans, next: nextActionFor('plan') ?? undefined }, null, 2),
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
        text: JSON.stringify({ success: true, plan_id: planId, message: 'Plan created', next: nextActionFor('plan') ?? undefined }, null, 2),
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
        text: JSON.stringify({ buckets, next: nextActionFor('plannerBucket') ?? undefined }, null, 2),
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
        text: JSON.stringify({ success: true, bucket_id: bucketId, message: 'Bucket created', next: nextActionFor('plannerBucket') ?? undefined }, null, 2),
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

    await this.repo.deleteBucketAsync((result.token!.targetId as string));
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
        text: JSON.stringify({ tasks, next: nextActionFor('plannerTask') ?? undefined }, null, 2),
      }],
    };
  }

  async listMyPlannerTasks(_params: ListMyPlannerTasksParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const tasks = await this.repo.listMyPlannerTasksAsync();
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ tasks, next: nextActionFor('plannerTask') ?? undefined }, null, 2),
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
        text: JSON.stringify({ success: true, task_id: taskId, message: 'Planner task created', next: nextActionFor('plannerTask') ?? undefined }, null, 2),
      }],
    };
  }

  async updatePlannerTask(params: UpdatePlannerTaskParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const updates: {
      title?: string; bucketId?: string; percentComplete?: number;
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

    await this.repo.deletePlannerTaskAsync((result.token!.targetId as string));
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

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2)
// =============================================================================

/**
 * Registry tool definitions for the planner domain.
 */
export function plannerToolDefinitions(): ToolDefinition[] {
  const tools = (ctx: ToolContext): PlannerTools => requireGraphToolset(ctx, 'planner');

  return [
    defineTool({
      name: 'list_plans',
      description: 'List all Planner plans the user has access to (Graph API)',
      input: ListPlansInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx) => tools(ctx).listPlans(),
    }),
    defineTool({
      name: 'get_plan',
      description: 'Get details for a specific Planner plan (Graph API)',
      input: GetPlanInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getPlan(params),
    }),
    defineTool({
      name: 'create_plan',
      description: 'Create a new Planner plan in a Microsoft 365 group (Graph API)',
      input: CreatePlanInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).createPlan(params),
    }),
    defineTool({
      name: 'update_plan',
      description: 'Update a Planner plan title (Graph API)',
      input: UpdatePlanInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).updatePlan(params),
    }),
    defineTool({
      name: 'list_buckets',
      description: 'List all buckets in a Planner plan (Graph API)',
      input: ListBucketsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listBuckets(params),
    }),
    defineTool({
      name: 'create_bucket',
      description: 'Create a new bucket in a Planner plan (Graph API)',
      input: CreateBucketInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).createBucket(params),
    }),
    defineTool({
      name: 'update_bucket',
      description: 'Update a Planner bucket name (Graph API)',
      input: UpdateBucketInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).updateBucket(params),
    }),
    defineTool({
      name: 'prepare_delete_bucket',
      description: 'Prepare to delete a Planner bucket. Returns an approval token. (Graph API)',
      input: PrepareDeleteBucketInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareDeleteBucket(params),
    }),
    defineTool({
      name: 'confirm_delete_bucket',
      description: 'Confirm deletion of a Planner bucket using the approval token from prepare_delete_bucket. (Graph API)',
      input: ConfirmDeleteBucketInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmDeleteBucket(params),
    }),
    defineTool({
      name: 'list_planner_tasks',
      description: 'List all tasks in a Planner plan (Graph API)',
      input: ListPlannerTasksInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listPlannerTasks(params),
    }),
    defineTool({
      name: 'list_my_planner_tasks',
      description: 'List all Planner tasks assigned to the signed-in user across every plan (Graph API)',
      input: ListMyPlannerTasksInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listMyPlannerTasks(params),
    }),
    defineTool({
      name: 'get_planner_task',
      description: 'Get details for a specific Planner task (Graph API)',
      input: GetPlannerTaskInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getPlannerTask(params),
    }),
    defineTool({
      name: 'create_planner_task',
      description: 'Create a new task in a Planner plan (Graph API)',
      input: CreatePlannerTaskInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).createPlannerTask(params),
    }),
    defineTool({
      name: 'update_planner_task',
      description: 'Update a Planner task (Graph API)',
      input: UpdatePlannerTaskInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).updatePlannerTask(params),
    }),
    defineTool({
      name: 'prepare_delete_planner_task',
      description: 'Prepare to delete a Planner task. Returns an approval token. (Graph API)',
      input: PrepareDeletePlannerTaskInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareDeletePlannerTask(params),
    }),
    defineTool({
      name: 'confirm_delete_planner_task',
      description: 'Confirm deletion of a Planner task using the approval token from prepare_delete_planner_task. (Graph API)',
      input: ConfirmDeletePlannerTaskInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmDeletePlannerTask(params),
    }),
    defineTool({
      name: 'get_planner_task_details',
      description: 'Get details for a Planner task (description, checklist, references). (Graph API)',
      input: GetPlannerTaskDetailsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getPlannerTaskDetails(params),
    }),
    defineTool({
      name: 'update_planner_task_details',
      description: 'Update details for a Planner task (description, checklist, references). Requires get_planner_task_details first for ETag. (Graph API)',
      input: UpdatePlannerTaskDetailsInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).updatePlannerTaskDetails(params),
    }),
  ];
}
