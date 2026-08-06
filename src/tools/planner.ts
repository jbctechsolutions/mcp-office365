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
import { approvalTokenLink } from '../registry/elicit-links.js';
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

/** Shared key schema for Planner's 25 label categories. */
const CategoryKey = z.string().regex(/^category([1-9]|1[0-9]|2[0-5])$/, 'Keys must be category1..category25');

const AppliedCategories = z.record(CategoryKey, z.boolean())
  .describe('Planner labels. Keys are category1..category25; true applies a label, false removes it; omitted keys are preserved on update. Label names come from get_plan_details.');

const OrderHint = z.string().min(1)
  .describe('Planner order hint: "<previous> <next>!" positions between neighbors (empty string for a missing neighbor); " !" appends. Never resend a service-returned hint verbatim (Graph 400).');

const ChecklistItems = z.record(z.string(), z.object({}).passthrough());

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

export const PrepareDeletePlanInput = z.strictObject({
  plan_id: Id.plan,
});

export const ConfirmDeletePlanInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_plan'),
});

export const GetPlanDetailsInput = z.strictObject({
  plan_id: Id.plan,
});

export const UpdatePlanDetailsInput = z.strictObject({
  plan_id: Id.plan,
  category_descriptions: z.record(CategoryKey, z.string().nullable())
    .optional().describe('Label display names. Keys are category1..category25; value is the new name, or null to reset to the default. Omitted keys are preserved.'),
});

export const UpdatePlanSharingInput = z.strictObject({
  plan_id: Id.plan,
  shared_with: z.record(z.string().min(1), z.boolean())
    .describe('Plan sharing map. Keys are Entra user GUIDs; true shares the plan with the user, false removes them — removal may revoke their plan access. Members of the owning M365 group keep access via group membership.'),
});

export const ListBucketsInput = z.strictObject({
  plan_id: Id.plan,
});

export const CreateBucketInput = z.strictObject({
  plan_id: Id.plan,
  name: z.string().min(1).describe('Bucket name'),
  order_hint: OrderHint.optional(),
});

export const UpdateBucketInput = z.strictObject({
  bucket_id: Id.plannerBucket,
  name: z.string().min(1).optional().describe('New bucket name'),
  order_hint: OrderHint.optional(),
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
  applied_categories: AppliedCategories.optional(),
  order_hint: OrderHint.optional(),
  percent_complete: z.number().int().min(0).max(100).optional().describe('Percent complete (0-100)'),
  description: z.string().optional().describe('Task description/notes — applied via a follow-up details write after creation'),
  checklist: ChecklistItems.optional().describe('Checklist items, applied via a follow-up details write. Keys are GUIDs, values have title (string) and isChecked (boolean)'),
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
  applied_categories: AppliedCategories.optional(),
  order_hint: OrderHint.optional(),
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
  checklist: ChecklistItems.optional().describe('Checklist items. Keys are GUIDs, values have title (string) and isChecked (boolean)'),
  references: z.record(z.string(), z.object({}).passthrough()).optional().describe('Reference links. Keys are encoded URLs, values have alias (string) and type (string)'),
});

export const ListPlannerTaskMessagesInput = z.strictObject({
  task_id: Id.plannerTask,
  skip_token: z.string().optional().describe('Paging token from a previous list_planner_task_messages response (`next_skip_token`)'),
});

export const CreatePlannerTaskMessageInput = z.strictObject({
  task_id: Id.plannerTask,
  content: z.string().min(1).describe('Comment text (plain text or sanitized HTML)'),
  mention_user_ids: z.array(z.string().min(1)).optional().describe('Entra user ids or email addresses to @mention. Mention HTML is built automatically for plain-text content.'),
});

export const UpdatePlannerTaskMessageInput = z.strictObject({
  message_id: Id.plannerTaskMessage,
  content: z.string().min(1).describe('Updated comment text (plain text or sanitized HTML)'),
  mention_user_ids: z.array(z.string().min(1)).optional().describe('Entra user ids or email addresses to @mention'),
});

export const PrepareDeletePlannerTaskMessageInput = z.strictObject({
  message_id: Id.plannerTaskMessage,
});

export const ConfirmDeletePlannerTaskMessageInput = z.strictObject({
  approval_token: z.string().describe('Approval token from prepare_delete_planner_task_message'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListPlansParams = z.infer<typeof ListPlansInput>;
export type GetPlanParams = z.infer<typeof GetPlanInput>;
export type CreatePlanParams = z.infer<typeof CreatePlanInput>;
export type UpdatePlanParams = z.infer<typeof UpdatePlanInput>;
export type PrepareDeletePlanParams = z.infer<typeof PrepareDeletePlanInput>;
export type ConfirmDeletePlanParams = z.infer<typeof ConfirmDeletePlanInput>;
export type GetPlanDetailsParams = z.infer<typeof GetPlanDetailsInput>;
export type UpdatePlanDetailsParams = z.infer<typeof UpdatePlanDetailsInput>;
export type UpdatePlanSharingParams = z.infer<typeof UpdatePlanSharingInput>;
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
export type ListPlannerTaskMessagesParams = z.infer<typeof ListPlannerTaskMessagesInput>;
export type CreatePlannerTaskMessageParams = z.infer<typeof CreatePlannerTaskMessageInput>;
export type UpdatePlannerTaskMessageParams = z.infer<typeof UpdatePlannerTaskMessageInput>;
export type PrepareDeletePlannerTaskMessageParams = z.infer<typeof PrepareDeletePlannerTaskMessageInput>;
export type ConfirmDeletePlannerTaskMessageParams = z.infer<typeof ConfirmDeletePlannerTaskMessageInput>;

// =============================================================================
// Repository Interface
// =============================================================================

export interface IPlannerRepository {
  listPlansAsync(): Promise<Array<{ id: string; title: string; owner: string; createdDateTime: string }>>;
  getPlanAsync(planId: string): Promise<{ id: string; title: string; owner: string; createdDateTime: string; etag: string }>;
  createPlanAsync(title: string, groupId: string): Promise<string>;
  updatePlanAsync(planId: string, updates: { title?: string }): Promise<void>;
  getPlanDetailsAsync(planId: string): Promise<{
    id: string;
    categoryDescriptions: Record<string, string | null>;
    sharedWith: Record<string, boolean>;
    etag: string;
  }>;
  updatePlanDetailsAsync(planId: string, updates: {
    categoryDescriptions?: Record<string, string | null>;
    sharedWith?: Record<string, boolean>;
  }): Promise<void>;
  updatePlanSharingAsync(planId: string, sharedWith: Record<string, boolean>): Promise<void>;
  deletePlanAsync(planId: string): Promise<void>;
  listBucketsAsync(planId: string): Promise<Array<{ id: string; name: string; planId: string; orderHint: string }>>;
  createBucketAsync(planId: string, name: string, orderHint?: string): Promise<string>;
  updateBucketAsync(bucketId: string, updates: { name?: string; orderHint?: string }): Promise<void>;
  deleteBucketAsync(bucketId: string): Promise<void>;
  listPlannerTasksAsync(planId: string): Promise<Array<{
    id: string; title: string; bucketId: string | null; assignees: string[];
    percentComplete: number; priority: number; startDateTime: string;
    dueDateTime: string; createdDateTime: string;
    appliedCategories: Record<string, boolean>;
  }>>;
  listMyPlannerTasksAsync(): Promise<Array<{
    id: string; title: string; planId: string; bucketId: string | null;
    assignees: string[]; percentComplete: number; priority: number;
    startDateTime: string; dueDateTime: string; createdDateTime: string;
    appliedCategories: Record<string, boolean>;
  }>>;
  getPlannerTaskAsync(taskId: string): Promise<{
    id: string; title: string; bucketId: string | null; assignees: string[];
    percentComplete: number; priority: number; startDateTime: string;
    dueDateTime: string; createdDateTime: string; conversationThreadId: string;
    orderHint: string; etag: string; appliedCategories: Record<string, boolean>;
  }>;
  createPlannerTaskAsync(planId: string, title: string, options: {
    bucketId?: string; assignments?: Record<string, object>; priority?: number;
    startDate?: string; dueDate?: string;
    appliedCategories?: Record<string, boolean>; orderHint?: string;
    percentComplete?: number; description?: string;
    checklist?: Record<string, object>;
  }): Promise<{ taskId: string; detailsError?: string }>;
  updatePlannerTaskAsync(taskId: string, updates: {
    title?: string; bucketId?: string; percentComplete?: number;
    priority?: number; startDate?: string; dueDate?: string;
    assignments?: Record<string, object>;
    appliedCategories?: Record<string, boolean>; orderHint?: string;
  }): Promise<void>;
  deletePlannerTaskAsync(taskId: string): Promise<void>;
  getPlannerTaskDetailsAsync(taskId: string): Promise<{
    id: string; description: string; checklist: Record<string, unknown>;
    references: Record<string, unknown>; etag: string;
  }>;
  updatePlannerTaskDetailsAsync(taskId: string, updates: {
    description?: string; checklist?: Record<string, object>;
    references?: Record<string, object>;
  }): Promise<void>;
  listPlannerTaskMessagesAsync(taskId: string, skipToken?: string): Promise<{
    messages: Array<{
      id: string; content: string; messageType: string; createdDateTime: string;
      editedTime: string | null; deletedTime: string | null; createdByUserId: string;
      mentions: unknown[];
    }>;
    nextSkipToken?: string;
  }>;
  createPlannerTaskMessageAsync(taskId: string, content: string, mentionUserIds?: string[]): Promise<string>;
  updatePlannerTaskMessageAsync(messageId: string, content: string, mentionUserIds?: string[]): Promise<void>;
  deletePlannerTaskMessageAsync(messageId: string): Promise<void>;
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

  prepareDeletePlan(params: PrepareDeletePlanParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'delete_plan',
      targetType: 'plan',
      targetId: params.plan_id,
      targetHash: String(params.plan_id),
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          plan_id: params.plan_id,
          action: `To confirm deleting plan ${params.plan_id} — including ALL buckets and tasks it contains — call confirm_delete_plan with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmDeletePlan(params: ConfirmDeletePlanParams): Promise<{
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

    const result = this.tokenManager.consumeToken(params.approval_token, 'delete_plan', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_plan again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_plan',
        TARGET_MISMATCH: 'Token was generated for a different plan',
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

    await this.repo.deletePlanAsync(result.token!.targetId);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Plan deleted' }, null, 2),
      }],
    };
  }

  async getPlanDetails(params: GetPlanDetailsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const details = await this.repo.getPlanDetailsAsync(params.plan_id);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ details }, null, 2),
      }],
    };
  }

  async updatePlanDetails(params: UpdatePlanDetailsParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const updates: { categoryDescriptions?: Record<string, string | null> } = {};
    if (params.category_descriptions != null) updates.categoryDescriptions = params.category_descriptions;
    await this.repo.updatePlanDetailsAsync(params.plan_id, updates);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Plan details updated' }, null, 2),
      }],
    };
  }

  async updatePlanSharing(params: UpdatePlanSharingParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    await this.repo.updatePlanSharingAsync(params.plan_id, params.shared_with);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Plan sharing updated' }, null, 2),
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
    const bucketId = await this.repo.createBucketAsync(params.plan_id, params.name, params.order_hint);
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
    const updates: { name?: string; orderHint?: string } = {};
    if (params.name != null) updates.name = params.name;
    if (params.order_hint != null) updates.orderHint = params.order_hint;
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

    await this.repo.deleteBucketAsync((result.token!.targetId));
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
    const options: Parameters<IPlannerRepository['createPlannerTaskAsync']>[2] = {};
    if (params.bucket_id != null) options.bucketId = params.bucket_id;
    if (params.assignments != null) options.assignments = params.assignments;
    if (params.priority != null) options.priority = params.priority;
    if (params.start_date != null) options.startDate = params.start_date;
    if (params.due_date != null) options.dueDate = params.due_date;
    if (params.applied_categories != null) options.appliedCategories = params.applied_categories;
    if (params.order_hint != null) options.orderHint = params.order_hint;
    if (params.percent_complete != null) options.percentComplete = params.percent_complete;
    if (params.description != null) options.description = params.description;
    if (params.checklist != null) options.checklist = params.checklist;
    const { taskId, detailsError } = await this.repo.createPlannerTaskAsync(params.plan_id, params.title, options);
    const detailsWarning = detailsError != null
      ? `Task created, but the description/checklist could not be applied (${detailsError}). ` +
        `Retry via update_planner_task_details with task_id ${taskId}.`
      : undefined;
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          success: true,
          task_id: taskId,
          message: 'Planner task created',
          details_warning: detailsWarning,
          next: nextActionFor('plannerTask') ?? undefined,
        }, null, 2),
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
      appliedCategories?: Record<string, boolean>; orderHint?: string;
    } = {};
    if (params.title != null) updates.title = params.title;
    if (params.bucket_id != null) updates.bucketId = params.bucket_id;
    if (params.percent_complete != null) updates.percentComplete = params.percent_complete;
    if (params.priority != null) updates.priority = params.priority;
    if (params.start_date != null) updates.startDate = params.start_date;
    if (params.due_date != null) updates.dueDate = params.due_date;
    if (params.assignments != null) updates.assignments = params.assignments;
    if (params.applied_categories != null) updates.appliedCategories = params.applied_categories;
    if (params.order_hint != null) updates.orderHint = params.order_hint;
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

    await this.repo.deletePlannerTaskAsync((result.token!.targetId));
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

  // ===========================================================================
  // Planner Task Chat Messages (beta)
  // ===========================================================================

  async listPlannerTaskMessages(params: ListPlannerTaskMessagesParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const { messages, nextSkipToken } = await this.repo.listPlannerTaskMessagesAsync(
      params.task_id,
      params.skip_token,
    );
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          messages,
          next_skip_token: nextSkipToken,
          next: nextActionFor('plannerTaskMessage') ?? undefined,
        }, null, 2),
      }],
    };
  }

  async createPlannerTaskMessage(params: CreatePlannerTaskMessageParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    const messageId = await this.repo.createPlannerTaskMessageAsync(
      params.task_id,
      params.content,
      params.mention_user_ids,
    );
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          success: true,
          message_id: messageId,
          message: 'Planner task comment posted',
          next: nextActionFor('plannerTaskMessage') ?? undefined,
        }, null, 2),
      }],
    };
  }

  async updatePlannerTaskMessage(params: UpdatePlannerTaskMessageParams): Promise<{
    content: Array<{ type: 'text'; text: string }>;
  }> {
    await this.repo.updatePlannerTaskMessageAsync(
      params.message_id,
      params.content,
      params.mention_user_ids,
    );
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Planner task comment updated' }, null, 2),
      }],
    };
  }

  prepareDeletePlannerTaskMessage(params: PrepareDeletePlannerTaskMessageParams): {
    content: Array<{ type: 'text'; text: string }>;
  } {
    const token = this.tokenManager.generateToken({
      operation: 'delete_planner_task_message',
      targetType: 'planner_task_message',
      targetId: params.message_id,
      targetHash: String(params.message_id),
    });

    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({
          approval_token: token.tokenId,
          expires_at: new Date(token.expiresAt).toISOString(),
          message_id: params.message_id,
          action: `To confirm deleting planner task comment ${params.message_id}, call confirm_delete_planner_task_message with the approval_token.`,
        }, null, 2),
      }],
    };
  }

  async confirmDeletePlannerTaskMessage(params: ConfirmDeletePlannerTaskMessageParams): Promise<{
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

    const result = this.tokenManager.consumeToken(params.approval_token, 'delete_planner_task_message', token.targetId);
    if (!result.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_planner_task_message again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_planner_task_message',
        TARGET_MISMATCH: 'Token was generated for a different comment',
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

    await this.repo.deletePlannerTaskMessageAsync(result.token!.targetId);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ success: true, message: 'Planner task comment deleted' }, null, 2),
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
      name: 'prepare_delete_plan',
      description: 'Prepare to delete a Planner plan INCLUDING all its buckets and tasks. Returns an approval token. (Graph API)',
      input: PrepareDeletePlanInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareDeletePlan(params),
      onElicit: approvalTokenLink('confirm_delete_plan'),
    }),
    defineTool({
      name: 'confirm_delete_plan',
      description: 'Confirm deletion of a Planner plan (and all contained buckets/tasks) using the approval token from prepare_delete_plan. (Graph API)',
      input: ConfirmDeletePlanInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmDeletePlan(params),
    }),
    defineTool({
      name: 'get_plan_details',
      description: 'Get Planner plan details: category label names (categoryDescriptions) and who the plan is shared with (Graph API)',
      input: GetPlanDetailsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getPlanDetails(params),
    }),
    defineTool({
      name: 'update_plan_details',
      description: 'Update Planner plan category label names (category1..category25; null resets a name to default) (Graph API)',
      input: UpdatePlanDetailsInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).updatePlanDetails(params),
    }),
    defineTool({
      name: 'update_plan_sharing',
      description: 'Share or unshare a Planner plan (sharedWith user GUIDs; true adds, false removes — removal may revoke plan access; owning-group members keep access via membership) (Graph API)',
      input: UpdatePlanSharingInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).updatePlanSharing(params),
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
      onElicit: approvalTokenLink('confirm_delete_bucket'),
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
      onElicit: approvalTokenLink('confirm_delete_planner_task'),
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
    defineTool({
      name: 'list_planner_task_messages',
      description: 'List chat comments on a Planner task (Comments tab). Uses Graph beta; delegated permissions only. (Graph API beta)',
      input: ListPlannerTaskMessagesInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listPlannerTaskMessages(params),
    }),
    defineTool({
      name: 'create_planner_task_message',
      description: 'Post a comment on a Planner task with optional @mentions (user ids or emails). Uses Graph beta; delegated permissions only. Requires Tasks.ReadWrite. (Graph API beta)',
      input: CreatePlannerTaskMessageInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).createPlannerTaskMessage(params),
    }),
    defineTool({
      name: 'update_planner_task_message',
      description: 'Update a Planner task comment. Uses Graph beta; delegated permissions only. Requires Tasks.ReadWrite. (Graph API beta)',
      input: UpdatePlannerTaskMessageInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).updatePlannerTaskMessage(params),
    }),
    defineTool({
      name: 'prepare_delete_planner_task_message',
      description: 'Prepare to delete a Planner task comment. Returns an approval token. Uses Graph beta. (Graph API beta)',
      input: PrepareDeletePlannerTaskMessageInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).prepareDeletePlannerTaskMessage(params),
      onElicit: approvalTokenLink('confirm_delete_planner_task_message'),
    }),
    defineTool({
      name: 'confirm_delete_planner_task_message',
      description: 'Confirm deletion of a Planner task comment using the approval token from prepare_delete_planner_task_message. (Graph API beta)',
      input: ConfirmDeletePlannerTaskMessageInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['planner'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).confirmDeletePlannerTaskMessage(params),
    }),
  ];
}
