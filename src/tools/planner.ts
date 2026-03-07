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
}
