/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Task Attachments tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { TaskAttachmentsTools, type ITaskAttachmentsRepository } from '../../../src/tools/task-attachments.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';

describe('TaskAttachmentsTools', () => {
  let repo: ITaskAttachmentsRepository;
  let tokenManager: ApprovalTokenManager;
  let tools: TaskAttachmentsTools;

  beforeEach(() => {
    repo = {
      listTaskAttachmentsAsync: vi.fn(),
      createTaskAttachmentAsync: vi.fn(),
      deleteTaskAttachmentAsync: vi.fn(),
    };
    tokenManager = new ApprovalTokenManager();
    tools = new TaskAttachmentsTools(repo, tokenManager);
  });

  describe('listTaskAttachments', () => {
    it('returns task attachments from the repository', async () => {
      const mockItems = [
        { id: 1, name: 'file1.pdf', size: 1024, contentType: 'application/pdf' },
        { id: 2, name: 'image.png', size: 2048, contentType: 'image/png' },
      ];
      vi.mocked(repo.listTaskAttachmentsAsync).mockResolvedValue(mockItems);

      const result = await tools.listTaskAttachments({ task_id: 42 });

      expect(repo.listTaskAttachmentsAsync).toHaveBeenCalledWith(42);
      expect(result.content).toHaveLength(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.task_attachments).toEqual(mockItems);
    });
  });

  describe('createTaskAttachment', () => {
    it('creates a task attachment and returns the ID', async () => {
      vi.mocked(repo.createTaskAttachmentAsync).mockResolvedValue(100);

      const result = await tools.createTaskAttachment({
        task_id: 42,
        name: 'document.pdf',
        content_bytes: 'dGVzdA==',
      });

      expect(repo.createTaskAttachmentAsync).toHaveBeenCalledWith(42, 'document.pdf', 'dGVzdA==', undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.task_attachment_id).toBe(100);
      expect(parsed.message).toBe('Task attachment created');
    });

    it('passes content_type when provided', async () => {
      vi.mocked(repo.createTaskAttachmentAsync).mockResolvedValue(101);

      await tools.createTaskAttachment({
        task_id: 42,
        name: 'image.png',
        content_bytes: 'iVBORw0KGgo=',
        content_type: 'image/png',
      });

      expect(repo.createTaskAttachmentAsync).toHaveBeenCalledWith(42, 'image.png', 'iVBORw0KGgo=', 'image/png');
    });
  });

  describe('prepareDeleteTaskAttachment', () => {
    it('generates an approval token', () => {
      const result = tools.prepareDeleteTaskAttachment({ task_attachment_id: 100 });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.task_attachment_id).toBe(100);
      expect(parsed.action).toContain('confirm_delete_task_attachment');
    });
  });

  describe('confirmDeleteTaskAttachment', () => {
    it('deletes the task attachment with a valid token', async () => {
      vi.mocked(repo.deleteTaskAttachmentAsync).mockResolvedValue(undefined);

      // Generate a token first
      const prepareResult = tools.prepareDeleteTaskAttachment({ task_attachment_id: 100 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeleteTaskAttachment({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Task attachment deleted');
      expect(repo.deleteTaskAttachmentAsync).toHaveBeenCalledWith(100);
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmDeleteTaskAttachment({ approval_token: 'invalid-token' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBe('Token not found or already used');
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.deleteTaskAttachmentAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeleteTaskAttachment({ task_attachment_id: 100 });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      // Consume the token
      await tools.confirmDeleteTaskAttachment({ approval_token });

      // Try to consume again
      const result = await tools.confirmDeleteTaskAttachment({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });
});
