/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for calendar permission tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { CalendarPermissionsTools, type ICalendarPermissionsRepository, PrepareDeleteCalendarPermissionInput } from '../../../src/tools/calendar-permissions.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';

describe('CalendarPermissionsTools schema', () => {
  it('accepts a cp_ permission token and rejects a legacy numeric permission id', () => {
    // Calendar permissions now emit cp_ tokens; a numeric permission_id must
    // fail validation so the producer/consumer contract can't silently desync.
    expect(PrepareDeleteCalendarPermissionInput.safeParse({ permission_id: 'cp_AbC' }).success).toBe(true);
    expect(PrepareDeleteCalendarPermissionInput.safeParse({ permission_id: 42 }).success).toBe(false);
  });
});

describe('CalendarPermissionsTools', () => {
  let repo: ICalendarPermissionsRepository;
  let tokenManager: ApprovalTokenManager;
  let tools: CalendarPermissionsTools;

  beforeEach(() => {
    repo = {
      listCalendarPermissionsAsync: vi.fn(),
      createCalendarPermissionAsync: vi.fn(),
      deleteCalendarPermissionAsync: vi.fn(),
    };
    tokenManager = new ApprovalTokenManager();
    tools = new CalendarPermissionsTools(repo, tokenManager);
  });

  describe('listCalendarPermissions', () => {
    it('returns permissions from the repository', async () => {
      const mockPermissions = [
        { id: 'cp_1', emailAddress: 'alice@example.com', role: 'read', isRemovable: true, isInsideOrganization: true },
        { id: 'cp_2', emailAddress: 'bob@example.com', role: 'write', isRemovable: true, isInsideOrganization: false },
      ];
      vi.mocked(repo.listCalendarPermissionsAsync).mockResolvedValue(mockPermissions);

      const result = await tools.listCalendarPermissions({ calendar_id: 'fd_10' });

      expect(repo.listCalendarPermissionsAsync).toHaveBeenCalledWith('fd_10');
      expect(result.content).toHaveLength(1);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.permissions).toEqual(mockPermissions);
    });
  });

  describe('createCalendarPermission', () => {
    it('creates a permission and returns the ID', async () => {
      vi.mocked(repo.createCalendarPermissionAsync).mockResolvedValue('cp_42');

      const result = await tools.createCalendarPermission({
        calendar_id: 'fd_10',
        email_address: 'alice@example.com',
        role: 'read',
      });

      expect(repo.createCalendarPermissionAsync).toHaveBeenCalledWith('fd_10', 'alice@example.com', 'read');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.permission_id).toBe('cp_42');
      expect(parsed.message).toBe('Calendar permission created');
    });
  });

  describe('prepareDeleteCalendarPermission', () => {
    it('generates an approval token', () => {
      const result = tools.prepareDeleteCalendarPermission({ permission_id: 'cp_42' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.approval_token).toBeDefined();
      expect(typeof parsed.approval_token).toBe('string');
      expect(parsed.expires_at).toBeDefined();
      expect(parsed.permission_id).toBe('cp_42');
      expect(parsed.action).toContain('confirm_delete_calendar_permission');
    });
  });

  describe('confirmDeleteCalendarPermission', () => {
    it('deletes the permission with a valid token', async () => {
      vi.mocked(repo.deleteCalendarPermissionAsync).mockResolvedValue(undefined);

      // Generate a token first
      const prepareResult = tools.prepareDeleteCalendarPermission({ permission_id: 'cp_42' });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      const result = await tools.confirmDeleteCalendarPermission({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Calendar permission deleted');
      expect(repo.deleteCalendarPermissionAsync).toHaveBeenCalledWith('cp_42');
    });

    it('returns error for invalid token', async () => {
      const result = await tools.confirmDeleteCalendarPermission({
        approval_token: '00000000-0000-0000-0000-000000000000',
      });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
      expect(parsed.error).toBeDefined();
      expect(repo.deleteCalendarPermissionAsync).not.toHaveBeenCalled();
    });

    it('returns error for already consumed token', async () => {
      vi.mocked(repo.deleteCalendarPermissionAsync).mockResolvedValue(undefined);

      const prepareResult = tools.prepareDeleteCalendarPermission({ permission_id: 'cp_42' });
      const { approval_token } = JSON.parse(prepareResult.content[0].text);

      // Consume the token
      await tools.confirmDeleteCalendarPermission({ approval_token });

      // Try to use it again
      const result = await tools.confirmDeleteCalendarPermission({ approval_token });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });
});
