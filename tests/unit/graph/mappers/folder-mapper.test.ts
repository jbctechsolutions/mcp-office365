/**
 * Tests for Graph folder mapper functions.
 */

import { describe, it, expect } from 'vitest';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import {
  mapMailFolderToRow,
  mapCalendarToFolderRow,
  mapTaskListToFolderRow,
} from '../../../../src/graph/mappers/folder-mapper.js';
import { hashStringToNumber } from '../../../../src/graph/mappers/utils.js';

describe('graph/mappers/folder-mapper', () => {
  describe('mapMailFolderToRow', () => {
    it('maps mail folder with all fields', () => {
      const folder: MicrosoftGraph.MailFolder = {
        id: 'folder-123',
        displayName: 'Inbox',
        parentFolderId: 'parent-456',
        totalItemCount: 100,
        unreadItemCount: 5,
      };

      const result = mapMailFolderToRow(folder);

      expect(result.id).toBe(hashStringToNumber('folder-123'));
      expect(result.name).toBe('Inbox');
      expect(result.parentId).toBe(hashStringToNumber('parent-456'));
      expect(result.messageCount).toBe(100);
      expect(result.unreadCount).toBe(5);
      expect(result.folderType).toBe(1); // Mail folder
      expect(result.specialType).toBe(0);
      expect(result.accountId).toBe(1);
    });

    it('handles folder with null id', () => {
      const folder: MicrosoftGraph.MailFolder = {
        id: undefined,
        displayName: 'Test',
      };

      const result = mapMailFolderToRow(folder);

      expect(result.id).toBe(hashStringToNumber(''));
    });

    it('handles folder with null displayName', () => {
      const folder: MicrosoftGraph.MailFolder = {
        id: 'folder-123',
        displayName: undefined,
      };

      const result = mapMailFolderToRow(folder);

      expect(result.name).toBeNull();
    });

    it('handles folder without parentFolderId', () => {
      const folder: MicrosoftGraph.MailFolder = {
        id: 'folder-123',
        displayName: 'Root Folder',
        parentFolderId: undefined,
      };

      const result = mapMailFolderToRow(folder);

      expect(result.parentId).toBeNull();
    });

    it('handles folder with zero counts', () => {
      const folder: MicrosoftGraph.MailFolder = {
        id: 'folder-123',
        displayName: 'Empty',
        totalItemCount: 0,
        unreadItemCount: 0,
      };

      const result = mapMailFolderToRow(folder);

      expect(result.messageCount).toBe(0);
      expect(result.unreadCount).toBe(0);
    });

    it('defaults counts to 0 when undefined', () => {
      const folder: MicrosoftGraph.MailFolder = {
        id: 'folder-123',
        displayName: 'No Counts',
        totalItemCount: undefined,
        unreadItemCount: undefined,
      };

      const result = mapMailFolderToRow(folder);

      expect(result.messageCount).toBe(0);
      expect(result.unreadCount).toBe(0);
    });
  });

  describe('mapCalendarToFolderRow', () => {
    it('maps calendar with all fields', () => {
      const calendar: MicrosoftGraph.Calendar = {
        id: 'calendar-123',
        name: 'My Calendar',
        color: 'blue',
        isDefaultCalendar: true,
        canEdit: true,
      };

      const result = mapCalendarToFolderRow(calendar);

      expect(result.id).toBe(hashStringToNumber('calendar-123'));
      expect(result.name).toBe('My Calendar');
      expect(result.parentId).toBeNull();
      expect(result.folderType).toBe(2); // Calendar folder
      expect(result.specialType).toBe(0);
      expect(result.accountId).toBe(1);
      expect(result.messageCount).toBe(0);
      expect(result.unreadCount).toBe(0);
    });

    it('handles calendar with null id', () => {
      const calendar: MicrosoftGraph.Calendar = {
        id: undefined,
        name: 'Test Calendar',
      };

      const result = mapCalendarToFolderRow(calendar);

      expect(result.id).toBe(hashStringToNumber(''));
    });

    it('handles calendar with null name', () => {
      const calendar: MicrosoftGraph.Calendar = {
        id: 'calendar-123',
        name: undefined,
      };

      const result = mapCalendarToFolderRow(calendar);

      expect(result.name).toBeNull();
    });
  });

  describe('mapTaskListToFolderRow', () => {
    it('maps task list with all fields', () => {
      const taskList: MicrosoftGraph.TodoTaskList = {
        id: 'tasklist-123',
        displayName: 'My Tasks',
        isOwner: true,
        isShared: false,
        wellknownListName: 'defaultList',
      };

      const result = mapTaskListToFolderRow(taskList);

      expect(result.id).toBe(hashStringToNumber('tasklist-123'));
      expect(result.name).toBe('My Tasks');
      expect(result.parentId).toBeNull();
      expect(result.folderType).toBe(3); // Task folder
      expect(result.specialType).toBe(0);
      expect(result.accountId).toBe(1);
      expect(result.messageCount).toBe(0);
      expect(result.unreadCount).toBe(0);
    });

    it('handles task list with null id', () => {
      const taskList: MicrosoftGraph.TodoTaskList = {
        id: undefined,
        displayName: 'Test List',
      };

      const result = mapTaskListToFolderRow(taskList);

      expect(result.id).toBe(hashStringToNumber(''));
    });

    it('handles task list with null displayName', () => {
      const taskList: MicrosoftGraph.TodoTaskList = {
        id: 'tasklist-123',
        displayName: undefined,
      };

      const result = mapTaskListToFolderRow(taskList);

      expect(result.name).toBeNull();
    });
  });
});
