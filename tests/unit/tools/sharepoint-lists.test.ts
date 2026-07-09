/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for SharePoint Lists tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { SharePointListsTools, type ISharePointListsRepository } from '../../../src/tools/sharepoint-lists.js';
import { ApprovalTokenManager } from '../../../src/approval/index.js';

describe('SharePointListsTools', () => {
  let repo: ISharePointListsRepository;
  let tokenManager: ApprovalTokenManager;
  let tools: SharePointListsTools;

  beforeEach(() => {
    repo = {
      listSharePointListsAsync: vi.fn(),
      getSharePointListAsync: vi.fn(),
      createSharePointListAsync: vi.fn(),
      listSharePointListColumnsAsync: vi.fn(),
      listSharePointListItemsAsync: vi.fn(),
      getSharePointListItemAsync: vi.fn(),
      createSharePointListItemAsync: vi.fn(),
      updateSharePointListItemAsync: vi.fn(),
      deleteSharePointListItemAsync: vi.fn(),
    };
    tokenManager = new ApprovalTokenManager();
    tools = new SharePointListsTools(repo, tokenManager);
  });

  describe('listLists', () => {
    it('returns lists from the repository', async () => {
      const mockLists = [
        { id: 'sl_aaa', name: 'a', displayName: 'Announcements', description: '', webUrl: '' },
      ];
      vi.mocked(repo.listSharePointListsAsync).mockResolvedValue(mockLists);

      const result = await tools.listLists({ site_id: 'si_x' });

      expect(repo.listSharePointListsAsync).toHaveBeenCalledWith('si_x');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.lists).toEqual(mockLists);
    });
  });

  describe('getList', () => {
    it('returns a single list', async () => {
      const list = { id: 'sl_x', name: 'a', displayName: 'A', description: '', webUrl: '' };
      vi.mocked(repo.getSharePointListAsync).mockResolvedValue(list);

      const result = await tools.getList({ list_id: 'sl_x' });

      expect(repo.getSharePointListAsync).toHaveBeenCalledWith('sl_x');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.list).toEqual(list);
    });
  });

  describe('createList', () => {
    it('creates a list and returns the minted token', async () => {
      vi.mocked(repo.createSharePointListAsync).mockResolvedValue('sl_new');

      const result = await tools.createList({ site_id: 'si_x', display_name: 'My List', description: 'd' });

      expect(repo.createSharePointListAsync).toHaveBeenCalledWith('si_x', 'My List', 'd');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.id).toBe('sl_new');
      expect(parsed.status).toBe('created');
    });
  });

  describe('listListColumns', () => {
    it('returns column metadata', async () => {
      const cols = [{ id: 'c1', name: 'Title', displayName: 'Title', columnType: 'text', required: true, readOnly: false }];
      vi.mocked(repo.listSharePointListColumnsAsync).mockResolvedValue(cols);

      const result = await tools.listListColumns({ list_id: 'sl_x' });

      expect(repo.listSharePointListColumnsAsync).toHaveBeenCalledWith('sl_x');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.columns).toEqual(cols);
    });
  });

  describe('listListItems', () => {
    it('passes the limit through and returns items', async () => {
      const items = [{ id: 'sn_1', fields: { Title: 'First' }, webUrl: '', createdDateTime: '', lastModifiedDateTime: '' }];
      vi.mocked(repo.listSharePointListItemsAsync).mockResolvedValue(items);

      const result = await tools.listListItems({ list_id: 'sl_x', limit: 25 });

      expect(repo.listSharePointListItemsAsync).toHaveBeenCalledWith('sl_x', 25);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.items).toEqual(items);
    });
  });

  describe('getListItem', () => {
    it('returns a single item', async () => {
      const item = { id: 'sn_1', fields: { Title: 'First' }, webUrl: '', createdDateTime: '', lastModifiedDateTime: '' };
      vi.mocked(repo.getSharePointListItemAsync).mockResolvedValue(item);

      const result = await tools.getListItem({ item_id: 'sn_1' });

      expect(repo.getSharePointListItemAsync).toHaveBeenCalledWith('sn_1');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.item).toEqual(item);
    });
  });

  describe('createListItem', () => {
    it('creates an item and returns the minted token', async () => {
      vi.mocked(repo.createSharePointListItemAsync).mockResolvedValue('sn_new');

      const result = await tools.createListItem({ list_id: 'sl_x', fields: { Title: 'New' } });

      expect(repo.createSharePointListItemAsync).toHaveBeenCalledWith('sl_x', { Title: 'New' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.id).toBe('sn_new');
      expect(parsed.status).toBe('created');
    });
  });

  describe('updateListItem', () => {
    it('updates the item fields', async () => {
      const result = await tools.updateListItem({ item_id: 'sn_1', fields: { Title: 'Updated' } });

      expect(repo.updateSharePointListItemAsync).toHaveBeenCalledWith('sn_1', { Title: 'Updated' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.status).toBe('updated');
    });
  });

  describe('delete list item (two-phase)', () => {
    it('prepare returns a token; confirm deletes with it', async () => {
      const prepared = tools.prepareDeleteListItem({ item_id: 'sn_1' });
      const preparedParsed = JSON.parse(prepared.content[0].text);
      expect(preparedParsed.token_id).toBeDefined();
      expect(preparedParsed.item_id).toBe('sn_1');

      const confirmed = await tools.confirmDeleteListItem({ token_id: preparedParsed.token_id, item_id: 'sn_1' });

      expect(repo.deleteSharePointListItemAsync).toHaveBeenCalledWith('sn_1');
      const confirmedParsed = JSON.parse(confirmed.content[0].text);
      expect(confirmedParsed.success).toBe(true);
    });

    it('confirm rejects an invalid token without deleting', async () => {
      const result = await tools.confirmDeleteListItem({
        token_id: '00000000-0000-0000-0000-000000000000',
        item_id: 'sn_1',
      });

      expect(repo.deleteSharePointListItemAsync).not.toHaveBeenCalled();
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });

    it('confirm rejects a token minted for a different item', async () => {
      const prepared = tools.prepareDeleteListItem({ item_id: 'sn_1' });
      const preparedParsed = JSON.parse(prepared.content[0].text);

      const result = await tools.confirmDeleteListItem({ token_id: preparedParsed.token_id, item_id: 'sn_2' });

      expect(repo.deleteSharePointListItemAsync).not.toHaveBeenCalled();
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(false);
    });
  });
});
