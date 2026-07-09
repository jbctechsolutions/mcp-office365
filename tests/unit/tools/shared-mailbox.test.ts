/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for shared-mailbox / delegate-access tools (#40).
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { SharedMailboxTools, type ISharedMailboxClient } from '../../../src/tools/shared-mailbox.js';
import { mintSelfEncoded } from '../../../src/ids/token.js';

describe('SharedMailboxTools', () => {
  let client: ISharedMailboxClient;
  let tools: SharedMailboxTools;

  beforeEach(() => {
    vi.clearAllMocks();
    client = {
      listSharedMailFolders: vi.fn(),
      listSharedMessages: vi.fn(),
      getSharedMessage: vi.fn(),
      searchSharedMessages: vi.fn(),
      listSharedEvents: vi.fn(),
      getSharedEvent: vi.fn(),
      listSharedDriveItems: vi.fn(),
      searchSharedDriveItems: vi.fn(),
    };
    tools = new SharedMailboxTools(client);
  });

  describe('listFolders', () => {
    it('maps folders and echoes the mailbox', async () => {
      vi.mocked(client.listSharedMailFolders).mockResolvedValue([
        { id: 'AAA', displayName: 'Inbox', parentFolderId: 'root', totalItemCount: 10, unreadItemCount: 3 },
        { id: 'BBB', displayName: null, parentFolderId: null, totalItemCount: null, unreadItemCount: null },
      ]);

      const result = await tools.listFolders({ mailbox: 'shared@example.com' });

      expect(client.listSharedMailFolders).toHaveBeenCalledWith('shared@example.com');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.mailbox).toBe('shared@example.com');
      expect(parsed.folders[0]).toEqual({ id: 'AAA', name: 'Inbox', parentFolderId: 'root', totalItemCount: 10, unreadItemCount: 3 });
      expect(parsed.folders[1]).toEqual({ id: 'BBB', name: null, parentFolderId: null, totalItemCount: 0, unreadItemCount: 0 });
    });
  });

  describe('listEmails', () => {
    it('maps messages with default limit and no folder', async () => {
      vi.mocked(client.listSharedMessages).mockResolvedValue([
        {
          id: 'msg1',
          subject: 'Hello',
          from: { emailAddress: { name: 'Alice', address: 'alice@example.com' } },
          toRecipients: [{ emailAddress: { address: 'shared@example.com' } }],
          ccRecipients: null,
          receivedDateTime: '2026-07-01T10:00:00Z',
          sentDateTime: '2026-07-01T09:59:00Z',
          isRead: false,
          hasAttachments: true,
          importance: 'normal',
          bodyPreview: 'Hi there',
          conversationId: 'conv1',
        },
      ]);

      const result = await tools.listEmails({ mailbox: 'shared@example.com' });

      expect(client.listSharedMessages).toHaveBeenCalledWith('shared@example.com', undefined, 25, false);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.mailbox).toBe('shared@example.com');
      expect(parsed.emails[0]).toEqual({
        id: 'msg1',
        subject: 'Hello',
        from: 'alice@example.com',
        to: ['shared@example.com'],
        cc: [],
        receivedDateTime: '2026-07-01T10:00:00Z',
        sentDateTime: '2026-07-01T09:59:00Z',
        isRead: false,
        hasAttachments: true,
        importance: 'normal',
        preview: 'Hi there',
        conversationId: 'conv1',
      });
    });

    it('passes folder_id, limit, and unread_only through', async () => {
      vi.mocked(client.listSharedMessages).mockResolvedValue([]);

      await tools.listEmails({ mailbox: 'shared@example.com', folder_id: 'FOLDER1', limit: 5, unread_only: true });

      expect(client.listSharedMessages).toHaveBeenCalledWith('shared@example.com', 'FOLDER1', 5, true);
    });

    it('rejects a durable token passed as folder_id', async () => {
      const token = mintSelfEncoded('folder', 'FOLDER1');

      await expect(tools.listEmails({ mailbox: 'shared@example.com', folder_id: token })).rejects.toThrow(/raw Graph id/);
      expect(client.listSharedMessages).not.toHaveBeenCalled();
    });
  });

  describe('getEmail', () => {
    it('includes body when requested', async () => {
      vi.mocked(client.getSharedMessage).mockResolvedValue({
        id: 'msg1',
        subject: 'Hello',
        from: { emailAddress: { address: 'alice@example.com' } },
        body: { contentType: 'text', content: 'Full body' },
      });

      const result = await tools.getEmail({ mailbox: 'shared@example.com', email_id: 'msg1', include_body: true });

      expect(client.getSharedMessage).toHaveBeenCalledWith('shared@example.com', 'msg1', true);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.body).toBe('Full body');
      expect(parsed.id).toBe('msg1');
    });

    it('omits body when not requested', async () => {
      vi.mocked(client.getSharedMessage).mockResolvedValue({ id: 'msg1', body: { content: 'Full body' } });

      const result = await tools.getEmail({ mailbox: 'shared@example.com', email_id: 'msg1' });

      expect(client.getSharedMessage).toHaveBeenCalledWith('shared@example.com', 'msg1', false);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.body).toBeNull();
    });

    it('strips html when requested', async () => {
      vi.mocked(client.getSharedMessage).mockResolvedValue({ id: 'msg1', body: { content: '<p>Hello <b>world</b></p>' } });

      const result = await tools.getEmail({ mailbox: 'shared@example.com', email_id: 'msg1', include_body: true, strip_html: true });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.body).toBe('Hello world');
    });

    it('rejects a durable token passed as email_id', async () => {
      const token = mintSelfEncoded('message', 'msg1');

      await expect(tools.getEmail({ mailbox: 'shared@example.com', email_id: token })).rejects.toThrow(/raw Graph id/);
      expect(client.getSharedMessage).not.toHaveBeenCalled();
    });
  });

  describe('searchEmails', () => {
    it('passes query and limit', async () => {
      vi.mocked(client.searchSharedMessages).mockResolvedValue([]);

      await tools.searchEmails({ mailbox: 'shared@example.com', query: 'invoice', limit: 10 });

      expect(client.searchSharedMessages).toHaveBeenCalledWith('shared@example.com', 'invoice', 10);
    });
  });

  describe('listEvents', () => {
    it('maps events without a window', async () => {
      vi.mocked(client.listSharedEvents).mockResolvedValue([
        {
          id: 'ev1',
          subject: 'Standup',
          start: { dateTime: '2026-07-01T15:00:00', timeZone: 'UTC' },
          end: { dateTime: '2026-07-01T15:30:00', timeZone: 'UTC' },
          location: { displayName: 'Room 1' },
          isAllDay: false,
          organizer: { emailAddress: { address: 'boss@example.com' } },
          attendees: [{ emailAddress: { address: 'a@example.com' } }, { emailAddress: { address: 'b@example.com' } }],
          bodyPreview: 'Daily',
        },
      ]);

      const result = await tools.listEvents({ mailbox: 'shared@example.com' });

      expect(client.listSharedEvents).toHaveBeenCalledWith('shared@example.com', 25, undefined, undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.events[0]).toEqual({
        id: 'ev1',
        subject: 'Standup',
        start: '2026-07-01T15:00:00',
        end: '2026-07-01T15:30:00',
        location: 'Room 1',
        isAllDay: false,
        organizer: 'boss@example.com',
        attendees: ['a@example.com', 'b@example.com'],
        preview: 'Daily',
      });
    });

    it('passes a start/end window as Dates', async () => {
      vi.mocked(client.listSharedEvents).mockResolvedValue([]);

      await tools.listEvents({ mailbox: 'shared@example.com', start: '2026-07-01T00:00:00Z', end: '2026-07-02T00:00:00Z' });

      const call = vi.mocked(client.listSharedEvents).mock.calls[0];
      expect(call[0]).toBe('shared@example.com');
      expect(call[2]).toBeInstanceOf(Date);
      expect(call[3]).toBeInstanceOf(Date);
      expect((call[2] as Date).toISOString()).toBe('2026-07-01T00:00:00.000Z');
    });

    it('rejects start without end', async () => {
      await expect(tools.listEvents({ mailbox: 'shared@example.com', start: '2026-07-01T00:00:00Z' })).rejects.toThrow(/together/);
      expect(client.listSharedEvents).not.toHaveBeenCalled();
    });

    it('rejects an invalid date', async () => {
      await expect(
        tools.listEvents({ mailbox: 'shared@example.com', start: 'not-a-date', end: '2026-07-02T00:00:00Z' }),
      ).rejects.toThrow(/ISO 8601/);
    });
  });

  describe('getEvent', () => {
    it('returns the event with body', async () => {
      vi.mocked(client.getSharedEvent).mockResolvedValue({
        id: 'ev1',
        subject: 'Standup',
        body: { content: 'Agenda' },
      });

      const result = await tools.getEvent({ mailbox: 'shared@example.com', event_id: 'ev1' });

      expect(client.getSharedEvent).toHaveBeenCalledWith('shared@example.com', 'ev1');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.id).toBe('ev1');
      expect(parsed.body).toBe('Agenda');
    });

    it('rejects a durable token passed as event_id', async () => {
      const token = mintSelfEncoded('event', 'ev1');

      await expect(tools.getEvent({ mailbox: 'shared@example.com', event_id: token })).rejects.toThrow(/raw Graph id/);
      expect(client.getSharedEvent).not.toHaveBeenCalled();
    });
  });

  describe('listDriveItems', () => {
    it('maps drive items and detects folders', async () => {
      vi.mocked(client.listSharedDriveItems).mockResolvedValue([
        { id: 'd1', name: 'Report.docx', size: 2048, webUrl: 'https://x/1', lastModifiedDateTime: '2026-07-01T00:00:00Z' },
        { id: 'd2', name: 'Sub', folder: { childCount: 2 } },
      ]);

      const result = await tools.listDriveItems({ mailbox: 'shared@example.com' });

      expect(client.listSharedDriveItems).toHaveBeenCalledWith('shared@example.com', undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.items[0]).toEqual({
        id: 'd1',
        name: 'Report.docx',
        size: 2048,
        webUrl: 'https://x/1',
        lastModifiedDateTime: '2026-07-01T00:00:00Z',
        isFolder: false,
      });
      expect(parsed.items[1].isFolder).toBe(true);
    });

    it('passes item_id through', async () => {
      vi.mocked(client.listSharedDriveItems).mockResolvedValue([]);

      await tools.listDriveItems({ mailbox: 'shared@example.com', item_id: 'ITEM1' });

      expect(client.listSharedDriveItems).toHaveBeenCalledWith('shared@example.com', 'ITEM1');
    });
  });

  describe('searchDriveItems', () => {
    it('passes query and limit', async () => {
      vi.mocked(client.searchSharedDriveItems).mockResolvedValue([]);

      await tools.searchDriveItems({ mailbox: 'shared@example.com', query: 'budget', limit: 7 });

      expect(client.searchSharedDriveItems).toHaveBeenCalledWith('shared@example.com', 'budget', 7);
    });
  });
});
