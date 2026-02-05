/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Graph email mapper functions.
 */

import { describe, it, expect } from 'vitest';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { mapMessageToEmailRow } from '../../../../src/graph/mappers/email-mapper.js';
import { hashStringToNumber } from '../../../../src/graph/mappers/utils.js';

describe('graph/mappers/email-mapper', () => {
  describe('mapMessageToEmailRow', () => {
    it('maps message with all fields', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        subject: 'Test Email',
        from: {
          emailAddress: { address: 'sender@example.com', name: 'Sender Name' },
        },
        toRecipients: [
          { emailAddress: { address: 'to@example.com', name: 'To User' } },
        ],
        ccRecipients: [
          { emailAddress: { address: 'cc@example.com', name: 'CC User' } },
        ],
        receivedDateTime: '2024-01-15T10:30:00Z',
        sentDateTime: '2024-01-15T10:29:00Z',
        isRead: true,
        hasAttachments: true,
        importance: 'high',
        flag: { flagStatus: 'flagged' },
        bodyPreview: 'This is a preview...',
        conversationId: 'conv-456',
        internetMessageId: '<msg123@example.com>',
        parentFolderId: 'folder-789',
      };

      const result = mapMessageToEmailRow(message);

      expect(result.id).toBe(hashStringToNumber('msg-123'));
      expect(result.subject).toBe('Test Email');
      expect(result.sender).toBe('Sender Name');
      expect(result.senderAddress).toBe('sender@example.com');
      expect(result.recipients).toBe('To User');
      expect(result.displayTo).toBe('To User');
      expect(result.toAddresses).toBe('to@example.com');
      expect(result.ccAddresses).toBe('cc@example.com');
      expect(result.preview).toBe('This is a preview...');
      expect(result.isRead).toBe(1);
      expect(result.hasAttachment).toBe(1);
      expect(result.priority).toBe(1); // high
      expect(result.flagStatus).toBe(1); // flagged
      expect(result.messageId).toBe('<msg123@example.com>');
      expect(result.conversationId).toBe(hashStringToNumber('conv-456'));
      expect(result.dataFilePath).toBe('graph-email:msg-123');
    });

    it('uses provided folderId over parentFolderId', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        parentFolderId: 'folder-parent',
      };

      const result = mapMessageToEmailRow(message, 'folder-override');

      expect(result.folderId).toBe(hashStringToNumber('folder-override'));
    });

    it('uses parentFolderId when folderId not provided', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        parentFolderId: 'folder-parent',
      };

      const result = mapMessageToEmailRow(message);

      expect(result.folderId).toBe(hashStringToNumber('folder-parent'));
    });

    it('handles message with null id', () => {
      const message: MicrosoftGraph.Message = {
        id: undefined,
        subject: 'Test',
      };

      const result = mapMessageToEmailRow(message);

      expect(result.id).toBe(hashStringToNumber(''));
      expect(result.dataFilePath).toBe('graph-email:');
    });

    it('handles message with null subject', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        subject: undefined,
      };

      const result = mapMessageToEmailRow(message);

      expect(result.subject).toBeNull();
    });

    it('handles message without sender', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        from: undefined,
      };

      const result = mapMessageToEmailRow(message);

      expect(result.sender).toBeNull();
      expect(result.senderAddress).toBeNull();
    });

    it('handles message without recipients', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        toRecipients: undefined,
        ccRecipients: undefined,
      };

      const result = mapMessageToEmailRow(message);

      expect(result.recipients).toBeNull();
      expect(result.displayTo).toBeNull();
      expect(result.toAddresses).toBeNull();
      expect(result.ccAddresses).toBeNull();
    });

    it('handles message with empty recipients arrays', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        toRecipients: [],
        ccRecipients: [],
      };

      const result = mapMessageToEmailRow(message);

      expect(result.recipients).toBeNull();
      expect(result.toAddresses).toBeNull();
      expect(result.ccAddresses).toBeNull();
    });

    it('handles unread message', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        isRead: false,
      };

      const result = mapMessageToEmailRow(message);

      expect(result.isRead).toBe(0);
    });

    it('handles message without attachments', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        hasAttachments: false,
      };

      const result = mapMessageToEmailRow(message);

      expect(result.hasAttachment).toBe(0);
    });

    it('handles low importance', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        importance: 'low',
      };

      const result = mapMessageToEmailRow(message);

      expect(result.priority).toBe(-1);
    });

    it('handles normal importance', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        importance: 'normal',
      };

      const result = mapMessageToEmailRow(message);

      expect(result.priority).toBe(0);
    });

    it('handles complete flag status', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        flag: { flagStatus: 'complete' },
      };

      const result = mapMessageToEmailRow(message);

      expect(result.flagStatus).toBe(2);
    });

    it('handles notFlagged status', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        flag: { flagStatus: 'notFlagged' },
      };

      const result = mapMessageToEmailRow(message);

      expect(result.flagStatus).toBe(0);
    });

    it('handles message without flag', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        flag: undefined,
      };

      const result = mapMessageToEmailRow(message);

      expect(result.flagStatus).toBe(0);
    });

    it('handles message without conversationId', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        conversationId: undefined,
      };

      const result = mapMessageToEmailRow(message);

      expect(result.conversationId).toBeNull();
    });

    it('handles message without internetMessageId', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        internetMessageId: undefined,
      };

      const result = mapMessageToEmailRow(message);

      expect(result.messageId).toBeNull();
    });

    it('handles message without dates', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        receivedDateTime: undefined,
        sentDateTime: undefined,
      };

      const result = mapMessageToEmailRow(message);

      expect(result.timeReceived).toBeNull();
      expect(result.timeSent).toBeNull();
    });

    it('sets size to 0 (not available in Graph)', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
      };

      const result = mapMessageToEmailRow(message);

      expect(result.size).toBe(0);
    });

    it('maps categories to buffer', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        categories: ['Work', 'Important'],
      };

      const result = mapMessageToEmailRow(message);

      expect(result.categories).toBeInstanceOf(Buffer);
      expect(result.categories!.toString('utf-8')).toBe('Work,Important');
    });

    it('sets categories to null when empty', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        categories: [],
      };

      const result = mapMessageToEmailRow(message);

      expect(result.categories).toBeNull();
    });

    it('sets categories to null when undefined', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
      };

      const result = mapMessageToEmailRow(message);

      expect(result.categories).toBeNull();
    });

    it('handles folderId 0 when no folder info', () => {
      const message: MicrosoftGraph.Message = {
        id: 'msg-123',
        parentFolderId: undefined,
      };

      const result = mapMessageToEmailRow(message);

      expect(result.folderId).toBe(0);
    });
  });
});
