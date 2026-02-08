/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { existsSync } from 'fs';

vi.mock('../../../src/applescript/executor.js', () => ({
  executeAppleScriptOrThrow: vi.fn(),
  escapeForAppleScript: (s: string) => s.replace(/\\/g, '\\\\').replace(/"/g, '\\"'),
}));

vi.mock('fs', () => ({
  existsSync: vi.fn(),
}));

import { AppleScriptMailSender } from '../../../src/applescript/mail-sender.js';
import { executeAppleScriptOrThrow } from '../../../src/applescript/executor.js';
import { AttachmentNotFoundError, MailSendError } from '../../../src/utils/errors.js';

const mockedExecute = vi.mocked(executeAppleScriptOrThrow);
const mockedExistsSync = vi.mocked(existsSync);

describe('AppleScriptMailSender', () => {
  let sender: AppleScriptMailSender;

  beforeEach(() => {
    vi.clearAllMocks();
    sender = new AppleScriptMailSender();
    // Default to files existing
    mockedExistsSync.mockReturnValue(true);
  });

  describe('sendEmail', () => {
    it('sends plain text email to single recipient', () => {
      mockedExecute.mockReturnValue(
        '{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}12345{{FIELD}}sentAt{{=}}2024-01-15T10:30:00Z'
      );

      const result = sender.sendEmail({
        to: ['test@example.com'],
        subject: 'Test Subject',
        body: 'Test body',
        bodyType: 'plain',
      });

      expect(result).toEqual({
        messageId: '12345',
        sentAt: '2024-01-15T10:30:00Z',
      });
      expect(mockedExecute).toHaveBeenCalledOnce();
      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('Test Subject');
      expect(script).toContain('Test body');
      expect(script).toContain('plain text content');
      expect(script).toContain('test@example.com');
    });

    it('sends HTML email', () => {
      mockedExecute.mockReturnValue(
        '{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}67890{{FIELD}}sentAt{{=}}2024-01-15T10:35:00Z'
      );

      const result = sender.sendEmail({
        to: ['recipient@example.com'],
        subject: 'HTML Email',
        body: '<p>HTML body</p>',
        bodyType: 'html',
      });

      expect(result.messageId).toBe('67890');
      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('html content');
      expect(script).toContain('HTML body');
    });

    it('includes CC and BCC recipients', () => {
      mockedExecute.mockReturnValue(
        '{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}11111{{FIELD}}sentAt{{=}}2024-01-15T10:40:00Z'
      );

      sender.sendEmail({
        to: ['to@example.com'],
        subject: 'Test',
        body: 'Body',
        bodyType: 'plain',
        cc: ['cc1@example.com', 'cc2@example.com'],
        bcc: ['bcc@example.com'],
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('cc1@example.com');
      expect(script).toContain('cc2@example.com');
      expect(script).toContain('bcc@example.com');
      expect(script).toContain('recipient cc');
      expect(script).toContain('recipient bcc');
    });

    it('includes reply-to address', () => {
      mockedExecute.mockReturnValue(
        '{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}22222{{FIELD}}sentAt{{=}}2024-01-15T10:45:00Z'
      );

      sender.sendEmail({
        to: ['test@example.com'],
        subject: 'Test',
        body: 'Body',
        bodyType: 'plain',
        replyTo: 'reply@example.com',
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('reply@example.com');
    });

    it('validates attachments exist before sending', () => {
      mockedExistsSync.mockReturnValue(false);

      expect(() => {
        sender.sendEmail({
          to: ['test@example.com'],
          subject: 'Test',
          body: 'Body',
          bodyType: 'plain',
          attachments: [{ path: '/nonexistent/file.pdf' }],
        });
      }).toThrow(AttachmentNotFoundError);

      expect(mockedExecute).not.toHaveBeenCalled();
      expect(mockedExistsSync).toHaveBeenCalledWith('/nonexistent/file.pdf');
    });

    it('includes attachments when they exist', () => {
      mockedExistsSync.mockReturnValue(true);
      mockedExecute.mockReturnValue(
        '{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}33333{{FIELD}}sentAt{{=}}2024-01-15T10:50:00Z'
      );

      sender.sendEmail({
        to: ['test@example.com'],
        subject: 'Test',
        body: 'Body',
        bodyType: 'plain',
        attachments: [
          { path: '/path/to/file.pdf' },
          { path: '/path/to/image.png', name: 'screenshot.png' },
        ],
      });

      expect(mockedExistsSync).toHaveBeenCalledWith('/path/to/file.pdf');
      expect(mockedExistsSync).toHaveBeenCalledWith('/path/to/image.png');

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('POSIX file "/path/to/file.pdf"');
      expect(script).toContain('POSIX file "/path/to/image.png"');
    });

    it('includes account ID', () => {
      mockedExecute.mockReturnValue(
        '{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}44444{{FIELD}}sentAt{{=}}2024-01-15T10:55:00Z'
      );

      sender.sendEmail({
        to: ['test@example.com'],
        subject: 'Test',
        body: 'Body',
        bodyType: 'plain',
        accountId: 123,
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('account id 123');
    });

    it('throws MailSendError when send fails', () => {
      mockedExecute.mockReturnValue(
        '{{RECORD}}success{{=}}false{{FIELD}}error{{=}}Recipient not found'
      );

      expect(() => {
        sender.sendEmail({
          to: ['invalid@example.com'],
          subject: 'Test',
          body: 'Body',
          bodyType: 'plain',
        });
      }).toThrow(MailSendError);
    });

    it('throws error when parser returns null', () => {
      mockedExecute.mockReturnValue('invalid output');

      expect(() => {
        sender.sendEmail({
          to: ['test@example.com'],
          subject: 'Test',
          body: 'Body',
          bodyType: 'plain',
        });
      }).toThrow('Failed to parse send email response');
    });

    it('handles multiple recipients', () => {
      mockedExecute.mockReturnValue(
        '{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}55555{{FIELD}}sentAt{{=}}2024-01-15T11:00:00Z'
      );

      sender.sendEmail({
        to: ['user1@example.com', 'user2@example.com', 'user3@example.com'],
        subject: 'Test',
        body: 'Body',
        bodyType: 'plain',
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('user1@example.com');
      expect(script).toContain('user2@example.com');
      expect(script).toContain('user3@example.com');
    });

    // =========================================================================
    // Inline Images
    // =========================================================================

    it('validates inline image files exist before sending', () => {
      mockedExistsSync.mockReturnValue(true);
      mockedExecute.mockReturnValue(
        '{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}66666{{FIELD}}sentAt{{=}}2024-01-15T11:05:00Z'
      );

      sender.sendEmail({
        to: ['test@example.com'],
        subject: 'Test',
        body: '<p>Hello</p>',
        bodyType: 'html',
        inlineImages: [
          { path: '/path/to/logo.png', contentId: 'logo123' },
        ],
      });

      expect(mockedExistsSync).toHaveBeenCalledWith('/path/to/logo.png');
    });

    it('throws AttachmentNotFoundError for missing inline image files', () => {
      mockedExistsSync.mockReturnValue(false);

      expect(() => {
        sender.sendEmail({
          to: ['test@example.com'],
          subject: 'Test',
          body: '<p>Hello</p>',
          bodyType: 'html',
          inlineImages: [
            { path: '/nonexistent/logo.png', contentId: 'logo123' },
          ],
        });
      }).toThrow(AttachmentNotFoundError);

      expect(mockedExecute).not.toHaveBeenCalled();
      expect(mockedExistsSync).toHaveBeenCalledWith('/nonexistent/logo.png');
    });

    it('includes content id in generated script for inline images', () => {
      mockedExistsSync.mockReturnValue(true);
      mockedExecute.mockReturnValue(
        '{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}77777{{FIELD}}sentAt{{=}}2024-01-15T11:10:00Z'
      );

      sender.sendEmail({
        to: ['test@example.com'],
        subject: 'Test',
        body: '<p><img src="cid:logo123"></p>',
        bodyType: 'html',
        inlineImages: [
          { path: '/path/to/logo.png', contentId: 'logo123' },
        ],
      });

      const script = mockedExecute.mock.calls[0]![0];
      expect(script).toContain('content id');
      expect(script).toContain('logo123');
      expect(script).toContain('POSIX file "/path/to/logo.png"');
    });
  });
});
