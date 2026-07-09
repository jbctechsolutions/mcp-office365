/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect } from 'vitest';
import {
  ErrorCode,
  OutlookMcpError,
  DatabaseNotFoundError,
  DatabaseLockedError,
  DatabaseError,
  ContentFileNotFoundError,
  ContentParseError,
  ValidationError,
  NotFoundError,
  GraphAuthRequiredError,
  GraphRateLimitedError,
  GraphPermissionDeniedError,
  GraphError,
  AttachmentNotFoundError,
  AttachmentTooLargeError,
  AttachmentSaveError,
  MailSendError,
  RecurringEventError,
  isOutlookMcpError,
  wrapError,
  toErrorEnvelope,
  isErrorEnvelope,
  ensureErrorEnvelopeText,
  ReadOnlyModeError,
} from '../../../src/utils/errors.js';

describe('errors', () => {
  describe('ErrorCode', () => {
    it('has all expected error codes', () => {
      expect(ErrorCode.DATABASE_NOT_FOUND).toBe('DATABASE_NOT_FOUND');
      expect(ErrorCode.DATABASE_LOCKED).toBe('DATABASE_LOCKED');
      expect(ErrorCode.DATABASE_ERROR).toBe('DATABASE_ERROR');
      expect(ErrorCode.CONTENT_FILE_NOT_FOUND).toBe('CONTENT_FILE_NOT_FOUND');
      expect(ErrorCode.CONTENT_PARSE_ERROR).toBe('CONTENT_PARSE_ERROR');
      expect(ErrorCode.VALIDATION_ERROR).toBe('VALIDATION_ERROR');
      expect(ErrorCode.NOT_FOUND).toBe('NOT_FOUND');
      expect(ErrorCode.AUTH_EXPIRED).toBe('AUTH_EXPIRED');
      expect(ErrorCode.GRAPH_RATE_LIMITED).toBe('GRAPH_RATE_LIMITED');
      expect(ErrorCode.GRAPH_PERMISSION_DENIED).toBe('GRAPH_PERMISSION_DENIED');
      expect(ErrorCode.GRAPH_ERROR).toBe('GRAPH_ERROR');
      expect(ErrorCode.ATTACHMENT_NOT_FOUND).toBe('ATTACHMENT_NOT_FOUND');
      expect(ErrorCode.MAIL_SEND_ERROR).toBe('MAIL_SEND_ERROR');
      expect(ErrorCode.RECURRING_EVENT_ERROR).toBe('RECURRING_EVENT_ERROR');
    });
  });

  describe('DatabaseNotFoundError', () => {
    it('creates error with correct message', () => {
      const error = new DatabaseNotFoundError('/path/to/db.sqlite');
      expect(error.message).toContain('/path/to/db.sqlite');
      expect(error.message).toContain('Outlook for Mac has been opened');
      expect(error.code).toBe(ErrorCode.DATABASE_NOT_FOUND);
      expect(error.name).toBe('DatabaseNotFoundError');
    });

    it('extends OutlookMcpError', () => {
      const error = new DatabaseNotFoundError('/path');
      expect(error).toBeInstanceOf(OutlookMcpError);
      expect(error).toBeInstanceOf(Error);
    });
  });

  describe('DatabaseLockedError', () => {
    it('creates error with correct message', () => {
      const error = new DatabaseLockedError();
      expect(error.message).toContain('locked');
      expect(error.message).toContain('try again');
      expect(error.code).toBe(ErrorCode.DATABASE_LOCKED);
      expect(error.name).toBe('DatabaseLockedError');
    });
  });

  describe('DatabaseError', () => {
    it('creates error with message', () => {
      const error = new DatabaseError('Something went wrong');
      expect(error.message).toBe('Something went wrong');
      expect(error.code).toBe(ErrorCode.DATABASE_ERROR);
      expect(error.cause).toBeUndefined();
    });

    it('captures cause', () => {
      const cause = new Error('Original error');
      const error = new DatabaseError('Wrapped error', cause);
      expect(error.cause).toBe(cause);
    });
  });

  describe('ContentFileNotFoundError', () => {
    it('creates error with path', () => {
      const error = new ContentFileNotFoundError('/path/to/file.olk15Message');
      expect(error.message).toContain('/path/to/file.olk15Message');
      expect(error.code).toBe(ErrorCode.CONTENT_FILE_NOT_FOUND);
    });
  });

  describe('ContentParseError', () => {
    it('creates error with path and cause', () => {
      const cause = new Error('Parse failed');
      const error = new ContentParseError('/path/to/file.olk15Message', cause);
      expect(error.message).toContain('/path/to/file.olk15Message');
      expect(error.code).toBe(ErrorCode.CONTENT_PARSE_ERROR);
      expect(error.cause).toBe(cause);
    });
  });

  describe('ValidationError', () => {
    it('creates error with message', () => {
      const error = new ValidationError('Invalid input');
      expect(error.message).toBe('Invalid input');
      expect(error.code).toBe(ErrorCode.VALIDATION_ERROR);
    });
  });

  describe('NotFoundError', () => {
    it('creates error for numeric ID', () => {
      const error = new NotFoundError('Email', 123);
      expect(error.message).toContain('Email');
      expect(error.message).toContain('123');
      expect(error.code).toBe(ErrorCode.NOT_FOUND);
    });

    it('creates error for string ID', () => {
      const error = new NotFoundError('Folder', 'inbox');
      expect(error.message).toContain('Folder');
      expect(error.message).toContain('inbox');
    });
  });

  describe('isOutlookMcpError', () => {
    it('returns true for OutlookMcpError instances', () => {
      expect(isOutlookMcpError(new DatabaseNotFoundError('/path'))).toBe(true);
      expect(isOutlookMcpError(new DatabaseLockedError())).toBe(true);
      expect(isOutlookMcpError(new DatabaseError('msg'))).toBe(true);
      expect(isOutlookMcpError(new ContentFileNotFoundError('/path'))).toBe(true);
      expect(isOutlookMcpError(new ContentParseError('/path'))).toBe(true);
      expect(isOutlookMcpError(new ValidationError('msg'))).toBe(true);
      expect(isOutlookMcpError(new NotFoundError('type', 1))).toBe(true);
    });

    it('returns false for regular Error', () => {
      expect(isOutlookMcpError(new Error('test'))).toBe(false);
    });

    it('returns false for non-error values', () => {
      expect(isOutlookMcpError('string')).toBe(false);
      expect(isOutlookMcpError(null)).toBe(false);
      expect(isOutlookMcpError(undefined)).toBe(false);
      expect(isOutlookMcpError({})).toBe(false);
    });
  });

  describe('wrapError', () => {
    it('returns OutlookMcpError as-is', () => {
      const original = new DatabaseLockedError();
      const wrapped = wrapError(original, 'default');
      expect(wrapped).toBe(original);
    });

    it('wraps regular Error in DatabaseError', () => {
      const original = new Error('Original message');
      const wrapped = wrapError(original, 'default');
      expect(wrapped).toBeInstanceOf(DatabaseError);
      expect(wrapped.message).toBe('Original message');
      expect((wrapped as DatabaseError).cause).toBe(original);
    });

    it('wraps non-Error values with default message', () => {
      const wrapped = wrapError('string error', 'Default message');
      expect(wrapped).toBeInstanceOf(DatabaseError);
      expect(wrapped.message).toBe('Default message');
    });

    it('wraps null/undefined with default message', () => {
      expect(wrapError(null, 'Default').message).toBe('Default');
      expect(wrapError(undefined, 'Default').message).toBe('Default');
    });
  });

  // =========================================================================
  // Microsoft Graph API Errors
  // =========================================================================

  describe('GraphAuthRequiredError', () => {
    it('creates error with correct message and the AUTH_EXPIRED code (U9 consolidation)', () => {
      const error = new GraphAuthRequiredError();
      expect(error.message).toContain('authentication required');
      expect(error.message).toContain('npx @jbctechsolutions/mcp-office365 auth');
      expect(error.code).toBe(ErrorCode.AUTH_EXPIRED);
      expect(error.name).toBe('GraphAuthRequiredError');
    });

    it('reports a session-expired reason distinctly', () => {
      const error = new GraphAuthRequiredError('session_expired');
      expect(error.code).toBe(ErrorCode.AUTH_EXPIRED);
      expect(error.message).toContain('session expired');
    });

    it('extends OutlookMcpError', () => {
      const error = new GraphAuthRequiredError();
      expect(error).toBeInstanceOf(OutlookMcpError);
      expect(error).toBeInstanceOf(Error);
    });
  });

  describe('GraphRateLimitedError', () => {
    it('creates error with retryAfter', () => {
      const error = new GraphRateLimitedError(30);
      expect(error.message).toContain('rate limit exceeded');
      expect(error.message).toContain('30 seconds');
      expect(error.code).toBe(ErrorCode.GRAPH_RATE_LIMITED);
      expect(error.name).toBe('GraphRateLimitedError');
      expect(error.retryAfter).toBe(30);
    });

    it('creates error without retryAfter', () => {
      const error = new GraphRateLimitedError();
      expect(error.message).toContain('rate limit exceeded');
      expect(error.message).toContain('try again later');
      expect(error.retryAfter).toBeUndefined();
    });
  });

  describe('GraphPermissionDeniedError', () => {
    it('creates error with scope', () => {
      const error = new GraphPermissionDeniedError('Mail.Read');
      expect(error.message).toContain('permission denied');
      expect(error.message).toContain("'Mail.Read'");
      expect(error.code).toBe(ErrorCode.GRAPH_PERMISSION_DENIED);
      expect(error.name).toBe('GraphPermissionDeniedError');
    });

    it('creates error without scope', () => {
      const error = new GraphPermissionDeniedError();
      expect(error.message).toContain('permission denied');
      expect(error.message).toContain('Azure AD');
    });
  });

  describe('GraphError', () => {
    it('creates error with message', () => {
      const error = new GraphError('API call failed');
      expect(error.message).toBe('API call failed');
      expect(error.code).toBe(ErrorCode.GRAPH_ERROR);
      expect(error.cause).toBeUndefined();
    });

    it('captures cause', () => {
      const cause = new Error('Original error');
      const error = new GraphError('Wrapped error', cause);
      expect(error.cause).toBe(cause);
    });
  });

  describe('isOutlookMcpError with Graph errors', () => {
    it('returns true for Graph error instances', () => {
      expect(isOutlookMcpError(new GraphAuthRequiredError())).toBe(true);
      expect(isOutlookMcpError(new GraphRateLimitedError(30))).toBe(true);
      expect(isOutlookMcpError(new GraphPermissionDeniedError('scope'))).toBe(true);
      expect(isOutlookMcpError(new GraphError('msg'))).toBe(true);
    });
  });

  // =========================================================================
  // Event Management and Email Errors
  // =========================================================================

  describe('AttachmentNotFoundError', () => {
    it('creates error with file path', () => {
      const error = new AttachmentNotFoundError('/path/to/file.pdf');
      expect(error.code).toBe(ErrorCode.ATTACHMENT_NOT_FOUND);
      expect(error.message).toContain('/path/to/file.pdf');
      expect(error.message).toContain('not found');
      expect(error.name).toBe('AttachmentNotFoundError');
    });

    it('extends OutlookMcpError', () => {
      const error = new AttachmentNotFoundError('/path/to/file.pdf');
      expect(error).toBeInstanceOf(OutlookMcpError);
      expect(error).toBeInstanceOf(Error);
    });
  });

  describe('MailSendError', () => {
    it('creates error with reason', () => {
      const error = new MailSendError('Network timeout');
      expect(error.code).toBe(ErrorCode.MAIL_SEND_ERROR);
      expect(error.message).toContain('Failed to send email');
      expect(error.message).toContain('Network timeout');
      expect(error.name).toBe('MailSendError');
    });

    it('extends OutlookMcpError', () => {
      const error = new MailSendError('Test reason');
      expect(error).toBeInstanceOf(OutlookMcpError);
      expect(error).toBeInstanceOf(Error);
    });
  });

  describe('RecurringEventError', () => {
    it('creates error with custom message', () => {
      const error = new RecurringEventError('Invalid recurrence pattern');
      expect(error.code).toBe(ErrorCode.RECURRING_EVENT_ERROR);
      expect(error.message).toBe('Invalid recurrence pattern');
      expect(error.name).toBe('RecurringEventError');
    });

    it('extends OutlookMcpError', () => {
      const error = new RecurringEventError('Test message');
      expect(error).toBeInstanceOf(OutlookMcpError);
      expect(error).toBeInstanceOf(Error);
    });
  });

  // =========================================================================
  // Attachment Errors
  // =========================================================================

  describe('AttachmentTooLargeError', () => {
    it('has code ATTACHMENT_TOO_LARGE and message includes size info', () => {
      const error = new AttachmentTooLargeError('large-file.zip', 52_428_800, 25_165_824);
      expect(error.code).toBe(ErrorCode.ATTACHMENT_TOO_LARGE);
      expect(error.message).toContain('large-file.zip');
      expect(error.message).toContain('50MB');
      expect(error.message).toContain('24MB');
      expect(error.name).toBe('AttachmentTooLargeError');
    });

    it('extends OutlookMcpError', () => {
      const error = new AttachmentTooLargeError('file.zip', 1000000, 500000);
      expect(error).toBeInstanceOf(OutlookMcpError);
      expect(error).toBeInstanceOf(Error);
    });
  });

  describe('AttachmentSaveError', () => {
    it('has code ATTACHMENT_SAVE_ERROR and message includes name and reason', () => {
      const error = new AttachmentSaveError('report.pdf', 'Disk full');
      expect(error.code).toBe(ErrorCode.ATTACHMENT_SAVE_ERROR);
      expect(error.message).toContain('report.pdf');
      expect(error.message).toContain('Disk full');
      expect(error.name).toBe('AttachmentSaveError');
    });

    it('extends OutlookMcpError', () => {
      const error = new AttachmentSaveError('file.txt', 'Permission denied');
      expect(error).toBeInstanceOf(OutlookMcpError);
      expect(error).toBeInstanceOf(Error);
    });
  });

  describe('ReadOnlyModeError', () => {
    it('has code READ_ONLY_MODE, is non-retriable, names the tool, and maps to an envelope', () => {
      const error = new ReadOnlyModeError('confirm_delete_email');
      expect(error.code).toBe(ErrorCode.READ_ONLY_MODE);
      expect(error.retriable).toBe(false);
      expect(error.message).toContain('confirm_delete_email');
      expect(error).toBeInstanceOf(OutlookMcpError);

      const env = toErrorEnvelope(error);
      expect(env.code).toBe(ErrorCode.READ_ONLY_MODE);
      expect(env.retriable).toBe(false);
      expect(env.suggestion).toContain('--read-only');
    });
  });

  describe('toErrorEnvelope (D10)', () => {
    it('maps a typed OutlookMcpError to its code/message', () => {
      const env = toErrorEnvelope(new ValidationError('bad input'));
      expect(env).toEqual({
        code: ErrorCode.VALIDATION_ERROR,
        message: 'bad input',
        retriable: false,
      });
    });

    it('carries retriable + suggestion from GraphRateLimitedError', () => {
      const env = toErrorEnvelope(new GraphRateLimitedError(2));
      expect(env.code).toBe(ErrorCode.GRAPH_RATE_LIMITED);
      expect(env.retriable).toBe(true);
      expect(env.suggestion).toContain('2s');
    });

    it('carries the auth suggestion from GraphAuthRequiredError', () => {
      const env = toErrorEnvelope(new GraphAuthRequiredError());
      expect(env.code).toBe(ErrorCode.AUTH_EXPIRED);
      expect(env.retriable).toBe(false);
      expect(env.suggestion).toContain('auth');
    });

    it('maps a 401 Graph SDK error to AUTH_EXPIRED (not retriable)', () => {
      const env = toErrorEnvelope({ statusCode: 401, message: 'Access token expired' });
      expect(env.code).toBe(ErrorCode.AUTH_EXPIRED);
      expect(env.retriable).toBe(false);
      expect(env.suggestion).toContain('auth');
    });

    it('maps a 429 Graph SDK error to THROTTLED (retriable)', () => {
      const env = toErrorEnvelope({ statusCode: 429, message: 'Too many requests' });
      expect(env.code).toBe(ErrorCode.THROTTLED);
      expect(env.retriable).toBe(true);
    });

    it('maps 502/503/504 to GRAPH_UNAVAILABLE (retriable)', () => {
      for (const status of [502, 503, 504]) {
        const env = toErrorEnvelope({ statusCode: status, message: 'upstream' });
        expect(env.code).toBe(ErrorCode.GRAPH_UNAVAILABLE);
        expect(env.retriable).toBe(true);
      }
    });

    it('maps a 403 to permission denied and 404 to not found', () => {
      expect(toErrorEnvelope({ statusCode: 403, message: 'no' }).code).toBe(
        ErrorCode.GRAPH_PERMISSION_DENIED
      );
      expect(toErrorEnvelope({ statusCode: 404, message: 'gone' }).code).toBe(ErrorCode.NOT_FOUND);
    });

    it('maps an unrecognized 5xx to GRAPH_UNAVAILABLE and a 4xx to GRAPH_ERROR', () => {
      expect(toErrorEnvelope({ statusCode: 500, message: 'x' }).code).toBe(
        ErrorCode.GRAPH_UNAVAILABLE
      );
      expect(toErrorEnvelope({ statusCode: 418, message: 'teapot' }).code).toBe(
        ErrorCode.GRAPH_ERROR
      );
    });

    it('marks non-auto-retried 5xx (500/501) as NOT retriable, matching the D5 policy', () => {
      // Only 429/502/503/504 are auto-retried by the transport; the envelope's
      // retriable flag must not overstate coverage for bare 500/501.
      expect(toErrorEnvelope({ statusCode: 500, message: 'x' }).retriable).toBe(false);
      expect(toErrorEnvelope({ statusCode: 501, message: 'x' }).retriable).toBe(false);
      // ...but the auto-retried statuses stay retriable:true.
      expect(toErrorEnvelope({ statusCode: 503, message: 'x' }).retriable).toBe(true);
    });

    it('synthesizes a message when the SDK error has none', () => {
      const env = toErrorEnvelope({ statusCode: 503 });
      expect(env.message).toContain('503');
    });

    it('maps a plain Error to a non-retriable GRAPH_ERROR', () => {
      const env = toErrorEnvelope(new Error('boom'));
      expect(env).toEqual({ code: ErrorCode.GRAPH_ERROR, message: 'boom', retriable: false });
    });

    it('maps an unknown non-error value to a generic GRAPH_ERROR', () => {
      const env = toErrorEnvelope('just a string');
      expect(env.code).toBe(ErrorCode.GRAPH_ERROR);
      expect(env.retriable).toBe(false);
    });
  });

  describe('isErrorEnvelope', () => {
    it('recognizes a well-formed envelope', () => {
      expect(isErrorEnvelope({ code: 'GRAPH_ERROR', message: 'x', retriable: false })).toBe(true);
    });

    it('rejects shapes missing a required field or with wrong types', () => {
      expect(isErrorEnvelope({ code: 'GRAPH_ERROR', message: 'x' })).toBe(false); // no retriable
      expect(isErrorEnvelope({ code: 1, message: 'x', retriable: false })).toBe(false); // code not string
      expect(isErrorEnvelope(null)).toBe(false);
      expect(isErrorEnvelope('str')).toBe(false);
    });

    it('rejects an unknown code even when the rest of the shape matches', () => {
      // A tool payload that happens to carry code/message/retriable fields must
      // not be mistaken for an envelope.
      expect(isErrorEnvelope({ code: 'X', message: 'x', retriable: false })).toBe(false);
    });

    it('rejects a non-string suggestion', () => {
      expect(
        isErrorEnvelope({ code: 'GRAPH_ERROR', message: 'x', retriable: false, suggestion: 5 })
      ).toBe(false);
    });
  });

  describe('ensureErrorEnvelopeText (D10 handler-return normalization)', () => {
    it('wraps a plain error message into an envelope JSON string', () => {
      const out = ensureErrorEnvelopeText('Email not found');
      const parsed = JSON.parse(out) as unknown;
      expect(isErrorEnvelope(parsed)).toBe(true);
      expect((parsed as { code: string }).code).toBe(ErrorCode.GRAPH_ERROR);
      expect((parsed as { message: string }).message).toBe('Email not found');
      expect((parsed as { retriable: boolean }).retriable).toBe(false);
    });

    it('is idempotent — text that is already an envelope is returned unchanged', () => {
      const envelope = JSON.stringify(
        { code: 'NOT_FOUND', message: 'gone', retriable: false },
        null,
        2
      );
      expect(ensureErrorEnvelopeText(envelope)).toBe(envelope);
    });

    it('wraps JSON that is not an envelope (e.g. a tool payload) rather than passing it through', () => {
      const out = ensureErrorEnvelopeText('{"foo":1}');
      const parsed = JSON.parse(out) as { code: string; message: string };
      expect(parsed.code).toBe(ErrorCode.GRAPH_ERROR);
      expect(parsed.message).toBe('{"foo":1}');
    });

    it('wraps envelope-shaped JSON whose code is not a known ErrorCode', () => {
      const payload = '{"code":"X","message":"x","retriable":false}';
      const out = ensureErrorEnvelopeText(payload);
      const parsed = JSON.parse(out) as { code: string; message: string };
      expect(parsed.code).toBe(ErrorCode.GRAPH_ERROR);
      expect(parsed.message).toBe(payload);
    });
  });
});
