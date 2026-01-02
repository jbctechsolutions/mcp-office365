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
  OutlookNotRunningError,
  AppleScriptPermissionError,
  AppleScriptTimeoutError,
  AppleScriptError,
  GraphAuthRequiredError,
  GraphRateLimitedError,
  GraphPermissionDeniedError,
  GraphError,
  isOutlookMcpError,
  wrapError,
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
      expect(ErrorCode.OUTLOOK_NOT_RUNNING).toBe('OUTLOOK_NOT_RUNNING');
      expect(ErrorCode.APPLESCRIPT_PERMISSION_DENIED).toBe('APPLESCRIPT_PERMISSION_DENIED');
      expect(ErrorCode.APPLESCRIPT_TIMEOUT).toBe('APPLESCRIPT_TIMEOUT');
      expect(ErrorCode.APPLESCRIPT_ERROR).toBe('APPLESCRIPT_ERROR');
      expect(ErrorCode.GRAPH_AUTH_REQUIRED).toBe('GRAPH_AUTH_REQUIRED');
      expect(ErrorCode.GRAPH_RATE_LIMITED).toBe('GRAPH_RATE_LIMITED');
      expect(ErrorCode.GRAPH_PERMISSION_DENIED).toBe('GRAPH_PERMISSION_DENIED');
      expect(ErrorCode.GRAPH_ERROR).toBe('GRAPH_ERROR');
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
  // AppleScript Errors
  // =========================================================================

  describe('OutlookNotRunningError', () => {
    it('creates error with correct message', () => {
      const error = new OutlookNotRunningError();
      expect(error.message).toContain('not running');
      expect(error.message).toContain('start Outlook');
      expect(error.code).toBe(ErrorCode.OUTLOOK_NOT_RUNNING);
      expect(error.name).toBe('OutlookNotRunningError');
    });

    it('extends OutlookMcpError', () => {
      const error = new OutlookNotRunningError();
      expect(error).toBeInstanceOf(OutlookMcpError);
      expect(error).toBeInstanceOf(Error);
    });
  });

  describe('AppleScriptPermissionError', () => {
    it('creates error with correct message', () => {
      const error = new AppleScriptPermissionError();
      expect(error.message).toContain('Automation permission denied');
      expect(error.message).toContain('System Settings');
      expect(error.code).toBe(ErrorCode.APPLESCRIPT_PERMISSION_DENIED);
      expect(error.name).toBe('AppleScriptPermissionError');
    });
  });

  describe('AppleScriptTimeoutError', () => {
    it('creates error with operation name', () => {
      const error = new AppleScriptTimeoutError('listEmails');
      expect(error.message).toContain('timed out');
      expect(error.message).toContain('listEmails');
      expect(error.code).toBe(ErrorCode.APPLESCRIPT_TIMEOUT);
      expect(error.name).toBe('AppleScriptTimeoutError');
    });
  });

  describe('AppleScriptError', () => {
    it('creates error with message', () => {
      const error = new AppleScriptError('Script failed');
      expect(error.message).toBe('Script failed');
      expect(error.code).toBe(ErrorCode.APPLESCRIPT_ERROR);
      expect(error.cause).toBeUndefined();
    });

    it('captures cause', () => {
      const cause = new Error('Original error');
      const error = new AppleScriptError('Wrapped error', cause);
      expect(error.cause).toBe(cause);
    });
  });

  // =========================================================================
  // Microsoft Graph API Errors
  // =========================================================================

  describe('GraphAuthRequiredError', () => {
    it('creates error with correct message', () => {
      const error = new GraphAuthRequiredError();
      expect(error.message).toContain('authentication required');
      expect(error.message).toContain('device code flow');
      expect(error.code).toBe(ErrorCode.GRAPH_AUTH_REQUIRED);
      expect(error.name).toBe('GraphAuthRequiredError');
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

    it('returns true for AppleScript error instances', () => {
      expect(isOutlookMcpError(new OutlookNotRunningError())).toBe(true);
      expect(isOutlookMcpError(new AppleScriptPermissionError())).toBe(true);
      expect(isOutlookMcpError(new AppleScriptTimeoutError('op'))).toBe(true);
      expect(isOutlookMcpError(new AppleScriptError('msg'))).toBe(true);
    });
  });
});
