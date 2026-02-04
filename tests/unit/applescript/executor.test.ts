import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('node:child_process', () => ({
  execSync: vi.fn(),
}));

import { execSync } from 'node:child_process';
import {
  escapeForAppleScript,
  executeAppleScript,
  executeAppleScriptOrThrow,
  isOutlookRunning,
  launchOutlook,
  getOutlookVersion,
  AppleScriptExecutionError,
} from '../../../src/applescript/executor.js';

const mockedExecSync = vi.mocked(execSync);

describe('escapeForAppleScript', () => {
  it('escapes backslashes', () => {
    expect(escapeForAppleScript('path\\to\\file')).toBe('path\\\\to\\\\file');
  });

  it('escapes double quotes', () => {
    expect(escapeForAppleScript('say "hello"')).toBe('say \\"hello\\"');
  });

  it('escapes both backslashes and quotes', () => {
    expect(escapeForAppleScript('a\\b"c')).toBe('a\\\\b\\"c');
  });

  it('returns unchanged string without special characters', () => {
    expect(escapeForAppleScript('hello world')).toBe('hello world');
  });
});

describe('executeAppleScript', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('returns success result when execSync succeeds', () => {
    mockedExecSync.mockReturnValue('  output text  ');
    const result = executeAppleScript('tell app "Finder" to activate');
    expect(result.success).toBe(true);
    expect(result.output).toBe('output text');
    expect(result.error).toBeUndefined();
  });

  it('returns failure result when execSync throws', () => {
    mockedExecSync.mockImplementation(() => {
      throw new Error('Command failed');
    });
    const result = executeAppleScript('bad script');
    expect(result.success).toBe(false);
    expect(result.output).toBe('');
    expect(result.error).toContain('Command failed');
  });

  it('includes stderr in error message', () => {
    const error = new Error('exec error') as Error & { stderr: string };
    error.stderr = 'stderr content';
    mockedExecSync.mockImplementation(() => {
      throw error;
    });
    const result = executeAppleScript('bad script');
    expect(result.error).toContain('stderr content');
  });

  it('uses custom timeout when provided', () => {
    mockedExecSync.mockReturnValue('ok');
    executeAppleScript('test', { timeoutMs: 5000 });
    expect(mockedExecSync).toHaveBeenCalledWith(
      expect.any(String),
      expect.objectContaining({ timeout: 5000 })
    );
  });

  it('uses default timeout when not provided', () => {
    mockedExecSync.mockReturnValue('ok');
    executeAppleScript('test');
    expect(mockedExecSync).toHaveBeenCalledWith(
      expect.any(String),
      expect.objectContaining({ timeout: 30000 })
    );
  });
});

describe('executeAppleScriptOrThrow', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('returns output on success', () => {
    mockedExecSync.mockReturnValue('result');
    const output = executeAppleScriptOrThrow('test script');
    expect(output).toBe('result');
  });

  it('throws AppleScriptExecutionError on failure', () => {
    mockedExecSync.mockImplementation(() => {
      throw new Error('application isn\'t running');
    });
    expect(() => executeAppleScriptOrThrow('test')).toThrow(AppleScriptExecutionError);
  });

  it('sets errorType to not_running for not running errors', () => {
    mockedExecSync.mockImplementation(() => {
      throw new Error('application isn\'t running');
    });
    try {
      executeAppleScriptOrThrow('test');
    } catch (e) {
      expect(e).toBeInstanceOf(AppleScriptExecutionError);
      expect((e as AppleScriptExecutionError).errorType).toBe('not_running');
    }
  });

  it('sets errorType to permission_denied for permission errors', () => {
    mockedExecSync.mockImplementation(() => {
      throw new Error('not authorized to send Apple events');
    });
    try {
      executeAppleScriptOrThrow('test');
    } catch (e) {
      expect(e).toBeInstanceOf(AppleScriptExecutionError);
      expect((e as AppleScriptExecutionError).errorType).toBe('permission_denied');
    }
  });

  it('sets errorType to timeout for timeout errors', () => {
    mockedExecSync.mockImplementation(() => {
      throw new Error('timed out');
    });
    try {
      executeAppleScriptOrThrow('test');
    } catch (e) {
      expect(e).toBeInstanceOf(AppleScriptExecutionError);
      expect((e as AppleScriptExecutionError).errorType).toBe('timeout');
    }
  });

  it('sets errorType to unknown for other errors', () => {
    mockedExecSync.mockImplementation(() => {
      throw new Error('some random error');
    });
    try {
      executeAppleScriptOrThrow('test');
    } catch (e) {
      expect(e).toBeInstanceOf(AppleScriptExecutionError);
      expect((e as AppleScriptExecutionError).errorType).toBe('unknown');
    }
  });
});

describe('isOutlookRunning', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('returns true when output is "true"', () => {
    mockedExecSync.mockReturnValue('true');
    expect(isOutlookRunning()).toBe(true);
  });

  it('returns false when output is "false"', () => {
    mockedExecSync.mockReturnValue('false');
    expect(isOutlookRunning()).toBe(false);
  });

  it('returns false when execution fails', () => {
    mockedExecSync.mockImplementation(() => {
      throw new Error('failed');
    });
    expect(isOutlookRunning()).toBe(false);
  });
});

describe('launchOutlook', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('returns true on successful launch', () => {
    mockedExecSync.mockReturnValue('launched');
    expect(launchOutlook()).toBe(true);
  });

  it('returns false when launch fails', () => {
    mockedExecSync.mockImplementation(() => {
      throw new Error('failed');
    });
    expect(launchOutlook()).toBe(false);
  });
});

describe('getOutlookVersion', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('returns version string on success', () => {
    mockedExecSync.mockReturnValue('16.78.3');
    expect(getOutlookVersion()).toBe('16.78.3');
  });

  it('returns null on failure', () => {
    mockedExecSync.mockImplementation(() => {
      throw new Error('not running');
    });
    expect(getOutlookVersion()).toBeNull();
  });
});

describe('AppleScriptExecutionError', () => {
  it('has correct name', () => {
    const error = new AppleScriptExecutionError('test', 'unknown');
    expect(error.name).toBe('AppleScriptExecutionError');
  });

  it('has correct message', () => {
    const error = new AppleScriptExecutionError('test message', 'timeout');
    expect(error.message).toBe('test message');
  });

  it('has correct errorType', () => {
    const error = new AppleScriptExecutionError('test', 'not_running');
    expect(error.errorType).toBe('not_running');
  });

  it('is an instance of Error', () => {
    const error = new AppleScriptExecutionError('test', 'unknown');
    expect(error).toBeInstanceOf(Error);
  });
});
