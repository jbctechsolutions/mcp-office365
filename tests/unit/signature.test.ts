/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { join } from 'node:path';

// Use vi.hoisted so mock functions are available when vi.mock factories run
const {
  mockReadFileSync,
  mockWriteFileSync,
  mockExistsSync,
  mockMkdirSync,
  mockHomedir,
} = vi.hoisted(() => ({
  mockReadFileSync: vi.fn(),
  mockWriteFileSync: vi.fn(),
  mockExistsSync: vi.fn(),
  mockMkdirSync: vi.fn(),
  mockHomedir: vi.fn().mockReturnValue('/mock/home'),
}));

vi.mock('node:fs', () => ({
  readFileSync: mockReadFileSync,
  writeFileSync: mockWriteFileSync,
  existsSync: mockExistsSync,
  mkdirSync: mockMkdirSync,
}));

vi.mock('node:os', () => ({
  homedir: mockHomedir,
}));

import { readSignature, writeSignature, appendSignature } from '../../src/signature.js';

describe('signature', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe('readSignature', () => {
    it('returns signature content when file exists', () => {
      mockExistsSync.mockReturnValue(true);
      mockReadFileSync.mockReturnValue('<p>-- Joel</p>');

      const result = readSignature();

      expect(result).toBe('<p>-- Joel</p>');
      expect(mockExistsSync).toHaveBeenCalledWith(
        join('/mock/home', '.outlook-mcp', 'signature.html')
      );
      expect(mockReadFileSync).toHaveBeenCalledWith(
        join('/mock/home', '.outlook-mcp', 'signature.html'),
        'utf-8'
      );
    });

    it('returns null when file does not exist', () => {
      mockExistsSync.mockReturnValue(false);

      const result = readSignature();

      expect(result).toBeNull();
      expect(mockReadFileSync).not.toHaveBeenCalled();
    });
  });

  describe('writeSignature', () => {
    it('writes HTML content to signature file', () => {
      mockExistsSync.mockReturnValue(true);

      writeSignature('<p>Best regards,<br>Joel</p>');

      expect(mockWriteFileSync).toHaveBeenCalledWith(
        join('/mock/home', '.outlook-mcp', 'signature.html'),
        '<p>Best regards,<br>Joel</p>',
        { encoding: 'utf-8', mode: 0o600 }
      );
    });

    it('creates directory if it does not exist', () => {
      mockExistsSync.mockReturnValue(false);

      writeSignature('<p>Sig</p>');

      expect(mockMkdirSync).toHaveBeenCalledWith(
        join('/mock/home', '.outlook-mcp'),
        { recursive: true, mode: 0o700 }
      );
    });

    it('wraps plain text in <pre> tag when content_type is text', () => {
      mockExistsSync.mockReturnValue(true);

      writeSignature('-- Joel\nSenior Dev', 'text');

      expect(mockWriteFileSync).toHaveBeenCalledWith(
        join('/mock/home', '.outlook-mcp', 'signature.html'),
        '<pre>-- Joel\nSenior Dev</pre>',
        { encoding: 'utf-8', mode: 0o600 }
      );
    });

    it('stores HTML content directly when content_type is html', () => {
      mockExistsSync.mockReturnValue(true);

      writeSignature('<b>Joel</b>', 'html');

      expect(mockWriteFileSync).toHaveBeenCalledWith(
        join('/mock/home', '.outlook-mcp', 'signature.html'),
        '<b>Joel</b>',
        { encoding: 'utf-8', mode: 0o600 }
      );
    });
  });

  describe('appendSignature', () => {
    it('appends signature to HTML body with <br><br> separator', () => {
      mockExistsSync.mockReturnValue(true);
      mockReadFileSync.mockReturnValue('<p>-- Joel</p>');

      const result = appendSignature('<p>Hello</p>', 'html', true);

      expect(result).toBe('<p>Hello</p><br><br><p>-- Joel</p>');
    });

    it('appends signature to text body with \\n\\n--\\n separator and strips HTML', () => {
      mockExistsSync.mockReturnValue(true);
      mockReadFileSync.mockReturnValue('<p>Best regards,<br>Joel</p>');

      const result = appendSignature('Hello World', 'text', true);

      expect(result).toBe('Hello World\n\n--\nBest regards,\nJoel');
    });

    it('returns body unchanged when includeSignature is false', () => {
      const result = appendSignature('Hello', 'text', false);

      expect(result).toBe('Hello');
      expect(mockExistsSync).not.toHaveBeenCalled();
    });

    it('returns body unchanged when no signature file exists', () => {
      mockExistsSync.mockReturnValue(false);

      const result = appendSignature('Hello', 'html', true);

      expect(result).toBe('Hello');
    });

    it('handles signature with nested HTML tags for text stripping', () => {
      mockExistsSync.mockReturnValue(true);
      mockReadFileSync.mockReturnValue('<div><b>Joel</b> | <a href="https://example.com">Site</a></div>');

      const result = appendSignature('Hi', 'text', true);

      expect(result).toBe('Hi\n\n--\nJoel | Site');
    });
  });
});
