/**
 * Tests for Graph API token cache.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import {
  FileTokenCachePlugin,
  createTokenCachePlugin,
  hasTokenCache,
  clearTokenCache,
  getTokenCacheDir,
  getTokenCacheFile,
} from '../../../../src/graph/auth/token-cache.js';

// Mock fs module
vi.mock('node:fs', () => ({
  existsSync: vi.fn(),
  mkdirSync: vi.fn(),
  readFileSync: vi.fn(),
  writeFileSync: vi.fn(),
}));

// Mock os module
vi.mock('node:os', () => ({
  homedir: vi.fn(() => '/mock/home'),
}));

describe('graph/auth/token-cache', () => {
  const mockExistsSync = vi.mocked(existsSync);
  const mockMkdirSync = vi.mocked(mkdirSync);
  const mockReadFileSync = vi.mocked(readFileSync);
  const mockWriteFileSync = vi.mocked(writeFileSync);
  const mockHomedir = vi.mocked(homedir);

  beforeEach(() => {
    vi.clearAllMocks();
    mockHomedir.mockReturnValue('/mock/home');
  });

  describe('getTokenCacheDir', () => {
    it('returns path in home directory', () => {
      const dir = getTokenCacheDir();
      expect(dir).toBe(join('/mock/home', '.outlook-mcp'));
    });
  });

  describe('getTokenCacheFile', () => {
    it('returns tokens.json path', () => {
      const file = getTokenCacheFile();
      expect(file).toBe(join('/mock/home', '.outlook-mcp', 'tokens.json'));
    });
  });

  describe('hasTokenCache', () => {
    it('returns true when cache file exists', () => {
      mockExistsSync.mockReturnValue(true);

      const result = hasTokenCache();

      expect(result).toBe(true);
      expect(mockExistsSync).toHaveBeenCalled();
    });

    it('returns false when cache file does not exist', () => {
      mockExistsSync.mockReturnValue(false);

      const result = hasTokenCache();

      expect(result).toBe(false);
    });
  });

  describe('clearTokenCache', () => {
    it('writes empty object when cache file exists', () => {
      mockExistsSync.mockReturnValue(true);

      clearTokenCache();

      expect(mockWriteFileSync).toHaveBeenCalledWith(
        expect.stringContaining('tokens.json'),
        '{}',
        { mode: 0o600 }
      );
    });

    it('does nothing when cache file does not exist', () => {
      mockExistsSync.mockReturnValue(false);

      clearTokenCache();

      expect(mockWriteFileSync).not.toHaveBeenCalled();
    });

    it('ignores write errors', () => {
      mockExistsSync.mockReturnValue(true);
      mockWriteFileSync.mockImplementation(() => {
        throw new Error('Write failed');
      });

      // Should not throw
      expect(() => clearTokenCache()).not.toThrow();
    });
  });

  describe('createTokenCachePlugin', () => {
    it('returns a FileTokenCachePlugin instance', () => {
      const plugin = createTokenCachePlugin();
      expect(plugin).toBeInstanceOf(FileTokenCachePlugin);
    });
  });

  describe('FileTokenCachePlugin', () => {
    let plugin: FileTokenCachePlugin;

    beforeEach(() => {
      plugin = new FileTokenCachePlugin();
    });

    describe('beforeCacheAccess', () => {
      it('loads cache from file when it exists', async () => {
        mockExistsSync.mockReturnValue(true);
        mockReadFileSync.mockReturnValue('{"tokens": "data"}');

        const mockContext = {
          tokenCache: {
            deserialize: vi.fn(),
          },
        } as any;

        await plugin.beforeCacheAccess(mockContext);

        expect(mockReadFileSync).toHaveBeenCalled();
        expect(mockContext.tokenCache.deserialize).toHaveBeenCalledWith('{"tokens": "data"}');
      });

      it('does nothing when file does not exist', async () => {
        mockExistsSync.mockReturnValue(false);

        const mockContext = {
          tokenCache: {
            deserialize: vi.fn(),
          },
        } as any;

        await plugin.beforeCacheAccess(mockContext);

        expect(mockReadFileSync).not.toHaveBeenCalled();
        expect(mockContext.tokenCache.deserialize).not.toHaveBeenCalled();
      });

      it('ignores read errors', async () => {
        mockExistsSync.mockReturnValue(true);
        mockReadFileSync.mockImplementation(() => {
          throw new Error('Read failed');
        });

        const mockContext = {
          tokenCache: {
            deserialize: vi.fn(),
          },
        } as any;

        // Should not throw
        await expect(plugin.beforeCacheAccess(mockContext)).resolves.not.toThrow();
      });
    });

    describe('afterCacheAccess', () => {
      it('writes cache to file when changed', async () => {
        mockExistsSync.mockReturnValue(true);

        const mockContext = {
          cacheHasChanged: true,
          tokenCache: {
            serialize: vi.fn().mockReturnValue('{"updated": "tokens"}'),
          },
        } as any;

        await plugin.afterCacheAccess(mockContext);

        expect(mockContext.tokenCache.serialize).toHaveBeenCalled();
        expect(mockWriteFileSync).toHaveBeenCalledWith(
          expect.stringContaining('tokens.json'),
          '{"updated": "tokens"}',
          { mode: 0o600 }
        );
      });

      it('creates directory if it does not exist', async () => {
        mockExistsSync.mockReturnValueOnce(false); // Directory doesn't exist

        const mockContext = {
          cacheHasChanged: true,
          tokenCache: {
            serialize: vi.fn().mockReturnValue('{}'),
          },
        } as any;

        await plugin.afterCacheAccess(mockContext);

        expect(mockMkdirSync).toHaveBeenCalledWith(
          expect.stringContaining('.outlook-mcp'),
          { recursive: true, mode: 0o700 }
        );
      });

      it('does nothing when cache has not changed', async () => {
        const mockContext = {
          cacheHasChanged: false,
          tokenCache: {
            serialize: vi.fn(),
          },
        } as any;

        await plugin.afterCacheAccess(mockContext);

        expect(mockContext.tokenCache.serialize).not.toHaveBeenCalled();
        expect(mockWriteFileSync).not.toHaveBeenCalled();
      });

      it('ignores write errors', async () => {
        mockExistsSync.mockReturnValue(true);
        mockWriteFileSync.mockImplementation(() => {
          throw new Error('Write failed');
        });

        const mockContext = {
          cacheHasChanged: true,
          tokenCache: {
            serialize: vi.fn().mockReturnValue('{}'),
          },
        } as any;

        // Should not throw
        await expect(plugin.afterCacheAccess(mockContext)).resolves.not.toThrow();
      });
    });
  });
});
