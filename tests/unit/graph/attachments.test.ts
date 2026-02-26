/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for attachment upload/download helpers.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

// Mock fs before importing the module under test
vi.mock('fs', async () => {
  const actual = await vi.importActual<typeof import('fs')>('fs');
  return {
    ...actual,
    existsSync: vi.fn(),
    mkdirSync: vi.fn(),
    readFileSync: vi.fn(),
    writeFileSync: vi.fn(),
    statSync: vi.fn(),
  };
});

// We'll import the module under test after mocks are set up
import {
  getDownloadDir,
  sanitizeFilename,
  resolveFilePath,
  uploadAttachment,
  uploadInlineAttachment,
  downloadAttachment,
} from '../../../src/graph/attachments.js';
import type { GraphClient } from '../../../src/graph/client/index.js';

// =============================================================================
// sanitizeFilename
// =============================================================================

describe('sanitizeFilename', () => {
  it('returns a normal filename unchanged', () => {
    expect(sanitizeFilename('report.pdf')).toBe('report.pdf');
  });

  it('strips forward slash path components', () => {
    expect(sanitizeFilename('/etc/passwd')).toBe('passwd');
  });

  it('strips backslash path components', () => {
    expect(sanitizeFilename('C:\\Users\\file.txt')).toBe('file.txt');
  });

  it('strips dot-dot segments', () => {
    expect(sanitizeFilename('../../secret.txt')).toBe('secret.txt');
  });

  it('strips mixed path traversal', () => {
    expect(sanitizeFilename('../foo/../../bar/baz.txt')).toBe('baz.txt');
  });

  it('trims whitespace', () => {
    expect(sanitizeFilename('  file.txt  ')).toBe('file.txt');
  });

  it('falls back to "attachment" for empty string', () => {
    expect(sanitizeFilename('')).toBe('attachment');
  });

  it('falls back to "attachment" for whitespace-only', () => {
    expect(sanitizeFilename('   ')).toBe('attachment');
  });

  it('falls back to "attachment" for only path separators', () => {
    expect(sanitizeFilename('///')).toBe('attachment');
  });

  it('falls back to "attachment" for only dot-dot segments', () => {
    expect(sanitizeFilename('..')).toBe('attachment');
  });

  it('handles filenames with spaces', () => {
    expect(sanitizeFilename('my document.pdf')).toBe('my document.pdf');
  });
});

// =============================================================================
// resolveFilePath
// =============================================================================

describe('resolveFilePath', () => {
  beforeEach(() => {
    vi.mocked(fs.existsSync).mockReset();
  });

  it('returns the full path when file does not exist', () => {
    vi.mocked(fs.existsSync).mockReturnValue(false);

    const result = resolveFilePath('/tmp/downloads', 'report.pdf');

    expect(result).toBe(path.join('/tmp/downloads', 'report.pdf'));
  });

  it('sanitizes the filename before resolving', () => {
    vi.mocked(fs.existsSync).mockReturnValue(false);

    const result = resolveFilePath('/tmp/downloads', '../../evil.txt');

    expect(result).toBe(path.join('/tmp/downloads', 'evil.txt'));
  });

  it('appends numeric suffix when file exists', () => {
    vi.mocked(fs.existsSync)
      .mockReturnValueOnce(true)   // report.pdf exists
      .mockReturnValueOnce(false); // report(1).pdf does not

    const result = resolveFilePath('/tmp/downloads', 'report.pdf');

    expect(result).toBe(path.join('/tmp/downloads', 'report(1).pdf'));
  });

  it('increments suffix until unique name found', () => {
    vi.mocked(fs.existsSync)
      .mockReturnValueOnce(true)   // report.pdf exists
      .mockReturnValueOnce(true)   // report(1).pdf exists
      .mockReturnValueOnce(true)   // report(2).pdf exists
      .mockReturnValueOnce(false); // report(3).pdf does not

    const result = resolveFilePath('/tmp/downloads', 'report.pdf');

    expect(result).toBe(path.join('/tmp/downloads', 'report(3).pdf'));
  });

  it('handles filenames without extension', () => {
    vi.mocked(fs.existsSync)
      .mockReturnValueOnce(true)   // README exists
      .mockReturnValueOnce(false); // README(1) does not

    const result = resolveFilePath('/tmp/downloads', 'README');

    expect(result).toBe(path.join('/tmp/downloads', 'README(1)'));
  });

  it('handles dotfiles', () => {
    vi.mocked(fs.existsSync)
      .mockReturnValueOnce(true)   // .gitignore exists
      .mockReturnValueOnce(false); // .gitignore(1) does not

    const result = resolveFilePath('/tmp/downloads', '.gitignore');

    expect(result).toBe(path.join('/tmp/downloads', '.gitignore(1)'));
  });
});

// =============================================================================
// getDownloadDir
// =============================================================================

describe('getDownloadDir', () => {
  const originalEnv = process.env;

  beforeEach(() => {
    vi.mocked(fs.mkdirSync).mockReset();
    process.env = { ...originalEnv };
  });

  afterEach(() => {
    process.env = originalEnv;
  });

  it('uses MCP_OUTLOOK_DOWNLOAD_DIR env var when set', () => {
    process.env['MCP_OUTLOOK_DOWNLOAD_DIR'] = '/custom/downloads';
    vi.mocked(fs.mkdirSync).mockReturnValue(undefined);

    const result = getDownloadDir();

    expect(result).toBe(path.join('/custom/downloads', 'mcp-outlook-attachments'));
    expect(fs.mkdirSync).toHaveBeenCalledWith(
      path.join('/custom/downloads', 'mcp-outlook-attachments'),
      { recursive: true }
    );
  });

  it('falls back to os.tmpdir() when env var is not set', () => {
    delete process.env['MCP_OUTLOOK_DOWNLOAD_DIR'];
    vi.mocked(fs.mkdirSync).mockReturnValue(undefined);

    const result = getDownloadDir();

    const expected = path.join(os.tmpdir(), 'mcp-outlook-attachments');
    expect(result).toBe(expected);
    expect(fs.mkdirSync).toHaveBeenCalledWith(expected, { recursive: true });
  });
});

// =============================================================================
// uploadAttachment
// =============================================================================

describe('uploadAttachment', () => {
  let mockClient: {
    addAttachment: ReturnType<typeof vi.fn>;
    createUploadSession: ReturnType<typeof vi.fn>;
  };

  beforeEach(() => {
    mockClient = {
      addAttachment: vi.fn().mockResolvedValue({}),
      createUploadSession: vi.fn().mockResolvedValue({ uploadUrl: 'https://upload.example.com/session' }),
    };
    vi.mocked(fs.readFileSync).mockReset();
    vi.mocked(fs.statSync).mockReset();

    // Reset global fetch mock
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue({
      ok: true,
      json: () => Promise.resolve({}),
    }));
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('uses inline upload for files <= 3MB', async () => {
    const fileContent = Buffer.alloc(1024 * 1024); // 1MB
    vi.mocked(fs.statSync).mockReturnValue({ size: fileContent.length } as fs.Stats);
    vi.mocked(fs.readFileSync).mockReturnValue(fileContent);

    await uploadAttachment(
      mockClient as unknown as GraphClient,
      'msg-1',
      '/path/to/file.txt'
    );

    expect(mockClient.addAttachment).toHaveBeenCalledWith('msg-1', {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'file.txt',
      contentBytes: fileContent.toString('base64'),
      contentType: 'application/octet-stream',
    });
    expect(mockClient.createUploadSession).not.toHaveBeenCalled();
  });

  it('uses inline upload at exactly 3MB boundary', async () => {
    const threeBytes = 3 * 1024 * 1024;
    const fileContent = Buffer.alloc(threeBytes);
    vi.mocked(fs.statSync).mockReturnValue({ size: threeBytes } as fs.Stats);
    vi.mocked(fs.readFileSync).mockReturnValue(fileContent);

    await uploadAttachment(
      mockClient as unknown as GraphClient,
      'msg-1',
      '/path/to/file.pdf'
    );

    expect(mockClient.addAttachment).toHaveBeenCalled();
    expect(mockClient.createUploadSession).not.toHaveBeenCalled();
  });

  it('uses upload session for files > 3MB', async () => {
    const fileSize = 3 * 1024 * 1024 + 1; // just over 3MB
    const fileContent = Buffer.alloc(fileSize);
    vi.mocked(fs.statSync).mockReturnValue({ size: fileSize } as fs.Stats);
    vi.mocked(fs.readFileSync).mockReturnValue(fileContent);

    await uploadAttachment(
      mockClient as unknown as GraphClient,
      'msg-1',
      '/path/to/large-file.zip'
    );

    expect(mockClient.addAttachment).not.toHaveBeenCalled();
    expect(mockClient.createUploadSession).toHaveBeenCalledWith('msg-1', {
      AttachmentItem: {
        attachmentType: 'file',
        name: 'large-file.zip',
        size: fileSize,
      },
    });
    expect(fetch).toHaveBeenCalled();
  });

  it('sends correct Content-Range headers for chunked upload', async () => {
    const chunkSize = 3932160; // 3.75MB
    const fileSize = chunkSize * 2 + 1000; // ~8.5MB: needs 3 chunks
    const fileContent = Buffer.alloc(fileSize);
    vi.mocked(fs.statSync).mockReturnValue({ size: fileSize } as fs.Stats);
    vi.mocked(fs.readFileSync).mockReturnValue(fileContent);

    await uploadAttachment(
      mockClient as unknown as GraphClient,
      'msg-1',
      '/path/to/big-file.iso'
    );

    const fetchMock = vi.mocked(fetch);
    expect(fetchMock).toHaveBeenCalledTimes(3);

    // First chunk: bytes 0 to chunkSize-1
    expect(fetchMock.mock.calls[0]![1]!.headers).toEqual(
      expect.objectContaining({
        'Content-Range': `bytes 0-${chunkSize - 1}/${fileSize}`,
      })
    );

    // Second chunk: bytes chunkSize to 2*chunkSize-1
    expect(fetchMock.mock.calls[1]![1]!.headers).toEqual(
      expect.objectContaining({
        'Content-Range': `bytes ${chunkSize}-${2 * chunkSize - 1}/${fileSize}`,
      })
    );

    // Third chunk: remaining bytes
    const lastChunkEnd = fileSize - 1;
    expect(fetchMock.mock.calls[2]![1]!.headers).toEqual(
      expect.objectContaining({
        'Content-Range': `bytes ${2 * chunkSize}-${lastChunkEnd}/${fileSize}`,
      })
    );
  });

  it('uses custom name when provided', async () => {
    const fileContent = Buffer.alloc(1024);
    vi.mocked(fs.statSync).mockReturnValue({ size: 1024 } as fs.Stats);
    vi.mocked(fs.readFileSync).mockReturnValue(fileContent);

    await uploadAttachment(
      mockClient as unknown as GraphClient,
      'msg-1',
      '/path/to/file.txt',
      'custom-name.doc'
    );

    expect(mockClient.addAttachment).toHaveBeenCalledWith('msg-1', expect.objectContaining({
      name: 'custom-name.doc',
    }));
  });

  it('uses custom content type when provided', async () => {
    const fileContent = Buffer.alloc(1024);
    vi.mocked(fs.statSync).mockReturnValue({ size: 1024 } as fs.Stats);
    vi.mocked(fs.readFileSync).mockReturnValue(fileContent);

    await uploadAttachment(
      mockClient as unknown as GraphClient,
      'msg-1',
      '/path/to/file.txt',
      undefined,
      'text/plain'
    );

    expect(mockClient.addAttachment).toHaveBeenCalledWith('msg-1', expect.objectContaining({
      contentType: 'text/plain',
    }));
  });

  it('defaults name to path.basename(filePath)', async () => {
    const fileContent = Buffer.alloc(1024);
    vi.mocked(fs.statSync).mockReturnValue({ size: 1024 } as fs.Stats);
    vi.mocked(fs.readFileSync).mockReturnValue(fileContent);

    await uploadAttachment(
      mockClient as unknown as GraphClient,
      'msg-1',
      '/some/deep/path/to/document.pdf'
    );

    expect(mockClient.addAttachment).toHaveBeenCalledWith('msg-1', expect.objectContaining({
      name: 'document.pdf',
    }));
  });

  it('throws when chunked upload fetch returns non-ok response', async () => {
    const fileSize = 3 * 1024 * 1024 + 1;
    const fileContent = Buffer.alloc(fileSize);
    vi.mocked(fs.statSync).mockReturnValue({ size: fileSize } as fs.Stats);
    vi.mocked(fs.readFileSync).mockReturnValue(fileContent);

    vi.stubGlobal('fetch', vi.fn().mockResolvedValue({
      ok: false,
      status: 500,
      statusText: 'Internal Server Error',
    }));

    await expect(
      uploadAttachment(mockClient as unknown as GraphClient, 'msg-1', '/path/to/large.zip')
    ).rejects.toThrow('Upload chunk failed');
  });
});

// =============================================================================
// uploadInlineAttachment
// =============================================================================

describe('uploadInlineAttachment', () => {
  let mockClient: { addAttachment: ReturnType<typeof vi.fn> };

  beforeEach(() => {
    mockClient = { addAttachment: vi.fn().mockResolvedValue({}) };
    vi.mocked(fs.readFileSync).mockReset();
    vi.mocked(fs.statSync).mockReset();
  });

  it('posts file as inline attachment with contentId', async () => {
    const fileContent = Buffer.alloc(100);
    vi.mocked(fs.statSync).mockReturnValue({ size: 100 } as fs.Stats);
    vi.mocked(fs.readFileSync).mockReturnValue(fileContent);

    await uploadInlineAttachment(
      mockClient as unknown as GraphClient,
      'msg-1',
      '/path/to/logo.png',
      'logo'
    );

    expect(mockClient.addAttachment).toHaveBeenCalledWith('msg-1', {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'logo.png',
      contentBytes: fileContent.toString('base64'),
      contentType: 'image/png',
      isInline: true,
      contentId: 'logo',
    });
  });

  it('defaults to image/png for unknown extension', async () => {
    const fileContent = Buffer.alloc(50);
    vi.mocked(fs.statSync).mockReturnValue({ size: 50 } as fs.Stats);
    vi.mocked(fs.readFileSync).mockReturnValue(fileContent);

    await uploadInlineAttachment(
      mockClient as unknown as GraphClient,
      'msg-1',
      '/path/to/image.xyz',
      'img1'
    );

    expect(mockClient.addAttachment).toHaveBeenCalledWith('msg-1', expect.objectContaining({
      contentType: 'image/png',
      isInline: true,
      contentId: 'img1',
    }));
  });

  it('throws when file exceeds 3MB', async () => {
    const overLimit = 3 * 1024 * 1024 + 1;
    vi.mocked(fs.statSync).mockReturnValue({ size: overLimit } as fs.Stats);

    await expect(
      uploadInlineAttachment(
        mockClient as unknown as GraphClient,
        'msg-1',
        '/path/to/large.png',
        'big'
      )
    ).rejects.toThrow(/Inline image too large/);

    expect(mockClient.addAttachment).not.toHaveBeenCalled();
  });
});

// =============================================================================
// downloadAttachment
// =============================================================================

describe('downloadAttachment', () => {
  let mockClient: {
    getAttachment: ReturnType<typeof vi.fn>;
  };

  const originalEnv = process.env;

  beforeEach(() => {
    mockClient = {
      getAttachment: vi.fn(),
    };
    vi.mocked(fs.existsSync).mockReset();
    vi.mocked(fs.writeFileSync).mockReset();
    vi.mocked(fs.mkdirSync).mockReset().mockReturnValue(undefined);
    process.env = { ...originalEnv };
    delete process.env['MCP_OUTLOOK_DOWNLOAD_DIR'];
  });

  afterEach(() => {
    process.env = originalEnv;
  });

  it('downloads attachment and writes to file', async () => {
    const content = Buffer.from('Hello, World!');
    const base64Content = content.toString('base64');

    mockClient.getAttachment.mockResolvedValue({
      name: 'hello.txt',
      size: content.length,
      contentType: 'text/plain',
      contentBytes: base64Content,
    });

    vi.mocked(fs.existsSync).mockReturnValue(false);

    const result = await downloadAttachment(
      mockClient as unknown as GraphClient,
      'msg-1',
      'att-1'
    );

    expect(result.name).toBe('hello.txt');
    expect(result.size).toBe(content.length);
    expect(result.contentType).toBe('text/plain');
    expect(result.filePath).toContain('hello.txt');
    expect(fs.writeFileSync).toHaveBeenCalledWith(
      expect.stringContaining('hello.txt'),
      content
    );
  });

  it('falls back to "attachment" when name is empty', async () => {
    mockClient.getAttachment.mockResolvedValue({
      name: '',
      size: 10,
      contentType: 'application/octet-stream',
      contentBytes: Buffer.from('data').toString('base64'),
    });

    vi.mocked(fs.existsSync).mockReturnValue(false);

    const result = await downloadAttachment(
      mockClient as unknown as GraphClient,
      'msg-1',
      'att-2'
    );

    expect(result.name).toBe('attachment');
    expect(result.filePath).toContain('attachment');
  });

  it('calls client.getAttachment with correct parameters', async () => {
    mockClient.getAttachment.mockResolvedValue({
      name: 'file.pdf',
      size: 100,
      contentType: 'application/pdf',
      contentBytes: Buffer.from('pdf data').toString('base64'),
    });

    vi.mocked(fs.existsSync).mockReturnValue(false);

    await downloadAttachment(
      mockClient as unknown as GraphClient,
      'message-abc',
      'attachment-xyz'
    );

    expect(mockClient.getAttachment).toHaveBeenCalledWith('message-abc', 'attachment-xyz');
  });

  it('handles duplicate filenames via resolveFilePath', async () => {
    mockClient.getAttachment.mockResolvedValue({
      name: 'report.pdf',
      size: 50,
      contentType: 'application/pdf',
      contentBytes: Buffer.from('data').toString('base64'),
    });

    // First call: file exists, second: unique name found
    vi.mocked(fs.existsSync)
      .mockReturnValueOnce(true)
      .mockReturnValueOnce(false);

    const result = await downloadAttachment(
      mockClient as unknown as GraphClient,
      'msg-1',
      'att-1'
    );

    expect(result.filePath).toContain('report(1).pdf');
  });
});
