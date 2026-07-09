/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Microsoft OneNote tools (Graph API only).
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { OneNoteTools } from '../../../src/tools/onenote.js';
import { mintSelfEncoded } from '../../../src/ids/token.js';
import type { GraphRepository } from '../../../src/graph/repository.js';

describe('OneNoteTools', () => {
  let mockClient: {
    listNotebooks: ReturnType<typeof vi.fn>;
    listNoteSections: ReturnType<typeof vi.fn>;
    listNotePages: ReturnType<typeof vi.fn>;
    getNotePage: ReturnType<typeof vi.fn>;
    getNotePageContent: ReturnType<typeof vi.fn>;
    searchNotePages: ReturnType<typeof vi.fn>;
    createNotePage: ReturnType<typeof vi.fn>;
  };
  let repository: GraphRepository;
  let tools: OneNoteTools;

  beforeEach(() => {
    mockClient = {
      listNotebooks: vi.fn(),
      listNoteSections: vi.fn(),
      listNotePages: vi.fn(),
      getNotePage: vi.fn(),
      getNotePageContent: vi.fn(),
      searchNotePages: vi.fn(),
      createNotePage: vi.fn(),
    };
    repository = {
      getClient: () => mockClient,
    } as unknown as GraphRepository;
    tools = new OneNoteTools(repository);
  });

  // ===========================================================================
  // List Notebooks
  // ===========================================================================

  describe('listNotebooks', () => {
    it('returns notebooks with minted noteNotebook tokens', async () => {
      mockClient.listNotebooks.mockResolvedValue([
        { id: 'nb-graph-1', displayName: 'Work', createdDateTime: '2026-01-01T00:00:00Z', lastModifiedDateTime: '2026-01-02T00:00:00Z' },
      ]);

      const result = await tools.listNotebooks();

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.notebooks).toHaveLength(1);
      expect(parsed.notebooks[0].id).toBe(mintSelfEncoded('noteNotebook', 'nb-graph-1'));
      expect(parsed.notebooks[0].displayName).toBe('Work');
    });
  });

  // ===========================================================================
  // List Note Sections
  // ===========================================================================

  describe('listNoteSections', () => {
    it('lists all sections when no notebook_id given', async () => {
      mockClient.listNoteSections.mockResolvedValue([
        { id: 'sec-graph-1', displayName: 'General', createdDateTime: '2026-01-01T00:00:00Z', lastModifiedDateTime: '2026-01-02T00:00:00Z' },
      ]);

      const result = await tools.listNoteSections({});

      expect(mockClient.listNoteSections).toHaveBeenCalledWith(undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.sections[0].id).toBe(mintSelfEncoded('noteSection', 'sec-graph-1'));
    });

    it('resolves a notebook_id token to its Graph ID', async () => {
      mockClient.listNoteSections.mockResolvedValue([]);
      const notebookToken = mintSelfEncoded('noteNotebook', 'nb-graph-1');

      await tools.listNoteSections({ notebook_id: notebookToken });

      expect(mockClient.listNoteSections).toHaveBeenCalledWith('nb-graph-1');
    });
  });

  // ===========================================================================
  // List Note Pages
  // ===========================================================================

  describe('listNotePages', () => {
    it('lists all pages when no section_id given', async () => {
      mockClient.listNotePages.mockResolvedValue([
        { id: 'page-graph-1', title: 'Meeting Notes', createdDateTime: '2026-01-01T00:00:00Z', lastModifiedDateTime: '2026-01-02T00:00:00Z' },
      ]);

      const result = await tools.listNotePages({});

      expect(mockClient.listNotePages).toHaveBeenCalledWith(undefined);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.pages[0].id).toBe(mintSelfEncoded('notePage', 'page-graph-1'));
      expect(parsed.pages[0].title).toBe('Meeting Notes');
    });

    it('resolves a section_id token to its Graph ID', async () => {
      mockClient.listNotePages.mockResolvedValue([]);
      const sectionToken = mintSelfEncoded('noteSection', 'sec-graph-1');

      await tools.listNotePages({ section_id: sectionToken });

      expect(mockClient.listNotePages).toHaveBeenCalledWith('sec-graph-1');
    });
  });

  // ===========================================================================
  // Get Note Page
  // ===========================================================================

  describe('getNotePage', () => {
    it('resolves the page token, fetches metadata + content, and returns both', async () => {
      const pageToken = mintSelfEncoded('notePage', 'page-graph-1');
      mockClient.getNotePage.mockResolvedValue({
        id: 'page-graph-1', title: 'Meeting Notes', createdDateTime: '2026-01-01T00:00:00Z', lastModifiedDateTime: '2026-01-02T00:00:00Z',
      });
      mockClient.getNotePageContent.mockResolvedValue('<html><body><p>Notes</p></body></html>');

      const result = await tools.getNotePage({ page_id: pageToken });

      expect(mockClient.getNotePage).toHaveBeenCalledWith('page-graph-1');
      expect(mockClient.getNotePageContent).toHaveBeenCalledWith('page-graph-1');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.id).toBe(pageToken);
      expect(parsed.title).toBe('Meeting Notes');
      expect(parsed.content_html).toBe('<html><body><p>Notes</p></body></html>');
    });

    it('rejects a wrong-entity token (a section token as page_id) with ID_ENTITY_MISMATCH', async () => {
      const sectionToken = mintSelfEncoded('noteSection', 'sec-graph-1');
      await expect(tools.getNotePage({ page_id: sectionToken })).rejects.toMatchObject({
        code: 'ID_ENTITY_MISMATCH',
      });
      expect(mockClient.getNotePage).not.toHaveBeenCalled();
    });
  });

  // ===========================================================================
  // Search Note Pages
  // ===========================================================================

  describe('searchNotePages', () => {
    it('returns pages matching the query with minted tokens', async () => {
      mockClient.searchNotePages.mockResolvedValue([
        { id: 'page-graph-2', title: 'Budget', createdDateTime: '2026-01-01T00:00:00Z', lastModifiedDateTime: '2026-01-02T00:00:00Z' },
      ]);

      const result = await tools.searchNotePages({ query: 'budget' });

      expect(mockClient.searchNotePages).toHaveBeenCalledWith('budget');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.pages[0].id).toBe(mintSelfEncoded('notePage', 'page-graph-2'));
    });
  });

  // ===========================================================================
  // Create Note Page
  // ===========================================================================

  describe('createNotePage', () => {
    it('resolves the section token, wraps content in HTML, and returns the minted page id', async () => {
      const sectionToken = mintSelfEncoded('noteSection', 'sec-graph-1');
      mockClient.createNotePage.mockResolvedValue({ id: 'page-graph-new' });

      const result = await tools.createNotePage({
        section_id: sectionToken,
        title: 'New Page',
        content_html: '<p>Hello</p>',
      });

      expect(mockClient.createNotePage).toHaveBeenCalledWith('sec-graph-1', expect.stringContaining('<p>Hello</p>'));
      const [, html] = mockClient.createNotePage.mock.calls[0] as [string, string];
      expect(html).toContain('<title>New Page</title>');

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.id).toBe(mintSelfEncoded('notePage', 'page-graph-new'));
      expect(parsed.title).toBe('New Page');
      expect(parsed.status).toBe('created');
    });

    it('escapes HTML special characters in the title', async () => {
      const sectionToken = mintSelfEncoded('noteSection', 'sec-graph-1');
      mockClient.createNotePage.mockResolvedValue({ id: 'page-graph-new' });

      await tools.createNotePage({
        section_id: sectionToken,
        title: '<script>alert(1)</script>',
        content_html: '<p>Body</p>',
      });

      const [, html] = mockClient.createNotePage.mock.calls[0] as [string, string];
      expect(html).not.toContain('<script>alert(1)</script>');
      expect(html).toContain('&lt;script&gt;');
    });
  });
});
