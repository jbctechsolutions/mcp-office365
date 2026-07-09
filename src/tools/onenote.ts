/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Microsoft OneNote MCP tools (Graph API only).
 *
 * Covers the OneNote hierarchy — notebooks, sections, pages — with durable
 * self-encoding `nb_`/`ns_`/`np_` tokens (U5 / D1), since all three levels
 * are independently addressable.
 */

import { z } from 'zod';
import { Id } from '../ids/schema.js';
import { nextActionFor } from '../ids/next-action.js';
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolContext, ToolDefinition, ToolResult } from '../registry/types.js';
import type { GraphRepository } from '../graph/repository.js';
import { resolveId } from '../ids/resolver.js';
import { mintSelfEncoded, type EntityType } from '../ids/token.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    onenote: OneNoteTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const ListNotebooksInput = z.strictObject({});

export const ListNoteSectionsInput = z.strictObject({
  notebook_id: Id.noteNotebook.optional().describe('Notebook ID from list_notebooks (omit to list all sections).'),
});

export const ListNotePagesInput = z.strictObject({
  section_id: Id.noteSection.optional().describe('Section ID from list_note_sections (omit to list recent pages across all notebooks).'),
});

export const GetNotePageInput = z.strictObject({
  page_id: Id.notePage,
});

export const SearchNotePagesInput = z.strictObject({
  query: z.string().min(1).describe('Search query'),
});

export const CreateNotePageInput = z.strictObject({
  section_id: Id.noteSection,
  title: z.string().min(1).describe('Page title'),
  content_html: z.string().describe('Page body content as HTML'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListNotebooksParams = z.infer<typeof ListNotebooksInput>;
export type ListNoteSectionsParams = z.infer<typeof ListNoteSectionsInput>;
export type ListNotePagesParams = z.infer<typeof ListNotePagesInput>;
export type GetNotePageParams = z.infer<typeof GetNotePageInput>;
export type SearchNotePagesParams = z.infer<typeof SearchNotePagesInput>;
export type CreateNotePageParams = z.infer<typeof CreateNotePageInput>;

// =============================================================================
// Helpers
// =============================================================================

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// =============================================================================
// OneNote Tools
// =============================================================================

/**
 * Microsoft OneNote tools for browsing and creating notebooks, sections, and
 * pages via the Graph `/me/onenote` API.
 */
export class OneNoteTools {
  constructor(
    private readonly repository: GraphRepository,
  ) {}

  private toGraphId(id: string, entityType: EntityType): string {
    return resolveId(id, 'default', undefined, entityType).graphId;
  }

  async listNotebooks(): Promise<ToolResult> {
    const notebooks = await this.repository.getClient().listNotebooks();
    return jsonResult({
      next: nextActionFor('noteNotebook') ?? undefined,
      notebooks: notebooks.map((nb) => ({
        id: mintSelfEncoded('noteNotebook', nb.id ?? ''),
        displayName: nb.displayName ?? null,
        createdDateTime: nb.createdDateTime ?? null,
        lastModifiedDateTime: nb.lastModifiedDateTime ?? null,
      })),
    });
  }

  async listNoteSections(params: ListNoteSectionsParams): Promise<ToolResult> {
    const notebookGraphId = params.notebook_id != null
      ? this.toGraphId(params.notebook_id, 'noteNotebook')
      : undefined;
    const sections = await this.repository.getClient().listNoteSections(notebookGraphId);
    return jsonResult({
      notebook_id: params.notebook_id ?? null,
      next: nextActionFor('noteSection') ?? undefined,
      sections: sections.map((s) => ({
        id: mintSelfEncoded('noteSection', s.id ?? ''),
        displayName: s.displayName ?? null,
        createdDateTime: s.createdDateTime ?? null,
        lastModifiedDateTime: s.lastModifiedDateTime ?? null,
      })),
    });
  }

  async listNotePages(params: ListNotePagesParams): Promise<ToolResult> {
    const sectionGraphId = params.section_id != null
      ? this.toGraphId(params.section_id, 'noteSection')
      : undefined;
    const pages = await this.repository.getClient().listNotePages(sectionGraphId);
    return jsonResult({
      section_id: params.section_id ?? null,
      next: nextActionFor('notePage') ?? undefined,
      pages: pages.map((p) => ({
        id: mintSelfEncoded('notePage', p.id ?? ''),
        title: p.title ?? null,
        createdDateTime: p.createdDateTime ?? null,
        lastModifiedDateTime: p.lastModifiedDateTime ?? null,
      })),
    });
  }

  async getNotePage(params: GetNotePageParams): Promise<ToolResult> {
    const graphId = this.toGraphId(params.page_id, 'notePage');
    const client = this.repository.getClient();
    const page = await client.getNotePage(graphId);
    const contentHtml = await client.getNotePageContent(graphId);
    return jsonResult({
      id: mintSelfEncoded('notePage', page.id ?? ''),
      title: page.title ?? null,
      createdDateTime: page.createdDateTime ?? null,
      lastModifiedDateTime: page.lastModifiedDateTime ?? null,
      content_html: contentHtml,
    });
  }

  async searchNotePages(params: SearchNotePagesParams): Promise<ToolResult> {
    const pages = await this.repository.getClient().searchNotePages(params.query);
    return jsonResult({
      next: nextActionFor('notePage') ?? undefined,
      pages: pages.map((p) => ({
        id: mintSelfEncoded('notePage', p.id ?? ''),
        title: p.title ?? null,
        createdDateTime: p.createdDateTime ?? null,
        lastModifiedDateTime: p.lastModifiedDateTime ?? null,
      })),
    });
  }

  async createNotePage(params: CreateNotePageParams): Promise<ToolResult> {
    const sectionGraphId = this.toGraphId(params.section_id, 'noteSection');
    const html = `<!DOCTYPE html><html><head><title>${escapeHtml(params.title)}</title></head><body>${params.content_html}</body></html>`;
    const created = await this.repository.getClient().createNotePage(sectionGraphId, html);
    return jsonResult({
      id: mintSelfEncoded('notePage', created.id ?? ''),
      title: params.title,
      status: 'created',
      next: nextActionFor('notePage') ?? undefined,
    });
  }
}

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2)
// =============================================================================

/**
 * Registry tool definitions for the OneNote domain.
 */
export function onenoteToolDefinitions(): ToolDefinition[] {
  const tools = (ctx: ToolContext): OneNoteTools => requireGraphToolset(ctx, 'onenote');

  return [
    defineTool({
      name: 'list_notebooks',
      description: 'List OneNote notebooks for the current user (Graph API)',
      input: ListNotebooksInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['notes'],
      backends: ['graph'],
      handler: (ctx) => tools(ctx).listNotebooks(),
    }),
    defineTool({
      name: 'list_note_sections',
      description: 'List OneNote sections, optionally scoped to a notebook (Graph API)',
      input: ListNoteSectionsInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['notes'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listNoteSections(params),
    }),
    defineTool({
      name: 'list_note_pages',
      description: 'List OneNote pages, optionally scoped to a section (Graph API)',
      input: ListNotePagesInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['notes'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).listNotePages(params),
    }),
    defineTool({
      name: 'get_note_page',
      description: 'Get a OneNote page\'s metadata and HTML content (Graph API)',
      input: GetNotePageInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['notes'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).getNotePage(params),
    }),
    defineTool({
      name: 'search_note_pages',
      description: 'Search OneNote pages by keyword (Graph API)',
      input: SearchNotePagesInput,
      annotations: { readOnlyHint: true, openWorldHint: true },
      destructive: false,
      presets: ['notes'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).searchNotePages(params),
    }),
    defineTool({
      name: 'create_note_page',
      description: 'Create a new OneNote page in a section (Graph API)',
      input: CreateNotePageInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['notes'],
      backends: ['graph'],
      handler: (ctx, params) => tools(ctx).createNotePage(params),
    }),
  ];
}
