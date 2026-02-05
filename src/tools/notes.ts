/**
 * Notes-related MCP tools.
 *
 * Provides tools for listing, searching, and getting notes.
 */

import { z } from 'zod';
import type { IRepository, NoteRow } from '../database/repository.js';
import type { NoteSummary, Note } from '../types/index.js';
import { appleTimestampToIso } from '../utils/dates.js';

// =============================================================================
// Input Schemas
// =============================================================================

export const ListNotesInput = z.strictObject({
  limit: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(50)
    .describe('Maximum number of notes to return (1-100)'),
  offset: z.number().int().min(0).default(0).describe('Number of notes to skip'),
});

export const GetNoteInput = z.strictObject({
  note_id: z.number().int().positive().describe('The note ID to retrieve'),
});

export const SearchNotesInput = z.strictObject({
  query: z.string().min(1).describe('Search query for note content'),
  limit: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(50)
    .describe('Maximum number of notes to return (1-100)'),
});

// =============================================================================
// Type Definitions
// =============================================================================

export type ListNotesParams = z.infer<typeof ListNotesInput>;
export type GetNoteParams = z.infer<typeof GetNoteInput>;
export type SearchNotesParams = z.infer<typeof SearchNotesInput>;

// =============================================================================
// Content Reader Interface
// =============================================================================

/**
 * Interface for reading note content from data files.
 */
export interface INoteContentReader {
  /**
   * Reads note details from the given data file path.
   */
  readNoteDetails(dataFilePath: string | null): NoteDetails | null;
}

/**
 * Note details from content file.
 */
export interface NoteDetails {
  readonly title: string | null;
  readonly body: string | null;
  readonly preview: string | null;
  readonly createdDate: string | null;
  readonly categories: readonly string[];
}

/**
 * Default note content reader that returns null.
 */
export const nullNoteContentReader: INoteContentReader = {
  readNoteDetails: (): NoteDetails | null => null,
};

// =============================================================================
// Transformers
// =============================================================================

/**
 * Transforms a database note row to NoteSummary.
 */
function transformNoteSummary(row: NoteRow, details: NoteDetails | null): NoteSummary {
  return {
    id: row.id,
    folderId: row.folderId,
    title: details?.title ?? null,
    preview: details?.preview ?? null,
    modifiedDate: appleTimestampToIso(row.modifiedDate),
  };
}

/**
 * Transforms a database note row to full Note.
 */
function transformNote(row: NoteRow, details: NoteDetails | null): Note {
  const summary = transformNoteSummary(row, details);

  return {
    ...summary,
    body: details?.body ?? null,
    createdDate: details?.createdDate ?? null,
    categories: details?.categories ?? [],
  };
}

// =============================================================================
// Notes Tools Class
// =============================================================================

/**
 * Notes tools implementation with dependency injection.
 */
export class NotesTools {
  constructor(
    private readonly repository: IRepository,
    private readonly contentReader: INoteContentReader = nullNoteContentReader
  ) {}

  /**
   * Lists notes with pagination.
   */
  listNotes(params: ListNotesParams): NoteSummary[] {
    const { limit, offset } = params;
    const rows = this.repository.listNotes(limit, offset);
    return rows.map((row) => {
      const details = this.contentReader.readNoteDetails(row.dataFilePath);
      return transformNoteSummary(row, details);
    });
  }

  /**
   * Gets a single note by ID.
   */
  getNote(params: GetNoteParams): Note | null {
    const { note_id } = params;

    const row = this.repository.getNote(note_id);
    if (row == null) {
      return null;
    }

    const details = this.contentReader.readNoteDetails(row.dataFilePath);
    return transformNote(row, details);
  }

  /**
   * Searches notes by content (requires content reader).
   */
  searchNotes(params: SearchNotesParams): NoteSummary[] {
    const { query, limit } = params;
    const queryLower = query.toLowerCase();

    // Get all notes and filter by content (from content reader)
    const rows = this.repository.listNotes(limit * 2, 0); // Fetch more to filter
    const results: NoteSummary[] = [];

    for (const row of rows) {
      if (results.length >= limit) break;

      const details = this.contentReader.readNoteDetails(row.dataFilePath);
      const searchText = [details?.title ?? '', details?.body ?? '', details?.preview ?? '']
        .join(' ')
        .toLowerCase();

      if (searchText.includes(queryLower)) {
        results.push(transformNoteSummary(row, details));
      }
    }

    return results;
  }
}

/**
 * Creates notes tools with the given repository.
 */
export function createNotesTools(
  repository: IRepository,
  contentReader: INoteContentReader = nullNoteContentReader
): NotesTools {
  return new NotesTools(repository, contentReader);
}
