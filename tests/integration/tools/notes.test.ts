import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { createTestDatabase, SAMPLE_COUNTS } from '../../fixtures/database.js';
import { createConnection, type IConnection } from '../../../src/database/connection.js';
import { createRepository, type IRepository } from '../../../src/database/repository.js';
import {
  NotesTools,
  createNotesTools,
  ListNotesInput,
  GetNoteInput,
  SearchNotesInput,
  type INoteContentReader,
  type NoteDetails,
} from '../../../src/tools/notes.js';

describe('NotesTools', () => {
  let testDb: { path: string; cleanup: () => void };
  let connection: IConnection;
  let repository: IRepository;
  let notesTools: NotesTools;

  beforeEach(() => {
    testDb = createTestDatabase();
    connection = createConnection(testDb.path);
    repository = createRepository(connection);
    notesTools = createNotesTools(repository);
  });

  afterEach(() => {
    connection.close();
    testDb.cleanup();
  });

  // ---------------------------------------------------------------------------
  // Input Validation
  // ---------------------------------------------------------------------------

  describe('input validation', () => {
    it('validates ListNotesInput with defaults', () => {
      const parsed = ListNotesInput.parse({});
      expect(parsed.limit).toBe(50);
      expect(parsed.offset).toBe(0);
    });

    it('validates ListNotesInput with options', () => {
      const input = { limit: 25, offset: 10 };
      const parsed = ListNotesInput.parse(input);
      expect(parsed).toEqual(input);
    });

    it('validates GetNoteInput', () => {
      const parsed = GetNoteInput.parse({ note_id: 1 });
      expect(parsed.note_id).toBe(1);
    });

    it('validates SearchNotesInput', () => {
      const parsed = SearchNotesInput.parse({ query: 'meeting' });
      expect(parsed.query).toBe('meeting');
      expect(parsed.limit).toBe(50);
    });
  });

  // ---------------------------------------------------------------------------
  // listNotes
  // ---------------------------------------------------------------------------

  describe('listNotes', () => {
    it('returns notes', () => {
      const notes = notesTools.listNotes({ limit: 50, offset: 0 });
      expect(notes.length).toBe(SAMPLE_COUNTS.notes);
    });

    it('returns notes with correct structure', () => {
      const notes = notesTools.listNotes({ limit: 1, offset: 0 });
      const note = notes[0];

      expect(note).toHaveProperty('id');
      expect(note).toHaveProperty('folderId');
      expect(note).toHaveProperty('title');
      expect(note).toHaveProperty('preview');
      expect(note).toHaveProperty('modifiedDate');
    });

    it('respects limit parameter', () => {
      const notes = notesTools.listNotes({ limit: 1, offset: 0 });
      expect(notes.length).toBe(1);
    });

    it('respects offset parameter', () => {
      const allNotes = notesTools.listNotes({ limit: 50, offset: 0 });
      const offsetNotes = notesTools.listNotes({ limit: 50, offset: 1 });
      expect(offsetNotes.length).toBe(allNotes.length - 1);
    });
  });

  // ---------------------------------------------------------------------------
  // getNote
  // ---------------------------------------------------------------------------

  describe('getNote', () => {
    it('returns note by ID', () => {
      const notes = notesTools.listNotes({ limit: 1, offset: 0 });
      const firstNote = notes[0];

      if (firstNote) {
        const note = notesTools.getNote({ note_id: firstNote.id });
        expect(note).not.toBeNull();
        expect(note?.id).toBe(firstNote.id);
      }
    });

    it('returns null for non-existent ID', () => {
      const note = notesTools.getNote({ note_id: 99999 });
      expect(note).toBeNull();
    });

    it('includes additional fields in full note', () => {
      const notes = notesTools.listNotes({ limit: 1, offset: 0 });
      const firstNote = notes[0];

      if (firstNote) {
        const note = notesTools.getNote({ note_id: firstNote.id });
        expect(note).toHaveProperty('body');
        expect(note).toHaveProperty('createdDate');
        expect(note).toHaveProperty('categories');
      }
    });
  });

  // ---------------------------------------------------------------------------
  // searchNotes
  // ---------------------------------------------------------------------------

  describe('searchNotes', () => {
    it('returns empty array when no content reader (content comes from files)', () => {
      const notes = notesTools.searchNotes({ query: 'meeting', limit: 50 });
      // Without a content reader, search can't find anything
      expect(notes.length).toBe(0);
    });
  });

  // ---------------------------------------------------------------------------
  // Content Reader Integration
  // ---------------------------------------------------------------------------

  describe('content reader integration', () => {
    it('uses content reader for note details', () => {
      const mockDetails: NoteDetails = {
        title: 'Meeting Notes',
        body: 'Discussion about project timeline',
        preview: 'Discussion about...',
        createdDate: '2024-01-15T10:00:00.000Z',
        categories: ['Work'],
      };

      const mockContentReader: INoteContentReader = {
        readNoteDetails: () => mockDetails,
      };

      const toolsWithReader = createNotesTools(repository, mockContentReader);
      const notes = toolsWithReader.listNotes({ limit: 1, offset: 0 });

      expect(notes[0]?.title).toBe('Meeting Notes');
      expect(notes[0]?.preview).toBe('Discussion about...');
    });

    it('can search notes when content reader provides content', () => {
      const mockContentReader: INoteContentReader = {
        readNoteDetails: () => ({
          title: 'Meeting Notes',
          body: 'Discussion about project timeline',
          preview: 'Discussion...',
          createdDate: null,
          categories: [],
        }),
      };

      const toolsWithReader = createNotesTools(repository, mockContentReader);
      const notes = toolsWithReader.searchNotes({ query: 'Meeting', limit: 50 });

      expect(notes.length).toBeGreaterThan(0);
      expect(notes[0]?.title).toBe('Meeting Notes');
    });

    it('gets full note details from content reader', () => {
      const mockDetails: NoteDetails = {
        title: 'Test Note',
        body: 'Full note body content',
        preview: 'Full note...',
        createdDate: '2024-01-15T10:00:00.000Z',
        categories: ['Personal', 'Ideas'],
      };

      const mockContentReader: INoteContentReader = {
        readNoteDetails: () => mockDetails,
      };

      const toolsWithReader = createNotesTools(repository, mockContentReader);
      const notes = toolsWithReader.listNotes({ limit: 1, offset: 0 });

      if (notes[0]) {
        const note = toolsWithReader.getNote({ note_id: notes[0].id });
        expect(note?.body).toBe('Full note body content');
        expect(note?.createdDate).toBe('2024-01-15T10:00:00.000Z');
        expect(note?.categories).toEqual(['Personal', 'Ideas']);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Factory Function
  // ---------------------------------------------------------------------------

  describe('createNotesTools', () => {
    it('creates a NotesTools instance', () => {
      const tools = createNotesTools(repository);
      expect(tools).toBeInstanceOf(NotesTools);
    });
  });
});
