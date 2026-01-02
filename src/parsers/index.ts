/**
 * Parsers for Outlook content.
 *
 * Re-exports all parsers for convenient importing.
 */

export {
  stripHtml,
  containsHtml,
  extractPlainText,
  type StripHtmlOptions,
} from './html-stripper.js';

export {
  parseOlk15File,
  getDefaultDataPath,
  Olk15EmailContentReader,
  Olk15EventContentReader,
  Olk15ContactContentReader,
  Olk15TaskContentReader,
  Olk15NoteContentReader,
  createContentReaders,
  createDefaultContentReaders,
  type Olk15ParseResult,
  type ContentReaders,
} from './olk15.js';
