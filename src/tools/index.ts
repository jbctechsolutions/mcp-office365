/**
 * MCP Tool exports.
 *
 * Re-exports all tools and schemas for convenient importing.
 */

// Mail tools
export {
  MailTools,
  createMailTools,
  ListFoldersInput,
  ListEmailsInput,
  SearchEmailsInput,
  GetEmailInput,
  GetUnreadCountInput,
  type ListFoldersParams,
  type ListEmailsParams,
  type SearchEmailsParams,
  type GetEmailParams,
  type GetUnreadCountParams,
  type IContentReader,
  nullContentReader,
} from './mail.js';

// Calendar tools
export {
  CalendarTools,
  createCalendarTools,
  ListCalendarsInput,
  ListEventsInput,
  GetEventInput,
  SearchEventsInput,
  CreateEventInput,
  type ListCalendarsParams,
  type ListEventsParams,
  type GetEventParams,
  type SearchEventsParams,
  type CreateEventParams,
  type CreateEventResult,
  type RecurrenceParams,
  type IEventContentReader,
  type EventDetails,
  nullEventContentReader,
} from './calendar.js';

// Contacts tools
export {
  ContactsTools,
  createContactsTools,
  ListContactsInput,
  SearchContactsInput,
  GetContactInput,
  type ListContactsParams,
  type SearchContactsParams,
  type GetContactParams,
  type IContactContentReader,
  type ContactDetails,
  nullContactContentReader,
} from './contacts.js';

// Tasks tools
export {
  TasksTools,
  createTasksTools,
  ListTasksInput,
  SearchTasksInput,
  GetTaskInput,
  type ListTasksParams,
  type SearchTasksParams,
  type GetTaskParams,
  type ITaskContentReader,
  type TaskDetails,
  nullTaskContentReader,
} from './tasks.js';

// Notes tools
export {
  NotesTools,
  createNotesTools,
  ListNotesInput,
  GetNoteInput,
  SearchNotesInput,
  type ListNotesParams,
  type GetNoteParams,
  type SearchNotesParams,
  type INoteContentReader,
  type NoteDetails,
  nullNoteContentReader,
} from './notes.js';
