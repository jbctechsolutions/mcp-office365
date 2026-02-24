/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

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
  ListAttachmentsInput,
  DownloadAttachmentInput,
  type ListFoldersParams,
  type ListEmailsParams,
  type SearchEmailsParams,
  type GetEmailParams,
  type GetUnreadCountParams,
  type ListAttachmentsParams,
  type DownloadAttachmentParams,
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
  RespondToEventInput,
  type ListCalendarsParams,
  type ListEventsParams,
  type GetEventParams,
  type SearchEventsParams,
  type CreateEventParams,
  type CreateEventResult,
  type RecurrenceParams,
  type RespondToEventParams,
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

// Mailbox organization tools
export {
  MailboxOrganizationTools,
  createMailboxOrganizationTools,
  PrepareDeleteEmailInput,
  ConfirmDeleteEmailInput,
  PrepareMoveEmailInput,
  ConfirmMoveEmailInput,
  PrepareArchiveEmailInput,
  ConfirmArchiveEmailInput,
  PrepareJunkEmailInput,
  ConfirmJunkEmailInput,
  PrepareDeleteFolderInput,
  ConfirmDeleteFolderInput,
  PrepareEmptyFolderInput,
  ConfirmEmptyFolderInput,
  PrepareBatchDeleteEmailsInput,
  PrepareBatchMoveEmailsInput,
  ConfirmBatchOperationInput,
  MarkEmailReadInput,
  MarkEmailUnreadInput,
  SetEmailFlagInput,
  ClearEmailFlagInput,
  SetEmailCategoriesInput,
  CreateFolderInput,
  RenameFolderInput,
  MoveFolderInput,
} from './mailbox-organization.js';

// Mail send tools
export {
  MailSendTools,
  createMailSendTools,
  CreateDraftInput,
  UpdateDraftInput,
  ListDraftsInput,
  PrepareSendDraftInput,
  ConfirmSendDraftInput,
  PrepareSendEmailInput,
  ConfirmSendEmailInput,
  PrepareReplyEmailInput,
  ConfirmReplyEmailInput,
  PrepareForwardEmailInput,
  ConfirmForwardEmailInput,
  ReplyAsDraftInput,
  ForwardAsDraftInput,
  type IMailSendRepository,
  type CreateDraftResult,
  type CreateDraftParams,
  type UpdateDraftParams,
  type ListDraftsParams,
  type PrepareSendDraftParams,
  type ConfirmSendDraftParams,
  type PrepareSendEmailParams,
  type ConfirmSendEmailParams,
  type PrepareReplyEmailParams,
  type ConfirmReplyEmailParams,
  type PrepareForwardEmailParams,
  type ConfirmForwardEmailParams,
  type ReplyAsDraftParams,
  type ForwardAsDraftParams,
} from './mail-send.js';
