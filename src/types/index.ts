/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Type definitions for Outlook MCP Server.
 *
 * Re-exports all domain types for convenient importing.
 */

// Mail types
export {
  SpecialFolderType,
  type SpecialFolderTypeValue,
  Priority,
  type PriorityValue,
  FlagStatus,
  type FlagStatusValue,
  type Folder,
  type EmailSummary,
  type Email,
  type UnreadCount,
} from './mail.js';

// Calendar types
export {
  type CalendarFolder,
  type EventSummary,
  type Event,
  type Attendee,
  AttendeeStatus,
} from './calendar.js';

// Contact types
export {
  ContactType,
  type ContactTypeValue,
  type ContactSummary,
  type Contact,
  type ContactEmail,
  EmailType,
  type ContactPhone,
  PhoneType,
  type ContactAddress,
  AddressType,
} from './contacts.js';

// Task types
export { type TaskSummary, type Task } from './tasks.js';

// Note types
export { type NoteSummary, type Note } from './notes.js';
