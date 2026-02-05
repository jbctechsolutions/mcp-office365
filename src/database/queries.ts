/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * SQL query definitions for Outlook database.
 *
 * All queries are read-only SELECT statements.
 */

// =============================================================================
// Folder Queries
// =============================================================================

/**
 * Lists all mail folders with message and unread counts.
 */
export const LIST_FOLDERS = `
SELECT
  f.Record_RecordID as id,
  f.Folder_Name as name,
  f.Folder_ParentID as parentId,
  f.Folder_SpecialFolderType as specialType,
  f.Folder_FolderType as folderType,
  f.Record_AccountUID as accountId,
  (SELECT COUNT(*) FROM Mail m WHERE m.Record_FolderID = f.Record_RecordID) as messageCount,
  (SELECT COUNT(*) FROM Mail m WHERE m.Record_FolderID = f.Record_RecordID AND m.Message_ReadFlag = 0) as unreadCount
FROM Folders f
WHERE f.Folder_FolderClass = 0
ORDER BY f.Folder_FolderOrder, f.Folder_Name
`;

/**
 * Gets a single folder by ID.
 */
export const GET_FOLDER = `
SELECT
  f.Record_RecordID as id,
  f.Folder_Name as name,
  f.Folder_ParentID as parentId,
  f.Folder_SpecialFolderType as specialType,
  f.Folder_FolderType as folderType,
  f.Record_AccountUID as accountId,
  (SELECT COUNT(*) FROM Mail m WHERE m.Record_FolderID = f.Record_RecordID) as messageCount,
  (SELECT COUNT(*) FROM Mail m WHERE m.Record_FolderID = f.Record_RecordID AND m.Message_ReadFlag = 0) as unreadCount
FROM Folders f
WHERE f.Record_RecordID = ?
`;

// =============================================================================
// Mail Queries
// =============================================================================

/**
 * Lists emails in a folder with pagination.
 */
export const LIST_EMAILS = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Message_NormalizedSubject as subject,
  Message_SenderList as sender,
  Message_SenderAddressList as senderAddress,
  Message_RecipientList as recipients,
  Message_DisplayTo as displayTo,
  Message_ToRecipientAddressList as toAddresses,
  Message_CCRecipientAddressList as ccAddresses,
  Message_Preview as preview,
  Message_ReadFlag as isRead,
  Message_TimeReceived as timeReceived,
  Message_TimeSent as timeSent,
  Message_HasAttachment as hasAttachment,
  Message_Size as size,
  Record_Priority as priority,
  Record_FlagStatus as flagStatus,
  PathToDataFile as dataFilePath
FROM Mail
WHERE Record_FolderID = ?
ORDER BY Message_TimeReceived DESC
LIMIT ? OFFSET ?
`;

/**
 * Lists unread emails in a folder with pagination.
 */
export const LIST_UNREAD_EMAILS = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Message_NormalizedSubject as subject,
  Message_SenderList as sender,
  Message_SenderAddressList as senderAddress,
  Message_RecipientList as recipients,
  Message_DisplayTo as displayTo,
  Message_Preview as preview,
  Message_ReadFlag as isRead,
  Message_TimeReceived as timeReceived,
  Message_TimeSent as timeSent,
  Message_HasAttachment as hasAttachment,
  Message_Size as size,
  Record_Priority as priority,
  Record_FlagStatus as flagStatus,
  PathToDataFile as dataFilePath
FROM Mail
WHERE Record_FolderID = ? AND Message_ReadFlag = 0
ORDER BY Message_TimeReceived DESC
LIMIT ? OFFSET ?
`;

/**
 * Searches emails by subject, sender, or preview.
 */
export const SEARCH_EMAILS = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Message_NormalizedSubject as subject,
  Message_SenderList as sender,
  Message_SenderAddressList as senderAddress,
  Message_RecipientList as recipients,
  Message_Preview as preview,
  Message_ReadFlag as isRead,
  Message_TimeReceived as timeReceived,
  Message_HasAttachment as hasAttachment,
  PathToDataFile as dataFilePath
FROM Mail
WHERE (
  Message_NormalizedSubject LIKE ?
  OR Message_SenderList LIKE ?
  OR Message_Preview LIKE ?
)
ORDER BY Message_TimeReceived DESC
LIMIT ?
`;

/**
 * Searches emails within a specific folder.
 */
export const SEARCH_EMAILS_IN_FOLDER = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Message_NormalizedSubject as subject,
  Message_SenderList as sender,
  Message_SenderAddressList as senderAddress,
  Message_RecipientList as recipients,
  Message_Preview as preview,
  Message_ReadFlag as isRead,
  Message_TimeReceived as timeReceived,
  Message_HasAttachment as hasAttachment,
  PathToDataFile as dataFilePath
FROM Mail
WHERE Record_FolderID = ? AND (
  Message_NormalizedSubject LIKE ?
  OR Message_SenderList LIKE ?
  OR Message_Preview LIKE ?
)
ORDER BY Message_TimeReceived DESC
LIMIT ?
`;

/**
 * Gets a single email by ID.
 */
export const GET_EMAIL = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Message_NormalizedSubject as subject,
  Message_SenderList as sender,
  Message_SenderAddressList as senderAddress,
  Message_RecipientList as recipients,
  Message_DisplayTo as displayTo,
  Message_ToRecipientAddressList as toAddresses,
  Message_CCRecipientAddressList as ccAddresses,
  Message_Preview as preview,
  Message_ReadFlag as isRead,
  Message_TimeReceived as timeReceived,
  Message_TimeSent as timeSent,
  Message_HasAttachment as hasAttachment,
  Message_Size as size,
  Record_Priority as priority,
  Record_FlagStatus as flagStatus,
  Record_Categories as categories,
  Message_MessageID as messageId,
  Conversation_ConversationID as conversationId,
  PathToDataFile as dataFilePath
FROM Mail
WHERE Record_RecordID = ?
`;

/**
 * Gets unread count for all folders or a specific folder.
 */
export const GET_UNREAD_COUNT = `
SELECT COUNT(*) as count
FROM Mail
WHERE Message_ReadFlag = 0
`;

export const GET_UNREAD_COUNT_BY_FOLDER = `
SELECT COUNT(*) as count
FROM Mail
WHERE Message_ReadFlag = 0 AND Record_FolderID = ?
`;

// =============================================================================
// Calendar Queries
// =============================================================================

/**
 * Lists calendar folders.
 */
export const LIST_CALENDARS = `
SELECT
  f.Record_RecordID as id,
  f.Folder_Name as name,
  f.Record_AccountUID as accountId
FROM Folders f
WHERE f.Folder_SpecialFolderType = 4
ORDER BY f.Folder_Name
`;

/**
 * Lists events with optional date range and folder filter.
 */
export const LIST_EVENTS = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Calendar_StartDateUTC as startDate,
  Calendar_EndDateUTC as endDate,
  Calendar_IsRecurring as isRecurring,
  Calendar_HasReminder as hasReminder,
  Calendar_AttendeeCount as attendeeCount,
  Calendar_UID as uid,
  PathToDataFile as dataFilePath
FROM CalendarEvents
ORDER BY Calendar_StartDateUTC
LIMIT ?
`;

export const LIST_EVENTS_BY_FOLDER = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Calendar_StartDateUTC as startDate,
  Calendar_EndDateUTC as endDate,
  Calendar_IsRecurring as isRecurring,
  Calendar_HasReminder as hasReminder,
  Calendar_AttendeeCount as attendeeCount,
  Calendar_UID as uid,
  PathToDataFile as dataFilePath
FROM CalendarEvents
WHERE Record_FolderID = ?
ORDER BY Calendar_StartDateUTC
LIMIT ?
`;

export const LIST_EVENTS_BY_DATE_RANGE = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Calendar_StartDateUTC as startDate,
  Calendar_EndDateUTC as endDate,
  Calendar_IsRecurring as isRecurring,
  Calendar_HasReminder as hasReminder,
  Calendar_AttendeeCount as attendeeCount,
  Calendar_UID as uid,
  PathToDataFile as dataFilePath
FROM CalendarEvents
WHERE Calendar_StartDateUTC >= ? AND Calendar_EndDateUTC <= ?
ORDER BY Calendar_StartDateUTC
LIMIT ?
`;

/**
 * Gets a single event by ID.
 */
export const GET_EVENT = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Calendar_StartDateUTC as startDate,
  Calendar_EndDateUTC as endDate,
  Calendar_IsRecurring as isRecurring,
  Calendar_HasReminder as hasReminder,
  Calendar_AttendeeCount as attendeeCount,
  Calendar_UID as uid,
  Calendar_MasterRecordID as masterRecordId,
  Calendar_RecurrenceID as recurrenceId,
  PathToDataFile as dataFilePath
FROM CalendarEvents
WHERE Record_RecordID = ?
`;

// =============================================================================
// Contact Queries
// =============================================================================

/**
 * Lists contacts with pagination.
 */
export const LIST_CONTACTS = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Contact_DisplayName as displayName,
  Contact_DisplayNameSort as sortName,
  Contact_ContactRecType as contactType,
  PathToDataFile as dataFilePath
FROM Contacts
ORDER BY Contact_DisplayNameSort
LIMIT ? OFFSET ?
`;

/**
 * Searches contacts by name.
 */
export const SEARCH_CONTACTS = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Contact_DisplayName as displayName,
  Contact_DisplayNameSort as sortName,
  PathToDataFile as dataFilePath
FROM Contacts
WHERE Contact_DisplayName LIKE ? OR Contact_DisplayNameSort LIKE ?
ORDER BY Contact_DisplayNameSort
LIMIT ?
`;

/**
 * Gets a single contact by ID.
 */
export const GET_CONTACT = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Contact_DisplayName as displayName,
  Contact_DisplayNameSort as sortName,
  Contact_ContactRecType as contactType,
  PathToDataFile as dataFilePath
FROM Contacts
WHERE Record_RecordID = ?
`;

// =============================================================================
// Task Queries
// =============================================================================

/**
 * Lists tasks with pagination.
 */
export const LIST_TASKS = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Task_Name as name,
  Task_Completed as isCompleted,
  Record_DueDate as dueDate,
  Record_StartDate as startDate,
  Record_Priority as priority,
  PathToDataFile as dataFilePath
FROM Tasks
ORDER BY Record_DueDate, Task_Name
LIMIT ? OFFSET ?
`;

/**
 * Lists incomplete tasks only.
 */
export const LIST_INCOMPLETE_TASKS = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Task_Name as name,
  Task_Completed as isCompleted,
  Record_DueDate as dueDate,
  Record_StartDate as startDate,
  Record_Priority as priority,
  PathToDataFile as dataFilePath
FROM Tasks
WHERE Task_Completed = 0
ORDER BY Record_DueDate, Task_Name
LIMIT ? OFFSET ?
`;

/**
 * Searches tasks by name.
 */
export const SEARCH_TASKS = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Task_Name as name,
  Task_Completed as isCompleted,
  Record_DueDate as dueDate,
  Record_Priority as priority,
  PathToDataFile as dataFilePath
FROM Tasks
WHERE Task_Name LIKE ?
ORDER BY Record_DueDate, Task_Name
LIMIT ?
`;

/**
 * Gets a single task by ID.
 */
export const GET_TASK = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Task_Name as name,
  Task_Completed as isCompleted,
  Record_DueDate as dueDate,
  Record_StartDate as startDate,
  Record_Priority as priority,
  Record_HasReminder as hasReminder,
  PathToDataFile as dataFilePath
FROM Tasks
WHERE Record_RecordID = ?
`;

// =============================================================================
// Note Queries
// =============================================================================

/**
 * Lists notes with pagination.
 */
export const LIST_NOTES = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Record_ModDate as modifiedDate,
  PathToDataFile as dataFilePath
FROM Notes
ORDER BY Record_ModDate DESC
LIMIT ? OFFSET ?
`;

/**
 * Gets a single note by ID.
 */
export const GET_NOTE = `
SELECT
  Record_RecordID as id,
  Record_FolderID as folderId,
  Record_ModDate as modifiedDate,
  PathToDataFile as dataFilePath
FROM Notes
WHERE Record_RecordID = ?
`;
