/**
 * AppleScript template strings for Outlook operations.
 *
 * All scripts output data in a delimiter-based format for reliable parsing:
 * - Records are separated by {{RECORD}}
 * - Fields are separated by {{FIELD}}
 * - Field names and values are separated by {{=}}
 *
 * Example output: {{RECORD}}id{{=}}1{{FIELD}}name{{=}}Inbox{{FIELD}}unread{{=}}5{{RECORD}}...
 */

import { escapeForAppleScript } from './executor.js';

// =============================================================================
// Delimiters
// =============================================================================

export const DELIMITERS = {
  RECORD: '{{RECORD}}',
  FIELD: '{{FIELD}}',
  EQUALS: '{{=}}',
  NULL: '{{NULL}}',
} as const;

// =============================================================================
// Mail Scripts
// =============================================================================

/**
 * Lists all mail folders with their properties.
 */
export const LIST_MAIL_FOLDERS = `
tell application "Microsoft Outlook"
  set output to ""
  set allFolders to mail folders
  repeat with f in allFolders
    try
      set fId to id of f
      set fName to name of f
      set uCount to unread count of f
      set output to output & "{{RECORD}}id{{=}}" & fId & "{{FIELD}}name{{=}}" & fName & "{{FIELD}}unreadCount{{=}}" & uCount
    end try
  end repeat
  return output
end tell
`;

/**
 * Gets messages from a specific folder.
 */
export function listMessages(folderId: number, limit: number, offset: number, unreadOnly: boolean): string {
  const unreadFilter = unreadOnly ? ' whose is read is false' : '';
  const totalToFetch = limit + offset;

  return `
tell application "Microsoft Outlook"
  set output to ""
  set targetFolder to mail folder id ${folderId}
  set allMsgs to (messages of targetFolder${unreadFilter})
  set msgCount to count of allMsgs
  set startIdx to ${offset + 1}
  set endIdx to ${totalToFetch}
  if endIdx > msgCount then set endIdx to msgCount
  if startIdx > msgCount then return ""

  repeat with i from startIdx to endIdx
    try
      set m to item i of allMsgs
      set mId to id of m
      set mSubject to subject of m
      set mSender to ""
      try
        set mSender to address of sender of m
      end try
      set mSenderName to ""
      try
        set mSenderName to name of sender of m
      end try
      set mDate to ""
      try
        set mDate to time received of m as «class isot» as string
      end try
      set mRead to is read of m
      set mPriority to "normal"
      try
        set p to priority of m
        if p is priority high then
          set mPriority to "high"
        else if p is priority low then
          set mPriority to "low"
        end if
      end try
      set mPreview to ""
      try
        set mPreview to text 1 thru 200 of plain text content of m
      on error
        try
          set mPreview to plain text content of m
        end try
      end try

      set output to output & "{{RECORD}}id{{=}}" & mId & "{{FIELD}}subject{{=}}" & mSubject & "{{FIELD}}senderEmail{{=}}" & mSender & "{{FIELD}}senderName{{=}}" & mSenderName & "{{FIELD}}dateReceived{{=}}" & mDate & "{{FIELD}}isRead{{=}}" & mRead & "{{FIELD}}priority{{=}}" & mPriority & "{{FIELD}}preview{{=}}" & mPreview
    end try
  end repeat
  return output
end tell
`;
}

/**
 * Searches messages by query.
 */
export function searchMessages(query: string, folderId: number | null, limit: number): string {
  const escapedQuery = escapeForAppleScript(query);
  const folderClause = folderId != null ? `of mail folder id ${folderId}` : '';

  return `
tell application "Microsoft Outlook"
  set output to ""
  set searchResults to (messages ${folderClause} whose subject contains "${escapedQuery}" or (address of sender) contains "${escapedQuery}")
  set resultCount to count of searchResults
  set maxResults to ${limit}
  if resultCount < maxResults then set maxResults to resultCount

  repeat with i from 1 to maxResults
    try
      set m to item i of searchResults
      set mId to id of m
      set mSubject to subject of m
      set mSender to ""
      try
        set mSender to address of sender of m
      end try
      set mSenderName to ""
      try
        set mSenderName to name of sender of m
      end try
      set mDate to ""
      try
        set mDate to time received of m as «class isot» as string
      end try
      set mRead to is read of m
      set mPreview to ""
      try
        set mPreview to text 1 thru 200 of plain text content of m
      on error
        try
          set mPreview to plain text content of m
        end try
      end try

      set output to output & "{{RECORD}}id{{=}}" & mId & "{{FIELD}}subject{{=}}" & mSubject & "{{FIELD}}senderEmail{{=}}" & mSender & "{{FIELD}}senderName{{=}}" & mSenderName & "{{FIELD}}dateReceived{{=}}" & mDate & "{{FIELD}}isRead{{=}}" & mRead & "{{FIELD}}preview{{=}}" & mPreview
    end try
  end repeat
  return output
end tell
`;
}

/**
 * Gets a single message by ID with full content.
 */
export function getMessage(messageId: number): string {
  return `
tell application "Microsoft Outlook"
  set m to message id ${messageId}
  set mId to id of m
  set mSubject to subject of m
  set mSender to ""
  try
    set mSender to address of sender of m
  end try
  set mSenderName to ""
  try
    set mSenderName to name of sender of m
  end try
  set mDateReceived to ""
  try
    set mDateReceived to time received of m as «class isot» as string
  end try
  set mDateSent to ""
  try
    set mDateSent to time sent of m as «class isot» as string
  end try
  set mRead to is read of m
  set mPriority to "normal"
  try
    set p to priority of m
    if p is priority high then
      set mPriority to "high"
    else if p is priority low then
      set mPriority to "low"
    end if
  end try
  set mHtml to ""
  try
    set mHtml to content of m
  end try
  set mPlain to ""
  try
    set mPlain to plain text content of m
  end try
  set mHasHtml to has html of m
  set mFolderId to ""
  try
    set mFolderId to id of folder of m
  end try

  -- Get recipients
  set toList to ""
  try
    repeat with r in to recipients of m
      set toList to toList & (address of r) & ","
    end repeat
  end try
  set ccList to ""
  try
    repeat with r in cc recipients of m
      set ccList to ccList & (address of r) & ","
    end repeat
  end try

  -- Get attachments
  set attachList to ""
  try
    repeat with a in attachments of m
      set attachList to attachList & (name of a) & ","
    end repeat
  end try

  return "{{RECORD}}id{{=}}" & mId & "{{FIELD}}subject{{=}}" & mSubject & "{{FIELD}}senderEmail{{=}}" & mSender & "{{FIELD}}senderName{{=}}" & mSenderName & "{{FIELD}}dateReceived{{=}}" & mDateReceived & "{{FIELD}}dateSent{{=}}" & mDateSent & "{{FIELD}}isRead{{=}}" & mRead & "{{FIELD}}priority{{=}}" & mPriority & "{{FIELD}}htmlContent{{=}}" & mHtml & "{{FIELD}}plainContent{{=}}" & mPlain & "{{FIELD}}hasHtml{{=}}" & mHasHtml & "{{FIELD}}folderId{{=}}" & mFolderId & "{{FIELD}}toRecipients{{=}}" & toList & "{{FIELD}}ccRecipients{{=}}" & ccList & "{{FIELD}}attachments{{=}}" & attachList
end tell
`;
}

/**
 * Gets unread count for a folder.
 */
export function getUnreadCount(folderId: number): string {
  return `
tell application "Microsoft Outlook"
  set f to mail folder id ${folderId}
  return unread count of f
end tell
`;
}

// =============================================================================
// Calendar Scripts
// =============================================================================

/**
 * Lists all calendars.
 */
export const LIST_CALENDARS = `
tell application "Microsoft Outlook"
  set output to ""
  set allCalendars to calendars
  repeat with c in allCalendars
    try
      set cId to id of c
      set cName to name of c
      set output to output & "{{RECORD}}id{{=}}" & cId & "{{FIELD}}name{{=}}" & cName
    end try
  end repeat
  return output
end tell
`;

/**
 * Lists events from a calendar within a date range.
 */
export function listEvents(calendarId: number | null, startDate: string | null, endDate: string | null, limit: number): string {
  const calendarClause = calendarId != null ? `of calendar id ${calendarId}` : '';

  return `
tell application "Microsoft Outlook"
  set output to ""
  set allEvents to calendar events ${calendarClause}
  set eventCount to count of allEvents
  set maxEvents to ${limit}
  if eventCount < maxEvents then set maxEvents to eventCount

  repeat with i from 1 to maxEvents
    try
      set e to item i of allEvents
      set eId to id of e
      set eSubject to subject of e
      set eStart to ""
      try
        set eStart to start time of e as «class isot» as string
      end try
      set eEnd to ""
      try
        set eEnd to end time of e as «class isot» as string
      end try
      set eLocation to ""
      try
        set eLocation to location of e
      end try
      set eAllDay to all day flag of e
      set eRecurring to is recurring of e

      set output to output & "{{RECORD}}id{{=}}" & eId & "{{FIELD}}subject{{=}}" & eSubject & "{{FIELD}}startTime{{=}}" & eStart & "{{FIELD}}endTime{{=}}" & eEnd & "{{FIELD}}location{{=}}" & eLocation & "{{FIELD}}isAllDay{{=}}" & eAllDay & "{{FIELD}}isRecurring{{=}}" & eRecurring
    end try
  end repeat
  return output
end tell
`;
}

/**
 * Gets a single event by ID.
 */
export function getEvent(eventId: number): string {
  return `
tell application "Microsoft Outlook"
  set e to calendar event id ${eventId}
  set eId to id of e
  set eSubject to subject of e
  set eStart to ""
  try
    set eStart to start time of e as «class isot» as string
  end try
  set eEnd to ""
  try
    set eEnd to end time of e as «class isot» as string
  end try
  set eLocation to ""
  try
    set eLocation to location of e
  end try
  set eContent to ""
  try
    set eContent to content of e
  end try
  set ePlain to ""
  try
    set ePlain to plain text content of e
  end try
  set eAllDay to all day flag of e
  set eRecurring to is recurring of e
  set eOrganizer to ""
  try
    set eOrganizer to organizer of e
  end try
  set eCalId to ""
  try
    set eCalId to id of calendar of e
  end try

  -- Get attendees
  set attendeeList to ""
  try
    repeat with a in attendees of e
      set aEmail to email address of a
      set aName to name of a
      set attendeeList to attendeeList & aEmail & "|" & aName & ","
    end repeat
  end try

  return "{{RECORD}}id{{=}}" & eId & "{{FIELD}}subject{{=}}" & eSubject & "{{FIELD}}startTime{{=}}" & eStart & "{{FIELD}}endTime{{=}}" & eEnd & "{{FIELD}}location{{=}}" & eLocation & "{{FIELD}}htmlContent{{=}}" & eContent & "{{FIELD}}plainContent{{=}}" & ePlain & "{{FIELD}}isAllDay{{=}}" & eAllDay & "{{FIELD}}isRecurring{{=}}" & eRecurring & "{{FIELD}}organizer{{=}}" & eOrganizer & "{{FIELD}}calendarId{{=}}" & eCalId & "{{FIELD}}attendees{{=}}" & attendeeList
end tell
`;
}

/**
 * Searches events by query.
 */
export function searchEvents(query: string, limit: number): string {
  const escapedQuery = escapeForAppleScript(query);

  return `
tell application "Microsoft Outlook"
  set output to ""
  set searchResults to (calendar events whose subject contains "${escapedQuery}")
  set resultCount to count of searchResults
  set maxResults to ${limit}
  if resultCount < maxResults then set maxResults to resultCount

  repeat with i from 1 to maxResults
    try
      set e to item i of searchResults
      set eId to id of e
      set eSubject to subject of e
      set eStart to ""
      try
        set eStart to start time of e as «class isot» as string
      end try
      set eEnd to ""
      try
        set eEnd to end time of e as «class isot» as string
      end try
      set eLocation to ""
      try
        set eLocation to location of e
      end try

      set output to output & "{{RECORD}}id{{=}}" & eId & "{{FIELD}}subject{{=}}" & eSubject & "{{FIELD}}startTime{{=}}" & eStart & "{{FIELD}}endTime{{=}}" & eEnd & "{{FIELD}}location{{=}}" & eLocation
    end try
  end repeat
  return output
end tell
`;
}

/**
 * Recurrence parameters for AppleScript generation.
 */
export interface RecurrenceScriptParams {
  readonly frequency: 'daily' | 'weekly' | 'monthly' | 'yearly';
  readonly interval: number;
  readonly daysOfWeek?: readonly string[];
  readonly dayOfMonth?: number;
  readonly weekOfMonth?: string;
  readonly dayOfWeekMonthly?: string;
  readonly endDate?: {
    readonly year: number;
    readonly month: number;
    readonly day: number;
    readonly hours: number;
    readonly minutes: number;
  };
  readonly endAfterCount?: number;
}

/**
 * Builds AppleScript to set recurrence on a newly created event.
 * Assumes `newEvent` variable is already in scope from createEvent.
 */
function buildRecurrenceScript(params: RecurrenceScriptParams): string {
  const isOrdinalMonthly = params.frequency === 'monthly' && params.weekOfMonth != null;
  const frequencyMap: Record<string, string> = {
    daily: 'daily recurrence',
    weekly: 'weekly recurrence',
    monthly: isOrdinalMonthly ? 'month nth recurrence' : 'monthly recurrence',
    yearly: 'yearly recurrence',
  };

  const recurrenceType = frequencyMap[params.frequency]!;
  const capitalize = (s: string): string => s.charAt(0).toUpperCase() + s.slice(1);

  let script = `
  set is recurring of newEvent to true
  set theRecurrence to recurrence of newEvent
  set recurrence type of theRecurrence to ${recurrenceType}
  set occurrence interval of theRecurrence to ${params.interval}`;

  // Weekly: days of week mask
  if (params.frequency === 'weekly' && params.daysOfWeek != null) {
    const daysList = params.daysOfWeek.map(capitalize).join(', ');
    script += `\n  set day of week mask of theRecurrence to {${daysList}}`;
  }

  // Monthly by date
  if (params.frequency === 'monthly' && params.dayOfMonth != null && params.weekOfMonth == null) {
    script += `\n  set day of month of theRecurrence to ${params.dayOfMonth}`;
  }

  // Monthly ordinal (e.g., 3rd Thursday)
  if (params.frequency === 'monthly' && params.weekOfMonth != null && params.dayOfWeekMonthly != null) {
    const instanceMap: Record<string, number> = { first: 1, second: 2, third: 3, fourth: 4, last: 5 };
    const instance = instanceMap[params.weekOfMonth] ?? 1;
    script += `\n  set day of week mask of theRecurrence to {${capitalize(params.dayOfWeekMonthly)}}`;
    script += `\n  set instance of theRecurrence to ${instance}`;
  }

  // End after count
  if (params.endAfterCount != null) {
    script += `\n  set occurrences of theRecurrence to ${params.endAfterCount}`;
  }

  // End by date (component-based for locale safety)
  if (params.endDate != null) {
    script += `
  set theEndRecurrenceDate to current date
  set day of theEndRecurrenceDate to 1
  set year of theEndRecurrenceDate to ${params.endDate.year}
  set month of theEndRecurrenceDate to ${params.endDate.month}
  set day of theEndRecurrenceDate to ${params.endDate.day}
  set hours of theEndRecurrenceDate to ${params.endDate.hours}
  set minutes of theEndRecurrenceDate to ${params.endDate.minutes}
  set seconds of theEndRecurrenceDate to 0
  set pattern end date of theRecurrence to theEndRecurrenceDate`;
  }

  return script;
}

/**
 * Parameters for responding to an event invitation (RSVP).
 */
export interface RespondToEventParams {
  readonly eventId: number;
  readonly response: 'accept' | 'decline' | 'tentative';
  readonly sendResponse: boolean;
  readonly comment?: string;
}

/**
 * Responds to an event invitation (RSVP).
 */
export function respondToEvent(params: RespondToEventParams): string {
  const { eventId, response, sendResponse, comment } = params;

  // Map response to AppleScript status value
  const statusMap = {
    accept: 'accept',
    decline: 'decline',
    tentative: 'tentative accept',
  };
  const status = statusMap[response];

  // Escape comment if provided
  const commentLine = comment != null
    ? `set comment of myEvent to "${escapeForAppleScript(comment)}"`
    : '';

  return `
tell application "Microsoft Outlook"
  try
    set myEvent to calendar event id ${eventId}
    set response status of myEvent to ${status}
    ${commentLine}

    -- Return success
    set output to "{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}" & ${eventId}
    return output
  on error errMsg
    -- Return failure
    set output to "{{RECORD}}success{{=}}false{{FIELD}}error{{=}}" & errMsg
    return output
  end try
end tell
`;
}

/**
 * Parameters for deleting an event.
 */
export interface DeleteEventParams {
  readonly eventId: number;
  readonly applyTo: 'this_instance' | 'all_in_series';
}

/**
 * Deletes an event. For recurring events, can delete single instance or entire series.
 */
export function deleteEvent(params: DeleteEventParams): string {
  const { eventId, applyTo } = params;

  const comment = applyTo === 'all_in_series'
    ? '-- Deleting entire series'
    : '-- Deleting single instance';

  return `
tell application "Microsoft Outlook"
  try
    ${comment}
    set myEvent to calendar event id ${eventId}
    delete myEvent

    -- Return success
    set output to "{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}" & ${eventId}
    return output
  on error errMsg
    -- Return failure
    set output to "{{RECORD}}success{{=}}false{{FIELD}}error{{=}}" & errMsg
    return output
  end try
end tell
`;
}

/**
 * Parameters for updating an event.
 */
export interface UpdateEventParams {
  readonly eventId: number;
  readonly applyTo: 'this_instance' | 'all_in_series';
  readonly updates: {
    readonly title?: string;
    readonly startDate?: string; // ISO 8601
    readonly endDate?: string; // ISO 8601
    readonly location?: string;
    readonly description?: string;
    readonly isAllDay?: boolean;
  };
}

/**
 * Converts an ISO 8601 date string into individual UTC date components.
 */
function isoToDateComponents(isoString: string): {
  year: number;
  month: number;
  day: number;
  hours: number;
  minutes: number;
} {
  const date = new Date(isoString);
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    hours: date.getUTCHours(),
    minutes: date.getUTCMinutes(),
  };
}

/**
 * Updates an event. For recurring events, can update single instance or entire series.
 * All update fields are optional - only specified fields will be updated.
 */
export function updateEvent(params: UpdateEventParams): string {
  const { eventId, applyTo, updates } = params;
  const updatedFields: string[] = [];

  let updateStatements = '';

  if (updates.title != null) {
    updateStatements += `    set subject of myEvent to "${escapeForAppleScript(updates.title)}"\n`;
    updatedFields.push('title');
  }

  if (updates.location != null) {
    updateStatements += `    set location of myEvent to "${escapeForAppleScript(updates.location)}"\n`;
    updatedFields.push('location');
  }

  if (updates.description != null) {
    updateStatements += `    set content of myEvent to "${escapeForAppleScript(updates.description)}"\n`;
    updatedFields.push('description');
  }

  if (updates.startDate != null) {
    const start = isoToDateComponents(updates.startDate);
    updateStatements += `    set start time of myEvent to date "${start.year}-${start.month}-${start.day} ${start.hours}:${start.minutes}:00"\n`;
    updatedFields.push('startDate');
  }

  if (updates.endDate != null) {
    const end = isoToDateComponents(updates.endDate);
    updateStatements += `    set end time of myEvent to date "${end.year}-${end.month}-${end.day} ${end.hours}:${end.minutes}:00"\n`;
    updatedFields.push('endDate');
  }

  if (updates.isAllDay != null) {
    updateStatements += `    set all day flag of myEvent to ${updates.isAllDay}\n`;
    updatedFields.push('isAllDay');
  }

  const comment = applyTo === 'all_in_series'
    ? '-- Updating entire series'
    : '-- Updating single instance';

  const fieldsOutput = updatedFields.join(',');

  return `
tell application "Microsoft Outlook"
  try
    ${comment}
    set myEvent to calendar event id ${eventId}

${updateStatements}
    -- Return success
    set output to "{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}" & ${eventId} & "{{FIELD}}updatedFields{{=}}${fieldsOutput}"
    return output
  on error errMsg
    -- Return failure
    set output to "{{RECORD}}success{{=}}false{{FIELD}}error{{=}}" & errMsg
    return output
  end try
end tell
`;
}

/**
 * Creates a new calendar event.
 * Uses component-based date construction for locale safety.
 */
export function createEvent(params: {
  title: string;
  startYear: number;
  startMonth: number;
  startDay: number;
  startHours: number;
  startMinutes: number;
  endYear: number;
  endMonth: number;
  endDay: number;
  endHours: number;
  endMinutes: number;
  calendarId?: number;
  location?: string;
  description?: string;
  isAllDay?: boolean;
  recurrence?: RecurrenceScriptParams;
}): string {
  const escapedTitle = escapeForAppleScript(params.title);
  const escapedLocation = params.location ? escapeForAppleScript(params.location) : '';
  const escapedDescription = params.description ? escapeForAppleScript(params.description) : '';

  // Build properties list
  let properties = `subject:"${escapedTitle}", start time:theStartDate, end time:theEndDate`;
  if (params.location) {
    properties += `, location:"${escapedLocation}"`;
  }
  if (params.isAllDay) {
    properties += ', all day flag:true';
  }

  // Build the target clause for calendar
  const targetClause = params.calendarId != null
    ? `at calendar id ${params.calendarId} `
    : '';

  return `
tell application "Microsoft Outlook"
  set theStartDate to current date
  set day of theStartDate to 1
  set year of theStartDate to ${params.startYear}
  set month of theStartDate to ${params.startMonth}
  set day of theStartDate to ${params.startDay}
  set hours of theStartDate to ${params.startHours}
  set minutes of theStartDate to ${params.startMinutes}
  set seconds of theStartDate to 0

  set theEndDate to current date
  set day of theEndDate to 1
  set year of theEndDate to ${params.endYear}
  set month of theEndDate to ${params.endMonth}
  set day of theEndDate to ${params.endDay}
  set hours of theEndDate to ${params.endHours}
  set minutes of theEndDate to ${params.endMinutes}
  set seconds of theEndDate to 0

  set newEvent to make new calendar event ${targetClause}with properties {${properties}}
  ${escapedDescription ? `set plain text content of newEvent to "${escapedDescription}"` : ''}${params.recurrence != null ? buildRecurrenceScript(params.recurrence) : ''}
  set eId to id of newEvent
  set eCalId to ""
  try
    set eCalId to id of calendar of newEvent
  end try
  return "{{RECORD}}id{{=}}" & eId & "{{FIELD}}calendarId{{=}}" & eCalId
end tell
`;
}

// =============================================================================
// Contact Scripts
// =============================================================================

/**
 * Lists all contacts.
 */
export function listContacts(limit: number, offset: number): string {
  const totalToFetch = limit + offset;

  return `
tell application "Microsoft Outlook"
  set output to ""
  set allContacts to contacts
  set contactCount to count of allContacts
  set startIdx to ${offset + 1}
  set endIdx to ${totalToFetch}
  if endIdx > contactCount then set endIdx to contactCount
  if startIdx > contactCount then return ""

  repeat with i from startIdx to endIdx
    try
      set c to item i of allContacts
      set cId to id of c
      set cDisplay to display name of c
      set cFirst to ""
      try
        set cFirst to first name of c
      end try
      set cLast to ""
      try
        set cLast to last name of c
      end try
      set cCompany to ""
      try
        set cCompany to company of c
      end try
      set cEmail to ""
      try
        set emailAddrs to email addresses of c
        if (count of emailAddrs) > 0 then
          set cEmail to address of item 1 of emailAddrs
        end if
      end try

      set output to output & "{{RECORD}}id{{=}}" & cId & "{{FIELD}}displayName{{=}}" & cDisplay & "{{FIELD}}firstName{{=}}" & cFirst & "{{FIELD}}lastName{{=}}" & cLast & "{{FIELD}}company{{=}}" & cCompany & "{{FIELD}}email{{=}}" & cEmail
    end try
  end repeat
  return output
end tell
`;
}

/**
 * Searches contacts by name.
 */
export function searchContacts(query: string, limit: number): string {
  const escapedQuery = escapeForAppleScript(query);

  return `
tell application "Microsoft Outlook"
  set output to ""
  set searchResults to (contacts whose display name contains "${escapedQuery}")
  set resultCount to count of searchResults
  set maxResults to ${limit}
  if resultCount < maxResults then set maxResults to resultCount

  repeat with i from 1 to maxResults
    try
      set c to item i of searchResults
      set cId to id of c
      set cDisplay to display name of c
      set cFirst to ""
      try
        set cFirst to first name of c
      end try
      set cLast to ""
      try
        set cLast to last name of c
      end try
      set cCompany to ""
      try
        set cCompany to company of c
      end try
      set cEmail to ""
      try
        set emailAddrs to email addresses of c
        if (count of emailAddrs) > 0 then
          set cEmail to address of item 1 of emailAddrs
        end if
      end try

      set output to output & "{{RECORD}}id{{=}}" & cId & "{{FIELD}}displayName{{=}}" & cDisplay & "{{FIELD}}firstName{{=}}" & cFirst & "{{FIELD}}lastName{{=}}" & cLast & "{{FIELD}}company{{=}}" & cCompany & "{{FIELD}}email{{=}}" & cEmail
    end try
  end repeat
  return output
end tell
`;
}

/**
 * Gets a single contact by ID with full details.
 */
export function getContact(contactId: number): string {
  return `
tell application "Microsoft Outlook"
  set c to contact id ${contactId}
  set cId to id of c
  set cDisplay to display name of c
  set cFirst to ""
  try
    set cFirst to first name of c
  end try
  set cLast to ""
  try
    set cLast to last name of c
  end try
  set cMiddle to ""
  try
    set cMiddle to middle name of c
  end try
  set cNickname to ""
  try
    set cNickname to nickname of c
  end try
  set cCompany to ""
  try
    set cCompany to company of c
  end try
  set cTitle to ""
  try
    set cTitle to job title of c
  end try
  set cDept to ""
  try
    set cDept to department of c
  end try
  set cNotes to ""
  try
    set cNotes to description of c
  end try

  -- Phones
  set cHomePhone to ""
  try
    set cHomePhone to home phone number of c
  end try
  set cWorkPhone to ""
  try
    set cWorkPhone to business phone number of c
  end try
  set cMobile to ""
  try
    set cMobile to mobile number of c
  end try

  -- Emails
  set emailList to ""
  try
    repeat with e in email addresses of c
      set emailList to emailList & (address of e) & ","
    end repeat
  end try

  -- Address
  set cHomeStreet to ""
  try
    set cHomeStreet to home street address of c
  end try
  set cHomeCity to ""
  try
    set cHomeCity to home city of c
  end try
  set cHomeState to ""
  try
    set cHomeState to home state of c
  end try
  set cHomeZip to ""
  try
    set cHomeZip to home zip of c
  end try
  set cHomeCountry to ""
  try
    set cHomeCountry to home country of c
  end try

  return "{{RECORD}}id{{=}}" & cId & "{{FIELD}}displayName{{=}}" & cDisplay & "{{FIELD}}firstName{{=}}" & cFirst & "{{FIELD}}lastName{{=}}" & cLast & "{{FIELD}}middleName{{=}}" & cMiddle & "{{FIELD}}nickname{{=}}" & cNickname & "{{FIELD}}company{{=}}" & cCompany & "{{FIELD}}jobTitle{{=}}" & cTitle & "{{FIELD}}department{{=}}" & cDept & "{{FIELD}}notes{{=}}" & cNotes & "{{FIELD}}homePhone{{=}}" & cHomePhone & "{{FIELD}}workPhone{{=}}" & cWorkPhone & "{{FIELD}}mobilePhone{{=}}" & cMobile & "{{FIELD}}emails{{=}}" & emailList & "{{FIELD}}homeStreet{{=}}" & cHomeStreet & "{{FIELD}}homeCity{{=}}" & cHomeCity & "{{FIELD}}homeState{{=}}" & cHomeState & "{{FIELD}}homeZip{{=}}" & cHomeZip & "{{FIELD}}homeCountry{{=}}" & cHomeCountry
end tell
`;
}

// =============================================================================
// Task Scripts
// =============================================================================

/**
 * Lists all tasks.
 */
export function listTasks(limit: number, offset: number, includeCompleted: boolean): string {
  const totalToFetch = limit + offset;
  const completedFilter = includeCompleted ? '' : ' whose is completed is false';

  return `
tell application "Microsoft Outlook"
  set output to ""
  set allTasks to (tasks${completedFilter})
  set taskCount to count of allTasks
  set startIdx to ${offset + 1}
  set endIdx to ${totalToFetch}
  if endIdx > taskCount then set endIdx to taskCount
  if startIdx > taskCount then return ""

  repeat with i from startIdx to endIdx
    try
      set t to item i of allTasks
      set tId to id of t
      set tName to name of t
      set tDue to ""
      try
        set tDue to due date of t as «class isot» as string
      end try
      set tCompleted to is completed of t
      set tPriority to "normal"
      try
        set p to priority of t
        if p is priority high then
          set tPriority to "high"
        else if p is priority low then
          set tPriority to "low"
        end if
      end try

      set output to output & "{{RECORD}}id{{=}}" & tId & "{{FIELD}}name{{=}}" & tName & "{{FIELD}}dueDate{{=}}" & tDue & "{{FIELD}}isCompleted{{=}}" & tCompleted & "{{FIELD}}priority{{=}}" & tPriority
    end try
  end repeat
  return output
end tell
`;
}

/**
 * Searches tasks by name.
 */
export function searchTasks(query: string, limit: number): string {
  const escapedQuery = escapeForAppleScript(query);

  return `
tell application "Microsoft Outlook"
  set output to ""
  set searchResults to (tasks whose name contains "${escapedQuery}")
  set resultCount to count of searchResults
  set maxResults to ${limit}
  if resultCount < maxResults then set maxResults to resultCount

  repeat with i from 1 to maxResults
    try
      set t to item i of searchResults
      set tId to id of t
      set tName to name of t
      set tDue to ""
      try
        set tDue to due date of t as «class isot» as string
      end try
      set tCompleted to is completed of t

      set output to output & "{{RECORD}}id{{=}}" & tId & "{{FIELD}}name{{=}}" & tName & "{{FIELD}}dueDate{{=}}" & tDue & "{{FIELD}}isCompleted{{=}}" & tCompleted
    end try
  end repeat
  return output
end tell
`;
}

/**
 * Gets a single task by ID.
 */
export function getTask(taskId: number): string {
  return `
tell application "Microsoft Outlook"
  set t to task id ${taskId}
  set tId to id of t
  set tName to name of t
  set tContent to ""
  try
    set tContent to content of t
  end try
  set tPlain to ""
  try
    set tPlain to plain text content of t
  end try
  set tDue to ""
  try
    set tDue to due date of t as «class isot» as string
  end try
  set tStart to ""
  try
    set tStart to start date of t as «class isot» as string
  end try
  set tCompletedDate to ""
  try
    set tCompletedDate to completed date of t as «class isot» as string
  end try
  set tCompleted to is completed of t
  set tPriority to "normal"
  try
    set p to priority of t
    if p is priority high then
      set tPriority to "high"
    else if p is priority low then
      set tPriority to "low"
    end if
  end try
  set tFolderId to ""
  try
    set tFolderId to id of folder of t
  end try

  return "{{RECORD}}id{{=}}" & tId & "{{FIELD}}name{{=}}" & tName & "{{FIELD}}htmlContent{{=}}" & tContent & "{{FIELD}}plainContent{{=}}" & tPlain & "{{FIELD}}dueDate{{=}}" & tDue & "{{FIELD}}startDate{{=}}" & tStart & "{{FIELD}}completedDate{{=}}" & tCompletedDate & "{{FIELD}}isCompleted{{=}}" & tCompleted & "{{FIELD}}priority{{=}}" & tPriority & "{{FIELD}}folderId{{=}}" & tFolderId
end tell
`;
}

// =============================================================================
// Note Scripts
// =============================================================================

/**
 * Lists all notes.
 */
export function listNotes(limit: number, offset: number): string {
  const totalToFetch = limit + offset;

  return `
tell application "Microsoft Outlook"
  set output to ""
  set allNotes to notes
  set noteCount to count of allNotes
  set startIdx to ${offset + 1}
  set endIdx to ${totalToFetch}
  if endIdx > noteCount then set endIdx to noteCount
  if startIdx > noteCount then return ""

  repeat with i from startIdx to endIdx
    try
      set n to item i of allNotes
      set nId to id of n
      set nName to name of n
      set nCreated to ""
      try
        set nCreated to creation date of n as «class isot» as string
      end try
      set nModified to ""
      try
        set nModified to modification date of n as «class isot» as string
      end try
      set nPreview to ""
      try
        set nPreview to text 1 thru 200 of plain text content of n
      on error
        try
          set nPreview to plain text content of n
        end try
      end try

      set output to output & "{{RECORD}}id{{=}}" & nId & "{{FIELD}}name{{=}}" & nName & "{{FIELD}}createdDate{{=}}" & nCreated & "{{FIELD}}modifiedDate{{=}}" & nModified & "{{FIELD}}preview{{=}}" & nPreview
    end try
  end repeat
  return output
end tell
`;
}

/**
 * Searches notes by name.
 */
export function searchNotes(query: string, limit: number): string {
  const escapedQuery = escapeForAppleScript(query);

  return `
tell application "Microsoft Outlook"
  set output to ""
  set searchResults to (notes whose name contains "${escapedQuery}")
  set resultCount to count of searchResults
  set maxResults to ${limit}
  if resultCount < maxResults then set maxResults to resultCount

  repeat with i from 1 to maxResults
    try
      set n to item i of searchResults
      set nId to id of n
      set nName to name of n
      set nCreated to ""
      try
        set nCreated to creation date of n as «class isot» as string
      end try
      set nPreview to ""
      try
        set nPreview to text 1 thru 200 of plain text content of n
      on error
        try
          set nPreview to plain text content of n
        end try
      end try

      set output to output & "{{RECORD}}id{{=}}" & nId & "{{FIELD}}name{{=}}" & nName & "{{FIELD}}createdDate{{=}}" & nCreated & "{{FIELD}}preview{{=}}" & nPreview
    end try
  end repeat
  return output
end tell
`;
}

/**
 * Gets a single note by ID.
 */
export function getNote(noteId: number): string {
  return `
tell application "Microsoft Outlook"
  set n to note id ${noteId}
  set nId to id of n
  set nName to name of n
  set nContent to ""
  try
    set nContent to content of n
  end try
  set nPlain to ""
  try
    set nPlain to plain text content of n
  end try
  set nCreated to ""
  try
    set nCreated to creation date of n as «class isot» as string
  end try
  set nModified to ""
  try
    set nModified to modification date of n as «class isot» as string
  end try
  set nFolderId to ""
  try
    set nFolderId to id of folder of n
  end try

  return "{{RECORD}}id{{=}}" & nId & "{{FIELD}}name{{=}}" & nName & "{{FIELD}}htmlContent{{=}}" & nContent & "{{FIELD}}plainContent{{=}}" & nPlain & "{{FIELD}}createdDate{{=}}" & nCreated & "{{FIELD}}modifiedDate{{=}}" & nModified & "{{FIELD}}folderId{{=}}" & nFolderId
end tell
`;
}

// =============================================================================
// Email Sending Scripts
// =============================================================================

export interface SendEmailParams {
  readonly to: readonly string[];
  readonly subject: string;
  readonly body: string;
  readonly bodyType: 'plain' | 'html';
  readonly cc?: readonly string[];
  readonly bcc?: readonly string[];
  readonly replyTo?: string;
  readonly attachments?: readonly { path: string; name?: string }[];
  readonly accountId?: number;
}

/**
 * Sends an email with optional CC, BCC, attachments, and account selection.
 */
export function sendEmail(params: SendEmailParams): string {
  const { to, subject, body, bodyType, cc, bcc, replyTo, attachments, accountId } = params;

  const escapedSubject = escapeForAppleScript(subject);
  const escapedBody = escapeForAppleScript(body);

  const toRecipients = to.map(email =>
    `    make new recipient at newMessage with properties {email address:{address:"${email}"}}`
  ).join('\n');

  const ccRecipients = cc != null && cc.length > 0
    ? cc.map(email =>
        `    make new recipient at newMessage with properties {email address:{address:"${email}"}, recipient type:recipient cc}`
      ).join('\n')
    : '';

  const bccRecipients = bcc != null && bcc.length > 0
    ? bcc.map(email =>
        `    make new recipient at newMessage with properties {email address:{address:"${email}"}, recipient type:recipient bcc}`
      ).join('\n')
    : '';

  const contentProperty = bodyType === 'html'
    ? `html content:"${escapedBody}"`
    : `plain text content:"${escapedBody}"`;

  const replyToStatement = replyTo != null
    ? `    set reply to of newMessage to "${replyTo}"`
    : '';

  const attachmentStatements = attachments != null && attachments.length > 0
    ? attachments.map(att =>
        `    make new attachment at newMessage with properties {file:(POSIX file "${att.path}")}`
      ).join('\n')
    : '';

  const accountStatement = accountId != null
    ? `    set sending account of newMessage to account id ${accountId}`
    : '';

  return `
tell application "Microsoft Outlook"
  try
    set newMessage to make new outgoing message with properties {subject:"${escapedSubject}", ${contentProperty}}

${toRecipients}
${ccRecipients}
${bccRecipients}
${replyToStatement}
${accountStatement}
${attachmentStatements}

    send newMessage

    -- Get message ID and timestamp
    set msgId to id of newMessage as string
    set sentTime to current date
    set sentISO to sentTime as «class isot» as string

    -- Return success
    set output to "{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}" & msgId & "{{FIELD}}sentAt{{=}}" & sentISO
    return output
  on error errMsg
    -- Return failure
    set output to "{{RECORD}}success{{=}}false{{FIELD}}error{{=}}" & errMsg
    return output
  end try
end tell
`;
}
