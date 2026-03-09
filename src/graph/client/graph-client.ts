/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Microsoft Graph API client wrapper.
 *
 * Provides a typed interface to the Graph API with:
 * - Automatic token management
 * - Response caching
 * - Pagination support
 * - Error handling
 */

import 'isomorphic-fetch';
import { Client, type PageCollection } from '@microsoft/microsoft-graph-client';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { getAccessToken, type DeviceCodeCallback } from '../auth/index.js';
import { ResponseCache, CacheTTL, createCacheKey } from './cache.js';

/**
 * Graph client wrapper with caching and token management.
 */
export class GraphClient {
  private client: Client | null = null;
  private readonly cache = new ResponseCache();
  private readonly deviceCodeCallback: DeviceCodeCallback | undefined;

  constructor(deviceCodeCallback?: DeviceCodeCallback) {
    this.deviceCodeCallback = deviceCodeCallback;
  }

  /**
   * Gets or creates the Graph client instance.
   */
  // eslint-disable-next-line @typescript-eslint/require-await
  private async getClient(): Promise<Client> {
    if (this.client == null) {
      this.client = Client.init({
        // eslint-disable-next-line @typescript-eslint/no-misused-promises
        authProvider: async (done) => {
          try {
            const token = await getAccessToken(this.deviceCodeCallback);
            done(null, token);
          } catch (error) {
            done(error as Error, null);
          }
        },
      });
    }
    return this.client;
  }

  /**
   * Clears the response cache.
   */
  clearCache(): void {
    this.cache.clear();
  }

  // ===========================================================================
  // Mail Folders
  // ===========================================================================

  /**
   * Lists all mail folders.
   */
  async listMailFolders(): Promise<MicrosoftGraph.MailFolder[]> {
    const cacheKey = createCacheKey('listMailFolders');
    const cached = this.cache.get<MicrosoftGraph.MailFolder[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();
    const result: MicrosoftGraph.MailFolder[] = [];

    // Get top-level folders with pagination
    let response = await client
      .api('/me/mailFolders')
      .select('id,displayName,parentFolderId,totalItemCount,unreadItemCount')
      .top(100)
      .get() as PageCollection;

    result.push(...(response.value as MicrosoftGraph.MailFolder[]));

    // Handle pagination
    while (response['@odata.nextLink'] != null) {
      response = await client.api(response['@odata.nextLink']).get() as PageCollection;
      result.push(...(response.value as MicrosoftGraph.MailFolder[]));
    }

    // Also get child folders (one level deep)
    for (const folder of [...result]) {
      try {
        const children = await client
          .api(`/me/mailFolders/${folder.id}/childFolders`)
          .select('id,displayName,parentFolderId,totalItemCount,unreadItemCount')
          .get() as PageCollection;

        result.push(...(children.value as MicrosoftGraph.MailFolder[]));
      } catch {
        // Some folders may not have children or may not be accessible
      }
    }

    this.cache.set(cacheKey, result, CacheTTL.FOLDERS);
    return result;
  }

  /**
   * Gets a specific mail folder by ID.
   */
  async getMailFolder(folderId: string): Promise<MicrosoftGraph.MailFolder | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/mailFolders/${folderId}`)
        .select('id,displayName,parentFolderId,totalItemCount,unreadItemCount')
        .get() as MicrosoftGraph.MailFolder;
    } catch {
      return null;
    }
  }

  // ===========================================================================
  // Messages (Emails)
  // ===========================================================================

  /**
   * Lists messages in a folder with pagination.
   */
  async listMessages(
    folderId: string,
    limit: number = 50,
    skip: number = 0
  ): Promise<MicrosoftGraph.Message[]> {
    const cacheKey = createCacheKey('listMessages', folderId, limit, skip);
    const cached = this.cache.get<MicrosoftGraph.Message[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();

    const response = await client
      .api(`/me/mailFolders/${folderId}/messages`)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId')
      .orderby('receivedDateTime desc')
      .top(limit)
      .skip(skip)
      .get() as PageCollection;

    const result = response.value as MicrosoftGraph.Message[];
    this.cache.set(cacheKey, result, CacheTTL.EMAILS);
    return result;
  }

  /**
   * Lists unread messages in a folder.
   */
  async listUnreadMessages(
    folderId: string,
    limit: number = 50,
    skip: number = 0
  ): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();

    const response = await client
      .api(`/me/mailFolders/${folderId}/messages`)
      .filter('isRead eq false')
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId')
      .orderby('receivedDateTime desc')
      .top(limit)
      .skip(skip)
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Message[];
  }

  /**
   * Searches messages across all folders.
   */
  async searchMessages(query: string, limit: number = 50): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();

    const response = await client
      .api('/me/messages')
      .search(`"${query}"`)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId,parentFolderId')
      .top(limit)
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Message[];
  }

  /**
   * Searches messages in a specific folder.
   */
  async searchMessagesInFolder(
    folderId: string,
    query: string,
    limit: number = 50
  ): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();

    const response = await client
      .api(`/me/mailFolders/${folderId}/messages`)
      .search(`"${query}"`)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId')
      .top(limit)
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Message[];
  }

  /**
   * Searches messages using raw KQL (Keyword Query Language).
   * Unlike searchMessages, the query is passed directly without quote-wrapping,
   * enabling KQL operators like from:, subject:, hasAttachments:, received>=, AND, OR.
   */
  async searchMessagesKql(query: string, limit: number = 50): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();
    const response = await client
      .api('/me/messages')
      .search(query)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId,parentFolderId')
      .top(limit)
      .get() as PageCollection;
    return response.value as MicrosoftGraph.Message[];
  }

  /**
   * Searches messages in a specific folder using raw KQL.
   */
  async searchMessagesKqlInFolder(
    folderId: string,
    query: string,
    limit: number = 50
  ): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();
    const response = await client
      .api(`/me/mailFolders/${folderId}/messages`)
      .search(query)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId,parentFolderId')
      .top(limit)
      .get() as PageCollection;
    return response.value as MicrosoftGraph.Message[];
  }

  /**
   * Lists messages in a conversation by conversationId.
   */
  async listConversationMessages(
    conversationId: string,
    limit: number = 50
  ): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();
    const response = await client
      .api('/me/messages')
      .filter(`conversationId eq '${conversationId}'`)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId,parentFolderId')
      .orderby('receivedDateTime asc')
      .top(limit)
      .get() as PageCollection;
    return response.value as MicrosoftGraph.Message[];
  }

  /**
   * Gets a specific message with full body.
   */
  async getMessage(messageId: string): Promise<MicrosoftGraph.Message | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/messages/${messageId}`)
        .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,body,bodyPreview,conversationId,internetMessageId,parentFolderId')
        .get() as MicrosoftGraph.Message;
    } catch {
      return null;
    }
  }

  /**
   * Gets message delta for incremental sync.
   */
  async getMessagesDelta(
    folderId: string,
    deltaLink?: string
  ): Promise<{ messages: MicrosoftGraph.Message[]; deltaLink: string }> {
    const client = await this.getClient();
    let response;

    if (deltaLink != null) {
      response = await client.api(deltaLink).get() as PageCollection;
    } else {
      response = await client
        .api(`/me/mailFolders/${folderId}/messages/delta`)
        .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId,parentFolderId')
        .top(50)
        .get() as PageCollection;
    }

    const messages: MicrosoftGraph.Message[] = [...(response.value ?? [])];

    let nextLink = response['@odata.nextLink'];
    while (nextLink != null) {
      const nextPage = await client.api(nextLink).get() as PageCollection;
      messages.push(...(nextPage.value ?? []));
      nextLink = nextPage['@odata.nextLink'];
    }

    const newDeltaLink = response['@odata.deltaLink'] ?? '';
    return { messages, deltaLink: newDeltaLink };
  }

  // ===========================================================================
  // Mail Rules
  // ===========================================================================

  /**
   * Lists all inbox mail rules.
   */
  async listMailRules(): Promise<MicrosoftGraph.MessageRule[]> {
    const client = await this.getClient();
    const response = await client
      .api('/me/mailFolders/inbox/messageRules')
      .get() as PageCollection;
    return response.value as MicrosoftGraph.MessageRule[];
  }

  /**
   * Creates a new inbox mail rule.
   */
  async createMailRule(rule: Record<string, unknown>): Promise<MicrosoftGraph.MessageRule> {
    const client = await this.getClient();
    const result = await client
      .api('/me/mailFolders/inbox/messageRules')
      .post(rule) as MicrosoftGraph.MessageRule;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes an inbox mail rule.
   */
  async deleteMailRule(ruleId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/mailFolders/inbox/messageRules/${ruleId}`)
      .delete();
    this.cache.clear();
  }

  // ===========================================================================
  // Automatic Replies (Out of Office)
  // ===========================================================================

  /**
   * Gets the automatic replies (OOF) settings.
   */
  async getAutomaticReplies(): Promise<Record<string, unknown>> {
    const client = await this.getClient();
    return await client.api('/me/mailboxSettings/automaticRepliesSetting').get() as Record<string, unknown>;
  }

  /**
   * Sets the automatic replies (OOF) settings.
   */
  async setAutomaticReplies(settings: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client.api('/me/mailboxSettings').patch({ automaticRepliesSetting: settings });
  }

  // ===========================================================================
  // Mailbox Settings
  // ===========================================================================

  /**
   * Gets the full mailbox settings for the current user.
   */
  async getMailboxSettings(): Promise<Record<string, unknown>> {
    const client = await this.getClient();
    return await client.api('/me/mailboxSettings').get() as Record<string, unknown>;
  }

  /**
   * Updates mailbox settings for the current user.
   */
  async updateMailboxSettings(settings: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client.api('/me/mailboxSettings').patch(settings);
  }

  // ===========================================================================
  // Calendars
  // ===========================================================================

  /**
   * Lists all calendars.
   */
  async listCalendars(): Promise<MicrosoftGraph.Calendar[]> {
    const cacheKey = createCacheKey('listCalendars');
    const cached = this.cache.get<MicrosoftGraph.Calendar[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();

    const response = await client
      .api('/me/calendars')
      .select('id,name,color,isDefaultCalendar,canEdit')
      .get() as PageCollection;

    const result = response.value as MicrosoftGraph.Calendar[];
    this.cache.set(cacheKey, result, CacheTTL.FOLDERS);
    return result;
  }

  // ===========================================================================
  // Events
  // ===========================================================================

  /**
   * Lists events with optional date range.
   */
  async listEvents(
    limit: number = 50,
    calendarId?: string,
    startDate?: Date,
    endDate?: Date
  ): Promise<MicrosoftGraph.Event[]> {
    const client = await this.getClient();

    // If date range provided, use calendarView
    if (startDate != null && endDate != null) {
      const baseUrl = calendarId != null
        ? `/me/calendars/${calendarId}/calendarView`
        : '/me/calendarView';

      const response = await client
        .api(baseUrl)
        .query({
          startDateTime: startDate.toISOString(),
          endDateTime: endDate.toISOString(),
        })
        .select('id,subject,start,end,location,isAllDay,organizer,attendees,body,recurrence,iCalUId')
        .orderby('start/dateTime')
        .top(limit)
        .get() as PageCollection;

      return response.value as MicrosoftGraph.Event[];
    }

    // Otherwise, get upcoming events
    const baseUrl = calendarId != null
      ? `/me/calendars/${calendarId}/events`
      : '/me/events';

    const response = await client
      .api(baseUrl)
      .select('id,subject,start,end,location,isAllDay,organizer,attendees,body,recurrence,iCalUId')
      .orderby('start/dateTime')
      .top(limit)
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Event[];
  }

  /**
   * Gets a specific event.
   */
  async getEvent(eventId: string): Promise<MicrosoftGraph.Event | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/events/${eventId}`)
        .select('id,subject,start,end,location,isAllDay,organizer,attendees,body,recurrence,iCalUId')
        .get() as MicrosoftGraph.Event;
    } catch {
      return null;
    }
  }

  /**
   * Lists instances of a recurring event within a date range.
   */
  async listEventInstances(
    eventId: string,
    startDateTime: string,
    endDateTime: string
  ): Promise<MicrosoftGraph.Event[]> {
    const client = await this.getClient();
    const response = await client
      .api(`/me/events/${eventId}/instances`)
      .query({ startDateTime, endDateTime })
      .select('id,subject,start,end,location,isAllDay,isCancelled,organizer,recurrence,bodyPreview')
      .get() as PageCollection;
    return response.value as MicrosoftGraph.Event[];
  }

  // ===========================================================================
  // Calendar Write Operations
  // ===========================================================================

  /**
   * Creates a new calendar event.
   */
  async createEvent(
    event: Record<string, unknown>,
    calendarId?: string
  ): Promise<MicrosoftGraph.Event> {
    const client = await this.getClient();
    const url = calendarId != null
      ? `/me/calendars/${calendarId}/events`
      : '/me/events';

    const result = await client
      .api(url)
      .post(event) as MicrosoftGraph.Event;
    this.cache.clear();
    return result;
  }

  /**
   * Updates an existing calendar event.
   */
  async updateEvent(eventId: string, updates: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/events/${eventId}`)
      .patch(updates);
    this.cache.clear();
  }

  /**
   * Deletes a calendar event.
   */
  async deleteEvent(eventId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/events/${eventId}`)
      .delete();
    this.cache.clear();
  }

  /**
   * Responds to a calendar event invitation.
   */
  async respondToEvent(
    eventId: string,
    response: 'accept' | 'decline' | 'tentative',
    sendResponse: boolean,
    comment?: string
  ): Promise<void> {
    const client = await this.getClient();
    const actionMap: Record<string, string> = {
      accept: 'accept',
      decline: 'decline',
      tentative: 'tentativelyAccept',
    };
    const action = actionMap[response];

    await client
      .api(`/me/events/${eventId}/${action}`)
      .post({ sendResponse, comment: comment ?? '' });
    this.cache.clear();
  }

  // ===========================================================================
  // Contacts
  // ===========================================================================

  /**
   * Lists contacts with pagination.
   */
  async listContacts(limit: number = 50, skip: number = 0): Promise<MicrosoftGraph.Contact[]> {
    const cacheKey = createCacheKey('listContacts', limit, skip);
    const cached = this.cache.get<MicrosoftGraph.Contact[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();

    const response = await client
      .api('/me/contacts')
      .select('id,displayName,givenName,surname,middleName,nickName,companyName,jobTitle,department,emailAddresses,homePhones,businessPhones,mobilePhone,homeAddress,businessAddress,personalNotes')
      .orderby('displayName')
      .top(limit)
      .skip(skip)
      .get() as PageCollection;

    const result = response.value as MicrosoftGraph.Contact[];
    this.cache.set(cacheKey, result, CacheTTL.CONTACTS);
    return result;
  }

  /**
   * Searches contacts by display name.
   */
  async searchContacts(query: string, limit: number = 50): Promise<MicrosoftGraph.Contact[]> {
    const client = await this.getClient();

    const response = await client
      .api('/me/contacts')
      .filter(`contains(displayName,'${query}')`)
      .select('id,displayName,givenName,surname,middleName,nickName,companyName,jobTitle,department,emailAddresses,homePhones,businessPhones,mobilePhone,homeAddress,businessAddress,personalNotes')
      .top(limit)
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Contact[];
  }

  /**
   * Gets a specific contact.
   */
  async getContact(contactId: string): Promise<MicrosoftGraph.Contact | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/contacts/${contactId}`)
        .select('id,displayName,givenName,surname,middleName,nickName,companyName,jobTitle,department,emailAddresses,homePhones,businessPhones,mobilePhone,homeAddress,businessAddress,personalNotes')
        .get() as MicrosoftGraph.Contact;
    } catch {
      return null;
    }
  }

  // ===========================================================================
  // Contact Write Operations
  // ===========================================================================

  /**
   * Creates a new contact.
   */
  async createContact(contact: Record<string, unknown>): Promise<MicrosoftGraph.Contact> {
    const client = await this.getClient();
    const result = await client
      .api('/me/contacts')
      .post(contact) as MicrosoftGraph.Contact;
    this.cache.clear();
    return result;
  }

  /**
   * Updates an existing contact.
   */
  async updateContact(contactId: string, updates: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/contacts/${contactId}`)
      .patch(updates);
    this.cache.clear();
  }

  /**
   * Deletes a contact.
   */
  async deleteContact(contactId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/contacts/${contactId}`)
      .delete();
    this.cache.clear();
  }

  // ===========================================================================
  // Contact Photos
  // ===========================================================================

  /**
   * Gets the photo for a contact as raw binary data.
   */
  async getContactPhoto(contactId: string): Promise<ArrayBuffer> {
    const client = await this.getClient();
    return await client
      .api(`/me/contacts/${contactId}/photo/$value`)
      .get() as ArrayBuffer;
  }

  /**
   * Sets or updates the photo for a contact.
   */
  async setContactPhoto(contactId: string, photoData: Buffer, contentType: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/contacts/${contactId}/photo/$value`)
      .header('Content-Type', contentType)
      .put(photoData);
    this.cache.clear();
  }

  // ===========================================================================
  // Contact Folders
  // ===========================================================================

  /**
   * Lists all contact folders.
   */
  async listContactFolders(): Promise<MicrosoftGraph.ContactFolder[]> {
    const client = await this.getClient();
    const response = await client.api('/me/contactFolders').get() as PageCollection;
    return response.value as MicrosoftGraph.ContactFolder[];
  }

  /**
   * Creates a new contact folder.
   */
  async createContactFolder(displayName: string): Promise<MicrosoftGraph.ContactFolder> {
    const client = await this.getClient();
    const result = await client.api('/me/contactFolders').post({ displayName }) as MicrosoftGraph.ContactFolder;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes a contact folder.
   */
  async deleteContactFolder(folderId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/contactFolders/${folderId}`).delete();
    this.cache.clear();
  }

  /**
   * Lists contacts in a specific contact folder.
   */
  async listContactsInFolder(folderId: string, limit: number = 100): Promise<MicrosoftGraph.Contact[]> {
    const client = await this.getClient();
    const response = await client
      .api(`/me/contactFolders/${folderId}/contacts`)
      .select('id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,jobTitle,companyName')
      .top(limit)
      .get() as PageCollection;
    return response.value as MicrosoftGraph.Contact[];
  }

  // ===========================================================================
  // Tasks (Microsoft To Do)
  // ===========================================================================

  /**
   * Lists task lists.
   */
  async listTaskLists(): Promise<MicrosoftGraph.TodoTaskList[]> {
    const cacheKey = createCacheKey('listTaskLists');
    const cached = this.cache.get<MicrosoftGraph.TodoTaskList[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();

    const response = await client
      .api('/me/todo/lists')
      .select('id,displayName,isOwner,isShared,wellknownListName')
      .get() as PageCollection;

    const result = response.value as MicrosoftGraph.TodoTaskList[];
    this.cache.set(cacheKey, result, CacheTTL.FOLDERS);
    return result;
  }

  /**
   * Lists tasks in a task list.
   */
  async listTasks(
    taskListId: string,
    limit: number = 50,
    skip: number = 0,
    includeCompleted: boolean = true
  ): Promise<MicrosoftGraph.TodoTask[]> {
    const client = await this.getClient();

    let api = client
      .api(`/me/todo/lists/${taskListId}/tasks`)
      .select('id,title,status,importance,dueDateTime,completedDateTime,body,createdDateTime,lastModifiedDateTime,isReminderOn,reminderDateTime')
      .top(limit)
      .skip(skip);

    if (!includeCompleted) {
      api = api.filter("status ne 'completed'");
    }

    const response = await api.get() as PageCollection;
    return response.value as MicrosoftGraph.TodoTask[];
  }

  /**
   * Lists all tasks across all task lists.
   */
  async listAllTasks(
    limit: number = 50,
    skip: number = 0,
    includeCompleted: boolean = true
  ): Promise<Array<MicrosoftGraph.TodoTask & { taskListId: string }>> {
    const taskLists = await this.listTaskLists();
    const allTasks: Array<MicrosoftGraph.TodoTask & { taskListId: string }> = [];

    for (const list of taskLists) {
      if (list.id == null) continue;

      const tasks = await this.listTasks(list.id, 100, 0, includeCompleted);

      for (const task of tasks) {
        allTasks.push({ ...task, taskListId: list.id });
      }
    }

    // Sort by due date, then slice for pagination
    allTasks.sort((a, b) => {
      if (a.dueDateTime == null && b.dueDateTime == null) return 0;
      if (a.dueDateTime == null) return 1;
      if (b.dueDateTime == null) return -1;
      return new Date(a.dueDateTime.dateTime ?? '').getTime() - new Date(b.dueDateTime.dateTime ?? '').getTime();
    });

    return allTasks.slice(skip, skip + limit);
  }

  /**
   * Gets a specific task.
   */
  async getTask(taskListId: string, taskId: string): Promise<MicrosoftGraph.TodoTask | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/todo/lists/${taskListId}/tasks/${taskId}`)
        .select('id,title,status,importance,dueDateTime,completedDateTime,body,createdDateTime,lastModifiedDateTime,isReminderOn,reminderDateTime')
        .get() as MicrosoftGraph.TodoTask;
    } catch {
      return null;
    }
  }

  // ===========================================================================
  // Write Operations
  // ===========================================================================

  /**
   * Moves a message to a different folder.
   */
  async moveMessage(messageId: string, destinationFolderId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: destinationFolderId });
    this.cache.clear(); // Invalidate cache after mutation
  }

  /**
   * Deletes a message (moves to Deleted Items).
   */
  async deleteMessage(messageId: string): Promise<void> {
    const client = await this.getClient();
    // Move to deletedItems well-known folder
    await client
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: 'deleteditems' });
    this.cache.clear();
  }

  /**
   * Archives a message (moves to Archive folder).
   */
  async archiveMessage(messageId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: 'archive' });
    this.cache.clear();
  }

  /**
   * Moves a message to the Junk folder.
   */
  async junkMessage(messageId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: 'junkemail' });
    this.cache.clear();
  }

  /**
   * Updates message properties (read status, flag, categories).
   */
  async updateMessage(messageId: string, updates: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}`)
      .patch(updates);
    this.cache.clear();
  }

  /**
   * Creates a new mail folder.
   */
  async createMailFolder(
    displayName: string,
    parentFolderId?: string
  ): Promise<MicrosoftGraph.MailFolder> {
    const client = await this.getClient();
    const url = parentFolderId != null
      ? `/me/mailFolders/${parentFolderId}/childFolders`
      : '/me/mailFolders';

    const result = await client
      .api(url)
      .post({ displayName }) as MicrosoftGraph.MailFolder;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes a mail folder.
   */
  async deleteMailFolder(folderId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/mailFolders/${folderId}`)
      .delete();
    this.cache.clear();
  }

  /**
   * Renames a mail folder.
   */
  async renameMailFolder(folderId: string, newName: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/mailFolders/${folderId}`)
      .patch({ displayName: newName });
    this.cache.clear();
  }

  /**
   * Moves a mail folder to a new parent.
   */
  async moveMailFolder(folderId: string, destinationParentId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/mailFolders/${folderId}/move`)
      .post({ destinationId: destinationParentId });
    this.cache.clear();
  }

  /**
   * Deletes all messages in a folder.
   */
  async emptyMailFolder(folderId: string): Promise<void> {
    const client = await this.getClient();
    // Get all messages in the folder
    let response = await client
      .api(`/me/mailFolders/${folderId}/messages`)
      .select('id')
      .top(100)
      .get() as PageCollection;

    // Delete each message
    for (const message of response.value as MicrosoftGraph.Message[]) {
      if (message.id != null) {
        await client
          .api(`/me/messages/${message.id}/move`)
          .post({ destinationId: 'deleteditems' });
      }
    }

    // Handle pagination
    while (response['@odata.nextLink'] != null) {
      response = await client.api(response['@odata.nextLink']).get() as PageCollection;
      for (const message of response.value as MicrosoftGraph.Message[]) {
        if (message.id != null) {
          await client
            .api(`/me/messages/${message.id}/move`)
            .post({ destinationId: 'deleteditems' });
        }
      }
    }

    this.cache.clear();
  }

  // ===========================================================================
  // Draft & Send Operations
  // ===========================================================================

  /**
   * Creates a new draft message.
   */
  async createDraft(message: {
    subject: string;
    body: MicrosoftGraph.ItemBody;
    toRecipients?: MicrosoftGraph.Recipient[];
    ccRecipients?: MicrosoftGraph.Recipient[];
    bccRecipients?: MicrosoftGraph.Recipient[];
    isDraft?: boolean;
  }): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api('/me/messages')
      .post(message) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  /**
   * Updates an existing draft message.
   */
  async updateDraft(messageId: string, updates: Record<string, unknown>): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/messages/${messageId}`)
      .patch(updates) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  /**
   * Sends an existing draft message.
   */
  async sendDraft(messageId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}/send`)
      .post(null);
    this.cache.clear();
  }

  /**
   * Sends a new email directly without creating a draft.
   */
  async sendMail(message: {
    subject: string;
    body: MicrosoftGraph.ItemBody;
    toRecipients: MicrosoftGraph.Recipient[];
    ccRecipients?: MicrosoftGraph.Recipient[];
    bccRecipients?: MicrosoftGraph.Recipient[];
  }): Promise<void> {
    const client = await this.getClient();
    await client
      .api('/me/sendMail')
      .post({ message });
    this.cache.clear();
  }

  /**
   * Replies to a message, or replies to all recipients.
   */
  async replyMessage(messageId: string, comment: string, replyAll: boolean): Promise<void> {
    const client = await this.getClient();
    const action = replyAll ? 'replyAll' : 'reply';
    await client
      .api(`/me/messages/${messageId}/${action}`)
      .post({ comment });
    this.cache.clear();
  }

  /**
   * Forwards a message to specified recipients.
   */
  async forwardMessage(
    messageId: string,
    toRecipients: MicrosoftGraph.Recipient[],
    comment?: string
  ): Promise<void> {
    const client = await this.getClient();
    const body: { toRecipients: MicrosoftGraph.Recipient[]; comment?: string } = { toRecipients };
    if (comment != null) {
      body.comment = comment;
    }
    await client
      .api(`/me/messages/${messageId}/forward`)
      .post(body);
    this.cache.clear();
  }

  /**
   * Creates a reply draft for a message.
   */
  async createReplyDraft(messageId: string): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/messages/${messageId}/createReply`)
      .post(null) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  /**
   * Creates a reply-all draft for a message.
   */
  async createReplyAllDraft(messageId: string): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/messages/${messageId}/createReplyAll`)
      .post(null) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  /**
   * Creates a forward draft for a message.
   */
  async createForwardDraft(messageId: string): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/messages/${messageId}/createForward`)
      .post(null) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  // ---------------------------------------------------------------------------
  // Calendar Scheduling
  // ---------------------------------------------------------------------------

  /**
   * Gets the free/busy schedule for one or more people.
   * POST /me/calendar/getSchedule
   */
  async getSchedule(params: {
    schedules: string[];
    startTime: { dateTime: string; timeZone: string };
    endTime: { dateTime: string; timeZone: string };
    availabilityViewInterval?: number;
  }): Promise<unknown[]> {
    const client = await this.getClient();
    const response = await client.api('/me/calendar/getSchedule').post(params) as { value: unknown[] };
    return response.value;
  }

  /**
   * Suggests meeting times for a set of attendees.
   * POST /me/findMeetingTimes
   */
  async findMeetingTimes(params: {
    attendees: Array<{ emailAddress: { address: string }; type: string }>;
    meetingDuration: string;
    timeConstraint?: {
      timeslots: Array<{
        start: { dateTime: string; timeZone: string };
        end: { dateTime: string; timeZone: string };
      }>;
    };
    maxCandidates?: number;
  }): Promise<unknown> {
    const client = await this.getClient();
    return (await client.api('/me/findMeetingTimes').post(params)) as unknown;
  }

  // ===========================================================================
  // Attachment Operations
  // ===========================================================================

  /**
   * Lists attachments on a message.
   */
  async listAttachments(messageId: string): Promise<MicrosoftGraph.Attachment[]> {
    const client = await this.getClient();

    const response = await client
      .api(`/me/messages/${messageId}/attachments`)
      .select('id,name,size,contentType,isInline')
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Attachment[];
  }

  /**
   * Gets a specific attachment with full content (including contentBytes).
   */
  async getAttachment(messageId: string, attachmentId: string): Promise<MicrosoftGraph.FileAttachment> {
    const client = await this.getClient();

    return await client
      .api(`/me/messages/${messageId}/attachments/${attachmentId}`)
      .get() as MicrosoftGraph.FileAttachment;
  }

  /**
   * Adds an inline base64 attachment to a message (<= 3MB).
   */
  async addAttachment(messageId: string, attachment: Record<string, unknown>): Promise<MicrosoftGraph.Attachment> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/messages/${messageId}/attachments`)
      .post(attachment) as MicrosoftGraph.Attachment;
    this.cache.clear();
    return result;
  }

  /**
   * Creates an upload session for large file attachments (> 3MB).
   */
  async createUploadSession(messageId: string, body: Record<string, unknown>): Promise<{ uploadUrl: string }> {
    const client = await this.getClient();
    return await client
      .api(`/me/messages/${messageId}/attachments/createUploadSession`)
      .post(body) as { uploadUrl: string };
  }

  // ===========================================================================
  // Tasks (Microsoft To Do) - continued
  // ===========================================================================

  /**
   * Searches tasks by title.
   */
  async searchTasks(query: string, limit: number = 50): Promise<Array<MicrosoftGraph.TodoTask & { taskListId: string }>> {
    const allTasks = await this.listAllTasks(1000, 0, true);

    const queryLower = query.toLowerCase();
    const matched = allTasks.filter(
      (task) => task.title?.toLowerCase().includes(queryLower) ?? false
    );

    return matched.slice(0, limit);
  }

  // ===========================================================================
  // Task Write Operations
  // ===========================================================================

  /**
   * Creates a new task in a task list.
   */
  async createTask(
    taskListId: string,
    task: Record<string, unknown>
  ): Promise<MicrosoftGraph.TodoTask> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/todo/lists/${taskListId}/tasks`)
      .post(task) as MicrosoftGraph.TodoTask;
    this.cache.clear();
    return result;
  }

  /**
   * Updates an existing task.
   */
  async updateTask(
    taskListId: string,
    taskId: string,
    updates: Record<string, unknown>
  ): Promise<MicrosoftGraph.TodoTask> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/todo/lists/${taskListId}/tasks/${taskId}`)
      .patch(updates) as MicrosoftGraph.TodoTask;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes a task.
   */
  async deleteTask(taskListId: string, taskId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/todo/lists/${taskListId}/tasks/${taskId}`)
      .delete();
    this.cache.clear();
  }

  /**
   * Creates a new task list.
   */
  async createTaskList(displayName: string): Promise<MicrosoftGraph.TodoTaskList> {
    const client = await this.getClient();
    const result = await client
      .api('/me/todo/lists')
      .post({ displayName }) as MicrosoftGraph.TodoTaskList;
    this.cache.clear();
    return result;
  }

  /**
   * Updates a task list (e.g. rename).
   */
  async updateTaskList(listId: string, updates: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/todo/lists/${listId}`).patch(updates);
    this.cache.clear();
  }

  /**
   * Deletes a task list.
   */
  async deleteTaskList(listId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/todo/lists/${listId}`).delete();
    this.cache.clear();
  }

  // ===========================================================================
  // Master Categories
  // ===========================================================================

  /**
   * Lists all master categories.
   */
  async listMasterCategories(): Promise<MicrosoftGraph.OutlookCategory[]> {
    const client = await this.getClient();
    const response = await client.api('/me/outlook/masterCategories').get() as PageCollection;
    return response.value as MicrosoftGraph.OutlookCategory[];
  }

  /**
   * Creates a new master category.
   */
  async createMasterCategory(displayName: string, color: string): Promise<MicrosoftGraph.OutlookCategory> {
    const client = await this.getClient();
    const result = await client.api('/me/outlook/masterCategories').post({ displayName, color }) as MicrosoftGraph.OutlookCategory;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes a master category.
   */
  async deleteMasterCategory(categoryId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/outlook/masterCategories/${categoryId}`).delete();
    this.cache.clear();
  }

  // ===========================================================================
  // Focused Inbox Overrides
  // ===========================================================================

  /**
   * Lists all focused inbox overrides.
   */
  async listFocusedOverrides(): Promise<MicrosoftGraph.InferenceClassificationOverride[]> {
    const client = await this.getClient();
    const response = await client.api('/me/inferenceClassification/overrides').get() as PageCollection;
    return response.value as MicrosoftGraph.InferenceClassificationOverride[];
  }

  /**
   * Creates a focused inbox override.
   */
  async createFocusedOverride(senderAddress: string, classifyAs: 'focused' | 'other'): Promise<MicrosoftGraph.InferenceClassificationOverride> {
    const client = await this.getClient();
    const result = await client.api('/me/inferenceClassification/overrides').post({
      classifyAs,
      senderEmailAddress: { address: senderAddress },
    }) as MicrosoftGraph.InferenceClassificationOverride;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes a focused inbox override.
   */
  async deleteFocusedOverride(overrideId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/inferenceClassification/overrides/${overrideId}`).delete();
    this.cache.clear();
  }

  // ===========================================================================
  // Message Headers & MIME
  // ===========================================================================

  /**
   * Gets internet message headers for a message.
   */
  async getMessageHeaders(messageId: string): Promise<Array<{ name: string; value: string }>> {
    const client = await this.getClient();
    const message = await client
      .api(`/me/messages/${messageId}`)
      .select('internetMessageHeaders')
      .get() as MicrosoftGraph.Message;
    return (message.internetMessageHeaders ?? []) as Array<{ name: string; value: string }>;
  }

  /**
   * Gets the MIME content of a message.
   */
  async getMessageMime(messageId: string): Promise<string> {
    const client = await this.getClient();
    return await client.api(`/me/messages/${messageId}/$value`).get() as string;
  }

  // ===========================================================================
  // Calendar Groups
  // ===========================================================================

  /**
   * Lists all calendar groups.
   */
  async listCalendarGroups(): Promise<MicrosoftGraph.CalendarGroup[]> {
    const client = await this.getClient();
    const response = await client.api('/me/calendarGroups').get() as PageCollection;
    return response.value as MicrosoftGraph.CalendarGroup[];
  }

  /**
   * Creates a new calendar group.
   */
  async createCalendarGroup(name: string): Promise<MicrosoftGraph.CalendarGroup> {
    const client = await this.getClient();
    const result = await client.api('/me/calendarGroups').post({ name }) as MicrosoftGraph.CalendarGroup;
    this.cache.clear();
    return result;
  }

  // ===========================================================================
  // Calendar Permissions
  // ===========================================================================

  /**
   * Lists all permissions for a calendar.
   */
  async listCalendarPermissions(calendarId: string): Promise<MicrosoftGraph.CalendarPermission[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/calendars/${calendarId}/calendarPermissions`).get() as PageCollection;
    return response.value as MicrosoftGraph.CalendarPermission[];
  }

  /**
   * Creates a calendar permission (shares a calendar).
   */
  async createCalendarPermission(calendarId: string, permission: Record<string, unknown>): Promise<MicrosoftGraph.CalendarPermission> {
    const client = await this.getClient();
    const result = await client.api(`/me/calendars/${calendarId}/calendarPermissions`).post(permission) as MicrosoftGraph.CalendarPermission;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes a calendar permission.
   */
  async deleteCalendarPermission(calendarId: string, permissionId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/calendars/${calendarId}/calendarPermissions/${permissionId}`).delete();
    this.cache.clear();
  }

  // ===========================================================================
  // Room Lists & Rooms
  // ===========================================================================

  /**
   * GET /me/findRoomLists
   */
  async listRoomLists(): Promise<MicrosoftGraph.EmailAddress[]> {
    const client = await this.getClient();
    const response = await client.api('/me/findRoomLists').get() as { value: MicrosoftGraph.EmailAddress[] };
    return response.value;
  }

  /**
   * GET /me/findRooms or /me/findRooms(RoomList='...')
   */
  async listRooms(roomListEmail?: string): Promise<MicrosoftGraph.EmailAddress[]> {
    const client = await this.getClient();
    const endpoint = roomListEmail != null
      ? `/me/findRooms(RoomList='${roomListEmail}')`
      : '/me/findRooms';
    const response = await client.api(endpoint).get() as { value: MicrosoftGraph.EmailAddress[] };
    return response.value;
  }

  // ===========================================================================
  // Mail Tips
  // ===========================================================================

  /**
   * Gets mail tips for the specified email addresses.
   */
  async getMailTips(emailAddresses: string[]): Promise<Record<string, unknown>[]> {
    const client = await this.getClient();
    const response = await client.api('/me/getMailTips').post({
      emailAddresses,
      mailTipsOptions: 'automaticReplies,mailboxFullStatus,maxMessageSize,deliveryRestriction,externalMemberCount',
    }) as { value: Record<string, unknown>[] };
    return response.value;
  }

  // ===========================================================================
  // Teams
  // ===========================================================================

  /**
   * Lists all teams the current user has joined.
   */
  async listJoinedTeams(): Promise<MicrosoftGraph.Team[]> {
    const client = await this.getClient();
    const response = await client.api('/me/joinedTeams').get() as PageCollection;
    return response.value as MicrosoftGraph.Team[];
  }

  /**
   * Lists all channels in a team.
   */
  async listChannels(teamId: string): Promise<MicrosoftGraph.Channel[]> {
    const client = await this.getClient();
    const response = await client.api(`/teams/${teamId}/channels`).get() as PageCollection;
    return response.value as MicrosoftGraph.Channel[];
  }

  /**
   * Gets a specific channel.
   */
  async getChannel(teamId: string, channelId: string): Promise<MicrosoftGraph.Channel> {
    const client = await this.getClient();
    return await client.api(`/teams/${teamId}/channels/${channelId}`).get() as MicrosoftGraph.Channel;
  }

  /**
   * Creates a new channel in a team.
   */
  async createChannel(teamId: string, displayName: string, description?: string): Promise<MicrosoftGraph.Channel> {
    const client = await this.getClient();
    const body: Record<string, unknown> = { displayName };
    if (description != null) body['description'] = description;
    return await client.api(`/teams/${teamId}/channels`).post(body) as MicrosoftGraph.Channel;
  }

  /**
   * Updates a channel's properties.
   */
  async updateChannel(teamId: string, channelId: string, updates: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client.api(`/teams/${teamId}/channels/${channelId}`).patch(updates);
  }

  /**
   * Deletes a channel.
   */
  async deleteChannel(teamId: string, channelId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/teams/${teamId}/channels/${channelId}`).delete();
  }

  /**
   * Lists members of a team.
   */
  async listTeamMembers(teamId: string): Promise<MicrosoftGraph.ConversationMember[]> {
    const client = await this.getClient();
    const response = await client.api(`/teams/${teamId}/members`).get() as PageCollection;
    return response.value as MicrosoftGraph.ConversationMember[];
  }

  // ===========================================================================
  // Channel Messages
  // ===========================================================================

  /**
   * Lists recent messages in a channel.
   */
  async listChannelMessages(teamId: string, channelId: string, top: number = 25): Promise<MicrosoftGraph.ChatMessage[]> {
    const client = await this.getClient();
    const response = await client.api(`/teams/${teamId}/channels/${channelId}/messages`).top(top).get() as PageCollection;
    return response.value as MicrosoftGraph.ChatMessage[];
  }

  /**
   * Gets a specific channel message.
   */
  async getChannelMessage(teamId: string, channelId: string, messageId: string): Promise<MicrosoftGraph.ChatMessage> {
    const client = await this.getClient();
    return await client.api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}`).get() as MicrosoftGraph.ChatMessage;
  }

  /**
   * Lists replies to a channel message.
   */
  async listChannelMessageReplies(teamId: string, channelId: string, messageId: string): Promise<MicrosoftGraph.ChatMessage[]> {
    const client = await this.getClient();
    const response = await client.api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`).get() as PageCollection;
    return response.value as MicrosoftGraph.ChatMessage[];
  }

  /**
   * Sends a new message to a channel.
   */
  async sendChannelMessage(teamId: string, channelId: string, body: string, contentType: string = 'html'): Promise<MicrosoftGraph.ChatMessage> {
    const client = await this.getClient();
    return await client.api(`/teams/${teamId}/channels/${channelId}/messages`).post({
      body: { contentType, content: body },
    }) as MicrosoftGraph.ChatMessage;
  }

  /**
   * Replies to a channel message.
   */
  async replyToChannelMessage(teamId: string, channelId: string, messageId: string, body: string, contentType: string = 'html'): Promise<MicrosoftGraph.ChatMessage> {
    const client = await this.getClient();
    return await client.api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`).post({
      body: { contentType, content: body },
    }) as MicrosoftGraph.ChatMessage;
  }

  // ===========================================================================
  // Chats
  // ===========================================================================

  async listChats(top: number = 25): Promise<MicrosoftGraph.Chat[]> {
    const client = await this.getClient();
    const response = await client.api('/me/chats')
      .top(top)
      .orderby('lastMessagePreview/createdDateTime desc')
      .expand('lastMessagePreview')
      .get() as PageCollection;
    return response.value as MicrosoftGraph.Chat[];
  }

  async getChat(chatId: string): Promise<MicrosoftGraph.Chat> {
    const client = await this.getClient();
    return await client.api(`/me/chats/${chatId}`).get() as MicrosoftGraph.Chat;
  }

  async listChatMessages(chatId: string, top: number = 25): Promise<MicrosoftGraph.ChatMessage[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/chats/${chatId}/messages`).top(top).get() as PageCollection;
    return response.value as MicrosoftGraph.ChatMessage[];
  }

  async sendChatMessage(chatId: string, body: string, contentType: string = 'html'): Promise<MicrosoftGraph.ChatMessage> {
    const client = await this.getClient();
    return await client.api(`/me/chats/${chatId}/messages`).post({
      body: { contentType, content: body },
    }) as MicrosoftGraph.ChatMessage;
  }

  async listChatMembers(chatId: string): Promise<MicrosoftGraph.ConversationMember[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/chats/${chatId}/members`).get() as PageCollection;
    return response.value as MicrosoftGraph.ConversationMember[];
  }

  // ===========================================================================
  // Checklist Items
  // ===========================================================================

  /**
   * Lists checklist items on a task.
   */
  async listChecklistItems(taskListId: string, taskId: string): Promise<MicrosoftGraph.ChecklistItem[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/checklistItems`).get() as PageCollection;
    return response.value as MicrosoftGraph.ChecklistItem[];
  }

  /**
   * Creates a checklist item on a task.
   */
  async createChecklistItem(taskListId: string, taskId: string, displayName: string, isChecked: boolean = false): Promise<MicrosoftGraph.ChecklistItem> {
    const client = await this.getClient();
    return await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/checklistItems`).post({
      displayName,
      isChecked,
    }) as MicrosoftGraph.ChecklistItem;
  }

  /**
   * Updates a checklist item.
   */
  async updateChecklistItem(taskListId: string, taskId: string, checklistItemId: string, updates: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/checklistItems/${checklistItemId}`).patch(updates);
  }

  /**
   * Deletes a checklist item.
   */
  async deleteChecklistItem(taskListId: string, taskId: string, checklistItemId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/checklistItems/${checklistItemId}`).delete();
  }

  // ===========================================================================
  // Linked Resources
  // ===========================================================================

  async listLinkedResources(taskListId: string, taskId: string): Promise<MicrosoftGraph.LinkedResource[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/linkedResources`).get() as PageCollection;
    return response.value as MicrosoftGraph.LinkedResource[];
  }

  async createLinkedResource(taskListId: string, taskId: string, webUrl: string, applicationName: string, displayName?: string): Promise<MicrosoftGraph.LinkedResource> {
    const client = await this.getClient();
    const body: Record<string, unknown> = { webUrl, applicationName };
    if (displayName != null) body['displayName'] = displayName;
    return await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/linkedResources`).post(body) as MicrosoftGraph.LinkedResource;
  }

  async deleteLinkedResource(taskListId: string, taskId: string, linkedResourceId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/linkedResources/${linkedResourceId}`).delete();
  }

  // ===========================================================================
  // Task Attachments
  // ===========================================================================

  async listTaskAttachments(taskListId: string, taskId: string): Promise<MicrosoftGraph.AttachmentBase[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/attachments`).get() as PageCollection;
    return response.value as MicrosoftGraph.AttachmentBase[];
  }

  async createTaskAttachment(taskListId: string, taskId: string, name: string, contentBytes: string, contentType: string = 'application/octet-stream'): Promise<MicrosoftGraph.AttachmentBase> {
    const client = await this.getClient();
    return await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/attachments`).post({
      '@odata.type': '#microsoft.graph.taskFileAttachment',
      name,
      contentBytes,
      contentType,
    }) as MicrosoftGraph.AttachmentBase;
  }

  async deleteTaskAttachment(taskListId: string, taskId: string, attachmentId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/todo/lists/${taskListId}/tasks/${taskId}/attachments/${attachmentId}`).delete();
  }

  // ===========================================================================
  // Planner
  // ===========================================================================

  async listPlans(): Promise<MicrosoftGraph.PlannerPlan[]> {
    const client = await this.getClient();
    const response = await client.api('/me/planner/plans').get() as PageCollection;
    return response.value as MicrosoftGraph.PlannerPlan[];
  }

  async getPlan(planId: string): Promise<MicrosoftGraph.PlannerPlan> {
    const client = await this.getClient();
    return await client.api(`/planner/plans/${planId}`).get() as MicrosoftGraph.PlannerPlan;
  }

  async createPlan(title: string, groupId: string): Promise<MicrosoftGraph.PlannerPlan> {
    const client = await this.getClient();
    return await client.api('/planner/plans').post({
      title,
      owner: groupId,
      container: { url: `https://graph.microsoft.com/v1.0/groups/${groupId}`, type: 'group' },
    }) as MicrosoftGraph.PlannerPlan;
  }

  async updatePlan(planId: string, updates: Record<string, unknown>, etag: string): Promise<MicrosoftGraph.PlannerPlan> {
    const client = await this.getClient();
    return await client.api(`/planner/plans/${planId}`).header('If-Match', etag).patch(updates) as MicrosoftGraph.PlannerPlan;
  }

  async listBuckets(planId: string): Promise<MicrosoftGraph.PlannerBucket[]> {
    const client = await this.getClient();
    const response = await client.api(`/planner/plans/${planId}/buckets`).get() as PageCollection;
    return response.value as MicrosoftGraph.PlannerBucket[];
  }

  async createBucket(planId: string, name: string): Promise<MicrosoftGraph.PlannerBucket> {
    const client = await this.getClient();
    return await client.api('/planner/buckets').post({ planId, name }) as MicrosoftGraph.PlannerBucket;
  }

  async updateBucket(bucketId: string, updates: Record<string, unknown>, etag: string): Promise<MicrosoftGraph.PlannerBucket> {
    const client = await this.getClient();
    return await client.api(`/planner/buckets/${bucketId}`).header('If-Match', etag).patch(updates) as MicrosoftGraph.PlannerBucket;
  }

  async deleteBucket(bucketId: string, etag: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/planner/buckets/${bucketId}`).header('If-Match', etag).delete();
  }

  // ===========================================================================
  // Planner Tasks
  // ===========================================================================

  async listPlannerTasks(planId: string): Promise<MicrosoftGraph.PlannerTask[]> {
    const client = await this.getClient();
    const response = await client.api(`/planner/plans/${planId}/tasks`).get() as PageCollection;
    return response.value as MicrosoftGraph.PlannerTask[];
  }

  async getPlannerTask(taskId: string): Promise<MicrosoftGraph.PlannerTask> {
    const client = await this.getClient();
    return await client.api(`/planner/tasks/${taskId}`).get() as MicrosoftGraph.PlannerTask;
  }

  async createPlannerTask(task: Record<string, unknown>): Promise<MicrosoftGraph.PlannerTask> {
    const client = await this.getClient();
    return await client.api('/planner/tasks').post(task) as MicrosoftGraph.PlannerTask;
  }

  async updatePlannerTask(taskId: string, updates: Record<string, unknown>, etag: string): Promise<MicrosoftGraph.PlannerTask> {
    const client = await this.getClient();
    return await client.api(`/planner/tasks/${taskId}`).header('If-Match', etag).patch(updates) as MicrosoftGraph.PlannerTask;
  }

  async deletePlannerTask(taskId: string, etag: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/planner/tasks/${taskId}`).header('If-Match', etag).delete();
  }

  async getPlannerTaskDetails(taskId: string): Promise<MicrosoftGraph.PlannerTaskDetails> {
    const client = await this.getClient();
    return await client.api(`/planner/tasks/${taskId}/details`).get() as MicrosoftGraph.PlannerTaskDetails;
  }

  async updatePlannerTaskDetails(taskId: string, updates: Record<string, unknown>, etag: string): Promise<MicrosoftGraph.PlannerTaskDetails> {
    const client = await this.getClient();
    return await client.api(`/planner/tasks/${taskId}/details`).header('If-Match', etag).patch(updates) as MicrosoftGraph.PlannerTaskDetails;
  }

  // ===========================================================================
  // People & Presence
  // ===========================================================================

  async listRelevantPeople(top: number = 25): Promise<MicrosoftGraph.Person[]> {
    const client = await this.getClient();
    const response = await client.api('/me/people').top(top).get() as PageCollection;
    return response.value as MicrosoftGraph.Person[];
  }

  async searchPeople(query: string, top: number = 25): Promise<MicrosoftGraph.Person[]> {
    const client = await this.getClient();
    const response = await client.api('/me/people').search('"' + query + '"').top(top).get() as PageCollection;
    return response.value as MicrosoftGraph.Person[];
  }

  async getManager(): Promise<MicrosoftGraph.DirectoryObject> {
    const client = await this.getClient();
    return await client.api('/me/manager').get() as MicrosoftGraph.DirectoryObject;
  }

  async getDirectReports(): Promise<MicrosoftGraph.DirectoryObject[]> {
    const client = await this.getClient();
    const response = await client.api('/me/directReports').get() as PageCollection;
    return response.value as MicrosoftGraph.DirectoryObject[];
  }

  async getUserProfile(identifier: string): Promise<MicrosoftGraph.User> {
    const client = await this.getClient();
    return await client.api(`/users/${identifier}`).get() as MicrosoftGraph.User;
  }

  async getUserPhoto(identifier: string): Promise<ArrayBuffer> {
    const client = await this.getClient();
    return await client.api(`/users/${identifier}/photo/$value`).get() as ArrayBuffer;
  }

  async getUserPresence(identifier: string): Promise<MicrosoftGraph.Presence> {
    const client = await this.getClient();
    return await client.api(`/users/${identifier}/presence`).get() as MicrosoftGraph.Presence;
  }

  async getUsersPresence(userIds: string[]): Promise<MicrosoftGraph.Presence[]> {
    const client = await this.getClient();
    const response = await client.api('/communications/getPresencesByUserId').post({ ids: userIds });
    return response.value as MicrosoftGraph.Presence[];
  }
  // ===========================================================================
  // OneDrive
  // ===========================================================================

  async listDriveItems(itemId?: string): Promise<any[]> {
    const client = await this.getClient();
    const path = itemId ? `/me/drive/items/${itemId}/children` : '/me/drive/root/children';
    const response = await client.api(path).get();
    return response.value;
  }

  async searchDriveItems(query: string, limit: number = 25): Promise<any[]> {
    const client = await this.getClient();
    const response = await client.api(`/me/drive/root/search(q='${encodeURIComponent(query)}')`).top(limit).get();
    return response.value;
  }

  async getDriveItem(itemId: string): Promise<any> {
    const client = await this.getClient();
    return await client.api(`/me/drive/items/${itemId}`).get();
  }

  async downloadDriveItem(itemId: string): Promise<ArrayBuffer> {
    const client = await this.getClient();
    return await client.api(`/me/drive/items/${itemId}/content`).get();
  }

  async uploadDriveItem(parentPath: string, fileName: string, content: Buffer): Promise<any> {
    const client = await this.getClient();
    return await client.api(`/me/drive/root:/${parentPath}/${fileName}:/content`)
      .header('Content-Type', 'application/octet-stream')
      .put(content);
  }

  async listRecentDriveItems(): Promise<any[]> {
    const client = await this.getClient();
    const response = await client.api('/me/drive/recent').get();
    return response.value;
  }

  async listSharedWithMe(): Promise<any[]> {
    const client = await this.getClient();
    const response = await client.api('/me/drive/sharedWithMe').get();
    return response.value;
  }

  async createSharingLink(itemId: string, type: string, scope: string): Promise<any> {
    const client = await this.getClient();
    return await client.api(`/me/drive/items/${itemId}/createLink`).post({ type, scope });
  }

  async deleteDriveItem(itemId: string): Promise<void> {
    const client = await this.getClient();
    await client.api(`/me/drive/items/${itemId}`).delete();
  }
}

/**
 * Creates a new Graph client instance.
 */
export function createGraphClient(deviceCodeCallback?: DeviceCodeCallback): GraphClient {
  return new GraphClient(deviceCodeCallback);
}
