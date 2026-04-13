/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Microsoft People API MCP tools.
 *
 * Provides tools for discovering relevant people, searching by name/email,
 * viewing org chart (manager/direct reports), user profiles, photos, and presence.
 */

import { z } from 'zod';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

// =============================================================================
// Input Schemas
// =============================================================================

export const ListRelevantPeopleInput = z.strictObject({
  limit: z.number().int().min(1).max(100).optional().describe('Max people to return (default 25, max 100)'),
});

export const SearchPeopleInput = z.strictObject({
  query: z.string().min(1).describe('Search query (name or email)'),
  limit: z.number().int().min(1).max(100).optional().describe('Max results to return (default 25, max 100)'),
});

export const GetManagerInput = z.strictObject({});

export const GetDirectReportsInput = z.strictObject({});

export const GetUserProfileInput = z.strictObject({
  identifier: z.string().min(1).describe('User email address or user ID'),
});

export const GetUserPhotoInput = z.strictObject({
  identifier: z.string().min(1).describe('User email address or user ID'),
  save_path: z.string().optional().describe('File path to save the photo (defaults to ~/Downloads/{identifier}_photo.jpg)'),
});

export const GetUserPresenceInput = z.strictObject({
  identifier: z.string().min(1).describe('User email address or user ID'),
});

export const GetUsersPresenceInput = z.strictObject({
  user_ids: z.array(z.string().min(1)).min(1).max(650).describe('Array of user IDs (max 650)'),
});

// =============================================================================
// Type Exports
// =============================================================================

export type ListRelevantPeopleParams = z.infer<typeof ListRelevantPeopleInput>;
export type SearchPeopleParams = z.infer<typeof SearchPeopleInput>;
export type GetManagerParams = z.infer<typeof GetManagerInput>;
export type GetDirectReportsParams = z.infer<typeof GetDirectReportsInput>;
export type GetUserProfileParams = z.infer<typeof GetUserProfileInput>;
export type GetUserPhotoParams = z.infer<typeof GetUserPhotoInput>;
export type GetUserPresenceParams = z.infer<typeof GetUserPresenceInput>;
export type GetUsersPresenceParams = z.infer<typeof GetUsersPresenceInput>;

// =============================================================================
// Client Interface
// =============================================================================

export interface IPeopleClient {
  listRelevantPeople(top?: number): Promise<Array<{
    displayName?: string | null;
    givenName?: string | null;
    surname?: string | null;
    scoredEmailAddresses?: Array<{ address?: string | null }> | null;
    jobTitle?: string | null;
    department?: string | null;
    officeLocation?: string | null;
  }>>;
  searchPeople(query: string, top?: number): Promise<Array<{
    displayName?: string | null;
    givenName?: string | null;
    surname?: string | null;
    scoredEmailAddresses?: Array<{ address?: string | null }> | null;
    jobTitle?: string | null;
    department?: string | null;
    officeLocation?: string | null;
  }>>;
  getManager(): Promise<{
    id?: string | null; displayName?: string | null; mail?: string | null;
    jobTitle?: string | null; department?: string | null; officeLocation?: string | null;
  }>;
  getDirectReports(): Promise<Array<{
    id?: string | null; displayName?: string | null; mail?: string | null;
    jobTitle?: string | null; department?: string | null; officeLocation?: string | null;
  }>>;
  getUserProfile(identifier: string): Promise<{
    id?: string | null; displayName?: string | null; mail?: string | null;
    jobTitle?: string | null; department?: string | null; officeLocation?: string | null;
    mobilePhone?: string | null; businessPhones?: string[] | null;
  }>;
  getUserPhoto(identifier: string): Promise<ArrayBuffer>;
  getUserPresence(identifier: string): Promise<{
    availability?: string | null;
    activity?: string | null;
    statusMessage?: { message?: { content?: string | null } | null } | null;
  }>;
  getUsersPresence(userIds: string[]): Promise<Array<{
    id?: string | null;
    availability?: string | null;
    activity?: string | null;
  }>>;
}

// =============================================================================
// Result Type
// =============================================================================

type ToolResult = {
  content: Array<{ type: 'text'; text: string }>;
};

// =============================================================================
// Helper Functions
// =============================================================================

function mapPerson(person: {
  displayName?: string | null;
  givenName?: string | null;
  surname?: string | null;
  scoredEmailAddresses?: Array<{ address?: string | null }> | null;
  jobTitle?: string | null;
  department?: string | null;
  officeLocation?: string | null;
}): {
  displayName: string | null;
  givenName: string | null;
  surname: string | null;
  emailAddresses: string[];
  jobTitle: string | null;
  department: string | null;
  officeLocation: string | null;
} {
  return {
    displayName: person.displayName ?? null,
    givenName: person.givenName ?? null,
    surname: person.surname ?? null,
    emailAddresses: (person.scoredEmailAddresses ?? [])
      .map((e) => e.address)
      .filter((a): a is string => a != null),
    jobTitle: person.jobTitle ?? null,
    department: person.department ?? null,
    officeLocation: person.officeLocation ?? null,
  };
}

function mapDirectoryObject(obj: {
  id?: string | null; displayName?: string | null; mail?: string | null;
  jobTitle?: string | null; department?: string | null; officeLocation?: string | null;
}): {
  id: string | null;
  displayName: string | null;
  mail: string | null;
  jobTitle: string | null;
  department: string | null;
  officeLocation: string | null;
} {
  return {
    id: obj.id ?? null,
    displayName: obj.displayName ?? null,
    mail: obj.mail ?? null,
    jobTitle: obj.jobTitle ?? null,
    department: obj.department ?? null,
    officeLocation: obj.officeLocation ?? null,
  };
}

// =============================================================================
// People Tools
// =============================================================================

/**
 * Microsoft People API tools for discovering and querying user information.
 */
export class PeopleTools {
  constructor(
    private readonly client: IPeopleClient,
  ) {}

  async listRelevantPeople(params: ListRelevantPeopleParams): Promise<ToolResult> {
    const limit = params.limit ?? 25;
    const people = await this.client.listRelevantPeople(limit);
    const mapped = people.map(mapPerson);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ people: mapped }, null, 2),
      }],
    };
  }

  async searchPeople(params: SearchPeopleParams): Promise<ToolResult> {
    const limit = params.limit ?? 25;
    const people = await this.client.searchPeople(params.query, limit);
    const mapped = people.map(mapPerson);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ people: mapped }, null, 2),
      }],
    };
  }

  async getManager(): Promise<ToolResult> {
    const manager = await this.client.getManager();
    const mapped = mapDirectoryObject(manager);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ manager: mapped }, null, 2),
      }],
    };
  }

  async getDirectReports(): Promise<ToolResult> {
    const reports = await this.client.getDirectReports();
    const mapped = reports.map(mapDirectoryObject);
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ direct_reports: mapped }, null, 2),
      }],
    };
  }

  async getUserProfile(params: GetUserProfileParams): Promise<ToolResult> {
    const user = await this.client.getUserProfile(params.identifier);
    const mapped = {
      id: user.id ?? null,
      displayName: user.displayName ?? null,
      mail: user.mail ?? null,
      jobTitle: user.jobTitle ?? null,
      department: user.department ?? null,
      officeLocation: user.officeLocation ?? null,
      mobilePhone: user.mobilePhone ?? null,
      businessPhones: user.businessPhones ?? [],
    };
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ user: mapped }, null, 2),
      }],
    };
  }

  async getUserPhoto(params: GetUserPhotoParams): Promise<ToolResult> {
    const savePath = params.save_path ?? path.join(os.homedir(), 'Downloads', `${params.identifier}_photo.jpg`);

    try {
      const photoData = await this.client.getUserPhoto(params.identifier);
      const buffer = Buffer.from(photoData);

      // Ensure directory exists
      const dir = path.dirname(savePath);
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }

      fs.writeFileSync(savePath, buffer);

      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ saved_to: savePath, size: buffer.length }, null, 2),
        }],
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to get user photo';
      return {
        content: [{
          type: 'text' as const,
          text: JSON.stringify({ error: message }, null, 2),
        }],
      };
    }
  }

  async getUserPresence(params: GetUserPresenceParams): Promise<ToolResult> {
    const presence = await this.client.getUserPresence(params.identifier);
    const mapped = {
      availability: presence.availability ?? null,
      activity: presence.activity ?? null,
      statusMessage: presence.statusMessage?.message?.content ?? null,
    };
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ presence: mapped }, null, 2),
      }],
    };
  }

  async getUsersPresence(params: GetUsersPresenceParams): Promise<ToolResult> {
    const presences = await this.client.getUsersPresence(params.user_ids);
    const mapped = presences.map((p) => ({
      id: p.id ?? null,
      availability: p.availability ?? null,
      activity: p.activity ?? null,
    }));
    return {
      content: [{
        type: 'text' as const,
        text: JSON.stringify({ presences: mapped }, null, 2),
      }],
    };
  }
}
