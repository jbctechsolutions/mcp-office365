/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Microsoft People API tools.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { PeopleTools, type IPeopleClient } from '../../../src/tools/people.js';

// Mock fs module
vi.mock('fs', () => ({
  existsSync: vi.fn().mockReturnValue(true),
  writeFileSync: vi.fn(),
  mkdirSync: vi.fn(),
}));

import * as fs from 'fs';

describe('PeopleTools', () => {
  let client: IPeopleClient;
  let tools: PeopleTools;

  beforeEach(() => {
    vi.clearAllMocks();
    client = {
      listRelevantPeople: vi.fn(),
      searchPeople: vi.fn(),
      getManager: vi.fn(),
      getDirectReports: vi.fn(),
      getUserProfile: vi.fn(),
      getUserPhoto: vi.fn(),
      getUserPresence: vi.fn(),
      getUsersPresence: vi.fn(),
    };
    tools = new PeopleTools(client);
  });

  describe('listRelevantPeople', () => {
    it('returns mapped people with default limit', async () => {
      const mockPeople = [
        {
          displayName: 'Alice Smith',
          givenName: 'Alice',
          surname: 'Smith',
          scoredEmailAddresses: [{ address: 'alice@example.com' }],
          jobTitle: 'Engineer',
          department: 'Engineering',
          officeLocation: 'Building A',
        },
        {
          displayName: 'Bob Jones',
          givenName: 'Bob',
          surname: 'Jones',
          scoredEmailAddresses: [{ address: 'bob@example.com' }, { address: 'bob.jones@example.com' }],
          jobTitle: 'Manager',
          department: 'Product',
          officeLocation: null,
        },
      ];
      vi.mocked(client.listRelevantPeople).mockResolvedValue(mockPeople);

      const result = await tools.listRelevantPeople({});

      expect(client.listRelevantPeople).toHaveBeenCalledWith(25);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.people).toHaveLength(2);
      expect(parsed.people[0]).toEqual({
        displayName: 'Alice Smith',
        givenName: 'Alice',
        surname: 'Smith',
        emailAddresses: ['alice@example.com'],
        jobTitle: 'Engineer',
        department: 'Engineering',
        officeLocation: 'Building A',
      });
      expect(parsed.people[1].emailAddresses).toEqual(['bob@example.com', 'bob.jones@example.com']);
      expect(parsed.people[1].officeLocation).toBeNull();
    });

    it('passes custom limit', async () => {
      vi.mocked(client.listRelevantPeople).mockResolvedValue([]);

      await tools.listRelevantPeople({ limit: 50 });

      expect(client.listRelevantPeople).toHaveBeenCalledWith(50);
    });

    it('handles people with null/undefined fields', async () => {
      const mockPeople = [
        {
          displayName: null,
          givenName: undefined,
          surname: null,
          scoredEmailAddresses: null,
          jobTitle: null,
          department: undefined,
          officeLocation: null,
        },
      ];
      vi.mocked(client.listRelevantPeople).mockResolvedValue(mockPeople as any);

      const result = await tools.listRelevantPeople({});

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.people[0]).toEqual({
        displayName: null,
        givenName: null,
        surname: null,
        emailAddresses: [],
        jobTitle: null,
        department: null,
        officeLocation: null,
      });
    });
  });

  describe('searchPeople', () => {
    it('returns search results', async () => {
      const mockPeople = [
        {
          displayName: 'Alice Smith',
          givenName: 'Alice',
          surname: 'Smith',
          scoredEmailAddresses: [{ address: 'alice@example.com' }],
          jobTitle: 'Engineer',
          department: 'Engineering',
          officeLocation: 'Building A',
        },
      ];
      vi.mocked(client.searchPeople).mockResolvedValue(mockPeople);

      const result = await tools.searchPeople({ query: 'Alice' });

      expect(client.searchPeople).toHaveBeenCalledWith('Alice', 25);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.people).toHaveLength(1);
      expect(parsed.people[0].displayName).toBe('Alice Smith');
    });

    it('passes custom limit', async () => {
      vi.mocked(client.searchPeople).mockResolvedValue([]);

      await tools.searchPeople({ query: 'test', limit: 10 });

      expect(client.searchPeople).toHaveBeenCalledWith('test', 10);
    });
  });

  describe('getManager', () => {
    it('returns mapped manager', async () => {
      const mockManager = {
        id: 'mgr-123',
        displayName: 'Manager Person',
        mail: 'manager@example.com',
        jobTitle: 'VP Engineering',
        department: 'Engineering',
        officeLocation: 'Floor 5',
      };
      vi.mocked(client.getManager).mockResolvedValue(mockManager);

      const result = await tools.getManager();

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.manager).toEqual({
        id: 'mgr-123',
        displayName: 'Manager Person',
        mail: 'manager@example.com',
        jobTitle: 'VP Engineering',
        department: 'Engineering',
        officeLocation: 'Floor 5',
      });
    });

    it('handles missing fields gracefully', async () => {
      vi.mocked(client.getManager).mockResolvedValue({ id: 'mgr-123' });

      const result = await tools.getManager();

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.manager.id).toBe('mgr-123');
      expect(parsed.manager.displayName).toBeNull();
      expect(parsed.manager.mail).toBeNull();
    });
  });

  describe('getDirectReports', () => {
    it('returns mapped direct reports', async () => {
      const mockReports = [
        {
          id: 'user-1',
          displayName: 'Report One',
          mail: 'report1@example.com',
          jobTitle: 'SWE',
          department: 'Engineering',
          officeLocation: 'Building B',
        },
        {
          id: 'user-2',
          displayName: 'Report Two',
          mail: 'report2@example.com',
          jobTitle: 'SWE II',
          department: 'Engineering',
          officeLocation: null,
        },
      ];
      vi.mocked(client.getDirectReports).mockResolvedValue(mockReports);

      const result = await tools.getDirectReports();

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.direct_reports).toHaveLength(2);
      expect(parsed.direct_reports[0].displayName).toBe('Report One');
      expect(parsed.direct_reports[1].officeLocation).toBeNull();
    });

    it('returns empty array when no direct reports', async () => {
      vi.mocked(client.getDirectReports).mockResolvedValue([]);

      const result = await tools.getDirectReports();

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.direct_reports).toEqual([]);
    });
  });

  describe('getUserProfile', () => {
    it('returns mapped user profile', async () => {
      const mockUser = {
        id: 'user-abc',
        displayName: 'Test User',
        mail: 'test@example.com',
        jobTitle: 'Developer',
        department: 'Engineering',
        officeLocation: 'Remote',
        mobilePhone: '+1-555-1234',
        businessPhones: ['+1-555-5678'],
      };
      vi.mocked(client.getUserProfile).mockResolvedValue(mockUser);

      const result = await tools.getUserProfile({ identifier: 'test@example.com' });

      expect(client.getUserProfile).toHaveBeenCalledWith('test@example.com');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.user).toEqual({
        id: 'user-abc',
        displayName: 'Test User',
        mail: 'test@example.com',
        jobTitle: 'Developer',
        department: 'Engineering',
        officeLocation: 'Remote',
        mobilePhone: '+1-555-1234',
        businessPhones: ['+1-555-5678'],
      });
    });

    it('handles missing optional fields', async () => {
      vi.mocked(client.getUserProfile).mockResolvedValue({
        id: 'user-abc',
        displayName: 'Test User',
      });

      const result = await tools.getUserProfile({ identifier: 'user-abc' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.user.mobilePhone).toBeNull();
      expect(parsed.user.businessPhones).toEqual([]);
    });
  });

  describe('getUserPhoto', () => {
    it('saves photo to disk and returns path and size', async () => {
      const photoBuffer = new ArrayBuffer(1024);
      vi.mocked(client.getUserPhoto).mockResolvedValue(photoBuffer);

      const result = await tools.getUserPhoto({
        identifier: 'test@example.com',
        save_path: '/tmp/test_photo.jpg',
      });

      expect(client.getUserPhoto).toHaveBeenCalledWith('test@example.com');
      expect(fs.writeFileSync).toHaveBeenCalled();
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.saved_to).toBe('/tmp/test_photo.jpg');
      expect(parsed.size).toBe(1024);
    });

    it('uses default path when save_path not provided', async () => {
      const photoBuffer = new ArrayBuffer(512);
      vi.mocked(client.getUserPhoto).mockResolvedValue(photoBuffer);

      const result = await tools.getUserPhoto({ identifier: 'user@example.com' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.saved_to).toContain('user@example.com_photo.jpg');
      expect(parsed.size).toBe(512);
    });

    it('returns error when photo fetch fails', async () => {
      vi.mocked(client.getUserPhoto).mockRejectedValue(new Error('No photo found'));

      const result = await tools.getUserPhoto({ identifier: 'nophoto@example.com' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBe('No photo found');
    });

    it('creates directory if it does not exist', async () => {
      vi.mocked(fs.existsSync).mockReturnValue(false);
      const photoBuffer = new ArrayBuffer(256);
      vi.mocked(client.getUserPhoto).mockResolvedValue(photoBuffer);

      await tools.getUserPhoto({
        identifier: 'test@example.com',
        save_path: '/tmp/newdir/photo.jpg',
      });

      expect(fs.mkdirSync).toHaveBeenCalledWith('/tmp/newdir', { recursive: true });
    });
  });

  describe('getUserPresence', () => {
    it('returns mapped presence', async () => {
      const mockPresence = {
        availability: 'Available',
        activity: 'Available',
        statusMessage: {
          message: { content: 'Working from home' },
        },
      };
      vi.mocked(client.getUserPresence).mockResolvedValue(mockPresence);

      const result = await tools.getUserPresence({ identifier: 'user@example.com' });

      expect(client.getUserPresence).toHaveBeenCalledWith('user@example.com');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.presence).toEqual({
        availability: 'Available',
        activity: 'Available',
        statusMessage: 'Working from home',
      });
    });

    it('handles null status message', async () => {
      const mockPresence = {
        availability: 'Busy',
        activity: 'InACall',
        statusMessage: null,
      };
      vi.mocked(client.getUserPresence).mockResolvedValue(mockPresence);

      const result = await tools.getUserPresence({ identifier: 'user@example.com' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.presence.statusMessage).toBeNull();
    });

    it('handles missing availability/activity', async () => {
      vi.mocked(client.getUserPresence).mockResolvedValue({});

      const result = await tools.getUserPresence({ identifier: 'user@example.com' });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.presence.availability).toBeNull();
      expect(parsed.presence.activity).toBeNull();
    });
  });

  describe('getUsersPresence', () => {
    it('returns mapped presences for multiple users', async () => {
      const mockPresences = [
        { id: 'user-1', availability: 'Available', activity: 'Available' },
        { id: 'user-2', availability: 'Busy', activity: 'InACall' },
        { id: 'user-3', availability: 'DoNotDisturb', activity: 'Presenting' },
      ];
      vi.mocked(client.getUsersPresence).mockResolvedValue(mockPresences);

      const result = await tools.getUsersPresence({ user_ids: ['user-1', 'user-2', 'user-3'] });

      expect(client.getUsersPresence).toHaveBeenCalledWith(['user-1', 'user-2', 'user-3']);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.presences).toHaveLength(3);
      expect(parsed.presences[0]).toEqual({ id: 'user-1', availability: 'Available', activity: 'Available' });
      expect(parsed.presences[1]).toEqual({ id: 'user-2', availability: 'Busy', activity: 'InACall' });
      expect(parsed.presences[2]).toEqual({ id: 'user-3', availability: 'DoNotDisturb', activity: 'Presenting' });
    });

    it('handles null fields', async () => {
      const mockPresences = [
        { id: null, availability: null, activity: null },
      ];
      vi.mocked(client.getUsersPresence).mockResolvedValue(mockPresences);

      const result = await tools.getUsersPresence({ user_ids: ['user-1'] });

      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.presences[0]).toEqual({ id: null, availability: null, activity: null });
    });
  });
});
