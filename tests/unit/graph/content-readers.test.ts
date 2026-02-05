/**
 * Tests for Graph API content readers.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import {
  GRAPH_EMAIL_PATH_PREFIX,
  GRAPH_EVENT_PATH_PREFIX,
  GRAPH_CONTACT_PATH_PREFIX,
  GRAPH_TASK_PATH_PREFIX,
  GraphEmailContentReader,
  GraphEventContentReader,
  GraphContactContentReader,
  GraphTaskContentReader,
  GraphNoteContentReader,
  createGraphContentReaders,
  createGraphContentReadersWithClient,
} from '../../../src/graph/content-readers.js';
import { GraphClient } from '../../../src/graph/client/index.js';

// Mock the GraphClient
vi.mock('../../../src/graph/client/index.js', () => ({
  GraphClient: vi.fn().mockImplementation(function() {
    return {
      getMessage: vi.fn(),
      getEvent: vi.fn(),
      getContact: vi.fn(),
      getTask: vi.fn(),
    };
  }),
}));

describe('graph/content-readers', () => {
  describe('Path Prefixes', () => {
    it('has correct email path prefix', () => {
      expect(GRAPH_EMAIL_PATH_PREFIX).toBe('graph-email:');
    });

    it('has correct event path prefix', () => {
      expect(GRAPH_EVENT_PATH_PREFIX).toBe('graph-event:');
    });

    it('has correct contact path prefix', () => {
      expect(GRAPH_CONTACT_PATH_PREFIX).toBe('graph-contact:');
    });

    it('has correct task path prefix', () => {
      expect(GRAPH_TASK_PATH_PREFIX).toBe('graph-task:');
    });
  });

  describe('GraphEmailContentReader', () => {
    let client: any;
    let reader: GraphEmailContentReader;

    beforeEach(() => {
      client = {
        getMessage: vi.fn(),
      };
      reader = new GraphEmailContentReader(client);
    });

    describe('readEmailBody (sync)', () => {
      it('returns null (sync not supported)', () => {
        const result = reader.readEmailBody('graph-email:msg-123');
        expect(result).toBeNull();
      });
    });

    describe('readEmailBodyAsync', () => {
      it('returns email body content', async () => {
        client.getMessage.mockResolvedValue({
          body: { content: '<p>Email content</p>', contentType: 'html' },
        });

        const result = await reader.readEmailBodyAsync('graph-email:msg-123');

        expect(result).toBe('<p>Email content</p>');
        expect(client.getMessage).toHaveBeenCalledWith('msg-123');
      });

      it('returns plain text content', async () => {
        client.getMessage.mockResolvedValue({
          body: { content: 'Plain text email', contentType: 'text' },
        });

        const result = await reader.readEmailBodyAsync('graph-email:msg-123');

        expect(result).toBe('Plain text email');
      });

      it('returns null for null path', async () => {
        const result = await reader.readEmailBodyAsync(null);

        expect(result).toBeNull();
        expect(client.getMessage).not.toHaveBeenCalled();
      });

      it('returns null for invalid path prefix', async () => {
        const result = await reader.readEmailBodyAsync('invalid:msg-123');

        expect(result).toBeNull();
        expect(client.getMessage).not.toHaveBeenCalled();
      });

      it('returns null when message not found', async () => {
        client.getMessage.mockResolvedValue(null);

        const result = await reader.readEmailBodyAsync('graph-email:msg-123');

        expect(result).toBeNull();
      });

      it('returns null when body is missing', async () => {
        client.getMessage.mockResolvedValue({});

        const result = await reader.readEmailBodyAsync('graph-email:msg-123');

        expect(result).toBeNull();
      });

      it('returns null on API error', async () => {
        client.getMessage.mockRejectedValue(new Error('API error'));

        const result = await reader.readEmailBodyAsync('graph-email:msg-123');

        expect(result).toBeNull();
      });
    });
  });

  describe('GraphEventContentReader', () => {
    let client: any;
    let reader: GraphEventContentReader;

    beforeEach(() => {
      client = {
        getEvent: vi.fn(),
      };
      reader = new GraphEventContentReader(client);
    });

    describe('readEventDetails (sync)', () => {
      it('returns null (sync not supported)', () => {
        const result = reader.readEventDetails('graph-event:evt-123');
        expect(result).toBeNull();
      });
    });

    describe('readEventDetailsAsync', () => {
      it('returns event details', async () => {
        client.getEvent.mockResolvedValue({
          subject: 'Meeting',
          location: { displayName: 'Conference Room A' },
          body: { content: 'Meeting agenda...' },
          organizer: { emailAddress: { name: 'John Doe', address: 'john@example.com' } },
          attendees: [
            {
              emailAddress: { address: 'a@example.com', name: 'User A' },
              status: { response: 'accepted' },
            },
            {
              emailAddress: { address: 'b@example.com', name: 'User B' },
              status: { response: 'declined' },
            },
          ],
        });

        const result = await reader.readEventDetailsAsync('graph-event:evt-123');

        expect(result).toEqual({
          title: 'Meeting',
          location: 'Conference Room A',
          description: 'Meeting agenda...',
          organizer: 'John Doe',
          attendees: [
            { email: 'a@example.com', name: 'User A', status: 'accepted' },
            { email: 'b@example.com', name: 'User B', status: 'declined' },
          ],
        });
        expect(client.getEvent).toHaveBeenCalledWith('evt-123');
      });

      it('maps tentativelyAccepted to tentative', async () => {
        client.getEvent.mockResolvedValue({
          attendees: [
            {
              emailAddress: { address: 'a@example.com' },
              status: { response: 'tentativelyAccepted' },
            },
          ],
        });

        const result = await reader.readEventDetailsAsync('graph-event:evt-123');

        expect(result?.attendees[0].status).toBe('tentative');
      });

      it('maps unknown status to unknown', async () => {
        client.getEvent.mockResolvedValue({
          attendees: [
            {
              emailAddress: { address: 'a@example.com' },
              status: { response: 'someOtherStatus' },
            },
          ],
        });

        const result = await reader.readEventDetailsAsync('graph-event:evt-123');

        expect(result?.attendees[0].status).toBe('unknown');
      });

      it('uses email address when organizer name is missing', async () => {
        client.getEvent.mockResolvedValue({
          organizer: { emailAddress: { address: 'org@example.com' } },
        });

        const result = await reader.readEventDetailsAsync('graph-event:evt-123');

        expect(result?.organizer).toBe('org@example.com');
      });

      it('returns null for null path', async () => {
        const result = await reader.readEventDetailsAsync(null);

        expect(result).toBeNull();
        expect(client.getEvent).not.toHaveBeenCalled();
      });

      it('returns null for invalid path prefix', async () => {
        const result = await reader.readEventDetailsAsync('invalid:evt-123');

        expect(result).toBeNull();
      });

      it('returns null when event not found', async () => {
        client.getEvent.mockResolvedValue(null);

        const result = await reader.readEventDetailsAsync('graph-event:evt-123');

        expect(result).toBeNull();
      });

      it('returns null on API error', async () => {
        client.getEvent.mockRejectedValue(new Error('API error'));

        const result = await reader.readEventDetailsAsync('graph-event:evt-123');

        expect(result).toBeNull();
      });

      it('handles event without attendees', async () => {
        client.getEvent.mockResolvedValue({
          subject: 'Solo Event',
        });

        const result = await reader.readEventDetailsAsync('graph-event:evt-123');

        expect(result?.attendees).toEqual([]);
      });
    });
  });

  describe('GraphContactContentReader', () => {
    let client: any;
    let reader: GraphContactContentReader;

    beforeEach(() => {
      client = {
        getContact: vi.fn(),
      };
      reader = new GraphContactContentReader(client);
    });

    describe('readContactDetails (sync)', () => {
      it('returns null (sync not supported)', () => {
        const result = reader.readContactDetails('graph-contact:contact-123');
        expect(result).toBeNull();
      });
    });

    describe('readContactDetailsAsync', () => {
      it('returns contact details', async () => {
        client.getContact.mockResolvedValue({
          givenName: 'John',
          surname: 'Doe',
          middleName: 'M',
          nickName: 'Johnny',
          companyName: 'Acme Corp',
          jobTitle: 'Developer',
          department: 'Engineering',
          emailAddresses: [
            { name: 'work', address: 'john@acme.com' },
            { name: 'personal', address: 'john@gmail.com' },
          ],
          homePhones: ['555-1234'],
          businessPhones: ['555-5678'],
          mobilePhone: '555-9999',
          homeAddress: {
            street: '123 Home St',
            city: 'Homeville',
            state: 'HO',
            postalCode: '12345',
            countryOrRegion: 'USA',
          },
          businessAddress: {
            street: '456 Work Ave',
            city: 'Worktown',
            state: 'WO',
            postalCode: '67890',
            countryOrRegion: 'USA',
          },
          personalNotes: 'Some notes about John',
        });

        const result = await reader.readContactDetailsAsync('graph-contact:contact-123');

        expect(result).toEqual({
          firstName: 'John',
          lastName: 'Doe',
          middleName: 'M',
          nickname: 'Johnny',
          company: 'Acme Corp',
          jobTitle: 'Developer',
          department: 'Engineering',
          emails: [
            { type: 'work', address: 'john@acme.com' },
            { type: 'personal', address: 'john@gmail.com' },
          ],
          phones: [
            { type: 'home', number: '555-1234' },
            { type: 'work', number: '555-5678' },
            { type: 'mobile', number: '555-9999' },
          ],
          addresses: [
            {
              type: 'home',
              street: '123 Home St',
              city: 'Homeville',
              state: 'HO',
              postalCode: '12345',
              country: 'USA',
            },
            {
              type: 'work',
              street: '456 Work Ave',
              city: 'Worktown',
              state: 'WO',
              postalCode: '67890',
              country: 'USA',
            },
          ],
          notes: 'Some notes about John',
        });
      });

      it('returns null for null path', async () => {
        const result = await reader.readContactDetailsAsync(null);

        expect(result).toBeNull();
        expect(client.getContact).not.toHaveBeenCalled();
      });

      it('returns null for invalid path prefix', async () => {
        const result = await reader.readContactDetailsAsync('invalid:contact-123');

        expect(result).toBeNull();
      });

      it('returns null when contact not found', async () => {
        client.getContact.mockResolvedValue(null);

        const result = await reader.readContactDetailsAsync('graph-contact:contact-123');

        expect(result).toBeNull();
      });

      it('returns null on API error', async () => {
        client.getContact.mockRejectedValue(new Error('API error'));

        const result = await reader.readContactDetailsAsync('graph-contact:contact-123');

        expect(result).toBeNull();
      });

      it('handles contact with minimal data', async () => {
        client.getContact.mockResolvedValue({});

        const result = await reader.readContactDetailsAsync('graph-contact:contact-123');

        expect(result).toEqual({
          firstName: null,
          lastName: null,
          middleName: null,
          nickname: null,
          company: null,
          jobTitle: null,
          department: null,
          emails: [],
          phones: [],
          addresses: [],
          notes: null,
        });
      });

      it('filters out emails without address', async () => {
        client.getContact.mockResolvedValue({
          emailAddresses: [
            { name: 'work', address: 'john@acme.com' },
            { name: 'invalid' }, // no address
          ],
        });

        const result = await reader.readContactDetailsAsync('graph-contact:contact-123');

        expect(result?.emails).toHaveLength(1);
        expect(result?.emails[0].address).toBe('john@acme.com');
      });
    });
  });

  describe('GraphTaskContentReader', () => {
    let client: any;
    let reader: GraphTaskContentReader;

    beforeEach(() => {
      client = {
        getTask: vi.fn(),
      };
      reader = new GraphTaskContentReader(client);
    });

    describe('readTaskDetails (sync)', () => {
      it('returns null (sync not supported)', () => {
        const result = reader.readTaskDetails('graph-task:list-123:task-456');
        expect(result).toBeNull();
      });
    });

    describe('readTaskDetailsAsync', () => {
      it('returns task details', async () => {
        client.getTask.mockResolvedValue({
          body: { content: 'Task description' },
          completedDateTime: { dateTime: '2024-01-15T10:00:00' },
          reminderDateTime: { dateTime: '2024-01-14T09:00:00' },
        });

        const result = await reader.readTaskDetailsAsync('graph-task:list-123:task-456');

        expect(result).toEqual({
          body: 'Task description',
          completedDate: '2024-01-15T10:00:00',
          reminderDate: '2024-01-14T09:00:00',
          categories: [],
        });
        expect(client.getTask).toHaveBeenCalledWith('list-123', 'task-456');
      });

      it('returns null for null path', async () => {
        const result = await reader.readTaskDetailsAsync(null);

        expect(result).toBeNull();
        expect(client.getTask).not.toHaveBeenCalled();
      });

      it('returns null for invalid path prefix', async () => {
        const result = await reader.readTaskDetailsAsync('invalid:list:task');

        expect(result).toBeNull();
      });

      it('returns null for malformed task path (missing parts)', async () => {
        const result = await reader.readTaskDetailsAsync('graph-task:only-one-part');

        expect(result).toBeNull();
        expect(client.getTask).not.toHaveBeenCalled();
      });

      it('returns null for malformed task path (too many parts)', async () => {
        // The current implementation splits by ':' and expects exactly 2 parts after the prefix
        const result = await reader.readTaskDetailsAsync('graph-task:part1:part2:part3');

        expect(result).toBeNull();
        expect(client.getTask).not.toHaveBeenCalled();
      });

      it('returns null when task not found', async () => {
        client.getTask.mockResolvedValue(null);

        const result = await reader.readTaskDetailsAsync('graph-task:list-123:task-456');

        expect(result).toBeNull();
      });

      it('returns null on API error', async () => {
        client.getTask.mockRejectedValue(new Error('API error'));

        const result = await reader.readTaskDetailsAsync('graph-task:list-123:task-456');

        expect(result).toBeNull();
      });

      it('handles task with minimal data', async () => {
        client.getTask.mockResolvedValue({});

        const result = await reader.readTaskDetailsAsync('graph-task:list-123:task-456');

        expect(result).toEqual({
          body: null,
          completedDate: null,
          reminderDate: null,
          categories: [],
        });
      });
    });
  });

  describe('GraphNoteContentReader', () => {
    let reader: GraphNoteContentReader;

    beforeEach(() => {
      reader = new GraphNoteContentReader();
    });

    describe('readNoteDetails (sync)', () => {
      it('returns null (notes not supported by Graph API)', () => {
        const result = reader.readNoteDetails('some-path');
        expect(result).toBeNull();
      });
    });

    describe('readNoteDetailsAsync', () => {
      it('returns null (notes not supported by Graph API)', async () => {
        const result = await reader.readNoteDetailsAsync('some-path');
        expect(result).toBeNull();
      });
    });
  });

  describe('createGraphContentReaders', () => {
    it('creates all content readers', () => {
      const readers = createGraphContentReaders();

      expect(readers.email).toBeInstanceOf(GraphEmailContentReader);
      expect(readers.event).toBeInstanceOf(GraphEventContentReader);
      expect(readers.contact).toBeInstanceOf(GraphContactContentReader);
      expect(readers.task).toBeInstanceOf(GraphTaskContentReader);
      expect(readers.note).toBeInstanceOf(GraphNoteContentReader);
    });
  });

  describe('createGraphContentReadersWithClient', () => {
    it('creates all content readers with provided client', () => {
      const mockClient = new GraphClient() as any;
      const readers = createGraphContentReadersWithClient(mockClient);

      expect(readers.email).toBeInstanceOf(GraphEmailContentReader);
      expect(readers.event).toBeInstanceOf(GraphEventContentReader);
      expect(readers.contact).toBeInstanceOf(GraphContactContentReader);
      expect(readers.task).toBeInstanceOf(GraphTaskContentReader);
      expect(readers.note).toBeInstanceOf(GraphNoteContentReader);
    });
  });
});
