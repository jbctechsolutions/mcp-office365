/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { createTestDatabase, SAMPLE_COUNTS } from '../../fixtures/database.js';
import { createConnection, type IConnection } from '../../../src/database/connection.js';
import { createRepository, type IRepository } from '../../../src/database/repository.js';
import {
  ContactsTools,
  createContactsTools,
  ListContactsInput,
  SearchContactsInput,
  GetContactInput,
  type IContactContentReader,
  type ContactDetails,
} from '../../../src/tools/contacts.js';

describe('ContactsTools', () => {
  let testDb: { path: string; cleanup: () => void };
  let connection: IConnection;
  let repository: IRepository;
  let contactsTools: ContactsTools;

  beforeEach(() => {
    testDb = createTestDatabase();
    connection = createConnection(testDb.path);
    repository = createRepository(connection);
    contactsTools = createContactsTools(repository);
  });

  afterEach(() => {
    connection.close();
    testDb.cleanup();
  });

  // ---------------------------------------------------------------------------
  // Input Validation
  // ---------------------------------------------------------------------------

  describe('input validation', () => {
    it('validates ListContactsInput with defaults', () => {
      const parsed = ListContactsInput.parse({});
      expect(parsed.limit).toBe(50);
      expect(parsed.offset).toBe(0);
    });

    it('validates ListContactsInput with options', () => {
      const input = { limit: 25, offset: 10 };
      const parsed = ListContactsInput.parse(input);
      expect(parsed).toEqual(input);
    });

    it('validates SearchContactsInput', () => {
      const parsed = SearchContactsInput.parse({ query: 'john' });
      expect(parsed.query).toBe('john');
      expect(parsed.limit).toBe(50);
    });

    it('validates GetContactInput', () => {
      const parsed = GetContactInput.parse({ contact_id: 1 });
      expect(parsed.contact_id).toBe(1);
    });
  });

  // ---------------------------------------------------------------------------
  // listContacts
  // ---------------------------------------------------------------------------

  describe('listContacts', () => {
    it('returns contacts', () => {
      const contacts = contactsTools.listContacts({ limit: 50, offset: 0 });
      expect(contacts.length).toBe(SAMPLE_COUNTS.contacts);
    });

    it('returns contacts with correct structure', () => {
      const contacts = contactsTools.listContacts({ limit: 1, offset: 0 });
      const contact = contacts[0];

      expect(contact).toHaveProperty('id');
      expect(contact).toHaveProperty('folderId');
      expect(contact).toHaveProperty('displayName');
      expect(contact).toHaveProperty('sortName');
      expect(contact).toHaveProperty('contactType');
    });

    it('respects limit parameter', () => {
      const contacts = contactsTools.listContacts({ limit: 1, offset: 0 });
      expect(contacts.length).toBe(1);
    });

    it('respects offset parameter', () => {
      const allContacts = contactsTools.listContacts({ limit: 50, offset: 0 });
      const offsetContacts = contactsTools.listContacts({ limit: 50, offset: 1 });
      expect(offsetContacts.length).toBe(allContacts.length - 1);
    });

    it('returns contacts sorted by sortName', () => {
      const contacts = contactsTools.listContacts({ limit: 50, offset: 0 });
      const sortNames = contacts.map((c) => c.sortName);
      const sorted = [...sortNames].sort();
      expect(sortNames).toEqual(sorted);
    });
  });

  // ---------------------------------------------------------------------------
  // searchContacts
  // ---------------------------------------------------------------------------

  describe('searchContacts', () => {
    it('finds contacts by name', () => {
      const contacts = contactsTools.searchContacts({ query: 'John', limit: 50 });
      expect(contacts.length).toBeGreaterThan(0);
    });

    it('returns empty array for no matches', () => {
      const contacts = contactsTools.searchContacts({ query: 'xyznonexistent', limit: 50 });
      expect(contacts.length).toBe(0);
    });
  });

  // ---------------------------------------------------------------------------
  // getContact
  // ---------------------------------------------------------------------------

  describe('getContact', () => {
    it('returns contact by ID', () => {
      const contacts = contactsTools.listContacts({ limit: 1, offset: 0 });
      const firstContact = contacts[0];

      if (firstContact) {
        const contact = contactsTools.getContact({ contact_id: firstContact.id });
        expect(contact).not.toBeNull();
        expect(contact?.id).toBe(firstContact.id);
      }
    });

    it('returns null for non-existent ID', () => {
      const contact = contactsTools.getContact({ contact_id: 99999 });
      expect(contact).toBeNull();
    });

    it('includes additional fields in full contact', () => {
      const contacts = contactsTools.listContacts({ limit: 1, offset: 0 });
      const firstContact = contacts[0];

      if (firstContact) {
        const contact = contactsTools.getContact({ contact_id: firstContact.id });
        expect(contact).toHaveProperty('firstName');
        expect(contact).toHaveProperty('lastName');
        expect(contact).toHaveProperty('emails');
        expect(contact).toHaveProperty('phones');
        expect(contact).toHaveProperty('addresses');
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Content Reader Integration
  // ---------------------------------------------------------------------------

  describe('content reader integration', () => {
    it('uses content reader for contact details', () => {
      const mockDetails: ContactDetails = {
        firstName: 'John',
        lastName: 'Doe',
        middleName: null,
        nickname: 'JD',
        company: 'Acme Corp',
        jobTitle: 'Engineer',
        department: 'R&D',
        emails: [{ type: 'work', address: 'john@acme.com' }],
        phones: [{ type: 'mobile', number: '555-1234' }],
        addresses: [
          {
            type: 'work',
            street: '123 Main St',
            city: 'New York',
            state: 'NY',
            postalCode: '10001',
            country: 'USA',
          },
        ],
        notes: 'Good contact',
      };

      const mockContentReader: IContactContentReader = {
        readContactDetails: () => mockDetails,
      };

      const toolsWithReader = createContactsTools(repository, mockContentReader);
      const contacts = toolsWithReader.listContacts({ limit: 1, offset: 0 });

      if (contacts[0]) {
        const contact = toolsWithReader.getContact({ contact_id: contacts[0].id });
        expect(contact?.firstName).toBe('John');
        expect(contact?.lastName).toBe('Doe');
        expect(contact?.company).toBe('Acme Corp');
        expect(contact?.emails).toHaveLength(1);
        expect(contact?.phones).toHaveLength(1);
        expect(contact?.addresses).toHaveLength(1);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Factory Function
  // ---------------------------------------------------------------------------

  describe('createContactsTools', () => {
    it('creates a ContactsTools instance', () => {
      const tools = createContactsTools(repository);
      expect(tools).toBeInstanceOf(ContactsTools);
    });
  });
});
