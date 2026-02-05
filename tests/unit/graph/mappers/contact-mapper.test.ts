/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Graph contact mapper functions.
 */

import { describe, it, expect } from 'vitest';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { mapContactToContactRow } from '../../../../src/graph/mappers/contact-mapper.js';
import { hashStringToNumber } from '../../../../src/graph/mappers/utils.js';

describe('graph/mappers/contact-mapper', () => {
  describe('mapContactToContactRow', () => {
    it('maps contact with all fields', () => {
      const contact: MicrosoftGraph.Contact = {
        id: 'contact-123',
        displayName: 'John Doe',
        surname: 'Doe',
      };

      const result = mapContactToContactRow(contact);

      expect(result.id).toBe(hashStringToNumber('contact-123'));
      expect(result.displayName).toBe('John Doe');
      expect(result.sortName).toBe('Doe');
      expect(result.folderId).toBe(0);
      expect(result.contactType).toBeNull();
      expect(result.dataFilePath).toBe('graph-contact:contact-123');
    });

    it('handles contact with null id', () => {
      const contact: MicrosoftGraph.Contact = {
        id: undefined,
        displayName: 'Test Contact',
      };

      const result = mapContactToContactRow(contact);

      expect(result.id).toBe(hashStringToNumber(''));
      expect(result.dataFilePath).toBe('graph-contact:');
    });

    it('handles contact with null displayName', () => {
      const contact: MicrosoftGraph.Contact = {
        id: 'contact-123',
        displayName: undefined,
      };

      const result = mapContactToContactRow(contact);

      expect(result.displayName).toBeNull();
    });

    it('uses surname as sortName when available', () => {
      const contact: MicrosoftGraph.Contact = {
        id: 'contact-123',
        displayName: 'John Doe',
        surname: 'Doe',
      };

      const result = mapContactToContactRow(contact);

      expect(result.sortName).toBe('Doe');
    });

    it('falls back to displayName when surname is null', () => {
      const contact: MicrosoftGraph.Contact = {
        id: 'contact-123',
        displayName: 'John Doe',
        surname: undefined,
      };

      const result = mapContactToContactRow(contact);

      expect(result.sortName).toBe('John Doe');
    });

    it('handles null sortName when both surname and displayName are null', () => {
      const contact: MicrosoftGraph.Contact = {
        id: 'contact-123',
        displayName: undefined,
        surname: undefined,
      };

      const result = mapContactToContactRow(contact);

      expect(result.sortName).toBeNull();
    });

    it('always sets folderId to 0', () => {
      const contact: MicrosoftGraph.Contact = {
        id: 'contact-123',
      };

      const result = mapContactToContactRow(contact);

      expect(result.folderId).toBe(0);
    });

    it('always sets contactType to null', () => {
      const contact: MicrosoftGraph.Contact = {
        id: 'contact-123',
      };

      const result = mapContactToContactRow(contact);

      expect(result.contactType).toBeNull();
    });
  });
});
