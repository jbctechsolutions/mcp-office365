/**
 * Maps Microsoft Graph Contact type to ContactRow.
 */

import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import type { ContactRow } from '../../database/repository.js';
import { hashStringToNumber, createGraphContentPath } from './utils.js';

/**
 * Maps a Graph Contact to a ContactRow.
 */
export function mapContactToContactRow(contact: MicrosoftGraph.Contact): ContactRow {
  const contactId = contact.id ?? '';

  return {
    id: hashStringToNumber(contactId),
    folderId: 0, // Graph contacts don't have folders in the same way
    displayName: contact.displayName ?? null,
    sortName: contact.surname ?? contact.displayName ?? null,
    contactType: null, // Not available in Graph
    dataFilePath: createGraphContentPath('contact', contactId),
  };
}
