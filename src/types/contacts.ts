/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Contact-related type definitions.
 */

/**
 * Contact record types.
 */
export const ContactType = {
  Person: 0,
  DistributionList: 1,
} as const;

export type ContactTypeValue = (typeof ContactType)[keyof typeof ContactType];

/**
 * Contact summary for list views.
 */
export interface ContactSummary {
  readonly id: number;
  readonly folderId: number;
  readonly displayName: string | null;
  readonly sortName: string | null;
  readonly contactType: ContactTypeValue;
}

/**
 * Full contact details including all fields.
 */
export interface Contact extends ContactSummary {
  readonly firstName: string | null;
  readonly lastName: string | null;
  readonly middleName: string | null;
  readonly nickname: string | null;
  readonly company: string | null;
  readonly jobTitle: string | null;
  readonly department: string | null;
  readonly emails: readonly ContactEmail[];
  readonly phones: readonly ContactPhone[];
  readonly addresses: readonly ContactAddress[];
  readonly notes: string | null;
}

/**
 * Contact email address.
 */
export interface ContactEmail {
  readonly type: EmailType;
  readonly address: string;
}

/**
 * Email address types.
 */
export const EmailType = {
  Work: 'work',
  Home: 'home',
  Other: 'other',
} as const;

export type EmailType = (typeof EmailType)[keyof typeof EmailType];

/**
 * Contact phone number.
 */
export interface ContactPhone {
  readonly type: PhoneType;
  readonly number: string;
}

/**
 * Phone number types.
 */
export const PhoneType = {
  Work: 'work',
  Home: 'home',
  Mobile: 'mobile',
  Fax: 'fax',
  Other: 'other',
} as const;

export type PhoneType = (typeof PhoneType)[keyof typeof PhoneType];

/**
 * Contact postal address.
 */
export interface ContactAddress {
  readonly type: AddressType;
  readonly street: string | null;
  readonly city: string | null;
  readonly state: string | null;
  readonly postalCode: string | null;
  readonly country: string | null;
}

/**
 * Address types.
 */
export const AddressType = {
  Work: 'work',
  Home: 'home',
  Other: 'other',
} as const;

export type AddressType = (typeof AddressType)[keyof typeof AddressType];
