/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, test, expect } from 'vitest';
import { createCalendarManager } from '../../../src/applescript/index.js';

const isOutlookAvailable = process.env.OUTLOOK_AVAILABLE === '1';
const testIf = (condition: boolean) => (condition ? test : test.skip);

describe('Event Management Integration', () => {
  const calendarManager = createCalendarManager();

  testIf(isOutlookAvailable)('responds to event', () => {
    expect(calendarManager.respondToEvent).toBeDefined();
  });

  testIf(isOutlookAvailable)('deletes event', () => {
    expect(calendarManager.deleteEvent).toBeDefined();
  });

  testIf(isOutlookAvailable)('updates event', () => {
    expect(calendarManager.updateEvent).toBeDefined();
  });
});
