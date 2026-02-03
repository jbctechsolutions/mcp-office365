/**
 * Unit tests for AppleScript template generation.
 */

import { describe, it, expect } from 'vitest';
import { respondToEvent, deleteEvent } from '../../../src/applescript/scripts.js';

describe('respondToEvent', () => {
  it('should generate accept script with comment', () => {
    const script = respondToEvent({
      eventId: 123,
      response: 'accept',
      sendResponse: true,
      comment: 'I will be there',
    });

    expect(script).toContain('calendar event id 123');
    expect(script).toContain('accept');
    expect(script).toContain('I will be there');
  });

  it('should generate decline script without sending response', () => {
    const script = respondToEvent({
      eventId: 456,
      response: 'decline',
      sendResponse: false,
    });

    expect(script).toContain('calendar event id 456');
    expect(script).toContain('decline');
  });

  it('should generate tentative accept script', () => {
    const script = respondToEvent({
      eventId: 789,
      response: 'tentative',
      sendResponse: true,
    });

    expect(script).toContain('calendar event id 789');
    expect(script).toContain('tentative');
  });
});

describe('deleteEvent', () => {
  it('should generate script for single instance', () => {
    const script = deleteEvent({ eventId: 123, applyTo: 'this_instance' });
    expect(script).toContain('calendar event id 123');
    expect(script).toContain('delete');
    expect(script).toContain('Deleting single instance');
  });

  it('should generate script for all in series', () => {
    const script = deleteEvent({ eventId: 456, applyTo: 'all_in_series' });
    expect(script).toContain('calendar event id 456');
    expect(script).toContain('delete');
    expect(script).toContain('Deleting entire series');
  });

  it('should include success output format', () => {
    const script = deleteEvent({ eventId: 789, applyTo: 'this_instance' });
    expect(script).toContain('success{{=}}true');
    expect(script).toContain('eventId{{=}}');
  });
});
