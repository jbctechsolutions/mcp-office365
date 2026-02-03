/**
 * Calendar management operations using AppleScript.
 *
 * Provides high-level calendar operations including RSVP, deletion, and updates.
 */

import { executeAppleScriptOrThrow } from './executor.js';
import * as scripts from './scripts.js';
import { parseRespondToEventResult, parseDeleteEventResult, type RespondToEventResult } from './parser.js';
import { AppleScriptError } from '../utils/errors.js';

// =============================================================================
// Types
// =============================================================================

export type ResponseType = 'accept' | 'decline' | 'tentative';
export type ApplyToScope = 'this_instance' | 'all_in_series';

export interface ICalendarManager {
  respondToEvent(eventId: number, response: ResponseType, sendResponse: boolean, comment?: string): RespondToEventResult;
  deleteEvent(eventId: number, applyTo: ApplyToScope): void;
  updateEvent(eventId: number, updates: EventUpdates, applyTo: ApplyToScope): UpdatedEvent;
}

export interface EventUpdates {
  readonly title?: string;
  readonly startDate?: string;
  readonly endDate?: string;
  readonly location?: string;
  readonly description?: string;
  readonly isAllDay?: boolean;
  readonly recurrence?: scripts.RecurrenceScriptParams;
}

export interface UpdatedEvent {
  readonly id: number;
  readonly updatedFields: readonly string[];
}

// =============================================================================
// Implementation
// =============================================================================

export class AppleScriptCalendarManager implements ICalendarManager {
  /**
   * Responds to an event invitation (accept, decline, or tentative).
   */
  respondToEvent(eventId: number, response: ResponseType, sendResponse: boolean, comment?: string): RespondToEventResult {
    const params: scripts.RespondToEventParams = comment != null
      ? { eventId, response, sendResponse, comment }
      : { eventId, response, sendResponse };
    const script = scripts.respondToEvent(params);
    const output = executeAppleScriptOrThrow(script);
    const result = parseRespondToEventResult(output);

    if (result == null) {
      throw new AppleScriptError('Failed to parse RSVP response');
    }

    if (!result.success) {
      throw new AppleScriptError(result.error ?? 'RSVP operation failed');
    }

    return result;
  }

  /**
   * Deletes an event (single instance or entire recurring series).
   */
  deleteEvent(eventId: number, applyTo: ApplyToScope): void {
    const script = scripts.deleteEvent({ eventId, applyTo });
    const output = executeAppleScriptOrThrow(script);
    const result = parseDeleteEventResult(output);

    if (result == null) {
      throw new AppleScriptError('Failed to parse delete response');
    }

    if (!result.success) {
      throw new AppleScriptError(result.error ?? 'Delete operation failed');
    }
  }

  /**
   * Updates an event (to be implemented in future tasks).
   */
  updateEvent(_eventId: number, _updates: EventUpdates, _applyTo: ApplyToScope): UpdatedEvent {
    throw new Error('Not yet implemented');
  }
}

// =============================================================================
// Factory
// =============================================================================

/**
 * Creates a new calendar manager instance.
 */
export function createCalendarManager(): ICalendarManager {
  return new AppleScriptCalendarManager();
}
