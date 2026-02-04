/**
 * Calendar management operations using AppleScript.
 *
 * Provides high-level calendar operations including RSVP, deletion, and updates.
 */

import { executeAppleScriptOrThrow } from './executor.js';
import * as scripts from './scripts.js';
import { parseRespondToEventResult, parseDeleteEventResult, parseUpdateEventResult, type RespondToEventResult } from './parser.js';
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
   * Updates an event (single instance or entire recurring series).
   * All fields in updates are optional - only specified fields will be updated.
   */
  updateEvent(eventId: number, updates: EventUpdates, applyTo: ApplyToScope): UpdatedEvent {
    const scriptUpdates: scripts.UpdateEventParams['updates'] = {
      ...(updates.title != null && { title: updates.title }),
      ...(updates.location != null && { location: updates.location }),
      ...(updates.description != null && { description: updates.description }),
      ...(updates.startDate != null && { startDate: updates.startDate }),
      ...(updates.endDate != null && { endDate: updates.endDate }),
      ...(updates.isAllDay != null && { isAllDay: updates.isAllDay }),
    };

    const scriptParams: scripts.UpdateEventParams = {
      eventId,
      applyTo,
      updates: scriptUpdates,
    };

    const script = scripts.updateEvent(scriptParams);
    const output = executeAppleScriptOrThrow(script);
    const result = parseUpdateEventResult(output);

    if (result == null) {
      throw new AppleScriptError('Failed to parse update response');
    }

    if (!result.success) {
      throw new AppleScriptError(result.error ?? 'Update operation failed');
    }

    return {
      id: result.id!,
      updatedFields: result.updatedFields ?? [],
    };
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
