/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Aggregated tool definitions across all registry-migrated domains.
 *
 * This is the single list the server registers and the contract harness
 * iterates. As U2 migrates each domain, add its `*ToolDefinitions()` here and
 * the harness covers it automatically — no per-domain test wiring.
 */

import type { ToolDefinition } from './types.js';
import { mailRulesToolDefinitions } from '../tools/mail-rules.js';
import { categoriesToolDefinitions } from '../tools/categories.js';
import { focusedOverridesToolDefinitions } from '../tools/focused-overrides.js';
import { calendarPermissionsToolDefinitions } from '../tools/calendar-permissions.js';
import { checklistItemsToolDefinitions } from '../tools/checklist-items.js';
import { linkedResourcesToolDefinitions } from '../tools/linked-resources.js';
import { taskAttachmentsToolDefinitions } from '../tools/task-attachments.js';
import { peopleToolDefinitions } from '../tools/people.js';
import { plannerVisualizationToolDefinitions } from '../tools/planner-visualization.js';
import { meetingsToolDefinitions } from '../tools/meetings.js';
import { sharePointToolDefinitions } from '../tools/sharepoint.js';
import { teamsToolDefinitions } from '../tools/teams.js';
import { plannerToolDefinitions } from '../tools/planner.js';
import { oneDriveToolDefinitions } from '../tools/onedrive.js';
import { excelToolDefinitions } from '../tools/excel.js';
import { notesToolDefinitions } from '../tools/notes.js';
import { contactsToolDefinitions } from '../tools/contacts.js';
import { contactFoldersToolDefinitions } from '../tools/contact-folders.js';
import { calendarToolDefinitions } from '../tools/calendar.js';
import { tasksToolDefinitions } from '../tools/tasks.js';
import { taskListsToolDefinitions } from '../tools/task-lists.js';
import { mailboxOrganizationToolDefinitions } from '../tools/mailbox-organization.js';
import { mailToolDefinitions } from '../tools/mail.js';
import { mailSendToolDefinitions } from '../tools/mail-send.js';
import { schedulingToolDefinitions } from '../tools/scheduling.js';
import { mailboxSettingsToolDefinitions } from '../tools/mailbox-settings.js';
import { accountsToolDefinitions } from '../tools/accounts.js';

export function allToolDefinitions(): ToolDefinition[] {
  return [
    ...mailRulesToolDefinitions(),
    ...categoriesToolDefinitions(),
    ...focusedOverridesToolDefinitions(),
    ...calendarPermissionsToolDefinitions(),
    ...checklistItemsToolDefinitions(),
    ...linkedResourcesToolDefinitions(),
    ...taskAttachmentsToolDefinitions(),
    ...peopleToolDefinitions(),
    ...plannerVisualizationToolDefinitions(),
    ...meetingsToolDefinitions(),
    ...sharePointToolDefinitions(),
    ...teamsToolDefinitions(),
    ...plannerToolDefinitions(),
    ...oneDriveToolDefinitions(),
    ...excelToolDefinitions(),
    ...notesToolDefinitions(),
    ...contactsToolDefinitions(),
    ...contactFoldersToolDefinitions(),
    ...calendarToolDefinitions(),
    ...tasksToolDefinitions(),
    ...taskListsToolDefinitions(),
    ...mailboxOrganizationToolDefinitions(),
    ...mailToolDefinitions(),
    ...mailSendToolDefinitions(),
    ...schedulingToolDefinitions(),
    ...mailboxSettingsToolDefinitions(),
    ...accountsToolDefinitions(),
    // U2: append each migrated domain's definitions here.
  ];
}
