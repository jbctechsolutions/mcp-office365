/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Note-related type definitions.
 */

/**
 * Note summary for list views.
 */
export interface NoteSummary {
  readonly id: number;
  readonly folderId: number;
  readonly title: string | null;
  readonly preview: string | null;
  readonly modifiedDate: string | null;
}

/**
 * Full note details including body content.
 */
export interface Note extends NoteSummary {
  readonly body: string | null;
  readonly createdDate: string | null;
  readonly categories: readonly string[];
}
