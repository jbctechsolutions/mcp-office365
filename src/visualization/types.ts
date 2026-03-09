/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Shared types for Planner visualization renderers.
 */

export type VisualizationFormat = 'html' | 'svg' | 'markdown' | 'mermaid';

export interface PlanVisualizationData {
  plan: {
    id: number;
    title: string;
    owner: string;
    createdDateTime: string;
  };
  buckets: Array<{
    id: number;
    name: string;
  }>;
  tasks: Array<{
    id: number;
    title: string;
    bucketId: number | null;
    assignees: string[];
    percentComplete: number;
    priority: number;
    startDateTime: string;
    dueDateTime: string;
    createdDateTime: string;
  }>;
}
