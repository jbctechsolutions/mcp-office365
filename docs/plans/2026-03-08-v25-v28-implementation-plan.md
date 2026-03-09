# v2.5–v2.8 Expansion Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Expand the M365 MCP server from 189 to ~225 tools with Planner visualization, Teams enhancements, OneDrive/SharePoint/Excel, and internal batch optimization. Also rewrite the README.

**Architecture:** Follows existing 4-layer pattern (GraphClient → Repository → Tool class → index.ts registration). New `src/visualization/` directory for format renderers. Dual-backend (AppleScript + Graph API) with `GRAPH_ONLY_TOOL_NAMES` gating.

**Tech Stack:** TypeScript, Zod (z.strictObject), Vitest, Microsoft Graph SDK, MCP SDK

---

## Phase 7: README Update (v2.4.1)

### Task 1: Rewrite README

**Files:**
- Modify: `README.md`
- Modify: `package.json` (bump to 2.4.1)

**Step 1: Read current README and all tool definitions**

Read `README.md` (689 lines) and scan `src/index.ts` TOOLS array (starts ~line 255) to catalogue every tool by category. Count tools per category.

**Step 2: Write the new README**

Structure:
1. Header with badges: npm version, license, node version, tool count (~189)
2. One-liner: "MCP server for Microsoft 365 on macOS — mail, calendar, contacts, tasks, teams, people, and planner"
3. Features overview table (category + count + brief description)
4. Quick Start: install (`npx @jbctechsolutions/mcp-office365-mac`), auth modes (AppleScript default, `USE_GRAPH_API=1` for Graph), Claude Desktop JSON config example
5. Authentication section: device code flow, Azure AD app registration, required scopes per feature
6. Tool Reference: collapsible `<details>` sections per category, each listing every tool with one-line description
7. Architecture: text description of dual-backend, two-phase approval, ID caching, ETag caching
8. Required Graph API Permissions: table of scope → feature area
9. Contributing, License

Categories for tool reference:
| Category | Approx Count |
|---|---|
| Mail — Reading | 10 |
| Mail — Sending & Drafts | 16 |
| Mail — Organization | 24 |
| Mail — Rules | 4 |
| Mail — Categories | 4 |
| Mail — Focused Inbox | 4 |
| Mail — Settings & Auto-Replies | 4 |
| Mail — Tips & Headers | 3 |
| Attachments | 2 |
| Calendar — Events | 11 |
| Calendar — Groups | 2 |
| Calendar — Permissions | 4 |
| Calendar — Rooms | 2 |
| Contacts & Folders | 13 |
| Tasks & Task Lists | 11 |
| Checklist Items | 5 |
| Linked Resources | 4 |
| Task Attachments | 4 |
| Notes (AppleScript only) | 3 |
| Scheduling | 2 |
| Teams — Channels | 8 |
| Teams — Channel Messages | 6 |
| Teams — Chats | 6 |
| People & Presence | 8 |
| Planner | 17 |
| Accounts | 1 |

**Step 3: Bump version to 2.4.1 in package.json**

**Step 4: Update CHANGELOG.md**

Add v2.4.1 entry: "docs: comprehensive README rewrite with full tool reference"

**Step 5: Run tests to ensure nothing broke**

Run: `npm test`
Expected: All 1717+ tests pass

**Step 6: Commit**

```bash
git add README.md package.json CHANGELOG.md
git commit -m "docs: comprehensive README rewrite with full tool reference"
```

**Step 7: Tag and release v2.4.1**

```bash
git tag v2.4.1
git push origin main --tags
gh release create v2.4.1 --title "v2.4.1" --notes "Comprehensive README rewrite with full tool reference for all 189 tools."
```

---

## Phase 8: Planner Visualization (v2.5.0)

### Task 2: Visualization renderers — Markdown & Mermaid

**Files:**
- Create: `src/visualization/types.ts`
- Create: `src/visualization/markdown.ts`
- Create: `src/visualization/mermaid.ts`
- Create: `tests/unit/visualization/markdown.test.ts`
- Create: `tests/unit/visualization/mermaid.test.ts`

**Step 1: Write the failing tests for types and markdown renderer**

`tests/unit/visualization/markdown.test.ts`:
```typescript
import { describe, it, expect } from 'vitest';
import { renderKanbanMarkdown, renderGanttMarkdown, renderSummaryMarkdown, renderBurndownMarkdown } from '../../../src/visualization/markdown.js';
import type { PlanVisualizationData } from '../../../src/visualization/types.js';

const mockData: PlanVisualizationData = {
  plan: { id: 1, title: 'Test Plan' },
  buckets: [
    { id: 10, name: 'To Do', orderHint: '1' },
    { id: 20, name: 'In Progress', orderHint: '2' },
    { id: 30, name: 'Done', orderHint: '3' },
  ],
  tasks: [
    { id: 100, title: 'Task A', bucketId: 10, percentComplete: 0, priority: 5, startDateTime: '2026-03-01', dueDateTime: '2026-03-10', assignments: ['Alice'] },
    { id: 101, title: 'Task B', bucketId: 20, percentComplete: 50, priority: 3, startDateTime: '2026-03-02', dueDateTime: '2026-03-08', assignments: ['Bob'] },
    { id: 102, title: 'Task C', bucketId: 30, percentComplete: 100, priority: 1, startDateTime: '2026-03-01', dueDateTime: '2026-03-05', assignments: ['Alice'], completedDateTime: '2026-03-04' },
  ],
};

describe('renderKanbanMarkdown', () => {
  it('renders tasks grouped by bucket as markdown tables', () => {
    const result = renderKanbanMarkdown(mockData);
    expect(result).toContain('## To Do');
    expect(result).toContain('## In Progress');
    expect(result).toContain('## Done');
    expect(result).toContain('Task A');
    expect(result).toContain('Task B');
    expect(result).toContain('Task C');
  });

  it('shows priority and assignees', () => {
    const result = renderKanbanMarkdown(mockData);
    expect(result).toContain('Alice');
    expect(result).toContain('Bob');
  });
});

describe('renderGanttMarkdown', () => {
  it('renders a mermaid gantt code block', () => {
    const result = renderGanttMarkdown(mockData);
    expect(result).toContain('```mermaid');
    expect(result).toContain('gantt');
    expect(result).toContain('Task A');
  });
});

describe('renderSummaryMarkdown', () => {
  it('renders task counts by status', () => {
    const result = renderSummaryMarkdown(mockData);
    expect(result).toContain('Total');
    expect(result).toContain('3');
  });

  it('renders assignee workload', () => {
    const result = renderSummaryMarkdown(mockData);
    expect(result).toContain('Alice');
    expect(result).toContain('2');
  });
});

describe('renderBurndownMarkdown', () => {
  it('renders a text-based burndown table', () => {
    const result = renderBurndownMarkdown(mockData);
    expect(result).toContain('Date');
    expect(result).toContain('Remaining');
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/visualization/markdown.test.ts`
Expected: FAIL — modules not found

**Step 3: Create types file**

`src/visualization/types.ts`:
```typescript
export interface PlanVisualizationData {
  plan: { id: number; title: string };
  buckets: Array<{ id: number; name: string; orderHint: string }>;
  tasks: Array<{
    id: number;
    title: string;
    bucketId: number;
    percentComplete: number;
    priority: number;
    startDateTime?: string | null;
    dueDateTime?: string | null;
    assignments: string[];
    completedDateTime?: string | null;
  }>;
}

export type VisualizationFormat = 'html' | 'svg' | 'markdown' | 'mermaid';
```

**Step 4: Implement markdown renderer**

`src/visualization/markdown.ts`:
- `renderKanbanMarkdown(data)`: Group tasks by bucket, render as tables with columns: Title, Priority, Assignees, % Complete, Due Date
- `renderGanttMarkdown(data)`: Wrap a Mermaid gantt diagram in a code block
- `renderSummaryMarkdown(data)`: Stats table (total, not started, in progress, completed, overdue) + assignee workload table
- `renderBurndownMarkdown(data)`: Date-indexed table showing remaining task count over time

**Step 5: Run markdown tests to verify they pass**

Run: `npx vitest run tests/unit/visualization/markdown.test.ts`
Expected: PASS

**Step 6: Write failing mermaid tests**

`tests/unit/visualization/mermaid.test.ts`:
```typescript
import { describe, it, expect } from 'vitest';
import { renderKanbanMermaid, renderGanttMermaid, renderSummaryMermaid, renderBurndownMermaid } from '../../../src/visualization/mermaid.js';
import type { PlanVisualizationData } from '../../../src/visualization/types.js';

// Use same mockData as markdown tests

describe('renderGanttMermaid', () => {
  it('renders valid mermaid gantt syntax', () => {
    const result = renderGanttMermaid(mockData);
    expect(result).toMatch(/^gantt\n/);
    expect(result).toContain('dateFormat YYYY-MM-DD');
    expect(result).toContain('Task A');
  });

  it('groups tasks by bucket section', () => {
    const result = renderGanttMermaid(mockData);
    expect(result).toContain('section To Do');
    expect(result).toContain('section In Progress');
  });
});

describe('renderKanbanMermaid', () => {
  it('renders a text-based kanban (mermaid has no native kanban)', () => {
    const result = renderKanbanMermaid(mockData);
    expect(typeof result).toBe('string');
    expect(result.length).toBeGreaterThan(0);
  });
});
```

**Step 7: Implement mermaid renderer**

`src/visualization/mermaid.ts`:
- `renderGanttMermaid(data)`: Pure Mermaid `gantt` syntax with sections per bucket
- `renderKanbanMermaid(data)`: Mermaid `block-beta` or flowchart-based columns (best approximation)
- `renderSummaryMermaid(data)`: Mermaid `pie` chart of task status distribution
- `renderBurndownMermaid(data)`: Mermaid `xychart-beta` line chart

**Step 8: Run mermaid tests to verify they pass**

Run: `npx vitest run tests/unit/visualization/mermaid.test.ts`
Expected: PASS

**Step 9: Commit**

```bash
git add src/visualization/ tests/unit/visualization/
git commit -m "feat: add markdown and mermaid visualization renderers"
```

---

### Task 3: Visualization renderers — HTML & SVG

**Files:**
- Create: `src/visualization/html.ts`
- Create: `src/visualization/svg.ts`
- Create: `tests/unit/visualization/html.test.ts`
- Create: `tests/unit/visualization/svg.test.ts`

**Step 1: Write failing HTML renderer tests**

`tests/unit/visualization/html.test.ts`:
```typescript
import { describe, it, expect } from 'vitest';
import { renderKanbanHtml, renderGanttHtml, renderSummaryHtml, renderBurndownHtml } from '../../../src/visualization/html.js';
// Use same mockData pattern

describe('renderKanbanHtml', () => {
  it('renders a self-contained HTML document', () => {
    const result = renderKanbanHtml(mockData);
    expect(result).toContain('<!DOCTYPE html>');
    expect(result).toContain('</html>');
  });

  it('contains bucket columns', () => {
    const result = renderKanbanHtml(mockData);
    expect(result).toContain('To Do');
    expect(result).toContain('In Progress');
    expect(result).toContain('Done');
  });

  it('contains task cards with priority colors', () => {
    const result = renderKanbanHtml(mockData);
    expect(result).toContain('Task A');
    expect(result).toContain('Task B');
  });

  it('includes inline CSS (no external dependencies)', () => {
    const result = renderKanbanHtml(mockData);
    expect(result).toContain('<style>');
    expect(result).not.toContain('<link rel="stylesheet"');
  });
});

describe('renderGanttHtml', () => {
  it('renders a self-contained HTML with inline JS for timeline', () => {
    const result = renderGanttHtml(mockData);
    expect(result).toContain('<!DOCTYPE html>');
    expect(result).toContain('<script>');
    expect(result).toContain('Task A');
  });

  it('renders task bars with date ranges', () => {
    const result = renderGanttHtml(mockData);
    expect(result).toContain('2026-03-01');
    expect(result).toContain('2026-03-10');
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/visualization/html.test.ts`
Expected: FAIL

**Step 3: Implement HTML renderer**

`src/visualization/html.ts`:
- `renderKanbanHtml(data)`: Self-contained HTML with CSS Grid columns per bucket, card elements per task, priority color coding (urgent=red, important=orange, medium=yellow, low=green), hover tooltips
- `renderGanttHtml(data)`: Self-contained HTML with inline JS that draws timeline bars using CSS absolute positioning, date axis, task labels, progress indicators
- `renderSummaryHtml(data)`: Dashboard with stat cards (total, status breakdown, overdue), assignee table, priority distribution bar
- `renderBurndownHtml(data)`: Line chart using inline SVG within HTML (no external JS libs)

**Step 4: Run HTML tests to verify they pass**

**Step 5: Write failing SVG renderer tests**

`tests/unit/visualization/svg.test.ts`:
```typescript
import { describe, it, expect } from 'vitest';
import { renderKanbanSvg, renderGanttSvg, renderSummarySvg, renderBurndownSvg } from '../../../src/visualization/svg.js';

describe('renderKanbanSvg', () => {
  it('renders valid SVG', () => {
    const result = renderKanbanSvg(mockData);
    expect(result).toContain('<svg');
    expect(result).toContain('</svg>');
  });

  it('contains bucket labels', () => {
    const result = renderKanbanSvg(mockData);
    expect(result).toContain('To Do');
    expect(result).toContain('In Progress');
  });
});

describe('renderGanttSvg', () => {
  it('renders valid SVG with timeline bars', () => {
    const result = renderGanttSvg(mockData);
    expect(result).toContain('<svg');
    expect(result).toContain('<rect');
    expect(result).toContain('Task A');
  });
});
```

**Step 6: Implement SVG renderer**

`src/visualization/svg.ts`:
- `renderKanbanSvg(data)`: SVG with `<rect>` columns, `<text>` labels, `<rect>` task cards with rounded corners
- `renderGanttSvg(data)`: SVG with horizontal bars, date axis, grid lines, task labels
- `renderSummarySvg(data)`: SVG pie chart + stat text
- `renderBurndownSvg(data)`: SVG line chart with axis labels

**Step 7: Run all visualization tests**

Run: `npx vitest run tests/unit/visualization/`
Expected: All PASS

**Step 8: Commit**

```bash
git add src/visualization/ tests/unit/visualization/
git commit -m "feat: add HTML and SVG visualization renderers"
```

---

### Task 4: Planner visualization tools

**Files:**
- Create: `src/tools/planner-visualization.ts`
- Create: `tests/unit/tools/planner-visualization.test.ts`
- Modify: `src/index.ts` — add tool definitions, GRAPH_ONLY_TOOL_NAMES entries, handler dispatch
- Modify: `src/graph/repository.ts` — add methods that aggregate plan+buckets+tasks for visualization

**Step 1: Write failing tests for PlannerVisualizationTools**

`tests/unit/tools/planner-visualization.test.ts`:
```typescript
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { PlannerVisualizationTools, type IPlannerVisualizationRepository } from '../../../src/tools/planner-visualization.js';
import type { PlanVisualizationData } from '../../../src/visualization/types.js';

describe('PlannerVisualizationTools', () => {
  let repo: IPlannerVisualizationRepository;
  let tools: PlannerVisualizationTools;

  const mockVisualizationData: PlanVisualizationData = {
    plan: { id: 1, title: 'Sprint 1' },
    buckets: [{ id: 10, name: 'To Do', orderHint: '1' }],
    tasks: [
      { id: 100, title: 'Task A', bucketId: 10, percentComplete: 0, priority: 5, startDateTime: '2026-03-01', dueDateTime: '2026-03-10', assignments: ['Alice'] },
    ],
  };

  beforeEach(() => {
    repo = {
      getPlanVisualizationDataAsync: vi.fn().mockResolvedValue(mockVisualizationData),
    };
    tools = new PlannerVisualizationTools(repo);
  });

  describe('generateKanbanBoard', () => {
    it('returns file_path and preview for html format', async () => {
      const result = await tools.generateKanbanBoard({ plan_id: 1, format: 'html' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.file_path).toContain('.html');
      expect(parsed.format).toBe('html');
      expect(parsed.preview).toBeDefined();
    });

    it('returns markdown content for markdown format', async () => {
      const result = await tools.generateKanbanBoard({ plan_id: 1, format: 'markdown' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.format).toBe('markdown');
    });

    it('defaults to html format', async () => {
      const result = await tools.generateKanbanBoard({ plan_id: 1 });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.format).toBe('html');
    });
  });

  describe('generateGanttChart', () => {
    it('returns file_path for svg format', async () => {
      const result = await tools.generateGanttChart({ plan_id: 1, format: 'svg' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.file_path).toContain('.svg');
      expect(parsed.format).toBe('svg');
    });
  });

  describe('generatePlanSummary', () => {
    it('returns summary stats', async () => {
      const result = await tools.generatePlanSummary({ plan_id: 1, format: 'markdown' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.format).toBe('markdown');
    });
  });

  describe('generateBurndownChart', () => {
    it('returns burndown visualization', async () => {
      const result = await tools.generateBurndownChart({ plan_id: 1, format: 'html' });
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.file_path).toContain('.html');
    });
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/tools/planner-visualization.test.ts`
Expected: FAIL

**Step 3: Implement PlannerVisualizationTools**

`src/tools/planner-visualization.ts`:
```typescript
import { z } from 'zod';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import type { PlanVisualizationData, VisualizationFormat } from '../visualization/types.js';
import { renderKanbanMarkdown, renderGanttMarkdown, renderSummaryMarkdown, renderBurndownMarkdown } from '../visualization/markdown.js';
import { renderKanbanMermaid, renderGanttMermaid, renderSummaryMermaid, renderBurndownMermaid } from '../visualization/mermaid.js';
import { renderKanbanHtml, renderGanttHtml, renderSummaryHtml, renderBurndownHtml } from '../visualization/html.js';
import { renderKanbanSvg, renderGanttSvg, renderSummarySvg, renderBurndownSvg } from '../visualization/svg.js';

export const GenerateKanbanBoardInput = z.strictObject({
  plan_id: z.number().describe('Numeric plan ID'),
  format: z.enum(['html', 'svg', 'markdown', 'mermaid']).default('html').describe('Output format'),
  output_path: z.string().optional().describe('File path to save output'),
});

// Similar schemas for gantt, summary, burndown...

export interface IPlannerVisualizationRepository {
  getPlanVisualizationDataAsync(planId: number): Promise<PlanVisualizationData>;
}

export class PlannerVisualizationTools {
  constructor(private readonly repo: IPlannerVisualizationRepository) {}

  async generateKanbanBoard(params: { plan_id: number; format?: VisualizationFormat; output_path?: string }) {
    const format = params.format ?? 'html';
    const data = await this.repo.getPlanVisualizationDataAsync(params.plan_id);
    const content = this.renderKanban(data, format);
    return this.writeAndReturn(content, format, 'kanban', params.output_path, data);
  }
  // ... generateGanttChart, generatePlanSummary, generateBurndownChart similar

  private renderKanban(data: PlanVisualizationData, format: VisualizationFormat): string {
    switch (format) {
      case 'html': return renderKanbanHtml(data);
      case 'svg': return renderKanbanSvg(data);
      case 'markdown': return renderKanbanMarkdown(data);
      case 'mermaid': return renderKanbanMermaid(data);
    }
  }

  private writeAndReturn(content: string, format: VisualizationFormat, viewType: string, outputPath: string | undefined, data: PlanVisualizationData) {
    const ext = format === 'mermaid' ? 'mmd' : format === 'markdown' ? 'md' : format;
    const filePath = outputPath ?? path.join(os.tmpdir(), `planner-${viewType}-${Date.now()}.${ext}`);
    fs.writeFileSync(filePath, content, 'utf-8');
    const preview = `${data.plan.title}: ${data.tasks.length} tasks across ${data.buckets.length} buckets`;
    return {
      content: [{ type: 'text' as const, text: JSON.stringify({ file_path: filePath, format, preview }, null, 2) }],
    };
  }
}
```

**Step 4: Run tool tests to verify they pass**

Run: `npx vitest run tests/unit/tools/planner-visualization.test.ts`
Expected: PASS

**Step 5: Add repository method for visualization data**

Modify `src/graph/repository.ts`: Add `getPlanVisualizationDataAsync(planId: number)` that calls existing `listPlannerTasksAsync` and `listBucketsAsync` and assembles `PlanVisualizationData`.

**Step 6: Register tools in index.ts**

- Add 4 tool definitions to TOOLS array (with `zodToJsonSchema` for inputSchema)
- Add 4 tool names to `GRAPH_ONLY_TOOL_NAMES`
- Add switch cases in `handleGraphToolCall` dispatching to `plannerVisualizationTools`
- Add `plannerVisualizationTools` parameter to `handleGraphToolCall` signature
- Instantiate `PlannerVisualizationTools` in initialization

**Step 7: Run full test suite**

Run: `npm test`
Expected: All tests pass

**Step 8: Commit**

```bash
git add src/tools/planner-visualization.ts src/visualization/ src/graph/repository.ts src/index.ts tests/
git commit -m "feat: add Planner visualization tools (kanban, gantt, summary, burndown)"
```

**Step 9: Bump version to 2.5.0, update CHANGELOG, tag, release**

---

## Phase 9: Teams Enhancements (v2.6.0)

### Task 5: Meeting recordings & transcripts — GraphClient + Repository

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add 6 new methods
- Modify: `src/graph/repository.ts` — add IdCache entries + 6 repository methods

**Step 1: Add GraphClient methods**

```typescript
// Online Meetings
async listOnlineMeetings(limit: number = 20): Promise<any[]> {
  const client = await this.getClient();
  const response = await client.api('/me/onlineMeetings')
    .top(limit)
    .orderby('startDateTime desc')
    .get();
  return response.value;
}

async getOnlineMeeting(meetingId: string): Promise<any> {
  const client = await this.getClient();
  return await client.api(`/me/onlineMeetings/${meetingId}`).get();
}

// Recordings
async listMeetingRecordings(meetingId: string): Promise<any[]> {
  const client = await this.getClient();
  const response = await client.api(`/me/onlineMeetings/${meetingId}/recordings`).get();
  return response.value;
}

async getMeetingRecordingContent(meetingId: string, recordingId: string): Promise<ArrayBuffer> {
  const client = await this.getClient();
  return await client.api(`/me/onlineMeetings/${meetingId}/recordings/${recordingId}/content`).get();
}

// Transcripts
async listMeetingTranscripts(meetingId: string): Promise<any[]> {
  const client = await this.getClient();
  const response = await client.api(`/me/onlineMeetings/${meetingId}/transcripts`).get();
  return response.value;
}

async getMeetingTranscriptContent(meetingId: string, transcriptId: string, format: string = 'text/vtt'): Promise<string> {
  const client = await this.getClient();
  return await client.api(`/me/onlineMeetings/${meetingId}/transcripts/${transcriptId}/content`)
    .header('Accept', format)
    .get();
}
```

**Step 2: Add IdCache entries in repository.ts**

```typescript
// Add to IdCache interface:
onlineMeetings: Map<number, string>;
recordings: Map<number, { meetingId: string; recordingId: string }>;
transcripts: Map<number, { meetingId: string; transcriptId: string }>;
```

**Step 3: Add repository methods**

```typescript
async listOnlineMeetingsAsync(limit?: number): Promise<Array<{ id: number; subject: string; startDateTime: string; endDateTime: string; joinUrl: string }>> { ... }
async getOnlineMeetingAsync(meetingId: number): Promise<{ ... } | undefined> { ... }
async listMeetingRecordingsAsync(meetingId: number): Promise<Array<{ id: number; createdDateTime: string }>> { ... }
async downloadMeetingRecordingAsync(recordingId: number, outputPath: string): Promise<string> { ... }
async listMeetingTranscriptsAsync(meetingId: number): Promise<Array<{ id: number; createdDateTime: string }>> { ... }
async getMeetingTranscriptContentAsync(transcriptId: number): Promise<string> { ... }
```

**Step 4: Commit**

```bash
git add src/graph/client/graph-client.ts src/graph/repository.ts
git commit -m "feat: add meeting recordings and transcripts Graph API methods"
```

---

### Task 6: Meeting recordings & transcripts — Tools + Tests

**Files:**
- Create: `src/tools/meetings.ts`
- Create: `tests/unit/tools/meetings.test.ts`
- Modify: `src/index.ts`

**Step 1: Write failing tests**

`tests/unit/tools/meetings.test.ts` — 12+ tests covering:
- `listOnlineMeetings` returns meetings list
- `getOnlineMeeting` returns meeting details
- `listMeetingRecordings` returns recordings for a meeting
- `downloadMeetingRecording` saves file and returns path
- `listMeetingTranscripts` returns transcripts
- `getMeetingTranscriptContent` returns VTT text

**Step 2: Implement MeetingsTools class**

`src/tools/meetings.ts`:
- `IMeetingsRepository` interface with 6 methods
- Zod schemas: `ListOnlineMeetingsInput`, `GetOnlineMeetingInput`, `ListMeetingRecordingsInput`, `DownloadMeetingRecordingInput`, `ListMeetingTranscriptsInput`, `GetMeetingTranscriptContentInput`
- `MeetingsTools` class with 6 handler methods
- No two-phase approval needed (read-only + local file download)

**Step 3: Register in index.ts**

- 6 tool definitions in TOOLS array
- 6 entries in GRAPH_ONLY_TOOL_NAMES
- 6 switch cases in handleGraphToolCall
- Instantiate MeetingsTools with repository

**Step 4: Run tests**

Run: `npm test`
Expected: All pass

**Step 5: Commit**

```bash
git add src/tools/meetings.ts tests/unit/tools/meetings.test.ts src/index.ts
git commit -m "feat: add meeting recordings and transcripts tools"
```

---

### Task 7: Message reactions — GraphClient + Repository + Tools + Tests

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add reaction methods
- Modify: `src/graph/repository.ts` — add reaction repository methods
- Modify: `src/tools/teams.ts` — extend TeamsTools with 4 reaction methods
- Modify: `src/approval/types.ts` — add `add_message_reaction` operation type
- Modify: `tests/unit/tools/teams.test.ts` — add reaction tests
- Modify: `src/index.ts` — add 4 tool definitions + handlers

**Step 1: Add to approval/types.ts**

Add `'add_message_reaction'` to OperationType, `'message_reaction'` to TargetType.

**Step 2: Add GraphClient methods**

```typescript
async listMessageReactions(teamId: string, channelId: string, messageId: string): Promise<any[]> { ... }
async setMessageReaction(teamId: string, channelId: string, messageId: string, reactionType: string): Promise<void> { ... }
async unsetMessageReaction(teamId: string, channelId: string, messageId: string, reactionType: string): Promise<void> { ... }
// Chat message reactions:
async listChatMessageReactions(chatId: string, messageId: string): Promise<any[]> { ... }
async setChatMessageReaction(chatId: string, messageId: string, reactionType: string): Promise<void> { ... }
async unsetChatMessageReaction(chatId: string, messageId: string, reactionType: string): Promise<void> { ... }
```

**Step 3: Add repository methods**

Extend repository with `listMessageReactionsAsync`, `addMessageReactionAsync`, `removeMessageReactionAsync` that handle both channel and chat messages based on `message_type` param.

**Step 4: Write failing tests**

Add to `tests/unit/tools/teams.test.ts`:
- `listMessageReactions` — returns reactions for a message
- `prepareAddMessageReaction` — generates approval token
- `confirmAddMessageReaction` — adds reaction with valid token
- `confirmAddMessageReaction` — rejects invalid token
- `removeMessageReaction` — removes own reaction

**Step 5: Extend TeamsTools class**

Add 4 methods + 4 Zod schemas to `src/tools/teams.ts`. Extend `ITeamsRepository` interface.

**Step 6: Register in index.ts**

4 tool definitions, 4 GRAPH_ONLY entries, 4 switch cases.

**Step 7: Run full test suite**

Run: `npm test`
Expected: All pass

**Step 8: Commit, bump to 2.6.0, CHANGELOG, tag, release**

```bash
git commit -m "feat: add message reaction tools"
# Then version bump + release
```

---

## Phase 10: OneDrive, SharePoint & Excel Online (v2.7.0)

### Task 8: OneDrive — GraphClient + Repository

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add 8 drive methods
- Modify: `src/graph/repository.ts` — add IdCache entry `driveItems: Map<number, string>`, 8 repository methods

**GraphClient methods:**
```typescript
async listDriveItems(itemId?: string): Promise<any[]> { ... }     // GET /me/drive/root/children or /me/drive/items/{id}/children
async searchDriveItems(query: string): Promise<any[]> { ... }     // GET /me/drive/root/search(q='{query}')
async getDriveItem(itemId: string): Promise<any> { ... }          // GET /me/drive/items/{id}
async downloadDriveItem(itemId: string): Promise<ArrayBuffer> { ... } // GET /me/drive/items/{id}/content
async uploadDriveItem(parentId: string, fileName: string, content: Buffer): Promise<any> { ... } // PUT /me/drive/items/{parentId}:/{fileName}:/content
async listRecentFiles(): Promise<any[]> { ... }                    // GET /me/drive/recent
async listSharedWithMe(): Promise<any[]> { ... }                   // GET /me/drive/sharedWithMe
async createSharingLink(itemId: string, type: string, scope: string): Promise<any> { ... } // POST /me/drive/items/{id}/createLink
async deleteDriveItem(itemId: string): Promise<void> { ... }      // DELETE /me/drive/items/{id}
```

**Step 1: Implement GraphClient methods**
**Step 2: Add IdCache entry and repository methods**
**Step 3: Commit**

```bash
git commit -m "feat: add OneDrive Graph API methods and repository"
```

---

### Task 9: OneDrive — Tools + Tests

**Files:**
- Create: `src/tools/onedrive.ts`
- Create: `tests/unit/tools/onedrive.test.ts`
- Modify: `src/index.ts`
- Modify: `src/approval/types.ts` — add `upload_file`, `delete_drive_item` operation types

**Tools (10):**
- `list_drive_items`, `search_drive_items`, `get_drive_item`, `download_file`, `prepare_upload_file`, `confirm_upload_file`, `list_recent_files`, `list_shared_with_me`, `create_sharing_link`, `prepare_delete_drive_item`, `confirm_delete_drive_item`

**Step 1: Write 15+ failing tests**
**Step 2: Implement OneDriveTools class**
**Step 3: Register in index.ts (11 tool defs, 11 GRAPH_ONLY, 11 switch cases)**
**Step 4: Run tests**
**Step 5: Commit**

```bash
git commit -m "feat: add OneDrive tools"
```

---

### Task 10: SharePoint — GraphClient + Repository + Tools + Tests

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add 5 site/library methods
- Modify: `src/graph/repository.ts` — add IdCache entries `sites`, `documentLibraries`
- Create: `src/tools/sharepoint.ts`
- Create: `tests/unit/tools/sharepoint.test.ts`
- Modify: `src/index.ts`

**GraphClient methods:**
```typescript
async listSites(): Promise<any[]> { ... }                    // GET /sites?search=*
async searchSites(query: string): Promise<any[]> { ... }     // GET /sites?search={query}
async getSite(siteId: string): Promise<any> { ... }          // GET /sites/{id}
async listDocumentLibraries(siteId: string): Promise<any[]> { ... } // GET /sites/{id}/drives
async listLibraryItems(driveId: string, itemId?: string): Promise<any[]> { ... } // GET /drives/{id}/root/children
async downloadLibraryFile(driveId: string, itemId: string): Promise<ArrayBuffer> { ... }
```

**Tools (6):** `list_sites`, `search_sites`, `get_site`, `list_document_libraries`, `list_library_items`, `download_library_file`

**Step 1: Add GraphClient methods**
**Step 2: Add repository methods + IdCache**
**Step 3: Write 10+ failing tests**
**Step 4: Implement SharePointTools class**
**Step 5: Register in index.ts**
**Step 6: Run tests**
**Step 7: Commit**

```bash
git commit -m "feat: add SharePoint tools"
```

---

### Task 11: Excel Online — GraphClient + Repository + Tools + Tests

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add 5 workbook methods
- Modify: `src/graph/repository.ts` — add repository methods
- Create: `src/tools/excel.ts`
- Create: `tests/unit/tools/excel.test.ts`
- Modify: `src/index.ts`
- Modify: `src/approval/types.ts` — add `update_excel_range` operation type

**GraphClient methods:**
```typescript
async listWorksheets(driveItemId: string): Promise<any[]> { ... }  // GET /me/drive/items/{id}/workbook/worksheets
async getWorksheetRange(driveItemId: string, worksheetName: string, range: string): Promise<any> { ... } // GET .../worksheets/{name}/range(address='{range}')
async getUsedRange(driveItemId: string, worksheetName: string): Promise<any> { ... } // GET .../worksheets/{name}/usedRange
async updateRange(driveItemId: string, worksheetName: string, range: string, values: unknown[][]): Promise<any> { ... } // PATCH .../range(address='{range}')
async getTableData(driveItemId: string, tableName: string): Promise<any> { ... } // GET .../workbook/tables/{name}/rows
```

**Tools (6):** `list_worksheets`, `get_worksheet_range`, `get_used_range`, `prepare_update_range`, `confirm_update_range`, `get_table_data`

Excel writes use two-phase approval (modifies remote file).

**Step 1: Add GraphClient methods**
**Step 2: Add repository methods**
**Step 3: Write 10+ failing tests**
**Step 4: Implement ExcelTools class**
**Step 5: Register in index.ts**
**Step 6: Run tests**
**Step 7: Commit, bump to 2.7.0, CHANGELOG, tag, release**

```bash
git commit -m "feat: add Excel Online tools"
```

---

## Phase 11: Graph $batch Optimization (v2.8.0)

### Task 12: Batch request infrastructure

**Files:**
- Create: `src/graph/client/batch.ts`
- Create: `tests/unit/graph/batch.test.ts`
- Modify: `src/graph/client/graph-client.ts` — add `batchRequests()` method

**Step 1: Write failing tests**

`tests/unit/graph/batch.test.ts`:
```typescript
import { describe, it, expect } from 'vitest';
import { buildBatchPayload, parseBatchResponse, splitIntoBatches } from '../../../src/graph/client/batch.js';

describe('buildBatchPayload', () => {
  it('builds a valid $batch request body', () => {
    const requests = [
      { id: '1', method: 'GET', url: '/me/messages' },
      { id: '2', method: 'GET', url: '/me/events' },
    ];
    const payload = buildBatchPayload(requests);
    expect(payload.requests).toHaveLength(2);
    expect(payload.requests[0].id).toBe('1');
  });
});

describe('splitIntoBatches', () => {
  it('splits requests into chunks of 20', () => {
    const requests = Array.from({ length: 25 }, (_, i) => ({
      id: String(i), method: 'GET', url: `/me/item/${i}`,
    }));
    const batches = splitIntoBatches(requests);
    expect(batches).toHaveLength(2);
    expect(batches[0]).toHaveLength(20);
    expect(batches[1]).toHaveLength(5);
  });
});

describe('parseBatchResponse', () => {
  it('maps response IDs to results', () => {
    const response = {
      responses: [
        { id: '1', status: 200, body: { value: [] } },
        { id: '2', status: 200, body: { value: [] } },
      ],
    };
    const results = parseBatchResponse(response);
    expect(results.get('1')?.status).toBe(200);
    expect(results.get('2')?.status).toBe(200);
  });

  it('handles partial failures', () => {
    const response = {
      responses: [
        { id: '1', status: 200, body: { value: [] } },
        { id: '2', status: 404, body: { error: { message: 'Not found' } } },
      ],
    };
    const results = parseBatchResponse(response);
    expect(results.get('1')?.status).toBe(200);
    expect(results.get('2')?.status).toBe(404);
  });
});
```

**Step 2: Implement batch.ts**

`src/graph/client/batch.ts`:
```typescript
export interface BatchRequest {
  id: string;
  method: string;
  url: string;
  headers?: Record<string, string>;
  body?: unknown;
}

export interface BatchResponseItem {
  id: string;
  status: number;
  headers?: Record<string, string>;
  body: unknown;
}

export function buildBatchPayload(requests: BatchRequest[]): { requests: BatchRequest[] } {
  return { requests };
}

export function splitIntoBatches(requests: BatchRequest[], maxPerBatch: number = 20): BatchRequest[][] {
  const batches: BatchRequest[][] = [];
  for (let i = 0; i < requests.length; i += maxPerBatch) {
    batches.push(requests.slice(i, i + maxPerBatch));
  }
  return batches;
}

export function parseBatchResponse(response: { responses: BatchResponseItem[] }): Map<string, BatchResponseItem> {
  const map = new Map<string, BatchResponseItem>();
  for (const item of response.responses) {
    map.set(item.id, item);
  }
  return map;
}
```

**Step 3: Add `batchRequests()` to GraphClient**

```typescript
async batchRequests(requests: BatchRequest[]): Promise<Map<string, BatchResponseItem>> {
  const client = await this.getClient();
  const batches = splitIntoBatches(requests);
  const allResults = new Map<string, BatchResponseItem>();

  for (const batch of batches) {
    const payload = buildBatchPayload(batch);
    const response = await client.api('/$batch').post(payload);
    const results = parseBatchResponse(response);
    for (const [id, result] of results) {
      allResults.set(id, result);
    }
  }

  return allResults;
}
```

**Step 4: Run tests**

Run: `npx vitest run tests/unit/graph/batch.test.ts`
Expected: PASS

**Step 5: Commit**

```bash
git add src/graph/client/batch.ts tests/unit/graph/batch.test.ts src/graph/client/graph-client.ts
git commit -m "feat: add Graph \$batch request infrastructure"
```

---

### Task 13: Opt-in batching in repository methods

**Files:**
- Modify: `src/graph/repository.ts` — use batch for N+1 patterns

**Step 1: Identify N+1 patterns**

Key candidates:
- `listPlannerTasksAsync` + fetching task details for each
- Multi-folder unread counts
- Loading channel messages with extended properties

**Step 2: Refactor with batch calls**

Example — batch loading planner task details:
```typescript
async listPlannerTasksWithDetailsAsync(planId: number): Promise<Array<{ task: ...; details: ... }>> {
  const tasks = await this.listPlannerTasksAsync(planId);
  const requests = tasks.map((t, i) => ({
    id: String(i),
    method: 'GET',
    url: `/planner/tasks/${this.idCache.plannerTasks.get(t.id)?.taskId}/details`,
  }));
  const results = await this.client.batchRequests(requests);
  // Merge results with tasks
}
```

**Step 3: Run full test suite to ensure no regressions**

Run: `npm test`
Expected: All pass

**Step 4: Commit, bump to 2.8.0, CHANGELOG, tag, release**

```bash
git commit -m "feat: add batch optimization for N+1 repository patterns"
```

---

## Release Checklist (per version)

For each version bump:

1. Update `package.json` version
2. Update `CHANGELOG.md` with new version section
3. Run `npm test` — all tests pass
4. Run `npm run typecheck` — no type errors
5. Run `npm run lint` — no lint errors
6. Commit: `git commit -m "chore: release vX.Y.Z"`
7. Tag: `git tag vX.Y.Z`
8. Push: `git push origin main --tags`
9. Release: `gh release create vX.Y.Z --title "vX.Y.Z" --notes "..."`

---

## Task Summary

| Task | Phase | Description | New Tools |
|---|---|---|---|
| T1 | 7 | README rewrite | 0 |
| T2 | 8 | Markdown & Mermaid renderers | 0 |
| T3 | 8 | HTML & SVG renderers | 0 |
| T4 | 8 | Planner visualization tools | 4 |
| T5 | 9 | Meeting recordings/transcripts — backend | 0 |
| T6 | 9 | Meeting recordings/transcripts — tools | 6 |
| T7 | 9 | Message reactions | 4 |
| T8 | 10 | OneDrive — backend | 0 |
| T9 | 10 | OneDrive — tools | 10 |
| T10 | 10 | SharePoint — tools | 6 |
| T11 | 10 | Excel Online — tools | 6 |
| T12 | 11 | Batch infrastructure | 0 |
| T13 | 11 | Batch optimization | 0 |
