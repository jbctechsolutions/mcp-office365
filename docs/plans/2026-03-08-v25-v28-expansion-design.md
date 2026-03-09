# M365 MCP Server v2.5–v2.8 Expansion Design

## Overview

Continuing the M365 MCP server expansion from v2.4.0 (189 tools). Target audience: knowledge workers (individual productivity) and team leads (project management). Five phases delivering README improvements, Planner visualization, Teams enhancements, OneDrive/SharePoint/Excel, and internal batch optimization.

## Phase 7: README Update (v2.4.1)

**Goal:** Full rewrite of README to accurately represent the 189-tool M365 MCP server.

**Structure:**
1. Header — new name, badges (npm, version, license, tools count)
2. One-liner — "MCP server for Microsoft 365 on macOS — mail, calendar, contacts, tasks, teams, people, and planner"
3. Features overview — grouped by service with tool counts
4. Quick start — install, auth setup (AppleScript vs Graph), Claude Desktop config
5. Tool reference — collapsible sections per category with every tool listed
6. Architecture — dual-backend diagram (AppleScript + Graph), two-phase approval explanation
7. Permissions/scopes — which Graph API scopes are needed per feature area
8. Contributing / License

**Tool categories in reference:**

| Category | Count |
|---|---|
| Mail (read, send, organize, rules, categories, focused inbox, settings, tips, headers) | ~70 |
| Calendar (events, groups, permissions, rooms) | ~19 |
| Contacts & folders | ~13 |
| Tasks, checklists, linked resources, attachments | ~24 |
| Notes (AppleScript only) | 3 |
| Scheduling | 2 |
| Teams (channels, messages, chats) | 20 |
| People & presence | 8 |
| Planner (plans, buckets, tasks, details) | 17 |
| Accounts | 1 |

No new code — documentation only. Patch release.

---

## Phase 8: Planner Visualization (v2.5.0)

**Goal:** Generate visual representations of Planner data in multiple formats.

### Tools (4)

| Tool | Description |
|---|---|
| `generate_kanban_board` | Kanban view of tasks grouped by bucket, colored by priority |
| `generate_gantt_chart` | Gantt timeline of tasks with start/due dates and dependencies |
| `generate_plan_summary` | Overview stats — task counts by status, assignee workload, overdue items |
| `generate_burndown_chart` | Burndown/burnup of completed vs remaining tasks over time |

### Input Schema (common)

Each tool accepts:
- `plan_id` (number) — numeric ID from ID cache
- `format` (enum): `"html"` | `"svg"` | `"markdown"` | `"mermaid"` — default: `"html"`
- `output_path` (string, optional) — where to save file; defaults to temp dir

### Format Details

- **HTML** (priority): Interactive — CSS grid Kanban, JS-powered Gantt with hover tooltips, responsive. Self-contained single file (inline CSS/JS).
- **SVG** (priority): Clean static visuals, embeddable anywhere. Generated programmatically (no external deps).
- **Markdown**: Tables and text — Kanban as columnar tables, Gantt as Mermaid code block, summary as stats table.
- **Mermaid**: Native Mermaid syntax — `gantt` diagram type for Gantt, list-based for others.

### Architecture

- New file: `src/tools/planner-visualization.ts`
- New directory: `src/visualization/` with renderers per format (`html.ts`, `svg.ts`, `markdown.ts`, `mermaid.ts`)
- Reads from existing Planner repository (no new Graph API calls)
- Returns `{ file_path, format, preview }` where `preview` is a truncated text summary
- No two-phase approval — read-only + local file generation
- All tools in `GRAPH_ONLY_TOOL_NAMES`

---

## Phase 9: Teams Enhancements (v2.6.0)

**Goal:** Meeting recordings/transcripts access + message reactions.

### Tools (10)

**Meeting Recordings & Transcripts (6):**

| Tool | Description |
|---|---|
| `list_online_meetings` | List user's online meetings (recent/upcoming) |
| `get_online_meeting` | Get meeting details by join URL or meeting ID |
| `list_meeting_recordings` | List recordings for a meeting |
| `download_meeting_recording` | Download recording to local file |
| `list_meeting_transcripts` | List transcripts for a meeting |
| `get_meeting_transcript_content` | Get transcript text content (VTT/text format) |

**Message Reactions (4):**

| Tool | Description |
|---|---|
| `list_message_reactions` | List reactions on a channel message or chat message |
| `prepare_add_message_reaction` | Preview adding a reaction (emoji name) to a message |
| `confirm_add_message_reaction` | Execute the reaction add |
| `remove_message_reaction` | Remove your own reaction from a message |

### Architecture

- Recordings/transcripts: New Graph API methods on `GraphClient` using `/me/onlineMeetings` and `/communications` endpoints
- Reactions: Extend existing channel message / chat message methods via Graph API `POST /reactions` and `DELETE /reactions`
- Reactions use two-phase approval for add (visible to others), direct for remove (only removes own)
- New IdCache entries: `onlineMeetings`, `recordings`, `transcripts`
- All tools in `GRAPH_ONLY_TOOL_NAMES`

### Graph Permissions

- `OnlineMeetings.Read` — meetings list/details
- `OnlineMeetingRecording.Read.All` — recordings
- `OnlineMeetingTranscript.Read.All` — transcripts
- Channel/chat message reactions use existing `ChannelMessage.Send` / `Chat.ReadWrite` scopes

---

## Phase 10: OneDrive, SharePoint & Excel Online (v2.7.0)

**Goal:** File operations, document libraries, and spreadsheet access.

### Tools (22)

**OneDrive — Personal Files (10):**

| Tool | Description |
|---|---|
| `list_drive_items` | List files/folders in a directory (root or by path/ID) |
| `search_drive_items` | Search files by name/content across OneDrive |
| `get_drive_item` | Get file metadata (size, modified, sharing info) |
| `download_file` | Download file to local path |
| `prepare_upload_file` | Preview uploading a local file to OneDrive |
| `confirm_upload_file` | Execute the upload |
| `list_recent_files` | List recently accessed files |
| `list_shared_with_me` | List files shared with the user |
| `create_sharing_link` | Create a sharing link (view/edit) for a file |
| `prepare_delete_drive_item` / `confirm_delete_drive_item` | Two-phase delete |

**SharePoint — Team Documents (6):**

| Tool | Description |
|---|---|
| `list_sites` | List SharePoint sites the user has access to |
| `search_sites` | Search for SharePoint sites by keyword |
| `get_site` | Get site details |
| `list_document_libraries` | List document libraries in a site |
| `list_library_items` | List files in a document library (supports path navigation) |
| `download_library_file` | Download a file from a document library |

**Excel Online (6):**

| Tool | Description |
|---|---|
| `list_worksheets` | List worksheets in an Excel file (OneDrive or SharePoint) |
| `get_worksheet_range` | Read cell values from a range (e.g., "A1:D10") |
| `get_used_range` | Read all data in the used range of a worksheet |
| `prepare_update_range` | Preview writing values to a cell range |
| `confirm_update_range` | Execute the write |
| `get_table_data` | Read a named table's rows and columns |

### Architecture

- New files: `src/tools/onedrive.ts`, `src/tools/sharepoint.ts`, `src/tools/excel.ts`
- New GraphClient methods for `/me/drive`, `/sites`, and `/me/drive/items/{id}/workbook`
- IdCache entries: `driveItems`, `sites`, `documentLibraries`
- Upload uses two-phase (creates content on remote), delete uses two-phase, Excel writes use two-phase
- Downloads save to local filesystem (following existing `download_attachment` pattern)

### Graph Permissions

- `Files.ReadWrite` — OneDrive files
- `Sites.Read.All` — SharePoint site access
- `Files.ReadWrite.All` — SharePoint file downloads (broader scope)

---

## Phase 11: Graph $batch Optimization (v2.8.0)

**Goal:** Internal performance optimization — no new user-facing tools.

### What Gets Batched

- Multi-item fetches: listing items then fetching details for each (e.g., plan tasks + their details)
- Parallel reads: loading multiple mailbox folders, multiple calendars
- Extending the pattern from `get_users_presence` (already batches up to 650)

### Implementation

- New method on `GraphClient`: `batchRequests(requests: BatchRequest[])` using `POST /$batch`
- Each `BatchRequest` = `{ id, method, url, headers?, body? }`
- Handles Graph's 20-request-per-batch limit internally (splits into multiple batches)
- Handles individual request failures within a batch (partial success)
- Repository layer opts in where it makes sense — no changes to tool layer

### Where It Helps Most

- `list_planner_tasks` + task details fetch (currently N+1 calls)
- Loading channel messages with reaction counts
- Multi-folder email counts

### Architecture

- New file: `src/graph/client/batch.ts` — batch request builder and response parser
- Modified: `graph-client.ts` — adds `batchRequests()` method
- Modified: `repository.ts` — opt-in batching in specific methods
- No new tools, no new IdCache entries, no approval changes

---

## Summary

| Phase | Feature | Version | New Tools | Total Tools |
|---|---|---|---|---|
| 7 | README Update | v2.4.1 | 0 | 189 |
| 8 | Planner Visualization | v2.5.0 | 4 | 193 |
| 9 | Teams Enhancements | v2.6.0 | 10 | 203 |
| 10 | OneDrive/SharePoint/Excel | v2.7.0 | 22 | 225 |
| 11 | Graph $batch | v2.8.0 | 0 | 225 |

Final tool count: ~225 tools across mail, calendar, contacts, tasks, teams, people, planner, visualization, OneDrive, SharePoint, and Excel.
