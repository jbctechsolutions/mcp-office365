# mcp-office365 v4.0.0 — Status & Next Steps

_Last updated: 2026-07-09. Working branch: `mcp-server-fixes` worktree. `main` is the release target._

**Session progress (this run):** #47 driveItems (dr_), #48 teams+channels pilot (tm_/cn_),
#49 chats/chatMessages/channelMessages (ch_/cm_/xm_), #51 tasks+taskLists (td_/tl_),
#52 task sub-resources (ci_/lr_/ta_). **Teams + To-Do domains fully durable.** Adversarial
review caught 2 real bugs pre-merge (driveItem empty-id guard #47; sub-entity schema break #51).
Follow-up issue #50 (channel reply operability). Remaining alias wave: mail/contact sub-resources,
calendar, meetings, SharePoint chain, Planner. Then U5b-3, union cleanup, U6/U11/U12, coverage
tools, release.

## Release strategy (decided)

The entire durable-ID rollout **plus** the AppleScript removal ship as **one major release: v4.0.0**.
Rationale: durable IDs change the ID format returned to clients (numeric hash → opaque token; a
legacy numeric id on Graph now returns `NUMERIC_ID_UNSUPPORTED`). That's breaking, and breaking
changes already landed on `main`, so per-entity minor releases aren't semver-honest — everything
accumulates to v4.0.0. Additive-only tools (OneNote, SharePoint Lists, shared mailboxes) ride into
v4.0.0 (or a v4.1.0 minor afterward).

Tag `v4.0.0` (via the OIDC release workflow) only after the work below is complete + CHANGELOG has
the breaking-changes migration matrix.

## Done (merged to `main` this program)

- **#33** OIDC tokenless npm publish (`release.yml` uses `npm publish --provenance`; Trusted Publisher
  registered on npmjs; `NPM_TOKEN` deleted).
- **#34** Durable-ID account identity — `src/graph/auth/account-id.ts` derives the stable MSAL
  `homeAccountId`; `ApprovalTokenManager` takes a lazy `accountId` thunk; self-heals on each tool call.
- **#35** `list_my_planner_tasks` (`GET /me/planner/tasks`).
- **Durable self-encoding IDs (the pattern):** contacts **#36**, events **#37**, messages **#41**,
  folders **#45**, driveItems **#47**. Each: mapper mints `mintSelfEncoded('<type>', graphId)`; repo
  resolves via the private `GraphRepository.toGraphId(id, '<type>')`; per-entity `idCache.*` reverse
  maps deleted; tool id params relaxed; characterization-first tests. `getGraphId` is fully deleted
  (folders was its last case).
- **#47** driveItems (`dr_`) — personal OneDrive + Excel single-id case (`/me/drive/items/{id}`).
  Migrated all 9 OneDrive + 5 Excel repo methods off `idCache.driveItems` (map deleted); `item_id`/
  `folder_id`/`file_id` relaxed to `z.string()`; confirm-flow `targetId` casts fixed `as number`→
  `as string` (delete + update-range). Empty-id mint guard added at the 4 list mappers (matches #46).
  **SharePoint document-library items (`idCache.libraryDriveItems`, `{driveId,itemId}` composite)
  deliberately left numeric** — they belong to the alias-composite wave (item #2 below).
- **#42** OneNote tools (`/me/onenote`, durable `nb_`/`ns_`/`np_` tokens) — replaces Apple Notes.
- **#43** **Removed the AppleScript backend** (−11,815 lines). Server is Graph-only. Deleted
  `src/applescript/`, `*-apple.ts`, `notes.ts`, `accounts.ts`; collapsed the dual-backend handlers;
  removed `USE_APPLESCRIPT`, `ToolContext.applescript`, `AppleScriptToolsets`, AppleScript error
  classes. Removed tools: `list_notes`/`get_note`/`search_notes` (→ OneNote), `list_accounts`
  (no Graph analog). Coverage `branches` floor recalibrated 64→63 (deleting well-tested code shifted
  the baseline; the floor already tracks "current actual" per its own comment).
- **#44** `Notes.ReadWrite` Graph scope (OneNote was calling `/me/onenote` without the scope).
- **#46** Mapper empty-id guards (`folderId`/`parentId` mint guarded against empty string).

**Tool count: 221.** Test suite ~1684 green. typecheck + lint clean (one pre-existing `cli.ts`
no-console warning is expected).

## Key architectural facts (read before continuing)

- **String-only from here.** AppleScript is gone, so NEW durable-ID migrations use plain `string` /
  `z.string()` — NOT the `string | number` union. (The older contact/event/message/`ApprovalToken`
  unions are legacy debt; see cleanup step below.)
- **Self-encoding vs alias-backed:** `src/ids/token.ts`. Self-encoding (`em_ ev_ ct_ fd_ dr_ td_`
  + OneNote `nb_ ns_ np_`) carries the Graph id (base64url), resolves by decode with no store —
  cold-durable. Alias-backed (`pl_ pt_ ch_ tm_ at_ cn_ cm_ ci_`) is a digest needing the SQLite
  alias table (`src/state/store.ts`, `registerComposite` in `src/ids/mint.ts`) — machine-scoped.
- **Resolver:** `resolveId(id, accountId, store, expectedEntityType?)` in `src/ids/resolver.ts`.
  Self-encoding needs no store. Numeric → `NUMERIC_ID_UNSUPPORTED`. Wrong entity type →
  `ID_ENTITY_MISMATCH`. Raw non-token string → passthrough as opaque Graph id.
- **Repository resolve helper:** `GraphRepository.toGraphId(id, entityType)` (private) — the choke
  point every migrated method uses.
- **Still numeric (NOT yet migrated):** library driveItems (`{driveId,itemId}` composite — personal
  driveItems done in #47), tasks, task lists, contact folders, planner
  (plan/bucket/task), teams/channels/chat/chatMessages, attachments, checklist items, linked
  resources, meetings/recordings/transcripts, sites/document libraries, focused overrides,
  categories, calendar permissions, calendar groups, mail rules move-target(? migrated), excel.
  These use `hashStringToNumber` + `idCache.*` maps. Their tool params are still `z.number()`.

## Alias-backed composite pattern (LOCKED by #48 — copy this for the wave)

Established in `repository.ts` by the teams+channels pilot (commit `8a3654b`):
- **`mintAlias(entityType, graphId)`** — single-Graph-id entities. `registerComposite`
  with `parts:{id:graphId}`, stored `graphId`.
- **`mintAliasComposite(entityType, parts)`** — multi-id entities. Stores
  `graphId: JSON.stringify(parts)`; `parts` is also the canonical key.
- **`toGraphParts<K>(id, entityType, keys)`** — generic resolver; JSON-parses the
  alias value, validates every required key is a non-empty string (else `ID_UNKNOWN`).
  Generic over keys so destructured fields are non-optional under `noUncheckedIndexedAccess`.
- **Parent cold-miss re-list** (`resolveTeamId`): try `toGraphId`; on `IdUnknownError`
  re-list the parent (deterministic re-mint) then retry. Only `IdUnknownError` is caught
  (not `ID_ENTITY_MISMATCH`/`NUMERIC_ID_UNSUPPORTED`/`ID_FOREIGN_ACCOUNT`).
- **Store**: `this.store` (always present in prod; in-memory-degraded still works).
  Repo tests thread `StateStore.open({dir})` in `beforeEach` (fs mocked → in-memory).
- **Confirm flows**: relax `z.number()`→`z.string()`; fix `targetId as number`→`as string`.
- **Known tradeoff (documented, accepted):** composite child tokens are NOT cold-durable
  and can't self-heal (no parent handle to re-list) — machine-scoped, `ID_UNKNOWN` on a
  cold store. On-disk store persists across restarts (common case fine). Adversarial review
  #48 flagged this asymmetry; it's the deliberate alias-backed choice (short + account-scoped).

## Alias prefix allocation (whole wave — avoid collisions)

Existing alias: `pl`=plan `pt`=plannerTask `ch`=chat `tm`=team `at`=attachment(mail)
`cn`=channel `cm`=chatMessage `ci`=checklistItem. **Decision: ALL remaining composites
are alias-backed** (one mechanical pattern; `td` task moves self→alias). New prefixes:
`td`=task(→alias) `tl`=taskList `xm`=channelMessage `lr`=linkedResource `ta`=taskAttachment
`cp`=calendarPermission `pb`=plannerBucket `rc`=recording `tr`=transcript `dl`=documentLibrary
`li`=libraryDriveItem `mr`=mailRule `cf`=contactFolder `cg`=category `fo`=focusedOverride
`om`=onlineMeeting `si`=site. Each entity PR adds its EntityType(s)+prefix(es) to `token.ts`
+ token.test.ts, then wires repo+tools per the locked pattern above.

## Process / cadence (keep doing this)

1. Cut a branch off latest `origin/main` (**use `git fetch --no-tags origin main:refs/remotes/origin/main`
   — a bare `git fetch` hangs in this env on tag/ref negotiation**).
2. Implement via a **sonnet** subagent with a detailed, file-by-file spec (model on the existing
   contact/event/message/folder migrations — tell it to study those). Opus stays on orchestration.
3. **Independently verify** the subagent's work: `npm run typecheck`, `npm run lint`, `npm test`,
   and grep that the target `idCache.*` maps are gone. Don't trust the subagent's self-report blind.
4. Push, open PR, dispatch an **opus adversarial review** (`ce-adversarial-reviewer`) focused on the
   entity-specific risks. It has caught a real regression on every migration — fix findings before merge.
5. Wait for CI + **CodeRabbit** (check inline comments BEFORE merging — don't merge in the same step).
6. Squash-merge. Commits are SSH-signed via 1Password — **if a commit stalls with
   "1Password: agent returned an error", it needs an interactive unlock**; retry after Joel unlocks.

## Remaining work to v4.0.0 (in order)

1. ~~**driveItems** (`dr_` self-encoding).~~ **DONE #47.** Resolved the open question: `/me/drive/items/{id}`
   single-id suffices for personal OneDrive + Excel (`dr_` self-encoding). SharePoint document-library
   items (`{driveId, itemId}`) ARE composite and are folded into the alias-composite wave below (they
   were already `idCache.libraryDriveItems`, untouched by #47).
2. **Alias-composite wave (U5b-4)** via `registerComposite`: **library driveItems (`{driveId, itemId}`,
   `src/tools/sharepoint.ts` `item_id`/`folder_id`)**, tasks (`{listId, taskId}`), task lists,
   contact folders, plan/plannerTask, chat/team/channel/attachment/chatMessage/checklistItem/
   linkedResource/taskAttachment. These need the store threaded + `ID_UNKNOWN`/`ID_STALE` on cold miss.
   **Watch:** the ~13 un-migrated confirm flows that cast `token.targetId as number` — each becomes a
   real string as its entity migrates; fix the cast (don't leave `as number` hiding a string).
3. **U5b-3 immutable-ID preference:** add `Prefer: IdType="ImmutableId"` header (middleware in
   `graph-client.ts`) + `translateExchangeIds` batch for `$search`-minted mutable ids +
   `mutable=1` alias fallback (`ID_STALE`, `degraded_ids` count).
4. **U5b-5 Planner fetch-before-update:** drop the cached ETag (rides inside `idCache.plans/*` today),
   fetch-then-PATCH with one 412 retry.
5. **Union cleanup:** simplify the legacy `string | number` on contact/event/message ids + the
   `*Summary.id` types + `ApprovalToken.targetId` to `string`; delete `hashStringToNumber` and its
   remaining call sites (only possible once EVERY entity is migrated). Remove the dead `CreateEventInput`
   (legacy schema; `create_event` registers `CreateEventGraphInput`).
6. **Coverage-gap tools (GitHub issues filed):** **#38** SharePoint Lists (`/sites/{id}/lists`),
   **#40** shared-mailbox / delegate access (`/users/{upn}/...`). (#39 OneNote already done.)
7. **U6** canonical per-entity Zod schemas + alias coercion + next-action hints.
8. **U11** elicitation (`--confirm elicit`, 60s wait degrading to durable token).
9. **U12** delta-sync local mirror + `what_changed` tool.
10. **Cut v4.0.0:** version bump, CHANGELOG breaking-changes migration matrix (numeric→token id format,
    removed tools/backend), tag `v4.0.0` (OIDC workflow publishes).

## Open follow-ups / known low-severity items (documented, deferred)

- **Polymorphic `fd_` token** (mail folder vs calendar share the prefix): fails safe via Graph 404,
  but `getMailFolder` swallows the 404 → a calendar token to a mail endpoint returns a silent empty
  result instead of a clear wrong-entity error. Could de-swallow later.
- **Seal-with-raw-id mismatch:** a two-phase approval prepared with a RAW Graph id can't be confirmed
  with the same raw id (seal keys on the minted token). Fails closed. Resolve in the union cleanup
  (normalize the seal to the resolved graph id).
- **Coverage:** `branches` floor is 63 (temporary "current actual"; real target 75). Raise as tests land.
- README/docs `--preset` list is missing some presets (files/onenote); minor doc debt.

## Coverage-gap analysis (for later feature work, beyond durable IDs)

Biggest "we'd hit Graph directly" gaps: **SharePoint Lists** (#38), **shared mailboxes** (#40),
Word/PPT read+convert-to-PDF, create-Teams-meeting (`POST /me/onlineMeetings`). Depth gaps: Excel
charts/pivots/sessions, OneDrive move/copy/versions, unified `/search/query`, groups/directory,
`/me/insights`.
