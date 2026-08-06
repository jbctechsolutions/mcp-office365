---
title: 'feat: Full Planner support — labels, plan details, delete plan, ordering, remote write scope'
type: feat
status: completed
date: 2026-08-06
---

# feat: Full Planner support

## Summary

Close the gaps found in the 2026-08-06 Planner audit (all six items approved for build): (1) the remote claude.ai connector exposes Planner write tools but its Entra API app was never granted any `Tasks.*` Graph scope, so every write 403s — fixed with a one-block jp-infrastructure change plus admin re-consent; (2) the tool surface itself is missing label support (`appliedCategories` + `plannerPlanDetails` category names), plan deletion, plan-details read/write, `orderHint` writes, and creation ergonomics (percent/description/checklist at task creation, audit gap item 6) — added here following the repo's existing planner tool/repository/client layering, ETag fetch-before-write pattern, and two-phase delete convention. Five new tools land (`get_plan_details`, `update_plan_details`, `update_plan_sharing`, `prepare_delete_plan`, `confirm_delete_plan`), taking the registry from 250 to 255 tools.

---

## Problem Frame

The original remote-connector requirements (R7/A2 in `docs/brainstorms/2026-07-11-remote-connector-mode-requirements.md`) intended JP staff to *read and write* Planner items, but the connector app registration shipped with Planner-read only (`Group.Read.All`), with writes deferred pending scope validation. That validation is now done: Microsoft's current v1.0 docs list `Tasks.ReadWrite` as the least-privileged delegated scope for **every** Planner write — no `Group.ReadWrite.All` needed. Separately, Planner labels are unusable through the MCP (no `appliedCategories`, no category-name access), plans can be created but never deleted, and board/bucket ordering can't be controlled.

---

## Assumptions

*This plan was authored in pipeline mode without synchronous user confirmation. The items below are agent inferences — review as bets, not settled decisions.*

- Remote pinned surface (`DEFAULT_TOOL_SURFACE`): `get_plan_details` and `update_plan_details` (labels — the pilot users' headline ask) are **added**; the `prepare_delete_plan`/`confirm_delete_plan` pair and `update_plan_sharing` are **excluded** — plan deletion and plan-sharing writes (access grant/revoke via `sharedWith`) both carry a larger blast radius than anything currently exposed to pilot users, and the same criterion is applied to both. Local stdio and `fullAccess` users get all five tools automatically. Whether Graph restricts `sharedWith` updates to plan owners (vs any member) is unverified — another reason to keep it off the remote surface for now.
- New tool names mirror existing conventions: `get_plan_details`/`update_plan_details` (matching `get_plan`, not `get_planner_plan_details`).
- Label input shape is a `applied_categories` record keyed `category1`…`category25` with boolean values, passed through to Graph with light zod validation only (no client-side defensive apparatus — per the probe-before-defenses learning, Graph's own 400s are surfaced instead).
- Admin re-consent for the new Graph permission is a manual tenant-admin action Joel performs after the jp-infrastructure PR is applied; the pipeline cannot execute it.
- Rendering labels in the visualization tools (kanban/gantt) is deferred to follow-up work.

---

## Requirements

- R1. Remote connector users can create and update Planner tasks (and all other already-exposed Planner writes) without Graph 403s, using least-privileged delegated scope.
- R2. Planner labels are fully usable: set/clear `appliedCategories` on task create and update, and read/rename the plan's category names.
- R3. Plans can be deleted via the standard two-phase prepare/confirm approval flow.
- R4. Plan details (`categoryDescriptions`, `sharedWith`) are readable and writable with correct per-resource ETag concurrency.
- R5. `orderHint` is writable on buckets (create/update) and tasks (create/update) using Graph's documented hint format.
- R6. Task creation supports `percent_complete` natively and description/checklist via an automatic follow-up details write.
- R7. Docs and pinned counts stay truthful: README tool tables/counts, permissions table, stale ETag paragraph, and `docs/remote/provisioning.md`'s "Planner write deferred" note are all updated; the e2e 250-tool assertions are bumped to 255.
- R8. The remote pinned surface changes only by explicit, reviewed edit (per the R7 pinning contract in `src/remote/entitlements.ts`).

---

## Scope Boundaries

- No new Graph scopes beyond `Tasks.ReadWrite` on the connector API app; the local stdio app already has it. No `Group.ReadWrite.All`.
- No changes to Planner task comment tools (beta) — they already work under `Tasks.ReadWrite` and stay out of the remote default surface.
- No `list_groups` tool in this pass — `list_teams` covers team-connected groups for `create_plan` discovery.
- No board-format tools (`assignedToTaskBoardFormat` etc.), no Planner delta/what_changed integration, no premium (Project-backed) plan support (not in Graph v1.0).

### Deferred to Follow-Up Work

- Label rendering in visualization tools (`generate_kanban_board` cards showing labels): future iteration once labels land.
- M365 group discovery tool (`list_groups`) for non-team groups: separate feature if pilot feedback asks for it.
- Local device-code app registration drift in jp-infrastructure (`stacks/azure/entra/mcp-office365/main.tf` is missing `Mail.Send`, `Files.ReadWrite`, `Sites.ReadWrite.All`, `Notes.ReadWrite`, shared-mailbox scopes that `GRAPH_SCOPES` requests): separate jp-infrastructure PR, unrelated to Planner.

---

## Context & Research

### Relevant Code and Patterns

- `src/tools/planner.ts` — canonical domain module: zod strict schemas, `PlannerTools` class, `plannerToolDefinitions()` with `defineTool`, two-phase delete via `ApprovalTokenManager` + `approvalTokenLink` (copy `prepareDeleteBucket`/`confirmDeleteBucket`).
- `src/graph/repository.ts` — `withFreshEtag(fetchEtag, write)` (line ~214): GET fresh `@odata.etag`, throw on empty, write with If-Match, retry once on 412. All new writes use it. `mintAlias`/`toGraphId`; `resolvePlanId` self-heals cold `pl_` tokens — plan-details tools must route through it.
- `src/graph/client/graph-client.ts` planner section (~2103–2250) — new client methods needed: `getPlanDetails`, `updatePlanDetails(etag)`, `deletePlan(etag)`; PATCH bodies pass `.header('If-Match', etag)`.
- `tests/unit/tools/planner.test.ts` (hand-built `IPlannerRepository` mock — every new interface method needs a `vi.fn()`), `tests/unit/graph/repository.test.ts` (literal GraphClient mock needs new methods; planner ETag tests ~line 5290 are the model), `tests/contract/invariants.test.ts` (auto-covers new tools; enforces `prepare_delete_plan`→`confirm_delete_plan` naming and destructive flags).
- `tests/e2e/mcp-client.test.ts` — hard-codes tool count 250 in three places (lines ~50, 106, 134).
- `src/remote/entitlements.ts` `DEFAULT_TOOL_SURFACE` — pinned allow-list; contract test only asserts pinned names exist, so new tools are remote-invisible until explicitly added.
- `src/ids/next-action.ts` `FOLLOWUP_TOOLS` — any tool name mentioned must exist in the registry (`tests/unit/ids/schema.test.ts`).
- jp-infrastructure `stacks/azure/entra/mcp-office365-connector/main.tf` (~line 139) — Planner-read-only `resource_access` block with the outdated "broader group scope" comment.

### Institutional Learnings

- `docs/solutions/design-patterns/fetch-before-update-for-mutable-etags.md` — never cache ETags; **never stash an ETag in an approval token**; details sub-resources have their *own* ETags distinct from the parent's.
- `docs/solutions/architecture-patterns/alias-backed-composite-durable-id-pattern.md` — plan details piggyback the `pl_` token (like task details on `pt_`); no new entity type needed.
- `docs/solutions/best-practices/test-external-api-assumptions-before-building-defenses.md` — keep appliedCategories/orderHint validation minimal; surface Graph 400s rather than pre-building defenses.
- `docs/solutions/architecture-patterns/stateless-http-transport-for-stdio-mcp-server.md` — approval tokens are store-bound (single replica); delete_plan inherits this constraint, no cross-machine redemption.
- `docs/solutions/conventions/adversarial-review-as-primary-gate.md` — review hunt list: empty `If-Match`, empty-id guards, cross-consumer schema breaks, tool-count assertion.
- `docs/solutions/integration-issues/claude-ai-entra-oauth-remote-mcp-connector-2026-07-17.md` — new delegated Graph permissions on the API app flow through OBO `.default` with **no PRM changes**; only tenant admin consent is needed post-apply.

### External References

- Graph v1.0 permission docs (verified 2026-08-06): `Tasks.ReadWrite` is least-privileged delegated for create/update/delete plannerPlan, plannerPlanDetails update, create/update/delete plannerTask, plannerTaskDetails update, create plannerBucket. `Group.ReadWrite.All` is everywhere only the higher-privileged alternative. Sources: learn.microsoft.com `planner-post-plans`, `plannerplan-update`, `plannerplan-delete`, `plannerplandetails-update`, `planner-post-tasks`, `plannertask-update`, `plannertask-delete` (view=graph-rest-1.0).
- `appliedCategories`: flat map `{categoryN: true}`; PATCH merges (omitted keys preserved), `false` removes the key; **25 categories** (`plannerCategoryDescriptions` page; the 6-category text on `plannerAppliedCategories` is stale). Settable on task POST.
- `plannerPlanDetails`: `categoryDescriptions` category1–25 (string|null; null resets), `sharedWith` user-GUID→bool map; PATCH requires the **details object's** ETag.
- Delete plannerPlan: `DELETE /planner/plans/{id}` + If-Match; contained buckets/tasks go with it; no pre-emptying required.
- orderHint format: `"<previous> <next>!"`, missing neighbor = empty string, append = `" !"`; echoing a service-returned hint verbatim → 400; use `Prefer: return=representation` to read back canonical hints.
- Task creation accepts `percentComplete`, `priority`, `appliedCategories`, `orderHint`; **details (description/checklist) cannot be set on POST** — follow-up PATCH to `/details` with the details' own ETag (fetched via GET after create).

---

## Key Technical Decisions

- **`Tasks.ReadWrite`, not `Group.ReadWrite.All`, for the connector app**: least-privileged per current docs; the original deferral comment predates this verification. No server code change needed — OBO uses `.default`.
- **Plan details piggyback the `pl_` token**: mirrors task details on `pt_`; no new entity type, no new prefix, no schema.ts changes.
- **Fresh-ETag-at-confirm for delete_plan**: the approval token stores only the plan id; the ETag is fetched inside `confirmDeletePlan` via `withFreshEtag` (learnings #1/#5).
- **Composite create for ergonomics**: `create_planner_task` accepts optional `description`/`checklist`; repository creates the task, then (only when those fields are present) GETs the auto-created details for its ETag and PATCHes. Partial failure returns `task_id` plus a `details_warning` field rather than failing the whole create — the task exists; hiding that would strand it.
- **Minimal input validation**: `applied_categories` keys validated by regex `^category([1-9]|1[0-9]|2[0-5])$`, values boolean; `order_hint` passed through as a plain string with the format documented in the description. Graph errors surface as-is.
- **Remote surface edit is explicit and partial**: `get_plan_details`/`update_plan_details` added to `DEFAULT_TOOL_SURFACE`; delete-plan pair and `update_plan_sharing` deliberately left out (see Assumptions).
- **Sharing writes are a separate tool** (`update_plan_sharing`), not a field on `update_plan_details`: entitlements are per-tool, so splitting is the only way to ship labels to remote users while keeping access grant/revoke off the remote surface. Both tools PATCH the same `/planner/plans/{id}/details` resource via `withFreshEtag`.

---

## Open Questions

### Resolved During Planning

- Does Planner write need `Group.ReadWrite.All`? **No** — `Tasks.ReadWrite` is least-privileged for all writes (verified against live learn.microsoft.com pages).
- Can task description/checklist be set at creation? **No** — details are a separate auto-created resource; requires follow-up PATCH with the details' own ETag.
- Does deleting a plan require emptying it first? **No** — contained buckets/tasks are deleted with the plan.
- How many label categories? **25** (category1–category25); the 6-category doc text is stale.

### Deferred to Implementation

- Exact canonical orderHint values Graph returns: service-normalized; tests assert passthrough of the request value, not the normalized response.
- Whether `Prefer: return=representation` is worth adding to planner PATCHes: only adopt if implementation finds the extra GET burdensome; not required for correctness.

---

## Implementation Units

### U1. jp-infrastructure: grant `Tasks.ReadWrite` to the connector API app

**Goal:** Remote connector Planner writes stop 403ing — the API app's Graph `required_resource_access` gains `Tasks.ReadWrite`.

**Requirements:** R1

**Dependencies:** None (separate repo; can land in parallel with U2–U7)

**Target repo:** `joshua-project/jp-infrastructure` (feature branch from `main`, PR targets `main`; use a fresh worktree — Joel has an unrelated feature branch checked out)

**Files:**
- Modify: `stacks/azure/entra/mcp-office365-connector/main.tf` — add a `resource_access` block for `oauth2_permission_scope_ids["Tasks.ReadWrite"]` in the Planner section; replace the "Planner (read)… broader group scope" comment with one stating `Tasks.ReadWrite` is least-privileged for all Planner writes (verified 2026-08-06 against Graph v1.0 docs).

**Approach:**
- One-block Terraform change following the existing `resource_access` pattern in the same file. No PRM/server changes (OBO `.default` picks it up). Admin consent is NOT codified in Terraform for this stack — the PR description must state the post-apply manual step: `az ad app permission admin-consent --id 484c0657-6a05-4aad-a175-dabac48acb05` (tenant admin), per `docs/remote/provisioning.md` Step 1.

**Test scenarios:**
- Test expectation: none — pure Terraform config; `terraform validate`/plan output in the PR is the verification. Do NOT run `terraform apply` (CI/Joel applies).

**Verification:**
- `terraform fmt -check` and `terraform validate` pass in the stack directory; the diff adds exactly one scope; PR body carries the admin-consent runbook step.

---

### U2. Labels: `applied_categories` on task create/update + read surfaces

**Goal:** Tasks can be labeled: `create_planner_task` and `update_planner_task` accept `applied_categories`; task reads return it.

**Requirements:** R2

**Dependencies:** None

**Files:**
- Modify: `src/tools/planner.ts` (schemas + params mapping), `src/graph/repository.ts` (`createPlannerTaskAsync`, `updatePlannerTaskAsync`, task mappers in `getPlannerTaskAsync`/`listPlannerTasksAsync`/`listMyPlannerTasksAsync`), `src/graph/client/graph-client.ts` (body passthrough — likely no change since bodies are built in repository)
- Test: `tests/unit/tools/planner.test.ts`, `tests/unit/graph/repository.test.ts`

**Approach:**
- Zod: `applied_categories: z.record(z.string().regex(/^category([1-9]|1[0-9]|2[0-5])$/), z.boolean()).optional()` with a description explaining `true` applies, `false` removes, omitted keys are preserved on update.
- Repository passes the map to Graph as `appliedCategories` unchanged. Reads include `appliedCategories` in task results (empty object when absent).

**Test scenarios:**
- Happy path: update with `{category3: true, category4: false}` → PATCH body carries `appliedCategories` verbatim; response reports success.
- Happy path: create with `applied_categories` → POST body includes it.
- Edge case: `category25` accepted; `category26` and `categoryX` rejected by zod with a clear error.
- Edge case: omitted `applied_categories` → field absent from PATCH body (no accidental clearing).
- Integration (mock-level): `get_planner_task` surfaces `appliedCategories` returned by the client mock.

**Verification:**
- New/updated unit tests pass; contract invariants suite passes unchanged.

---

### U3. Plan details tools: `get_plan_details` / `update_plan_details` / `update_plan_sharing`

**Goal:** Category label names become readable and writable (making U2's labels human-meaningful), and plan sharing is writable via a deliberately separate tool that stays off the remote surface.

**Requirements:** R2, R4

**Dependencies:** None (pairs naturally with U2)

**Files:**
- Modify: `src/tools/planner.ts` (three new input schemas, `IPlannerRepository` methods, `PlannerTools` methods, three `defineTool` entries), `src/graph/repository.ts` (`getPlanDetailsAsync`/`updatePlanDetailsAsync`/`updatePlanSharingAsync` using `resolvePlanId` + `withFreshEtag` on the **details** resource), `src/graph/client/graph-client.ts` (`getPlanDetails`, `updatePlanDetails(planId, updates, etag)` — shared by both update tools), `src/ids/next-action.ts` (extend `FOLLOWUP_TOOLS.plan` mention to include `get_plan_details`)
- Test: `tests/unit/tools/planner.test.ts`, `tests/unit/graph/repository.test.ts`

**Approach:**
- Mirror `get_planner_task_details`/`update_planner_task_details` exactly: details addressed by the plan's `pl_` token, no new entity type. ETag comes from `GET /planner/plans/{id}/details` inside `withFreshEtag` — never the plan's own ETag.
- `get_plan_details` returns `categoryDescriptions`, `sharedWith`, `etag`.
- `update_plan_details` accepts `category_descriptions` only (record `category1`–`25` → string|null, null resets to default).
- `update_plan_sharing` accepts `shared_with` (record user-GUID → boolean; `true` adds, `false` removes). Its description must state that removal may revoke a user's plan access and that group members retain access via group membership. Kept out of the remote pinned surface (see Assumptions).
- Both update tools converge on the same client PATCH (`updatePlanDetails`); the split exists purely at the tool/entitlement layer.
- Annotations: get = read-only; `update_plan_details` = non-destructive write; `update_plan_sharing` = `destructive: true`/`destructiveHint: true` (access revocation, per code-review finding — standalone destructive follows the `delete_event` precedent). All `presets: ['planner']`, `backends: ['graph']`, description suffix `(Graph API)`.

**Patterns to follow:**
- `getPlannerTaskDetailsAsync`/`updatePlannerTaskDetailsAsync` in `src/graph/repository.ts` (~2836/2915).

**Test scenarios:**
- Happy path: get returns `categoryDescriptions`, `sharedWith`, `etag`.
- Happy path: update with `{category_descriptions: {category1: "Blocked"}}` → PATCH to `/planner/plans/{gid}/details` with If-Match from the details GET (assert call ordering: details GET before PATCH).
- Happy path: `{category1: null}` passes through as null (reset semantics preserved — zod must allow null, not strip it).
- Happy path: `update_plan_sharing` with `{shared_with: {"<guid>": true}}` → PATCH body `{sharedWith: {"<guid>": true}}` with fresh details ETag.
- Edge case: `update_plan_details` schema rejects a `shared_with` key (strict object — sharing cannot ride in through the labels tool).
- Error path: details GET returns no `@odata.etag` → loud failure, no PATCH sent.
- Error path: first PATCH 412 → exactly one retry with re-fetched ETag (mirror existing planner ETag tests ~repository.test.ts:5290).
- Integration: cold-store `pl_` token self-heals via `resolvePlanId` re-list before the details call.

**Verification:**
- Both tools appear in the registry; contract suite passes; unit tests green.

---

### U4. Two-phase plan deletion: `prepare_delete_plan` / `confirm_delete_plan`

**Goal:** Plans can be deleted with the same approval-token safety as buckets/tasks.

**Requirements:** R3

**Dependencies:** None

**Files:**
- Modify: `src/tools/planner.ts` (schemas, `IPlannerRepository.deletePlanAsync`, prepare/confirm methods, two `defineTool` entries with `onElicit: approvalTokenLink('confirm_delete_plan')`), `src/graph/repository.ts` (`deletePlanAsync` via `withFreshEtag`), `src/graph/client/graph-client.ts` (`deletePlan(planId, etag)`)
- Test: `tests/unit/tools/planner.test.ts`, `tests/unit/graph/repository.test.ts`

**Approach:**
- Copy `prepareDeleteBucket`/`confirmDeleteBucket` verbatim shape: token `{operation: 'delete_plan', targetType: 'plan', targetId, targetHash}`; prepare result carries `approval_token`, `expires_at`, `plan_id`, and an `action` string that ALSO warns the delete removes all buckets and tasks in the plan. Token stores the id only — ETag fetched fresh at confirm.
- Contract test derives `confirm_delete_plan` from the prepare name mechanically — naming is forced. Both `destructive: true`; prepare `destructiveHint: false`, confirm `destructiveHint: true`.

**Test scenarios:**
- Happy path: prepare → token; confirm with token → client `deletePlan` called with the resolved Graph id and a freshly fetched ETag.
- Error paths: expired token, wrong-operation token, already-consumed token, unknown token → each returns the mapped error message, no delete call.
- Error path: 412 on delete → one retry with re-fetched ETag.
- Edge case: confirm for a token minted for a different plan (TARGET_MISMATCH) → refused.

**Verification:**
- Contract invariants (prepare/confirm pairing, destructive flags) pass; unit tests green.

---

### U5. orderHint writes on buckets and tasks

**Goal:** Board and bucket ordering is controllable: `order_hint` on `create_bucket`, `update_bucket`, `create_planner_task`, `update_planner_task`.

**Requirements:** R5

**Dependencies:** U2 (same schema/mapping regions in planner.ts/repository.ts — sequence to avoid conflicts)

**Files:**
- Modify: `src/tools/planner.ts`, `src/graph/repository.ts` (`createBucketAsync` signature gains optional orderHint, `updateBucketAsync`, task create/update mappers), `src/graph/client/graph-client.ts` (`createBucket` body)
- Test: `tests/unit/tools/planner.test.ts`, `tests/unit/graph/repository.test.ts`

**Approach:**
- `order_hint: z.string().min(1).optional()` with a description teaching the Graph format: `"<previous> <next>!"`, `" !"` to append, and the warning that echoing a previously returned hint verbatim causes a 400. No client-side format validation (probe-before-defenses).

**Test scenarios:**
- Happy path: `update_bucket` with `order_hint: " !"` → PATCH body `{orderHint: " !"}` alongside/without `name`.
- Happy path: task create with `order_hint` → POST body includes `orderHint`.
- Edge case: omitted → absent from body.

**Verification:**
- Unit tests green; existing bucket/task tests unaffected.

---

### U6. Creation ergonomics: `percent_complete`, `description`, `checklist` at task creation

**Goal:** One `create_planner_task` call can produce a complete task — progress, notes, and checklist included.

**Requirements:** R6

**Dependencies:** U2, U5 (same create-path code)

**Files:**
- Modify: `src/tools/planner.ts` (`CreatePlannerTaskInput` gains `percent_complete`, `description`, `checklist`), `src/graph/repository.ts` (`createPlannerTaskAsync` composite flow)
- Test: `tests/unit/tools/planner.test.ts`, `tests/unit/graph/repository.test.ts`

**Approach:**
- `percent_complete` goes straight into the POST body (Graph accepts it at creation).
- `description`/`checklist` trigger a follow-up: after create, GET `/planner/tasks/{id}/details` for the details ETag, PATCH details via `withFreshEtag`. When the follow-up fails, return `{success: true, task_id, details_warning: '<message>'}` — never fail the whole create for a details error, and never leave the failure silent. The `details_warning` text must include the underlying Graph error AND the exact remediation ("retry the description/checklist via update_planner_task_details with task_id <pt_…>") so an agent caller has a specified recovery path.
- Checklist input mirrors `update_planner_task_details`'s existing shape (record of GUID → `{title, isChecked}` passthrough objects).

**Test scenarios:**
- Happy path: create with title only → single POST, no details call.
- Happy path: create with `description` → POST, then details GET, then details PATCH with the details ETag (assert ordering).
- Happy path: create with `percent_complete: 50` → in POST body, no details call.
- Error path: details PATCH fails after successful create → result still `success: true` with `task_id` and `details_warning`; the warning names `update_planner_task_details` and the task token as the recovery path.
- Edge case: `percent_complete: 100` accepted; `101` rejected by zod.

**Verification:**
- Unit tests green; create-only path makes exactly one Graph call (no regression for the simple case).

---

### U7. Remote surface, docs, and pinned counts

**Goal:** Everything stays truthful: remote surface widened by explicit review, README/provisioning docs updated, tool-count assertions bumped. PR reviewers must diff the five added tool names, not just accept the new count — bumping the assertion in the same PR neutralizes it as an independent guard for this PR.

**Requirements:** R7, R8

**Dependencies:** U2, U3, U4, U5, U6 (final tool set must be settled)

**Files:**
- Modify: `src/remote/entitlements.ts` (add `get_plan_details`, `update_plan_details` to `DEFAULT_TOOL_SURFACE`; deliberately omit the delete-plan pair AND `update_plan_sharing` — add a comment noting both omissions are intentional blast-radius calls), `tests/e2e/mcp-client.test.ts` (250 → 255 in all three assertions), `README.md` (headline "250 tools" → 255 at lines ~9/224/644; Features-Overview Planner row AND the Total row at ~line 42 `**250**` → `**255**`; Planner detail table +5 rows and `(23)` → `(28)`; fix the stale ETag paragraph at ~line 656 to describe fetch-fresh-before-write, not "caches ETags"; permissions table: `Tasks.ReadWrite` purpose → "Read and manage To Do tasks and Planner"), `docs/remote/provisioning.md` (replace the "Planner write is intentionally not requested" callout with: Tasks.ReadWrite now requested via jp-infrastructure; post-apply admin re-consent required — keep the existing `az ad app permission admin-consent` command as the how, and add the connector-restart step from Operational Notes)
- Test: `tests/unit/remote/entitlements.test.ts` (contract auto-verifies new pinned names exist)

**Approach:**
- Single unit at the end so counts are computed once. Do not attempt to fix the README Features-Overview table's pre-existing global miscounts beyond the Planner row and totals actually touched — flag remaining staleness in the PR description instead of scope-creeping.

**Test scenarios:**
- Test expectation: existing contract/e2e suites are themselves the tests — entitlements contract (every pinned name resolves), e2e tool-count equality, invariants over the four new tools.

**Verification:**
- Full `npm test` + `lint` + `typecheck` green; `grep -nE '\b250\b' README.md` returns no tool-count matches (catches the bare Total-row `**250**`, not just the "250 tools" phrase).

---

## System-Wide Impact

- **Interaction graph:** New repository methods feed only the planner toolset; visualization (`getPlanVisualizationDataAsync`) composes existing list methods and is untouched (labels deliberately not rendered yet). `what_changed` has no Planner integration — untouched.
- **Error propagation:** All new writes go through `withFreshEtag` — empty-ETag throws loudly; 412 retries once; Graph 400s (bad orderHint echo, bad category key that slipped past zod) surface via the existing error mapping in `src/utils/errors.ts`. Planner 403s are overloaded three ways — missing scope, service limits (e.g. `MaximumPlannerPlans`), AND group membership: delegated writes only reach plans whose owning M365 group the signed-in user belongs to, so membership 403s are expected (not a regression) for non-members even after the scope fix. Don't remap any of these to "permission denied" blindly.
- **State lifecycle risks:** U6's composite create can partially succeed (task without details) — surfaced via `details_warning`, never silent. delete_plan approval tokens are single-replica/store-bound (documented constraint, unchanged).
- **API surface parity:** Remote pinned surface intentionally lags the full registry (delete-plan pair excluded) — this is the designed behavior of R7 pinning, restated in an entitlements comment.
- **Integration coverage:** Contract invariants suite executes every new handler against a proxy context; e2e client test re-counts the registry.
- **Unchanged invariants:** Existing planner tool schemas gain only optional fields — no breaking changes for current callers; `GRAPH_SCOPES` (local) unchanged; PRM/OAuth surface unchanged.

---

## Risks & Dependencies

| Risk | Mitigation |
|------|------------|
| Admin consent is manual — infra PR merges but writes still 403 until a JP tenant admin re-consents, and MSAL's OBO cache serves stale pre-consent tokens for up to ~90 min after that | PR body + provisioning.md call out consent + connector restart + membership-scoped verification write as one runbook sequence; U7 doc change keeps the runbook accurate |
| jp-infrastructure repo has an unrelated feature branch checked out locally | Work in a fresh `git worktree` from `origin/main`; never touch the existing checkout |
| Graph PATCH-merge semantics for `appliedCategories`/`sharedWith` subtly differ from docs | Tests assert our request bodies only (docs-verified); probe-before-defenses learning says don't pre-build correction logic — surface Graph errors |
| U2/U5/U6 touch the same functions | Sequence them (dependency-ordered), don't parallelize edits to `planner.ts`/`repository.ts` |
| Tool-count assertion enforced only on some CI legs | Run full local `npm test` before push; adversarial review hunt list includes the count |

---

## Documentation / Operational Notes

- After the jp-infrastructure PR is applied: tenant admin runs `az ad app permission admin-consent --id 484c0657-6a05-4aad-a175-dabac48acb05`, then verifies with `az ad app permission list-grants`. **Then restart the connector replica** — MSAL's in-memory OBO cache keeps serving each user's pre-consent Graph token (missing the `Tasks.ReadWrite` scp claim) until it expires (~60–90 min) or the process restarts, so "retry after consent" alone still 403s and misdiagnoses the fix. No user re-auth is expected (`.default` flows through).
- R1 acceptance: after consent + restart, perform one Planner write against a plan **whose owning group the tester is a member of** (delegated writes require membership regardless of scope) before declaring the scope fix done. Sequencing note: the infra apply + consent + restart can happen before or after the npm release — remote Planner writes 403 today either way — but the pilot announcement should wait for both.
- Release: this repo follows merge-to-main + tag; the new tools ship in the next minor (v4.5.0 candidate).
- Cairn vault write-back at session end per project CLAUDE.md.

---

## Sources & References

- Audit context: this session's Planner audit (2026-08-06), logged in Cairn `20-projects/JBC-MCP-Office365.md`.
- Origin intent: `docs/brainstorms/2026-07-11-remote-connector-mode-requirements.md` (R7, A2 — Planner write was always intended).
- Related code: `src/tools/planner.ts`, `src/graph/repository.ts`, `src/graph/client/graph-client.ts`, `src/remote/entitlements.ts`, jp-infrastructure `stacks/azure/entra/mcp-office365-connector/main.tf`.
- External docs: learn.microsoft.com Graph v1.0 Planner API pages (permissions verified 2026-08-06); `planner-order-hint-format`; `plannercategorydescriptions`.
