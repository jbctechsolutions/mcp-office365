## Vault write-back

At the end of any session with a meaningful decision, status change, or new fact about this project, append a dated 1–3 line entry under `## Log` in `~/vaults/cairn/20-projects/JBC-MCP-Office365.md`, then commit via Cairn's normal git flow. Never write secrets; restricted content (donor PII, personnel specifics, private financials) never enters the vault.

## Documented Solutions

`docs/solutions/` — documented solutions to past problems (bugs, best practices, architecture/design patterns, workflow learnings), organized by category with YAML frontmatter (`module`, `tags`, `problem_type`). Relevant when implementing or debugging in documented areas.
