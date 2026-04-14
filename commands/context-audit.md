---
name: context-audit
description: >
  Audit your Claude Code setup for token waste and context bloat. Use when
  the user says "audit my context", "check my settings", "why is Claude so
  slow", "token optimization", "context audit", or runs /context-audit.
  Starts by running /context to see real overhead, then audits MCP servers,
  CLAUDE.md rules, skills, settings, and file permissions. Returns a
  health score with specific fixes.
user-invocable: true
---

# Usage Audit

Bloated context costs more and produces worse output. This skill finds
the waste and tells you what to cut.

## Step 1: Get /context Data

Check the conversation history for /context output. If the user already
ran /context in this session, use that data. If not, ask:

"Run /context in this session terminal and let me know when you're done. I can't run
slash commands myself, but once I can see the breakdown I'll audit
everything it flags."

STOP HERE. Do NOT proceed to Step 2 until the user has ran /context. The context breakdown determines
what to audit and in what order. Without it, the audit is guessing.
Output the message above and wait for the user's next message.

## Step 2: Audit What's Bloated

Based on the /context output, audit each category from largest to
smallest. Run checks in parallel where possible.

### MCP Servers

Each server loads full tool definitions into context every turn
(~15,000-20,000 tokens each).

- Count configured servers from settings.json
- Flag any with CLI alternatives (Playwright, Google Workspace, GitHub
  all have CLIs that cost zero tokens when idle)
- Report total MCP overhead from /context output

### CLAUDE.md

Read all CLAUDE.md files (project root, .claude/, ~/.claude/).
Count lines. Then read every rule and test against five filters:

| Filter | Flag when... |
|--------|-------------|
| Default | Claude already does this without being told ("write clean code", "handle errors") |
| Contradiction | Conflicts with another rule in same or different file |
| Redundancy | Repeats something already covered elsewhere |
| Bandaid | Added to fix one bad output, not improve outputs generally |
| Vague | Interpreted differently every time ("be natural", "use good tone") |

If total CLAUDE.md lines > 200, check for progressive disclosure
opportunities: rules that only apply to specific tasks (API conventions,
deployment steps, testing guidelines) should move to reference files
with one-line pointers. Only recommend splitting when the file is
actually bloated -- a lean CLAUDE.md with universal context is fine
as a single file.
