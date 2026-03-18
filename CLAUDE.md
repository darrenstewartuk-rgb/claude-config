# Claude Code — Global Configuration

This directory contains Claude Code's global skills, commands, settings, and memory.
It is version-controlled and backed up to GitHub.

---

## GitHub Backup

- **Repository:** https://github.com/darrenstewartuk-rgb/claude-config
- **Clone URL:** https://github.com/darrenstewartuk-rgb/claude-config.git
- **Account:** darrenstewartuk-rgb

### How to push updates to GitHub

Run these commands from `C:\Users\Darren\.claude\` after making any changes to skills, commands, or settings:

```bash
cd "C:/Users/Darren/.claude"
git add -A
git commit -m "Update: <describe what changed>"
git push
```

Or use the shortcut — just tell Claude:
> "Back up my Claude config to GitHub"

Claude will stage, commit, and push all changes in `C:\Users\Darren\.claude\` automatically.

---

## Directory Structure

| Folder / File | Purpose |
|---|---|
| `commands/` | Custom slash commands (skills) — e.g. `/parts-data`, `/performance-pack-data-extractor` |
| `memory/` | Persistent memory files (user profile, feedback, project context) |
| `settings.json` | Global Claude Code settings and permissions |
| `settings.local.json` | Local overrides (not committed) |
| `CLAUDE.md` | This file — global instructions loaded into every conversation |

---

## Active Skills

| Skill | File | Purpose |
|---|---|---|
| `/parts-data` | `commands/parts-data.md` | Generates SGM Parts Analysis HTML report from DataBaseSearch.xlsm |
| `/performance-pack-data-extractor` | `commands/performance-pack-data-extractor.md` | Extracts weekly ticket data into styled Excel report |
| `/excel-editor` | `commands/excel-editor.md` | Formats and edits Excel workbooks |

---

## Key Paths (SGM)

| Resource | Path |
|---|---|
| Parts database | `S:\SGMWindows\Customer Care\2026 Parts Lists\DataBaseSearch.xlsm` |
| Parts list style ref | `C:\Users\Darren\Downloads\Parts_Top10_Report_Mar2026.html` |
| Report output folder | `S:\SGMWindows\Customer Care\Reports\` |

---

## General Preferences

- Keep responses concise and direct
- No emojis unless requested
- When generating reports: always validate totals before writing HTML
- When finished with a report: always open in Chrome via local server (`python -m http.server 8765`)
