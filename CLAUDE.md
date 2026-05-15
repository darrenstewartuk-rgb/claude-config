# Claude Code — Global Configuration

This directory contains Claude Code's global skills, commands, settings, and memory.
It is version-controlled and backed up to GitHub.

---

## GitHub Backup

Repo: https://github.com/darrenstewartuk-rgb/claude-config — push with `cd "C:/Users/Darren/.claude" && git add -A && git commit -m "Update: <describe what changed>" && git push`

---

## Key Paths (SGM)

| Resource | Path |
|---|---|
| Parts database | `S:\SGMWindows\Customer Care\2026 Parts Lists\DataBaseSearch.xlsm` — do not read directly; use /parts-data skill which reads source xlsx files instead |
| Parts list style ref | `C:\Users\Darren\Downloads\Parts_Top10_Report_Mar2026.html` |
| Report output folder | `S:\SGMWindows\Customer Care\Reports\` |
| PhotoCluster app | `C:\Users\Darren\OneDrive - SGM Windows\Desktop\Folders\PhotoCluster.exe` — standalone GUI for visual photo search and grouping. Source: `C:\Users\Darren\photo_cluster_app.py` |

---

## General Preferences

- No emojis unless requested
- When finished with a report: always open in Chrome via local server (`python -m http.server 8765`)