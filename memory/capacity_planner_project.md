---
name: SGM Capacity Planner Project
description: Ongoing Excel-based capacity planning tool for SGM Windows factory — location, status, and outstanding work
type: project
---

**File:** `C:\Users\Darren\OneDrive - SGM Windows\Desktop\Folders\New folder\SGM_Capacity_Planner.xlsm`
**Project log:** `C:\Users\Darren\OneDrive - SGM Windows\Desktop\Folders\New folder\CAPACITY_PLANNER_PROJECT.md`

**Why:** Tool to plan daily/weekly production capacity across 6 factory units, inputting units-to-produce per product type and outputting required vs available hours with SHORTFALL flags.

**How to apply:** Always read CAPACITY_PLANNER_PROJECT.md at the start of any session working on this — it is the authoritative record of bugs fixed and outstanding work.

## Status as at 2026-03-26 (v2 — complete redesign)
- Complete redesign to single-row-per-day structure covering all 257 days (Wk 1–52)
- Config: B–G = user-definable dept %, H = total hrs/unit; Customer Care removed
- Scenario Planner: 257 data rows pre-populated from Con_Diary; formulas for O–S
- Weekly Summary: SUMIF formulas aggregate 52 weeks
- Macro rewritten: opens Con_Diary.xlsx from same folder, no file picker
- Macro code is in Macro Setup sheet (text only) — user must install via ALT+F11

## Key outstanding items
- Hours/unit values in Config col H are placeholders — must be set to real values
- Dept % column headers (B–G row 4) need labelling with real dept names
- Con_Diary.xlsx Sheet1 needs renaming to "Con_Diary"
- Conditional formatting for Status column not yet applied
- Weekly Summary col B (date range) blank — manual fill or macro extension needed
