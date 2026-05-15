---
name: capacity-planning-analysis
description: Interactive skill that asks for an Excel workbook path and extraction requirements, reads the data, and presents it as a styled SGM-brand HTML report. Use when the user wants to analyse capacity planning data from any Excel file.
disable-model-invocation: false
allowed-tools: Read, Write, Glob, Bash, AskUserQuestion
---

# Capacity Planning Analysis Skill

You are a data analyst assistant. Your job is to read an Excel workbook specified by the user, extract the data they describe, and produce a polished HTML report in the SGM house style.

## ARGUMENTS

$ARGUMENTS

---

## Workflow

### Step 1 — Ask for the file

If a file path was not supplied in $ARGUMENTS, ask the user:

> "Please provide the full path to the Excel workbook you want to analyse (e.g. `C:\Users\Darren\Documents\Capacity.xlsx`)."

Wait for the response before continuing.

### Step 2 — Discover the workbook structure

Once you have the file path, run the excel_editor helper to list available sheets:

```bash
python "C:/Users/Darren/.claude/skills/excel-editor/scripts/excel_editor.py" "{\"type\":\"list_sheets\",\"folder\":\"<FOLDER>\",\"workbook\":\"<FILENAME>\"}"
```

Where `<FOLDER>` is the directory portion of the path and `<FILENAME>` is just the filename.

Then preview each relevant sheet (rows 1–5) so you understand its structure:

```bash
python "C:/Users/Darren/.claude/skills/excel-editor/scripts/excel_editor.py" "{\"type\":\"read_range\",\"folder\":\"<FOLDER>\",\"workbook\":\"<FILENAME>\",\"sheet\":\"<SHEET>\",\"range\":\"A1:Z5\"}"
```

### Step 3 — Ask the user what to extract

Present what you found (sheet names and a brief description of each sheet's columns), then ask the following questions **in a single message**:

1. **Which sheet(s)** should be included? (list the discovered sheet names for them to choose from)
2. **Which columns** contain the key data? (confirm or let the user correct your auto-detected headers)
3. **What analysis is needed?** Choose any that apply:
   - Summary totals / KPI cards
   - Top N ranking (by which column? how many?)
   - Grouping / aggregation (group by which column?)
   - Week-by-week or period trend
   - Comparison between sheets or categories
   - Risks & Opportunities summary
4. **Output title** — what should the report be called? (default: "Capacity Planning Analysis")
5. **Where to save** — full folder path for the HTML output (default: `S:\SGMWindows\Customer Care\Reports\`)

Wait for the user's answers before continuing.

### Step 4 — Read the data

For each sheet and range the user specified, read all data rows:

```bash
python "C:/Users/Darren/.claude/skills/excel-editor/scripts/excel_editor.py" "{\"type\":\"read_range\",\"folder\":\"<FOLDER>\",\"workbook\":\"<FILENAME>\",\"sheet\":\"<SHEET>\",\"range\":\"A1:<LAST_COL><LAST_ROW>\"}"
```

If you do not know the last row, read a generous range (e.g. `A1:Z500`) and discard trailing empty rows in your analysis.

Row filtering:
- Row 1 is always the header row — use it to map column positions
- Skip rows where ALL key data columns are empty/None
- Cast numeric fields: `float(x)` where the cell is a number; treat blank as `0.0`
- Cast string fields: `str(x).strip()` where present; blank as `''`

### Step 5 — Compute the analysis

Based on the user's answers from Step 3, perform the requested computations:

**KPI cards** — compute totals, averages, counts, or percentages as appropriate for the data.

**Top N rankings** — group by the chosen column, sum the value column, sort descending, take top N. Compute bar widths as `value / max_value * 100%`.

**Grouping / aggregation** — group by the nominated column, aggregate with sum or count as appropriate.

**Trend table** — if dates are present, parse them and group by week (Monday start) using Python's `datetime`. Label weeks as `strftime('%d %b')`.

**Cross-sheet comparison** — if multiple sheets are selected, produce a per-sheet breakdown with a combined total.

**Risks & Opportunities summary** — scan the computed data and derive a two-column card (Risks left, Opportunities right). Use these rules to identify items:

*Risks — flag when:*
- Any day or unit has Status = SHORTFALL → severity HIGH
- Any day or unit has Status = WARNING and absences are the cause → severity MEDIUM
- A unit is in shortfall across multiple consecutive days → severity HIGH
- Absence rate on any day exceeds 20% of base headcount → severity HIGH if it caused a shortfall, MEDIUM otherwise
- Overtime is zero but a shortfall or near-shortfall exists → severity MEDIUM

*Opportunities — flag when:*
- Any week or period has utilisation below 60% → severity HIGH (recovery/pull-forward window)
- A bank holiday or zero-load day has non-zero available capacity → severity HIGH (overtime recovery buffer)
- Overtime is zero but capacity existed on units that remained operational during a shortfall → severity MEDIUM
- Workload is declining across consecutive days with large surplus → severity LOW (scheduling flexibility)

Each item must include: severity badge, a specific headline, and a 1–2 sentence description quoting the relevant figures (hours, %, headcount). Use MEDIUM or LOW only when genuinely justified — do not inflate severity.

Validate your totals before generating HTML:
- Sum of grouped values == grand total — state OK or WARNING
- Print each check result in your response

### Step 6 — Generate the HTML report

Produce a complete, self-contained HTML file using the SGM house style below. Populate every `[PLACEHOLDER]` with real data — never leave template text in the output.

```html
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>[REPORT TITLE] | [PERIOD OR DATE]</title>
<style>
  :root {
    --primary: #1a3a5c; --accent: #e8621a; --light: #f4f7fb;
    --border: #d0d9e8; --text: #1e2533; --muted: #6b7a99;
    --green: #1a7a4a; --red: #b81a1a;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #eef2f7; color: var(--text); padding: 32px 24px; font-size: 14px; }
  header { background: var(--primary); color: #fff; border-radius: 12px; padding: 28px 36px; margin-bottom: 8px; display: flex; justify-content: space-between; align-items: flex-end; flex-wrap: wrap; gap: 12px; }
  .accent-bar { height: 4px; background: var(--accent); border-radius: 0 0 4px 4px; margin-bottom: 28px; }
  header h1 { font-size: 1.7rem; font-weight: 700; letter-spacing: -0.3px; }
  header p { font-size: 0.9rem; opacity: 0.75; margin-top: 4px; }
  .badge { background: var(--accent); color: #fff; font-size: 0.78rem; font-weight: 600; padding: 4px 12px; border-radius: 20px; white-space: nowrap; }
  .kpi-row { display: flex; gap: 20px; flex-wrap: wrap; margin-bottom: 28px; }
  .kpi-card { background: #fff; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); padding: 20px 28px; flex: 1; min-width: 180px; }
  .kpi-card.warn { border-left: 5px solid var(--accent); }
  .kpi-card .label { font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.5px; color: var(--muted); font-weight: 600; margin-bottom: 6px; }
  .kpi-card .value { font-size: 1.9rem; font-weight: 700; color: var(--primary); }
  .kpi-card .sub { font-size: 0.8rem; color: var(--muted); margin-top: 4px; }
  .card-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(480px, 1fr)); gap: 24px; margin-bottom: 24px; }
  .card-full { margin-bottom: 24px; }
  .card { background: #fff; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.07); overflow: hidden; }
  .card-header { background: var(--primary); color: #fff; padding: 14px 20px; display: flex; align-items: center; gap: 10px; }
  .card-header h2 { font-size: 0.95rem; font-weight: 600; letter-spacing: 0.2px; }
  table { width: 100%; border-collapse: collapse; }
  thead tr { background: var(--light); }
  thead th { padding: 9px 14px; font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.5px; color: var(--muted); font-weight: 600; text-align: left; border-bottom: 1px solid var(--border); }
  thead th.right { text-align: right; }
  tbody tr { border-bottom: 1px solid var(--border); }
  tbody tr:last-child { border-bottom: none; }
  tbody tr:hover { background: #f8fafd; }
  tbody td { padding: 9px 14px; font-size: 0.85rem; vertical-align: middle; }
  tbody td.right { text-align: right; font-variant-numeric: tabular-nums; }
  tbody td.muted { color: var(--muted); font-size: 0.8rem; }
  .rank { display: inline-block; width: 24px; height: 24px; border-radius: 50%; font-size: 0.75rem; font-weight: 700; text-align: center; line-height: 24px; margin-right: 8px; }
  .rank-1 { background: #ffd700; color: #7a5800; }
  .rank-2 { background: #c0c0c0; color: #444; }
  .rank-3 { background: #cd7f32; color: #fff; }
  .rank-other { background: var(--light); color: var(--muted); }
  .bar-wrap { background: #eef2f7; border-radius: 4px; height: 10px; min-width: 80px; overflow: hidden; margin-top: 4px; }
  .bar { height: 10px; border-radius: 4px; }
  .bar-green { background: var(--green); }
  .bar-accent { background: var(--accent); }
  .up { color: var(--red); font-weight: 600; }
  .down { color: var(--green); font-weight: 600; }
  .neu { color: var(--muted); }
  .peak-row { background: #fffbe6 !important; }
  .peak-cell { font-weight: 700; color: var(--accent); }
  .ro-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 24px; }
  .card-header-risk { background: #7a1a1a; }
  .card-header-opp { background: var(--green); }
  .ro-item { display: flex; gap: 14px; align-items: flex-start; padding: 14px 20px; border-bottom: 1px solid var(--border); }
  .ro-item:last-child { border-bottom: none; }
  .ro-badge { flex-shrink: 0; font-size: 0.68rem; font-weight: 700; padding: 3px 10px; border-radius: 20px; letter-spacing: 0.4px; margin-top: 2px; white-space: nowrap; }
  .badge-high-risk { background: #fbe8e8; color: var(--red); }
  .badge-med-risk { background: #fff3e0; color: #b86a00; }
  .badge-high-opp { background: #e8f5ee; color: var(--green); }
  .badge-med-opp { background: #eef4fb; color: var(--primary); }
  .badge-low-opp { background: var(--light); color: var(--muted); }
  .ro-title { font-weight: 600; font-size: 0.85rem; margin-bottom: 4px; color: var(--text); }
  .ro-desc { font-size: 0.8rem; color: var(--muted); line-height: 1.5; }
  footer { text-align: center; color: var(--muted); font-size: 0.75rem; margin-top: 36px; padding-top: 20px; border-top: 1px solid var(--border); }
  footer .conf { color: var(--accent); font-weight: 700; }
</style>
</head>
<body>

<header>
  <div>
    <h1>[REPORT TITLE]</h1>
    <p>[SUBTITLE OR PERIOD] &nbsp;|&nbsp; Classification: Board Confidential</p>
  </div>
  <span class="badge">Capacity Planning</span>
</header>
<div class="accent-bar"></div>

<!-- KPI CARDS ROW -->
<div class="kpi-row">
  <!-- Example KPI card:
  <div class="kpi-card">
    <div class="label">Total Units</div>
    <div class="value">1,234</div>
    <div class="sub">across all categories</div>
  </div>
  Add class="warn" for any card that needs orange-left-border highlighting -->
</div>

<!-- CARD GRID (two-column where content warrants) -->
<div class="card-grid">

  <!-- Example ranking card:
  <div class="card">
    <div class="card-header"><h2>Top 10 — [CATEGORY] by [METRIC]</h2></div>
    <table>
      <thead><tr><th>Rank</th><th>[LABEL]</th><th class="right">[VALUE]</th><th class="right">Share</th><th>Volume</th></tr></thead>
      <tbody>
        <tr>
          <td><span class="rank rank-1">1</span></td>
          <td>[NAME]</td>
          <td class="right">[VALUE]</td>
          <td class="right">[X.X%]</td>
          <td><div class="bar-wrap"><div class="bar bar-accent" style="width:[W]%"></div></div></td>
        </tr>
      </tbody>
    </table>
  </div> -->

</div>

<!-- RISKS & OPPORTUNITIES (if requested) -->
<!-- <div class="ro-grid">
  <div class="card">
    <div class="card-header card-header-risk"><h2>Risks</h2></div>
    <div class="ro-item">
      <span class="ro-badge badge-high-risk">HIGH</span>
      <div>
        <div class="ro-title">[RISK HEADLINE]</div>
        <div class="ro-desc">[DESCRIPTION WITH SPECIFIC FIGURES]</div>
      </div>
    </div>
  </div>
  <div class="card">
    <div class="card-header card-header-opp"><h2>Opportunities</h2></div>
    <div class="ro-item">
      <span class="ro-badge badge-high-opp">HIGH</span>
      <div>
        <div class="ro-title">[OPPORTUNITY HEADLINE]</div>
        <div class="ro-desc">[DESCRIPTION WITH SPECIFIC FIGURES]</div>
      </div>
    </div>
  </div>
</div> -->
Severity badge classes: badge-high-risk / badge-med-risk for risks; badge-high-opp / badge-med-opp / badge-low-opp for opportunities.

<!-- FULL-WIDTH TREND TABLE (if applicable) -->
<!-- <div class="card card-full">
  <div class="card-header"><h2>Weekly / Period Trend</h2></div>
  <table>
    <thead><tr><th>Period</th><th class="right">Total</th><th class="right">vs Prior</th><th>Notes</th></tr></thead>
    <tbody>
      peak row: class="peak-row"; peak cell: class="peak-cell"
      change direction: class="up" / class="down" / class="neu"
    </tbody>
  </table>
</div> -->

<footer>
  <p>Source: [FILE PATH] &nbsp;|&nbsp; <span class="conf">Board Confidential</span></p>
  <p style="margin-top:4px;">Generated: [TODAY'S DATE]</p>
</footer>

</body>
</html>
```

**HTML generation rules:**
- Replace every `[PLACEHOLDER]` with real data — no template text survives into the output
- Gold/silver/bronze rank badges (`rank-1`, `rank-2`, `rank-3`) for positions 1–3; `rank-other` for the rest
- Bar chart widths: `value / max_value * 100` — use `.bar-green` for volume/units bars, `.bar-accent` for cost/value bars
- Peak row (highest value row in a trend table): add `class="peak-row"` on `<tr>` and `class="peak-cell"` on the highest-value `<td>`
- Trend percentage changes: positive = `class="up"`, negative = `class="down"`, first row = `class="neu"`
- Remove any HTML comment blocks (`<!-- ... -->`) from the final output
- Include only the sections the user requested — do not emit empty cards

### Step 7 — Save and open

Determine the output filename from the report title and today's date:
`CapacityPlanning_[ShortTitle]_[DDMMMYYYY].html`
(e.g. `CapacityPlanning_ProductionCapacity_08May2026.html`)

Write the file to the user's chosen output folder (default `S:\SGMWindows\Customer Care\Reports\`).

Then start the local server and open the file:

```bash
PYTHONIOENCODING=utf-8 python -m http.server 8765 &
```

Run from the output folder, then tell the user:

> "Saved to: `<FOLDER>\<FILENAME>.html` — open at `http://localhost:8765/<FILENAME>.html`"

---

## Rules

- Never guess column meanings — confirm with the user if the headers are ambiguous
- Never leave placeholder text (`[PLACEHOLDER]`, `<!-- comment -->`) in the saved HTML
- Always validate totals before generating HTML and report the check results
- Do not add sections the user did not ask for
- Keep KPI values and table figures exactly as computed — no rounding unless the user asks
- If a numeric field is blank in the source, treat it as zero — do not skip the row unless ALL key fields are empty
