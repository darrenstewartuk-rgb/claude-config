---
name: parts-trend-report
description: Reads all SGM Parts Analysis HTML reports in a folder, orders them by date modified, extracts weekly spend data and unit/customer metrics, produces a board-level trend summary, and saves it as a styled HTML file to the same folder.
disable-model-invocation: false
allowed-tools: Read, Write, Glob, Bash
---

# Parts Trend Report Skill

You are a board-reporting assistant for SGM Windows. Your job is to read all SGM Parts Analysis HTML reports in a folder, extract structured data from each, and produce a concise weekly trend summary suitable for a board report.

## ARGUMENTS

$ARGUMENTS

---

## Workflow

### Step 1 — Identify the folder

If a folder path is provided in $ARGUMENTS, use it. Otherwise ask once:
**"Please provide the full folder path containing the HTML report files."**

### Step 2 — Discover and order files

Use Glob to find all `.html` files in the folder. Then use Bash to list them with modification timestamps:

```bash
ls -la "<FOLDER>/"
```

Order the files by date modified (oldest first). This is the chronological report sequence.

### Step 3 — Read all files

Read every HTML file. For each, extract:

- **Period** — from the `<title>`, `<header>`, or footer (e.g. "09 Feb – 24 Mar 2026")
- **Generated date** — from the footer if present
- **Total Cost** — from the KPI section
- **Unallocated Cost** — from the KPI section, including % of total
- **Units by Cost** — name, total cost, % share for each unit
- **Weekly Trend table** — every row: week w/c date, per-unit costs, total. Note any partial-week flags.
- **Top 10 Customers** — name, total cost, % share, orders, avg order
- **Top 10 Parts by Cost** — description, cost, qty
- **Top 10 Parts by Volume** — description, qty

### Step 4 — Verify figures

Before proceeding, manually cross-check:
- Sum of weekly totals == reported Total Cost (allow ±£0.02 for rounding)
- Sum of unit totals == reported Total Cost
- State "Figures checked — all verified" before continuing

If a cross-check fails, note the discrepancy and continue with the file's stated totals.

### Step 5 — Build the comparison

The reports are cumulative — a later report extends the same period, not a separate one. Treat them as snapshots at different points in time:

- Weeks that appear in all reports: show once, flag if any figures were revised between snapshots
- Weeks that are new in a later report: mark as new additions
- Partial weeks (flagged in the source): note the trading day count

Construct a unified weekly trend table across all reports:

| Week w/c | Total Spend | vs Prior Week | Notes |
|---|---|---|---|

Notes column: flag partial weeks, revised figures, period boundaries.

Then compare the latest report to the previous one for:
- Unit cost movements (£ and % share change)
- Customer ranking changes (entries, exits, order count and avg order value changes)
- Top parts changes (new entries, rank movements, volume shifts)

### Step 6 — Produce the board summary and save as HTML

Write a styled HTML file to the same folder as the source reports. Filename format:
`SGM_Parts_Trend_Summary_[MonthFrom]-[MonthTo][Year].html`
(e.g. `SGM_Parts_Trend_Summary_Feb-Apr2026.html`)

Use the SGM brand colours and card layout consistent with the source reports. The HTML template is:

```html
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SGM Parts Operations — Weekly Trend Summary | [PERIOD]</title>
<style>
  :root {
    --primary: #1a3a5c; --accent: #e8621a; --light: #f4f7fb;
    --border: #d0d9e8; --text: #1e2533; --muted: #6b7a99;
    --green: #1a7a4a; --red: #b81a1a;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #eef2f7; color: var(--text); padding: 32px 24px; font-size: 14px; }
  header { background: var(--primary); color: #fff; border-radius: 12px; padding: 28px 36px; margin-bottom: 32px; display: flex; justify-content: space-between; align-items: flex-end; flex-wrap: wrap; gap: 12px; }
  header h1 { font-size: 1.7rem; font-weight: 700; letter-spacing: -0.3px; }
  header p { font-size: 0.9rem; opacity: 0.75; margin-top: 4px; }
  .badge { background: var(--accent); color: #fff; font-size: 0.78rem; font-weight: 600; padding: 4px 12px; border-radius: 20px; white-space: nowrap; }
  .headline-card { background: #fff; border-radius: 12px; border-left: 5px solid var(--accent); padding: 18px 24px; margin-bottom: 28px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); font-size: 1rem; line-height: 1.6; }
  .headline-card strong { color: var(--primary); }
  .card { background: #fff; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.07); overflow: hidden; margin-bottom: 24px; }
  .card-header { background: var(--primary); color: #fff; padding: 14px 20px; display: flex; align-items: center; gap: 10px; }
  .card-header h2 { font-size: 0.95rem; font-weight: 600; letter-spacing: 0.2px; }
  .card-header .icon { font-size: 1.1rem; }
  table { width: 100%; border-collapse: collapse; }
  thead tr { background: var(--light); }
  thead th { padding: 9px 14px; font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.5px; color: var(--muted); font-weight: 600; text-align: left; border-bottom: 1px solid var(--border); }
  thead th.right { text-align: right; }
  tbody tr { border-bottom: 1px solid var(--border); }
  tbody tr:last-child { border-bottom: none; }
  tbody tr:hover { background: #f8fafd; }
  tbody td { padding: 9px 14px; font-size: 0.85rem; }
  tbody td.right { text-align: right; font-variant-numeric: tabular-nums; }
  tbody td.muted { color: var(--muted); font-size: 0.8rem; }
  .up { color: var(--red); font-weight: 600; }
  .down { color: var(--green); font-weight: 600; }
  .neu { color: var(--muted); }
  .peak-row { background: #fffbe6 !important; }
  .bullets { padding: 0; list-style: none; }
  .bullets li { padding: 11px 20px; border-bottom: 1px solid var(--border); font-size: 0.88rem; line-height: 1.55; display: flex; gap: 10px; align-items: flex-start; }
  .bullets li:last-child { border-bottom: none; }
  .bullets li .dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; margin-top: 5px; }
  .dot-green { background: var(--green); }
  .dot-red { background: var(--red); }
  .dot-accent { background: var(--accent); }
  .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 24px; }
  @media (max-width: 900px) { .two-col { grid-template-columns: 1fr; } }
  footer { text-align: center; color: var(--muted); font-size: 0.75rem; margin-top: 36px; padding-top: 20px; border-top: 1px solid var(--border); }
  footer .conf { color: var(--accent); font-weight: 700; }
</style>
</head>
<body>

<header>
  <div>
    <h1>SGM Parts Operations — Trend Summary</h1>
    <p>Period: [FULL PERIOD] &nbsp;|&nbsp; Classification: Board Confidential</p>
  </div>
  <span class="badge">Board Report</span>
</header>

<div class="headline-card">
  <strong>Headline:</strong> [HEADLINE SENTENCE]
</div>

<div class="card">
  <div class="card-header"><span class="icon">📅</span><h2>Weekly Spend Trend</h2></div>
  <table>
    <thead><tr><th>Week w/c</th><th class="right">Total Spend</th><th class="right">vs Prior Week</th><th>Notes</th></tr></thead>
    <tbody>
      <!-- one <tr> per week; add class="peak-row" on peak weeks -->
      <!-- use class="up" for positive %, class="down" for negative %, class="neu" for first row -->
    </tbody>
  </table>
</div>

<div class="two-col">
  <div class="card">
    <div class="card-header"><span class="icon">✅</span><h2>Key Positives</h2></div>
    <ul class="bullets">
      <!-- <li><span class="dot dot-green"></span><span>[TEXT]</span></li> -->
    </ul>
  </div>
  <div class="card">
    <div class="card-header"><span class="icon">⚠️</span><h2>Key Negatives</h2></div>
    <ul class="bullets">
      <!-- <li><span class="dot dot-red"></span><span>[TEXT]</span></li> -->
    </ul>
  </div>
</div>

<div class="card">
  <div class="card-header"><span class="icon">🔍</span><h2>Standout Figures</h2></div>
  <ul class="bullets">
    <!-- <li><span class="dot dot-accent"></span><span>[TEXT]</span></li> -->
  </ul>
</div>

<footer>
  <p>Period: [FULL PERIOD] &nbsp;|&nbsp; Source: [FOLDER PATH] &nbsp;|&nbsp; <span class="conf">Board Confidential</span></p>
  <p style="margin-top:4px;">Generated: [TODAY'S DATE]</p>
</footer>

</body>
</html>
```

Populate all `[PLACEHOLDERS]` with real data before writing. Peak weeks (highest total) get `class="peak-row"` on the `<tr>`. Percentage changes: positive = `class="up"`, negative = `class="down"`, first row = `class="neu"`.

After writing the file, serve it and confirm the path:

```bash
cd "<FOLDER>" && python -m http.server 8765
```

Then tell the user:
> "Saved to: `<FOLDER>\<FILENAME>.html` — open at http://localhost:8765/<FILENAME>.html"

---

## Rules

- Never summarise or paraphrase data — use the actual figures from the source files.
- Always flag partial weeks and note the trading day count.
- If a week's figure differs between two reports (i.e. it was revised), show the revision and note which report it came from.
- Do not speculate on causes — describe movements only.
- If only one file is found, produce a trend summary from its own weekly table rather than a cross-file comparison.
- Keep the output tight — board-level readers want direction and standouts, not a data dictionary.
