# Parts Data Report

Generate a styled HTML Top 10 report from SGM parts operations data.

---

## Source Data

- **File:** `S:\SGMWindows\Customer Care\2026 Parts Lists\DataBaseSearch.xlsm`
- **Sheet to use:** `Summary` — this sheet aggregates all unit tabs with a `Source Tab` column added
- **Do not** use Downloads folder, individual Parts List files, or any other sheet

### Summary sheet column structure (in order):
| Col | Name | Notes |
|-----|------|-------|
| A | Source.Name | Filename of the daily parts list (e.g. `Parts List - 02.03.26.xlsx`) |
| B | Part Code | SKU/part number — may be blank |
| C | Part Description | Free-text description |
| D | Quantity | Numeric |
| E | Size | Free-text |
| F | Customer | Customer name — may be blank or `-` for unallocated items |
| G | Job Number | Numeric job ref |
| H | Delivery Method | Free-text |
| I | Parts received? | YES/blank |
| J | DPD Code | Numeric |
| K | Paperwork? | YES/blank |
| L | Cost per item | Numeric |
| M | Total Cost | Numeric — **primary cost field used throughout** |
| N | Source Tab | Unit name: `Unit 1`, `Unit 3`, `Unit 4`, `Unit 5`, `Unit 6`, or `Customer Care` |

### Row filtering rules:
- Skip the header row (`Source.Name == 'Source.Name'`)
- Skip subtotal/blank rows — rows where **both** Part Code and Part Description are empty/None
- Keep all other rows, including those with a description but no part code

---

## Steps

1. **Load the workbook** using `openpyxl` with `read_only=True, data_only=True, keep_vba=False`. Cast all fields to safe types:
   - Strings: `str(x) if x else ''` for code, desc, cust, tab, src
   - Numerics: `float(x) if isinstance(x, (int, float)) else 0` for qty, cpi, tc

2. **Ask these questions before building:**
   - **Date range:** Filter by a specific date or date range, or use all available data?
     - Dates are parsed from `Source.Name` using regex `(\d{2})\.(\d{2})\.(\d{2})`
     - If filtered: reflect date range in report header and filename
   - Which sections to include? (default: Parts by Cost, Parts by Volume, Customers by Cost, Units by Cost, Weekly Trend)
   - Top 10 or different number?

3. **Extract the following data:**

   **KPI summary:**
   - Total Cost = `sum(tc)` across all rows
   - Transactions = total row count
   - Customers Served = count of unique non-empty, non-`-` customer names
   - Unique Parts = count of unique `(code + '|' + desc)` combinations
   - Unallocated Cost = `sum(tc)` for rows where `cust` is `''`, `'-'`, or `'None'`

   **Top 10 parts by cost:** group by description (fall back to code), sum qty and tc, sort by tc desc

   **Top 10 parts by volume:** same grouping, sort by qty desc; bar width = `qty/max_qty * 100%`

   **Top 10 customers by cost:** group by customer name, exclude blank/`-`/`None` rows; track unique source filenames as order count; compute % of total and avg order

   **Units by cost:** group by `Source Tab`; bar width = `unit_tc/max_unit_tc * 100%`

   **Weekly trend:** group by calendar week (Monday start) using `date - timedelta(days=date.weekday())`, label as `strftime('%d %b')`; columns = units; rows = weeks sorted chronologically; flag partial weeks (< 5 trading days) in footnote

4. **Validate totals — run all checks before generating HTML:**
   - `sum(all tc rows)` == Total Cost KPI  ✓/✗
   - `sum(unit totals)` == Total Cost KPI  ✓/✗
   - `sum(weekly totals per unit)` == each unit's total  ✓/✗
   - `weekly grand total` == Total Cost KPI  ✓/✗
   - Print each check to console with actual vs expected values
   - Raise a warning (do not abort) if any discrepancy > £0.01

5. **Generate the HTML** using the SGM house style:
   - Dark navy header (`#1a3a5c`) with orange accent (`#e8621a`)
   - CSS variables: `--primary:#1a3a5c`, `--accent:#e8621a`, `--light:#f4f7fb`, `--border:#d0d9e8`, `--text:#1e2533`, `--muted:#6b7a99`, `--green:#1a7a4a`
   - KPI cards row at top — warn card (orange border + value) for Unallocated Cost ⚠
   - Two-column card grid (`minmax(480px,1fr)`), weekly trend full-width (`grid-column: 1 / -1`) at bottom
   - Gold/silver/bronze rank badges for positions 1–3
   - Inline bar charts for units table (green bars) and volume table (accent/orange bars)
   - Peak week cell highlighted `style="background:#fffbe6"` with 🔝 on the highest unit that week
   - Footnote on weekly trend table flagging partial weeks with `*` / `†`
   - Footer: period, source file path, classification "Board Confidential", generated date
   - Data only — no action signal pills, no warning flags on customer names

6. **Save** as `SGM_Parts_Analysis_[DateRange].html` in `C:\Users\Darren\Downloads\`
   - Date range format: `Feb-Mar2026` for full months, `DD-DDMonYYYY` for filtered ranges

7. **Open in Chrome:**
   - Run `python -m http.server 8765` in background from `C:\Users\Darren\Downloads\`
   - Open with `start "" "http://localhost:8765/<filename>.html"`
   - Use Chrome browser tools to take a screenshot confirming it loaded correctly

---

## Style reference

Canonical style: `C:\Users\Darren\Downloads\Parts_Top10_Report_Mar2026.html`
Match its CSS variables, card layout, table structure, and rank badge colours exactly.

---

## Reference output

The report produced on 18 March 2026 for period 06 Feb – 18 Mar 2026 is saved at:
`C:\Users\Darren\Downloads\SGM_Parts_Analysis_Feb-Mar2026.html`

Key figures from that run (use to sanity-check future runs against same dataset):
- Total Cost: £6,234.35 | Transactions: 436 | Customers: 126 | Unique Parts: 281
- Unallocated Cost: £1,386.41
- Unit 6: £2,532.95 (40.6%) | Unit 5: £1,774.87 (28.5%) | Unit 4: £916.00 (14.7%)
- Top part by cost: Threshold Tray Silver — £268.60
- Top part by volume: 28mm Glazing Platform — 100 units
- Top customer: Stratton — £543.39
- Peak week: 09–13 Feb — £1,302.53
