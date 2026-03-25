# Parts Data Report

Generate a styled HTML Top 10 report from SGM parts operations data.

---

## Source Data

- **Source folder:** `S:\SGMWindows\Customer Care\2026 Parts Lists\Completed Sheets Here`
- **Files:** All `Parts List - DD.MM.YY.xlsx` files in that folder
- **Sheets to read per file:** `Unit 1`, `Unit 3`, `Unit 4`, `Unit 5`, `Unit 6`, `Customercare Office`
- **Do not** use `DataBaseSearch.xlsm` Summary sheet — Power Query cached values are unreliable when source files have been modified by openpyxl

### Source file column structure (row 1 = headers, data starts row 2):
| Col | Index (0-based) | Name | Notes |
|-----|-----------------|------|-------|
| B | 1 | Part Code | SKU/part number — may be blank |
| C | 2 | Part Description | Free-text description |
| D | 3 | Quantity | Numeric |
| E | 4 | Size | Free-text |
| F | 5 | Customer | Customer name — may be blank or `-` for unallocated items |
| G | 6 | Job Number | Numeric job ref |
| H | 7 | Delivery Method | Free-text |
| I | 8 | Parts received? | YES/blank |
| J | 9 | DPD Code | Numeric |
| K | 10 | Paperwork? | YES/blank |
| L | 11 | Cost per item | Numeric — XLOOKUP result; may be missing in openpyxl-modified files |
| M | 12 | Total Cost | Numeric — **primary cost field** — formula `=D*L`; may be missing in openpyxl-modified files |

### Sheet name → Source Tab mapping:
| Sheet name | Source Tab label |
|------------|-----------------|
| Unit 1 | Unit 1 |
| Unit 3 | Unit 3 |
| Unit 4 | Unit 4 |
| Unit 5 | Unit 5 |
| Unit 6 | Unit 6 |
| Customercare Office | Customer Care |

### Row filtering rules:
- Skip header row (row 1)
- Skip rows where **both** Part Code and Part Description are empty/None
- Keep all other rows

---

## Steps

1. **Load source files** using `glob` to find all `Parts List - *.xlsx` in the source folder. For each file:
   - Parse date from filename using regex `(\d{2})\.(\d{2})\.(\d{2})` → `datetime(2000+YY, MM, DD)`
   - Skip files outside the requested date range
   - Open with `openpyxl.load_workbook(fpath, read_only=True, data_only=True)`
   - **Build price lookup from `'.'` sheet** (for cost fallback): read rows from `ws['.']`, col A = part code, col B = price; store as `{code: price}` dict (~9,000+ entries)
   - Read each unit sheet (see mapping above), skip row 1, collect rows

   **Cost calculation — three-tier fallback (in order):**
   1. Use `r[12]` (Total Cost, col M) if non-zero
   2. Else if `r[11]` (Cost per item, col L) > 0: `tc = qty * cpi`
   3. Else if part code exists in price lookup: `cpi = price_lookup[code]; tc = qty * cpi`

   Cast all fields:
   - Strings: `str(x).strip() if x is not None else ''` for code, desc, cust
   - Numerics: `float(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else 0.0` for qty, cpi, tc

2. **Ask these questions before building:**
   - **Date range:** Filter by a specific date or date range, or use all available data?
     - If filtered: reflect date range in report header and filename
   - Which sections to include? (default: Parts by Cost, Parts by Volume, Customers by Cost, Units by Cost, Weekly Trend)
   - Top 10 or different number?

3. **Extract the following data:**

   **KPI summary (two cards only):**
   - Total Cost = `sum(tc)` across all rows
   - Unallocated Cost = `sum(tc)` for rows where `cust` is `''`, `'-'`, or `'None'`

   **Top 10 parts by cost:** group by description (fall back to code), sum qty and tc, sort by tc desc

   **Top 10 parts by volume:** same grouping, sort by qty desc; bar width = `qty/max_qty * 100%`

   **Top 10 customers by cost:** group by customer name, exclude blank/`-`/`None` rows; track unique source filenames as order count; compute % of total and avg order

   **Units by cost:** group by Source Tab; bar width = `unit_tc/max_unit_tc * 100%`

   **Weekly trend:** group by calendar week (Monday start) using `date - timedelta(days=date.weekday())`, label as `strftime('%d %b')`; columns = units; rows = weeks sorted chronologically; flag partial weeks (< 5 trading days) in footnote

4. **Validate totals — run all checks before generating HTML:**
   - `sum(all tc rows)` == Total Cost KPI  OK/WARNING
   - `sum(unit totals)` == Total Cost KPI  OK/WARNING
   - `sum(weekly totals per unit)` == each unit's total  OK/WARNING
   - `weekly grand total` == Total Cost KPI  OK/WARNING
   - Print each check to console with actual vs expected values
   - Raise a warning (do not abort) if any discrepancy > £0.01

5. **Generate the HTML** using the SGM house style:
   - Dark navy header (`#1a3a5c`) with orange accent (`#e8621a`)
   - CSS variables: `--primary:#1a3a5c`, `--accent:#e8621a`, `--light:#f4f7fb`, `--border:#d0d9e8`, `--text:#1e2533`, `--muted:#6b7a99`, `--green:#1a7a4a`
   - KPI cards row at top — warn card (orange border + value) for Unallocated Cost
   - Two-column card grid (`minmax(480px,1fr)`), weekly trend full-width (`grid-column: 1 / -1`) at bottom
   - Gold/silver/bronze rank badges for positions 1–3
   - Inline bar charts for units table (green bars) and volume table (accent/orange bars)
   - Peak week cell highlighted `style="background:#fffbe6"` with the highest unit that week marked
   - Footnote on weekly trend table flagging partial weeks with `*` / `†`
   - Footer: period, source folder path, classification "Board Confidential", generated date
   - Data only — no action signal pills, no warning flags on customer names

6. **Save** as `SGM_Parts_Analysis_[DateRange].html` in `S:\SGMWindows\Customer Care\Reports\`
   - Date range format: `Feb-Mar2026` for full months, `DD-DDMonYYYY` for filtered ranges

7. **Open in Chrome:**
   - Run `python -m http.server 8765` in background from `S:\SGMWindows\Customer Care\Reports\`
   - Open with `start "" "http://localhost:8765/<filename>.html"`
   - Use Chrome browser tools to take a screenshot confirming it loaded correctly

---

## Why direct file reading (not DataBaseSearch.xlsm)

Power Query in `DataBaseSearch.xlsm` reads formula cached values from source xlsx files. When files are modified by openpyxl (e.g. to fix sheet headers), openpyxl strips `<v>` cached value tags from formula cells. The affected columns are:

- **Col L (Cost per item):** array formula `XLOOKUP(B2:B29,'.'!A,'.'!B,0)` — row 2 master cell cache is cleared; rows 3+ retain individual cached values
- **Col M (Total Cost):** formula `=D*L` — all caches cleared

Power Query evaluates `D*L` at refresh time but returns null when the XLOOKUP master cache (L2) is empty, even if individual L3+ cells are present. Reading directly with `openpyxl data_only=True` bypasses this and the three-tier fallback recovers all costs correctly.

**Affected files (2026):** `Parts List - 23.02.26.xlsx` and `Parts List - 10.03.26.xlsx` through `Parts List - 24.03.26.xlsx` (excluding 25.03.26). These were modified to fix a `Customercare Office` sheet header issue.

---

## Style reference

Canonical style: `C:\Users\Darren\Downloads\Parts_Top10_Report_Mar2026.html`
Match its CSS variables, card layout, table structure, and rank badge colours exactly.

---

## Reference output

Benchmark figures are stored in memory (`memory/parts_report_benchmarks.md`) and updated after each run. Always compare a new run against the most recent benchmark — flag any unit that drops >20% without a known reason.

**Latest benchmark — 25 Mar 2026, period 06 Feb – 24 Mar 2026:**
- Total Cost: £6,912.41 | Unallocated: £1,556.18 (22.5%)
- Unit 6: £2,678.18 (38.7%) | Unit 5: £2,106.42 (30.5%) | Unit 4: £1,123.86 (16.3%)
- Unit 3: £438.56 (6.3%) | Customer Care: £317.12 (4.6%) | Unit 1: £248.26 (3.6%)

**KPI cards shown:** Total Cost and Unallocated Cost only (Transactions, Customers Served, Unique Parts removed).
