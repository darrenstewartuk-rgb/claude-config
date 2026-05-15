# Plan: Run SGM Parts Analysis Report

## Context
Generate today's SGM parts analysis HTML report by running the existing `parts_report_build.py` script against the completed parts list sheets on the S: drive. The last run was 22 Apr 2026 (benchmark: £10,118.98 total). This run will produce an updated report covering all files currently in the completed sheets folder.

## Steps

1. **Run the script**
   ```
   python C:\Users\Darren\parts_report_build.py
   ```
   - Script reads all `Parts List - *.xlsx` from `S:\SGMWindows\Customer Care\2026 Parts Lists\Completed Sheets Here`
   - Outputs HTML to `S:\SGMWindows\Customer Care\Reports\SGM_Parts_Analysis_<DateRange>.html`
   - Prints validation summary (unit totals vs weekly totals must match)

2. **Review console output**
   - Note file count, total rows, total cost, unallocated %
   - Check validation lines — all should say `OK`
   - Compare unit totals against benchmark (flag any unit down >20% without known reason)

3. **Serve and open in Chrome**
   ```
   python -m http.server 8765 --directory "S:\SGMWindows\Customer Care\Reports"
   ```
   Then open `http://localhost:8765/<output_filename>.html` in Chrome.

## Critical files
- Script: `C:\Users\Darren\parts_report_build.py`
- Source data: `S:\SGMWindows\Customer Care\2026 Parts Lists\Completed Sheets Here\Parts List - *.xlsx`
- Output folder: `S:\SGMWindows\Customer Care\Reports\`

## Benchmark comparison (22 Apr 2026)
| Unit | Last Total |
|---|---|
| Unit 6 | £3,559.84 (35.2%) |
| Unit 5 | £3,245.85 (32.1%) |
| Unit 4 | £1,677.00 (16.6%) |
| Unit 3 | £757.29 (7.5%) |
| Unit 1 | £496.43 (4.9%) |
| Customer Care | £382.57 (3.8%) |
| **Total** | **£10,118.98** |
| Unallocated | £2,418.02 (23.9%) |
