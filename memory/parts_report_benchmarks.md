---
name: parts_report_benchmarks
description: Historical benchmark figures from each SGM Parts Analysis report run — use to sanity-check new runs and spot anomalies
type: project
---

Use these figures when a new report is run to compare totals. Flag any run where a unit's cost drops >20% vs the prior period without a known reason (e.g. fewer working days, seasonal).

## Run history

### Run: 25 Mar 2026 — Period: 06 Feb – 24 Mar 2026 (34 working days)
- **Total Cost: £6,912.41**
- Unallocated Cost: £1,556.18 (22.5%)
- Unit 6: £2,678.18 (38.7%)
- Unit 5: £2,106.42 (30.5%)
- Unit 4: £1,123.86 (16.3%)
- Unit 3: £438.56 (6.3%)
- Customer Care: £317.12 (4.6%)
- Unit 1: £248.26 (3.6%)

**Why:** How to apply: on a new run, compare unit % shares and total cost. If a unit drops significantly check for missing source files or formula cache issues in affected xlsx files.

---

### Earlier reference: 18 Mar 2026 — Period: 06 Feb – 18 Mar 2026 (from DataBaseSearch Summary, pre-openpyxl modification)
- **Total Cost: £6,234.35**
- Note: computed from Power Query Summary sheet before source files were modified. Direct-read method gives slightly different figures due to formula cache recovery. Use the 25 Mar run as the canonical baseline going forward.
