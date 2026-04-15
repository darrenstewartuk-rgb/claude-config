---
name: Verify figures before publishing
description: Always manually verify all report figures before presenting them, and explicitly confirm to the user that checks have been done
type: feedback
---

Before presenting any report output (Parts Analysis or similar), manually verify all figures and explicitly state that checks have been done.

**Why:** User asked for this after a report was published without an explicit confirmation that the numbers had been cross-checked.

**How to apply:**
- After the report script runs, manually re-derive: sum units == Total Cost, each % share = unit_tc / total_tc, share sum ~= 100%
- State "Figures checked — all verified" (or list each check) before presenting results
- Do not just rely on the script's own validation output — show the cross-check explicitly in the response
