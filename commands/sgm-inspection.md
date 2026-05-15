# SGM Factory Inspection App

Launch, update, or manage the SGM Factory Inspection PWA.

---

## App location

**Primary (live — use this one):**
```
C:\Users\Darren\OneDrive - SGM Windows\Desktop\Folders\HS\sgm-inspection-pwa\
```

**Backup / dev copy:**
```
C:\Users\Darren\sgm-inspection-pwa\
```

Always edit the primary location. After any change, sync to the backup with:
```bash
cp -r "C:/Users/Darren/OneDrive - SGM Windows/Desktop/Folders/HS/sgm-inspection-pwa/." "C:/Users/Darren/sgm-inspection-pwa/"
```

---

## Files

| File | Purpose |
|---|---|
| `index.html` | App shell — loads all scripts |
| `js/jspdf.min.js` | jsPDF 2.5.1 — bundled locally, no CDN needed |
| `js/config.js` | Audit sections, questions, assignees, units |
| `js/app.js` | Full app logic, routing, scoring, UI |
| `js/db.js` | IndexedDB persistence |
| `js/pdf-gen.js` | PDF report generation (`generatePDF` + `_buildPDF`) |
| `css/styles.css` | Mobile-first SGM green theme |
| `sw.js` | Service worker — cache version `sgm-inspect-v3` |
| `manifest.json` | PWA manifest |

---

## Launch

Kill any old server first, then start from the primary folder:

```bash
taskkill /F /IM python.exe 2>nul
cd "C:/Users/Darren/OneDrive - SGM Windows/Desktop/Folders/HS/sgm-inspection-pwa" && python -m http.server 8765
```

Then open: `http://localhost:8765`

After any code change, do **Ctrl+Shift+R** in Chrome to hard-reload and clear the service worker cache.

---

## PDF generation

- jsPDF is served from `js/jspdf.min.js` (local file, no internet needed)
- `generatePDF(insp)` in `pdf-gen.js` is async — calls `_buildPDF(insp)` which is sync
- `_buildPDF` returns `{ doc, filename }`; `generatePDF` handles the save/share
- If PDF fails, an `alert()` shows the exact error — also check F12 console for `[PDF]` log lines
- Image format is auto-detected from dataURL prefix (JPEG/PNG/WebP all supported)
- After changing `pdf-gen.js`, bump the sw.js cache version (`sgm-inspect-v3` → `v4` etc.) so Chrome picks up the new file

---

## Modifying the audit

**Questions/sections:** edit `js/config.js` — the `sections` array. Each section has a `questions` array of plain strings. No other files need changing.

**Assignees:** edit `AUDIT_CONFIG.defaultAssignees` in `js/config.js`.

**Units:** edit `AUDIT_CONFIG.units` in `js/config.js`.

**Branding colours:** CSS custom properties at the top of `css/styles.css` — primary is `--green: #2e7d32`.

**PDF layout:** `js/pdf-gen.js` — the `_buildPDF` function. All nested drawing functions (`pdfHeader`, `drawMetaTable`, `drawScoreBanner`, `drawActionsSummary`, `drawFullAudit`, `drawPhotoAppendix`, `addPageNumbers`) are defined inside `_buildPDF`.

---

## PWA install on mobile (requires HTTPS)

```bash
ngrok http 8765
```

Share the HTTPS URL with the phone, open in mobile browser, tap "Add to Home Screen".

---

## Tech notes

- Data: **IndexedDB** — persists across sessions, survives app updates
- Photos: stored as **base64 dataURLs** inside each inspection record
- Scoring: Pass = 1pt, Action/Fail = 0pt, N/A excluded from denominator, unanswered excluded
- Section "Complete" = all questions have any non-null status (including N/A)
- Service worker cache name must be bumped (`v3` → `v4`) whenever files change, to force clients to update
