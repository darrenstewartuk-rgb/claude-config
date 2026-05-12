# /photo-cluster

Search photos visually for a user-defined term using 3-way AI classification, then group and copy confirmed matches and nearby no-reference photos into a named subfolder.

## Trigger
User invokes `/photo-cluster [directory]`

## Workflow

### Phase 1 — Scan
- If a directory argument was provided, use it. Otherwise ask: "Which folder should I scan for photos?"
- Run: `python C:\Users\Darren\photo_cluster.py --scan "<directory>"`
- Parse the JSON output.
- If `error` key is present, report it and stop.
- Report the total number of photos found and the date range (earliest to latest timestamp).

### Phase 2 — Query
- Ask the user: "What are you looking for? Enter a job number and/or customer reference (e.g. '226063' and 'R Holwell')."
- Store as `<TERM>` (job number) and `<REF>` (customer ref, optional).

### Phase 3 — Vision Inspect (3-way classification)
- Use the `Read` tool to open every file returned by `--scan`.
- For each image, classify it as exactly one of:
  - **MATCH** — clearly shows `<TERM>` or `<REF>`
  - **WRONG** — shows a *different* reference number or customer name
  - **NOREF** — no reference information visible (product, site, installation photo)
- After inspection, report:
  - Total photos inspected
  - MATCH count (with filename and one-line reason for each)
  - WRONG count (filenames — these will be excluded)
  - NOREF count

### Phase 4 — Group and Copy
- If no MATCH photos: report "No photos matched." and stop.
- The copy set is:
  - All **MATCH** photos
  - All **NOREF** photos within ±15 minutes of any MATCH *in the same folder*
  - **WRONG** photos are never included
- Run a single command passing all MATCH file paths as anchors:
  `python C:\Users\Darren\photo_cluster.py --copy-around "<directory>" "<TERM>" "<match1>" "<match2>" ...`
- Note: the script groups by time from the anchors; WRONG-classified photos that fall within the window will be copied by the script. To prevent this, only pass MATCH files as anchors and verify the copy list excludes known WRONG files.

### Phase 5 — Report
Parse the JSON copy results and output a summary table in active voice:

| File | Timestamp | Classification | Seconds from match | Action |
|------|-----------|---------------|-------------------|--------|

End with: "Copied N files into `<directory>\<TERM>\` — M confirmed match(es), K nearby photos, W wrong-reference photo(s) excluded."

## Notes
- **WRONG photos are never included** — if a photo shows a different job number or customer name it is excluded completely, not just flagged.
- NOREF photos (windows, installations, site shots with no label) are included only if within ±15 min of a confirmed MATCH in the same folder.
- Copies are non-destructive. Source files are never moved or deleted.
- The output subfolder is created inside the source directory as `<directory>\<TERM>\`.
- If EXIF data is missing, file modification time is used as the timestamp fallback.
- The standalone desktop app (`C:\Users\Darren\dist\PhotoCluster.exe`) performs the same workflow automatically with a GUI.
