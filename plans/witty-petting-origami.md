# Photo Temporal Clustering + Vision Search

## Context
Build a reusable Claude Code skill (`/photo-cluster`) that scans a folder of photos, groups them into temporal clusters (15-minute gap threshold), presents the clusters, then uses Claude's vision capabilities to find photos matching a user-defined semantic query — and copies entire matching clusters into a named subfolder.

---

## Architecture

Two components:

### 1. `C:\Users\Darren\photo_cluster.py` — Core Python script
Handles all filesystem and metadata work. Two modes:

**`--list <dir>`** — Scan directory, extract timestamps, return JSON clusters  
**`--copy <dir> <term> <cluster_id> [cluster_id ...]`** — Create `<dir>/<term>/` and copy all files from named clusters into it

**Timestamp extraction logic (Pillow only — no extra deps):**
```python
from PIL import Image
from PIL.ExifTags import TAGS

def get_timestamp(path):
    try:
        img = Image.open(path)
        exif = img._getexif()
        if exif:
            for tag_id, val in exif.items():
                if TAGS.get(tag_id) == "DateTimeOriginal":
                    return datetime.strptime(val, "%Y:%m:%d %H:%M:%S")
    except Exception:
        pass
    return datetime.fromtimestamp(os.path.getmtime(path))  # fallback
```

**Clustering algorithm (gap-based):**
- Sort files by timestamp ascending
- Start Cluster 1 with first file
- For each next file: if gap from previous file > 15 minutes → new cluster; else append to current cluster
- Cluster IDs use date-time stamp format: `2024-03-15_1430` (start of cluster)
- Output: `[{id, start, end, count, files: [...]}, ...]`

**Supported extensions:** `.jpg .jpeg .png .heic .tiff .tif .raw .cr2 .nef .arw .dng`

---

### 2. `C:\Users\Darren\.claude\skills\photo-cluster\skill.md` — Skill orchestration
Claude follows this multi-phase workflow when user invokes `/photo-cluster`:

**Phase 1 — Scan**
- Accept argument as target directory (default: current dir)
- Run `python C:\Users\Darren\photo_cluster.py --list <dir>` via Bash
- Parse JSON output; display cluster table:
  | Cluster | Start | End | Photos |
  |---------|-------|-----|--------|

**Phase 2 — Query**
- Ask user for search term (e.g., "industrial windows", "street photography")

**Phase 3 — Vision Inspect**
- For each cluster, use `Read` tool on each image file to visually inspect it
- Flag any cluster containing at least one image matching the search term
- Log which specific files matched

**Phase 4 — Copy**
- For each matching cluster, run `python photo_cluster.py --copy <dir> "<term>" <cluster_id>`
- Script creates `<dir>/<term>/` (inside the source directory) and copies all files from that cluster (not just matches)

**Phase 5 — Report**
- Output summary table in active voice:
  | File | Source Cluster | Destination | Action |
  |------|---------------|-------------|--------|

---

## Files to Create

| File | Purpose |
|------|---------|
| `C:\Users\Darren\photo_cluster.py` | Python CLI — scan, cluster, copy |
| `C:\Users\Darren\.claude\skills\photo-cluster\skill.md` | Skill definition and workflow |

---

## Dependencies
- **Pillow 12.1.1** — already installed, handles EXIF + most formats
- No additional installs needed for JPG/PNG/TIFF
- HEIC support: Pillow alone cannot decode HEIC pixel data but can still read EXIF headers on some systems. If HEIC EXIF fails, the script falls back to file mtime gracefully.

---

## Verification
1. Point `/photo-cluster` at a test folder with a few JPGs
2. Confirm cluster JSON looks correct (run `python photo_cluster.py --list <dir>` directly)
3. Run a vision query and confirm Claude reads the images and identifies matches
4. Confirm copied files land in `<dir>/<term>/` and source files are untouched (copy, not move)
5. Confirm summary table lists every moved file
