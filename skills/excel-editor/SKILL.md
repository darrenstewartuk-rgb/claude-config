---
name: excel-editor
description: Make formatting and data changes to Excel workbooks (.xlsx/.xlsm) in a named folder. Handles any number of workbooks and tabs. Use when the user asks to edit, format, update, or modify Excel files.
disable-model-invocation: false
allowed-tools: Read, Write, Edit, Glob, Bash
---

# Excel Editor Skill

You are an expert Excel automation agent. You can read and modify Excel workbooks in any folder using the helper script at:

`C:/Users/Darren/.claude/skills/excel-editor/scripts/excel_editor.py`

## ARGUMENTS

$ARGUMENTS

## Workflow

### Step 1 — Understand the request
Parse $ARGUMENTS to extract:
- **Folder path** — where the Excel files live (ask if not given)
- **Workbooks** — specific file(s) or all in the folder
- **Sheets/tabs** — specific tab(s) or all
- **Changes requested** — data edits, formatting, structure changes

If the folder path is missing, ask the user for it before proceeding.

### Step 2 — Discover what's in the folder
Run discovery to understand the workbooks and sheets:

```bash
python "C:/Users/Darren/.claude/skills/excel-editor/scripts/excel_editor.py" '{"type":"list_workbooks","folder":"<FOLDER>"}'
```

Then list sheets for relevant workbooks:
```bash
python "C:/Users/Darren/.claude/skills/excel-editor/scripts/excel_editor.py" '{"type":"list_sheets","folder":"<FOLDER>","workbook":"<FILE.xlsx>"}'
```

To preview data in a range:
```bash
python "C:/Users/Darren/.claude/skills/excel-editor/scripts/excel_editor.py" '{"type":"read_range","folder":"<FOLDER>","workbook":"<FILE.xlsx>","sheet":"Sheet1","range":"A1:E10"}'
```

### Step 3 — Plan the operations
Translate the user's request into one or more operation JSON calls. Choose the correct `type` from the list below.

### Step 4 — Execute operations
Run each operation via Bash, one at a time or in logical sequence:
```bash
python "C:/Users/Darren/.claude/skills/excel-editor/scripts/excel_editor.py" '<OPERATION_JSON>'
```

### Step 5 — Confirm and report
Tell the user what was changed, in which files and sheets.

---

## Available Operation Types

All operations share these common fields:
- `folder` — path to the folder (required for all file operations)
- `workbook` — specific filename (e.g. `"report.xlsx"`), or omit to target ALL workbooks in the folder
- `sheet` — sheet name or 0-based index; omit for active sheet; use `"__all__"` to target all sheets

### Data Operations

**Set a single cell value:**
```json
{"type":"set_cell_value","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","cell":"B2","value":42}
```

**Set a block of values (2D array, row-major):**
```json
{"type":"set_range_values","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","start_row":2,"start_col":"A","values":[["Name","Score"],["Alice",95],["Bob",87]]}
```

**Find and replace across all sheets:**
```json
{"type":"find_replace","folder":"C:/Data","workbook":"report.xlsx","sheet":"__all__","find":"OldText","replace":"NewText","match_case":false}
```

### Formatting Operations

**Format specific cells:**
```json
{"type":"set_cell_format","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","cells":["A1","B1"],"style":{"font":{"bold":true,"size":14,"color":"darkblue"},"fill":{"color":"lightblue"},"alignment":{"horizontal":"center"}}}
```

**Format a range:**
```json
{"type":"set_range_format","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","range":"A1:Z1","style":{"font":{"bold":true,"color":"white"},"fill":{"color":"darkblue"},"alignment":{"horizontal":"center","wrap_text":true}}}
```

**Number format:**
```json
{"type":"set_cell_format","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","cells":["C2"],"style":{"number_format":"$#,##0.00"}}
```

Common number formats:
- `"General"` — auto
- `"0"` — integer
- `"0.00"` — 2 decimal places
- `"$#,##0.00"` — currency
- `"0%"` — percentage
- `"DD/MM/YYYY"` — date
- `"@"` — text

**Column width:**
```json
{"type":"set_column_width","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","columns":["A","B","C"],"width":20}
```

**Row height:**
```json
{"type":"set_row_height","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","rows":[1,2],"height":30}
```

**Merge / unmerge cells:**
```json
{"type":"merge_cells","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","range":"A1:D1"}
{"type":"unmerge_cells","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","range":"A1:D1"}
```

**Borders** (in style dict):
```json
"border":{"top":"thin","bottom":"medium","left":"thin","right":"thin"}
```
Border styles: `"thin"`, `"medium"`, `"thick"`, `"dashed"`, `"dotted"`, `"double"`

**Font options** (in style dict):
```json
"font":{"bold":true,"italic":false,"underline":true,"size":12,"color":"red","name":"Arial","strikethrough":false}
```

### Structural Operations

**Add a sheet:**
```json
{"type":"add_sheet","folder":"C:/Data","workbook":"report.xlsx","name":"Summary","position":0}
```

**Rename a sheet:**
```json
{"type":"rename_sheet","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","new_name":"Sales Data"}
```

**Delete a sheet:**
```json
{"type":"delete_sheet","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet3"}
```

**Reorder sheets:**
```json
{"type":"reorder_sheets","folder":"C:/Data","workbook":"report.xlsx","order":["Summary","Jan","Feb","Mar"]}
```

**Set tab color:**
```json
{"type":"set_tab_color","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","color":"red"}
```

**Freeze panes:**
```json
{"type":"freeze_panes","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","cell":"B2"}
```

**Auto filter:**
```json
{"type":"apply_filter","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","range":"A1:F1"}
```

**Insert rows/columns:**
```json
{"type":"insert_rows","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","row":2,"count":3}
{"type":"insert_columns","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","col":"B","count":2}
```

**Delete rows/columns:**
```json
{"type":"delete_rows","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","row":5,"count":1}
{"type":"delete_columns","folder":"C:/Data","workbook":"report.xlsx","sheet":"Sheet1","col":"C","count":1}
```

---

## Color Names Supported
`red`, `green`, `blue`, `yellow`, `orange`, `purple`, `black`, `white`, `gray/grey`, `lightblue`, `lightgreen`, `lightgray/lightgrey`, `cyan`, `magenta`, `pink`, `darkblue`, `darkgreen`, `darkred`, `gold`, `teal`, `navy`, `lime`, `brown`, `silver`

Or use hex: `"#FF5733"` or `"FF5733"`

---

## Multi-workbook Targeting
Omit `"workbook"` to apply an operation to **all** Excel files in the folder:
```json
{"type":"set_range_format","folder":"C:/Reports","sheet":"Summary","range":"A1:Z1","style":{"font":{"bold":true}}}
```

---

## Prerequisites Check
If the first operation fails with an openpyxl error, install it:
```bash
pip install openpyxl
```

---

## Error Handling
- If a sheet name doesn't exist, report it clearly and ask the user to confirm the correct name (use `list_sheets` to show available names).
- If a workbook file doesn't exist, list available workbooks with `list_workbooks`.
- Never guess file paths — always confirm with the user if uncertain.
- Back up important files before bulk operations by noting to the user that changes are saved in-place.
