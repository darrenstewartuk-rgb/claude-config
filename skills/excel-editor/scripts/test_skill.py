#!/usr/bin/env python3
"""Smoke test for the excel-editor skill."""
import subprocess, json, tempfile, openpyxl, os, sys

SCRIPT = r'C:\Users\Darren\.claude\skills\excel-editor\scripts\excel_editor.py'

d = tempfile.mkdtemp()
folder = d.replace('\\', '/')

# Create a two-sheet workbook
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Sales'
wb.create_sheet('Returns')
ws['A1'] = 'Revenue'
ws['B1'] = 'Units'
ws['A2'] = 1000
ws['B2'] = 50
wb.save(os.path.join(d, 'report.xlsx'))

def run(op):
    r = subprocess.run([sys.executable, SCRIPT, json.dumps(op)], capture_output=True, text=True)
    out = r.stdout.strip()
    err = r.stderr.strip()
    parsed = json.loads(out) if out else {}
    print(f"  {op['type']:30s} -> {parsed}")
    if 'error' in parsed:
        print(f"    ERROR: {parsed['error']}")
    return parsed

print("=== Excel Editor Skill Smoke Test ===\n")

print("[Discovery]")
run({'type': 'list_workbooks', 'folder': folder})
run({'type': 'list_sheets', 'folder': folder, 'workbook': 'report.xlsx'})
run({'type': 'read_range', 'folder': folder, 'workbook': 'report.xlsx', 'sheet': 'Sales', 'range': 'A1:B2'})

print("\n[Data]")
run({'type': 'set_cell_value', 'folder': folder, 'workbook': 'report.xlsx', 'sheet': 'Sales', 'cell': 'C1', 'value': 'Profit'})
run({'type': 'set_range_values', 'folder': folder, 'workbook': 'report.xlsx', 'sheet': 'Sales', 'start_row': 2, 'start_col': 'C', 'values': [[500], [300]]})
run({'type': 'find_replace', 'folder': folder, 'workbook': 'report.xlsx', 'sheet': '__all__', 'find': 'Revenue', 'replace': 'Income', 'match_case': False})

print("\n[Formatting]")
run({'type': 'set_range_format', 'folder': folder, 'workbook': 'report.xlsx', 'sheet': 'Sales', 'range': 'A1:C1', 'style': {'font': {'bold': True, 'color': 'white', 'size': 12}, 'fill': {'color': 'darkblue'}, 'alignment': {'horizontal': 'center'}}})
run({'type': 'set_column_width', 'folder': folder, 'workbook': 'report.xlsx', 'sheet': 'Sales', 'columns': ['A', 'B', 'C'], 'width': 18})
run({'type': 'set_row_height', 'folder': folder, 'workbook': 'report.xlsx', 'sheet': 'Sales', 'rows': [1], 'height': 25})
run({'type': 'freeze_panes', 'folder': folder, 'workbook': 'report.xlsx', 'sheet': 'Sales', 'cell': 'A2'})
run({'type': 'apply_filter', 'folder': folder, 'workbook': 'report.xlsx', 'sheet': 'Sales', 'range': 'A1:C1'})

print("\n[Structure]")
run({'type': 'add_sheet', 'folder': folder, 'workbook': 'report.xlsx', 'name': 'Summary', 'position': 0})
run({'type': 'set_tab_color', 'folder': folder, 'workbook': 'report.xlsx', 'sheet': 'Summary', 'color': 'green'})
run({'type': 'rename_sheet', 'folder': folder, 'workbook': 'report.xlsx', 'sheet': 'Returns', 'new_name': 'Refunds'})
run({'type': 'list_sheets', 'folder': folder, 'workbook': 'report.xlsx'})

print(f"\nOutput workbook: {os.path.join(d, 'report.xlsx')}")
print("=== All tests completed ===")
