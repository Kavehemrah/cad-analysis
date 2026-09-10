import os
from datetime import datetime

import win32com.client


HEADERS = [
    "Drainage", "Status", "Confidence", "Sleeper", "Pipe Axis (deg)",
    "Sleeper Axis (deg)", "Angle Diff (deg)", "Pipe X", "Pipe Y",
    "Sleeper X", "Sleeper Y", "Local X", "Local Y", "Pipe Diameter (mm)",
    "Sleeper Width (mm)", "Current Clearance (mm)", "Required Move (mm)",
    "Final Clearance (mm)", "Move Dir X", "Move Dir Y", "Old X", "Old Y",
    "New X", "New Y", "Note"
]


def write_xlsx(path, results, clearance_mm):
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    wb = app.Workbooks.Add()

    summary = wb.Worksheets(1)
    summary.Name = "Summary"
    ws = wb.Worksheets.Add(After=summary)
    ws.Name = "Results"
    cfg = wb.Worksheets.Add(After=ws)
    cfg.Name = "Parameters"

    total = len(results)
    blocked = sum(r.status == "BLOCKED" for r in results)
    clear = sum(r.status == "CLEAR" for r in results)
    ambiguous = sum(r.status == "AMBIGUOUS" for r in results)
    unmatched = sum(r.status == "UNMATCHED" for r in results)
    errors = sum(r.status == "ERROR" for r in results)
    moved = sum(r.required_move_mm > 0 and r.status == "BLOCKED" for r in results)

    summary.Range("A1:D1").Merge()
    summary.Range("A1").Value = "Drainage / Sleeper Analysis"
    summary.Range("A1").Font.Bold = True
    summary.Range("A1").Font.Size = 16

    summary.Range("A3:B10").Value = [
        ["Run time", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["Total drainages", total],
        ["Blocked", blocked],
        ["Clear", clear],
        ["Ambiguous", ambiguous],
        ["Unmatched", unmatched],
        ["Errors", errors],
        ["Move required", moved],
    ]

    summary.Range("D3:E5").Value = [
        ["Pipe diameter (mm)", "Detected from block geometry"],
        ["Sleeper width (mm)", "Detected from block geometry"],
        ["Required clearance (mm)", clearance_mm],
    ]

    rows = [HEADERS]
    for r in results:
        rows.append([
            r.drainage, r.status, r.confidence, r.sleeper,
            r.pipe_axis_deg, r.sleeper_axis_deg, r.angle_diff_deg,
            r.pipe_x, r.pipe_y, r.sleeper_x, r.sleeper_y,
            r.local_x, r.local_y, r.pipe_diameter_mm, r.sleeper_width_mm,
            r.current_clearance_mm, r.required_move_mm, r.final_clearance_mm,
            r.move_dir_x, r.move_dir_y, r.old_x, r.old_y, r.new_x, r.new_y,
            r.note,
        ])
    ws.Range(ws.Cells(1, 1), ws.Cells(len(rows), len(HEADERS))).Value = rows
    ws.Range("A1:Y1").Font.Bold = True
    ws.Range("A1:Y1").AutoFilter()
    ws.Application.ActiveWindow.SplitRow = 1
    ws.Application.ActiveWindow.FreezePanes = True

    cfg.Range("A1:B8").Value = [
        ["Parameter", "Value"],
        ["Drainage block", "Drainage Pipe E"],
        ["Sleeper layer", "AG_Print"],
        ["Sleeper blocks", "AG_t, AG_tttt"],
        ["Expected pipe diameter (mm)", 63],
        ["Required clearance (mm)", clearance_mm],
        ["Sleeper angle tolerance (deg)", 8],
        ["Match margin (mm)", 80],
    ]
    cfg.Range("A1:B1").Font.Bold = True

    for sheet in (summary, ws, cfg):
        sheet.Columns.AutoFit()
        sheet.Rows.RowHeight = 18
    ws.Columns(25).ColumnWidth = 42
    summary.Columns(1).ColumnWidth = 28
    summary.Columns(2).ColumnWidth = 26

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.SaveAs(os.path.abspath(path), FileFormat=51)
    wb.Close(SaveChanges=True)
    app.Quit()
