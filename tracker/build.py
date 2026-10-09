"""Builds Filter_Log.xlsx and Filter_Order_Log.xlsx from data.json."""
import json, sys
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

data = json.load(open("data.json"))
job = data["job"]
HEAD = PatternFill("solid", fgColor="1F4E78")
BLANK = PatternFill("solid", fgColor="FFF2CC")   # yellow = still needs info
BAD = PatternFill("solid", fgColor="F8CBAD")


def header(ws, row, names):
    for c, n in enumerate(names, 1):
        cell = ws.cell(row=row, column=c, value=n)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEAD
        cell.alignment = Alignment(wrap_text=True, vertical="center")


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def job_block(ws):
    labels = [("Job number", "job_number"), ("Project name", "project"), ("Phase", "phase"),
              ("PO number", "po_number"), ("Airco PM", "airco_pm"), ("Ordered by", "ordered_by")]
    for r, (lab, key) in enumerate(labels, 1):
        ws.cell(row=r, column=1, value=lab).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=job[key] or None)
        if not job[key]:
            c.fill = BLANK
    return len(labels) + 2


# ---------- 1. Filter log ----------
wb = Workbook(); ws = wb.active; ws.title = "Filter Log"
top = job_block(ws)
header(ws, top, ["Unit", "Filter size (W x L x D)", "Qty per change", "Status", "Filter set installed", "Notes"])
for i, (unit, size, qty, status) in enumerate(data["units"], top + 1):
    for c, v in enumerate([unit, size, qty, status, None, None], 1):
        ws.cell(row=i, column=c, value=v)
    if status == "Not replaced":
        ws.cell(row=i, column=4).fill = BAD
    ws.cell(row=i, column=5).fill = BLANK
widths(ws, [14, 22, 14, 16, 20, 40])
ws.freeze_panes = ws.cell(row=top + 1, column=1)
last_unit_row = top + len(data["units"])
wb.save("Filter_Log.xlsx")

# ---------- 2. Order log ----------
wb = Workbook(); ws = wb.active; ws.title = "Order Log"
top = job_block(ws)
header(ws, top, ["Order date", "Filter set (1/2/3/Attic)", "Filter order #", "Order total price",
                 "Size", "Qty ordered", "Units it went to", "Notes"])
r = top + 1
for o in data["orders"]:
    first = True
    for size, qty in o["lines"]:
        vals = [o["date"], o["set"] or None, o["order_number"] or None, o["price"] if first else None,
                size, qty, None, o["note"] if first else None]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            if v is None and c in (2, 3, 7) or (first and c == 4 and v is None):
                cell.fill = BLANK
        ws.cell(row=r, column=4).number_format = '"$"#,##0.00'
        first = False; r += 1
order_last = r - 1
widths(ws, [12, 14, 14, 14, 12, 12, 24, 44])
ws.freeze_panes = ws.cell(row=top + 1, column=1)

# Reconciliation sheet: formulas read the other sheet, so edits flow through.
rs = wb.create_sheet("Needed vs Ordered")
sizes = sorted({s for _, s, _, _ in data["units"]} | {s for o in data["orders"] for s, _ in o["lines"]})
header(rs, 1, ["Size", "Ordered (all orders)", "Needed - units replaced", "Left over (+) / Short (-)",
               "Still needed - units not yet replaced"])
for i, s in enumerate(sizes, 2):
    rs.cell(row=i, column=1, value=s)
    rs.cell(row=i, column=2, value=f"=SUMIF('Order Log'!$E${top+1}:$E${order_last},A{i},'Order Log'!$F${top+1}:$F${order_last})")
    # replaced-unit totals typed from the filter log at build time
    need = sum(q for _, sz, q, st in data["units"] if sz == s and st == "Replaced")
    rem = sum(q for _, sz, q, st in data["units"] if sz == s and st != "Replaced")
    rs.cell(row=i, column=3, value=need)
    rs.cell(row=i, column=4, value=f"=B{i}-C{i}")
    rs.cell(row=i, column=5, value=rem)
widths(rs, [12, 20, 24, 24, 34])
wb.save("Filter_Order_Log.xlsx")

# console summary
ordered = defaultdict(int); need = defaultdict(int)
for o in data["orders"]:
    for s, q in o["lines"]: ordered[s] += q
for _, s, q, st in data["units"]:
    if st == "Replaced": need[s] += q
for s in sizes:
    print(f"{s:10} ordered {ordered[s]:3}  needed {need[s]:3}  diff {ordered[s]-need[s]:+d}")
