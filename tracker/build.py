"""Builds a job's spreadsheets from data.json (run inside the job folder).
Run with --template to make the blank PM templates in tracker/templates/."""
import json, re, sys
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

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
              ("Airco PM", "airco_pm"), ("Ordered by", "ordered_by")]
    if "phase" not in job:  # job without phases
        labels.remove(("Phase", "phase"))
    # optional fields, only shown for jobs that have them
    for lab, key in [("Airco APM", "apm"), ("PM history", "pm_history"), ("GC", "gc"),
                     ("Job type", "job_type"), ("Warranty expires", "warranty")]:
        if key in job:
            labels.append((lab, key))
    for r, (lab, key) in enumerate(labels, 1):
        ws.cell(row=r, column=1, value=lab).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=job[key] or None)
        if not job[key]:
            c.fill = BLANK
    if job.get("note"):
        ws.cell(row=len(labels) + 1, column=1, value=job["note"]).font = Font(italic=True)
        return len(labels) + 3
    return len(labels) + 2


# ---------- Blank templates for PMs: python3 build.py --template ----------
# Same columns and colors as the job files, but empty, with dropdowns, so a PM can fill one in by hand.
def build_templates():
    import os
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.formatting.rule import FormulaRule
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
    os.makedirs(out, exist_ok=True)
    RED = PatternFill("solid", fgColor="F8CBAD")

    def dropdown(ws, cells, choices):
        # error off: the list is a shortcut, PMs can still type something else
        dv = DataValidation(type="list", formula1='"' + ",".join(choices) + '"', allow_blank=True, showErrorMessage=False)
        ws.add_data_validation(dv); dv.add(cells)

    def blank_job_block(ws, legend):
        labels = ["Job number", "Project name", "Phase", "Airco PM", "Airco APM", "Ordered by",
                  "GC", "Job type", "PM history", "Warranty expires"]
        for r, lab in enumerate(labels, 1):
            ws.cell(row=r, column=1, value=lab).font = Font(bold=True)
        ws.cell(row=len(labels) + 1, column=1, value=legend).font = Font(italic=True)
        return len(labels) + 3

    # 1. Filter log
    wb = Workbook(); ws = wb.active; ws.title = "Filter Log"
    top = blank_job_block(ws, "Yellow = still needs info. Red = unit not replaced. One row per filter size "
                              "(a unit that takes 2 sizes gets 2 rows). Write sizes W x L x D and say if nominal or exact.")
    header(ws, top, ["Unit", "Filter size (W x L x D)", "QTY", "Status", "Filter set installed", "MERV", "Installed size if different from spec", "Notes"])
    last = top + 100
    dropdown(ws, f"D{top+1}:D{last}", ["Not started", "Needs filters", "Ordered", "Replaced", "Not replaced", "Washable - not replaced"])
    dropdown(ws, f"E{top+1}:E{last}", ["1", "2", "3", "Attic"])
    dropdown(ws, f"F{top+1}:F{last}", ["8", "13"])
    ws.conditional_formatting.add(f"D{top+1}:D{last}", FormulaRule(formula=[f'$D{top+1}="Not replaced"'], fill=RED))
    widths(ws, [14, 22, 14, 16, 20, 8, 26, 40])
    ws.freeze_panes = ws.cell(row=top + 1, column=1)
    wb.save(os.path.join(out, "TEMPLATE_FilterLog.xlsx"))

    # 2. Order log + Needed vs Ordered
    wb = Workbook(); ws = wb.active; ws.title = "Order Log"
    top = blank_job_block(ws, "Yellow = still needs info. Each order (new PO) starts a new row; one row per filter size. "
                              "Put the order total only on the first row of each order. Line total fills in by itself.")
    header(ws, top, ["Order date", "Filter set (1/2/3/Attic)", "Vendor", "Vendor SO #", "Our PO #", "MERV",
                     "Order total price", "Size", "QTY", "Unit price", "Line total", "Units it went to", "Delivery", "Notes"])
    last = top + 60
    for r in range(top + 1, last + 1):
        ws.cell(row=r, column=1).number_format = "mm/dd/yyyy"
        for c in (7, 10, 11):
            ws.cell(row=r, column=c).number_format = '"$"#,##0.00'
        ws.cell(row=r, column=11, value=f'=IF(OR(I{r}="",J{r}=""),"",I{r}*J{r})')
    dropdown(ws, f"B{top+1}:B{last}", ["1", "2", "3", "Attic"])
    dropdown(ws, f"F{top+1}:F{last}", ["8", "13"])
    dropdown(ws, f"M{top+1}:M{last}", ["Jobsite delivery", "Will call", "Customer pickup"])
    widths(ws, [12, 12, 16, 12, 12, 7, 14, 11, 7, 10, 11, 24, 20, 50])
    ws.freeze_panes = ws.cell(row=top + 1, column=1)

    rs = wb.create_sheet("Needed vs Ordered")
    header(rs, 1, ["Size", "Ordered (all orders)", "Needed - units replaced", "Left over (+) / Short (-)",
                   "Still needed - units not yet replaced"])
    rs.cell(row=1, column=7, value="Type each size in column A. Columns C and E: count them from the Filter Log. "
                                   "Columns B and D fill in by themselves.").font = Font(italic=True)
    for i in range(2, 42):
        rs.cell(row=i, column=2, value=f"=IF(A{i}=\"\",\"\",SUMIF('Order Log'!$H${top+1}:$H${last},A{i},'Order Log'!$I${top+1}:$I${last}))")
        rs.cell(row=i, column=4, value=f'=IF(A{i}="","",B{i}-C{i})')
    widths(rs, [12, 20, 24, 24, 34, 4, 60])
    wb.save(os.path.join(out, "TEMPLATE_FiltersOrders.xlsx"))

    # 3. Open questions
    wb = Workbook(); ws = wb.active; ws.title = "Open Questions"
    ws.cell(row=1, column=1, value="Job ________ - ________________").font = Font(bold=True, size=13)
    ws.cell(row=2, column=1, value='=COUNTIF(A4:A103,"Open")&" open - ready to ship when this is 0 and every unit and filter is accounted for"')
    header(ws, 3, ["Status", "Question", "Who knows"])
    for r in range(4, 104):
        for c in (1, 2, 3):
            ws.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical="top")
    dropdown(ws, "A4:A103", ["Open", "Answered", "Later"])
    ws.conditional_formatting.add("A4:A103", FormulaRule(formula=['$A4="Open"'], fill=BLANK))
    widths(ws, [11, 100, 18])
    wb.save(os.path.join(out, "TEMPLATE_OpenQuestions.xlsx"))
    print("Templates saved in", out)


if "--template" in sys.argv:
    build_templates()
    sys.exit()

data = json.load(open("data.json"))
job = data["job"]
# file names: jobnumber_PROJECTNAME_what-it-is, no spaces (e.g. 2250473_AISDLANGFORD_FiltersOrders.xlsx)
prefix = f"{job['job_number']}_{job.get('short_name') or re.sub(r'[^A-Za-z0-9]', '', job['project']).upper()}_"

# ---------- 1. Filter log ----------
wb = Workbook(); ws = wb.active; ws.title = "Filter Log"
top = job_block(ws)
header(ws, top, ["Unit", "Filter size (W x L x D)", "QTY", "Status", "Filter set installed", "MERV", "Installed size if different from spec", "Notes"])
for i, (unit, size, qty, status, *note) in enumerate(data["units"], top + 1):
    done = status == "Replaced"
    for c, v in enumerate([unit, size, qty, status, 2 if done else None, 8 if done else None, None, note[0] if note else None], 1):
        ws.cell(row=i, column=c, value=v)
    # unknown size or qty: mark yellow
    if "?" in size:
        ws.cell(row=i, column=2).fill = BLANK
    if qty is None:
        ws.cell(row=i, column=3).fill = BLANK
    if status == "Not replaced":
        ws.cell(row=i, column=4).fill = BAD
    if not done:
        ws.cell(row=i, column=5).fill = BLANK
        ws.cell(row=i, column=6).fill = BLANK
    if done and size.endswith('x2'):
        ws.cell(row=i, column=7).fill = BLANK
# no unit list yet: leave blank yellow rows to fill in
for i in range(top + 1, top + 1 + (data.get("blank_rows", 0) if not data["units"] else 0)):
    for c in range(1, 9):
        ws.cell(row=i, column=c).fill = BLANK
widths(ws, [14, 22, 14, 16, 20, 8, 26, 40])
ws.freeze_panes = ws.cell(row=top + 1, column=1)
last_unit_row = top + len(data["units"])
wb.save(prefix + "FilterLog.xlsx")

# ---------- 2. Order log ----------
wb = Workbook(); ws = wb.active; ws.title = "Order Log"
top = job_block(ws)
header(ws, top, ["Order date", "Filter set (1/2/3/Attic)", "Vendor", "Vendor SO #", "Our PO #", "MERV",
                 "Order total price", "Size", "QTY", "Unit price", "Line total", "Units it went to", "Delivery", "Notes"])
r = top + 1
for o in data["orders"]:
    first = True
    for size, qty, price in o["lines"]:
        vals = [o["date"], o["set"] or None, o["vendor"], o["vendor_so"], o["po_number"] or None, o["merv"],
                o["price"] if first else None, size, qty, price, f"=I{r}*J{r}", (o.get("units") if first else None),
                o["delivery"], o["note"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            if v is None and (c in (2, 5, 12) or (first and c in (1, 3, 4, 6, 7, 13))):
                cell.fill = BLANK
        ws.cell(row=r, column=7).number_format = '"$"#,##0.00'
        ws.cell(row=r, column=10).number_format = '"$"#,##0.00'
        ws.cell(row=r, column=11).number_format = '"$"#,##0.00'
        first = False; r += 1
order_last = r - 1
widths(ws, [12, 12, 16, 12, 12, 7, 14, 11, 7, 10, 11, 24, 20, 50])
ws.freeze_panes = ws.cell(row=top + 1, column=1)

# Reconciliation sheet: formulas read the other sheet, so edits flow through.
rs = wb.create_sheet("Needed vs Ordered")
sizes = sorted({u[1] for u in data["units"] if not u[3].startswith("Washable")} | {l[0] for o in data["orders"] for l in o["lines"]})
header(rs, 1, ["Size", "Ordered (all orders)", "Needed - units replaced", "Left over (+) / Short (-)",
               "Still needed - units not yet replaced"])
for i, s in enumerate(sizes, 2):
    rs.cell(row=i, column=1, value=s)
    rs.cell(row=i, column=2, value=f"=SUMIF('Order Log'!$H${top+1}:$H${order_last},A{i},'Order Log'!$I${top+1}:$I${order_last})")
    # replaced-unit totals typed from the filter log at build time
    need = sum(q or 0 for _, sz, q, st, *_ in data["units"] if sz == s and st == "Replaced")
    rem = sum(q or 0 for _, sz, q, st, *_ in data["units"] if sz == s and st != "Replaced")
    rs.cell(row=i, column=3, value=need)
    rs.cell(row=i, column=4, value=f"=B{i}-C{i}")
    rs.cell(row=i, column=5, value=rem)
widths(rs, [12, 20, 24, 24, 34])
wb.save(prefix + "FiltersOrders.xlsx")

# console summary
ordered = defaultdict(int); need = defaultdict(int)
for o in data["orders"]:
    for s, q, _ in o["lines"]: ordered[s] += q
for _, s, q, st, *_ in data["units"]:
    if st == "Replaced": need[s] += q or 0
for s in sizes:
    print(f"{s:10} ordered {ordered[s]:3}  needed {need[s]:3}  diff {ordered[s]-need[s]:+d}")

# ---------- 3. Open questions (working notes, not part of the two shipped files) ----------
wb = Workbook(); ws = wb.active; ws.title = "Open Questions"
ws.cell(row=1, column=1, value=f"Job {job['job_number']} - {job['project']}").font = Font(bold=True, size=13)
header(ws, 3, ["Status", "Question", "Who knows"])
for i, (st, q, who) in enumerate(data["questions"], 4):
    for c, v in enumerate([st, q, who], 1):
        cell = ws.cell(row=i, column=c, value=v)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    if st == "Open":
        ws.cell(row=i, column=1).fill = BLANK
n_open = sum(1 for q in data["questions"] if q[0] == "Open")
ws.cell(row=2, column=1, value=f"{n_open} open - ready to ship when this is 0 and every unit and filter is accounted for")
widths(ws, [11, 100, 18])
wb.save(prefix + "OpenQuestions.xlsx")
