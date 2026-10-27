"""Builds a job's spreadsheets from data.json (run inside the job folder):
  jobnumber_PROJECT_FilterLog.xlsx   - the file for the PM: tab 1 Filter Log, tab 2 Filter Orders
  jobnumber_PROJECT_OpenQuestions.xlsx - just for us, not filed
Run with --template to make the blank PM templates in tracker/templates/."""
import json, os, re, sys
from collections import defaultdict
from itertools import zip_longest
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

NAVY, ORANGE = "2B3C8E", "FC7C36"   # Airco logo colors
HEAD = PatternFill("solid", fgColor=NAVY)
BLANK = PatternFill("solid", fgColor="FFF2CC")   # yellow = still needs info
BAD = PatternFill("solid", fgColor="FFC7CE")     # red = unit not replaced
LABEL = PatternFill("solid", fgColor="DCE0F0")   # light navy
BAND = PatternFill("solid", fgColor="F2F2F2")    # light grey every other row
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

SET_NAMES = {"1": "1st (media roll)", "2": "2nd", "3": "3rd", "Attic": "Attic stock"}
SETS = ["1st (media roll)", "2nd", "3rd", "Attic stock", "RFI", "Change order"]
STATUSES = ["Not started", "Needs filters", "Ordered", "Replaced", "Not replaced", "Washable - not replaced"]
DELIVERY = ["Will call", "Delivered to site", "At Airco warehouse"]
LOG_COLS = ["Unit", "Filter size (W x L x D)", "QTY", "Status", "Filter set or reason", "MERV", "Notes"]
LOG_WIDTHS = [16, 20, 8, 22, 20, 8, 60]
ORDER_COLS = ["Order date", "Ordered by", "Filter set or reason", "Vendor", "Vendor SO #", "Our PO #", "MERV",
              "Order total price", "Size", "QTY", "Unit price", "Line total", "Delivery", "Notes"]
ORDER_WIDTHS = [14, 16, 18, 18, 13, 13, 8, 14, 14, 8, 11, 12, 20, 45]
JOB_LEFT = [("Job number", "job_number"), ("Project name", "project"), ("Phase", "phase"), ("Job type", "job_type")]
JOB_RIGHT = [("Airco PM", "airco_pm"), ("Airco APM", "apm"), ("GC", "gc")]   # Ordered by is a column on Filter Orders
MUST_HAVE = {"job_number", "project", "airco_pm"}   # yellow when empty; the rest are optional


def set_name(s):
    return SET_NAMES.get(str(s), s) if s not in (None, "") else None


def header(ws, row, names):
    for c, n in enumerate(names, 1):
        cell = ws.cell(row=row, column=c, value=n)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEAD
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=Side(style="medium", color=ORANGE))
    ws.row_dimensions[row].height = 34


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def dropdown(ws, cells, choices):
    # error off: the list is a shortcut, people can still type something else
    dv = DataValidation(type="list", formula1='"' + ",".join(choices) + '"', allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(dv); dv.add(cells)


def title(ws, text, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws.cell(row=1, column=1, value=text).font = Font(bold=True, size=16, color=NAVY)
    for c in range(1, ncols + 1):   # orange line under the title
        ws.cell(row=1, column=c).border = Border(bottom=Side(style="medium", color=ORANGE))
    ws.row_dimensions[1].height = 28


def field(ws, row, col, label, last_col, value=None, fill=None):
    lab = ws.cell(row=row, column=col, value=label)
    lab.font = Font(bold=True); lab.fill = LABEL; lab.border = BOX
    lab.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=row, start_column=col + 1, end_row=row, end_column=last_col)
    for c in range(col + 1, last_col + 1):
        ws.cell(row=row, column=c).border = BOX
    v = ws.cell(row=row, column=col + 1, value=value)
    v.alignment = Alignment(vertical="center", wrap_text=True)
    if fill:
        v.fill = fill
    # merged cells do not grow by themselves, so give long values (project name) two lines
    width = sum(ws.column_dimensions[get_column_letter(c)].width or 10 for c in range(col + 1, last_col + 1))
    lines = 2 if value and len(str(value)) * 1.1 > width else 1
    ws.row_dimensions[row].height = max(ws.row_dimensions[row].height or 0, 20 if lines == 1 else 32)
    return v.coordinate


def job_box(ws, ncols, job=None, copy_from=None):
    """Job info in two side by side columns (rows 2-5) so the top stays short.
    job: fill in the values. copy_from: (sheet, cells) to show the other tab's values by formula."""
    mid = ncols // 2 + 1
    where = {}
    for r, (a, b) in enumerate(zip_longest(JOB_LEFT, JOB_RIGHT), 2):
        for col, last, item in [(1, mid - 1, a), (mid, ncols, b)]:
            if not item:
                continue
            lab, key = item
            val = (job or {}).get(key) or None
            fill = BLANK if job is not None and key in MUST_HAVE and not val else None
            where[lab] = field(ws, r, col, lab, last, val, fill)
    for c in range(1, ncols + 1):   # orange line between the title and the job box
        cell = ws.cell(row=2, column=c)
        b = cell.border
        cell.border = Border(left=b.left, right=b.right, bottom=b.bottom, top=Side(style="medium", color=ORANGE))
    if copy_from:
        sheet, cells = copy_from
        for lab, ref in where.items():
            src = f"'{sheet}'!{cells[lab]}"
            ws[ref] = f'=IF({src}="","",{src})'
    return where


def note_row(ws, ncols, text):
    """Optional job note under the job box (row 6). Returns the header row."""
    text = re.sub(r"\s*Yellow = still needs info\.?", "", text or "").strip()
    if not text:
        ws.row_dimensions[6].height = 8
        return 7
    ws.merge_cells(start_row=6, start_column=1, end_row=6, end_column=ncols)
    c = ws.cell(row=6, column=1, value=text)
    c.font = Font(italic=True, color="404040"); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[6].height = 15 * (len(text) // 150 + 1) + 4
    return 8


def table(ws, names, top, nrows, center_cols, wrap_cols):
    """Header at row top, nrows bordered rows under it. Header stays on screen and prints on every page."""
    header(ws, top, names)
    first, last = top + 1, top + nrows
    for r in range(first, last + 1):
        # no fixed height, so rows grow by themselves when a long note wraps
        for c in range(1, len(names) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BOX
            cell.alignment = Alignment(vertical="center", wrap_text=c in wrap_cols,
                                       horizontal="center" if c in center_cols else None)
    ws.freeze_panes = ws.cell(row=first, column=1)
    ws.print_title_rows = f"{top}:{top}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return first, last


def stripes(ws, first, last, ncols):
    # grey every other row, but never over a yellow/red cell
    for r in range(first, last + 1):
        if r % 2 == 0:
            for c in range(1, ncols + 1):
                cell = ws.cell(row=r, column=c)
                if cell.fill.fill_type is None:
                    cell.fill = BAND


def log_sheet(ws, job=None):
    widths(ws, LOG_WIDTHS)
    title(ws, "FILTER LOG", len(LOG_COLS))
    box = job_box(ws, len(LOG_COLS), job)
    top = note_row(ws, len(LOG_COLS), job and job.get("note"))
    return box, top


def orders_sheet(ws, log_box, job=None):
    widths(ws, ORDER_WIDTHS)
    title(ws, "FILTER ORDERS", len(ORDER_COLS))
    job_box(ws, len(ORDER_COLS), job, ("Filter Log", log_box))
    ws.row_dimensions[6].height = 8
    return 7


def log_lists(ws, first, last):
    dropdown(ws, f"D{first}:D{last}", STATUSES)
    dropdown(ws, f"E{first}:E{last}", SETS)
    dropdown(ws, f"F{first}:F{last}", ["8", "13"])


def order_lists(ws, first, last):
    dropdown(ws, f"C{first}:C{last}", SETS)
    dropdown(ws, f"G{first}:G{last}", ["8", "13"])
    dropdown(ws, f"M{first}:M{last}", DELIVERY)


def order_formats(ws, first, last):
    for r in range(first, last + 1):
        ws.cell(row=r, column=1).number_format = "mm/dd/yyyy"
        for c in (8, 11, 12):
            ws.cell(row=r, column=c).number_format = '"$"#,##0.00'


def questions_sheet(ws, job=None):
    title(ws, "OPEN QUESTIONS", 3)
    for r, (lab, key) in enumerate([("Job number", "job_number"), ("Project name", "project"), ("Airco PM", "airco_pm")], 2):
        c = ws.cell(row=r, column=1, value=lab); c.font = Font(bold=True); c.fill = LABEL; c.border = BOX
        v = ws.cell(row=r, column=2, value=(job or {}).get(key) or None); v.border = BOX
        ws.row_dimensions[r].height = 20


# ---------- Blank templates for PMs: python3 build.py --template ----------
# Same layout as the job files, but empty, with dropdowns, so a PM can fill one in by hand.
def build_templates():
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
    os.makedirs(out, exist_ok=True)

    def bands(ws, rng):
        # light grey every other row; added last so the red/yellow rules win
        ws.conditional_formatting.add(rng, FormulaRule(formula=["MOD(ROW(),2)=0"], fill=BAND))

    # 1. Filter log (main tab)
    wb = Workbook(); ws = wb.active; ws.title = "Filter Log"
    log_box, top = log_sheet(ws)
    first, last = table(ws, LOG_COLS, top, 100, center_cols={2, 3, 4, 5, 6}, wrap_cols={7})
    log_lists(ws, first, last)
    ws.conditional_formatting.add(f"A{first}:G{last}", FormulaRule(formula=[f'$D{first}="Not replaced"'], fill=BAD, stopIfTrue=True))
    bands(ws, f"A{first}:G{last}")
    widths(ws, LOG_WIDTHS)

    # 2. Filter orders: second tab of the same file. Only a purchase log, unit/size work lives in the Filter Log.
    ws = wb.create_sheet("Filter Orders")
    top = orders_sheet(ws, log_box)
    first, last = table(ws, ORDER_COLS, top, 60, center_cols={1, 3, 5, 6, 7, 9, 10}, wrap_cols={2, 4, 13, 14})
    order_formats(ws, first, last)
    for r in range(first, last + 1):
        ws.cell(row=r, column=12, value=f'=IF(OR(J{r}="",K{r}=""),"",J{r}*K{r})')
    order_lists(ws, first, last)
    bands(ws, f"A{first}:N{last}")
    widths(ws, ORDER_WIDTHS)
    wb.save(os.path.join(out, "TEMPLATE_FilterLog.xlsx"))

    # 3. Open questions (just for us, not filed with the PM)
    wb = Workbook(); ws = wb.active; ws.title = "Open Questions"
    questions_sheet(ws)
    ws.cell(row=5, column=2, value='=COUNTIF(A8:A107,"Open")&" open - ready to send when this is 0 and every unit and filter is accounted for"').font = Font(bold=True)
    first, last = table(ws, ["Status", "Question", "Who knows"], 7, 100, center_cols={1}, wrap_cols={2, 3})
    dropdown(ws, f"A{first}:A{last}", ["Open", "Answered", "Later"])
    ws.conditional_formatting.add(f"A{first}:A{last}", FormulaRule(formula=[f'$A{first}="Open"'], fill=BLANK, stopIfTrue=True))
    bands(ws, f"A{first}:C{last}")
    widths(ws, [14, 90, 22])
    wb.save(os.path.join(out, "TEMPLATE_OpenQuestions.xlsx"))
    print("Templates saved in", out)


if "--template" in sys.argv:
    build_templates()
    sys.exit()

data = json.load(open("data.json"))
job = data["job"]
# file names: jobnumber_PROJECTNAME_what-it-is, no spaces (e.g. 2250473_AISDLANGFORD_FilterLog.xlsx)
prefix = f"{job['job_number']}_{job.get('short_name') or re.sub(r'[^A-Za-z0-9]', '', job['project']).upper()}_"
EXTRA = 20   # empty rows under the data, with the same dropdowns, for adding by hand

# ---------- 1. Filter log (tab 1 of the PM file) ----------
wb = Workbook(); ws = wb.active; ws.title = "Filter Log"
log_box, top = log_sheet(ws, job)
n = len(data["units"]) or data.get("blank_rows", 0)
first, last = table(ws, LOG_COLS, top, n + EXTRA, center_cols={2, 3, 4, 5, 6}, wrap_cols={7})
for i, (unit, size, qty, status, *note) in enumerate(data["units"], first):
    done = status == "Replaced"
    for c, v in enumerate([unit, size, qty, status, "2nd" if done else None, 8 if done else None, note[0] if note else None], 1):
        ws.cell(row=i, column=c, value=v)
    # unknown size or qty: mark yellow
    if "?" in size:
        ws.cell(row=i, column=2).fill = BLANK
    if qty is None:
        ws.cell(row=i, column=3).fill = BLANK
    if status == "Not replaced":
        for c in range(1, len(LOG_COLS) + 1):
            ws.cell(row=i, column=c).fill = BAD
    elif not done and not status.startswith("Washable"):
        ws.cell(row=i, column=5).fill = BLANK
        ws.cell(row=i, column=6).fill = BLANK
# no unit list yet: leave blank yellow rows to fill in
if not data["units"]:
    for i in range(first, first + n):
        for c in range(1, len(LOG_COLS) + 1):
            ws.cell(row=i, column=c).fill = BLANK
stripes(ws, first, last, len(LOG_COLS))
log_lists(ws, first, last)
widths(ws, LOG_WIDTHS)

# ---------- 2. Filter orders (tab 2 of the PM file) ----------
ws = wb.create_sheet("Filter Orders")
top = orders_sheet(ws, log_box, job)
n = sum(len(o["lines"]) for o in data["orders"])
first, last = table(ws, ORDER_COLS, top, n + EXTRA, center_cols={1, 3, 5, 6, 7, 9, 10}, wrap_cols={2, 4, 13, 14})
order_formats(ws, first, last)
r = first
for o in data["orders"]:
    for k, (size, qty, price) in enumerate(o["lines"]):
        first_line = k == 0
        # order-wide info (total, notes) only on the first line, so a big order is easy to read
        note = o["note"] if first_line else None
        if first_line and o.get("units"):
            note = (note + " " if note else "") + "Units: " + o["units"]
        # who placed this order; falls back to the job's "ordered_by" only for old data without it
        vals = [o["date"], o.get("ordered_by") or job.get("ordered_by") or None, set_name(o["set"]), o["vendor"], o["vendor_so"], o["po_number"] or None, o["merv"],
                o["price"] if first_line else None, size, qty, price, f"=J{r}*K{r}", o["delivery"], note]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            if v is None and (c in (2, 3, 6) or (first_line and c in (1, 4, 5, 7, 8, 13))):
                cell.fill = BLANK
        r += 1
for rr in range(r, last + 1):
    ws.cell(row=rr, column=12, value=f'=IF(OR(J{rr}="",K{rr}=""),"",J{rr}*K{rr})')
stripes(ws, first, last, len(ORDER_COLS))
order_lists(ws, first, last)
widths(ws, ORDER_WIDTHS)
wb.save(prefix + "FilterLog.xlsx")

# console summary: ordered vs needed per size (for us; no longer a sheet in the PM file)
ordered = defaultdict(int); need = defaultdict(int)
for o in data["orders"]:
    for s, q, _ in o["lines"]: ordered[s] += q
for _, s, q, st, *_ in data["units"]:
    if st == "Replaced": need[s] += q or 0
for s in sorted(set(ordered) | set(need)):
    print(f"{s:10} ordered {ordered[s]:3}  needed {need[s]:3}  diff {ordered[s]-need[s]:+d}")

# ---------- 3. Open questions (just for us, not filed with the PM) ----------
wb = Workbook(); ws = wb.active; ws.title = "Open Questions"
questions_sheet(ws, job)
n_open = sum(1 for q in data["questions"] if q[0] == "Open")
ws.cell(row=5, column=2, value=f"{n_open} open - ready to send when this is 0 and every unit and filter is accounted for").font = Font(bold=True)
first, last = table(ws, ["Status", "Question", "Who knows"], 7, len(data["questions"]) + EXTRA, center_cols={1}, wrap_cols={2, 3})
for i, (st, q, who) in enumerate(data["questions"], first):
    for c, v in enumerate([st, q, who], 1):
        ws.cell(row=i, column=c, value=v)
    if st == "Open":
        ws.cell(row=i, column=1).fill = BLANK
stripes(ws, first, last, 3)
dropdown(ws, f"A{first}:A{last}", ["Open", "Answered", "Later"])
widths(ws, [14, 90, 22])
wb.save(prefix + "OpenQuestions.xlsx")
