"""Build the G2 PSG Tracker workbook (.xlsx) for upload/conversion to Google Sheets.

Formulas are written in Google Sheets dialect (MAP, SCAN, MAKEARRAY, FILTER, SORT,
SPARKLINE...). They are evaluated by Google Sheets after conversion, not by Excel.

Usage:  python3 build_tracker.py [output.xlsx]
"""
import os
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.dimensions import ColumnDimension

try:  # real roster data stays local (git-ignored); the example set ships with the repo
    import seed_data_private as S
except ImportError:
    import seed_data_example as S

# ----------------------------------------------------------------------------
# Palette (validated reference palette; status colors reserved for state)
# ----------------------------------------------------------------------------
NAVY = "1F3A5F"
SLATE = "5B6B7F"
INK = "0B0B0B"
INK2 = "52514E"
SURF = "FCFCFB"
PAGE = "F4F5F7"
AUTO = "EEF1F5"      # computed-column tint
LINE = "D5DAE1"
SERIES = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "6250D6", "E34948"]
GOOD, WARN, CRIT = "0CA30C", "FAB219", "D03B3B"
GOOD_BG, WARN_BG, CRIT_BG, GREY_BG = "DFF3DF", "FFF1CC", "F9DADA", "E4E6EA"

FONT = "Arial"
thin = Side(style="thin", color=LINE)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

ROSTER_KEYS = [r[0] for r in S.ROSTERS]
ROSTER_ARR = "{" + ",".join('"%s"' % k for k in ROSTER_KEYS) + "}"
NR = len(S.ROSTERS)
# Personnel layout: A-I typed, then one BF # column per roster, grade/order/status, one eligibility column
# per roster, then details / last / next (all computed)
PC_BF = 10
PC_GRADE = PC_BF + NR
PC_ORD, PC_STAT, PC_ELIG = PC_GRADE + 1, PC_GRADE + 2, PC_GRADE + 3
PC_DET = PC_ELIG + NR
PC_LAST, PC_NEXT = PC_DET + 1, PC_DET + 2
L_ORD, L_STAT, L_DET, L_LAST, L_NEXT = (CL(c) for c in (PC_ORD, PC_STAT, PC_DET, PC_LAST, PC_NEXT))
L_E1, L_E2 = CL(PC_ELIG), CL(PC_ELIG + NR - 1)
# Next Up: one 6-column block per roster
NU_W = 6
NU_LAST = CL(NU_W * NR - 1)

# ----------------------------------------------------------------------------
# Range tokens (bounded ranges keep recalculation fast)
# ----------------------------------------------------------------------------
SLN, AVN, PEN = 1000, 1000, 400
SL = "'Duty Slots'!"
AV = "Availability!"
PE = "Personnel!"
ST = "Settings!"
T = {
    "SL_A": f"{SL}$A$4:$A${SLN}", "SL_B": f"{SL}$B$4:$B${SLN}", "SL_D": f"{SL}$D$4:$D${SLN}",
    "SL_E": f"{SL}$E$4:$E${SLN}", "SL_F": f"{SL}$F$4:$F${SLN}", "SL_G": f"{SL}$G$4:$G${SLN}",
    "SL_H": f"{SL}$H$4:$H${SLN}", "SL_I": f"{SL}$I$4:$I${SLN}", "SL_J": f"{SL}$J$4:$J${SLN}",
    "SL_L": f"{SL}$L$4:$L${SLN}", "SL_M": f"{SL}$M$4:$M${SLN}", "SL_P": f"{SL}$P$4:$P${SLN}",
    "AV_A": f"{AV}$A$4:$A${AVN}", "AV_B": f"{AV}$B$4:$B${AVN}", "AV_C": f"{AV}$C$4:$C${AVN}",
    "AV_D": f"{AV}$D$4:$D${AVN}", "AV_E": f"{AV}$E$4:$E${AVN}", "AV_G": f"{AV}$G$4:$G${AVN}",
    "AV_H": f"{AV}$H$4:$H${AVN}", "AV_I": f"{AV}$I$4:$I${AVN}", "AV_J": f"{AV}$J$4:$J${AVN}",
    "AV_M": f"{AV}$M$4:$M${AVN}",
    "PE_A": f"{PE}$A$4:$A${PEN}", "PE_B": f"{PE}$B$4:$B${PEN}", "PE_BC": f"{PE}$B$4:$C${PEN}",
    "PE_C": f"{PE}$C$4:$C${PEN}", "PE_D": f"{PE}$D$4:$D${PEN}", "PE_E": f"{PE}$E$4:$E${PEN}",
    "PE_H": f"{PE}$H$4:$H${PEN}", "PE_ORD": f"{PE}${L_ORD}$4:${L_ORD}${PEN}",
    "PE_ELIG": f"{PE}${L_E1}$4:${L_E2}${PEN}", "PE_DET": f"{PE}${L_DET}$4:${L_DET}${PEN}",
    "START": f"{ST}$B$4", "REST": f"{ST}$B$5", "TIE": f"{ST}$B$6", "UNIT": f"{ST}$B$7",
    "ASOF": f"{ST}$B$9", "LOOK": f"{ST}$B$10", "CYCLE": f"{ST}$B$11", "PSG": f"{ST}$B$12",
    "HOL": f"{ST}$Y$4:$Y$60", "RANKTBL": f"{ST}$E$4:$H$40", "DUTYTBL": f"{ST}$P$4:$Q$30",
    "ABSTBL": f"{ST}$T$4:$V$40", "ROSTER_ARR": ROSTER_ARR,
    "MS": "'Dash Data'!$B$2", "ME": "'Dash Data'!$D$2",
}


def F(s, **extra):
    """Expand <TOKEN> placeholders in a formula template."""
    toks = dict(T)
    toks.update(extra)
    # longest tokens first so <SL_A> never clobbers <SL_AB>
    for k in sorted(toks, key=len, reverse=True):
        s = s.replace("<%s>" % k, str(toks[k]))
    return s


# ----------------------------------------------------------------------------
# Style helpers
# ----------------------------------------------------------------------------
def font(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(hex_):
    return PatternFill("solid", fgColor=hex_)


def banner(ws, text, sub, last_col, row=1):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    c = ws.cell(row=row, column=1, value=text)
    c.font = font(14, True, "FFFFFF")
    c.fill = fill(NAVY)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 30
    ws.merge_cells(start_row=row + 1, start_column=1, end_row=row + 1, end_column=last_col)
    c = ws.cell(row=row + 1, column=1, value=sub)
    c.font = font(9, False, INK2, italic=True)
    c.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[row + 1].height = 30


def header_row(ws, row, headers, auto_from=None, widths=None):
    """Input headers are navy; computed (auto) headers are slate."""
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        is_auto = auto_from is not None and h in auto_from
        c.font = font(10, True, "FFFFFF")
        c.fill = fill(SLATE if is_auto else NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[row].height = 32
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[CL(i)].width = w


TINTS = []


def tint(ws, col_letters, r1, r2):
    """Computed-column tint, applied as a lowest-priority conditional format (keeps the file small)."""
    for col in col_letters:
        TINTS.append((ws, f"{col}{r1}:{col}{max(r2, 1000)}"))


def date_fmt(ws, col_letters, r1, r2, fmt="dd mmm yy"):
    r2 = min(r2, r1 + 30)   # Apps Script Setup formats whole columns
    for col in col_letters:
        for r in range(r1, r2 + 1):
            ws[f"{col}{r}"].number_format = fmt


def dv_list(ws, ref, sqref, allow_blank=True, error=True):
    dv = DataValidation(type="list", formula1=ref, allow_blank=allow_blank)
    dv.showErrorMessage = error
    dv.errorStyle = "warning"
    ws.add_data_validation(dv)
    dv.add(sqref)


def base_font(ws, r1, r2, c1, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            if cell.font is None or cell.font.name != FONT:
                cell.font = font()


# ----------------------------------------------------------------------------
wb = Workbook()
wb.remove(wb.active)
TABS = ["Dashboard", "Duty Slots", "Next Up"] + [r[1] for r in S.ROSTERS] + [
    "Personnel", "Availability", "Appointments", "Equipment", "Taskings", "Settings",
    "Dash Data"]
WS = {name: wb.create_sheet(name) for name in TABS}
TAB_COLOR = {"Dashboard": NAVY, "Duty Slots": "EDA100", "Next Up": "1BAF7A", "Personnel": SLATE,
             "Availability": SLATE, "Appointments": "6250D6", "Equipment": "6250D6",
             "Taskings": "6250D6", "Settings": "7F7F7F", "Dash Data": "BFBFBF"}
for r in S.ROSTERS:
    TAB_COLOR[r[1]] = "2A78D6"
for k, v in TAB_COLOR.items():
    WS[k].sheet_properties.tabColor = v

# ============================================================================
# SETTINGS
# ============================================================================
ws = WS["Settings"]
banner(ws, "SETTINGS · lists, rules and lookups",
       "Edit values in white cells. Every tab reads from here: rank bands per DA 6, duty types, "
       "absence codes (DA PAM 220-45 para 7c: A / D / U only), holidays and the tracker registry.", 35)
header_row(ws, 3, ["Setting", "Value", "Notes"])
settings_rows = [
    ("Roster FROM date", date(2026, 10, 1), "First posting date counted on every DA 6 (DA PAM 220-45 para 5a). Tracker history starts 1 Oct 26."),
    ("Rest window (days)", 1, "A Soldier with any other G2 duty within +/- this many days is coded D and is not suggested."),
    ("Tie-break (equal numbers)", "Roster order", "DA PAM 220-45 is silent on ties. 'Roster order' = senior/alphabetical as listed; 'Junior first' = lowest grade first."),
    ("Unit label on BDE roster", "G2", "Slots labeled with this section/unit are G2's to fill (used by the importer)."),
    ("As-of date override", None, "Leave blank to use TODAY(). Set a date to brief 'as of' a specific day."),
    ("As-of date (in use)", "=IF(B8=\"\",TODAY(),B8)", "Computed: drives every status, KPI and chart."),
    ("Look-ahead window (days)", 14, "Dashboard 'upcoming' and 'open slots' window."),
    ("Equipment inventory cycle (days)", 30, "Next inventory due = last inventory + this many days."),
    ("PSG", S.PSG_LABEL, "Shown on the dashboard banner."),
    ("Shared view file ID", None, "Filled automatically by the Apps Script 'Publish shared view'."),
    ("Exemption request form URL", None, "Filled automatically by the Apps Script 'Create exemption request form'."),
    ("Close-out (archive) log", None, "Filled automatically when you close/restart the DA 6s."),
    ("One-slider file ID", None, "Filled automatically by the Apps Script 'Build / refresh one-slider brief'."),
    ("Brief marking (optional)", None, "Banner text for the one-slider (top and footer), per your unit SOP. Blank = none."),
]
for i, (k, v, n) in enumerate(settings_rows, start=4):
    ws.cell(row=i, column=1, value=k).font = font(10, True)
    c = ws.cell(row=i, column=2, value=v)
    c.font = font(10, False, INK)
    c.border = BOX
    ws.cell(row=i, column=3, value=n).font = font(9, False, INK2, italic=True)
    ws.cell(row=i, column=3).alignment = Alignment(wrap_text=True, vertical="top")
ws["B4"].number_format = "dd mmm yyyy"
ws["B8"].number_format = "dd mmm yyyy"
ws["B9"].number_format = "dd mmm yyyy"
ws["B9"].fill = fill(AUTO)
dv_list(ws, '"Roster order,Junior first"', "B6")
for col, w in zip("ABC", (30, 22, 60)):
    ws.column_dimensions[col].width = w

# rank table E:H
for i, h in enumerate(["Rank", "Grade", "Order", "Category"]):
    c = ws.cell(row=3, column=5 + i, value=h)
for i, row in enumerate(S.RANKS, start=4):
    for j, v in enumerate(row):
        ws.cell(row=i, column=5 + j, value=v)
# rosters J:N
for i, h in enumerate(["DA 6 Roster", "Tab", "Min rank order", "Max rank order", "Nature of duty"]):
    ws.cell(row=3, column=10 + i, value=h)
for i, row in enumerate(S.ROSTERS, start=4):
    for j, v in enumerate(row):
        ws.cell(row=i, column=10 + j, value=v)
# duty types P:R
for i, h in enumerate(["Duty (as on BDE roster)", "DA 6 Roster", "Notes"]):
    ws.cell(row=3, column=16 + i, value=h)
for i, row in enumerate(S.DUTY_TYPES, start=4):
    for j, v in enumerate(row):
        ws.cell(row=i, column=16 + j, value=v)
# absence types T:W
for i, h in enumerate(["Availability type", "DA 6 code", "Blocks duty?", "Description"]):
    ws.cell(row=3, column=20 + i, value=h)
for i, row in enumerate(S.ABSENCE_TYPES, start=4):
    for j, v in enumerate(row):
        ws.cell(row=i, column=20 + j, value=v)
# holidays Y:Z
for i, h in enumerate(["Holiday / DONSA", "Name"]):
    ws.cell(row=3, column=25 + i, value=h)
for i, (d, n) in enumerate(S.HOLIDAYS, start=4):
    ws.cell(row=i, column=25, value=d).number_format = "ddd dd mmm yy"
    ws.cell(row=i, column=26, value=n)
# codes legend AB:AC
for i, h in enumerate(["DA 6 entry", "Meaning (DA PAM 220-45 paras 6, 7, 10, 11)"]):
    ws.cell(row=3, column=28 + i, value=h)
legend = [
    ("0", "Performed the duty on this date (green). Numbering restarts after it."),
    ("number", "Details made from this roster since the Soldier was last detailed."),
    ("A", "Not available: leave, pass, SD, TDY, illness in LOD, any authorized reason. Numbering is interrupted (paused)."),
    ("A + number", "First column of a new DA 6 for a Soldier still in A status: last number carried over (para 7c(1))."),
    ("D + number", "Eligible but not selectable because of a previous detail or other duty. Numbering continues."),
    ("U + number", "Not available through own misconduct (AWOL, arrest, confinement, sick not LOD). Numbering continues."),
    ("red column", "Weekend or holiday posting (para 11b)."),
]
for i, (a, b) in enumerate(legend, start=4):
    ws.cell(row=i, column=28, value=a)
    ws.cell(row=i, column=29, value=b)
# registry AE:AI
for i, h in enumerate(["Tracker table", "Tab", "Purpose", "Rows"]):
    ws.cell(row=3, column=31 + i, value=h)
for i, (t, tab, p) in enumerate(S.TRACKERS, start=4):
    ws.cell(row=i, column=31, value=t)
    ws.cell(row=i, column=32, value=tab)
    ws.cell(row=i, column=33, value=p)
for i in range(4, 31):
    ws.cell(row=i, column=34, value=f'=IF(AF{i}="","",IFERROR(COUNTA(INDIRECT("\'"&AF{i}&"\'!A4:A2000")),0))')
for col in range(5, 36):
    c = ws.cell(row=3, column=col)
    if c.value:
        c.font = font(10, True, "FFFFFF")
        c.fill = fill(NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
widths = {"D": 3, "E": 8, "F": 7, "G": 7, "H": 16, "I": 3, "J": 13, "K": 17, "L": 9, "M": 9, "N": 44,
          "O": 3, "P": 20, "Q": 13, "R": 34, "S": 3, "T": 24, "U": 9, "V": 9, "W": 46, "X": 3,
          "Y": 16, "Z": 30, "AA": 3, "AB": 12, "AC": 70, "AD": 3, "AE": 16, "AF": 16, "AG": 46, "AH": 8}
for k, w in widths.items():
    ws.column_dimensions[k].width = w
ws.freeze_panes = "A4"

# ============================================================================
# PERSONNEL
# ============================================================================
ws = WS["Personnel"]
P_HEAD = (["Section", "Name (LAST, FIRST)", "Rank", "Status", "DA6 Exclude (Y)", "Exclusion reason",
           "Phone", "Email", "Notes"] + [f"BF # {k}" for k in ROSTER_KEYS]
          + ["Grade", "Rank order", "Status today"] + ROSTER_KEYS + ["Details (total)", "Last detail", "Next detail"])
P_AUTO = set(P_HEAD[PC_GRADE - 1:])
banner(ws, "PERSONNEL · G2 master roster",
       "Single source of truth for every DA 6. Navy headers = you type; grey headers = automatic. "
       "Rank bands come from Settings. 'BF #' = brought-forward number from a previous DA 6 (leave 0 if unknown). "
       "Set Status to Departed or Exclude = Y (with a reason, DA PAM 220-45 para 10c) to drop a Soldier from all DA 6s.", PC_NEXT)
header_row(ws, 3, P_HEAD, P_AUTO,
           [11, 30, 7, 11, 9, 22, 14, 24, 24] + [8] * NR + [7, 7, 26] + [8] * NR + [9, 11, 11])
for i, (sec, name, rank, phone) in enumerate(S.PERSONNEL, start=4):
    ws.cell(row=i, column=1, value=sec)
    ws.cell(row=i, column=2, value=name)
    ws.cell(row=i, column=3, value=rank or None)
    ws.cell(row=i, column=4, value="Vacant" if name == "VACANT" else "Assigned")
    if phone:
        ws.cell(row=i, column=7, value=phone)
    if name == S.PSG_NAME:
        ws.cell(row=i, column=9, value="PSG")
r1, r2 = 4, PEN
ws[f"{CL(PC_GRADE)}4"] = F('=MAP(C4:C<N>,LAMBDA(r,IF(r="","",IFERROR(VLOOKUP(r,<RANKTBL>,2,FALSE),"?"))))', N=r2)
ws[f"{L_ORD}4"] = F('=MAP(C4:C<N>,LAMBDA(r,IF(r="","",IFERROR(VLOOKUP(r,<RANKTBL>,3,FALSE),""))))', N=r2)
ws[f"{L_STAT}4"] = F('=MAP(B4:B<N>,D4:D<N>,LAMBDA(n,st,IF(OR(n="",n="VACANT"),"",IF(st="Departed","Departed",'
             'LET(d,<ASOF>,IF(COUNTIFS(<SL_F>,n,<SL_A>,"<="&d,<SL_P>,">="&d,<SL_G>,"<>Cancelled")>0,'
             '"ON DUTY · "&IFERROR(INDEX(FILTER(<SL_D>,<SL_F>=n,<SL_A><=d,<SL_P>>=d,<SL_G><>"Cancelled"),1),""),'
             'IFERROR(INDEX(ARRAYFORMULA(FILTER(<AV_I>&" · "&<AV_B>&" until "&IF(<AV_D>="","open",TEXT(<AV_D>,"dd mmm")),'
             '<AV_A>=n,<AV_C><=d,<AV_M>>=d,<AV_J>="Y")),1),"✓ Available")))))))', N=r2)
for idx, (key, tab, mn, mx, _) in enumerate(S.ROSTERS):
    col = CL(PC_ELIG + idx)
    ws[f"{col}4"] = F('=MAP(B4:B<N>,D4:D<N>,E4:E<N>,<O>4:<O><N>,LAMBDA(n,st,ex,o,IF(AND(n<>"",n<>"VACANT",'
                      'st<>"Departed",ex<>"Y",ISNUMBER(o)),IF(AND(o>=<MIN>,o<=<MAX>),"✓",""),"")))',
                      N=r2, O=L_ORD, MIN=f"{ST}$L${4 + idx}", MAX=f"{ST}$M${4 + idx}")
ws[f"{L_DET}4"] = F('=MAP(B4:B<N>,LAMBDA(n,IF(OR(n="",n="VACANT"),"",COUNTIFS(<SL_F>,n,<SL_G>,"<>Cancelled"))))', N=r2)
ws[f"{L_LAST}4"] = F('=MAP(B4:B<N>,LAMBDA(n,IF(OR(n="",n="VACANT"),"",LET(m,MAXIFS(<SL_A>,<SL_F>,n,<SL_A>,"<="&<ASOF>,'
             '<SL_G>,"<>Cancelled"),IF(m=0,"—",m)))))', N=r2)
ws[f"{L_NEXT}4"] = F('=MAP(B4:B<N>,LAMBDA(n,IF(OR(n="",n="VACANT"),"",LET(m,MINIFS(<SL_A>,<SL_F>,n,<SL_A>,">"&<ASOF>,'
             '<SL_G>,"<>Cancelled"),IF(m=0,"—",m)))))', N=r2)
tint(ws, [CL(c) for c in range(PC_GRADE, PC_NEXT + 1)], 4, 120)
date_fmt(ws, [L_LAST, L_NEXT], 4, 120)
dv_list(ws, f"{ST}$E$4:$E$40", f"C4:C{r2}")
dv_list(ws, '"Assigned,Attached,Inbound,Departing,Departed,Vacant"', f"D4:D{r2}")
dv_list(ws, '"Y"', f"E4:E{r2}")
ws.conditional_formatting.add(f"{L_E1}4:{L_E2}{r2}", CellIsRule(operator="equal", formula=['"✓"'], fill=fill(GOOD_BG),
                                                      font=Font(name=FONT, color="006300", bold=True)))
ws.conditional_formatting.add(f"{L_STAT}4:{L_STAT}{r2}", FormulaRule(formula=[f'LEFT(${L_STAT}4,1)="A"'], fill=fill(WARN_BG)))
ws.conditional_formatting.add(f"{L_STAT}4:{L_STAT}{r2}", FormulaRule(formula=[f'LEFT(${L_STAT}4,2)="ON"'], fill=fill(GOOD_BG)))
ws.conditional_formatting.add(f"A4:{CL(PC_GRADE - 1)}{r2}", FormulaRule(formula=['$B4="VACANT"'], fill=fill(GREY_BG),
                                                       font=Font(name=FONT, italic=True, color=INK2)))
ws.freeze_panes = "C4"
ws.auto_filter.ref = f"A3:{L_NEXT}{3 + len(S.PERSONNEL)}"

# ============================================================================
# AVAILABILITY
# ============================================================================
ws = WS["Availability"]
A_HEAD = ["Soldier", "Type", "Start", "End (blank = open)", "Remarks", "Source", "Section", "Rank",
          "DA6 code", "Blocks duty", "Days", "Status", "End (calc)"]
banner(ws, "AVAILABILITY · leave, pass, TDY, school, exemptions",
       "One row per period. Codes follow DA PAM 220-45: A pauses the DA 6 number, D and U keep counting. "
       "'Request (non-binding)' only raises a flag on Duty Slots and never blocks a Soldier.", 13)
header_row(ws, 3, A_HEAD, set(A_HEAD[6:]), [30, 22, 11, 12, 46, 20, 10, 7, 8, 8, 7, 10, 11])
for i, (n, t, s, e, rm) in enumerate(S.AVAILABILITY, start=4):
    ws.cell(row=i, column=1, value=n)
    ws.cell(row=i, column=2, value=t)
    ws.cell(row=i, column=3, value=s)
    ws.cell(row=i, column=4, value=e)
    ws.cell(row=i, column=5, value=rm)
    ws.cell(row=i, column=6, value=S.SRC_NOV)
N = AVN
ws["G4"] = F('=MAP(A4:A<N>,LAMBDA(n,IF(n="","",IFERROR(INDEX(<PE_A>,MATCH(n,<PE_B>,0)),"NOT ON PERSONNEL"))))', N=N)
ws["H4"] = F('=MAP(A4:A<N>,LAMBDA(n,IF(n="","",IFERROR(VLOOKUP(n,<PE_BC>,2,FALSE),"?"))))', N=N)
ws["I4"] = F('=MAP(B4:B<N>,LAMBDA(t,IF(t="","",IFERROR(VLOOKUP(t,<ABSTBL>,2,FALSE)&"","?"))))', N=N)
ws["J4"] = F('=MAP(B4:B<N>,LAMBDA(t,IF(t="","",IFERROR(VLOOKUP(t,<ABSTBL>,3,FALSE),"Y"))))', N=N)
ws["K4"] = F('=MAP(C4:C<N>,D4:D<N>,LAMBDA(s,e,IF(s="","",IF(e="","open",e-s+1))))', N=N)
ws["L4"] = F('=MAP(C4:C<N>,M4:M<N>,LAMBDA(s,e,IF(s="","",IF(<ASOF><s,"Upcoming",IF(<ASOF>>e,"Past","ACTIVE")))))', N=N)
ws["M4"] = F('=MAP(C4:C<N>,D4:D<N>,LAMBDA(s,e,IF(s="","",IF(e="",DATE(2099,12,31),e))))', N=N)
date_fmt(ws, ["C", "D", "M"], 4, N)
tint(ws, list("GHIJKLM"), 4, 150)
dv_list(ws, f"{T['PE_B']}", f"A4:A{N}")
dv_list(ws, f"{ST}$T$4:$T$40", f"B4:B{N}")
ws.conditional_formatting.add(f"L4:L{N}", CellIsRule(operator="equal", formula=['"ACTIVE"'], fill=fill(WARN_BG),
                                                     font=Font(name=FONT, bold=True)))
ws.conditional_formatting.add(f"A4:F{N}", FormulaRule(formula=['$L4="Past"'], font=Font(name=FONT, color="9A9A9A")))
ws.conditional_formatting.add(f"G4:G{N}", CellIsRule(operator="equal", formula=['"NOT ON PERSONNEL"'],
                                                     fill=fill(CRIT_BG)))
ws.column_dimensions["M"].hidden = True
ws.freeze_panes = "B4"
ws.auto_filter.ref = f"A3:L{N}"

# ============================================================================
# DUTY SLOTS
# ============================================================================
ws = WS["Duty Slots"]
D_HEAD = ["Date", "End date (multi-day)", "Day", "Duty", "Position / detail", "Soldier", "Status",
          "Rank", "Section", "DA 6 roster", "Wknd / Hol", "Check", "Suggested (next up)", "Remarks",
          "Source", "End (calc)"]
D_AUTO = {"Day", "Rank", "Section", "DA 6 roster", "Wknd / Hol", "Check", "Suggested (next up)", "End (calc)"}
banner(ws, "DUTY SLOTS · every G2-labeled slot from the BDE roster",
       "Enter only slots labeled G2. Date + Duty create the requirement; picking a Soldier posts it to the right DA 6 "
       "automatically (CQ Runner and SD Runner share one DA 6). 'Suggested' shows the next eligible, available Soldier "
       "in DA 6 order; 'Check' flags double-tasking, back-to-back duty, rank band and availability.", 16)
header_row(ws, 3, D_HEAD, D_AUTO, [12, 12, 6, 17, 24, 28, 12, 7, 10, 12, 10, 44, 28, 40, 18, 11])
asof_seed = date(2026, 10, 13)
for i, (s, e, duty, pos, who, rm) in enumerate(S.SLOTS, start=4):
    ws.cell(row=i, column=1, value=s)
    ws.cell(row=i, column=2, value=e)
    ws.cell(row=i, column=4, value=duty)
    ws.cell(row=i, column=5, value=pos or None)
    ws.cell(row=i, column=6, value=who)
    ws.cell(row=i, column=7, value="Completed" if (e or s) < asof_seed else "Scheduled")
    ws.cell(row=i, column=14, value=rm or None)
    ws.cell(row=i, column=15, value=S.SRC_OCT)
N = SLN
ws["C4"] = F('=MAP(A4:A<N>,LAMBDA(a,IF(a="","",TEXT(a,"ddd"))))', N=N)
ws["H4"] = F('=MAP(F4:F<N>,LAMBDA(n,IF(n="","",IFERROR(VLOOKUP(n,<PE_BC>,2,FALSE),"?"))))', N=N)
ws["I4"] = F('=MAP(F4:F<N>,LAMBDA(n,IF(n="","",IFERROR(INDEX(<PE_A>,MATCH(n,<PE_B>,0)),"?"))))', N=N)
ws["J4"] = F('=MAP(D4:D<N>,LAMBDA(t,IF(t="","",IFERROR(VLOOKUP(t,<DUTYTBL>,2,FALSE)&"","N/A"))))', N=N)
ws["K4"] = F('=MAP(A4:A<N>,LAMBDA(a,IF(a="","",IF(COUNTIF(<HOL>,a)>0,"HOLIDAY",IF(WEEKDAY(a,2)>5,"WEEKEND","")))))', N=N)
ws["P4"] = F('=MAP(A4:A<N>,B4:B<N>,LAMBDA(a,b,IF(a="","",IF(b="",a,b))))', N=N)
ws["L4"] = F(
    '=MAP(A4:A<N>,F4:F<N>,G4:G<N>,J4:J<N>,P4:P<N>,LAMBDA(a,f,g,r,p,IF(a="","",IF(g="Cancelled","— cancelled",'
    'IF(f="","⚠ UNFILLED",LET(onp,ISNUMBER(MATCH(f,<PE_B>,0)),ri,IFERROR(MATCH(r,<ROSTER_ARR>,0),0),'
    'el,IF(OR(NOT(onp),ri=0),TRUE,INDEX(<PE_ELIG>,MATCH(f,<PE_B>,0),ri)="✓"),'
    'ov,COUNTIFS(<SL_F>,f,<SL_A>,"<="&p,<SL_P>,">="&a,<SL_G>,"<>Cancelled"),'
    'rs,COUNTIFS(<SL_F>,f,<SL_A>,"<="&(p+<REST>),<SL_P>,">="&(a-<REST>),<SL_G>,"<>Cancelled"),'
    'ab,IFERROR(INDEX(FILTER(<AV_B>,<AV_A>=f,<AV_C><=p,<AV_M>>=a,<AV_J>="Y"),1),""),'
    'rq,IFERROR(INDEX(FILTER(<AV_E>,<AV_A>=f,<AV_C><=p,<AV_M>>=a,<AV_J>="N"),1),""),'
    'm,TEXTJOIN(" · ",TRUE,IF(onp,"","NOT ON PERSONNEL"),IF(el,"","OUTSIDE RANK BAND / EXCLUDED"),'
    'IF(ov>1,"DOUBLE-TASKED",""),IF(rs>ov,"BACK-TO-BACK",""),IF(ab="","","UNAVAILABLE: "&ab),'
    'IF(rq="","","REQUEST: "&rq)),IF(m="","✓ OK","⚠ "&m)))))))', N=N)
NU_COLS = []
for idx in range(len(S.ROSTERS)):
    nc = CL(3 + idx * NU_W)
    NU_COLS.append(f"'Next Up'!${nc}$6:${nc}$125")
ws["M4"] = F(
    '=LET(out,SCAN("§",SEQUENCE(<CNT>),LAMBDA(acc,i,LET(used,IFERROR(REGEXEXTRACT(acc,"^([^§]*)§"),""),'
    'd,INDEX(A4:A<N>,i),r,INDEX(J4:J<N>,i),'
    'ok,AND(d<>"",INDEX(F4:F<N>,i)="",INDEX(G4:G<N>,i)<>"Cancelled",ISNUMBER(MATCH(r,<ROSTER_ARR>,0))),'
    'pk,IF(NOT(ok),"",IF(d<<ASOF>,"(past date)",IFERROR(LET(q,CHOOSE(MATCH(r,<ROSTER_ARR>,0),<QS>),'
    'pe,INDEX(P4:P<N>,i),INDEX(FILTER(q,MAP(q,LAMBDA(n,IF(n="",FALSE,AND('
    'COUNTIFS(<AV_A>,n,<AV_C>,"<="&pe,<AV_M>,">="&d,<AV_J>,"Y")=0,'
    'COUNTIFS(<SL_F>,n,<SL_A>,"<="&(pe+<REST>),<SL_P>,">="&(d-<REST>),<SL_G>,"<>Cancelled")=0,'
    'ISERROR(FIND("¦"&r&":"&n&"¦",used)),'
    'SUMPRODUCT(--ISNUMBER(FIND("¦"&n&"@"&(d+SEQUENCE(2*<REST>+1,1,-<REST>))&"¦",used)))=0))))),1)),'
    '"⚠ none available"))),'
    'nu,IF(OR(pk="",LEFT(pk,1)="⚠",LEFT(pk,1)="("),used,used&"¦"&r&":"&pk&"¦"&pk&"@"&d&"¦"),'
    'nu&"§"&pk))),MAP(out,LAMBDA(x,IFERROR(REGEXEXTRACT(x,"§(.*)$"),""))))',
    N=N, CNT=N - 3, QS=",".join(NU_COLS))
date_fmt(ws, ["A", "B", "P"], 4, N, "dd mmm yy")
tint(ws, list("CHIJKLMP"), 4, 300)
dv_list(ws, f"{ST}$P$4:$P$30", f"D4:D{N}")
dv_list(ws, f"{T['PE_B']}", f"F4:F{N}")
dv_list(ws, '"Scheduled,Completed,Swapped,Cancelled"', f"G4:G{N}")
ws.conditional_formatting.add(f"L4:L{N}", FormulaRule(formula=['LEFT($L4,1)="⚠"'], fill=fill(CRIT_BG),
                                                      font=Font(name=FONT, bold=True, color="8A1C1C")))
ws.conditional_formatting.add(f"L4:L{N}", FormulaRule(formula=['LEFT($L4,1)="✓"'], font=Font(name=FONT, color="006300")))
ws.conditional_formatting.add(f"A4:G{N}", FormulaRule(formula=['AND($A4<>"",$F4="",$G4<>"Cancelled")'], fill=fill(WARN_BG)))
ws.conditional_formatting.add(f"A4:O{N}", FormulaRule(formula=['$G4="Cancelled"'], font=Font(name=FONT, strike=True, color="9A9A9A")))
ws.conditional_formatting.add(f"K4:K{N}", FormulaRule(formula=['$K4<>""'], font=Font(name=FONT, bold=True, color="C00000")))
ws.conditional_formatting.add(f"M4:M{N}", FormulaRule(formula=['AND($M4<>"",LEFT($M4,1)<>"⚠",LEFT($M4,1)<>"(")'],
                                                      fill=fill(GOOD_BG), font=Font(name=FONT, bold=True)))
ws.conditional_formatting.add(f"A4:O{N}", FormulaRule(formula=['AND($A4<>"",$A4=INDIRECT("Settings!B9"))'], fill=fill("DCE9F9")))
ws.column_dimensions["P"].hidden = True
ws.freeze_panes = "B4"
ws.auto_filter.ref = f"A3:O{N}"

# ============================================================================
# DA 6 TABS
# ============================================================================
GR1, GR2 = 8, 127            # soldier rows
GC1, GCN = 9, 71             # grid: I .. (I+70)
GC2 = GC1 + GCN - 1          # CA
SC1 = GC2 + 2                # status matrix start (hidden)
SC2 = SC1 + GCN - 1
G1L, G2L, S1L, S2L = CL(GC1), CL(GC2), CL(SC1), CL(SC2)
for idx, (key, tab, mn, mx, nature) in enumerate(S.ROSTERS):
    ws = WS[tab]
    ws.sheet_view.showGridLines = False
    col_elig = f"{PE}${CL(PC_ELIG + idx)}$4:${CL(PC_ELIG + idx)}${PEN}"
    bf_col = f"{PE}${CL(PC_BF + idx)}$4:${CL(PC_BF + idx)}${PEN}"
    last = GC2
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    ws["A1"] = f"DUTY ROSTER · {key}"
    ws["A1"].font = font(16, True, "FFFFFF")
    ws["A1"].fill = fill(NAVY)
    ws["A1"].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
    ws["A2"] = F('="NATURE OF DUTY: ' + nature + '   ·   ORGANIZATION: "&<UNIT>&"   ·   FROM: "&TEXT(<START>,"dd mmm yyyy")'
                 '&"   ·   TO: "&IF(COUNT(<G1>4:<G2>4)=0,"(open)",TEXT(MAX(<G1>4:<G2>4),"dd mmm yyyy"))', G1=G1L, G2=G2L)
    ws["A2"].font = font(10, True, INK)
    ws["A2"].alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells(start_row=3, start_column=1, end_row=4, end_column=8)
    ws["A3"] = ("0 = performed duty (green)  ·  number = details made since last detailed  ·  A = not available, number paused (A3 = carried over)  ·  "
                "D# = other duty, number continues  ·  U# = misconduct, number continues  ·  red columns = weekend / holiday. "
                "Fully automatic from Duty Slots + Availability. Do not type in this tab.")
    ws["A3"].font = font(9, False, INK2, italic=True)
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="top", indent=1)
    heads = ["Rank", "Name", "Section", "Ord", "BF #", "Next-up", "Today", "Current #"]
    for j, h in enumerate(heads, start=1):
        c = ws.cell(row=5, column=j, value=h)
        c.font = font(10, True, "FFFFFF")
        c.fill = fill(NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
        ws.merge_cells(start_row=5, start_column=j, end_row=7, end_column=j)
    ws.row_dimensions[5].height = 46
    # posting dates (one column per date a detail was made from this roster)
    dates_f = ('=IFERROR(TRANSPOSE(SORT(UNIQUE(FILTER(<SL_A>,<SL_J>="<KEY>",<SL_F><>"",<SL_G><>"Cancelled",'
               '<SL_A>>=<START>)))),"")')
    # row 4 (masked): numeric posting dates · row 3 (masked): details made that date
    # rows 5-7: printable header (day / month / weekday, "x2" when two details on one date)
    ws[f"{G1L}4"] = F(dates_f, KEY=key)
    ws[f"{G1L}3"] = F('=MAP(<G1>4:<G2>4,LAMBDA(d,IF(d="",0,COUNTIFS(<SL_A>,d,<SL_J>,"<KEY>",<SL_F>,"<>",<SL_G>,"<>Cancelled"))))',
                      G1=G1L, G2=G2L, KEY=key)
    ws[f"{G1L}5"] = F('=MAP(<G1>4:<G2>4,LAMBDA(d,IF(d="","",TEXT(d,"dd"))))', G1=G1L, G2=G2L)
    ws[f"{G1L}6"] = F('=MAP(<G1>4:<G2>4,LAMBDA(d,IF(d="","",UPPER(TEXT(d,"mmm")))))', G1=G1L, G2=G2L)
    ws[f"{G1L}7"] = F('=MAP(<G1>4:<G2>4,<G1>3:<G2>3,LAMBDA(d,k,IF(d="","",LEFT(TEXT(d,"ddd"),2)&IF(k>1,"×"&k,""))))',
                      G1=G1L, G2=G2L)
    ws.column_dimensions[G1L] = ColumnDimension(ws, index=G1L, min=GC1, max=GC2, width=4.6)
    for c in range(GC1, GC2 + 1):
        ws.cell(row=4, column=c).number_format = "dd mmm yy"
    for r in range(GR1, GR1 + 70):
        for c in list(range(1, 9)) + list(range(GC1, GC1 + 40)):
            if c != 2:
                ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
    # names: alphabetical within grade, highest grade first (DA PAM 220-45 para 5c)
    ws["B8"] = F('=IFERROR(CHOOSECOLS(SORT(FILTER(HSTACK(<PE_B>,<PE_ORD>),<ELIG>="✓"),2,FALSE,1,TRUE),1),"")', ELIG=col_elig)
    ws["A8"] = F('=MAP(B8:B127,LAMBDA(n,IF(n="","",IFERROR(VLOOKUP(n,<PE_BC>,2,FALSE),""))))')
    ws["C8"] = F('=MAP(B8:B127,LAMBDA(n,IF(n="","",IFERROR(INDEX(<PE_A>,MATCH(n,<PE_B>,0)),""))))')
    ws["D8"] = F('=MAP(B8:B127,LAMBDA(n,IF(n="","",IFERROR(INDEX(<PE_ORD>,MATCH(n,<PE_B>,0)),""))))')
    ws["E8"] = F('=MAP(B8:B127,LAMBDA(n,IF(n="","",IFERROR(INDEX(<BF>,MATCH(n,<PE_B>,0))*1,0))))', BF=bf_col)
    ws["G8"] = F('=MAP(B8:B127,LAMBDA(n,IF(n="","",LET(d,<ASOF>,IF(COUNTIFS(<SL_F>,n,<SL_A>,"<="&d,<SL_P>,">="&d,'
                 '<SL_G>,"<>Cancelled")>0,"ON DUTY",IFERROR(INDEX(ARRAYFORMULA(FILTER(<AV_I>&" · "&<AV_B>,<AV_A>=n,<AV_C><=d,<AV_M>>=d,'
                 '<AV_J>="Y")),1),"✓"))))))')
    # status matrix (hidden): X = detailed, A / U / D, "-" = available & chargeable
    ws[f"{S1L}8"] = F(
        '=LET(nr,COUNTIF($B$8:$B$127,"?*"),nc,COUNT($<G1>$4:$<G2>$4),IF(OR(nr=0,nc=0),"",MAKEARRAY(nr,nc,LAMBDA(i,j,'
        'LET(n,INDEX($B$8:$B$127,i),d,INDEX($<G1>$4:$<G2>$4,1,j),'
        'IF(COUNTIFS(<SL_F>,n,<SL_A>,d,<SL_J>,"<KEY>",<SL_G>,"<>Cancelled")>0,"X",'
        'IF(COUNTIFS(<AV_A>,n,<AV_C>,"<="&d,<AV_M>,">="&d,<AV_I>,"A")>0,"A",'
        'IF(COUNTIFS(<AV_A>,n,<AV_C>,"<="&d,<AV_M>,">="&d,<AV_I>,"U")>0,"U",'
        'IF(COUNTIFS(<AV_A>,n,<AV_C>,"<="&d,<AV_M>,">="&d,<AV_I>,"D")+COUNTIFS(<SL_F>,n,<SL_A>,"<="&(d+<REST>),'
        '<SL_P>,">="&(d-<REST>),<SL_G>,"<>Cancelled")>0,"D","-")))))))))',
        G1=G1L, G2=G2L, KEY=key)
    cnt_core = ('IF(L=0,INDEX($E$8:$E$127,i),0)+SUMPRODUCT((idx>L)*(idx<=j)*(s<>"A")*(s<>"X")*$<G1>$3:$<G2>$3)')
    ws[f"{G1L}8"] = F(
        '=LET(nr,COUNTIF($B$8:$B$127,"?*"),nc,COUNT($<G1>$4:$<G2>$4),IF(OR(nr=0,nc=0),"",MAKEARRAY(nr,nc,LAMBDA(i,j,'
        'LET(idx,SEQUENCE(1,<W>),s,INDEX($<S1>$8:$<S2>$127,i,0),sj,INDEX(s,1,j),'
        'L,IFERROR(MAX(FILTER(idx,idx<=j,s="X")),0),c,' + cnt_core + ','
        'IF(sj="X",0,IF(sj="A",IF(AND(j=1,c>0),"A"&c,"A"),IF(sj="-",c,sj&c))))))))',
        G1=G1L, G2=G2L, S1=S1L, S2=S2L, W=GCN)
    ws["H8"] = F(
        '=LET(nr,COUNTIF($B$8:$B$127,"?*"),IF(nr=0,"",MAKEARRAY(nr,1,LAMBDA(i,z,'
        'LET(j,COUNT($<G1>$4:$<G2>$4),idx,SEQUENCE(1,<W>),s,INDEX($<S1>$8:$<S2>$127,i,0),'
        'L,IFERROR(MAX(FILTER(idx,idx<=j,s="X")),0),' + cnt_core + ')))))',
        G1=G1L, G2=G2L, S1=S1L, S2=S2L, W=GCN)
    ws["F8"] = F(
        '=LET(nr,COUNTIF($B$8:$B$127,"?*"),IF(nr=0,"",MAKEARRAY(nr,1,LAMBDA(i,z,'
        'LET(hc,$H$8:$H$127,dc,$D$8:$D$127,sq,SEQUENCE(120),hv,INDEX(hc,i),ov,INDEX(dc,i),'
        'COUNTIF(hc,">"&hv)+IF(<TIE>="Junior first",SUMPRODUCT((hc=hv)*((dc<ov)+(dc=ov)*(sq<=i))),SUMPRODUCT((hc=hv)*(sq<=i))))))))')
    # body cosmetics (borders/alignment) are applied by Apps Script Setup to keep the file small
    for c in range(1, 9):
        ws.column_dimensions[CL(c)].width = [7, 28, 10, 5, 6, 8, 18, 9][c - 1]
    ws.column_dimensions["D"].hidden = True
    ws.column_dimensions[CL(SC1 - 1)] = ColumnDimension(ws, index=CL(SC1 - 1), min=SC1 - 1, max=SC2, hidden=True)
    # Google sizes converted sheets to the used range: a marker in the last helper column keeps
    # the hidden status-matrix range ($CC:$EU) inside the grid.
    ws.cell(row=1, column=SC2, value="end")
    grid = f"{G1L}{GR1}:{G2L}{GR2}"
    hol = 'INDIRECT("Settings!Y4:Y60")'
    ws.conditional_formatting.add(grid, FormulaRule(formula=[f'AND(ISNUMBER({G1L}{GR1}),{G1L}{GR1}=0)'],
                                                    fill=fill("BFE6BF"), font=Font(name=FONT, bold=True, color="006300")))
    ws.conditional_formatting.add(grid, FormulaRule(formula=[f'LEFT({G1L}{GR1},1)="A"'], fill=fill(GREY_BG),
                                                    font=Font(name=FONT, color="6B6B6B")))
    ws.conditional_formatting.add(grid, FormulaRule(formula=[f'LEFT({G1L}{GR1},1)="D"'], fill=fill(WARN_BG)))
    ws.conditional_formatting.add(grid, FormulaRule(formula=[f'LEFT({G1L}{GR1},1)="U"'], fill=fill(CRIT_BG)))
    wk = f'AND(ISNUMBER({G1L}$4),OR(WEEKDAY({G1L}$4,2)>5,COUNTIF({hol},{G1L}$4)>0))'
    ws.conditional_formatting.add(grid, FormulaRule(formula=[wk], font=Font(name=FONT, bold=True, color="C00000")))
    hdr = f"{G1L}5:{G2L}7"
    ws.conditional_formatting.add(hdr, FormulaRule(formula=[wk], fill=fill("C00000"), font=Font(name=FONT, bold=True, color="FFFFFF")))
    ws.conditional_formatting.add(hdr, FormulaRule(formula=["TRUE"], fill=fill(SLATE), font=Font(name=FONT, bold=True, color="FFFFFF")))
    ws.conditional_formatting.add(f"{G1L}3:{G2L}4", FormulaRule(formula=["TRUE"], font=Font(name=FONT, color="FFFFFF")))
    ws.conditional_formatting.add(f"A{GR1}:H{GR2}", FormulaRule(formula=['$F8=1'], fill=fill("FFF4CC"),
                                                                font=Font(name=FONT, bold=True)))
    ws.conditional_formatting.add(f"G{GR1}:G{GR2}", FormulaRule(formula=['LEFT($G8,1)="A"'], font=Font(name=FONT, color="9C5700")))
    ws.conditional_formatting.add(f"G{GR1}:G{GR2}", FormulaRule(formula=['$G8="ON DUTY"'], fill=fill(GOOD_BG)))
    ws.freeze_panes = f"{G1L}{GR1}"

# ============================================================================
# NEXT UP
# ============================================================================
ws = WS["Next Up"]
ws.sheet_view.showGridLines = False
banner(ws, "NEXT UP · who is next on every DA 6",
       "Sorted by DA 6 position (highest number first, tie-break per Settings). 'Today' shows availability. "
       "The first line of each block shows the next open G2 slot and the Soldier the tracker suggests for it.", NU_W * NR - 1)
for idx, (key, tab, mn, mx, nature) in enumerate(S.ROSTERS):
    c0 = 1 + idx * NU_W
    ws.merge_cells(start_row=3, start_column=c0, end_row=3, end_column=c0 + 4)
    c = ws.cell(row=3, column=c0, value=f"{key} · {nature}")
    c.font = font(11, True, "FFFFFF" if idx in (0, 1, 5) else INK)
    c.fill = fill(SERIES[idx])
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = 34
    ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + 4)
    ws.cell(row=4, column=c0, value=F(
        '=IFERROR(LET(t,SORT(FILTER(HSTACK(<SL_A>,<SL_M>),<SL_J>="<KEY>",<SL_F>="",<SL_G><>"Cancelled",<SL_A>>=<ASOF>),1,TRUE),'
        '"Next open: "&TEXT(INDEX(t,1,1),"ddd dd mmm")&"  →  "&INDEX(t,1,2)),"No open G2 slots")', KEY=key))
    ws.cell(row=4, column=c0).font = font(9, True, INK)
    ws.cell(row=4, column=c0).fill = fill(PAGE)
    ws.cell(row=4, column=c0).alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[4].height = 30
    for j, h in enumerate(["Pos", "Rank", "Name", "Cur #", "Today"]):
        cc = ws.cell(row=5, column=c0 + j, value=h)
        cc.font = font(9, True, "FFFFFF")
        cc.fill = fill(NAVY)
        cc.alignment = Alignment(horizontal="center")
    t = f"'{tab}'!"
    ws.cell(row=6, column=c0, value=(
        f'=IFERROR(SORT(FILTER(HSTACK({t}$F$8:$F$127,{t}$A$8:$A$127,{t}$B$8:$B$127,{t}$H$8:$H$127,{t}$G$8:$G$127),'
        f'{t}$B$8:$B$127<>""),1,TRUE),"No eligible Soldiers")'))
    for w, cw in zip(range(5), [5, 6, 26, 6, 14]):
        ws.column_dimensions[CL(c0 + w)].width = cw
    ws.column_dimensions[CL(c0 + 5)].width = 2
    rng = f"{CL(c0)}6:{CL(c0 + 4)}125"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'${CL(c0)}6=1'], fill=fill("FFF4CC"),
                                                   font=Font(name=FONT, bold=True)))
    ws.conditional_formatting.add(f"{CL(c0 + 4)}6:{CL(c0 + 4)}125",
                                  FormulaRule(formula=[f'AND(${CL(c0 + 4)}6<>"",LEFT(${CL(c0 + 4)}6,1)<>"✓")'],
                                              font=Font(name=FONT, color="9C5700")))
ws.freeze_panes = "A6"

# ============================================================================
# APPOINTMENTS / EQUIPMENT / TASKINGS
# ============================================================================
ws = WS["Appointments"]
AP_HEAD = ["Date", "Time", "Soldier", "Type", "Location", "Status", "Notes", "Section", "Rank", "Duty conflict", "Days out"]
banner(ws, "APPOINTMENTS",
       "Track Soldier appointments. 'Duty conflict' warns if the Soldier is on a G2 duty that day. "
       "If an appointment should block duty, also add it on Availability (type Appointment).", 11)
header_row(ws, 3, AP_HEAD, set(AP_HEAD[7:]), [12, 9, 30, 16, 22, 12, 40, 10, 7, 22, 9])
for i, row in enumerate(S.APPOINTMENTS, start=4):
    for j, v in enumerate(row, start=1):
        ws.cell(row=i, column=j, value=v)
N = 1000
ws["H4"] = F('=MAP(C4:C<N>,LAMBDA(n,IF(n="","",IFERROR(INDEX(<PE_A>,MATCH(n,<PE_B>,0)),"?"))))', N=N)
ws["I4"] = F('=MAP(C4:C<N>,LAMBDA(n,IF(n="","",IFERROR(VLOOKUP(n,<PE_BC>,2,FALSE),"?"))))', N=N)
ws["J4"] = F('=MAP(A4:A<N>,C4:C<N>,LAMBDA(d,n,IF(OR(d="",n=""),"",IF(COUNTIFS(<SL_F>,n,<SL_A>,"<="&d,<SL_P>,">="&d,'
             '<SL_G>,"<>Cancelled")>0,"⚠ ON DUTY THAT DAY","✓ none"))))', N=N)
ws["K4"] = F('=MAP(A4:A<N>,LAMBDA(d,IF(d="","",d-<ASOF>)))', N=N)
date_fmt(ws, ["A"], 4, N, "ddd dd mmm yy")
tint(ws, list("HIJK"), 4, 150)
dv_list(ws, f"{T['PE_B']}", f"C4:C{N}")
dv_list(ws, '"Medical,Dental,Behavioral Health,Admin / Finance,Legal,Counseling,Training,Other"', f"D4:D{N}")
dv_list(ws, '"Scheduled,Completed,Rescheduled,No-Show,Cancelled"', f"F4:F{N}")
ws.conditional_formatting.add(f"J4:J{N}", FormulaRule(formula=['LEFT($J4,1)="⚠"'], fill=fill(CRIT_BG)))
ws.conditional_formatting.add(f"A4:K{N}", FormulaRule(formula=['AND(ISNUMBER($K4),$K4>=0,$K4<=7,$F4="Scheduled")'],
                                                      fill=fill("DCE9F9")))
ws.conditional_formatting.add(f"F4:F{N}", CellIsRule(operator="equal", formula=['"No-Show"'], fill=fill(CRIT_BG)))
ws.freeze_panes = "A4"
ws.auto_filter.ref = f"A3:K{N}"

ws = WS["Equipment"]
EQ_HEAD = ["Item / nomenclature", "LIN / NSN", "Serial #", "Hand receipt holder", "Location", "Status",
           "Last inventory", "Notes", "Section", "Next inventory", "Inventory status"]
banner(ws, "EQUIPMENT · hand receipt & status",
       "One row per serialized item. Next inventory = last inventory + cycle days (Settings). "
       "Statuses feed the dashboard.", 11)
header_row(ws, 3, EQ_HEAD, set(EQ_HEAD[8:]), [30, 16, 16, 30, 16, 12, 13, 36, 10, 13, 16])
ws["I4"] = F('=MAP(D4:D<N>,LAMBDA(n,IF(n="","",IFERROR(INDEX(<PE_A>,MATCH(n,<PE_B>,0)),"?"))))', N=N)
ws["J4"] = F('=MAP(A4:A<N>,G4:G<N>,LAMBDA(a,g,IF(OR(a="",g=""),"",g+<CYCLE>)))', N=N)
ws["K4"] = F('=MAP(A4:A<N>,F4:F<N>,J4:J<N>,LAMBDA(a,s,nx,IF(a="","",IF(s="Turned In","—",IF(nx="","NO INVENTORY DATE",'
             'IF(nx<<ASOF>,"OVERDUE",IF(nx-<ASOF><=7,"DUE ≤7 DAYS","OK")))))))', N=N)
date_fmt(ws, ["G", "J"], 4, N)
tint(ws, list("IJK"), 4, 150)
dv_list(ws, f"{T['PE_B']}", f"D4:D{N}")
dv_list(ws, '"FMC,NMC,Missing,On Loan,Turned In"', f"F4:F{N}")
ws.conditional_formatting.add(f"F4:F{N}", FormulaRule(formula=['OR($F4="NMC",$F4="Missing")'], fill=fill(CRIT_BG),
                                                      font=Font(name=FONT, bold=True)))
ws.conditional_formatting.add(f"F4:F{N}", CellIsRule(operator="equal", formula=['"FMC"'], fill=fill(GOOD_BG)))
ws.conditional_formatting.add(f"K4:K{N}", FormulaRule(formula=['OR($K4="OVERDUE",$K4="NO INVENTORY DATE")'], fill=fill(CRIT_BG)))
ws.conditional_formatting.add(f"K4:K{N}", FormulaRule(formula=['LEFT($K4,3)="DUE"'], fill=fill(WARN_BG)))
ws.freeze_panes = "B4"
ws.auto_filter.ref = f"A3:K{N}"

ws = WS["Taskings"]
TK_HEAD = ["Tasking #", "Received", "Suspense", "From", "Requirement", "Slots req", "Slots filled",
           "Personnel tasked", "Status", "POC", "Notes", "Days left", "Flag"]
banner(ws, "TASKINGS · incoming requirements & suspense",
       "Log every tasking that lands on G2 (details, augmentees, reports). Flag turns red when overdue or short.", 13)
header_row(ws, 3, TK_HEAD, {"Days left", "Flag"}, [14, 11, 11, 14, 40, 8, 8, 34, 12, 18, 30, 8, 22])
ws["L4"] = F('=MAP(A4:A<N>,C4:C<N>,LAMBDA(a,c,IF(OR(a="",c=""),"",c-<ASOF>)))', N=N)
ws["M4"] = F('=MAP(A4:A<N>,C4:C<N>,F4:F<N>,G4:G<N>,I4:I<N>,LAMBDA(a,c,rq,fl,s,IF(a="","",IF(OR(s="Closed",s="Submitted"),'
             '"✓ "&s,TEXTJOIN(" · ",TRUE,IF(c="","NO SUSPENSE",IF(c<<ASOF>,"OVERDUE",IF(c-<ASOF><=3,"DUE ≤3 DAYS",""))),'
             'IF(IF(ISNUMBER(rq),rq,0)>IF(ISNUMBER(fl),fl,0),"SHORT "&(IF(ISNUMBER(rq),rq,0)-IF(ISNUMBER(fl),fl,0)),""),IF(AND(c<>"",c>=<ASOF>,IF(ISNUMBER(rq),rq,0)<=IF(ISNUMBER(fl),fl,0)),"ON TRACK",""))))))', N=N)
date_fmt(ws, ["B", "C"], 4, N)
tint(ws, list("LM"), 4, 150)
dv_list(ws, '"Received,In Progress,Filled,Submitted,Closed"', f"I4:I{N}")
ws.conditional_formatting.add(f"M4:M{N}", FormulaRule(formula=['OR(LEFT($M4,7)="OVERDUE",ISNUMBER(SEARCH("SHORT",$M4)))'],
                                                      fill=fill(CRIT_BG), font=Font(name=FONT, bold=True)))
ws.conditional_formatting.add(f"M4:M{N}", FormulaRule(formula=['LEFT($M4,3)="DUE"'], fill=fill(WARN_BG)))
ws.conditional_formatting.add(f"M4:M{N}", FormulaRule(formula=['LEFT($M4,1)="✓"'], font=Font(name=FONT, color="006300")))
ws.freeze_panes = "B4"
ws.auto_filter.ref = f"A3:M{N}"

# ============================================================================
# DASH DATA (chart sources + KPIs)
# ============================================================================
ws = WS["Dash Data"]
ws["A1"] = "DASH DATA · chart sources and KPI math (automatic, do not edit)"
ws["A1"].font = font(12, True, NAVY)
ws["A2"] = "Month start"
ws["B2"] = F("=EOMONTH(<ASOF>,-1)+1")
ws["C2"] = "Month end"
ws["D2"] = F("=EOMONTH(<ASOF>,0)")
ws["B2"].number_format = ws["D2"].number_format = "dd mmm yy"


def table_title(r, c, text):
    cell = ws.cell(row=r, column=c, value=text)
    cell.font = font(10, True, NAVY)


# T1 duty load by section x roster (A3:F12)
table_title(3, 1, "Details by section (since roster FROM date)")
ws.cell(row=4, column=1, value="Section")
for j, k in enumerate(ROSTER_KEYS, start=2):
    ws.cell(row=4, column=j, value=k)
for i, sec in enumerate(S.SECTIONS, start=5):
    ws.cell(row=i, column=1, value=sec)
ws["B5"] = F('=MAKEARRAY(<R>,<NR>,LAMBDA(i,j,COUNTIFS(<SL_I>,INDEX($A$5:$A$<RE>,i),<SL_J>,INDEX($B$4:$<LR>$4,1,j),<SL_F>,"<>",'
             '<SL_G>,"<>Cancelled",<SL_A>,">="&<START>)))', R=len(S.SECTIONS), RE=4 + len(S.SECTIONS), NR=NR, LR=CL(1 + NR))
# T2 this month by roster (H3:M9)
table_title(3, 8, "This month by roster")
for j, h in enumerate(["Roster", "Slots", "Filled", "Open", "Conflicts", "Next open"], start=8):
    ws.cell(row=4, column=j, value=h)
for i, k in enumerate(ROSTER_KEYS, start=5):
    ws.cell(row=i, column=8, value=k)
    ws.cell(row=i, column=9, value=F(f'=COUNTIFS(<SL_J>,H{i},<SL_A>,">="&<MS>,<SL_A>,"<="&<ME>,<SL_G>,"<>Cancelled")'))
    ws.cell(row=i, column=10, value=F(f'=COUNTIFS(<SL_J>,H{i},<SL_A>,">="&<MS>,<SL_A>,"<="&<ME>,<SL_G>,"<>Cancelled",<SL_F>,"<>")'))
    ws.cell(row=i, column=11, value=f"=I{i}-J{i}")
    ws.cell(row=i, column=12, value=F(f'=COUNTIFS(<SL_J>,H{i},<SL_A>,">="&<MS>,<SL_L>,"⚠*",<SL_F>,"<>")'))
    ws.cell(row=i, column=13, value=F(f'=LET(m,MINIFS(<SL_A>,<SL_J>,H{i},<SL_F>,"",<SL_G>,"<>Cancelled",<SL_A>,">="&<ASOF>),'
                                      f'IF(m=0,"—",TEXT(m,"ddd dd mmm")))'))
# T3 unavailable next 30 days (O3:P34)
table_title(3, 15, "Soldiers unavailable, next 30 days")
ws.cell(row=4, column=15, value="Date")
ws.cell(row=4, column=16, value="Unavailable")
ws["O5"] = F("=SEQUENCE(30,1,<ASOF>)")
ws["P5"] = F('=MAP(O5:O34,LAMBDA(d,IFERROR(ROWS(UNIQUE(FILTER(<AV_A>,<AV_C><=d,<AV_M>>=d,<AV_J>="Y"))),0)))')
for i in range(5, 35):
    ws.cell(row=i, column=15).number_format = "dd mmm"
# T4 availability records this month by type (R3:S17)
table_title(3, 18, "Availability records this month, by type")
ws.cell(row=4, column=18, value="Type")
ws.cell(row=4, column=19, value="Records")
for i, (t, *_rest) in enumerate(S.ABSENCE_TYPES, start=5):
    ws.cell(row=i, column=18, value=t)
    ws.cell(row=i, column=19, value=F(f'=COUNTIFS(<AV_B>,R{i},<AV_C>,"<="&<ME>,<AV_M>,">="&<MS>)'))
# T5 details per soldier top 15 (U3:V19)
table_title(3, 21, "Details per Soldier (top 15)")
ws.cell(row=4, column=21, value="Soldier")
ws.cell(row=4, column=22, value="Details")
ws["U5"] = F('=IFERROR(SORTN(FILTER(HSTACK(<PE_C>&" "&<PE_B>,<PE_DET>),ISNUMBER(<PE_DET>),<PE_DET>>0),15,0,2,FALSE),"")')
# T6 equipment status (X3:Y9)
table_title(3, 24, "Equipment by status")
ws.cell(row=4, column=24, value="Status")
ws.cell(row=4, column=25, value="Items")
for i, s in enumerate(["FMC", "NMC", "Missing", "On Loan", "Turned In"], start=5):
    ws.cell(row=i, column=24, value=s)
    ws.cell(row=i, column=25, value=f"=COUNTIF(Equipment!$F$4:$F$1000,X{i})")
# T7 taskings status (AA3:AB9)
table_title(3, 27, "Taskings by status")
ws.cell(row=4, column=27, value="Status")
ws.cell(row=4, column=28, value="Taskings")
for i, s in enumerate(["Received", "In Progress", "Filled", "Submitted", "Closed"], start=5):
    ws.cell(row=i, column=27, value=s)
    ws.cell(row=i, column=28, value=f"=COUNTIF(Taskings!$I$4:$I$1000,AA{i})")
# KPIs (AD3:AE25)
table_title(3, 30, "KPIs")
kpis = [
    ("Assigned", F('=COUNTIFS(<PE_B>,"?*",<PE_B>,"<>VACANT",<PE_D>,"<>Departed")')),
    ("Vacant", F('=COUNTIF(<PE_B>,"VACANT")')),
    ("Unavailable today", "=P5"),
    ("Available today", "=AE5-AE7"),
    ("Present %", "=IF(AE5=0,0,AE8/AE5)"),
    ("G2 slots this month", F('=COUNTIFS(<SL_A>,">="&<MS>,<SL_A>,"<="&<ME>,<SL_G>,"<>Cancelled")')),
    ("Filled this month", F('=COUNTIFS(<SL_A>,">="&<MS>,<SL_A>,"<="&<ME>,<SL_G>,"<>Cancelled",<SL_F>,"<>")')),
    ("Fill %", "=IF(AE10=0,1,AE11/AE10)"),
    ("Open in window", F('=COUNTIFS(<SL_A>,">="&<ASOF>,<SL_A>,"<="&(<ASOF>+<LOOK>),<SL_F>,"",<SL_G>,"<>Cancelled")')),
    ("Conflicts", F('=COUNTIFS(<SL_L>,"⚠*",<SL_F>,"<>",<SL_A>,">="&<MS>)')),
    ("Open taskings", '=COUNTIFS(Taskings!$A$4:$A$1000,"<>",Taskings!$I$4:$I$1000,"<>Closed",Taskings!$I$4:$I$1000,"<>Submitted")'),
    ("Overdue / short taskings", '=COUNTIF(Taskings!$M$4:$M$1000,"OVERDUE*")+COUNTIF(Taskings!$M$4:$M$1000,"*SHORT*")'),
    ("Equipment NMC / Missing", "=Y6+Y7"),
    ("Inventories due / overdue", '=COUNTIF(Equipment!$K$4:$K$1000,"DUE*")+COUNTIF(Equipment!$K$4:$K$1000,"OVERDUE")+COUNTIF(Equipment!$K$4:$K$1000,"NO INVENTORY DATE")'),
    ("Appointments next 7 days", F('=COUNTIFS(Appointments!$A$4:$A$1000,">="&<ASOF>,Appointments!$A$4:$A$1000,"<="&(<ASOF>+7),Appointments!$F$4:$F$1000,"Scheduled")')),
    ("Duty-eligible enlisted", F('=ARRAYFORMULA(SUMPRODUCT(--(MMULT(--(<PE_ELIG>="✓"),SEQUENCE(<NR>,1,1,0))>0)))', NR=NR)),
    ("Fullest DA 6 (columns used of 71)", "=MAX(" + ",".join(f"COUNT('{r[1]}'!{G1L}4:{G2L}4)" for r in S.ROSTERS) + ")"),
]
for i, (k, v) in enumerate(kpis, start=5):
    ws.cell(row=i, column=30, value=k)
    ws.cell(row=i, column=31, value=v)
ws["AE9"].number_format = "0%"
ws["AE12"].number_format = "0%"
# Mind map nodes (AG3:AH30): node, parent
table_title(3, 33, "Mind map (org chart) nodes")
ws.cell(row=4, column=33, value="Node")
ws.cell(row=4, column=34, value="Parent")
mm = [("=\"G2 PSG · \"&TEXT(" + T["ASOF"] + ",\"dd mmm yy\")", "")]
mm.append(("Duty Rosters (DA 6)", "=$AG$5"))
for idx, (key, tab, *_r) in enumerate(S.ROSTERS):
    rk_c, nm_c = CL(2 + idx * NU_W), CL(3 + idx * NU_W)
    mm.append((f"=\"{key} → \"&IFERROR(INDEX('Next Up'!${rk_c}$6:${rk_c}$125,1)&\" \"&INDEX('Next Up'!${nm_c}$6:${nm_c}$125,1),\"none\")",
               "=$AG$6"))
mm.append(('="Personnel · "&AE5&" assigned"', "=$AG$5"))
for i, sec in enumerate(S.SECTIONS):
    mm.append((f'="{sec} ("&COUNTIFS({T["PE_A"]},"{sec}",{T["PE_D"]},"<>Departed")&")"', f"=$AG${7 + NR}"))
mm.append(('="Unavailable today · "&AE7', "=$AG$5"))
mm.append(('="G2 slots " & TEXT(B2,"mmm") & " · "&AE11&"/"&AE10&" filled"', "=$AG$5"))
mm.append(('="Taskings · "&AE15&" open"', "=$AG$5"))
mm.append(('="Appointments · "&AE19&" next 7d"', "=$AG$5"))
mm.append(('="Equipment · "&AE17&" NMC/missing"', "=$AG$5"))
for i, (n, p) in enumerate(mm, start=5):
    ws.cell(row=i, column=33, value=n)
    ws.cell(row=i, column=34, value=p)
for col, w in {"A": 12, "B": 9, "C": 9, "D": 9, "E": 9, "F": 11, "H": 12, "I": 7, "J": 7, "K": 7, "L": 9, "M": 13,
               "O": 9, "P": 11, "R": 22, "S": 8, "U": 30, "V": 8, "X": 10, "Y": 7, "AA": 12, "AB": 9, "AD": 30,
               "AE": 9, "AG": 36, "AH": 30}.items():
    ws.column_dimensions[col].width = w
for r in range(4, 5):
    for c in range(1, 35):
        cell = ws.cell(row=r, column=c)
        if cell.value:
            cell.font = font(9, True, "FFFFFF")
            cell.fill = fill(SLATE)

# ============================================================================
# DASHBOARD
# ============================================================================
ws = WS["Dashboard"]
ws.sheet_view.showGridLines = False
DD = "'Dash Data'!"
widths = {"A": 2, "B": 24, "C": 10, "D": 15, "E": 15, "F": 15, "G": 15, "H": 2, "I": 12, "J": 15, "K": 30,
          "L": 11, "M": 38, "N": 2}
for k, w in widths.items():
    ws.column_dimensions[k].width = w
ws.merge_cells("B1:M1")
ws["B1"] = "G2 · PLATOON SERGEANT DASHBOARD"
ws["B1"].font = font(18, True, "FFFFFF")
ws["B1"].fill = fill(NAVY)
ws["B1"].alignment = Alignment(vertical="center", indent=1)
ws.row_dimensions[1].height = 36
ws.merge_cells("B2:M2")
ws["B2"] = F('="As of "&TEXT(<ASOF>,"dddd, dd mmmm yyyy")&"   ·   PSG: "&<PSG>&"   ·   DA 6 FROM "&TEXT(<START>,"dd mmm yy")'
             '&"   ·   AR 220-45 / DA PAM 220-45 method"')
ws["B2"].font = font(10, False, "FFFFFF", italic=True)
ws["B2"].fill = fill(NAVY)
ws["B2"].alignment = Alignment(vertical="center", indent=1)
ws.row_dimensions[2].height = 22


def tile(r, c, label, value_f, note_f, fmt=None):
    """KPI tile: 2 columns x 3 rows (label / value / note)."""
    for rr in range(r, r + 3):
        ws.merge_cells(start_row=rr, start_column=c, end_row=rr, end_column=c + 1)
        for cc in (c, c + 1):
            cell = ws.cell(row=rr, column=cc)
            cell.fill = fill("FFFFFF")
            cell.border = Border(left=thin if cc == c else None, right=thin if cc == c + 1 else None,
                                 top=thin if rr == r else None, bottom=thin if rr == r + 2 else None)
    a = ws.cell(row=r, column=c, value=label)
    a.font = font(9, True, INK2)
    a.alignment = Alignment(horizontal="center", vertical="bottom")
    v = ws.cell(row=r + 1, column=c, value=value_f)
    v.font = font(24, True, NAVY)
    ws.row_dimensions[r + 1].height = 36
    v.alignment = Alignment(horizontal="center", vertical="center")
    if fmt:
        v.number_format = fmt
    n = ws.cell(row=r + 2, column=c, value=note_f)
    n.font = font(9, False, INK2)
    n.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)


tiles = [
    ("ASSIGNED STRENGTH", f"={DD}AE5", f'="Vacant: "&{DD}AE6&"  ·  duty-eligible: "&{DD}AE20', None),
    ("PRESENT FOR DUTY", f"={DD}AE9", f'={DD}AE8&" available  ·  "&{DD}AE7&" out today"', "0%"),
    ("G2 SLOTS FILLED (MONTH)", f"={DD}AE12", f'={DD}AE11&" of "&{DD}AE10&" · "&TEXT({DD}B2,"mmm yy")', "0%"),
    ("OPEN SLOTS (LOOK-AHEAD)", f"={DD}AE13", F('=IF(<DD>AE13=0,"✓ none in next "&<LOOK>&" days","⚠ fill within "&<LOOK>&" days")', DD=DD), None),
    ("CONFLICTS FLAGGED", f"={DD}AE14", f'=IF({DD}AE14=0,"✓ none","⚠ see Duty Slots ▸ Check")', None),
    ("NEXT 7 DAYS · APPOINTMENTS", f"={DD}AE19", '="scheduled"', None),
    ("OPEN TASKINGS", f"={DD}AE15", f'=IF({DD}AE16=0,"✓ none overdue / short","⚠ "&{DD}AE16&" overdue or short")', None),
    ("EQUIPMENT NMC / MISSING", f"={DD}AE17", f'={DD}AE18&" inventories due or overdue"', None),
    ("DA 6 CAPACITY", f'={DD}AE21&"/71"', '="columns used on fullest DA 6"', None),
]
for k, (lab, vf, nf, fmt) in enumerate(tiles):
    r = 4 + (k // 3) * 4
    c = 2 + (k % 3) * 2
    tile(r, c, lab, vf, nf, fmt)
# status coloring on tiles (icon + label already in the note)
ws.conditional_formatting.add("B9", FormulaRule(formula=['INDIRECT("\'Dash Data\'!AE13")>0'], font=Font(name=FONT, bold=True, size=24, color=CRIT)))
ws.conditional_formatting.add("D9", FormulaRule(formula=['INDIRECT("\'Dash Data\'!AE14")>0'], font=Font(name=FONT, bold=True, size=24, color="C77700")))
ws.conditional_formatting.add("F5", FormulaRule(formula=['INDIRECT("\'Dash Data\'!AE12")<1'], font=Font(name=FONT, bold=True, size=24, color="C77700")))
ws.conditional_formatting.add("B13", FormulaRule(formula=['INDIRECT("\'Dash Data\'!AE16")>0'], font=Font(name=FONT, bold=True, size=24, color=CRIT)))
ws.conditional_formatting.add("D13", FormulaRule(formula=['INDIRECT("\'Dash Data\'!AE17")>0'], font=Font(name=FONT, bold=True, size=24, color=CRIT)))
ws.conditional_formatting.add("F13", FormulaRule(formula=['INDIRECT("\'Dash Data\'!AE21")>60'], font=Font(name=FONT, bold=True, size=24, color=CRIT)))

# BLUF (I4:M7)
ws.merge_cells("I4:M4")
ws["I4"] = "BLUF"
ws["I4"].font = font(10, True, "FFFFFF")
ws["I4"].fill = fill(NAVY)
ws["I4"].alignment = Alignment(indent=1, vertical="center")
ws.merge_cells("I5:M7")
ws["I5"] = F(
    '=<DD>AE11&" of "&<DD>AE10&" G2 duty slots filled for "&TEXT(<DD>B2,"mmmm yyyy")&" ("&TEXT(<DD>AE12,"0%")&"). "'
    '&IF(<DD>AE13=0,"No open slots in the next "&<LOOK>&" days. ",<DD>AE13&" slot(s) open in the next "&<LOOK>&" days: action required. ")'
    '&IF(<DD>AE14=0,"No duty conflicts. ",<DD>AE14&" duty conflict(s) flagged. ")'
    '&<DD>AE7&" of "&<DD>AE5&" Soldiers unavailable today ("&TEXT(<DD>AE9,"0%")&" present). "'
    '&<DD>AE15&" open tasking(s)"&IF(<DD>AE16>0," ("&<DD>AE16&" overdue/short)","")&". "'
    '&IF(<DD>AE17>0,<DD>AE17&" equipment item(s) NMC/missing.","Equipment: no NMC/missing items.")', DD=DD)
ws["I5"].font = font(11, False, INK)
ws["I5"].fill = fill("FFFFFF")
ws["I5"].alignment = Alignment(wrap_text=True, vertical="top", indent=1)
for r in range(5, 8):
    ws.row_dimensions[r].height = 22

# NEXT UP (I9:M14)
ws.merge_cells("I9:M9")
ws["I9"] = "NEXT UP · top 3 per DA 6 (number in brackets; ⛔ = not available today)"
ws["I9"].font = font(10, True, "FFFFFF")
ws["I9"].fill = fill(NAVY)
ws["I9"].alignment = Alignment(indent=1, vertical="center")
for idx, (key, tab, *_r) in enumerate(S.ROSTERS):
    r = 10 + idx
    b = idx * NU_W
    rk, nm, cu, td = (f"'Next Up'!${CL(b + k)}$6:${CL(b + k)}$125" for k in (2, 3, 4, 5))
    ws.cell(row=r, column=9, value=key).font = font(10, True, "FFFFFF" if idx in (0, 1, 5) else INK)
    ws.cell(row=r, column=9).fill = fill(SERIES[idx])
    ws.cell(row=r, column=9).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=r, column=10, value=(
        f"=IFERROR(SPARKLINE(FILTER('{tab}'!$H$8:$H$127,'{tab}'!$B$8:$B$127<>\"\"),"
        f"{{\"charttype\",\"column\";\"color\",\"#{SERIES[idx]}\"}}),\"\")"))
    ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=13)
    ws.cell(row=r, column=11, value=(
        f'=IFERROR(TEXTJOIN("    ",TRUE,MAP(SEQUENCE(3),LAMBDA(i,IFERROR(i&") "&INDEX({rk},i)&" "&INDEX({nm},i)'
        f'&" ["&INDEX({cu},i)&"]"&IF(LEFT(INDEX({td},i),1)="✓",""," ⛔"),"")))),"—")'))
    ws.cell(row=r, column=11).font = font(10, False, INK)
    for c in range(10, 14):
        ws.cell(row=r, column=c).fill = fill("FFFFFF")
        ws.cell(row=r, column=c).border = Border(bottom=Side(style="hair", color=LINE))
    ws.cell(row=r, column=11).alignment = Alignment(vertical="center", indent=1, shrink_to_fit=True)
    ws.row_dimensions[r].height = 22

# THIS MONTH BY ROSTER (B16:G17+NR)
ws.merge_cells("B16:G16")
ws["B16"] = F('="THIS MONTH BY DA 6 ROSTER · "&UPPER(TEXT(<DD>B2,"mmmm yyyy"))', DD=DD)
ws["B16"].font = font(10, True, "FFFFFF")
ws["B16"].fill = fill(NAVY)
ws["B16"].alignment = Alignment(indent=1, vertical="center")
for j, h in enumerate(["Roster", "G2 slots", "Filled", "Open", "Conflicts", "Next open"], start=2):
    c = ws.cell(row=17, column=j, value=h)
    c.font = font(9, True, INK2)
    c.alignment = Alignment(horizontal="center")
    c.fill = fill("FFFFFF")
    c.border = Border(bottom=thin)
for i in range(NR):
    r = 18 + i
    for j, src in enumerate(["H", "I", "J", "K", "L", "M"], start=2):
        c = ws.cell(row=r, column=j, value=f"={DD}{src}{5 + i}")
        c.font = font(10, j == 2)
        c.alignment = Alignment(horizontal="center")
        c.fill = fill("FFFFFF")
        c.border = Border(bottom=Side(style="hair", color=LINE))
ws.conditional_formatting.add(f"E18:E{17 + NR}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, bold=True, color=CRIT)))
ws.conditional_formatting.add(f"F18:F{17 + NR}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, bold=True, color="C77700")))

# UNAVAILABLE TODAY (B24:G35)
ws.merge_cells("B24:G24")
ws["B24"] = "UNAVAILABLE TODAY (blocking availability records)"
ws["B24"].font = font(10, True, "FFFFFF")
ws["B24"].fill = fill(NAVY)
ws["B24"].alignment = Alignment(indent=1, vertical="center")
for j, h in enumerate(["Soldier", "Section", "Until", "Type", "", ""], start=2):
    c = ws.cell(row=25, column=j, value=h or None)
    c.font = font(9, True, INK2)
    c.fill = fill("FFFFFF")
    c.border = Border(bottom=thin)
ws["B26"] = F(
    '=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(SORT(FILTER(HSTACK(<AV_H>&" "&<AV_A>,<AV_G>,IF(<AV_D>="","open",TEXT(<AV_D>,"dd mmm yy")),'
    '<AV_B>),<AV_C><=<ASOF>,<AV_M>>=<ASOF>,<AV_J>="Y"),1,TRUE),10,4),'
    '"✓ No one unavailable today"))')
for r in range(26, 36):
    for c in range(2, 8):
        cell = ws.cell(row=r, column=c)
        cell.fill = fill("FFFFFF")
        cell.font = font(9, c == 2)
        cell.border = Border(bottom=Side(style="hair", color=LINE))

# UPCOMING G2 DUTIES (I16:M29)
ws.merge_cells("I16:M16")
ws["I16"] = F('="UPCOMING G2 DUTIES · NEXT "&<LOOK>&" DAYS"')
ws["I16"].font = font(10, True, "FFFFFF")
ws["I16"].fill = fill(NAVY)
ws["I16"].alignment = Alignment(indent=1, vertical="center")
for j, h in enumerate(["Date", "Duty", "Soldier", "Section", "Check"], start=9):
    c = ws.cell(row=17, column=j, value=h)
    c.font = font(9, True, INK2)
    c.fill = fill("FFFFFF")
    c.border = Border(bottom=thin)
ws["I18"] = F(
    '=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(SORT(FILTER(HSTACK(<SL_A>,<SL_D>,IF(<SL_F>="","— OPEN —",<SL_H>&" "&<SL_F>),'
    'IF(<SL_F>="","",<SL_I>),<SL_L>),<SL_A>>=<ASOF>,<SL_A><=<ASOF>+<LOOK>,<SL_G><>"Cancelled"),1,TRUE),14,5),'
    '"✓ No G2 duties in the window"))')
for r in range(18, 32):
    ws[f"I{r}"].number_format = "ddd dd mmm"
    for c in range(9, 14):
        cell = ws.cell(row=r, column=c)
        cell.fill = fill("FFFFFF")
        cell.font = font(9, c == 11)
        cell.border = Border(bottom=Side(style="hair", color=LINE))
ws.conditional_formatting.add("I18:M31", FormulaRule(formula=['$K18="— OPEN —"'], fill=fill(WARN_BG), font=Font(name=FONT, bold=True)))
ws.conditional_formatting.add("M18:M31", FormulaRule(formula=['LEFT($M18,1)="⚠"'], font=Font(name=FONT, bold=True, color=CRIT)))

# TRACKERS (I33:M40) - registry, grows when Apps Script adds tables
ws.merge_cells("I33:M33")
ws["I33"] = "TRACKER TABLES (add more from the G2 PSG menu ▸ Add tracker table)"
ws["I33"].font = font(10, True, "FFFFFF")
ws["I33"].fill = fill(NAVY)
ws["I33"].alignment = Alignment(indent=1, vertical="center")
ws["I34"] = f'=IFERROR(FILTER(HSTACK({ST}$AE$4:$AE$30,{ST}$AH$4:$AH$30),{ST}$AE$4:$AE$30<>""),"")'
ws["K34"] = f'=IFERROR(FILTER({ST}$AG$4:$AG$30,{ST}$AE$4:$AE$30<>""),"")'
for r in range(34, 44):
    for c in range(9, 14):
        cell = ws.cell(row=r, column=c)
        cell.fill = fill("FFFFFF")
        cell.font = font(9, c == 9)
        cell.border = Border(bottom=Side(style="hair", color=LINE))
    ws[f"J{r}"].number_format = '0" rows"'
ws.merge_cells("B37:G37")
ws["B37"] = "CHARTS · refresh automatically from the tabs (Dash Data holds the math)"
ws["B37"].font = font(10, True, "FFFFFF")
ws["B37"].fill = fill(NAVY)
ws["B37"].alignment = Alignment(indent=1, vertical="center")


# ---------------------------------------------------------------- charts
def style_chart(ch, title, y_title=None):
    ch.title = title
    ch.style = 10
    ch.height = 8.5
    ch.width = 16.5
    ch.legend.position = "b"
    if y_title:
        ch.y_axis.title = y_title
    ch.y_axis.majorGridlines = None
    return ch


dd = WS["Dash Data"]
# 1 stacked bar: details by section x roster
c1 = BarChart()
c1.type = "col"
c1.grouping = "stacked"
c1.overlap = 100
c1.add_data(Reference(dd, min_col=2, max_col=1 + NR, min_row=4, max_row=4 + len(S.SECTIONS)), titles_from_data=True)
c1.set_categories(Reference(dd, min_col=1, min_row=5, max_row=4 + len(S.SECTIONS)))
for i, s in enumerate(c1.series):
    s.graphicalProperties.solidFill = SERIES[i]
    s.graphicalProperties.line.solidFill = SURF
style_chart(c1, "G2 duty details by section", "Details")
ws.add_chart(c1, "B39")
# 2 roster fill this month
c2 = BarChart()
c2.type = "col"
c2.grouping = "stacked"
c2.overlap = 100
c2.add_data(Reference(dd, min_col=10, max_col=11, min_row=4, max_row=4 + NR), titles_from_data=True)
c2.set_categories(Reference(dd, min_col=8, min_row=5, max_row=4 + NR))
c2.series[0].graphicalProperties.solidFill = SERIES[0]
c2.series[1].graphicalProperties.solidFill = "C9CED6"
style_chart(c2, "This month: filled vs open G2 slots", "Slots")
ws.add_chart(c2, "I45")
# 3 unavailable next 30 days
c3 = LineChart()
c3.add_data(Reference(dd, min_col=16, min_row=4, max_row=34), titles_from_data=True)
c3.set_categories(Reference(dd, min_col=15, min_row=5, max_row=34))
c3.series[0].graphicalProperties.line.solidFill = SERIES[0]
c3.series[0].graphicalProperties.line.width = 25400
c3.series[0].smooth = False
style_chart(c3, "Soldiers unavailable, next 30 days", "Soldiers")
c3.legend = None
ws.add_chart(c3, "B57")
# 4 details per soldier (fairness)
c4 = BarChart()
c4.type = "bar"
c4.add_data(Reference(dd, min_col=22, min_row=4, max_row=19), titles_from_data=True)
c4.set_categories(Reference(dd, min_col=21, min_row=5, max_row=19))
c4.series[0].graphicalProperties.solidFill = SERIES[0]
style_chart(c4, "Duty load per Soldier (top 15): fairness check", None)
c4.legend = None
c4.height = 10
c4.dataLabels = DataLabelList()
c4.dataLabels.showVal = True
ws.add_chart(c4, "I63")
# 5 availability types this month
c5 = BarChart()
c5.type = "bar"
c5.add_data(Reference(dd, min_col=19, min_row=4, max_row=4 + len(S.ABSENCE_TYPES)), titles_from_data=True)
c5.set_categories(Reference(dd, min_col=18, min_row=5, max_row=4 + len(S.ABSENCE_TYPES)))
c5.series[0].graphicalProperties.solidFill = SERIES[2]
style_chart(c5, "Availability records this month, by type", None)
c5.legend = None
c5.height = 10
ws.add_chart(c5, "B75")
# 6 equipment + 7 taskings
c6 = BarChart()
c6.type = "col"
c6.add_data(Reference(dd, min_col=25, min_row=4, max_row=9), titles_from_data=True)
c6.set_categories(Reference(dd, min_col=24, min_row=5, max_row=9))
c6.series[0].graphicalProperties.solidFill = SERIES[6]
style_chart(c6, "Equipment by status", "Items")
c6.legend = None
ws.add_chart(c6, "I84")
c7 = BarChart()
c7.type = "col"
c7.add_data(Reference(dd, min_col=28, min_row=4, max_row=9), titles_from_data=True)
c7.set_categories(Reference(dd, min_col=27, min_row=5, max_row=9))
c7.series[0].graphicalProperties.solidFill = SERIES[3]
style_chart(c7, "Taskings by status", "Taskings")
c7.legend = None
ws.add_chart(c7, "B95")
ws.merge_cells("I102:M102")
ws["I102"] = "MIND MAP · run G2 PSG ▸ Setup once to draw the org-chart mind map here"
ws["I102"].font = font(10, True, "FFFFFF")
ws["I102"].fill = fill(NAVY)
ws["I102"].alignment = Alignment(indent=1, vertical="center")

# ============================================================================
# HOW TO (quick start; the full guide is the "G2 PSG Tracker - Guide" document)
# ============================================================================
ws = wb.create_sheet("How To")
ws.sheet_properties.tabColor = "BFBFBF"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 130
howto = [
    ("G2 PSG TRACKER · QUICK START", "h1"),
    ("First time (5 minutes)", "h2"),
    ("1. Extensions ▸ Apps Script ▸ paste Code.gs ▸ Save ▸ reload this sheet ▸ menu G2 PSG ▸ Setup / repair ▸ approve permissions.", ""),
    ("2. Settings: confirm rank bands, holidays/DONSAs, rest window and tie-break. Personnel: enter any brought-forward (BF #) numbers from the outgoing PSG.", ""),
    ("3. " + S.HOWTO_AVAILABILITY, ""),
    ("Daily loop", "h2"),
    ("1. Dashboard: read the BLUF; act on OPEN SLOTS and CONFLICTS (red/amber tiles).", ""),
    ("2. New BDE roster: paste it into a new tab ▸ G2 PSG ▸ Import G2 slots (or type the G2-labeled slots on Duty Slots).", ""),
    ("3. Open slots: use 'Suggested (next up)' or G2 PSG ▸ Auto-fill open G2 slots. Leave/pass/TDY/appointments go on Availability.", ""),
    ("4. G2 PSG ▸ Publish shared view gives section sergeants a no-PII copy; the request form feeds the Requests tab (tick Approve).", ""),
    ("DA 6 rules applied (AR 220-45 / DA PAM 220-45, 2024)", "h2"),
    ("Number = details made from the roster since last detailed · 0 = did the duty · A = authorized absence, number paused · D = other duty, "
     "number continues · U = misconduct, number continues · highest number is next; if unavailable, next highest · names alphabetical "
     "within grade, highest grade first · weekend/holiday postings in red · a separate DA 6 for each duty (AR 220-45 para 4b): CQ Runner and SD Runner each have their own, and a Soldier tasked on one is coded D on the other, so no one is double tasked.", ""),
    ("Extra duty under UCMJ Art. 15 is punishment, not a DA 6 detail: log it on Availability as 'Extra duty / restriction' (U). Never use an honor duty such as the flag detail as extra duty (AR 27-10 para 3-19b(4)(b)).", ""),
    ("AR 220-45 app B-4a: whoever maintains the DA 6 needs a training certificate or MFR on AR 220-45 and DA PAM 220-45.", ""),
    ("Colors", "h2"),
    ("Navy headers = you type. Grey headers / tinted columns = automatic: never type over them. Green = OK · amber = attention · red = action.", ""),
    ("Verify from the build", "h2"),
    (S.HOWTO_VERIFY, ""),
]
for i, (text, kind) in enumerate(howto, start=1):
    c = ws.cell(row=i, column=2, value=text)
    if kind == "h1":
        c.font = font(16, True, "FFFFFF")
        c.fill = fill(NAVY)
        ws.row_dimensions[i].height = 30
    elif kind == "h2":
        c.font = font(11, True, NAVY)
    else:
        c.font = font(10)
        c.alignment = Alignment(wrap_text=True, vertical="top")

from openpyxl.utils import column_index_from_string as CI
by_sheet = {}
for w, rng in TINTS:
    col = rng.split(":")[0].rstrip("0123456789")
    r1 = int(rng.split(":")[0][len(col):])
    by_sheet.setdefault((w.title, r1), (w, set()))[1].add(CI(col))
for (title, r1), (w, cols) in by_sheet.items():
    cols = sorted(cols)
    runs, start, prev = [], cols[0], cols[0]
    for c in cols[1:] + [None]:
        if c is not None and c == prev + 1:
            prev = c
            continue
        runs.append(f"{CL(start)}{r1}:{CL(prev)}1000")
        if c is not None:
            start = prev = c
    w.conditional_formatting.add(" ".join(runs), FormulaRule(formula=["TRUE"], fill=fill(AUTO)))

out = sys.argv[1] if len(sys.argv) > 1 else "G2_PSG_Tracker.xlsx"
if os.environ.get("VERIFY"):
    # trimmed copy used only to verify the DA 6 engine in Google Sheets
    keep = {"Settings", "Personnel", "Availability", "Duty Slots", "DA6 CQ RUNNER"}
    for name in list(wb.sheetnames):
        if name not in keep:
            del wb[name]
    for w in wb.worksheets:
        w.conditional_formatting = type(w.conditional_formatting)()
        w.data_validations.dataValidation = []
        for row in w.iter_rows():
            for c in row:
                c.style = "Normal"
                if (isinstance(c.value, str) and len(c.value) > 60 and not c.value.startswith("=")
                        and c.__class__.__name__ == "Cell"):
                    c.value = None
    ws_ = wb["Duty Slots"]
    ws_["L4"] = ws_["M4"] = None
    for w, cols in ((wb["Duty Slots"], "ENO"), (wb["Availability"], "EF"), (wb["Personnel"], "GI")):
        for col in cols:
            for r in range(1, 200):
                if w[f"{col}{r}"].__class__.__name__ == "Cell":
                    w[f"{col}{r}"] = None
    for w in wb.worksheets:
        for r in (1, 2):
            for c in range(1, 30):
                if w.cell(row=r, column=c).__class__.__name__ == "Cell":
                    w.cell(row=r, column=c).value = None
        w.merged_cells.ranges = set()
    ws_ = wb["Personnel"]
    for col in ("Q", "W", "X", "Y"):
        ws_[col + "4"] = None
    ws_ = wb["Settings"]
    for row in ws_.iter_rows(min_col=25):
        for c in row:
            if c.column >= 28 and c.__class__.__name__ == "Cell":
                c.value = None
wb.save(out)

# re-pack at maximum compression (the file is uploaded inline as base64)
import os
import re
import zipfile
tmp = out + ".tmp"
SST = {}


def _to_shared(m):
    """inline string cell -> shared string cell (smaller once compressed)."""
    attrs, text = m.group(1), m.group(2)
    idx = SST.setdefault(text, len(SST))
    return '<c%s t="s"><v>%d</v></c>' % (attrs, idx)


with zipfile.ZipFile(out) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
    names = [i.filename for i in zin.infolist()]
    for item in zin.infolist():
        name = item.filename
        if name.startswith("docProps/") or name.startswith("xl/theme/"):
            continue
        data = zin.read(name)
        if name == "xl/_rels/workbook.xml.rels":
            data = re.sub(rb'<Relationship[^>]*theme[^>]*/>', b"", data)
        if name == "[Content_Types].xml":
            data = re.sub(rb'<Override[^>]*(theme|docProps)[^>]*/>', b"", data)
        if name == "_rels/.rels":
            data = re.sub(rb'<Relationship[^>]*docProps[^>]*/>', b"", data)
        if name.startswith("xl/worksheets/sheet"):
            txt = data.decode("utf-8")
            txt = re.sub(r'<c([^>]*?) t="inlineStr"><is><t(?: xml:space="preserve")?>(.*?)</t></is></c>', _to_shared, txt, flags=re.S)
            data = txt.encode("utf-8")
        if name == "xl/_rels/workbook.xml.rels" and SST is not None:
            data = data.replace(b"</Relationships>",
                                b'<Relationship Id="rIdSST" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                                b'relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>')
        if name == "[Content_Types].xml":
            data = data.replace(b"</Types>", b'<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
                                b'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>')
        zout.writestr(name, data)
    sst = "".join('<si><t xml:space="preserve">%s</t></si>' % t for t in SST)
    zout.writestr("xl/sharedStrings.xml",
                  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/'
                  'spreadsheetml/2006/main" count="%d" uniqueCount="%d">%s</sst>' % (len(SST), len(SST), sst))
os.replace(tmp, out)
print("saved", out, os.path.getsize(out), "bytes")
