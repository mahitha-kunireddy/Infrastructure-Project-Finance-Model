"""
PROJECT 2 — Infrastructure Project Finance Model (Toll Road BOT Concession)
===========================================================================
Builds a fully formula-driven DCF / project finance model in Excel.

Structure:
    Assumptions    — every input in one labelled cell (blue = editable)
    Cash Flow      — 22-year build + operate projection
    Returns        — NPV, project IRR, equity IRR, payback, DSCR, sensitivity

All operating assumptions are ILLUSTRATIVE placeholders and are flagged as such
in the workbook. Replace with figures sourced from a real concession agreement
or NHAI/NIP project document before using this in an interview.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- styling ---
F_TITLE = Font(name="Arial", size=14, bold=True, color="FFFFFF")
F_HDR = Font(name="Arial", size=10, bold=True, color="FFFFFF")
F_LBL = Font(name="Arial", size=10, bold=True)
F_TXT = Font(name="Arial", size=10)
F_IN = Font(name="Arial", size=10, color="0000FF")          # blue = input
F_CALC = Font(name="Arial", size=10)                        # black = formula
F_LINK = Font(name="Arial", size=10, color="008000")        # green = x-sheet
F_NOTE = Font(name="Arial", size=8, italic=True, color="595959")
F_KEY = Font(name="Arial", size=11, bold=True)

FILL_DARK = PatternFill("solid", fgColor="1F1F1F")
FILL_HDR = PatternFill("solid", fgColor="404040")
FILL_KEY = PatternFill("solid", fgColor="FFFF00")           # key assumption
FILL_SUB = PatternFill("solid", fgColor="D9D9D9")
FILL_OUT = PatternFill("solid", fgColor="F2F2F2")

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOPLINE = Border(top=Side(style="medium", color="404040"))

CUR = '#,##0;(#,##0);-'          # INR mn
CUR2 = '#,##0.00;(#,##0.00);-'
PCT = '0.0%'
PCT2 = '0.00%'
NUM = '#,##0'
MULT = '0.00x'

wb = Workbook()

# ============================================================================
# SHEET 1 — ASSUMPTIONS
# ============================================================================
ws = wb.active
ws.title = "Assumptions"
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 46
ws.column_dimensions["C"].width = 16
ws.column_dimensions["D"].width = 12
ws.column_dimensions["E"].width = 58

ws["B2"] = "TOLL ROAD BOT CONCESSION — PROJECT FINANCE MODEL"
ws["B2"].font = F_TITLE
ws["B2"].fill = FILL_DARK
for c in "CDE":
    ws[f"{c}2"].fill = FILL_DARK
ws.row_dimensions[2].height = 24

ws["B3"] = "Blue = input cell (edit these)   |   Black = formula   |   Yellow = key assumption driving returns"
ws["B3"].font = F_NOTE


def block(row, title):
    ws.cell(row, 2, title).font = F_HDR
    for c in range(2, 6):
        ws.cell(row, c).fill = FILL_HDR
    return row + 1


def inp(row, label, value, fmt=NUM, note="", key=False, unit=""):
    ws.cell(row, 2, label).font = F_LBL
    c = ws.cell(row, 3, value)
    c.font = F_IN
    c.number_format = fmt
    c.border = BOX
    if key:
        c.fill = FILL_KEY
    ws.cell(row, 4, unit).font = F_NOTE
    ws.cell(row, 5, note).font = F_NOTE
    return row + 1


r = 5
r = block(r, "PROJECT & TIMELINE")
r = inp(r, "Concession length — construction", 2, NUM, "Typical 2-3 yr build for a 4-lane highway package", unit="years")
R_CONSTR = r - 1
r = inp(r, "Concession length — operations", 20, NUM, "NHAI BOT concessions commonly run 15-30 yrs post-COD", unit="years")
R_OPS = r - 1
r = inp(r, "Highway length", 60, NUM, "Placeholder — set from the concession agreement", unit="km")
R_KM = r - 1

r += 1
r = block(r, "CAPITAL COST")
r = inp(r, "Base construction cost per km", 95, CUR, "ILLUSTRATIVE. Source from NHAI bid/NIP docs.", unit="INR mn/km")
R_CPK = r - 1
r = inp(r, "Total project cost", None, CUR, "Length x cost per km", unit="INR mn")
R_TPC = r - 1
ws.cell(R_TPC, 3, f"=C{R_KM}*C{R_CPK}").font = F_CALC
ws.cell(R_TPC, 3).number_format = CUR
r = inp(r, "Capex split — construction year 1", 0.45, PCT, "Drawdown profile across the build period", key=True)
R_SPLIT1 = r - 1
r = inp(r, "Contingency on base cost", 0.05, PCT, "Cost overrun buffer")
R_CONT = r - 1

r += 1
r = block(r, "TRAFFIC & TARIFF")
r = inp(r, "Opening year traffic (AADT)", 24000, NUM, "Annual Average Daily Traffic, all vehicle classes (PCU)", key=True, unit="vehicles/day")
R_AADT = r - 1
r = inp(r, "Traffic CAGR", 0.055, PCT, "ILLUSTRATIVE. Ties to regional GDP/freight growth.", key=True)
R_TGROW = r - 1
r = inp(r, "Traffic ramp-up — year 1 of ops", 0.75, PCT, "Traffic builds to full run-rate over first 2-3 yrs")
R_RAMP1 = r - 1
r = inp(r, "Traffic ramp-up — year 2 of ops", 0.90, PCT, "")
R_RAMP2 = r - 1
r = inp(r, "Blended toll per vehicle (opening)", 115, CUR2, "Weighted across car/LCV/truck mix", key=True, unit="INR/trip")
R_TOLL = r - 1
r = inp(r, "Annual toll escalation", 0.04, PCT, "NHAI formula: ~3% + 40% of WPI inflation", key=True)
R_TESC = r - 1
r = inp(r, "Toll leakage / exemptions", 0.07, PCT, "Local exempt traffic, evasion, FASTag failures")
R_LEAK = r - 1

r += 1
r = block(r, "OPERATING COST")
r = inp(r, "Routine O&M — year 1 of ops", 210, CUR, "Patrolling, toll plaza staffing, minor repairs", unit="INR mn")
R_OM = r - 1
r = inp(r, "O&M cost inflation", 0.055, PCT, "")
R_OMINF = r - 1
r = inp(r, "Major maintenance cycle", 5, NUM, "Periodic resurfacing / overlay", unit="years")
R_MMCYC = r - 1
r = inp(r, "Major maintenance cost (per event)", 720, CUR, "Bituminous overlay across full length", unit="INR mn")
R_MMCOST = r - 1

r += 1
r = block(r, "FINANCING & TAX")
r = inp(r, "Debt share of project cost", 0.70, PCT, "70:30 D/E is standard for Indian BOT/HAM", key=True)
R_DEBT = r - 1
r = inp(r, "Interest rate on debt", 0.095, PCT, "ILLUSTRATIVE. Set from term sheet.", key=True)
R_INT = r - 1
r = inp(r, "Debt tenor (post-COD)", 15, NUM, "Amortising term loan", unit="years")
R_TENOR = r - 1
r = inp(r, "Principal moratorium (post-COD)", 2, NUM, "Standard in BOT — principal holiday during traffic ramp-up", unit="years")
R_MORAT = r - 1
r = inp(r, "Corporate tax rate", 0.25, PCT, "India headline corporate rate (Sec 115BAA regime)")
R_TAX = r - 1
r = inp(r, "Depreciation life (SLM)", 20, NUM, "Straight line over concession", unit="years")
R_DEP = r - 1
r = inp(r, "Discount rate (WACC)", 0.11, PCT, "Blended cost of capital", key=True)
R_WACC = r - 1
r = inp(r, "Cost of equity", 0.15, PCT, "Used for equity IRR benchmarking", key=True)
R_COE = r - 1

ws.cell(r + 2, 2, "NOTE: every figure above is an illustrative placeholder. Replace with sourced "
                 "values from a concession agreement, NHAI bid document or NIP project record "
                 "before presenting this model.").font = F_NOTE

A = "Assumptions"

# ============================================================================
# SHEET 2 — CASH FLOW
# ============================================================================
cf = wb.create_sheet("Cash Flow")
cf.column_dimensions["A"].width = 3
cf.column_dimensions["B"].width = 38
cf.freeze_panes = "C6"

N_CONSTR, N_OPS = 2, 20
N_YRS = N_CONSTR + N_OPS

cf["B2"] = "CASH FLOW PROJECTION  (INR mn)"
cf["B2"].font = F_TITLE
cf["B2"].fill = FILL_DARK
for i in range(3, 3 + N_YRS):
    cf.cell(2, i).fill = FILL_DARK
cf.row_dimensions[2].height = 24

# year headers
cf.cell(4, 2, "Year").font = F_HDR
cf.cell(4, 2).fill = FILL_HDR
cf.cell(5, 2, "Phase").font = F_HDR
cf.cell(5, 2).fill = FILL_HDR
for i in range(N_YRS):
    col = 3 + i
    cf.column_dimensions[get_column_letter(col)].width = 11
    c = cf.cell(4, col, str(i + 1))
    c.font = F_HDR
    c.fill = FILL_HDR
    c.alignment = Alignment(horizontal="center")
    p = cf.cell(5, col, "Build" if i < N_CONSTR else "Operate")
    p.font = F_NOTE
    p.alignment = Alignment(horizontal="center")
    p.fill = FILL_SUB


def rowlabel(row, text, bold=False, fill=None, top=False):
    c = cf.cell(row, 2, text)
    c.font = F_LBL if bold else F_TXT
    if fill:
        c.fill = fill
    if top:
        c.border = TOPLINE
    return row


def fillrow(row, formula_fn, fmt=CUR, bold=False, fill=None, top=False):
    for i in range(N_YRS):
        col = 3 + i
        val = formula_fn(i, col)
        c = cf.cell(row, col, val)
        c.font = Font(name="Arial", size=10, bold=bold)
        c.number_format = fmt
        if fill:
            c.fill = fill
        if top:
            c.border = TOPLINE


L = get_column_letter  # shorthand

# --- traffic ---
R = 7
rowlabel(R, "OPERATIONS", bold=True, fill=FILL_SUB)
R += 1

r_traffic = R
rowlabel(R, "Traffic (AADT, vehicles/day)")
def f_traffic(i, col):
    if i < N_CONSTR:
        return 0
    ops_yr = i - N_CONSTR                      # 0-indexed operating year
    ramp = (f"{A}!$C${R_RAMP1}" if ops_yr == 0 else
            f"{A}!$C${R_RAMP2}" if ops_yr == 1 else "1")
    return f"={A}!$C${R_AADT}*(1+{A}!$C${R_TGROW})^{ops_yr}*{ramp}"
fillrow(R, f_traffic, NUM)
R += 1

r_toll = R
rowlabel(R, "Blended toll (INR/trip)")
fillrow(R, lambda i, col: 0 if i < N_CONSTR else
        f"={A}!$C${R_TOLL}*(1+{A}!$C${R_TESC})^{i - N_CONSTR}", CUR2)
R += 1

r_gross = R
rowlabel(R, "Gross toll revenue")
fillrow(R, lambda i, col: 0 if i < N_CONSTR else
        f"={L(col)}{r_traffic}*{L(col)}{r_toll}*365/1000000")
R += 1

r_leak = R
rowlabel(R, "Less: leakage / exemptions")
fillrow(R, lambda i, col: 0 if i < N_CONSTR else
        f"=-{L(col)}{r_gross}*{A}!$C${R_LEAK}")
R += 1

r_netrev = R
rowlabel(R, "Net toll revenue", bold=True, top=True)
fillrow(R, lambda i, col: f"={L(col)}{r_gross}+{L(col)}{r_leak}", CUR, bold=True, top=True)
R += 2

# --- opex ---
r_om = R
rowlabel(R, "Routine O&M")
fillrow(R, lambda i, col: 0 if i < N_CONSTR else
        f"=-{A}!$C${R_OM}*(1+{A}!$C${R_OMINF})^{i - N_CONSTR}")
R += 1

r_mm = R
rowlabel(R, "Major maintenance")
def f_mm(i, col):
    if i < N_CONSTR:
        return 0
    ops_yr = i - N_CONSTR + 1
    # charge an overlay every Nth operating year, escalated
    return (f"=IF(MOD({ops_yr},{A}!$C${R_MMCYC})=0,"
            f"-{A}!$C${R_MMCOST}*(1+{A}!$C${R_OMINF})^{ops_yr},0)")
fillrow(R, f_mm)
R += 1

r_mmra = R
rowlabel(R, "Major maintenance reserve accrual")
fillrow(R, lambda i, col: 0 if i < N_CONSTR else
        f"=-{A}!$C${R_MMCOST}*(1+{A}!$C${R_OMINF})^{i - N_CONSTR}/{A}!$C${R_MMCYC}")
R += 1

r_ebitda = R
rowlabel(R, "EBITDA", bold=True, top=True)
fillrow(R, lambda i, col: f"={L(col)}{r_netrev}+{L(col)}{r_om}+{L(col)}{r_mm}",
        CUR, bold=True, top=True)
R += 1

r_margin = R
rowlabel(R, "EBITDA margin")
fillrow(R, lambda i, col: 0 if i < N_CONSTR else
        f"=IF({L(col)}{r_netrev}=0,0,{L(col)}{r_ebitda}/{L(col)}{r_netrev})", PCT)
R += 2

# --- capex ---
rowlabel(R, "CAPITAL & FINANCING", bold=True, fill=FILL_SUB)
R += 1

r_capex = R
rowlabel(R, "Capex")
def f_capex(i, col):
    if i >= N_CONSTR:
        return 0
    share = f"{A}!$C${R_SPLIT1}" if i == 0 else f"(1-{A}!$C${R_SPLIT1})"
    return f"=-{A}!$C${R_TPC}*(1+{A}!$C${R_CONT})*{share}"
fillrow(R, f_capex)
R += 1

r_cumcap = R
rowlabel(R, "Cumulative capex")
fillrow(R, lambda i, col: (f"=-{L(col)}{r_capex}" if i == 0 else
                           f"={L(col-1)}{r_cumcap}-{L(col)}{r_capex}"))
R += 1

r_dep = R
rowlabel(R, "Depreciation (SLM)")
fillrow(R, lambda i, col: 0 if i < N_CONSTR else
        f"=-${L(3+N_CONSTR-1)}${r_cumcap}/{A}!$C${R_DEP}")
R += 1

r_principal = R + 2   # forward declaration (row assigned below)
r_debtbal = R
rowlabel(R, "Debt balance (opening)")
def f_debtbal(i, col):
    if i < N_CONSTR - 1:
        return 0
    if i == N_CONSTR - 1:
        return f"=${L(3+N_CONSTR-1)}${r_cumcap}*{A}!$C${R_DEBT}"
    return f"=MAX(0,{L(col-1)}{r_debtbal}+{L(col-1)}{r_principal})"
fillrow(R, f_debtbal)
R += 1

r_interest = R
rowlabel(R, "Interest expense")
fillrow(R, lambda i, col: 0 if i < N_CONSTR else
        f"=-{L(col)}{r_debtbal}*{A}!$C${R_INT}")
R += 1

assert r_principal == R, f"principal row mismatch {r_principal} vs {R}"
rowlabel(R, "Principal repayment")
fillrow(R, lambda i, col: 0 if i < N_CONSTR else
        f"=IF({i - N_CONSTR}<{A}!$C${R_MORAT},0,"
        f"-MIN({L(col)}{r_debtbal},${L(3+N_CONSTR-1)}${r_debtbal}/"
        f"({A}!$C${R_TENOR}-{A}!$C${R_MORAT})))")
R += 2

# --- P&L to FCF ---
rowlabel(R, "RETURNS", bold=True, fill=FILL_SUB)
R += 1

r_ebit = R
rowlabel(R, "EBIT")
fillrow(R, lambda i, col: f"={L(col)}{r_ebitda}+{L(col)}{r_dep}")
R += 1

r_pbt = R
rowlabel(R, "PBT")
fillrow(R, lambda i, col: f"={L(col)}{r_ebit}+{L(col)}{r_interest}")
R += 1

r_taxrow = R
rowlabel(R, "Tax")
fillrow(R, lambda i, col: f"=-MAX(0,{L(col)}{r_pbt})*{A}!$C${R_TAX}")
R += 1

r_pat = R
rowlabel(R, "PAT", bold=True, top=True)
fillrow(R, lambda i, col: f"={L(col)}{r_pbt}+{L(col)}{r_taxrow}", CUR, bold=True, top=True)
R += 2

r_fcff = R
rowlabel(R, "Free Cash Flow to Firm", bold=True, fill=FILL_OUT)
fillrow(R, lambda i, col:
        f"={L(col)}{r_ebitda}+{L(col)}{r_taxrow}+{L(col)}{r_capex}",
        CUR, bold=True, fill=FILL_OUT)
R += 1

r_fcfe = R
rowlabel(R, "Free Cash Flow to Equity", bold=True, fill=FILL_OUT)
fillrow(R, lambda i, col:
        f"={L(col)}{r_fcff}+{L(col)}{r_interest}+{L(col)}{r_principal}"
        + (f"+{A}!$C${R_TPC}*(1+{A}!$C${R_CONT})*{A}!$C${R_DEBT}"
           f"*{A}!$C${R_SPLIT1}" if i == 0 else
           f"+{A}!$C${R_TPC}*(1+{A}!$C${R_CONT})*{A}!$C${R_DEBT}"
           f"*(1-{A}!$C${R_SPLIT1})" if i == 1 else ""),
        CUR, bold=True, fill=FILL_OUT)
R += 1

r_disc = R
rowlabel(R, "Discount factor")
fillrow(R, lambda i, col: f"=1/(1+{A}!$C${R_WACC})^{i+1}", '0.000')
R += 1

r_pv = R
rowlabel(R, "PV of FCFF")
fillrow(R, lambda i, col: f"={L(col)}{r_fcff}*{L(col)}{r_disc}")
R += 1

r_cumpv = R
rowlabel(R, "Cumulative PV")
fillrow(R, lambda i, col: (f"={L(col)}{r_pv}" if i == 0 else
                           f"={L(col-1)}{r_cumpv}+{L(col)}{r_pv}"))
R += 2

r_dscr = R
rowlabel(R, "DSCR (smoothed CFADS / debt service)", bold=True)
fillrow(R, lambda i, col: 0 if i < N_CONSTR else
        f"=IF(({L(col)}{r_interest}+{L(col)}{r_principal})=0,\"\","
        f"({L(col)}{r_ebitda}-{L(col)}{r_mm}+{L(col)}{r_mmra}+{L(col)}{r_taxrow})/"
        f"-({L(col)}{r_interest}+{L(col)}{r_principal}))", MULT, bold=True)

LAST = L(3 + N_YRS - 1)

# ============================================================================
# SHEET 3 — RETURNS & SENSITIVITY
# ============================================================================
rt = wb.create_sheet("Returns")
rt.column_dimensions["A"].width = 3
rt.column_dimensions["B"].width = 40
rt.column_dimensions["C"].width = 16
rt.column_dimensions["D"].width = 50

rt["B2"] = "RETURNS SUMMARY"
rt["B2"].font = F_TITLE
rt["B2"].fill = FILL_DARK
for c in "CD":
    rt[f"{c}2"].fill = FILL_DARK
rt.row_dimensions[2].height = 24

CFS = "'Cash Flow'"
out = [
    ("Total project cost (incl. contingency)", f"={A}!C{R_TPC}*(1+{A}!C{R_CONT})", CUR,
     "Equity + debt funded"),
    ("Peak net revenue (final year)", f"={CFS}!{LAST}{r_netrev}", CUR, "Terminal-year toll income"),
    ("", None, None, ""),
    ("NPV @ WACC", f"=SUM({CFS}!C{r_pv}:{LAST}{r_pv})", CUR,
     "Positive = project creates value at the hurdle rate"),
    ("Project IRR (FCFF)", f"=IRR({CFS}!C{r_fcff}:{LAST}{r_fcff})", PCT2,
     "Compare against WACC"),
    ("Equity IRR (FCFE)", f"=IRR({CFS}!C{r_fcfe}:{LAST}{r_fcfe})", PCT2,
     "Compare against cost of equity"),
    ("", None, None, ""),
    ("WACC (hurdle)", f"={A}!C{R_WACC}", PCT2, "From Assumptions"),
    ("Spread over WACC", f"=IRR({CFS}!C{r_fcff}:{LAST}{r_fcff})-{A}!C{R_WACC}", PCT2,
     "Project IRR less hurdle rate"),
    ("", None, None, ""),
    ("Minimum DSCR (operating years)", f"=MIN({CFS}!{L(3+N_CONSTR)}{r_dscr}:{LAST}{r_dscr})", MULT,
     "Lenders typically covenant a floor around 1.20x-1.30x"),
    ("Average DSCR (operating years)", f"=AVERAGE({CFS}!{L(3+N_CONSTR)}{r_dscr}:{LAST}{r_dscr})", MULT,
     "Debt service headroom"),
    ("", None, None, ""),
    ("DSCR covenant floor (lender)", "=0.2*6", MULT,
     "Illustrative 1.20x floor - typical for Indian BOT term loans"),
    ("Operating years breaching 1.20x",
     f"=COUNTIF({CFS}!{L(3+N_CONSTR)}{r_dscr}:{LAST}{r_dscr},\"<1.2\")", NUM,
     "Years where CFADS fails the covenant - drives DSRA sizing"),
    ("", None, None, ""),
    ("Discounted payback (year)",
     f"=IFERROR(MATCH(TRUE,INDEX({CFS}!C{r_cumpv}:{LAST}{r_cumpv}>0,0),0),\"beyond concession\")",
     NUM, "First year cumulative PV turns positive"),
]

rr = 4
KEYROWS = {}
for label, formula, fmt, note in out:
    if formula is None:
        rr += 1
        continue
    rt.cell(rr, 2, label).font = F_KEY if "IRR" in label or "NPV" in label else F_LBL
    c = rt.cell(rr, 3, formula)
    c.font = F_LINK
    c.number_format = fmt
    c.border = BOX
    if "IRR" in label or label.startswith("NPV") or "DSCR" in label:
        c.fill = FILL_KEY
    rt.cell(rr, 4, note).font = F_NOTE
    KEYROWS[label] = rr
    rr += 1

# --- sensitivity grid: traffic CAGR vs toll escalation ---
rr += 2
rt.cell(rr, 2, "SENSITIVITY — Project IRR").font = F_HDR
for c in range(2, 9):
    rt.cell(rr, c).fill = FILL_HDR
rr += 1
rt.cell(rr, 2, "Rows: traffic CAGR   |   Columns: toll escalation").font = F_NOTE
rr += 1

SENS_TOP = rr
traffic_vals = [0.035, 0.045, 0.055, 0.065, 0.075]
toll_vals = [0.02, 0.03, 0.04, 0.05, 0.06]

rt.cell(SENS_TOP, 2, "Traffic \\ Toll").font = F_LBL
rt.cell(SENS_TOP, 2).fill = FILL_SUB
for j, tv in enumerate(toll_vals):
    c = rt.cell(SENS_TOP, 3 + j, tv)
    c.font = F_LBL
    c.number_format = PCT
    c.fill = FILL_SUB
    c.alignment = Alignment(horizontal="center")

for i, gv in enumerate(traffic_vals):
    c = rt.cell(SENS_TOP + 1 + i, 2, gv)
    c.font = F_LBL
    c.number_format = PCT
    c.fill = FILL_SUB
    for j in range(len(toll_vals)):
        cc = rt.cell(SENS_TOP + 1 + i, 3 + j)
        cc.number_format = PCT2
        cc.border = BOX

rt.cell(SENS_TOP + len(traffic_vals) + 2, 2,
        "Sensitivity grid is populated by the Monte Carlo script "
        "(run_monte_carlo.py) — values written as static results.").font = F_NOTE

wb.save("Toll_Road_Project_Finance_Model.xlsx")
print("Workbook written.")
print(f"  Cash Flow rows -> FCFF={r_fcff}, FCFE={r_fcfe}, DSCR={r_dscr}, cumPV={r_cumpv}")
print(f"  Sensitivity grid anchored at Returns!B{SENS_TOP}")
