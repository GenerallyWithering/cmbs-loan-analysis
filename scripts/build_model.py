"""Builds model/700_6th_Street_CMBS_Loan_Analysis.xlsx.

Every number typed into the model is an input in BLUE with its source (document and page). Every calculated number is
an Excel formula in BLACK that points at those inputs (no typed numbers inside formulas). Run:

    .venv/bin/python scripts/build_model.py
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName

OUT = Path(__file__).resolve().parents[1] / "model" / "700_6th_Street_CMBS_Loan_Analysis.xlsx"

DOC = "Benchmark 2026-B44 Mortgage Trust, final prospectus (Form 424B2, filed Nov 10, 2026)"
A3 = "Annex A-3 (700 6th Street), p. "
ASSUME = "My assumption (not from the prospectus)"

BLUE = Font(name="Calibri", size=11, color="0000FF")
BLACK = Font(name="Calibri", size=11, color="000000")
BOLD = Font(name="Calibri", size=11, bold=True)
TITLE = Font(name="Calibri", size=14, bold=True)
GREY = Font(name="Calibri", size=10, color="595959", italic=True)
HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
SECTION_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

USD = '"$"#,##0'
USD2 = '"$"#,##0.00'
PCT1 = "0.0%"
PCT2 = "0.00%"
MULT = '0.00"x"'
NUM = "#,##0"
DATE = "mmm d, yyyy"

wb = Workbook()


def name(ws, cell: str, nm: str) -> None:
    """A workbook-level name for one cell, so formulas read like =UW_NCF/AnnualDS_Mine."""
    col = "".join(ch for ch in cell if ch.isalpha())
    row = "".join(ch for ch in cell if ch.isdigit())
    ref = f"'{ws.title}'!${col}${row}"
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def header(ws, row: int, labels: list, widths: list | None = None) -> None:
    for i, text in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=text)
        c.font = BOLD
        c.fill = HEAD_FILL
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + i)].width = w


def section(ws, row: int, text: str, span: int = 5) -> None:
    for col in range(1, span + 1):
        ws.cell(row=row, column=col).fill = SECTION_FILL
    ws.cell(row=row, column=1, value=text).font = BOLD


# ---------------------------------------------------------------------------------------------------------------- Inputs
ws = wb.active
ws.title = "Inputs"
ws["A1"] = "700 6th Street — CMBS loan inputs (typed in blue, each with its source)"
ws["A1"].font = TITLE
ws["A2"] = f"Source document: {DOC}. Page numbers are the printed page labels in the prospectus."
ws["A2"].font = GREY
header(ws, 4, ["Input", "Value", "Unit", "Source (document, page)", "Note"], [46, 18, 10, 52, 60])

INPUTS = [
    ("Loan terms", None),
    ("Loan amount (cut-off date balance)", 54197000, USD, "Loan_Amount", A3 + "A-3-28", "Original balance is the same; single note"),
    ("Interest rate (fixed)", 0.0738, PCT2, "Rate", A3 + "A-3-28", ""),
    ("Day count: days per year used", 365, NUM, "Days_Year", A3 + "A-3-28 (\"accrues interest ... on an Actual/360 basis\")", "Actual/360 = a 365-day year of interest on a 360-day basis"),
    ("Day count: basis", 360, NUM, "Day_Basis", A3 + "A-3-28", ""),
    ("Payments per year", 12, NUM, "Pmts_Year", A3 + "A-3-28 (monthly payment dates, first Nov 6, 2026)", ""),
    ("Original term (months)", 120, NUM, "Term_Months", A3 + "A-3-28", ""),
    ("Interest-only period (months)", 120, NUM, "IO_Months", A3 + "A-3-28", "Interest-only for the full term"),
    ("Amortization after IO (months)", 0, NUM, "Amort_Months", A3 + "A-3-28 (\"Original Amortization: NAP\")", "0 = none (balloon = full loan amount)"),
    ("Note (origination) date", "2026-10-02", DATE, "Note_Date", A3 + "A-3-28", ""),
    ("Maturity date", "2036-10-06", DATE, "Maturity_Date", A3 + "A-3-28", ""),
    ("Property", None),
    ("Net rentable area (SF)", 304947, NUM, "NRA", A3 + "A-3-28", "Class A office, built 2009, Washington DC (East End)"),
    ("Appraised value, as-is (Sep 8, 2026)", 86000000, USD, "Appraised_Value", A3 + "A-3-28 and A-3-32", ""),
    ("Appraisal cap rate", 0.11, PCT2, "Appraisal_Cap", A3 + "A-3-32", ""),
    ("Purchase price (acquisition loan)", 79500000, USD, "Purchase_Price", A3 + "A-3-28 (Sources and Uses)", ""),
    ("Occupancy (leased, Oct 1, 2026)", 0.761, PCT1, "Occupancy", A3 + "A-3-28", ""),
    ("Underwritten economic occupancy", 0.794, PCT1, "UW_Econ_Occ", A3 + "A-3-28 and A-3-31", ""),
    ("Underwritten cash flow", None),
    ("Underwritten NOI", 10442065, USD, "UW_NOI", A3 + "A-3-28 and A-3-31", "Includes $322,786 of contractual rent steps (A-3-28 note 3)"),
    ("Underwritten NCF", 9847208, USD, "UW_NCF", A3 + "A-3-28 and A-3-31", ""),
    ("UW replacement reserves", 60989, USD, "UW_RepRes", A3 + "A-3-31", ""),
    ("UW normalized TI/LC", 533868, USD, "UW_TILC", A3 + "A-3-31", ""),
    ("Historical NOI", None),
    ("NOI 2023", 13170098, USD, "NOI_2023", A3 + "A-3-31", "Occupancy 99.5%"),
    ("NOI 2024", 12388800, USD, "NOI_2024", A3 + "A-3-31", "Occupancy 85.7%"),
    ("NOI 2025", 9245340, USD, "NOI_2025", A3 + "A-3-31", "Free rent to Monumental + two tenants left (A-3-31 note 2)"),
    ("NOI TTM Aug 31, 2026", 9460195, USD, "NOI_TTM", A3 + "A-3-31", ""),
    ("Tenants (underwritten rent roll, Oct 1, 2026)", None),
    ("Eversheds Sutherland: SF leased", 146254, NUM, "EV_SF", A3 + "A-3-29", "48.0% of NRA; law firm; tenant since 2010"),
    ("Eversheds Sutherland: UW base rent", 11784591, USD, "EV_Rent", A3 + "A-3-29", "68.5% of UW base rent"),
    ("Eversheds Sutherland: office SF", 143160, NUM, "EV_Office_SF", A3 + "A-3-29 (note 2)", ""),
    ("Eversheds Sutherland: office rent per SF", 81.58, USD2, "EV_Office_PSF", A3 + "A-3-29 (note 2)", ""),
    ("Eversheds Sutherland: lease expiration", "2033-04-30", DATE, "EV_Exp", A3 + "A-3-29", "One 5-year renewal option, no termination option"),
    ("Monumental Sports & Ent.: SF leased", 85669, NUM, "MSE_SF", A3 + "A-3-29", "28.1% of NRA"),
    ("Monumental Sports & Ent.: UW base rent", 5330894, USD, "MSE_Rent", A3 + "A-3-29", "31.0% of UW base rent"),
    ("Monumental Sports & Ent.: lease expiration", "2036-09-30", DATE, "MSE_Exp", A3 + "A-3-29", "Two 5-year renewal options"),
    ("Monumental Sports & Ent.: termination option date", "2031-09-30", DATE, "MSE_Term", A3 + "A-3-29 (note 4)", "Needs 18 months' notice and a fee"),
    ("Antenna leases expiring 2029: UW base rent", 15934, USD, "Ant_2029", A3 + "A-3-30 (note 3)", ""),
    ("Antenna lease expiring 2035: UW base rent", 30600, USD, "Ant_2035", A3 + "A-3-30 (note 4)", ""),
    ("Management office: UW base rent (no expiry given)", 53491, USD, "Mgmt_Rent", A3 + "A-3-30 (note 5)", ""),
    ("Total UW base rent (occupied)", 17215510, USD, "Total_Rent", A3 + "A-3-29", "Includes $322,786 of rent steps"),
    ("Vacant SF", 73024, NUM, "Vacant_SF", A3 + "A-3-29", "23.9% of NRA"),
    ("Market and structure", None),
    ("East End submarket vacancy (Oct 21, 2026)", 0.227, PCT1, "Mkt_Vacancy", A3 + "A-3-32", "Third-party market report"),
    ("East End market asking rent per SF", 60.53, USD2, "Mkt_Rent", A3 + "A-3-32", ""),
    ("Cash sweep trigger: DSCR below", 1.50, MULT, "Trigger_DSCR", A3 + "A-3-33", "Cures at 1.55x for two quarters"),
    ("Specified tenant cash sweep cap per SF", 100, USD, "ST_Cap_PSF", A3 + "A-3-34", "Cap = $100 x Eversheds' SF"),
    ("TI/LC reserve: monthly deposit", 44489, USD, "TILC_Monthly", A3 + "A-3-28", ""),
    ("Free rent reserve at closing (Monumental)", 3845865, USD, "FreeRent_Res", A3 + "A-3-28", ""),
    ("Reported by the prospectus (to check my numbers)", None),
    ("Reported annual debt service (IO)", 4055290.56, USD2, "Rep_ADS", DOC + ", Annex A-1, p. A-1-9 (loan 3 row)", ""),
    ("Reported monthly debt service (IO)", 337940.88, USD2, "Rep_MDS", DOC + ", Annex A-1, p. A-1-9 (loan 3 row)", ""),
    ("Reported UW NCF DSCR", 2.43, MULT, "Rep_DSCR", A3 + "A-3-28", ""),
    ("Reported cut-off date LTV", 0.63, PCT1, "Rep_LTV", A3 + "A-3-28", ""),
    ("Reported maturity date LTV", 0.63, PCT1, "Rep_LTV_Mat", A3 + "A-3-28", ""),
    ("Reported UW NOI debt yield", 0.193, PCT1, "Rep_DY", A3 + "A-3-28", ""),
    ("My assumptions for the stress tests", None),
    ("NOI decline, case 1", 0.10, PCT1, "Stress_1", ASSUME, ""),
    ("NOI decline, case 2", 0.20, PCT1, "Stress_2", ASSUME, ""),
    ("Refi rate, case A", 0.07, PCT2, "Refi_Rate_A", ASSUME, "New loan rate in Oct 2036"),
    ("Refi rate, case B", 0.08, PCT2, "Refi_Rate_B", ASSUME, ""),
    ("Refi rate, case C", 0.09, PCT2, "Refi_Rate_C", ASSUME, ""),
    ("Refi amortization (months)", 360, NUM, "Refi_Amort", ASSUME, "30-year schedule, typical for a new conduit loan"),
    ("Refi lender minimum DSCR", 1.25, MULT, "Min_DSCR", ASSUME, ""),
    ("Refi lender minimum debt yield", 0.10, PCT1, "Min_DY", ASSUME, "Office lenders often want 10%+"),
    ("Refi lender maximum LTV", 0.65, PCT1, "Max_LTV", ASSUME, ""),
    ("Exit cap rate (value at maturity)", 0.10, PCT2, "Exit_Cap", ASSUME, "Appraisal used 11.0% (A-3-32)"),
]

row = 5
for item in INPUTS:
    if item[1] is None:
        section(ws, row, item[0])
        row += 1
        continue
    label, value, fmt, nm, src, note = item
    ws.cell(row=row, column=1, value=label).font = BLACK
    if isinstance(value, str) and value[:2] == "20":
        from datetime import date
        y, m, d = (int(x) for x in value.split("-"))
        value = date(y, m, d)
    c = ws.cell(row=row, column=2, value=value)
    c.font = BLUE
    c.number_format = fmt
    ws.cell(row=row, column=3, value={USD: "$", USD2: "$", PCT1: "%", PCT2: "%", MULT: "x", NUM: "#", DATE: "date"}[fmt])
    ws.cell(row=row, column=4, value=src).font = GREY
    ws.cell(row=row, column=5, value=note).font = GREY
    name(ws, f"B{row}", nm)
    row += 1
ws.freeze_panes = "A5"
ws.cell(row=row + 1, column=1, value="Sponsor: FitzWalter Capital Partners (AIV) II, LP (A-3-28, A-3-32). Lender: CREFI. "
        "Lockbox: hard, springing cash management (A-3-28, A-3-33).").font = GREY

# --------------------------------------------------------------------------------------------------------------- Metrics
m = wb.create_sheet("Metrics")
m["A1"] = "Key credit metrics — my calculation vs. the prospectus"
m["A1"].font = TITLE
m["A2"] = "All black cells are formulas that use the blue inputs on the Inputs sheet."
m["A2"].font = GREY
header(m, 4, ["Metric", "My calculation", "Reported", "Difference", "Why they differ"], [42, 18, 16, 16, 70])

METRICS = [
    ("Annual debt service (IO, Actual/360)",
     "=IF(Amort_Months=0,Loan_Amount*Rate*Days_Year/Day_Basis,PMT(Rate/Pmts_Year,Amort_Months,-Loan_Amount)*Pmts_Year)",
     "=Rep_ADS", USD2, "AnnualDS_Mine",
     "The prospectus multiplies a monthly payment rounded to the cent by 12; I use the unrounded annual figure."),
    ("Monthly debt service (average)", "=AnnualDS_Mine/Pmts_Year", "=Rep_MDS", USD2, "MonthlyDS",
     "Average month. Actual months differ (28-31 days of interest)."),
    ("UW NCF DSCR = NCF / annual debt service", "=UW_NCF/AnnualDS_Mine", "=Rep_DSCR", MULT, "DSCR_NCF",
     "Reported figure is rounded to two decimals."),
    ("UW NOI DSCR = NOI / annual debt service", "=UW_NOI/AnnualDS_Mine", None, MULT, "DSCR_NOI",
     "Not shown on the summary page; NOI is before reserves for replacements and TI/LC."),
    ("LTV at origination = loan / appraised value", "=Loan_Amount/Appraised_Value", "=Rep_LTV", PCT1, "LTV_Orig",
     "Reported figure is rounded to one decimal."),
    ("Balloon at maturity", "=IF(IO_Months>=Term_Months,Loan_Amount,Loan_Amount)", None, USD, "Balloon",
     "Interest-only for the full term, so nothing is paid down: the whole loan is due at maturity."),
    ("LTV at maturity = balloon / appraised value", "=Balloon/Appraised_Value", "=Rep_LTV_Mat", PCT1, "LTV_Mat",
     "Same as at origination because there is no amortization (uses today's appraisal, not a future value)."),
    ("Debt yield = NOI / loan amount", "=UW_NOI/Loan_Amount", "=Rep_DY", PCT1, "DY_NOI",
     "Reported figure is rounded to one decimal."),
    ("NCF debt yield = NCF / loan amount", "=UW_NCF/Loan_Amount", None, PCT1, "DY_NCF", ""),
    ("Loan per SF", "=Loan_Amount/NRA", None, USD, "Loan_PSF", "Reported as $178 (A-3-28)."),
    ("Loan-to-cost = loan / purchase price", "=Loan_Amount/Purchase_Price", None, PCT1, "LTC",
     "Sponsor's cash equity cushion below the loan."),
    ("Going-in cap rate on purchase = UW NOI / price", "=UW_NOI/Purchase_Price", None, PCT2, "Cap_Purchase",
     "High cap rate = the buyer priced in real risk (weak office market, rollover)."),
    ("Value implied by TTM NOI at the appraisal cap rate", "=NOI_TTM/Appraisal_Cap", None, USD, "Implied_Value",
     "Close to the $86.0M appraisal: the appraiser capitalized roughly current (TTM) NOI."),
    ("Cushion before DSCR hits the 1.50x cash sweep", "=1-(Trigger_DSCR*AnnualDS_Mine+(UW_NOI-UW_NCF))/UW_NOI", None,
     PCT1, "Cushion_Trigger", "How far NOI can fall before excess cash gets trapped."),
    ("Cushion before DSCR hits 1.00x (break-even)", "=1-(AnnualDS_Mine+(UW_NOI-UW_NCF))/UW_NOI", None, PCT1,
     "Cushion_BE", "How far NOI can fall before cash flow no longer covers debt service."),
]

r = 5
for label, mine, rep, fmt, nm, why in METRICS:
    m.cell(row=r, column=1, value=label)
    c = m.cell(row=r, column=2, value=mine)
    c.number_format = fmt
    c.font = BLACK
    name(m, f"B{r}", nm)
    if rep:
        c2 = m.cell(row=r, column=3, value=rep)
        c2.number_format = fmt
        d = m.cell(row=r, column=4, value=f"=B{r}-C{r}")
        d.number_format = fmt if fmt in (USD, USD2) else ("0.0000" if fmt == MULT else "0.00%")
    else:
        m.cell(row=r, column=3, value="n/a").font = GREY
    m.cell(row=r, column=5, value=why).font = GREY
    m.cell(row=r, column=5).alignment = Alignment(wrap_text=True)
    r += 1
m.freeze_panes = "A5"

# ------------------------------------------------------------------------------------------------------------ Historical
h = wb.create_sheet("Historical")
h["A1"] = "NOI history — the trend behind the underwriting"
h["A1"].font = TITLE
header(h, 3, ["Period", "NOI", "Change vs 2023", "NCF DSCR on today's debt service", "Note"], [22, 16, 16, 20, 60])
HIST = [("2023", "=NOI_2023", "Fully leased year (99.5%)"), ("2024", "=NOI_2024", ""),
        ("2025", "=NOI_2025", "Free rent + two tenants left"), ("TTM Aug 2026", "=NOI_TTM", ""),
        ("Underwritten", "=UW_NOI", "Uses UW NCF for the DSCR (after reserves)")]
for i, (p, f, note) in enumerate(HIST):
    rr = 4 + i
    h.cell(row=rr, column=1, value=p)
    h.cell(row=rr, column=2, value=f).number_format = USD
    h.cell(row=rr, column=3, value=f"=B{rr}/$B$4-1").number_format = PCT1
    dscr = "=UW_NCF/AnnualDS_Mine" if p == "Underwritten" else f"=B{rr}/AnnualDS_Mine"
    h.cell(row=rr, column=4, value=dscr).number_format = MULT
    h.cell(row=rr, column=5, value=note).font = GREY
h["A10"] = ("The prospectus shows NCF DSCRs of 3.25x, 3.05x, 2.28x and 2.33x for 2023-TTM (A-3-31); historical NCF there "
            "equals NOI because no reserves were taken out.")
h["A10"].font = GREY

# --------------------------------------------------------------------------------------------------------------- Stress
s = wb.create_sheet("Stress")
s["A1"] = "Stress tests"
s["A1"].font = TITLE
s["A2"] = ("Reserves (NOI minus NCF) are held constant in every case: a lower NOI does not make capital spending smaller. "
           "Tenant cases change base rent only (they ignore lost expense reimbursements and new leasing costs).")
s["A2"].font = GREY
s["A2"].alignment = Alignment(wrap_text=True)
s.row_dimensions[2].height = 30

header(s, 4, ["A. NOI decline", "Base (UW)", "Case 1", "Case 2", ""], [44, 18, 18, 18, 18])
s["A5"] = "NOI decline"
s["B5"] = 0
s["B5"].font = BLACK
s["C5"] = "=Stress_1"
s["D5"] = "=Stress_2"
for col in "BCD":
    s[f"{col}5"].number_format = PCT1
rowsA = [("NOI", "=UW_NOI*(1-{c}5)", USD), ("Reserves (NOI - NCF)", "=UW_NOI-UW_NCF", USD),
         ("NCF", "={c}6-{c}7", USD), ("Annual debt service", "=AnnualDS_Mine", USD),
         ("DSCR (NCF)", "={c}8/{c}9", MULT), ("Debt yield (NOI)", "={c}6/Loan_Amount", PCT1),
         ("Above the 1.50x cash sweep trigger?", '=IF({c}10>=Trigger_DSCR,"Yes","No - cash trapped")', None)]
for i, (lab, f, fmt) in enumerate(rowsA):
    rr = 6 + i
    s.cell(row=rr, column=1, value=lab)
    for col in "BCD":
        c = s[f"{col}{rr}"]
        c.value = f.format(c=col)
        if fmt:
            c.number_format = fmt
s["A5"].font = BOLD

header(s, 15, ["B. Tenant events", "Rent change", "NOI after", "DSCR (NCF)", "Debt yield (NOI)"])
TEN = [
    ("Eversheds renews at market rent (2033)", "=(Mkt_Rent-EV_Office_PSF)*EV_Office_SF"),
    ("Eversheds leaves, space stays empty", "=-EV_Rent"),
    ("Monumental terminates in Sep 2031", "=-MSE_Rent"),
    ("Both big tenants gone", "=-EV_Rent-MSE_Rent"),
]
for i, (lab, f) in enumerate(TEN):
    rr = 16 + i
    s.cell(row=rr, column=1, value=lab)
    s.cell(row=rr, column=2, value=f).number_format = USD
    s.cell(row=rr, column=3, value=f"=UW_NOI+B{rr}").number_format = USD
    s.cell(row=rr, column=4, value=f"=(C{rr}-(UW_NOI-UW_NCF))/AnnualDS_Mine").number_format = MULT
    s.cell(row=rr, column=5, value=f"=C{rr}/Loan_Amount").number_format = PCT1

header(s, 22, ["C. Rent rolling before the loan matures", "UW base rent", "Share of UW base rent", "", ""])
ROLL = [("Eversheds (Apr 2033)", "=IF(EV_Exp<=Maturity_Date,EV_Rent,0)"),
        ("Monumental (Sep 2036; can leave Sep 2031)", "=IF(MSE_Exp<=Maturity_Date,MSE_Rent,0)"),
        ("Antennas (2029 and 2035)", "=Ant_2029+Ant_2035")]
for i, (lab, f) in enumerate(ROLL):
    rr = 23 + i
    s.cell(row=rr, column=1, value=lab)
    s.cell(row=rr, column=2, value=f).number_format = USD
    s.cell(row=rr, column=3, value=f"=B{rr}/Total_Rent").number_format = PCT1
s["A26"] = "Total expiring before maturity"
s["A26"].font = BOLD
s["B26"] = "=SUM(B23:B25)"
s["B26"].number_format = USD
s["C26"] = "=B26/Total_Rent"
s["C26"].number_format = PCT1
name(s, "C26", "Rollover_Share")

header(s, 28, ["D. Refinance at maturity (Oct 2036): biggest new loan", "Rate case A", "Rate case B", "Rate case C",
               "Refinances at case C?"])
s["A29"] = "Refi rate"
s["B29"], s["C29"], s["D29"] = "=Refi_Rate_A", "=Refi_Rate_B", "=Refi_Rate_C"
s["A30"] = "Loan constant (annual payment per $1, 30-yr amortization)"
for col in "BCD":
    s[f"{col}29"].number_format = PCT2
    s[f"{col}30"] = f"=PMT({col}29/Pmts_Year,Refi_Amort,-1)*Pmts_Year"
    s[f"{col}30"].number_format = PCT2
SCEN = [("NOI at maturity = UW NOI", "=UW_NOI"),
        ("NOI down by case 1", "=UW_NOI*(1-Stress_1)"),
        ("NOI down by case 2", "=UW_NOI*(1-Stress_2)"),
        ("Eversheds renewed at market rent", "=UW_NOI+B16")]
s["A31"] = "Max new loan = lowest of debt-yield, DSCR and LTV limits; balloon to repay:"
s["A31"].font = GREY
s["E31"] = "=Balloon"
s["E31"].number_format = USD
for i, (lab, f) in enumerate(SCEN):
    rr = 32 + i
    s.cell(row=rr, column=1, value=lab)
    s.cell(row=rr, column=6, value=f).number_format = USD    # the scenario's NOI (column F)
    for col in "BCD":
        s[f"{col}{rr}"] = (f"=MIN(F{rr}/Min_DY,(F{rr}-(UW_NOI-UW_NCF))/(Min_DSCR*{col}$30),"
                           f"F{rr}/Exit_Cap*Max_LTV)")
        s[f"{col}{rr}"].number_format = USD
    s[f"E{rr}"] = f'=IF(D{rr}>=Balloon,"Yes ("&TEXT(D{rr}/Balloon,"0.00")&"x the balloon)","No - short by $"&TEXT(Balloon-D{rr},"#,##0"))'
s["F31"] = "NOI used"
s["F31"].font = BOLD
s.column_dimensions["F"].width = 16
s["A37"] = "NOI needed to refinance the balloon at case C (the highest of the three limits)"
s["B37"] = "=MAX(Balloon*Min_DY,Balloon*Min_DSCR*D30+(UW_NOI-UW_NCF),Balloon*Exit_Cap/Max_LTV)"
s["B37"].number_format = USD
s["C37"] = "=B37/UW_NOI-1"
s["C37"].number_format = PCT1
s["D37"] = "vs UW NOI"
s["D37"].font = GREY

header(s, 39, ["E. Which lender limit binds (NOI = UW NOI)", "Rate case A", "Rate case B", "Rate case C", ""])
LIMITS = [("Debt-yield limit = NOI / min debt yield", "=UW_NOI/Min_DY"),
          ("DSCR limit = NCF / (min DSCR x loan constant)", "=UW_NCF/(Min_DSCR*{c}$30)"),
          ("LTV limit = (NOI / exit cap) x max LTV", "=UW_NOI/Exit_Cap*Max_LTV")]
for i, (lab, f) in enumerate(LIMITS):
    rr = 40 + i
    s.cell(row=rr, column=1, value=lab)
    for col in "BCD":
        s[f"{col}{rr}"] = f.format(c=col)
        s[f"{col}{rr}"].number_format = USD
s["A43"] = "Refi rate at which the DSCR limit drops to the balloon"
s["B43"] = "=RATE(Refi_Amort,-(UW_NCF/(Min_DSCR*Pmts_Year)),Balloon)*Pmts_Year"
s["B43"].number_format = PCT2
s["C43"] = "(with UW NCF; above this rate, the DSCR limit is too small to repay the balloon)"
s["C43"].font = GREY
s.freeze_panes = "A4"

# ---------------------------------------------------------------------------------------------------------------- Notes
n = wb.create_sheet("Read me")
n.column_dimensions["A"].width = 120
NOTES = [
    "700 6th Street (Washington, DC) — loan 3 of Benchmark 2026-B44, a $54.2M, 10-year, interest-only office loan at 7.38%.",
    "",
    "Colours: BLUE = a typed input (each one has its source on the Inputs sheet). BLACK = a formula. No typed numbers inside formulas.",
    "Sheets: Inputs → Metrics (my numbers vs the prospectus) → Historical (NOI trend) → Stress (NOI cuts, tenant events,",
    "rollover before maturity, refinance test).",
    f"Source: {DOC}. SEC EDGAR accession 0001539497-26-002925.",
    "Assumptions that are mine (stress sizes, refinance rates and lender limits, exit cap rate) are labelled as such on Inputs.",
    "Everything else comes from the prospectus, with the page.",
]
for i, t in enumerate(NOTES, start=1):
    n.cell(row=i, column=1, value=t)
wb.move_sheet("Read me", offset=-4)

OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)
print("saved", OUT)
