"""Generate the BI Migration Pre-sales Estimator (ballpark, 3 tabs).

Run:  python build_bi_presales.py [output.xlsx]

A simple companion to BI_Migration_Estimator_v2.xlsx. Same baseline numbers,
far fewer inputs: Estimate (inputs + results), Plan (sprint grid like the
original PoP sheet) and Reference (numbers behind it).
"""
import sys
import datetime as dt

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1] if len(sys.argv) > 1 else "BI_Migration_Presales_Estimator.xlsx"

FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_TITLE = Font(name=FONT, size=16, bold=True, color="1F3864")
F_SUB = Font(name=FONT, size=10, italic=True, color="595959")
F_SECTION = Font(name=FONT, size=12, bold=True, color="1F3864")
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_KPI = Font(name=FONT, size=14, bold=True, color="1F3864")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_GRID = PatternFill("solid", fgColor="FFFBE6")
FILL_HEAD = PatternFill("solid", fgColor="1F3864")
FILL_SECTION = PatternFill("solid", fgColor="D9E1F2")
FILL_CALC = PatternFill("solid", fgColor="F2F2F2")
FILL_GOOD = PatternFill("solid", fgColor="C6EFCE")
FILL_BAD = PatternFill("solid", fgColor="FFC7CE")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

FMT_H = '#,##0;(#,##0);"-"'
FMT_H1 = '#,##0.0;(#,##0.0);"-"'
FMT_USD = '$#,##0;($#,##0);"-"'
FMT_USD2 = '$#,##0.00;($#,##0.00);"-"'
FMT_PCT0 = '0%;(0%);"-"'
FMT_PCT = '0.0%;(0.0%);"-"'
FMT_X = '0.00"x"'
FMT_DATE = "dd-mmm-yyyy"
FMT_FTE = '0.0;-0.0;""'

NWEEKS = 52
NPAIRS = 5
PHASES = [  # name, colour
    ("Plan (assessment & design)", "FFF59D"),
    ("Data model & report dev", "BDD7EE"),
    ("Testing & validation (SIT)", "F8CBAD"),
    ("UAT support", "F4B183"),
    ("Prod deployment", "C6E0B4"),
    ("Hypercare / parallel run", "A9D08E"),
    ("Project management", "D9D9D9"),
]
NPH = len(PHASES)
ROLES = [("BI Architect", "Architect II | G11 | India"), ("Lead BI Engineer", "Lead II | G11 | India"),
         ("Sr BI Engineer", "Sr. Engineer II | G10 | India"), ("BI Engineer", "Engineer II | G8 | India"),
         ("QA Engineer", "Sr. Engineer I | G9 | India"), ("PM", "MSE I | G12 | India")]
NROLES = len(ROLES)
MIX = [  # phase x role (Architect, Lead, Sr, Eng, QA, PM)
    [.50, .50, 0, 0, 0, 0],
    [.10, .15, .35, .40, 0, 0],
    [0, 0, 0, 0, 1, 0],
    [0, .30, .70, 0, 0, 0],
    [0, 0, 1, 0, 0, 0],
    [0, .50, .50, 0, 0, 0],
    [0, 0, 0, 0, 0, 1],
]

wb = Workbook()
wb.remove(wb.active)


def name(nm, ws, ref):
    def ab(a):
        col = "".join(c for c in a if c.isalpha())
        row = "".join(c for c in a if c.isdigit())
        return f"${col}${row}"
    txt = f"'{ws.title}'!" + ":".join(ab(p) for p in ref.split(":"))
    wb.defined_names[nm] = DefinedName(nm, attr_text=txt)


def put(ws, ref, value, font=F_BASE, fmt=None, fill=None, align=None, border=BOX):
    c = ws[ref]
    c.value = value
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    if border:
        c.border = border
    return c


def inp(ws, ref, value, fmt=None, note=None, fill=FILL_INPUT):
    c = put(ws, ref, value, font=F_INPUT, fmt=fmt, fill=fill)
    c.protection = Protection(locked=False)
    if note:
        c.comment = Comment(note, "BI Estimator")
    return c


def head(ws, row, col0, labels, height=32):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col0 + i, value=lab)
        c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BOX
    ws.row_dimensions[row].height = height


def section(ws, ref, text, span):
    put(ws, ref, text, font=F_SECTION, fill=FILL_SECTION, border=None)
    for i in range(1, span):
        ws.cell(row=ws[ref].row, column=ws[ref].column + i).fill = FILL_SECTION


def dv_list(ws, rng, src, free=False):
    dv = DataValidation(type="list", formula1=src, allow_blank=True)
    dv.showErrorMessage = not free
    ws.add_data_validation(dv)
    dv.add(rng)


def dv_num(ws, rng, lo, hi):
    dv = DataValidation(type="decimal", operator="between", formula1=str(lo), formula2=str(hi), allow_blank=True)
    dv.showErrorMessage = True
    dv.error = f"Enter a value between {lo} and {hi}."
    ws.add_data_validation(dv)
    dv.add(rng)


def interp(x, X, Y):
    k = f"MATCH({x},{X},1)"
    return (f"IF({x}<=INDEX({X},1),INDEX({Y},1),IF({x}>=INDEX({X},5),INDEX({Y},5)+({x}-INDEX({X},5))*(INDEX({Y},5)-INDEX({Y},4))/(INDEX({X},5)-INDEX({X},4)),"
            f"INDEX({Y},{k})+({x}-INDEX({X},{k}))*(INDEX({Y},{k}+1)-INDEX({Y},{k}))/(INDEX({X},{k}+1)-INDEX({X},{k}))))")


ws_est = wb.create_sheet("Estimate")
ws_plan = wb.create_sheet("Plan")
ws_ref = wb.create_sheet("Reference")

# =====================================================================
# Reference
# =====================================================================
ws = ws_ref
ws.sheet_view.showGridLines = False
put(ws, "A1", "Reference", font=F_TITLE, border=None)
put(ws, "A2", "Numbers behind the ballpark. Same baseline as the full v2 estimator. Change only with BI-team sign-off.", font=F_SUB, border=None)
for c, w in {"A": 44, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 3, "I": 22, "J": 26, "K": 10, "L": 3,
             "M": 22, "N": 20, "O": 9, "P": 30}.items():
    ws.column_dimensions[c].width = w
rr = [4]


def rline(label, value, nm, fmt=None, note=None):
    r = rr[0]
    put(ws, f"A{r}", label)
    inp(ws, f"B{r}", value, fmt=fmt, note=note)
    name(nm, ws, f"B{r}")
    rr[0] += 1


def rsec(t, span=2):
    rr[0] += 1
    section(ws, f"A{rr[0]}", t, span)
    rr[0] += 1


rsec("General")
rline("Productive hours per person-week", 40, "HrsWk", FMT_H)
rline("Ballpark range - low side", -0.10, "RangeLow", FMT_PCT0, "Expected price x (1 + this)")
rline("Ballpark range - high side", 0.25, "RangeHigh", FMT_PCT0, "Expected price x (1 + this)")

rsec("Report effort (hours per page)", 5)
head(ws, rr[0], 1, ["Complexity profile", "% Simple", "% Medium", "% Complex", "% Very complex"])
rr[0] += 1
put(ws, f"A{rr[0]}", "Hours per report page")
r_hpp = rr[0]
for j, v in enumerate([3.625, 6.625, 9.125, 12.625]):
    inp(ws, f"{L(2 + j)}{r_hpp}", v, fmt="0.000")
name("HrsPage", ws, f"B{r_hpp}:E{r_hpp}")
rr[0] += 1
r_pre = rr[0]
PRESETS = [("Mostly simple", [.60, .30, .10, 0]), ("Balanced", [.40, .40, .15, .05]), ("Mostly complex", [.20, .40, .30, .10])]
for i, (nm_, mix) in enumerate(PRESETS):
    put(ws, f"A{r_pre + i}", nm_)
    for j, v in enumerate(mix):
        inp(ws, f"{L(2 + j)}{r_pre + i}", v, fmt=FMT_PCT0)
put(ws, f"A{r_pre + 3}", "Custom")
for j in range(4):
    put(ws, f"{L(2 + j)}{r_pre + 3}", "enter on Estimate", font=F_SUB)
name("ProfileNames", ws, f"A{r_pre}:A{r_pre + 3}")
name("ProfileMix", ws, f"B{r_pre}:E{r_pre + 2}")
rr[0] = r_pre + 4
rline("Average learning gain (0/10/15/20% by quarter of sprints)", 0.1125, "LearnAvg", FMT_PCT)
rline("Data platform change uplift - reports", 0.05, "DPRep", FMT_PCT0, "Custom SQL re-pointing; ~25% uplift on ~20% of reports")
rline("Data platform change uplift - semantic models", 0.15, "DPModel", FMT_PCT0)

rsec("Semantic models (hours per target model)", 5)
head(ws, rr[0], 1, ["Model complexity", "Simple", "Medium", "Complex", "Very complex"])
rr[0] += 1
put(ws, f"A{rr[0]}", "Build hours per model")
for j, v in enumerate([80, 120, 200, 320]):
    inp(ws, f"{L(2 + j)}{rr[0]}", v, fmt=FMT_H)
name("ModelHrs", ws, f"B{rr[0]}:E{rr[0]}")
put(ws, f"A{rr[0] + 1}", "Level names")
for j, v in enumerate(["Simple", "Medium", "Complex", "Very complex"]):
    put(ws, f"{L(2 + j)}{rr[0] + 1}", v)
name("ModelLevels", ws, f"B{rr[0] + 1}:E{rr[0] + 1}")
rr[0] += 2
rline("AI readiness hours per model (when included)", 16, "AIHrs", FMT_H)
rline("Security, ops & handover - hours per pair", 40, "HandoverHrs", FMT_H)
rline("Security, ops & handover - hours per model", 8, "SecPerModel", FMT_H)

rsec("Assessment and architecture", 4)
head(ws, rr[0], 1, ["Anchor (inventory pages)", "Pages", "Initiation h", "Analysis h"])
rr[0] += 1
r0 = rr[0]
for i, (px, ini, ana) in enumerate([(100, 40, 80), (500, 60, 160), (1000, 80, 240), (5000, 120, 360), (10000, 160, 480)]):
    put(ws, f"A{r0 + i}", f"Anchor {i + 1}")
    inp(ws, f"B{r0 + i}", px, fmt=FMT_H)
    inp(ws, f"C{r0 + i}", ini, fmt=FMT_H)
    inp(ws, f"D{r0 + i}", ana, fmt=FMT_H)
name("AnchorPages", ws, f"B{r0}:B{r0 + 4}")
name("AnchorInit", ws, f"C{r0}:C{r0 + 4}")
name("AnchorAnalysis", ws, f"D{r0}:D{r0 + 4}")
rr[0] = r0 + 5
rline("Assessment hours per semantic model", 4, "AssessPerModel", FMT_H)
rline("Architecture hours per target BI tool", 80, "ArchPerTool", FMT_H)
rline("Architecture hours per pair", 24, "ArchPerPair", FMT_H)

rsec("Automation caps (share of a phase automation can remove)")
for lab, v, nm_ in [("Assessment", .6, "CapAssess"), ("Semantic model", .4, "CapSem"), ("Report development", .6, "CapRep"),
                    ("SIT", .3, "CapSIT"), ("UAT support", .2, "CapUAT"), ("Deployment", .6, "CapDep")]:
    rline(f"Cap - {lab}", v, nm_, FMT_PCT0)

rsec("Testing, UAT, deployment, hypercare (per sprint)")
rline("SIT crew sizing (% of dev hours)", 0.15, "SITPct", FMT_PCT0)
rline("UAT support engineers when pages per sprint <= threshold", 1, "UATLow")
rline("UAT support engineers above threshold", 2, "UATHigh")
rline("Volume threshold (report pages per sprint)", 150, "UATThreshold", FMT_H)
rline("UAT support weeks per sprint", 1, "UATWk")
rline("Deployment engineers", 1, "DepEng")
rline("Deployment weeks per sprint", 1, "DepWk")
rline("Hypercare crew (FTE)", 2, "HCCrew")

# tools / pairs / matrices (right-hand side)
ARCH = ["Enterprise semantic-layer", "Paginated / pixel-perfect", "Modern visual / self-service", "Modern code-based semantic"]
TOOLS = [("SAP BusinessObjects", 0), ("OBIEE", 0), ("Oracle Analytics Server", 0), ("MicroStrategy", 0), ("Cognos", 0),
         ("SSRS", 1), ("Crystal Reports", 1), ("Tableau", 2), ("Power BI", 2), ("QuickSight", 2), ("Qlik Sense", 2),
         ("QlikView", 2), ("Looker Studio", 2), ("Looker", 3)]
section(ws, "I4", "BI tools", 3)
head(ws, 5, 9, ["BI tool", "Archetype"])
for i in range(30):
    t = TOOLS[i] if i < len(TOOLS) else None
    inp(ws, f"I{6 + i}", t[0] if t else None)
    inp(ws, f"J{6 + i}", ARCH[t[1]] if t else None)
name("ToolNames", ws, "I6:I35")
name("ToolArch", ws, "J6:J35")
section(ws, "M4", "Known pair factors", 4)
head(ws, 5, 13, ["Source", "Target", "Factor", "Key"])
KNOWN = [("OBIEE", "Power BI", 1.00), ("Cognos", "Power BI", 1.10), ("SAP BusinessObjects", "Power BI", 1.25),
         ("Tableau", "Power BI", 0.95), ("SSRS", "Power BI", 1.05), ("OBIEE", "QuickSight", 1.20),
         ("Cognos", "QuickSight", 1.25), ("Tableau", "QuickSight", 1.10), ("SSRS", "QuickSight", 1.20),
         ("SAP BusinessObjects", "QuickSight", 1.30), ("Cognos", "Looker", 1.25), ("OBIEE", "Looker", 1.20),
         ("Tableau", "Looker", 1.10), ("SSRS", "Looker", 1.20), ("SAP BusinessObjects", "Looker", 1.30)]
for i in range(30):
    k = KNOWN[i] if i < len(KNOWN) else (None, None, None)
    inp(ws, f"M{6 + i}", k[0])
    inp(ws, f"N{6 + i}", k[1])
    inp(ws, f"O{6 + i}", k[2], fmt="0.00")
    put(ws, f"P{6 + i}", f'=IF(M{6 + i}="","",M{6 + i}&"|"&N{6 + i})')
name("KnownKey", ws, "P6:P35")
name("KnownFactor", ws, "O6:O35")

r = 38
section(ws, f"I{r}", "Pair factor fallback (source row x target column)", 5)
head(ws, r + 1, 9, ["Source \\ Target"] + ARCH, height=45)
AF = [[1.10, 1.00, 1.10, 1.20], [1.00, 0.90, 1.05, 1.20], [1.10, 1.00, 0.95, 1.10], [1.00, 1.00, 1.00, 1.00]]
for i in range(4):
    put(ws, f"I{r + 2 + i}", ARCH[i], font=F_BOLD)
    for j in range(4):
        inp(ws, f"{L(10 + j)}{r + 2 + i}", AF[i][j], fmt="0.00")
name("ArchNames", ws, f"I{r + 2}:I{r + 5}")
name("ArchFactor", ws, f"J{r + 2}:M{r + 5}")
r = 46
section(ws, f"I{r}", "Semantic-model approach share (source row x target column)", 5)
head(ws, r + 1, 9, ["Source \\ Target"] + ARCH, height=45)
AS = [[1, 1, 1, 1], [.3, .3, .3, 1], [1, .3, .4, 1], [.4, .3, .4, .4]]
for i in range(4):
    put(ws, f"I{r + 2 + i}", ARCH[i], font=F_BOLD)
    for j in range(4):
        inp(ws, f"{L(10 + j)}{r + 2 + i}", AS[i][j], fmt=FMT_PCT0)
name("ApproachShare", ws, f"J{r + 2}:M{r + 5}")
put(ws, f"I{r + 6}", "100% = redesign / author LookML; 40% = translate dataset; 30% = shared datasets for paginated reports.", font=F_SUB, border=None)

r = 56
section(ws, f"I{r}", "Rate card (bill rates, USD/h)", 4)
head(ws, r + 1, 9, ["Role", "Grade | Location", "Rate", "Key"])
RATES = [("MSE II", "G13", "India", 87), ("MSE I", "G12", "India", 82), ("Architect II", "G12", "India", 70),
         ("Architect I", "G11", "India", 65), ("Lead II", "G11", "India", 70), ("Lead II", "G10", "India", 70),
         ("Lead I", "G10", "India", 62), ("Lead I", "G9", "India", 62), ("Sr. Engineer II", "G10", "India", 62),
         ("Sr. Engineer II", "G9", "India", 62), ("Sr. Engineer I", "G9", "India", 57.5), ("Sr. Engineer I", "G8", "India", 57.5),
         ("Engineer II", "G8", "India", 50), ("Engineer II", "G6", "India", 50), ("Engineer I", "G6", "India", 42),
         ("Engineer I", "G4", "India", 42), ("Jr. Engineer", "G4", "India", 35), ("Architect II", "G11", "India", 70),
         ("Sr. Architect I", "G13", "India", 87), ("Sr. Architect II", "G14", "India", 105),
         ("Sr. Architect II", "G13", "India", 105), ("Architect II", "G12", "USA", 185), ("MSE I", "G12", "USA", 190),
         ("MSE II", "G13", "USA", 210)]
R0 = r + 2
for i, (ro, g, loc, rate) in enumerate(RATES):
    put(ws, f"I{R0 + i}", ro)
    put(ws, f"J{R0 + i}", f"{g} | {loc}")
    inp(ws, f"K{R0 + i}", rate, fmt=FMT_USD2)
    put(ws, f"L{R0 + i}", f'=I{R0 + i}&" | "&J{R0 + i}')
ws.column_dimensions["L"].width = 32
name("RateKeys", ws, f"L{R0}:L{R0 + len(RATES) - 1}")
name("RateVals", ws, f"K{R0}:K{R0 + len(RATES) - 1}")

r = rr[0] + 2
section(ws, f"A{r}", "Role mix by phase (each row totals 100%)", 8)
head(ws, r + 1, 1, ["Phase"] + [ro for ro, _k in ROLES] + ["Total"], height=36)
MX0 = r + 2
for p, (ph, _c) in enumerate(PHASES):
    put(ws, f"A{MX0 + p}", ph)
    for j in range(NROLES):
        inp(ws, f"{L(2 + j)}{MX0 + p}", MIX[p][j], fmt=FMT_PCT0)
    put(ws, f"H{MX0 + p}", f"=SUM(B{MX0 + p}:G{MX0 + p})", fmt=FMT_PCT0, font=F_BOLD)
name("RoleMix", ws, f"B{MX0}:G{MX0 + NPH - 1}")
name("RoleMixTotals", ws, f"H{MX0}:H{MX0 + NPH - 1}")
ws.column_dimensions["H"].width = 9

LISTS = {"YesNo": ["Yes", "No"], "SprintLens": [1, 2]}
col = 18
for nm_, vals in LISTS.items():
    head(ws, 5, col, [nm_])
    for i, v in enumerate(vals):
        put(ws, f"{L(col)}{6 + i}", v)
    name(nm_, ws, f"{L(col)}6:{L(col)}{5 + len(vals)}")
    col += 1

# =====================================================================
# Estimate
# =====================================================================
ws = ws_est
ws.sheet_view.showGridLines = False
widths = {"A": 2, "B": 38, "C": 22, "D": 12, "E": 10, "F": 10, "G": 16, "H": 10, "I": 10, "J": 10, "K": 14, "L": 11,
          "M": 11, "N": 14, "O": 10, "P": 12, "Q": 12}
for c, w in widths.items():
    ws.column_dimensions[c].width = w
put(ws, "B1", "BI Migration - Pre-sales Ballpark Estimate", font=F_TITLE, border=None)
put(ws, "B2", "Fill the yellow cells. Results update instantly. For a quotable fixed price use the full v2 estimator.", font=F_SUB, border=None)

section(ws, "B4", "1. Project settings", 2)
SET = [
    ("Client", "Example Client", "Client", None, None),
    ("Opportunity", "BI modernisation", "Opportunity", None, None),
    ("Planned start date (Monday)", dt.date(2027, 1, 4), "StartDate", FMT_DATE, None),
    ("Sprint length (weeks)", 2, "SprintLen", None, "1 or 2 weeks. SIT and UAT follow each sprint."),
    ("Development team (engineers)", 6, "DevTeam", None, "Engineers building models and reports at peak."),
    ("Assessment team (FTE)", 2, "AssessTeam", None, None),
    ("Automation % (LeapLogic / accelerators)", 0, "Auto", FMT_PCT0, "Applied to every phase, capped per phase on Reference."),
    ("Contingency %", 0.15, "ContPct", FMT_PCT0, None),
    ("Include project manager?", "Yes", "IncPM", None, None),
    ("Include AI readiness?", "No", "IncAI", None, "Copilot / Q&A / Amazon Q / Gemini readiness per semantic model."),
    ("Hypercare weeks after go-live", 2, "HCWk", None, None),
]
for i, (lab, v, nm_, f, note) in enumerate(SET):
    r = 5 + i
    put(ws, f"B{r}", lab)
    inp(ws, f"C{r}", v, fmt=f, note=note)
    name(nm_, ws, f"C{r}")
dv_list(ws, "C8", "=SprintLens")
dv_list(ws, "C13:C14", "=YesNo")
dv_num(ws, "C9:C10", 0.5, 100)
dv_num(ws, "C11:C12", 0, 1)
dv_num(ws, "C15", 0, 26)

# results block
section(ws, "G4", "Results", 5)
RES = [
    ("Expected price", "=ExpPrice", FMT_USD),
    ("Ballpark range", '=TEXT(ExpPrice*(1+RangeLow),"$#,##0")&"  to  "&TEXT(ExpPrice*(1+RangeHigh),"$#,##0")', None),
    ("Total effort (hours)", "=PlanHrs", FMT_H),
    ("Duration (weeks)", "=DurWeeks", FMT_H),
    ("Sprints", '=Sprints&" x "&SprintLen&"-week"', None),
    ("Peak team (FTE)", "=PeakTeam", FMT_H1),
    ("Average team (FTE)", "=AvgTeam", FMT_H1),
    ("Blended rate (USD/h)", "=BlendedRate", FMT_USD2),
    ("Planned end date", "=StartDate+DurWeeks*7-3", FMT_DATE),
    ("Status", '=IF(CheckFails=0,"OK - all checks pass",CheckFails&" check(s) to fix - see section 4")', None),
]
for i, (lab, fml, f) in enumerate(RES):
    r = 5 + i
    ws.merge_cells(f"G{r}:I{r}")
    put(ws, f"G{r}", lab, font=F_BOLD, fill=FILL_CALC)
    ws.merge_cells(f"J{r}:K{r}")
    put(ws, f"J{r}", fml, font=F_KPI if i == 0 else F_BOLD, fmt=f, fill=FILL_CALC, align=Alignment(horizontal="right"))
ws.row_dimensions[5].height = 22
ws.conditional_formatting.add("J14", FormulaRule(formula=["CheckFails=0"], fill=FILL_GOOD))
ws.conditional_formatting.add("J14", FormulaRule(formula=["CheckFails>0"], fill=FILL_BAD))

# scope table
SR = 18
section(ws, f"B{SR - 1}", "2. Scope - one row per source -> target pair (most deals need one row)", 16)
SCOPE_HEAD = ["Source BI tool", "Target BI tool", "Data platform changing?", "Reports", "Avg pages per report",
              "Complexity profile", "% Simple", "% Medium", "% Complex", "% Very complex", "% retired / duplicates",
              "Semantic models (target)", "Model complexity", "Pair factor", "Report hours", "Model hours"]
head(ws, SR, 2, SCOPE_HEAD, height=45)
EX = [("OBIEE", "Power BI", "Yes", 300, 2, "Balanced", None, None, None, None, 0.2, 6, "Medium")]
for i in range(NPAIRS):
    r = SR + 1 + i
    ex = EX[i] if i < len(EX) else (None,) * 13
    fmts = [None, None, None, FMT_H, "0.0", None, FMT_PCT0, FMT_PCT0, FMT_PCT0, FMT_PCT0, FMT_PCT0, FMT_H, None]
    for j in range(13):
        inp(ws, f"{L(2 + j)}{r}", ex[j], fmt=fmts[j])
P1, PN = SR + 1, SR + NPAIRS
dv_list(ws, f"B{P1}:C{PN}", "=ToolNames", free=True)
dv_list(ws, f"D{P1}:D{PN}", "=YesNo")
dv_list(ws, f"G{P1}:G{PN}", "=ProfileNames")
dv_list(ws, f"N{P1}:N{PN}", '"Simple,Medium,Complex,Very complex"')
dv_num(ws, f"H{P1}:L{PN}", 0, 1)
dv_num(ws, f"F{P1}:F{PN}", 1, 100)
dv_num(ws, f"E{P1}:E{PN}", 0, 1000000)
dv_num(ws, f"M{P1}:M{PN}", 0, 10000)
put(ws, f"B{PN + 1}", "% columns are used only when the profile is Custom. A new BI tool must be added on Reference with its archetype.", font=F_SUB, border=None)

# workings (grouped columns S:AE)
WK = ["Active", "Mix S", "Mix M", "Mix C", "Mix VC", "Src archetype", "Tgt archetype", "Approach share",
      "Inventory pages", "Pages to build", "Analysis h", "Mix flag", "Platform change"]
WC0 = 19  # S
head(ws, SR, WC0, WK, height=45)
for i in range(NPAIRS):
    r = SR + 1 + i
    c = {k: f"{L(WC0 + j)}{r}" for j, k in enumerate(["act", "ms", "mm", "mc", "mv", "sa", "ta", "ash", "inv", "pb", "ana", "flag", "dp"])}
    put(ws, c["act"], f'=IF(AND(B{r}<>"",C{r}<>"",N(E{r})>0),1,0)', fill=FILL_CALC)
    for k, (key, pc) in enumerate([("ms", "H"), ("mm", "I"), ("mc", "J"), ("mv", "K")]):
        put(ws, c[key], f'=IF(G{r}="Custom",N({pc}{r}),IFERROR(INDEX(ProfileMix,MATCH(G{r},ProfileNames,0),{k + 1}),INDEX(ProfileMix,2,{k + 1})))', fmt=FMT_PCT0, fill=FILL_CALC)
    put(ws, c["sa"], f'=IFERROR(INDEX(ToolArch,MATCH(B{r},ToolNames,0)),"Unmapped")', fill=FILL_CALC)
    put(ws, c["ta"], f'=IFERROR(INDEX(ToolArch,MATCH(C{r},ToolNames,0)),"Unmapped")', fill=FILL_CALC)
    put(ws, c["ash"], f'=IFERROR(INDEX(ApproachShare,MATCH({c["sa"]},ArchNames,0),MATCH({c["ta"]},ArchNames,0)),1)', fmt=FMT_PCT0, fill=FILL_CALC)
    put(ws, c["inv"], f'={c["act"]}*N(E{r})*MAX(1,N(F{r}))', fmt=FMT_H, fill=FILL_CALC)
    put(ws, c["pb"], f'={c["inv"]}*(1-MIN(1,N(L{r})))', fmt=FMT_H, fill=FILL_CALC)
    put(ws, c["ana"], f'=IF({c["act"]}=1,{interp(c["inv"], "AnchorPages", "AnchorAnalysis")}*O{r},0)', fmt=FMT_H, fill=FILL_CALC)
    put(ws, c["flag"], f'=IF(AND({c["act"]}=1,OR(ABS({c["ms"]}+{c["mm"]}+{c["mc"]}+{c["mv"]}-1)>0.001,{c["sa"]}="Unmapped",{c["ta"]}="Unmapped")),1,0)', fill=FILL_CALC)
    put(ws, c["dp"], f'=IF(AND({c["act"]}=1,D{r}="Yes"),1,0)', fill=FILL_CALC)
    # visible computed columns O:Q
    put(ws, f"O{r}", f'=IF({c["act"]}=0,0,IFERROR(INDEX(KnownFactor,MATCH(B{r}&"|"&C{r},KnownKey,0)),IFERROR(INDEX(ArchFactor,MATCH({c["sa"]},ArchNames,0),MATCH({c["ta"]},ArchNames,0)),1)))', fmt=FMT_X, fill=FILL_CALC)
    hpp = f'({c["ms"]}*INDEX(HrsPage,1)+{c["mm"]}*INDEX(HrsPage,2)+{c["mc"]}*INDEX(HrsPage,3)+{c["mv"]}*INDEX(HrsPage,4))'
    put(ws, f"P{r}", f'={c["pb"]}*{hpp}*O{r}*(1+{c["dp"]}*DPRep)*(1-LearnAvg)*(1-Auto*CapRep)', fmt=FMT_H, fill=FILL_CALC)
    put(ws, f"Q{r}", f'={c["act"]}*N(M{r})*(IFERROR(INDEX(ModelHrs,MATCH(N{r},ModelLevels,0)),INDEX(ModelHrs,2))*{c["ash"]}*O{r}*(1+{c["dp"]}*DPModel)*(1-Auto*CapSem)+IF(IncAI="Yes",AIHrs,0))', fmt=FMT_H, fill=FILL_CALC)
for j in range(len(WK)):
    ws.column_dimensions[L(WC0 + j)].width = 12
    ws.column_dimensions[L(WC0 + j)].outlineLevel = 1
    ws.column_dimensions[L(WC0 + j)].hidden = True
ws.column_dimensions["R"].width = 3
put(ws, f"S{SR - 1}", "Workings (expand the + above column R to see)", font=F_SUB, border=None)


def wcol(k):
    return L(WC0 + ["act", "ms", "mm", "mc", "mv", "sa", "ta", "ash", "inv", "pb", "ana", "flag", "dp"].index(k))


TOTR = PN + 2
put(ws, f"N{TOTR}", "Total", font=F_BOLD)
put(ws, f"P{TOTR}", f"=SUM(P{P1}:P{PN})", fmt=FMT_H, font=F_BOLD)
put(ws, f"Q{TOTR}", f"=SUM(Q{P1}:Q{PN})", fmt=FMT_H, font=F_BOLD)
name("RepHrs", ws, f"P{TOTR}")
name("ModHrs", ws, f"Q{TOTR}")

# ---- hidden calculation block for phases (right of results, rows 26+ in columns S:V)
CB = TOTR + 3
section(ws, f"B{CB}", "3. Effort and cost by phase", 7)
head(ws, CB + 1, 2, ["Phase", "Start week", "Weeks", "Avg FTE", "Hours", "Cost (USD)", "Basis"], height=30)
act = f"{wcol('act')}{P1}:{wcol('act')}{PN}"
calcs = {
    "ActivePairs": f"=SUM({act})",
    "InvPages": f"=SUM({wcol('inv')}{P1}:{wcol('inv')}{PN})",
    "PagesBuilt": f"=SUM({wcol('pb')}{P1}:{wcol('pb')}{PN})",
    "Models": f"=SUMPRODUCT({act},M{P1}:M{PN})",
    "TargetTools": f'=SUMPRODUCT((C{P1}:C{PN}<>"")/COUNTIF(C{P1}:C{PN},C{P1}:C{PN}&""))',
    "AssessHrs": f"=({interp('InvPages', 'AnchorPages', 'AnchorInit')}+SUM({wcol('ana')}{P1}:{wcol('ana')}{PN})+Models*AssessPerModel)*(1-Auto*CapAssess)+TargetTools*ArchPerTool+ActivePairs*ArchPerPair",
    "DevHrs": "=RepHrs+ModHrs+ActivePairs*HandoverHrs+Models*SecPerModel",
    "PlanWeeks": "=MAX(1,ROUNDUP(AssessHrs/(AssessTeam*HrsWk),0))",
    "Sprints": "=MAX(1,ROUNDUP(DevHrs/(DevTeam*HrsWk)/SprintLen,0))",
    "DevWeeks": "=Sprints*SprintLen",
    "DevStart": "=PlanWeeks+1",
    "PagesPerSprint": "=PagesBuilt/Sprints",
    "UATEnd": "=DevStart+2*SprintLen+DevWeeks-1",
    "DurWeeks": "=UATEnd+HCWk",
}
KR = CB + 15
put(ws, f"S{KR - 1}", "Phase workings", font=F_BOLD, border=None)
for i, (nm_, fml) in enumerate(calcs.items()):
    r = KR + i
    put(ws, f"S{r}", nm_, fill=FILL_CALC)
    put(ws, f"T{r}", fml, fmt=FMT_H1, fill=FILL_CALC)
    name(nm_, ws, f"T{r}")
PH = [  # start, weeks, FTE, basis
    ("1", "PlanWeeks", "CEILING(AssessHrs/(PlanWeeks*HrsWk),0.5)", "Initiation + analysis (interpolated) + architecture"),
    ("DevStart", "DevWeeks", "MIN(DevTeam,CEILING(DevHrs/(DevWeeks*HrsWk),0.5))", "Models + reports + security/handover, paced in sprints"),
    ("DevStart+SprintLen", "DevWeeks", "MAX(0.5,CEILING(SITPct*DevHrs*(1-Auto*CapSIT)/(DevWeeks*HrsWk),0.5))", "One sprint behind development"),
    ("DevStart+2*SprintLen", "DevWeeks", "IF(PagesPerSprint<=UATThreshold,UATLow,UATHigh)*UATWk/SprintLen*(1-Auto*CapUAT)", "Two sprints behind; 1-2 engineers by volume"),
    ("DevStart+2*SprintLen", "DevWeeks", "DepEng*DepWk/SprintLen*(1-Auto*CapDep)", "Deployment support each sprint"),
    ("UATEnd+1", "HCWk", "HCCrew", "After final go-live"),
    ("1", "DurWeeks", 'IF(IncPM="Yes",1,0)', "Whole duration"),
]
PT0 = CB + 2
for p, (st, wk, fte, basis) in enumerate(PH):
    r = PT0 + p
    put(ws, f"B{r}", PHASES[p][0], fill=PatternFill("solid", fgColor=PHASES[p][1]))
    put(ws, f"C{r}", f"={st}", fmt="0")
    put(ws, f"D{r}", f"={wk}", fmt="0")
    put(ws, f"E{r}", f"=IF(ActivePairs=0,0,{fte})", fmt="0.0")
    put(ws, f"F{r}", f"=D{r}*E{r}*HrsWk", fmt=FMT_H)
    put(ws, f"G{r}", f"=F{r}*SUMPRODUCT(MMULT(Reference!$B${MX0 + p}:$G${MX0 + p},PlanRates))", fmt=FMT_USD)
    ws.merge_cells(f"H{r}:K{r}")
    put(ws, f"H{r}", basis, font=F_SUB)
name("PhStart", ws, f"C{PT0}:C{PT0 + NPH - 1}")
name("PhWeeks", ws, f"D{PT0}:D{PT0 + NPH - 1}")
name("PhFTE", ws, f"E{PT0}:E{PT0 + NPH - 1}")
name("PhHours", ws, f"F{PT0}:F{PT0 + NPH - 1}")
r = PT0 + NPH
put(ws, f"B{r}", "Total (from phases)", font=F_BOLD)
put(ws, f"F{r}", "=SUM(PhHours)", fmt=FMT_H, font=F_BOLD)
put(ws, f"G{r}", f"=SUM(G{PT0}:G{r - 1})", fmt=FMT_USD, font=F_BOLD)
name("EstHrs", ws, f"F{r}")
put(ws, f"B{r + 1}", "Team cost (from Plan grid)", font=F_BOLD)
put(ws, f"G{r + 1}", "=PlanCost", fmt=FMT_USD, font=F_BOLD)
put(ws, f"B{r + 2}", "Contingency", font=F_BOLD)
put(ws, f"G{r + 2}", "=PlanCost*ContPct", fmt=FMT_USD, font=F_BOLD)
put(ws, f"B{r + 3}", "EXPECTED PRICE", font=F_BOLD, fill=FILL_SECTION)
put(ws, f"G{r + 3}", "=PlanCost*(1+ContPct)", fmt=FMT_USD, font=F_BOLD, fill=FILL_SECTION)
name("ExpPrice", ws, f"G{r + 3}")

CK = r + 6
section(ws, f"B{CK}", "4. Checks", 7)
CHECKS = [
    ("At least one scope row filled in", "ActivePairs>=1"),
    ("Mixes total 100% and every BI tool is known", f"SUM({wcol('flag')}{P1}:{wcol('flag')}{PN})=0"),
    ("Role mix rows total 100% (Reference)", "SUMPRODUCT(--(ABS(RoleMixTotals-1)>0.0001))=0"),
    (f"Duration fits the {NWEEKS}-week plan", f"DurWeeks<={NWEEKS}"),
    ("Plan grid matches the phase estimate (within 5%)", "ABS(PlanHrs-EstHrs)<=0.05*MAX(1,EstHrs)"),
]
for i, (lab, cond) in enumerate(CHECKS):
    rr_ = CK + 1 + i
    put(ws, f"B{rr_}", lab)
    ws.merge_cells(f"B{rr_}:F{rr_}")
    put(ws, f"G{rr_}", f'=IF(IFERROR({cond},FALSE),"PASS","FAIL")', font=F_BOLD)
CK1 = CK + len(CHECKS)
ws.conditional_formatting.add(f"G{CK + 1}:G{CK1}", CellIsRule(operator="equal", formula=['"PASS"'], fill=FILL_GOOD))
ws.conditional_formatting.add(f"G{CK + 1}:G{CK1}", CellIsRule(operator="equal", formula=['"FAIL"'], fill=FILL_BAD))
KF = KR + len(calcs)
put(ws, f"S{KF}", "CheckFails", fill=FILL_CALC)
put(ws, f"T{KF}", f'=COUNTIF(G{CK + 1}:G{CK1},"FAIL")', fill=FILL_CALC)
name("CheckFails", ws, f"T{KF}")
for col_ in ("S", "T"):
    ws.column_dimensions[col_].hidden = True
ws.freeze_panes = "A4"

# =====================================================================
# Plan (sample-style sprint grid)
# =====================================================================
ws = ws_plan
ws.sheet_view.showGridLines = False
put(ws, "A1", '="Delivery plan - "&Client', font=F_TITLE, border=None)
put(ws, "A2", "Phase bars show FTE per week. Team rows are pre-filled from the phases - overwrite any yellow cell to hand-plan; cost follows the grid.", font=F_SUB, border=None)
W0 = 3  # column C = week 1
WL = W0 + NWEEKS - 1
ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 10
for i in range(NWEEKS):
    ws.column_dimensions[L(W0 + i)].width = 5.5
TC = WL + 1  # totals columns
for j, (h, w) in enumerate([("Person-weeks", 12), ("Hours", 10), ("Cost (USD)", 13)]):
    ws.column_dimensions[L(TC + j)].width = w
put(ws, "A4", "Sprint", font=F_BOLD, fill=FILL_CALC)
put(ws, "A5", "Week", font=F_BOLD, fill=FILL_CALC)
put(ws, "A6", "Week starting", font=F_SUB)
for i in range(NWEEKS):
    c = L(W0 + i)
    put(ws, f"{c}5", f"Wk{i + 1}", font=F_BOLD, fill=FILL_CALC, align=CENTER)
    put(ws, f"{c}4", f'=IF(AND(MOD({i},SprintLen)=0,{i + 1}<=DurWeeks),"S"&({i}/SprintLen+1),"")', font=F_BOLD, fill=FILL_CALC, align=CENTER)
    put(ws, f"{c}6", f"=StartDate+{i}*7", fmt="dd-mmm", font=Font(name=FONT, size=8, color="595959"), align=CENTER)
PR0 = 8
put(ws, f"A{PR0 - 1}", "Phases (FTE per week)", font=F_SECTION, border=None)
for p, (ph, colr) in enumerate(PHASES):
    r = PR0 + p
    put(ws, f"A{r}", ph, fill=PatternFill("solid", fgColor=colr))
    for i in range(NWEEKS):
        c = L(W0 + i)
        put(ws, f"{c}{r}", f"=IF(AND({i + 1}>=INDEX(PhStart,{p + 1}),{i + 1}<INDEX(PhStart,{p + 1})+INDEX(PhWeeks,{p + 1})),INDEX(PhFTE,{p + 1}),0)", fmt=FMT_FTE, align=CENTER)
    ws.conditional_formatting.add(f"{L(W0)}{r}:{L(WL)}{r}", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor=colr)))
    put(ws, f"{L(TC)}{r}", f"=SUM({L(W0)}{r}:{L(WL)}{r})", fmt=FMT_H1)
    put(ws, f"{L(TC + 1)}{r}", f"={L(TC)}{r}*HrsWk", fmt=FMT_H)
r = PR0 + NPH
put(ws, f"A{r}", "Total FTE", font=F_BOLD)
for i in range(NWEEKS):
    c = L(W0 + i)
    put(ws, f"{c}{r}", f"=SUM({c}{PR0}:{c}{r - 1})", fmt=FMT_FTE, font=F_BOLD, align=CENTER)

TR0 = PR0 + NPH + 3
head(ws, TR0 - 1, 1, ["Team", "Rate (USD/h)"] + [f"Wk{i + 1}" for i in range(NWEEKS)] + ["Person-weeks", "Hours", "Cost (USD)"], height=20)
for j, (ro, key) in enumerate(ROLES):
    r = TR0 + j
    put(ws, f"A{r}", ro, font=F_BOLD)
    inp(ws, f"B{r}", f"=IFERROR(INDEX(RateVals,MATCH(\"{key}\",RateKeys,0)),0)", fmt=FMT_USD2,
        note=f"Priced as {key} from the Reference rate card. Overwrite to use another rate.")
    mixcol = L(2 + j)
    for i in range(NWEEKS):
        c = L(W0 + i)
        inp(ws, f"{c}{r}", f"=SUMPRODUCT({c}${PR0}:{c}${PR0 + NPH - 1},Reference!${mixcol}${MX0}:${mixcol}${MX0 + NPH - 1})", fmt=FMT_FTE, fill=FILL_GRID)
        ws[f"{c}{r}"].alignment = CENTER
    put(ws, f"{L(TC)}{r}", f"=SUM({L(W0)}{r}:{L(WL)}{r})", fmt=FMT_H1)
    put(ws, f"{L(TC + 1)}{r}", f"={L(TC)}{r}*HrsWk", fmt=FMT_H)
    put(ws, f"{L(TC + 2)}{r}", f"={L(TC + 1)}{r}*B{r}", fmt=FMT_USD)
name("PlanRates", ws, f"B{TR0}:B{TR0 + NROLES - 1}")
r = TR0 + NROLES
put(ws, f"A{r}", "Total team", font=F_BOLD)
for i in range(NWEEKS):
    c = L(W0 + i)
    put(ws, f"{c}{r}", f"=SUM({c}{TR0}:{c}{r - 1})", fmt=FMT_FTE, font=F_BOLD, align=CENTER)
for j, f in enumerate([FMT_H1, FMT_H, FMT_USD]):
    cc = L(TC + j)
    put(ws, f"{cc}{r}", f"=SUM({cc}{TR0}:{cc}{r - 1})", fmt=f, font=F_BOLD)
name("PlanHeadcount", ws, f"{L(W0)}{r}:{L(WL)}{r}")
name("PlanPW", ws, f"{L(TC)}{r}")
name("PlanHrs", ws, f"{L(TC + 1)}{r}")
name("PlanCost", ws, f"{L(TC + 2)}{r}")
S0 = r + 2
SUMM = [("Hrs per week", "=HrsWk", FMT_H), ("Peak team (FTE)", "=MAX(PlanHeadcount)", FMT_H1),
        ("Average team (FTE)", "=IF(DurWeeks>0,PlanPW/DurWeeks,0)", FMT_H1),
        ("Blended rate (USD/h)", "=IF(PlanHrs>0,PlanCost/PlanHrs,0)", FMT_USD2),
        ("Plan hours vs phase estimate", "=IF(EstHrs>0,PlanHrs/EstHrs-1,0)", FMT_PCT)]
for i, (lab, fml, f) in enumerate(SUMM):
    put(ws, f"A{S0 + i}", lab, font=F_BOLD, fill=FILL_CALC)
    put(ws, f"B{S0 + i}", fml, fmt=f, font=F_BOLD)
name("PeakTeam", ws, f"B{S0 + 1}")
name("AvgTeam", ws, f"B{S0 + 2}")
name("BlendedRate", ws, f"B{S0 + 3}")
ws.freeze_panes = f"{L(W0)}7"

# =====================================================================
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.font is not None and c.font.name != FONT:
                f = c.font
                c.font = Font(name=FONT, size=f.size or 10, bold=f.bold, italic=f.italic, color=f.color)
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
ws_est.page_setup.orientation = "landscape"
ws_est.page_setup.fitToWidth = 1
ws_est.sheet_properties.pageSetUpPr.fitToPage = True
wb.active = 0
wb.save(OUT)
print("wrote", OUT)
