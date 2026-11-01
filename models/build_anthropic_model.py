"""
Builds the Anthropic three-statement model (models/Anthropic_3_Statement_Model.xlsx).

Every forecast cell is a live Excel formula driven from the Assumptions sheet.
Historical figures are taken from Anthropic's Form S-1 (filed 2026-10-26) and
company disclosures to investors (Q3'26 revenue). Columns are aligned across
all sheets so cross-sheet formulas are column-stable:

    C FY2023A | D FY2024A | E FY2025A | F H1'26A | G H2'26E | H FY2026E |
    I FY2027E | J FY2028E | K FY2029E | L FY2030E

Run:  python3 models/build_anthropic_model.py
"""
from __future__ import annotations

import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = Path(__file__).resolve().parent / "Anthropic_3_Statement_Model.xlsx"

# --------------------------------------------------------------------------- #
# Layout constants
# --------------------------------------------------------------------------- #
COLS = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]
HIST = ["C", "D", "E", "F"]          # filed actuals (FY23, FY24, FY25, H1'26)
FC = ["G", "I", "J", "K", "L"]       # forecast calculation columns
PREV = {"G": "F", "I": "H", "J": "I", "K": "J", "L": "K"}
FRAC = {"C": 1, "D": 1, "E": 1, "F": 0.5, "G": 0.5, "H": 1, "I": 1, "J": 1, "K": 1, "L": 1}
FLOW_LABELS = ["FY2023A", "FY2024A", "FY2025A", "H1'26A", "H2'26E",
               "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"]
BS_LABELS = ["Dec-23A", "Dec-24A", "Dec-25A", "Jun-26A", "Dec-26E",
             "Dec-26E", "Dec-27E", "Dec-28E", "Dec-29E", "Dec-30E"]
FY_COLS = ["C", "D", "E", "H", "I", "J", "K", "L"]   # fiscal-year-only view
NOTE_COL = "M"

NUM = '#,##0;(#,##0);"–"'
NUM1 = '#,##0.0;(#,##0.0);"–"'
PCT = '0.0%;(0.0%);"–"'
MULT = '0.00"x";(0.00"x");"–"'
INT = '#,##0;(#,##0);"–"'

BLUE = Font(name="Calibri", size=10, color="0000FF")
BLUE_B = Font(name="Calibri", size=10, color="0000FF", bold=True)
BLACK = Font(name="Calibri", size=10, color="000000")
BLACK_B = Font(name="Calibri", size=10, color="000000", bold=True)
GREEN = Font(name="Calibri", size=10, color="007A33")
GREEN_B = Font(name="Calibri", size=10, color="007A33", bold=True)
GREY_I = Font(name="Calibri", size=9, color="7F7F7F", italic=True)
WHITE_B = Font(name="Calibri", size=10, color="FFFFFF", bold=True)
TITLE = Font(name="Calibri", size=16, bold=True, color="1F3864")
SUB = Font(name="Calibri", size=10, italic=True, color="595959")
SECTION = Font(name="Calibri", size=10, bold=True, color="1F3864")

FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_HIST = PatternFill("solid", fgColor="D9E1F2")
FILL_FC = PatternFill("solid", fgColor="E2EFDA")
FILL_SEC = PatternFill("solid", fgColor="F2F2F2")
FILL_LIVE = PatternFill("solid", fgColor="FFF2CC")
FILL_CHECK_OK = PatternFill("solid", fgColor="C6EFCE")
THIN = Side(style="thin", color="808080")
TOP = Border(top=THIN)
TOPBOT = Border(top=THIN, bottom=Side(style="double", color="808080"))

TOKEN = re.compile(r"\{([^!{}]+)!([^@}]+?)(?:@([A-Z]+))?\}")


# --------------------------------------------------------------------------- #
# Workbook helper with deferred cross-reference resolution
# --------------------------------------------------------------------------- #
class Model:
    def __init__(self):
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.rows: dict[tuple[str, str], int] = {}
        self.pending: list[tuple[str, str, str]] = []   # (sheet, coord, raw formula)
        self.scalar_col = "C"

    def sheet(self, name: str, flows: bool = True) -> "Sheet":
        return Sheet(self, name, flows)

    def resolve(self, sheet: str, raw: str) -> str:
        def sub(m: re.Match) -> str:
            s, key, col = m.group(1), m.group(2), m.group(3)
            if (s, key) not in self.rows:
                raise KeyError(f"Unknown reference {s}!{key} (from sheet {sheet})")
            row = self.rows[(s, key)]
            if col is None:  # scalar input stored in column C, absolute
                ref = f"${self.scalar_col}${row}"
            else:
                ref = f"{col}{row}"
            return ref if s == sheet else f"'{s}'!{ref}"
        return TOKEN.sub(sub, raw)

    def finalize(self):
        for sheet, coord, raw in self.pending:
            ws = self.wb[sheet]
            formula = self.resolve(sheet, raw)
            ws[coord].value = formula
            cell = ws[coord]
            bold = cell.font.bold
            if "'" in formula:        # cross-sheet link -> green
                cell.font = GREEN_B if bold else GREEN
            else:
                cell.font = BLACK_B if bold else BLACK


class Sheet:
    def __init__(self, model: Model, name: str, flows: bool):
        self.m = model
        self.name = name
        self.ws = model.wb.create_sheet(name)
        self.flows = flows
        self.r = 1
        self.ws.column_dimensions["A"].width = 50
        self.ws.column_dimensions["B"].width = 9
        for c in COLS:
            self.ws.column_dimensions[c].width = 11.5
        self.ws.column_dimensions[NOTE_COL].width = 70
        self.ws.sheet_view.showGridLines = False

    # -- primitive writes ---------------------------------------------------- #
    def put(self, coord: str, value, fmt: str | None = None, bold: bool = False,
            fill: PatternFill | None = None, font: Font | None = None,
            border: Border | None = None, align: str | None = None):
        cell = self.ws[coord]
        if isinstance(value, str) and value.startswith("="):
            cell.value = None
            cell.font = BLACK_B if bold else BLACK
            self.m.pending.append((self.name, coord, value))
        else:
            cell.value = value
            if font is not None:
                cell.font = font
            elif isinstance(value, (int, float)):
                cell.font = BLUE_B if bold else BLUE
            else:
                cell.font = BLACK_B if bold else BLACK
        if font is not None:
            cell.font = font
        if fmt:
            cell.number_format = fmt
        if fill:
            cell.fill = fill
        if border:
            cell.border = border
        if align:
            cell.alignment = Alignment(horizontal=align)
        return cell

    def title(self, text: str, subtitle: str | None = None):
        self.put(f"A{self.r}", text, font=TITLE)
        self.r += 1
        if subtitle:
            self.put(f"A{self.r}", subtitle, font=SUB)
            self.r += 1
        self.r += 1

    def period_header(self, labels=None, unit_label="$M unless noted"):
        labels = labels or (FLOW_LABELS if self.flows else BS_LABELS)
        row = self.r
        self.put(f"A{row}", unit_label, font=WHITE_B, fill=FILL_HDR)
        self.put(f"B{row}", "Unit", font=WHITE_B, fill=FILL_HDR, align="center")
        for c, lab in zip(COLS, labels):
            self.put(f"{c}{row}", lab, font=WHITE_B, fill=FILL_HDR, align="center")
        self.put(f"{NOTE_COL}{row}", "Notes / source", font=WHITE_B, fill=FILL_HDR)
        # status row (Actual / Forecast)
        self.r += 1
        for c in COLS:
            is_hist = c in HIST
            txt = "Actual" if is_hist else "Forecast"
            if c == "H":
                txt = "H1A + H2E" if self.flows else "same as H2'26E"
            self.put(f"{c}{self.r}", txt, font=GREY_I, fill=FILL_HIST if is_hist else FILL_FC, align="center")
        self.r += 1
        self.ws.freeze_panes = f"C{self.r}"
        return row

    def section(self, text: str):
        self.r += 1
        for col in ["A", "B"] + COLS + [NOTE_COL]:
            self.ws[f"{col}{self.r}"].fill = FILL_SEC
        self.put(f"A{self.r}", text, font=SECTION, fill=FILL_SEC)
        self.r += 1

    def blank(self, n: int = 1):
        self.r += n

    def note(self, text: str):
        self.put(f"A{self.r}", text, font=GREY_I)
        self.r += 1

    # -- the main row writer ------------------------------------------------- #
    def line(self, key: str | None, label: str, hist=None, fc=None, h=None,
             fmt: str = NUM, unit: str = "$M", bold: bool = False, note: str | None = None,
             hist_cols=None, fc_cols=None, border: Border | None = None,
             fill: PatternFill | None = None, indent: int = 0) -> int:
        row = self.r
        if key:
            self.m.rows[(self.name, key)] = row
        self.put(f"A{row}", ("   " * indent) + label, bold=bold)
        self.put(f"B{row}", unit, font=GREY_I, align="center")
        hist_cols = HIST if hist_cols is None else hist_cols
        fc_cols = FC if fc_cols is None else fc_cols

        if hist is not None:
            if callable(hist):
                for c in hist_cols:
                    self.put(f"{c}{row}", hist(c), fmt, bold)
            else:
                for c, v in zip(hist_cols, hist):
                    if v is None:
                        continue
                    self.put(f"{c}{row}", v, fmt, bold)
        if fc is not None:
            if callable(fc):
                for c in fc_cols:
                    self.put(f"{c}{row}", fc(c, PREV[c]), fmt, bold)
            else:
                for c, v in zip(fc_cols, fc):
                    if v is None:
                        continue
                    self.put(f"{c}{row}", v, fmt, bold)
        if h is not None:
            if h == "sum":
                self.put(f"H{row}", f"=F{row}+G{row}", fmt, bold)
            elif h == "same":
                self.put(f"H{row}", f"=G{row}", fmt, bold)
            elif callable(h):
                self.put(f"H{row}", h("H"), fmt, bold)
            else:
                self.put(f"H{row}", h, fmt, bold)
        if note:
            self.put(f"{NOTE_COL}{row}", note, font=GREY_I)
        if border or fill:
            for c in ["A", "B"] + COLS:
                if border:
                    self.ws[f"{c}{row}"].border = border
                if fill:
                    self.ws[f"{c}{row}"].fill = fill
        self.r += 1
        return row

    def scalar(self, key: str, label: str, value, fmt=NUM, unit="", note=None):
        row = self.r
        self.m.rows[(self.name, key)] = row
        self.put(f"A{row}", label)
        self.put(f"B{row}", unit, font=GREY_I, align="center")
        self.put(f"C{row}", value, fmt)
        if note:
            self.put(f"{NOTE_COL}{row}", note, font=GREY_I)
        self.r += 1
        return row

    def scenario_block(self, key: str, label: str, base, bull, bear, fmt=PCT, unit="%",
                       note=None, fc_cols=None):
        """Three input rows + a live row selected by the scenario switch."""
        fc_cols = FC if fc_cols is None else fc_cols
        r_base = self.line(f"{key}_base", f"{label} – Base", fc=base, fmt=fmt, unit=unit, fc_cols=fc_cols, indent=1)
        r_bull = self.line(f"{key}_bull", f"{label} – Bull", fc=bull, fmt=fmt, unit=unit, fc_cols=fc_cols, indent=1)
        r_bear = self.line(f"{key}_bear", f"{label} – Bear", fc=bear, fmt=fmt, unit=unit, fc_cols=fc_cols, indent=1)
        row = self.line(
            key, f"{label} – LIVE (selected scenario)",
            fc=lambda c, p: f"=CHOOSE({{Assumptions!scenario}},{c}{r_base},{c}{r_bull},{c}{r_bear})",
            fmt=fmt, unit=unit, bold=True, note=note, fc_cols=fc_cols, fill=FILL_LIVE)
        return row


# --------------------------------------------------------------------------- #
# Filed historical data (USD millions) – Form S-1, 26 Oct 2026
# --------------------------------------------------------------------------- #
# Order: FY2023A, FY2024A, FY2025A, H1'26A
D = {
    "revenue": [41, 386, 4585, 16286],
    "rev_cons": [39, 288, 3796, 13064],
    "rev_sub": [2, 98, 789, 3222],
    # nature of expense
    "compute": [433, 2522, 7326, 8945],
    "personnel": [95, 454, 1839, 3152],
    "bartz": [0, 0, 1500, 0],
    "smp": [2, 43, 568, 3168],
    "other_oa": [62, 347, 1412, 1714],
    "opex_total": [592, 3366, 12645, 16979],
    # function of expense
    "cor": [26, 752, 2773, 2380],
    "rnd": [422, 2275, 6338, 9741],
    "snm": [76, 163, 911, 3808],
    "gna": [68, 176, 1123, 1050],
    "op_income": [-551, -2980, -8060, -693],
    "interest_income": [69, 106, 431, 1111],
    "fv_change": [-1045, -5434, -34308, -81108],   # incl. $8,720M "other financial instruments" in H1'26
    "other_inc": [-2, 0, -8, -37],
    "pretax": [-1529, -8308, -41945, -80727],
    "tax": [2, 1, 20, 21],
    "net_income": [-1531, -8309, -41965, -80748],
    # cash flow
    "sbc": [29, 192, 1158, 2185],
    "da": [3, 21, 40, 97],
    "donation": [0, 0, 541, 126],
    "cf_other": [0, 7, 11, -7],
    "d_ar": [-10, -90, -834, -4982],
    "d_prepaid": [-1253, -3190, 1771, 2094],
    "d_ap": [84, 222, 2079, 2679],
    "d_lease": [-1, 0, 13, -3],
    "d_defrev": [7, 21, 241, 1380],
    "cfo": [-1627, -5692, -2637, 3929],
    "capex": [-2, -15, -131, -482],
    "sti_net": [0, 0, -2869, -16362],
    "other_inv": [0, -7, -77, -337],
    "cfi": [-2, -22, -3077, -17181],
    "pref_proceeds": [445, 1188, 16548, 110432],
    "conv_notes_cash": [1750, 6050, 3450, 0],
    "common_proceeds": [3, 10, 111, 355],
    "other_fin": [122, 0, -149, -13],   # SAFEs, repurchases, offering costs, finance lease principal
    "cff": [2320, 7248, 19960, 110774],
    "fx": [0, 0, 1, -5],
    "net_change_cash_filed": [691, 1534, 14247, 97517],
    "cash_end_filed": [1620, 3154, 17401, 114918],
    "sbc_is": [29, 192, 1165, 2185],     # IS SBC per non-GAAP reconciliation
    "headcount": [None, 940, 2700, 4300],  # Dec-24 filed; Dec-25 & Jun-26 are estimates (Sep-26 filed: ~5,300)
}
# Balance sheet: Dec-24A, Dec-25A, Jun-26A  (columns D, E, F)
B = {
    "cash": [3141, 17388, 112389],
    "sti": [0, 2889, 19262],
    "ar": [95, 909, 5833],
    "prepaid": [2076, 2560, 496],
    "restricted": [13, 13, 2529],
    "ppe": [19, 146, 1187],
    "rou": [35, 126, 1157],
    "goodwill": [0, 154, 576],
    "other_nca": [2373, 476, 3929],
    "total_assets": [7752, 24661, 147358],
    "ap": [333, 2511, 5422],
    "lease_liab": [52, 117, 1079],      # current + noncurrent operating lease liabilities
    "defrev": [28, 269, 1650],
    "conv_notes": [14116, 41677, 108566],
    "other_ncl": [150, 823, 750],
    "total_liab": [14679, 45397, 117467],
    "pref": [2874, 29769, 157510],
    "apic": [239, 1644, 5336],
    "aoci": [0, 5, -53],
    "deficit": [-10040, -52154, -132902],
    "total_equity": [-9801, -50505, -127619],
}
QUARTERS = [  # label, revenue, compute & infrastructure
    ("Q4'24A", 197, None), ("Q1'25A", 376, 944), ("Q2'25A", 787, 1390), ("Q3'25A", 1382, 2047),
    ("Q4'25A", 2040, 2945), ("Q1'26A", 4731, 3387), ("Q2'26A", 11555, 5558),
]
Q3_26_REV = 16800          # disclosed to investors 14 Oct 2026 (not in S-1 dollar form)
Q3_26_REV_PER_COMPUTE = 1.70   # disclosed to investors: $1.70 revenue per $1 of compute in Q3'26

# Compute commitments (estimated annual expense schedule, $M; see notes) -------- #
# (name, total $B, term, non_cancellable flag, [H2'26E, FY27E, FY28E, FY29E, FY30E], source note)
COMMITMENTS = [
    ("Google Cloud – TPUs / cloud (≥$114.5B, Apr-26 to Jul-33)", 114.5, 1, [4000, 10000, 16000, 20000, 22000],
     "True-up if spend falls short; non-cancellable except Google breach."),
    ("Amazon / AWS – Trainium & other (≥$110B, May-26 to Apr-36)", 110.0, 1, [3000, 7000, 10000, 13000, 15000],
     "Up to 5 GW; non-cancellable except Amazon breach."),
    ("Broadcom – TPU equipment finance leases (~$161.2B rent, 5-yr)", 161.2, 1, [0, 12000, 28000, 36000, 38000],
     "Finance leases (D&A + interest in GAAP); modeled here as compute expense for simplicity."),
    ("SpaceXAI – NVIDIA GPU capacity (up to $120.5B, ~$3.1B/mo at full)", 120.5, 0, [3000, 20000, 37000, 37000, 23000],
     "Cancellable on 90 days' notice after Jan/Mar 2027 – the main pressure valve."),
    ("Microsoft Azure – NVIDIA GPUs (≥$44.9B, Nov-26 to May-33)", 44.9, 1, [500, 5000, 7000, 8000, 8000],
     "Amended upward Oct-26 from ~$30B."),
    ("Nscale – NVIDIA GPUs (up to ~$44.6B, into 2033)", 44.6, 1, [500, 3000, 6000, 7000, 7000],
     "Paid as delivered; limited cancellation."),
    ("Lambda – NVIDIA GPUs (~$41.3B, into 2034)", 41.3, 1, [500, 3000, 5000, 6000, 7000],
     "Large tranche depends on a facility lease."),
    ("AMD – Instinct MI450 / Helios (>$20B, first tranche 2027)", 20.0, 1, [0, 2000, 4000, 5000, 5000],
     "AMD to buy up to $5B of Anthropic stock as chips deliver."),
]


# --------------------------------------------------------------------------- #
# Sheet builders
# --------------------------------------------------------------------------- #
def build_cover(m: Model):
    s = m.sheet("Cover")
    ws = s.ws
    ws.column_dimensions["A"].width = 110
    s.title("Anthropic, PBC – Three-Statement Financial Model",
            "Revenue, training vs. inference compute, income statement, balance sheet and cash flow | FY2023A–FY2030E | USD millions")
    lines = [
        ("Purpose", True),
        ("Integrated 3-statement model for Anthropic built off the Form S-1 (filed 26 Oct 2026, Nasdaq: ANT) and subsequent investor disclosures. "
         "The core analytical question is how revenue scales against the compute bill, split between inference (serving customers) and training/research.", False),
        ("", False),
        ("Structure", True),
        ("1. Summary – key outputs, charts and a valuation snapshot (fiscal-year view).", False),
        ("2. Assumptions – all forecast drivers. Scenario switch (cell C4): 1 = Base, 2 = Bull, 3 = Bear. Blue cells are inputs.", False),
        ("3. Revenue Build – total revenue (Q4'26 off the disclosed Q3'26 run-rate, then annual growth), consumption vs. subscription mix, marketplace vs. direct.", False),
        ("4. Compute Build – inference compute (paid, free-tier, internal) vs. training & research compute, with a take-or-pay floor from contracted minimums.", False),
        ("5. Opex Build – personnel (headcount × cash comp + SBC), channel-partner fees, marketing, other opex, D&A, and the functional (CoR / R&D / S&M / G&A) allocation.", False),
        ("6. Income Statement – GAAP functional presentation, memo by nature of expense, non-GAAP adjusted operating income / EBITDA.", False),
        ("7. Balance Sheet – rolls from the Jun-26A filed balance sheet; convertible notes and preferred stock convert to common at the IPO (H2'26E).", False),
        ("8. Cash Flow – indirect method; cash and short-term investments are modelled as one pool in the forecast.", False),
        ("9. Compute Commitments – estimated annual expense schedule for the contracted compute agreements vs. modelled demand.", False),
        ("10. Historical Data – filed statements (FY2023A–FY2025A audited, H1'25/H1'26 unaudited) and quarterly revenue/compute.", False),
        ("11. Checks – balance-sheet balance, cash tie-out, nature vs. function reconciliation.", False),
        ("", False),
        ("Period convention", True),
        ("Columns are aligned across all sheets: FY2023A | FY2024A | FY2025A | H1'26A | H2'26E (stub) | FY2026E (= H1A + H2E) | FY2027E–FY2030E. "
         "Flows in the H2'26E stub are half-year amounts; balance-sheet ratios annualise them via the period-fraction row on Assumptions.", False),
        ("", False),
        ("Colour legend", True),
        ("Blue font = hard-coded input or filed historical figure.  Black = formula.  Green = link to another sheet.  Yellow fill = live scenario-selected driver.", False),
        ("", False),
        ("Key simplifications", True),
        ("• All compute (including Broadcom TPU finance leases) is expensed as 'compute & infrastructure' in the period the capacity is used; GAAP would show part of this as D&A plus interest with a lease liability on the balance sheet.", False),
        ("• The training vs. inference split is not disclosed. Customer-serving inference is anchored to filed cost of revenue; training & research is the residual, and in the forecast it is driven by frontier runs × cost per run × a research multiplier.", False),
        ("• Convertible notes (Amazon) and the redeemable preferred stock convert into common equity at the IPO, assumed to close in H2'26E. The pre-conversion fair-value charge is an input and is non-cash.", False),
        ("• Marketplace revenue (AWS / Google Cloud / Azure) is recognised gross, with channel-partner fees in S&M, consistent with the S-1.", False),
        ("• Tax uses a simple NOL schedule; fair-value charges and the donation-matching expense are treated as non-deductible.", False),
        ("", False),
        ("Sources", True),
        ("Anthropic Form S-1 (SEC, 26 Oct 2026); Reuters / FT / CNBC reporting of the 14 Oct 2026 investor day (Q3'26 revenue $16.8B, adjusted operating margin ~11%, $1.70 revenue per $1 of compute); "
         "Bloomberg / CNBC (Jul-26 run-rate >$65B). Deal values for compute commitments as extracted from the S-1 by multiple outlets; the annual phasing is an estimate.", False),
        ("", False),
        ("This model is for research and illustration. Forecasts are the author's assumptions, not company guidance.", False),
    ]
    for text, bold in lines:
        c = ws[f"A{s.r}"]
        c.value = text
        c.font = SECTION if bold else BLACK
        c.alignment = Alignment(wrap_text=True, vertical="top")
        s.r += 1


def build_assumptions(m: Model):
    s = m.sheet("Assumptions")
    s.ws.column_dimensions["A"].width = 72
    s.title("Assumptions & Drivers", "Blue = input. Yellow rows are the live scenario-selected drivers used by the model.")
    # Scenario switch
    s.put(f"A{s.r}", "Scenario switch (1 = Base, 2 = Bull, 3 = Bear)", bold=True)
    s.m.rows[("Assumptions", "scenario")] = s.r
    s.put(f"C{s.r}", 1, INT, bold=True, fill=FILL_LIVE)
    s.put(f"D{s.r}", "=CHOOSE(C4,\"Base\",\"Bull\",\"Bear\")", bold=True)
    s.put(f"{NOTE_COL}{s.r}", "Change this cell to flip every scenario-driven input in the model.", font=GREY_I)
    s.r += 2
    s.period_header()

    s.section("Period")
    s.line("frac", "Period length (fraction of a year)", hist=[1, 1, 1, 0.5], fc=[0.5, 1, 1, 1, 1], h=1, fmt="0.0", unit="yrs",
           note="Used to annualise stub-period flows for balance-sheet ratios, interest and personnel cost.")

    # ------------------------------------------------------------------ Revenue
    s.section("Revenue")
    s.scenario_block("q4_qoq", "Q4'26 revenue growth vs. Q3'26 (QoQ)", base=[0.15], bull=[0.25], bear=[0.05],
                     fc_cols=["G"], note="Q3'26 revenue $16.8B (+45% QoQ) disclosed 14 Oct 2026. Run-rate ~$70–74B (Oct) implies Q4 ≈ $18–19B.")
    s.scenario_block("rev_growth", "Total revenue growth (YoY)", base=[0.95, 0.50, 0.32, 0.22],
                     bull=[1.30, 0.65, 0.40, 0.28], bear=[0.55, 0.28, 0.18, 0.12], fc_cols=["I", "J", "K", "L"],
                     note="Growth decelerates from +1,088% (FY25) as the base scales; S-1 warns growth will not continue at comparable rates.")
    s.line("sub_share", "Subscription revenue as % of total", fc=[0.20, 0.20, 0.21, 0.22, 0.22], fmt=PCT, unit="%",
           note="H1'26: 20% subscription / 80% consumption (filed). Consumer & team plans (Pro, Max, Team, Enterprise seats).")
    s.line("mkt_share", "Revenue via cloud marketplaces (AWS / GCP / Azure) as % of total", fc=[0.48, 0.47, 0.45, 0.43, 0.41], fmt=PCT, unit="%",
           note="H1'26: 48% (filed). Recognised gross; channel-partner fees sit in S&M.")

    # ------------------------------------------------------------ Inference
    s.section("Compute – inference (serving)")
    s.scenario_block("inf_cons_pct", "Paid inference compute as % of consumption revenue",
                     base=[0.10, 0.105, 0.10, 0.095, 0.09], bull=[0.095, 0.095, 0.09, 0.085, 0.08],
                     bear=[0.11, 0.125, 0.125, 0.12, 0.115],
                     note="≈ 1 – API serving gross margin. H1'26 implied ~10% (CoR $2.38B 'substantially all' serving compute). S-1 flags near-term "
                          "compression from higher-cost short-term capacity when demand exceeds contracted supply.")
    s.scenario_block("inf_sub_pct", "Paid inference compute as % of subscription revenue",
                     base=[0.30, 0.29, 0.28, 0.27, 0.26], bull=[0.28, 0.26, 0.24, 0.22, 0.20],
                     bear=[0.32, 0.33, 0.33, 0.32, 0.31],
                     note="Subscriptions are the less efficient slice of the inference pool (flat fee vs. heavy usage). SemiAnalysis est. ~50% compute margin on subs.")
    s.line("free_int_pct", "Free-tier + internal-use inference as % of paid inference compute", fc=[0.25] * 5, fmt=PCT, unit="%",
           note="Not disclosed; S-1 puts free-tier serving compute in S&M and internal Claude usage in R&D.")
    s.scalar("free_share", "   of which free-tier (S&M) share; remainder is internal usage (R&D)", 0.40, fmt=PCT, unit="%")

    # ------------------------------------------------------------- Training
    s.section("Compute – training & research")
    s.line("runs", "Frontier training runs completed in period", fc=[1, 2, 2, 2, 2], fmt="0", unit="runs",
           note="Roughly two frontier-model generations per year.")
    s.scenario_block("cost_per_run", "Compute cost per frontier training run", base=[5500, 7000, 10000, 13000, 16000],
                     bull=[6000, 8500, 12500, 16500, 20000], bear=[5000, 6000, 8000, 10000, 12000], fmt=NUM, unit="$M",
                     note="Calibrated so H2'26E total compute ≈ $21B (Q3'26 at $1.70 revenue/$ compute ≈ $9.9B, Q4 higher as new capacity lands).")
    s.line("research_mult", "Research, experimentation & post-training compute (multiple of frontier-run cost)",
           fc=[1.7, 1.6, 1.5, 1.4, 1.3], fmt=MULT, unit="x",
           note="Experiments, ablations, RL / post-training and safety research typically exceed the final pre-training run.")
    s.scalar("floor_toggle", "Floor compute expense at non-cancellable contracted minimums (1 = on, 0 = off)", 1, fmt="0", unit="flag",
             note="Take-or-pay: if contracted non-cancellable capacity exceeds modelled demand, the shortfall is expensed (shown in R&D as idle/research capacity).")

    # ------------------------------------------------------------ Personnel
    s.section("Personnel")
    s.line("headcount", "Year-end headcount", hist=D["headcount"], fc=[6000, 9000, 12000, 14500, 17000], h="same", fmt=INT, unit="FTE",
           note="Filed: ~940 (Dec-24), ~5,300 (Sep-26). Dec-25 and Jun-26 are estimates.")
    s.line("cash_comp", "Average cash compensation per employee (salary, bonus, benefits, payroll tax)",
           fc=[550, 550, 566, 583, 601], fmt=INT, unit="$K/yr",
           note="H1'26 implied ≈ $550K annualised ((personnel $3,152M – SBC $2,185M) / avg. headcount).")
    s.line("sbc_run", "Stock-based compensation – ongoing", fc=[2600, 9000, 10000, 11000, 12000], fmt=NUM, unit="$M",
           note="H1'26 SBC $2,185M. Unrecognised RSU expense $26.6B at Jun-26 (performance condition met at IPO) plus new grants.")
    s.line("sbc_ipo", "Stock-based compensation – IPO RSU catch-up (one-off)", fc=[8000, 0, 0, 0, 0], fmt=NUM, unit="$M",
           note="Cumulative catch-up on RSUs whose liquidity-event condition is satisfied on IPO effectiveness.")
    s.scalar("pers_cor", "Personnel allocation – cost of revenue", 0.03, fmt=PCT, unit="%")
    s.scalar("pers_rnd", "Personnel allocation – R&D", 0.73, fmt=PCT, unit="%")
    s.scalar("pers_snm", "Personnel allocation – S&M", 0.12, fmt=PCT, unit="%")
    s.scalar("pers_gna", "Personnel allocation – G&A", 0.12, fmt=PCT, unit="%",
             note="Allocations must sum to 100%. Majority of staff are research/engineering (S-1).")

    # -------------------------------------------------- Channel / marketing / other
    s.section("Channel partners, marketing & other operating costs")
    s.scenario_block("channel_fee", "Channel-partner & partnership costs as % of marketplace revenue",
                     base=[0.34, 0.33, 0.32, 0.31, 0.30], bull=[0.34, 0.32, 0.30, 0.28, 0.26], bear=[0.36] * 5,
                     note="H1'26 'sales, marketing & partnerships' $3,168M = 19.5% of revenue; net of ~2% advertising, implies ~36% of marketplace revenue. "
                          "Reuters: ~16c per marketplace $ in 2025, so partnership costs broadened in 2026.")
    s.line("mktg_pct", "Advertising & marketing as % of revenue", fc=[0.02] * 5, fmt=PCT, unit="%",
           note="FY25 advertising $147M (3.2% of revenue); brand spend scaling with consumer push.")
    s.line("other_pct", "Other operating & administrative as % of revenue (ex-donation matching, ex-D&A)",
           fc=[0.08, 0.065, 0.055, 0.05, 0.045], fmt=PCT, unit="%",
           note="H1'26: $1,588M ex-donation = 9.8% of revenue (data vendors, IT/cloud software, payment processing, facilities, legal, non-income taxes).")
    s.line("donation", "Donation-matching programme expense (non-cash, settled in Class N shares)",
           fc=[2000, 1500, 1500, 1500, 1500], fmt=NUM, unit="$M",
           note="Q3'26: $2.0B recognised (S-1 subsequent events). FY25: $541M; Q1'26: $126M.")
    s.scalar("oth_cor", "Other opex allocation – cost of revenue", 0.05, fmt=PCT, unit="%")
    s.scalar("oth_rnd", "Other opex allocation – R&D (data / eval vendors)", 0.50, fmt=PCT, unit="%")
    s.scalar("oth_snm", "Other opex allocation – S&M", 0.05, fmt=PCT, unit="%")
    s.scalar("oth_gna", "Other opex allocation – G&A", 0.40, fmt=PCT, unit="%", note="Allocations must sum to 100%.")

    # ------------------------------------------------------------ Capex / D&A
    s.section("Capex & D&A (owned assets)")
    s.line("capex_pct", "Capital expenditure as % of revenue", fc=[0.03, 0.03, 0.025, 0.025, 0.025], fmt=PCT, unit="%",
           note="H1'26 capex $482M = 3.0% of revenue. Most compute is rented, so owned PP&E is small.")
    s.scalar("da_rate", "D&A as % of (opening PP&E + ½ capex)", 0.35, fmt=PCT, unit="%",
             note="3–5 year useful lives; includes amortisation of acquired intangibles.")

    # -------------------------------------------------------- Working capital
    s.section("Working capital (balance as % of annualised period flow)")
    s.line("ar_pct", "Accounts receivable as % of annualised revenue", fc=[0.16, 0.14, 0.13, 0.125, 0.12], fmt=PCT, unit="%",
           note="Jun-26: $5,833M = 17.9% of annualised H1 revenue (inflated by in-period growth; ~46 days of Q2 revenue). Trends to ~45 DSO.")
    s.line("prepaid_pct", "Prepaid expenses & other current assets as % of annualised compute expense", fc=[0.05] * 5, fmt=PCT, unit="%",
           note="Largely prepaid cloud capacity. Jun-26: $496M; Dec-25: $2,560M (lumpy).")
    s.line("ap_pct", "Accounts payable & accrued as % of annualised cash operating costs", fc=[0.15] * 5, fmt=PCT, unit="%",
           note="Jun-26: $5,422M ≈ 18% of annualised H1 cash opex. Includes $1.2B Bartz settlement payable (instalments to Sep-27).")
    s.line("defrev_pct", "Deferred revenue as % of annualised subscription revenue", fc=[0.25] * 5, fmt=PCT, unit="%",
           note="Jun-26: $1,650M = 25.6% of annualised H1 subscription revenue (prepaid annual plans).")

    # ------------------------------------------------------------- Financing
    s.section("Financing, interest & tax")
    s.line("ipo_proceeds", "IPO primary proceeds, net of fees", fc=[60000, 0, 0, 0, 0], fmt=NUM, unit="$M",
           note="Press reports: raise could exceed $100B including secondary; primary portion assumed $60B. S-1 share count / price are blank.")
    s.line("fv_charge", "Fair-value charge on convertible notes & other instruments before conversion",
           fc=[50000, 0, 0, 0, 0], fmt=NUM, unit="$M",
           note="Amazon notes ($6.2B principal) carried at $108.6B at Jun-26 on a $965B mark; a ~$1.5T IPO mark implies a further non-cash charge. "
                "Notes convert to Class N common at IPO, after which the charge stops.")
    s.line("esop_proceeds", "Proceeds from option exercises / ESPP", fc=[150, 300, 300, 300, 300], fmt=NUM, unit="$M")
    s.line("debt_draw", "Debt / supplier-financing drawdown (repayment)", fc=[0, 0, 0, 0, 0], fmt=NUM, unit="$M",
           note="Broadcom offered up to $42B of 9% PIK convertible notes to fund TPU rent – not drawn in the base case.")
    s.scalar("debt_rate", "Interest rate on debt", 0.09, fmt=PCT, unit="%")
    s.line("yield", "Yield on cash & short-term investments", fc=[0.035, 0.0325, 0.03, 0.03, 0.03], fmt=PCT, unit="%",
           note="Applied to opening cash & STI balance (avoids circularity). H1'26 interest income $1,111M.")
    s.scalar("tax_rate", "Blended statutory tax rate (federal + state)", 0.23, fmt=PCT, unit="%")
    s.scalar("nol_open", "Net operating loss carryforward at Jun-26A (estimate)", 12000, fmt=NUM, unit="$M",
             note="Cumulative GAAP operating losses FY21–H1'26 ≈ $12B; fair-value charges are not deductible. 80% limitation ignored.")
    s.scalar("other_tax_pct", "Other / minimum taxes as % of revenue", 0.001, fmt=PCT, unit="%")
    return s


def build_historical(m: Model):
    s = m.sheet("Historical Data")
    s.title("Historical Data – as filed (Form S-1, 26 Oct 2026)",
            "USD millions. FY2023–FY2025 audited; six-month periods unaudited. Blue = filed figure or flagged estimate.")
    s.period_header()

    s.section("Statement of operations – by nature of expense (primary presentation in the S-1)")
    s.line("revenue", "Revenue", hist=D["revenue"], bold=True)
    s.line("rev_cons", "   Consumption-based (pay-for-intelligence)", hist=D["rev_cons"])
    s.line("rev_sub", "   Subscription-based", hist=D["rev_sub"])
    s.line("compute", "Compute and infrastructure", hist=D["compute"], note="Single fungible pool across training, research, serving and internal use.")
    s.line("personnel", "Personnel (incl. SBC)", hist=D["personnel"])
    s.line("bartz", "Bartz copyright settlement", hist=D["bartz"], note="$1.5B settled 2025; $300M paid in 2025, $1.2B payable through Sep-27.")
    s.line("smp", "Sales, marketing and partnerships", hist=D["smp"], note="Channel-partner costs (AWS/GCP/Azure marketplaces) + advertising.")
    s.line("other_oa", "Other operating and administrative", hist=D["other_oa"], note="Incl. donation-matching expense (FY25 $541M; H1'26 $126M).")
    s.line("opex_total", "Total operating expenses", hist=D["opex_total"], bold=True, border=TOP)
    s.line("op_income", "Operating income (loss)", hist=D["op_income"], bold=True)
    s.line("interest_income", "Interest income", hist=D["interest_income"])
    s.line("fv_change", "Change in FV of convertible notes, derivatives, SAFEs & other financial instruments", hist=D["fv_change"],
           note="H1'26 = $(72,388)M on notes/SAFEs + $(8,720)M recognition of 2026 note purchase agreements (reported within 'other').")
    s.line("other_inc", "Other income (expense), net", hist=D["other_inc"], note="H1'26 residual after moving the $8,720M to the FV line.")
    s.line("pretax", "Income (loss) before income taxes", hist=D["pretax"], bold=True, border=TOP)
    s.line("tax", "Provision for income taxes", hist=D["tax"])
    s.line("net_income", "Net income (loss)", hist=D["net_income"], bold=True, border=TOPBOT)

    s.section("Statement of operations – by function (S-1 supplemental note)")
    s.line("cor", "Cost of revenue", hist=D["cor"], note="'Substantially all' compute for customer-serving workloads + some personnel/overhead.")
    s.line("rnd", "Research and development", hist=D["rnd"], note="Primarily compute for model development, research and internal usage.")
    s.line("snm", "Sales and marketing", hist=D["snm"], note="Channel-partner costs, marketing, personnel, free-tier compute.")
    s.line("gna", "General and administrative", hist=D["gna"], note="Donation matching, personnel, facilities, legal, non-income taxes.")
    s.line("func_total", "Functional total + Bartz settlement", hist=lambda c: f"={c}{m.rows[('Historical Data','cor')]}+{c}{m.rows[('Historical Data','rnd')]}+{c}{m.rows[('Historical Data','snm')]}+{c}{m.rows[('Historical Data','gna')]}+{c}{m.rows[('Historical Data','bartz')]}",
           bold=True, border=TOP, note="Ties to total operating expenses above.")

    s.section("Cash flow statement (as filed)")
    s.line("sbc", "Stock-based compensation (cash-flow add-back)", hist=D["sbc"])
    s.line("sbc_is", "Stock-based compensation (income statement, per non-GAAP reconciliation)", hist=D["sbc_is"],
           note="FY25 IS expense $1,165M vs. $1,158M cash-flow add-back (small capitalised amount).")
    s.line("da", "Depreciation and amortization", hist=D["da"])
    s.line("donation", "Donation matching programme liability (non-cash)", hist=D["donation"])
    s.line("cf_other", "Other non-cash items", hist=D["cf_other"])
    s.line("d_ar", "Change in accounts receivable", hist=D["d_ar"])
    s.line("d_prepaid", "Change in prepaid expenses and other assets", hist=D["d_prepaid"])
    s.line("d_ap", "Change in accounts payable, accrued and other liabilities", hist=D["d_ap"])
    s.line("d_lease", "Change in lease liabilities", hist=D["d_lease"])
    s.line("d_defrev", "Change in deferred revenue", hist=D["d_defrev"])
    s.line("cfo", "Net cash from operating activities", hist=D["cfo"], bold=True, border=TOP)
    s.line("capex", "Purchases of property and equipment", hist=D["capex"])
    s.line("sti_net", "Purchases of short-term investments, net of sales/maturities", hist=D["sti_net"])
    s.line("other_inv", "Other investing activities", hist=D["other_inv"])
    s.line("cfi", "Net cash used in investing activities", hist=D["cfi"], bold=True, border=TOP)
    s.line("pref_proceeds", "Proceeds from redeemable convertible preferred stock, net", hist=D["pref_proceeds"])
    s.line("conv_notes_cash", "Issuance of convertible notes for cash", hist=D["conv_notes_cash"])
    s.line("common_proceeds", "Proceeds from issuance of common stock", hist=D["common_proceeds"])
    s.line("other_fin", "Other financing (SAFEs, repurchases, offering costs, finance-lease principal)", hist=D["other_fin"])
    s.line("cff", "Net cash from financing activities", hist=D["cff"], bold=True, border=TOP)
    s.line("fx", "Effect of exchange rates", hist=D["fx"])
    s.line("net_change_cash_filed", "Net increase in cash, cash equivalents and restricted cash", hist=D["net_change_cash_filed"], bold=True)
    s.line("cash_end_filed", "Cash, cash equivalents and restricted cash – end of period", hist=D["cash_end_filed"])

    s.section("Balance sheet (as filed) – column aligns with period end")
    bs_cols = ["D", "E", "F"]
    for key, label in [
        ("cash", "Cash and cash equivalents"), ("sti", "Short-term investments"), ("ar", "Accounts receivable, net"),
        ("prepaid", "Prepaid expenses and other current assets"), ("restricted", "Restricted cash"),
        ("ppe", "Property and equipment, net"), ("rou", "Operating lease right-of-use assets"), ("goodwill", "Goodwill"),
        ("other_nca", "Other assets, noncurrent"), ("total_assets", "Total assets"),
        ("ap", "Accounts payable, accrued expenses and other current liabilities"),
        ("lease_liab", "Operating lease liabilities (current + noncurrent)"), ("defrev", "Deferred revenue"),
        ("conv_notes", "Convertible notes payable (at fair value)"), ("other_ncl", "Other noncurrent liabilities"),
        ("total_liab", "Total liabilities"), ("pref", "Redeemable convertible preferred stock (mezzanine)"),
        ("apic", "Additional paid-in capital (incl. common stock at par)"), ("aoci", "Accumulated other comprehensive income (loss)"),
        ("deficit", "Accumulated deficit"), ("total_equity", "Total stockholders' equity (deficit)"),
    ]:
        bold = key in ("total_assets", "total_liab", "total_equity")
        s.line(f"bs_{key}", label, hist=B[key], hist_cols=bs_cols, bold=bold, border=TOP if bold else None)
    s.note("Balance-sheet columns: D = Dec-24A, E = Dec-25A, F = Jun-26A. Dec-23 balance sheet not presented in the S-1 in full.")

    s.section("Quarterly revenue and compute (S-1 narrative; Q3'26 from 14 Oct 2026 investor day)")
    # quarterly table placed in columns C.. with labels in row
    qrow = s.r
    s.put(f"A{qrow}", "Quarter", bold=True)
    qcols = COLS[: len(QUARTERS) + 1]
    for c, (lab, _, _) in zip(qcols, QUARTERS):
        s.put(f"{c}{qrow}", lab, bold=True, align="center")
    s.put(f"{qcols[-1]}{qrow}", "Q3'26", bold=True, align="center")
    s.r += 1
    m.rows[("Historical Data", "q_rev")] = s.r
    s.put(f"A{s.r}", "Revenue")
    s.put(f"B{s.r}", "$M", font=GREY_I, align="center")
    for c, (_, rev, _) in zip(qcols, QUARTERS):
        s.put(f"{c}{s.r}", rev, NUM)
    s.put(f"{qcols[-1]}{s.r}", Q3_26_REV, NUM)
    s.put(f"{NOTE_COL}{s.r}", "Q3'26 $16.8B disclosed to investors (S-1 states it as 86× Q4'24 revenue ≈ $16.9B).", font=GREY_I)
    s.r += 1
    m.rows[("Historical Data", "q_compute")] = s.r
    s.put(f"A{s.r}", "Compute and infrastructure")
    s.put(f"B{s.r}", "$M", font=GREY_I, align="center")
    for c, (_, _, comp) in zip(qcols, QUARTERS):
        if comp is not None:
            s.put(f"{c}{s.r}", comp, NUM)
    s.put(f"{qcols[-1]}{s.r}", f"={qcols[-1]}{m.rows[('Historical Data','q_rev')]}/{qcols[-1]}{s.r + 1}", NUM)
    s.put(f"{NOTE_COL}{s.r}", "Q3'26 compute derived from the disclosed $1.70 of revenue per $1 of compute.", font=GREY_I)
    s.r += 1
    m.rows[("Historical Data", "q_rev_per_compute")] = s.r
    s.put(f"A{s.r}", "Revenue per $1 of compute")
    s.put(f"B{s.r}", "x", font=GREY_I, align="center")
    for c, (_, rev, comp) in zip(qcols, QUARTERS):
        if comp is not None:
            s.put(f"{c}{s.r}", f"={c}{m.rows[('Historical Data','q_rev')]}/{c}{m.rows[('Historical Data','q_compute')]}", MULT)
    s.put(f"{qcols[-1]}{s.r}", Q3_26_REV_PER_COMPUTE, MULT)
    s.r += 1
    s.q3_col = qcols[-1]

    s.section("Other disclosed data points")
    s.scalar("rev_h1_25", "Revenue, six months ended 30 Jun 2025 (unaudited)", 1163, unit="$M",
             note="Implies H2'25 revenue of $3,422M – the comparison base for H2'26E growth.")
    s.scalar("runrate_jul26", "Run-rate revenue, end-July 2026 (company to investors)", 65000, unit="$M",
             note="Run-rate = trailing 28-day consumption fees × 13 + monthly subscription fees × 12.")
    s.scalar("runrate_dec25", "Run-rate revenue, end-2025", 9000, unit="$M")
    s.scalar("adj_op_margin_q3", "Adjusted operating margin, Q3'26 (company to investors)", 0.11, fmt=PCT, unit="%")
    s.scalar("commit_min", "Fixed / minimum compute purchase commitments at Jun-26A", 265600, unit="$M", note="Terms 1–20 years; up from $54.6B at Dec-25.")
    s.scalar("commit_leases", "Leases signed but not yet commenced (undiscounted) at Jun-26A", 321400, unit="$M", note="Primarily compute; commencement 2026–2028.")
    s.scalar("capital_raised", "Aggregate capital raised through Jun-26A", 141100, unit="$M")
    s.scalar("unrec_sbc", "Unrecognised SBC on unvested RSUs at Jun-26A", 26600, unit="$M")
    return s


def build_revenue(m: Model):
    s = m.sheet("Revenue Build")
    s.title("Revenue Build", "Q4'26 is grown off the disclosed Q3'26; FY2027E+ use annual growth rates. Mix splits feed the compute and channel-cost builds.")
    s.period_header()
    H = "Historical Data"
    q3 = m.wb["Historical Data"]  # noqa: F841 (sheet exists)
    q3col = m.q3_col

    s.section("Quarterly bridge for the H2'26E stub")
    s.line("q3_rev", "Q3'26A revenue (disclosed)", fc=[f"={{{H}!q_rev@{q3col}}}", None, None, None, None], fc_cols=["G"],
           note="Reuters / FT / CNBC, 14 Oct 2026.")
    s.line("q4_rev", "Q4'26E revenue", fc=[f"=G{m.rows[('Revenue Build','q3_rev')]}*(1+{{Assumptions!q4_qoq@G}})", None, None, None, None], fc_cols=["G"])
    s.line("exit_rr", "Implied Q4'26E exit run-rate (Q4 × 4)", fc=[f"=G{m.rows[('Revenue Build','q4_rev')]}*4", None, None, None, None], fc_cols=["G"],
           note="Compare: ~$65B at end-July 2026; third-party estimates ~$70–74B for October.")

    s.section("Total revenue")
    s.line("revenue", "Total revenue", hist=lambda c: f"={{{H}!revenue@{c}}}",
           fc=lambda c, p: (f"=G{m.rows[('Revenue Build','q3_rev')]}+G{m.rows[('Revenue Build','q4_rev')]}" if c == "G"
                            else f"={p}{m.rows[('Revenue Build','revenue')]}*(1+{{Assumptions!rev_growth@{c}}})"),
           h="sum", bold=True)
    rev = m.rows[("Revenue Build", "revenue")]
    s.line("rev_growth", "   YoY growth", hist=[None, f"=D{rev}/C{rev}-1", f"=E{rev}/D{rev}-1", None],
           fc=lambda c, p: (f"=G{rev}/({{{H}!revenue@E}}-{{{H}!rev_h1_25}})-1" if c == "G" else f"={c}{rev}/{p}{rev}-1"),
           h=f"=H{rev}/E{rev}-1", fmt=PCT, unit="%", note="H2'26E growth is vs. H2'25A (FY25 less H1'25 = $3,422M).")
    s.line("rev_hoh", "   Sequential growth (H2 vs. H1)", fc=[f"=G{rev}/F{rev}-1", None, None, None, None], fc_cols=["G"], fmt=PCT, unit="%")

    s.section("Mix – consumption vs. subscription")
    s.line("rev_cons", "Consumption-based revenue (API, Claude Code, agents)", hist=lambda c: f"={{{H}!rev_cons@{c}}}",
           fc=lambda c, p: f"={c}{rev}*(1-{{Assumptions!sub_share@{c}}})", h="sum")
    s.line("rev_sub", "Subscription-based revenue (Pro, Max, Team, Enterprise seats)", hist=lambda c: f"={{{H}!rev_sub@{c}}}",
           fc=lambda c, p: f"={c}{rev}*{{Assumptions!sub_share@{c}}}", h="sum")
    rc, rs = m.rows[("Revenue Build", "rev_cons")], m.rows[("Revenue Build", "rev_sub")]
    s.line("sub_mix", "   Subscription share of revenue", hist=lambda c: f"={c}{rs}/{c}{rev}",
           fc=lambda c, p: f"={c}{rs}/{c}{rev}", h=f"=H{rs}/H{rev}", fmt=PCT, unit="%")
    s.line("cons_growth", "   Consumption revenue growth (YoY)", hist=[None, f"=D{rc}/C{rc}-1", f"=E{rc}/D{rc}-1", None],
           fc=lambda c, p: (None if c == "G" else f"={c}{rc}/{p}{rc}-1"), h=f"=H{rc}/E{rc}-1", fmt=PCT, unit="%", fc_cols=["I", "J", "K", "L"])
    s.line("sub_growth", "   Subscription revenue growth (YoY)", hist=[None, f"=D{rs}/C{rs}-1", f"=E{rs}/D{rs}-1", None],
           fc=lambda c, p: (None if c == "G" else f"={c}{rs}/{p}{rs}-1"), h=f"=H{rs}/E{rs}-1", fmt=PCT, unit="%", fc_cols=["I", "J", "K", "L"])

    s.section("Mix – channel")
    s.line("rev_mkt", "Revenue through cloud marketplaces (AWS Bedrock, Vertex AI, Azure Foundry)",
           hist=[None, None, None, f"=F{rev}*0.48"], fc=lambda c, p: f"={c}{rev}*{{Assumptions!mkt_share@{c}}}", h="sum",
           note="H1'26: 48% of revenue (filed). Recognised gross; partner fees in S&M.")
    rm = m.rows[("Revenue Build", "rev_mkt")]
    s.line("rev_direct", "Direct revenue (first-party API, apps, subscriptions)", hist=[None, None, None, f"=F{rev}-F{rm}"],
           fc=lambda c, p: f"={c}{rev}-{c}{rm}", h="sum")

    s.section("Memo")
    s.line("rev_per_emp", "Revenue per average employee (annualised)",
           hist=[None, None, f"=E{rev}/AVERAGE({{Assumptions!headcount@D}},{{Assumptions!headcount@E}})",
                 f"=F{rev}/{{Assumptions!frac@F}}/AVERAGE({{Assumptions!headcount@E}},{{Assumptions!headcount@F}})"],
           fc=lambda c, p: f"={c}{rev}/{{Assumptions!frac@{c}}}/AVERAGE({{Assumptions!headcount@{p}}},{{Assumptions!headcount@{c}}})",
           h=f"=H{rev}/AVERAGE({{Assumptions!headcount@E}},{{Assumptions!headcount@H}})", fmt=NUM1, unit="$M/FTE",
           note="S-1: TTM revenue per employee > $6.6M at Sep-26.")
    return s


def build_compute(m: Model):
    s = m.sheet("Compute Build")
    s.title("Compute Build – Inference vs. Training",
            "Inference is driven by revenue × cost ratios; training & research by frontier runs × cost per run × research multiple. "
            "Historical split is an estimate anchored to filed cost of revenue (S-1 discloses only total compute).")
    s.period_header()
    R, H, A = "Revenue Build", "Historical Data", "Assumptions"

    s.section("Inference compute – customer-serving (cost of revenue)")
    s.line("inf_cons", "Paid inference – consumption revenue",
           fc=lambda c, p: f"={{{R}!rev_cons@{c}}}*{{{A}!inf_cons_pct@{c}}}", h="sum")
    s.line("inf_sub", "Paid inference – subscription revenue",
           fc=lambda c, p: f"={{{R}!rev_sub@{c}}}*{{{A}!inf_sub_pct@{c}}}", h="sum")
    ic, isub = m.rows[("Compute Build", "inf_cons")], m.rows[("Compute Build", "inf_sub")]
    s.line("inf_paid", "Total paid inference compute (in cost of revenue)",
           hist=lambda c: f"={{{H}!cor@{c}}}-{{{A}!pers_cor}}*{{{H}!personnel@{c}}}-{{{A}!oth_cor}}*{{{H}!other_oa@{c}}}",
           fc=lambda c, p: f"={c}{ic}+{c}{isub}", h="sum", bold=True, border=TOP,
           note="Historical: cost of revenue less allocated personnel/other ('substantially all' CoR is serving compute).")
    ip = m.rows[("Compute Build", "inf_paid")]
    s.line("inf_free", "Free-tier inference (S&M)",
           hist=lambda c: f"={c}{ip}*{{{A}!free_int_pct@G}}*{{{A}!free_share}}",
           fc=lambda c, p: f"={c}{ip}*{{{A}!free_int_pct@{c}}}*{{{A}!free_share}}", h="sum",
           note="Historical uses the H2'26E ratio – estimate.")
    s.line("inf_int", "Internal-usage inference (R&D)",
           hist=lambda c: f"={c}{ip}*{{{A}!free_int_pct@G}}*(1-{{{A}!free_share}})",
           fc=lambda c, p: f"={c}{ip}*{{{A}!free_int_pct@{c}}}*(1-{{{A}!free_share}})", h="sum")
    ifr, iin = m.rows[("Compute Build", "inf_free")], m.rows[("Compute Build", "inf_int")]
    s.line("inf_total", "Total inference compute", hist=lambda c: f"={c}{ip}+{c}{ifr}+{c}{iin}",
           fc=lambda c, p: f"={c}{ip}+{c}{ifr}+{c}{iin}", h="sum", bold=True, border=TOP)
    it = m.rows[("Compute Build", "inf_total")]

    s.section("Training & research compute (R&D)")
    s.line("runs", "Frontier training runs", fc=lambda c, p: f"={{{A}!runs@{c}}}", h="sum", fmt="0", unit="runs")
    s.line("cost_run", "Compute cost per frontier run", fc=lambda c, p: f"={{{A}!cost_per_run@{c}}}", fmt=NUM)
    rr, cr = m.rows[("Compute Build", "runs")], m.rows[("Compute Build", "cost_run")]
    s.line("train_frontier", "Frontier pre-training compute", fc=lambda c, p: f"={c}{rr}*{c}{cr}", h="sum")
    tf = m.rows[("Compute Build", "train_frontier")]
    s.line("train_research", "Research, experimentation & post-training compute",
           fc=lambda c, p: f"={c}{tf}*{{{A}!research_mult@{c}}}", h="sum")
    tr = m.rows[("Compute Build", "train_research")]
    s.line("train_total", "Total training & research compute",
           hist=lambda c: f"={{{H}!compute@{c}}}-{c}{it}",
           fc=lambda c, p: f"={c}{tf}+{c}{tr}", h="sum", bold=True, border=TOP,
           note="Historical: residual of filed total compute less estimated inference.")
    tt = m.rows[("Compute Build", "train_total")]

    s.section("Contracted capacity and total compute expense")
    s.line("demand", "Driver-based compute demand (inference + training)", hist=lambda c: f"={c}{it}+{c}{tt}",
           fc=lambda c, p: f"={c}{it}+{c}{tt}", h="sum")
    dm = m.rows[("Compute Build", "demand")]
    s.line("contracted", "Estimated non-cancellable contracted minimums coming on-line",
           fc=lambda c, p: f"={{Compute Commitments!noncancel_total@{c}}}", h="sum",
           note="From the Compute Commitments sheet (estimated phasing of S-1 deal terms).")
    ct = m.rows[("Compute Build", "contracted")]
    s.line("excess", "Take-or-pay shortfall: contracted capacity above modelled demand (expensed in R&D)",
           fc=lambda c, p: f"=IF({{{A}!floor_toggle}}=1,MAX(0,{c}{ct}-{c}{dm}),0)", h="sum",
           note="Zero when demand ≥ contracted minimums. Capacity is expensed when available, whether or not customers use it.")
    ex = m.rows[("Compute Build", "excess")]
    s.line("compute_total", "Total compute & infrastructure expense", hist=lambda c: f"={{{H}!compute@{c}}}",
           fc=lambda c, p: f"={c}{dm}+{c}{ex}", h="sum", bold=True, border=TOPBOT)
    tot = m.rows[("Compute Build", "compute_total")]

    s.section("Ratios")
    rev = f"{{{R}!revenue@{{c}}}}"
    s.line("compute_pct_rev", "Compute & infrastructure as % of revenue", hist=lambda c: f"={c}{tot}/{{{R}!revenue@{c}}}",
           fc=lambda c, p: f"={c}{tot}/{{{R}!revenue@{c}}}", h=f"=H{tot}/{{{R}!revenue@H}}", fmt=PCT, unit="%",
           note="FY25 160%; H1'26 55%; Q2'26 48%; Q3'26 ~59% ($1.70 revenue per $1).")
    s.line("rev_per_compute", "Revenue per $1 of compute", hist=lambda c: f"={{{R}!revenue@{c}}}/{c}{tot}",
           fc=lambda c, p: f"={{{R}!revenue@{c}}}/{c}{tot}", h=f"={{{R}!revenue@H}}/H{tot}", fmt=MULT, unit="x")
    s.line("inf_pct_rev", "Paid inference compute as % of revenue (1 – serving gross margin)", hist=lambda c: f"={c}{ip}/{{{R}!revenue@{c}}}",
           fc=lambda c, p: f"={c}{ip}/{{{R}!revenue@{c}}}", h=f"=H{ip}/{{{R}!revenue@H}}", fmt=PCT, unit="%")
    s.line("train_share", "Training & research share of total compute", hist=lambda c: f"={c}{tt}/{c}{tot}",
           fc=lambda c, p: f"=({c}{tt}+{c}{ex})/{c}{tot}", h=f"=(H{tt}+H{ex})/H{tot}", fmt=PCT, unit="%")
    s.line("inf_share", "Inference share of total compute", hist=lambda c: f"={c}{it}/{c}{tot}",
           fc=lambda c, p: f"={c}{it}/{c}{tot}", h=f"=H{it}/H{tot}", fmt=PCT, unit="%")
    s.line("train_pct_rev", "Training & research compute as % of revenue", hist=lambda c: f"={c}{tt}/{{{R}!revenue@{c}}}",
           fc=lambda c, p: f"=({c}{tt}+{c}{ex})/{{{R}!revenue@{c}}}", h=f"=(H{tt}+H{ex})/{{{R}!revenue@H}}", fmt=PCT, unit="%")
    s.line("compute_growth", "Total compute growth (YoY)", hist=[None, f"=D{tot}/C{tot}-1", f"=E{tot}/D{tot}-1", None],
           fc=lambda c, p: (None if c == "G" else f"={c}{tot}/{p}{tot}-1"), h=f"=H{tot}/E{tot}-1", fmt=PCT, unit="%", fc_cols=["I", "J", "K", "L"])
    return s


def build_opex(m: Model):
    s = m.sheet("Opex Build")
    s.title("Operating Expense Build", "Personnel, channel-partner fees, marketing, other operating costs and D&A, then allocated to GAAP functional lines.")
    s.period_header()
    R, H, A, C = "Revenue Build", "Historical Data", "Assumptions", "Compute Build"

    s.section("Personnel")
    s.line("hc_ye", "Year-end headcount", hist=lambda c: f"={{{A}!headcount@{c}}}", fc=lambda c, p: f"={{{A}!headcount@{c}}}", h="same", fmt=INT, unit="FTE",
           hist_cols=["D", "E", "F"])
    hc = m.rows[("Opex Build", "hc_ye")]
    s.line("hc_avg", "Average headcount", hist=[None, None, f"=AVERAGE(D{hc},E{hc})", f"=AVERAGE(E{hc},F{hc})"],
           fc=lambda c, p: f"=AVERAGE({p}{hc},{c}{hc})", h=f"=AVERAGE(E{hc},H{hc})", fmt=INT, unit="FTE")
    ha = m.rows[("Opex Build", "hc_avg")]
    s.line("sbc", "Stock-based compensation", hist=lambda c: f"={{{H}!sbc_is@{c}}}",
           fc=lambda c, p: f"={{{A}!sbc_run@{c}}}+{{{A}!sbc_ipo@{c}}}", h="sum")
    sbc = m.rows[("Opex Build", "sbc")]
    s.line("pers_cash", "Cash personnel cost (salary, bonus, benefits, payroll tax)", hist=lambda c: f"={{{H}!personnel@{c}}}-{c}{sbc}",
           fc=lambda c, p: f"={c}{ha}*{{{A}!cash_comp@{c}}}/1000*{{{A}!frac@{c}}}", h="sum")
    pc = m.rows[("Opex Build", "pers_cash")]
    s.line("pers_comp_k", "   Implied cash comp per average employee", hist=[None, None, f"=E{pc}/E{ha}*1000", f"=F{pc}/F{ha}*1000/{{{A}!frac@F}}"],
           fc=lambda c, p: f"={c}{pc}/{c}{ha}*1000/{{{A}!frac@{c}}}", h=f"=H{pc}/H{ha}*1000", fmt=INT, unit="$K/yr")
    s.line("personnel", "Total personnel expense", hist=lambda c: f"={{{H}!personnel@{c}}}", fc=lambda c, p: f"={c}{pc}+{c}{sbc}", h="sum", bold=True, border=TOP)
    pers = m.rows[("Opex Build", "personnel")]

    s.section("Sales, marketing & partnerships")
    s.line("channel", "Channel-partner & partnership costs (marketplace fees)",
           fc=lambda c, p: f"={{{R}!rev_mkt@{c}}}*{{{A}!channel_fee@{c}}}", h="sum")
    ch = m.rows[("Opex Build", "channel")]
    s.line("mktg", "Advertising & marketing", fc=lambda c, p: f"={{{R}!revenue@{c}}}*{{{A}!mktg_pct@{c}}}", h="sum")
    mk = m.rows[("Opex Build", "mktg")]
    s.line("smp", "Total sales, marketing & partnerships", hist=lambda c: f"={{{H}!smp@{c}}}", fc=lambda c, p: f"={c}{ch}+{c}{mk}", h="sum", bold=True, border=TOP)
    smp = m.rows[("Opex Build", "smp")]
    s.line("smp_pct", "   as % of revenue", hist=lambda c: f"={c}{smp}/{{{R}!revenue@{c}}}", fc=lambda c, p: f"={c}{smp}/{{{R}!revenue@{c}}}",
           h=f"=H{smp}/{{{R}!revenue@H}}", fmt=PCT, unit="%")

    s.section("Other operating & administrative")
    s.line("other_core", "Other operating & administrative (ex-donation, ex-D&A)",
           hist=lambda c: f"={{{H}!other_oa@{c}}}-{{{H}!donation@{c}}}-{{{H}!da@{c}}}",
           fc=lambda c, p: f"={{{R}!revenue@{c}}}*{{{A}!other_pct@{c}}}", h="sum")
    oc = m.rows[("Opex Build", "other_core")]
    s.line("donation", "Donation-matching programme expense (non-cash)", hist=lambda c: f"={{{H}!donation@{c}}}",
           fc=lambda c, p: f"={{{A}!donation@{c}}}", h="sum")
    dn = m.rows[("Opex Build", "donation")]
    s.line("capex", "Capital expenditure (owned PP&E)", hist=lambda c: f"=-{{{H}!capex@{c}}}",
           fc=lambda c, p: f"={{{R}!revenue@{c}}}*{{{A}!capex_pct@{c}}}", h="sum")
    cx = m.rows[("Opex Build", "capex")]
    s.line("da", "Depreciation & amortisation", hist=lambda c: f"={{{H}!da@{c}}}",
           fc=lambda c, p: f"=({{Balance Sheet!ppe@{p}}}+0.5*{c}{cx})*{{{A}!da_rate}}*{{{A}!frac@{c}}}", h="sum")
    da = m.rows[("Opex Build", "da")]
    s.line("other_oa", "Total other operating & administrative", hist=lambda c: f"={{{H}!other_oa@{c}}}",
           fc=lambda c, p: f"={c}{oc}+{c}{dn}+{c}{da}", h="sum", bold=True, border=TOP)
    ooa = m.rows[("Opex Build", "other_oa")]

    s.section("Functional allocation (forecast) – ties to total expenses by nature")
    s.line("f_cor", "Cost of revenue", hist=lambda c: f"={{{H}!cor@{c}}}",
           fc=lambda c, p: f"={{{C}!inf_paid@{c}}}+{{{A}!pers_cor}}*{c}{pers}+{{{A}!oth_cor}}*{c}{oc}", h="sum",
           note="Paid inference compute + allocated personnel and other.")
    s.line("f_rnd", "Research & development", hist=lambda c: f"={{{H}!rnd@{c}}}",
           fc=lambda c, p: f"={{{C}!train_total@{c}}}+{{{C}!inf_int@{c}}}+{{{C}!excess@{c}}}+{{{A}!pers_rnd}}*{c}{pers}+{{{A}!oth_rnd}}*{c}{oc}", h="sum",
           note="Training & research compute + internal inference + take-or-pay shortfall + allocated personnel/other.")
    s.line("f_snm", "Sales & marketing", hist=lambda c: f"={{{H}!snm@{c}}}",
           fc=lambda c, p: f"={{{C}!inf_free@{c}}}+{c}{smp}+{{{A}!pers_snm}}*{c}{pers}+{{{A}!oth_snm}}*{c}{oc}", h="sum",
           note="Free-tier inference + channel & marketing + allocated personnel/other.")
    s.line("f_gna", "General & administrative", hist=lambda c: f"={{{H}!gna@{c}}}",
           fc=lambda c, p: f"={{{A}!pers_gna}}*{c}{pers}+{{{A}!oth_gna}}*{c}{oc}+{c}{dn}+{c}{da}", h="sum",
           note="Allocated personnel/other + donation matching + D&A.")
    s.line("f_bartz", "Bartz settlement", hist=lambda c: f"={{{H}!bartz@{c}}}", fc=[0, 0, 0, 0, 0], h="sum")
    fc_, fr, fs, fg, fb = (m.rows[("Opex Build", k)] for k in ("f_cor", "f_rnd", "f_snm", "f_gna", "f_bartz"))
    s.line("f_total", "Total operating expenses (functional)", hist=lambda c: f"={c}{fc_}+{c}{fr}+{c}{fs}+{c}{fg}+{c}{fb}",
           fc=lambda c, p: f"={c}{fc_}+{c}{fr}+{c}{fs}+{c}{fg}+{c}{fb}", h="sum", bold=True, border=TOP)
    ft = m.rows[("Opex Build", "f_total")]
    s.line("n_total", "Total operating expenses (by nature: compute + personnel + S,M&P + other + Bartz)",
           hist=lambda c: f"={{{C}!compute_total@{c}}}+{c}{pers}+{c}{smp}+{c}{ooa}+{c}{fb}",
           fc=lambda c, p: f"={{{C}!compute_total@{c}}}+{c}{pers}+{c}{smp}+{c}{ooa}+{c}{fb}", h="sum", bold=True)
    nt = m.rows[("Opex Build", "n_total")]
    s.line("alloc_check", "Check: functional – nature (should be 0)", hist=lambda c: f"={c}{ft}-{c}{nt}", fc=lambda c, p: f"={c}{ft}-{c}{nt}",
           h=f"=H{ft}-H{nt}", note="Historical: the S-1's functional note ties to the nature presentation after adding the Bartz line.")
    return s


def build_is(m: Model):
    s = m.sheet("Income Statement")
    s.title("Income Statement", "GAAP functional presentation (as in the S-1 supplemental note) with memo by nature of expense and non-GAAP adjusted metrics.")
    s.period_header()
    R, H, A, C, O = "Revenue Build", "Historical Data", "Assumptions", "Compute Build", "Opex Build"

    s.section("GAAP income statement")
    s.line("revenue", "Revenue", hist=lambda c: f"={{{R}!revenue@{c}}}", fc=lambda c, p: f"={{{R}!revenue@{c}}}", h="sum", bold=True)
    rev = m.rows[("Income Statement", "revenue")]
    s.line("growth", "   YoY growth", hist=lambda c: f"={{{R}!rev_growth@{c}}}", fc=lambda c, p: f"={{{R}!rev_growth@{c}}}", h=f"={{{R}!rev_growth@H}}",
           fmt=PCT, unit="%", hist_cols=["D", "E"])
    s.line("cor", "Cost of revenue", hist=lambda c: f"={{{O}!f_cor@{c}}}", fc=lambda c, p: f"={{{O}!f_cor@{c}}}", h="sum")
    cor = m.rows[("Income Statement", "cor")]
    s.line("gp", "Gross profit", hist=lambda c: f"={c}{rev}-{c}{cor}", fc=lambda c, p: f"={c}{rev}-{c}{cor}", h=f"=H{rev}-H{cor}", bold=True, border=TOP)
    gp = m.rows[("Income Statement", "gp")]
    s.line("gm", "   Gross margin", hist=lambda c: f"={c}{gp}/{c}{rev}", fc=lambda c, p: f"={c}{gp}/{c}{rev}", h=f"=H{gp}/H{rev}", fmt=PCT, unit="%",
           note="Functional margin: excludes training, research, internal and free-tier compute (in R&D / S&M).")
    s.line("rnd", "Research & development", hist=lambda c: f"={{{O}!f_rnd@{c}}}", fc=lambda c, p: f"={{{O}!f_rnd@{c}}}", h="sum")
    s.line("snm", "Sales & marketing", hist=lambda c: f"={{{O}!f_snm@{c}}}", fc=lambda c, p: f"={{{O}!f_snm@{c}}}", h="sum")
    s.line("gna", "General & administrative", hist=lambda c: f"={{{O}!f_gna@{c}}}", fc=lambda c, p: f"={{{O}!f_gna@{c}}}", h="sum")
    s.line("bartz", "Bartz copyright settlement", hist=lambda c: f"={{{O}!f_bartz@{c}}}", fc=lambda c, p: f"={{{O}!f_bartz@{c}}}", h="sum")
    rnd, snm, gna, bz = (m.rows[("Income Statement", k)] for k in ("rnd", "snm", "gna", "bartz"))
    s.line("opex", "Total operating expenses (incl. cost of revenue)", hist=lambda c: f"={c}{cor}+{c}{rnd}+{c}{snm}+{c}{gna}+{c}{bz}",
           fc=lambda c, p: f"={c}{cor}+{c}{rnd}+{c}{snm}+{c}{gna}+{c}{bz}", h=f"=H{cor}+H{rnd}+H{snm}+H{gna}+H{bz}", border=TOP)
    opx = m.rows[("Income Statement", "opex")]
    s.line("op_income", "Operating income (loss)", hist=lambda c: f"={c}{rev}-{c}{opx}", fc=lambda c, p: f"={c}{rev}-{c}{opx}", h=f"=H{rev}-H{opx}", bold=True, border=TOP)
    oi = m.rows[("Income Statement", "op_income")]
    s.line("op_margin", "   Operating margin", hist=lambda c: f"={c}{oi}/{c}{rev}", fc=lambda c, p: f"={c}{oi}/{c}{rev}", h=f"=H{oi}/H{rev}", fmt=PCT, unit="%")
    s.line("interest_income", "Interest income", hist=lambda c: f"={{{H}!interest_income@{c}}}",
           fc=lambda c, p: f"={{Balance Sheet!cash@{p}}}*{{{A}!yield@{c}}}*{{{A}!frac@{c}}}", h="sum",
           note="Opening cash & STI × yield × period fraction.")
    s.line("interest_expense", "Interest expense", hist=[0, 0, 0, 0],
           fc=lambda c, p: f"=-{{Balance Sheet!debt@{p}}}*{{{A}!debt_rate}}*{{{A}!frac@{c}}}", h="sum")
    s.line("fv_change", "Change in FV of convertible notes, SAFEs & other financial instruments (non-cash)",
           hist=lambda c: f"={{{H}!fv_change@{c}}}", fc=lambda c, p: f"=-{{{A}!fv_charge@{c}}}", h="sum",
           note="Mark-to-market of Amazon notes etc. as the private valuation rose; ends on conversion at IPO.")
    s.line("other_inc", "Other income (expense), net", hist=lambda c: f"={{{H}!other_inc@{c}}}", fc=[0, 0, 0, 0, 0], h="sum")
    ii, ie, fv, oth = (m.rows[("Income Statement", k)] for k in ("interest_income", "interest_expense", "fv_change", "other_inc"))
    s.line("pretax", "Income (loss) before income taxes", hist=lambda c: f"={c}{oi}+{c}{ii}+{c}{ie}+{c}{fv}+{c}{oth}",
           fc=lambda c, p: f"={c}{oi}+{c}{ii}+{c}{ie}+{c}{fv}+{c}{oth}", h=f"=H{oi}+H{ii}+H{ie}+H{fv}+H{oth}", bold=True, border=TOP)
    pt = m.rows[("Income Statement", "pretax")]
    s.line("tax", "Provision for income taxes", hist=lambda c: f"=-{{{H}!tax@{c}}}", fc=lambda c, p: f"=-{{Balance Sheet!tax_total@{c}}}", h="sum",
           note="Forecast from the NOL schedule on the Balance Sheet sheet.")
    tx = m.rows[("Income Statement", "tax")]
    s.line("net_income", "Net income (loss)", hist=lambda c: f"={c}{pt}+{c}{tx}", fc=lambda c, p: f"={c}{pt}+{c}{tx}", h=f"=H{pt}+H{tx}", bold=True, border=TOPBOT)
    ni = m.rows[("Income Statement", "net_income")]
    s.line("ni_margin", "   Net margin", hist=lambda c: f"={c}{ni}/{c}{rev}", fc=lambda c, p: f"={c}{ni}/{c}{rev}", h=f"=H{ni}/H{rev}", fmt=PCT, unit="%")
    s.line("ni_check", "   Check vs. filed net loss (0 = ties)", hist=lambda c: f"={c}{ni}-{{{H}!net_income@{c}}}", fmt=NUM)

    s.section("Memo: operating expenses by nature")
    s.line("n_compute", "Compute & infrastructure", hist=lambda c: f"={{{C}!compute_total@{c}}}", fc=lambda c, p: f"={{{C}!compute_total@{c}}}", h="sum")
    s.line("n_inf", "   of which inference (paid + free-tier + internal)", hist=lambda c: f"={{{C}!inf_total@{c}}}", fc=lambda c, p: f"={{{C}!inf_total@{c}}}", h="sum")
    s.line("n_train", "   of which training & research (incl. take-or-pay shortfall)", hist=lambda c: f"={{{C}!train_total@{c}}}",
           fc=lambda c, p: f"={{{C}!train_total@{c}}}+{{{C}!excess@{c}}}", h="sum")
    s.line("n_pers", "Personnel", hist=lambda c: f"={{{O}!personnel@{c}}}", fc=lambda c, p: f"={{{O}!personnel@{c}}}", h="sum")
    s.line("n_sbc", "   of which stock-based compensation", hist=lambda c: f"={{{O}!sbc@{c}}}", fc=lambda c, p: f"={{{O}!sbc@{c}}}", h="sum")
    s.line("n_smp", "Sales, marketing & partnerships", hist=lambda c: f"={{{O}!smp@{c}}}", fc=lambda c, p: f"={{{O}!smp@{c}}}", h="sum")
    s.line("n_other", "Other operating & administrative", hist=lambda c: f"={{{O}!other_oa@{c}}}", fc=lambda c, p: f"={{{O}!other_oa@{c}}}", h="sum")
    s.line("n_bartz", "Bartz settlement", hist=lambda c: f"={c}{bz}", fc=lambda c, p: f"={c}{bz}", h="sum")
    nc, np_, ns, no, nb = (m.rows[("Income Statement", k)] for k in ("n_compute", "n_pers", "n_smp", "n_other", "n_bartz"))
    s.line("n_total", "Total operating expenses (by nature)", hist=lambda c: f"={c}{nc}+{c}{np_}+{c}{ns}+{c}{no}+{c}{nb}",
           fc=lambda c, p: f"={c}{nc}+{c}{np_}+{c}{ns}+{c}{no}+{c}{nb}", h=f"=H{nc}+H{np_}+H{ns}+H{no}+H{nb}", bold=True, border=TOP)
    ntot = m.rows[("Income Statement", "n_total")]
    s.line("n_check", "   Check: nature – functional (0 = ties)", hist=lambda c: f"={c}{ntot}-{c}{opx}", fc=lambda c, p: f"={c}{ntot}-{c}{opx}", h=f"=H{ntot}-H{opx}")
    s.line("compute_pct", "Compute & infrastructure as % of revenue", hist=lambda c: f"={c}{nc}/{c}{rev}", fc=lambda c, p: f"={c}{nc}/{c}{rev}", h=f"=H{nc}/H{rev}", fmt=PCT, unit="%")

    s.section("Non-GAAP (company definitions)")
    s.line("adj_op", "Adjusted operating income (op. income + SBC + non-cash donation matching + Bartz)",
           hist=lambda c: f"={c}{oi}+{{{O}!sbc@{c}}}+{{{O}!donation@{c}}}+{c}{bz}",
           fc=lambda c, p: f"={c}{oi}+{{{O}!sbc@{c}}}+{{{O}!donation@{c}}}+{c}{bz}",
           h=f"=H{oi}+{{{O}!sbc@H}}+{{{O}!donation@H}}+H{bz}", bold=True,
           note="Filed: FY25 $(4,854)M; H1'26 $1,618M. Q3'26 margin ~11% per management.")
    ao = m.rows[("Income Statement", "adj_op")]
    s.line("adj_op_margin", "   Adjusted operating margin", hist=lambda c: f"={c}{ao}/{c}{rev}", fc=lambda c, p: f"={c}{ao}/{c}{rev}", h=f"=H{ao}/H{rev}", fmt=PCT, unit="%")
    s.line("adj_ebitda", "Adjusted EBITDA (adjusted operating income + D&A)", hist=lambda c: f"={c}{ao}+{{{O}!da@{c}}}",
           fc=lambda c, p: f"={c}{ao}+{{{O}!da@{c}}}", h=f"=H{ao}+{{{O}!da@H}}", bold=True)
    ae = m.rows[("Income Statement", "adj_ebitda")]
    s.line("adj_ebitda_margin", "   Adjusted EBITDA margin", hist=lambda c: f"={c}{ae}/{c}{rev}", fc=lambda c, p: f"={c}{ae}/{c}{rev}", h=f"=H{ae}/H{rev}", fmt=PCT, unit="%")
    s.line("adj_ni", "Adjusted net income (net income + SBC + donation + Bartz + FV changes)",
           hist=lambda c: f"={c}{ni}+{{{O}!sbc@{c}}}+{{{O}!donation@{c}}}+{c}{bz}-{c}{fv}",
           fc=lambda c, p: f"={c}{ni}+{{{O}!sbc@{c}}}+{{{O}!donation@{c}}}+{c}{bz}-{c}{fv}",
           h=f"=H{ni}+{{{O}!sbc@H}}+{{{O}!donation@H}}+H{bz}-H{fv}", bold=True, note="Filed: FY25 $(4,448)M; H1'26 $2,671M.")
    return s


def build_bs(m: Model):
    s = m.sheet("Balance Sheet", flows=False)
    s.title("Balance Sheet", "Rolls forward from the Jun-26A filed balance sheet. Convertible notes and preferred stock convert to common equity at the IPO in H2'26E.")
    s.period_header()
    H, A, R, C, O, CF, IS = "Historical Data", "Assumptions", "Revenue Build", "Compute Build", "Opex Build", "Cash Flow", "Income Statement"
    hc = ["D", "E", "F"]

    def hist_link(key):
        return lambda c: f"={{{H}!bs_{key}@{c}}}"

    s.section("Assets")
    s.line("cash", "Cash, cash equivalents & short-term investments", hist=lambda c: f"={{{H}!bs_cash@{c}}}+{{{H}!bs_sti@{c}}}", hist_cols=hc,
           fc=lambda c, p: f"={{{CF}!cash_end@{c}}}", h="same", bold=True, note="Forecast: one pool (STI treated as cash).")
    s.line("ar", "Accounts receivable, net", hist=hist_link("ar"), hist_cols=hc,
           fc=lambda c, p: f"={{{A}!ar_pct@{c}}}*{{{R}!revenue@{c}}}/{{{A}!frac@{c}}}", h="same")
    s.line("prepaid", "Prepaid expenses & other current assets", hist=hist_link("prepaid"), hist_cols=hc,
           fc=lambda c, p: f"={{{A}!prepaid_pct@{c}}}*{{{C}!compute_total@{c}}}/{{{A}!frac@{c}}}", h="same")
    ca = [m.rows[("Balance Sheet", k)] for k in ("cash", "ar", "prepaid")]
    s.line("tca", "Total current assets", hist=lambda c: "=" + "+".join(f"{c}{r}" for r in ca), hist_cols=hc,
           fc=lambda c, p: "=" + "+".join(f"{c}{r}" for r in ca), h="same", bold=True, border=TOP)
    s.line("restricted", "Restricted cash", hist=hist_link("restricted"), hist_cols=hc, fc=lambda c, p: f"={p}{m.rows[('Balance Sheet','restricted')]}", h="same",
           note="Held flat.")
    s.line("ppe", "Property & equipment, net (incl. finance-lease assets)", hist=hist_link("ppe"), hist_cols=hc,
           fc=lambda c, p: f"={p}{m.rows[('Balance Sheet','ppe')]}+{{{O}!capex@{c}}}-{{{O}!da@{c}}}", h="same")
    s.line("rou", "Operating lease right-of-use assets", hist=hist_link("rou"), hist_cols=hc, fc=lambda c, p: f"={p}{m.rows[('Balance Sheet','rou')]}", h="same",
           note="Held flat (offices); matched by lease liabilities below.")
    s.line("goodwill", "Goodwill", hist=hist_link("goodwill"), hist_cols=hc, fc=lambda c, p: f"={p}{m.rows[('Balance Sheet','goodwill')]}", h="same")
    s.line("other_nca", "Other assets, noncurrent", hist=hist_link("other_nca"), hist_cols=hc, fc=lambda c, p: f"={p}{m.rows[('Balance Sheet','other_nca')]}", h="same",
           note="Held flat (deposits / prepayments on compute capacity).")
    nca = [m.rows[("Balance Sheet", k)] for k in ("restricted", "ppe", "rou", "goodwill", "other_nca")]
    tca = m.rows[("Balance Sheet", "tca")]
    s.line("total_assets", "Total assets", hist=lambda c: f"={c}{tca}+" + "+".join(f"{c}{r}" for r in nca), hist_cols=hc,
           fc=lambda c, p: f"={c}{tca}+" + "+".join(f"{c}{r}" for r in nca), h="same", bold=True, border=TOPBOT)
    ta = m.rows[("Balance Sheet", "total_assets")]

    s.section("Liabilities")
    s.line("ap", "Accounts payable, accrued expenses & other current liabilities", hist=hist_link("ap"), hist_cols=hc,
           fc=lambda c, p: f"={{{A}!ap_pct@{c}}}*({{{IS}!opex@{c}}}-{{{O}!sbc@{c}}}-{{{O}!donation@{c}}}-{{{O}!da@{c}}})/{{{A}!frac@{c}}}", h="same",
           note="% of annualised cash operating costs (total opex less SBC, donation matching and D&A).")
    s.line("lease_liab", "Operating lease liabilities (current + noncurrent)", hist=hist_link("lease_liab"), hist_cols=hc,
           fc=lambda c, p: f"={p}{m.rows[('Balance Sheet','lease_liab')]}", h="same")
    s.line("defrev", "Deferred revenue", hist=hist_link("defrev"), hist_cols=hc,
           fc=lambda c, p: f"={{{A}!defrev_pct@{c}}}*{{{R}!rev_sub@{c}}}/{{{A}!frac@{c}}}", h="same")
    s.line("conv_notes", "Convertible notes payable (at fair value)", hist=hist_link("conv_notes"), hist_cols=hc,
           fc=lambda c, p: f"={p}{m.rows[('Balance Sheet','conv_notes')]}+{{{A}!fv_charge@{c}}}-{c}{{notes_conv_placeholder}}", h="same",
           note="Amazon notes: $6.2B principal, $108.6B fair value at Jun-26; convert to Class N common at IPO.")
    s.line("debt", "Debt / supplier financing", hist=[0, 0, 0], hist_cols=hc,
           fc=lambda c, p: f"={p}{m.rows[('Balance Sheet','debt')]}+{{{A}!debt_draw@{c}}}", h="same")
    s.line("other_ncl", "Other noncurrent liabilities", hist=hist_link("other_ncl"), hist_cols=hc,
           fc=lambda c, p: f"={p}{m.rows[('Balance Sheet','other_ncl')]}", h="same")
    liabs = [m.rows[("Balance Sheet", k)] for k in ("ap", "lease_liab", "defrev", "conv_notes", "debt", "other_ncl")]
    s.line("total_liab", "Total liabilities", hist=lambda c: "=" + "+".join(f"{c}{r}" for r in liabs), hist_cols=hc,
           fc=lambda c, p: "=" + "+".join(f"{c}{r}" for r in liabs), h="same", bold=True, border=TOP)
    tl = m.rows[("Balance Sheet", "total_liab")]

    s.section("Mezzanine & stockholders' equity")
    s.line("pref", "Redeemable convertible preferred stock (mezzanine)", hist=hist_link("pref"), hist_cols=hc,
           fc=lambda c, p: f"={p}{m.rows[('Balance Sheet','pref')]}-{c}{{pref_conv_placeholder}}", h="same",
           note="Series A–H; converts 1:1 into common at IPO.")
    s.line("apic", "Common stock & additional paid-in capital", hist=hist_link("apic"), hist_cols=hc,
           fc=lambda c, p: (f"={p}{m.rows[('Balance Sheet','apic')]}+{{{O}!sbc@{c}}}+{{{O}!donation@{c}}}+{{{A}!esop_proceeds@{c}}}"
                            f"+{{{A}!ipo_proceeds@{c}}}+{c}{{pref_conv_placeholder}}+{c}{{notes_conv_placeholder}}"), h="same",
           note="+ SBC, share-settled donation matching, option/ESPP proceeds, IPO proceeds, conversion of preferred and notes.")
    s.line("aoci", "Accumulated other comprehensive income (loss)", hist=hist_link("aoci"), hist_cols=hc,
           fc=lambda c, p: f"={p}{m.rows[('Balance Sheet','aoci')]}", h="same")
    s.line("deficit", "Accumulated deficit", hist=hist_link("deficit"), hist_cols=hc,
           fc=lambda c, p: f"={p}{m.rows[('Balance Sheet','deficit')]}+{{{IS}!net_income@{c}}}", h="same")
    eq = [m.rows[("Balance Sheet", k)] for k in ("pref", "apic", "aoci", "deficit")]
    s.line("total_equity", "Total equity incl. mezzanine preferred", hist=lambda c: "=" + "+".join(f"{c}{r}" for r in eq), hist_cols=hc,
           fc=lambda c, p: "=" + "+".join(f"{c}{r}" for r in eq), h="same", bold=True, border=TOP)
    te = m.rows[("Balance Sheet", "total_equity")]
    s.line("total_le", "Total liabilities, mezzanine & equity", hist=lambda c: f"={c}{tl}+{c}{te}", hist_cols=hc,
           fc=lambda c, p: f"={c}{tl}+{c}{te}", h="same", bold=True, border=TOPBOT)
    tle = m.rows[("Balance Sheet", "total_le")]
    s.line("bs_check", "Check: assets – (liabilities + equity) (0 = balances)", hist=lambda c: f"=ROUND({c}{ta}-{c}{tle},3)", hist_cols=hc,
           fc=lambda c, p: f"=ROUND({c}{ta}-{c}{tle},3)", h="same", bold=True, fill=FILL_CHECK_OK)

    s.section("Supporting schedules")
    s.line("notes_conv", "Conversion of convertible notes into common equity at IPO (non-cash)",
           fc=lambda c, p: (f"={p}{m.rows[('Balance Sheet','conv_notes')]}+{{{A}!fv_charge@{c}}}" if c == "G" else "=0"), h="same",
           note="All outstanding notes (carrying value incl. the pre-IPO fair-value charge) convert in H2'26E.")
    s.line("pref_conv", "Conversion of preferred stock into common equity at IPO (non-cash)",
           fc=lambda c, p: (f"={p}{m.rows[('Balance Sheet','pref')]}" if c == "G" else "=0"), h="same")
    s.line("nwc", "Net working capital (AR + prepaid – AP – deferred revenue)",
           hist=lambda c: f"={c}{ca[1]}+{c}{ca[2]}-{c}{liabs[0]}-{c}{liabs[2]}", hist_cols=hc,
           fc=lambda c, p: f"={c}{ca[1]}+{c}{ca[2]}-{c}{liabs[0]}-{c}{liabs[2]}", h="same")

    s.section("Tax schedule (simplified NOL)")
    s.line("taxable", "Taxable income proxy (pre-tax income + non-deductible FV charge + donation matching)",
           fc=lambda c, p: f"={{{IS}!pretax@{c}}}-{{{IS}!fv_change@{c}}}+{{{O}!donation@{c}}}", h="sum")
    txb = m.rows[("Balance Sheet", "taxable")]
    nol_open_row = s.line("nol_open", "NOL carryforward – opening", note="Opening balance is an input on Assumptions; thereafter rolls from prior closing.")
    s.line("nol_used", "NOL utilised", fc=lambda c, p: f"=MIN({c}{nol_open_row},MAX(0,{c}{txb}))", h="sum")
    nu = m.rows[("Balance Sheet", "nol_used")]
    s.line("nol_close", "NOL carryforward – closing", fc=lambda c, p: f"={c}{nol_open_row}-{c}{nu}+MAX(0,-{c}{txb})", h="same")
    ncl = m.rows[("Balance Sheet", "nol_close")]
    for c in FC:
        s.put(f"{c}{nol_open_row}", f"={{{A}!nol_open}}" if c == "G" else f"={PREV[c]}{ncl}", NUM)
    s.put(f"H{nol_open_row}", f"=G{nol_open_row}", NUM)
    s.line("tax_total", "Income tax expense (cash & deferred, simplified)",
           fc=lambda c, p: f"=MAX(0,{c}{txb}-{c}{nu})*{{{A}!tax_rate}}+{{{A}!other_tax_pct}}*{{{R}!revenue@{c}}}", h="sum",
           note="Statutory rate on taxable income after NOLs, plus minimum/other taxes as % of revenue.")

    # resolve placeholders inside pending formulas for this sheet
    nc_row, pc_row = m.rows[("Balance Sheet", "notes_conv")], m.rows[("Balance Sheet", "pref_conv")]
    m.pending = [(sh, coord, raw.replace("{notes_conv_placeholder}", str(nc_row)).replace("{pref_conv_placeholder}", str(pc_row)))
                 for sh, coord, raw in m.pending]
    return s


def build_cf(m: Model):
    s = m.sheet("Cash Flow")
    s.title("Cash Flow Statement", "Indirect method. Historical columns as filed (restated to this line structure); forecast treats cash and short-term investments as one pool.")
    s.period_header()
    H, A, O, IS, BS = "Historical Data", "Assumptions", "Opex Build", "Income Statement", "Balance Sheet"

    def d_bs(key, sign):  # change in balance-sheet item (asset: prev - cur; liability: cur - prev)
        if sign < 0:
            return lambda c, p: f"={{{BS}!{key}@{p}}}-{{{BS}!{key}@{c}}}"
        return lambda c, p: f"={{{BS}!{key}@{c}}}-{{{BS}!{key}@{p}}}"

    s.section("Operating activities")
    s.line("ni", "Net income (loss)", hist=lambda c: f"={{{IS}!net_income@{c}}}", fc=lambda c, p: f"={{{IS}!net_income@{c}}}", h="sum")
    s.line("sbc", "Stock-based compensation", hist=lambda c: f"={{{H}!sbc@{c}}}", fc=lambda c, p: f"={{{O}!sbc@{c}}}", h="sum",
           note="Historical: cash-flow add-back (differs slightly from IS SBC in FY25 due to capitalised amounts).")
    s.line("da", "Depreciation & amortisation", hist=lambda c: f"={{{H}!da@{c}}}", fc=lambda c, p: f"={{{O}!da@{c}}}", h="sum")
    s.line("fv", "Change in FV of convertible notes, SAFEs & other financial instruments (non-cash)", hist=lambda c: f"=-{{{H}!fv_change@{c}}}",
           fc=lambda c, p: f"=-{{{IS}!fv_change@{c}}}", h="sum")
    s.line("donation", "Donation-matching expense settled in shares (non-cash)", hist=lambda c: f"={{{H}!donation@{c}}}", fc=lambda c, p: f"={{{O}!donation@{c}}}", h="sum")
    s.line("other_nc", "Other non-cash items", hist=lambda c: f"={{{H}!cf_other@{c}}}", fc=[0, 0, 0, 0, 0], h="sum")
    s.line("d_ar", "(Increase) decrease in accounts receivable", hist=lambda c: f"={{{H}!d_ar@{c}}}", fc=d_bs("ar", -1), h="sum")
    s.line("d_prepaid", "(Increase) decrease in prepaid expenses & other assets", hist=lambda c: f"={{{H}!d_prepaid@{c}}}", fc=d_bs("prepaid", -1), h="sum")
    s.line("d_ap", "Increase (decrease) in accounts payable, accrued & other liabilities", hist=lambda c: f"={{{H}!d_ap@{c}}}", fc=d_bs("ap", +1), h="sum")
    s.line("d_lease", "Increase (decrease) in lease liabilities", hist=lambda c: f"={{{H}!d_lease@{c}}}", fc=d_bs("lease_liab", +1), h="sum")
    s.line("d_defrev", "Increase (decrease) in deferred revenue", hist=lambda c: f"={{{H}!d_defrev@{c}}}", fc=d_bs("defrev", +1), h="sum")
    ops = [m.rows[("Cash Flow", k)] for k in ("ni", "sbc", "da", "fv", "donation", "other_nc", "d_ar", "d_prepaid", "d_ap", "d_lease", "d_defrev")]
    s.line("cfo", "Net cash from operating activities", hist=lambda c: "=" + "+".join(f"{c}{r}" for r in ops),
           fc=lambda c, p: "=" + "+".join(f"{c}{r}" for r in ops), h="sum", bold=True, border=TOP)
    cfo = m.rows[("Cash Flow", "cfo")]
    s.line("cfo_check", "   Check vs. filed operating cash flow (0 = ties)", hist=lambda c: f"={c}{cfo}-{{{H}!cfo@{c}}}")

    s.section("Investing activities")
    s.line("capex", "Purchases of property & equipment", hist=lambda c: f"={{{H}!capex@{c}}}", fc=lambda c, p: f"=-{{{O}!capex@{c}}}", h="sum")
    s.line("sti", "Purchases of short-term investments, net", hist=lambda c: f"={{{H}!sti_net@{c}}}", fc=[0, 0, 0, 0, 0], h="sum",
           note="Forecast: STI modelled within the cash pool, so no separate flow.")
    s.line("other_inv", "Other investing activities", hist=lambda c: f"={{{H}!other_inv@{c}}}", fc=[0, 0, 0, 0, 0], h="sum")
    inv = [m.rows[("Cash Flow", k)] for k in ("capex", "sti", "other_inv")]
    s.line("cfi", "Net cash from investing activities", hist=lambda c: "=" + "+".join(f"{c}{r}" for r in inv),
           fc=lambda c, p: "=" + "+".join(f"{c}{r}" for r in inv), h="sum", bold=True, border=TOP)
    cfi = m.rows[("Cash Flow", "cfi")]

    s.section("Financing activities")
    s.line("pref", "Proceeds from preferred stock, net", hist=lambda c: f"={{{H}!pref_proceeds@{c}}}", fc=[0, 0, 0, 0, 0], h="sum")
    s.line("notes", "Issuance of convertible notes for cash", hist=lambda c: f"={{{H}!conv_notes_cash@{c}}}", fc=[0, 0, 0, 0, 0], h="sum")
    s.line("ipo", "IPO primary proceeds, net", hist=[0, 0, 0, 0], fc=lambda c, p: f"={{{A}!ipo_proceeds@{c}}}", h="sum")
    s.line("common", "Proceeds from common stock (options / ESPP)", hist=lambda c: f"={{{H}!common_proceeds@{c}}}", fc=lambda c, p: f"={{{A}!esop_proceeds@{c}}}", h="sum")
    s.line("debt", "Debt / supplier financing drawdown (repayment)", hist=[0, 0, 0, 0], fc=lambda c, p: f"={{{A}!debt_draw@{c}}}", h="sum")
    s.line("other_fin", "Other financing", hist=lambda c: f"={{{H}!other_fin@{c}}}", fc=[0, 0, 0, 0, 0], h="sum")
    fin = [m.rows[("Cash Flow", k)] for k in ("pref", "notes", "ipo", "common", "debt", "other_fin")]
    s.line("cff", "Net cash from financing activities", hist=lambda c: "=" + "+".join(f"{c}{r}" for r in fin),
           fc=lambda c, p: "=" + "+".join(f"{c}{r}" for r in fin), h="sum", bold=True, border=TOP)
    cff = m.rows[("Cash Flow", "cff")]

    s.section("Cash reconciliation")
    s.line("fx", "Effect of exchange rates", hist=lambda c: f"={{{H}!fx@{c}}}", fc=[0, 0, 0, 0, 0], h="sum")
    fx = m.rows[("Cash Flow", "fx")]
    s.line("net_change", "Net change in cash (forecast: cash & STI; historical: cash & restricted cash as filed)",
           hist=lambda c: f"={c}{cfo}+{c}{cfi}+{c}{cff}+{c}{fx}", fc=lambda c, p: f"={c}{cfo}+{c}{cfi}+{c}{cff}+{c}{fx}", h="sum", bold=True)
    nch = m.rows[("Cash Flow", "net_change")]
    s.line("cash_begin", "Cash & STI – beginning of period", fc=lambda c, p: f"={{{BS}!cash@{p}}}", h=f"={{{BS}!cash@E}}")
    cb = m.rows[("Cash Flow", "cash_begin")]
    s.line("cash_end", "Cash & STI – end of period", fc=lambda c, p: f"={c}{cb}+{c}{nch}", h=f"={{{BS}!cash@G}}", bold=True, border=TOPBOT,
           hist=lambda c: f"={{{BS}!cash@{c}}}", hist_cols=["D", "E", "F"],
           note="Forecast: beginning + net change (feeds the balance sheet). Historical: per balance sheet.")
    s.line("cash_end_filed", "Memo: cash, equivalents & restricted cash – end of period (as filed)", hist=lambda c: f"={{{H}!cash_end_filed@{c}}}")

    s.section("Free cash flow")
    cx = m.rows[("Cash Flow", "capex")]
    s.line("fcf", "Free cash flow (CFO – capex)", hist=lambda c: f"={c}{cfo}+{c}{cx}", fc=lambda c, p: f"={c}{cfo}+{c}{cx}", h="sum", bold=True)
    fcf = m.rows[("Cash Flow", "fcf")]
    s.line("fcf_margin", "   FCF margin", hist=lambda c: f"={c}{fcf}/{{{IS}!revenue@{c}}}", fc=lambda c, p: f"={c}{fcf}/{{{IS}!revenue@{c}}}",
           h=f"=H{fcf}/{{{IS}!revenue@H}}", fmt=PCT, unit="%", note="The Information: H1'26 FCF ≈ $3.5B.")
    s.line("fcf_ex_sbc", "FCF less SBC (shareholder-adjusted)", hist=lambda c: f"={c}{fcf}-{{{O}!sbc@{c}}}", fc=lambda c, p: f"={c}{fcf}-{{{O}!sbc@{c}}}",
           h=f"=H{fcf}-{{{O}!sbc@H}}")
    return s


def build_commitments(m: Model):
    s = m.sheet("Compute Commitments")
    s.title("Compute Commitments – contracted capacity vs. modelled demand",
            "Estimated annual expense phasing of the agreements disclosed in the S-1 (deal values as extracted by multiple outlets). "
            "Phasing is the author's estimate; the S-1 gives deal totals and terms, not annual schedules. Column B: 1 = non-cancellable.")
    s.period_header()
    s.ws.column_dimensions["A"].width = 62
    s.ws.column_dimensions["B"].width = 12
    s.section("Estimated expense recognised as capacity comes on-line ($M)")
    first = s.r
    for name, total_b, flag, sched, note in COMMITMENTS:
        row = s.r
        s.put(f"A{row}", name)
        s.put(f"B{row}", flag, "0", align="center")
        for c, v in zip(FC, sched):
            s.put(f"{c}{row}", v, NUM)
        s.put(f"H{row}", f"=F{row}+G{row}", NUM)
        s.put(f"{NOTE_COL}{row}", f"Deal value ~${total_b:,.1f}B. {note}", font=GREY_I)
        s.r += 1
    last = s.r - 1
    s.line("total", "Total estimated contracted compute expense", fc=lambda c, p: f"=SUM({c}{first}:{c}{last})", h="sum", bold=True, border=TOP)
    s.line("noncancel_total", "   of which non-cancellable (flag = 1)", fc=lambda c, p: f"=SUMPRODUCT($B${first}:$B${last},{c}{first}:{c}{last})", h="sum", bold=True,
           note="Feeds the take-or-pay floor on the Compute Build sheet.")
    s.line("cancel_total", "   of which cancellable (SpaceXAI)", fc=lambda c, p: f"={c}{m.rows[('Compute Commitments','total')]}-{c}{m.rows[('Compute Commitments','noncancel_total')]}", h="sum")
    s.line("cum", "Cumulative contracted expense 2026–2030 (memo)", fc=lambda c, p: (f"=G{m.rows[('Compute Commitments','total')]}" if c == "G" else f"={p}{m.rows[('Compute Commitments','cum')]}+{c}{m.rows[('Compute Commitments','total')]}"),
           h="same", note="Compare: $265.6B minimum purchase commitments + $321.4B uncommenced leases at Jun-26A (through ~2036), plus ~$245B signed after Jun-26.")

    s.section("Comparison with modelled compute demand")
    C = "Compute Build"
    s.line("demand", "Driver-based compute demand (Compute Build)", fc=lambda c, p: f"={{{C}!demand@{c}}}", h="sum")
    s.line("total_expense", "Total compute & infrastructure expense in model", fc=lambda c, p: f"={{{C}!compute_total@{c}}}", h="sum", bold=True)
    tot, dm, te = (m.rows[("Compute Commitments", k)] for k in ("total", "demand", "total_expense"))
    s.line("gap", "Modelled expense – total contracted (positive = legacy capacity / on-demand / spot purchases)",
           fc=lambda c, p: f"={c}{te}-{c}{tot}", h="sum",
           note="Pre-2026 Amazon/Google capacity and short-term purchases are not in the schedule above, so a positive gap is expected in 2026–27.")
    s.line("idle", "Take-or-pay shortfall expensed (from Compute Build)", fc=lambda c, p: f"={{{C}!excess@{c}}}", h="sum",
           note="Non-zero when non-cancellable minimums exceed driver-based demand (check the Bear case).")
    s.line("coverage", "Contracted capacity coverage of demand (total contracted / demand)", fc=lambda c, p: f"={c}{tot}/{c}{dm}", h=f"=H{tot}/H{dm}", fmt=MULT, unit="x")
    return s


def build_checks(m: Model):
    s = m.sheet("Checks")
    s.title("Model Integrity Checks", "All rows should read 0 (or TRUE). Overall status at the top.")
    s.period_header()
    BS, CF, IS, O, C = "Balance Sheet", "Cash Flow", "Income Statement", "Opex Build", "Compute Build"
    s.line("bs", "Balance sheet balances (assets – liabilities – equity)", hist=lambda c: f"={{{BS}!bs_check@{c}}}", hist_cols=["D", "E", "F"],
           fc=lambda c, p: f"={{{BS}!bs_check@{c}}}", h=f"={{{BS}!bs_check@H}}")
    s.line("cash", "Cash flow ending cash ties to balance sheet", fc=lambda c, p: f"=ROUND({{{CF}!cash_end@{c}}}-{{{BS}!cash@{c}}},3)", h=f"=ROUND({{{CF}!cash_end@H}}-{{{BS}!cash@H}},3)")
    s.line("nature", "Opex by nature ties to functional presentation", hist=lambda c: f"=ROUND({{{IS}!n_check@{c}}},3)", fc=lambda c, p: f"=ROUND({{{IS}!n_check@{c}}},3)", h=f"=ROUND({{{IS}!n_check@H}},3)")
    s.line("ni", "Net income ties to filed figures (historical)", hist=lambda c: f"=ROUND({{{IS}!ni_check@{c}}},3)")
    s.line("cfo", "Operating cash flow ties to filed figures (historical)", hist=lambda c: f"=ROUND({{{CF}!cfo_check@{c}}},3)")
    s.line("alloc", "Expense allocation percentages sum to 100%",
           fc=[("=ROUND({Assumptions!pers_cor}+{Assumptions!pers_rnd}+{Assumptions!pers_snm}+{Assumptions!pers_gna}-1,6)"
                "+ROUND({Assumptions!oth_cor}+{Assumptions!oth_rnd}+{Assumptions!oth_snm}+{Assumptions!oth_gna}-1,6)"), None, None, None, None], fc_cols=["G"])
    s.line("deficit", "Accumulated deficit rolls with net income (Jun-26A → Dec-26E)",
           fc=[f"=ROUND({{{BS}!deficit@G}}-{{{BS}!deficit@F}}-{{{IS}!net_income@G}},3)", None, None, None, None], fc_cols=["G"])
    rows = [m.rows[("Checks", k)] for k in ("bs", "cash", "nature", "ni", "cfo", "alloc", "deficit")]
    s.blank()
    s.put(f"A{s.r}", "OVERALL MODEL STATUS", bold=True)
    m.rows[("Checks", "status")] = s.r
    rng = ",".join(f"C{r}:L{r}" for r in rows)
    s.put(f"C{s.r}", f'=IF(SUMPRODUCT(ABS(C{rows[0]}:L{rows[-1]}))<0.01,"OK – all checks pass","ERROR – review checks")', bold=True, fill=FILL_CHECK_OK)
    s.r += 1
    return s


def build_summary(m: Model):
    s = m.sheet("Summary")
    s.title("Summary – key outputs (fiscal-year view)", "Active scenario shown below (switch it on Assumptions!C4). Charts update with the Assumptions sheet.")
    s.put(f"A{s.r}", "Active scenario", bold=True)
    s.put(f"C{s.r}", "={Assumptions!scenario}", INT)
    s.put(f"D{s.r}", '=CHOOSE({Assumptions!scenario},"Base","Bull","Bear")', bold=True)
    s.r += 2
    # Fiscal-year-only view: label in A, FY2023A..FY2030E in B..I (no unit column so chart series read cleanly)
    fy_labels = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"]
    out_cols = ["B", "C", "D", "E", "F", "G", "H", "I"]
    s.ws.column_dimensions["A"].width = 62
    for oc in out_cols:
        s.ws.column_dimensions[oc].width = 12
    hdr = s.r
    s.put(f"A{hdr}", "$M unless noted in label", font=WHITE_B, fill=FILL_HDR)
    for oc, lab in zip(out_cols, fy_labels):
        s.put(f"{oc}{hdr}", lab, font=WHITE_B, fill=FILL_HDR, align="center")
    s.r += 1
    s.ws.freeze_panes = f"B{s.r}"

    def srow(key, label, sheet, src_key, fmt=NUM, bold=False, border=None):
        row = s.r
        m.rows[("Summary", key)] = row
        s.put(f"A{row}", label, bold=bold)
        for oc, src in zip(out_cols, FY_COLS):
            s.put(f"{oc}{row}", f"={{{sheet}!{src_key}@{src}}}", fmt, bold)
        if border:
            for col in ["A"] + out_cols:
                s.ws[f"{col}{row}"].border = border
        s.r += 1
        return row

    s.section("Growth")
    srow("rev", "Revenue", "Income Statement", "revenue", bold=True)
    r_rev = m.rows[("Summary", "rev")]
    row = s.r
    s.put(f"A{row}", "   YoY growth (%)")
    for i, oc in enumerate(out_cols[1:], start=1):
        s.put(f"{oc}{row}", f"={oc}{r_rev}/{out_cols[i-1]}{r_rev}-1", PCT)
    s.r += 1
    srow("hc", "Year-end headcount (FTE)", "Opex Build", "hc_ye", fmt=INT)

    s.section("Compute – training vs. inference")
    srow("inf", "Inference compute (paid + free-tier + internal)", "Compute Build", "inf_total")
    srow("train", "Training & research compute (incl. take-or-pay shortfall)", "Income Statement", "n_train")
    srow("compute", "Total compute & infrastructure", "Compute Build", "compute_total", bold=True, border=TOP)
    srow("compute_pct", "   Compute as % of revenue", "Compute Build", "compute_pct_rev", fmt=PCT)
    srow("train_share", "   Training & research share of compute (%)", "Compute Build", "train_share", fmt=PCT)
    srow("rpc", "   Revenue per $1 of compute (x)", "Compute Build", "rev_per_compute", fmt=MULT)

    s.section("Profitability")
    srow("gm", "Gross margin (functional, %)", "Income Statement", "gm", fmt=PCT)
    srow("op", "Operating income (loss) – GAAP", "Income Statement", "op_income", bold=True)
    srow("opm", "   GAAP operating margin (%)", "Income Statement", "op_margin", fmt=PCT)
    srow("adj_op", "Adjusted operating income (ex-SBC, donation matching, Bartz)", "Income Statement", "adj_op", bold=True)
    srow("adj_opm", "   Adjusted operating margin (%)", "Income Statement", "adj_op_margin", fmt=PCT)
    srow("ni", "Net income (loss) – GAAP", "Income Statement", "net_income", bold=True)
    srow("adj_ni", "Adjusted net income", "Income Statement", "adj_ni")

    s.section("Cash flow & balance sheet")
    srow("cfo", "Operating cash flow", "Cash Flow", "cfo")
    srow("fcf", "Free cash flow", "Cash Flow", "fcf", bold=True)
    srow("fcfm", "   FCF margin (%)", "Cash Flow", "fcf_margin", fmt=PCT)
    srow("cash", "Cash & short-term investments (year-end)", "Balance Sheet", "cash", bold=True)
    srow("equity", "Total equity incl. preferred (year-end)", "Balance Sheet", "total_equity")

    s.section("Valuation snapshot (illustrative)")
    row = s.r
    m.rows[("Summary", "mcap")] = row
    s.put(f"A{row}", "Assumed equity value at IPO ($M) – input")
    s.put(f"B{row}", 1500000, NUM)
    s.put(f"{NOTE_COL}{row}", "Press reports: $1.5–2.0T discussed; the S-1 price range is blank.", font=GREY_I)
    s.r += 1
    mc = row
    row = s.r
    m.rows[("Summary", "ev")] = row
    s.put(f"A{row}", "Implied enterprise value (equity value – Dec-26E cash & STI + debt)")
    s.put(f"B{row}", f"=B{mc}-{{Balance Sheet!cash@G}}+{{Balance Sheet!debt@G}}", NUM)
    s.r += 1
    for key, label, sheet, src_key in [("ev_rev", "EV / revenue (x)", "Income Statement", "revenue"),
                                       ("ev_adj", "EV / adjusted operating income (x)", "Income Statement", "adj_op"),
                                       ("ev_fcf", "Equity value / free cash flow (x)", "Cash Flow", "fcf")]:
        row = s.r
        m.rows[("Summary", key)] = row
        s.put(f"A{row}", label)
        for oc, src in zip(out_cols[3:], FY_COLS[3:]):
            base = f"$B${m.rows[('Summary','ev')]}" if key != "ev_fcf" else f"$B${mc}"
            s.put(f"{oc}{row}", f"=IF({{{sheet}!{src_key}@{src}}}>0,{base}/{{{sheet}!{src_key}@{src}}},\"n.m.\")", '0.0"x"')
        s.r += 1

    # ---- Charts (FY2025A–FY2030E; FY2023–24 ratios are too extreme to plot) ----
    s.r += 1
    chart_row = s.r
    first_chart_col = 4  # column D = FY2025A
    cats = Reference(s.ws, min_col=first_chart_col, max_col=9, min_row=hdr, max_row=hdr)

    def series(ch, keys):
        for k in keys:
            r = m.rows[("Summary", k)]
            ser = Series(Reference(s.ws, min_col=first_chart_col, max_col=9, min_row=r, max_row=r),
                         title=str(s.ws[f"A{r}"].value).strip())
            ch.series.append(ser)
        ch.set_categories(cats)
        ch.height, ch.width = 7.5, 16
        ch.legend.position = "b"

    def add_bar(title, keys, anchor, stacked=False):
        ch = BarChart()
        ch.type = "col"
        ch.title = title
        if stacked:
            ch.grouping = "stacked"
            ch.overlap = 100
        series(ch, keys)
        s.ws.add_chart(ch, anchor)

    def add_line(title, keys, anchor):
        ch = LineChart()
        ch.title = title
        series(ch, keys)
        ch.y_axis.numFmt = "0%"
        ch.y_axis.scaling.min = -2.0
        ch.y_axis.scaling.max = 1.0
        s.ws.add_chart(ch, anchor)

    add_bar("Revenue ($M)", ["rev"], f"A{chart_row}")
    add_bar("Compute: inference vs. training & research ($M)", ["inf", "train"], f"E{chart_row}", stacked=True)
    add_line("Margins and compute intensity", ["gm", "opm", "adj_opm", "compute_pct"], f"A{chart_row + 16}")
    add_bar("Free cash flow and year-end cash ($M)", ["fcf", "cash"], f"E{chart_row + 16}")
    return s


def main():
    m = Model()
    build_cover(m)
    build_assumptions(m)
    hist = build_historical(m)
    m.q3_col = hist.q3_col
    build_revenue(m)
    build_compute(m)
    build_opex(m)
    build_is(m)
    build_bs(m)
    build_cf(m)
    build_commitments(m)
    build_checks(m)
    build_summary(m)
    m.finalize()
    # order sheets: Cover, Summary, Assumptions, ...
    order = ["Cover", "Summary", "Assumptions", "Revenue Build", "Compute Build", "Opex Build", "Income Statement",
             "Balance Sheet", "Cash Flow", "Compute Commitments", "Historical Data", "Checks"]
    m.wb._sheets = [m.wb[n] for n in order]
    m.wb.active = 1
    # print setup
    for ws in m.wb.worksheets:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    OUT.parent.mkdir(parents=True, exist_ok=True)
    m.wb.save(OUT)
    print(f"Saved {OUT}")


if __name__ == "__main__":
    main()
