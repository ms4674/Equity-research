"""Recalculate the workbook with the `formulas` engine and report key outputs
and the integrity checks (balance-sheet balance + cash tie-out)."""
import formulas

XL = "SpaceX_Cursor_Pro_Forma_Model.xlsx"
xl = formulas.ExcelModel().loads(XL).finish()
sol = xl.calculate()

# Build a lookup that is tolerant of the funky key format formulas uses.
vals = {}
for k, v in sol.items():
    vals[k.upper()] = v


def cell(sheet, a1):
    key = f"'[{XL.upper()}]{sheet.upper()}'!{a1.upper()}"
    v = vals.get(key)
    try:
        return float(v.value[0, 0])
    except Exception:
        try:
            return v.value
        except Exception:
            return v


YEARS = [2025, 2026, 2027, 2028, 2029, 2030]
COLS_ISCF = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G", 2030: "H"}
BS = {2025: "C", "Open": "F", 2026: "G", 2027: "H", 2028: "I", 2029: "J", 2030: "K"}


def fmt(x):
    try:
        return f"{x:>12,.0f}"
    except Exception:
        return f"{str(x):>12}"


print("\n=== INTEGRITY CHECKS (Checks tab) ===")
# Checks rows: bs row, cf row
# find by reading the Checks sheet directly via openpyxl for row numbers
import openpyxl
wb = openpyxl.load_workbook(XL)
chk = wb["Checks"]
rows = {}
for r in range(1, chk.max_row + 1):
    lab = chk.cell(row=r, column=1).value
    if lab:
        rows[lab] = r

bs_row = next(r for l, r in rows.items() if l.startswith("Balance sheet balances"))
cf_row = next(r for l, r in rows.items() if l.startswith("Cash flow ties"))
ppa_row = next(r for l, r in rows.items() if l.startswith("PPA bridge"))

print("Year        " + "".join(f"{y:>12}" for y in YEARS))
print("BS balance: " + "".join(fmt(cell("Checks", f"{COLS_ISCF[y]}{bs_row}")) for y in YEARS))
print("CF tie:     " + "".join(fmt(cell("Checks", f"{COLS_ISCF[y]}{cf_row}")) for y in YEARS))
print("PPA open:   " + fmt(cell("Checks", f"C{ppa_row}")))


def is_row(label):
    s = wb["Income Statement"]
    for r in range(1, s.max_row + 1):
        if s.cell(row=r, column=1).value == label:
            return r


print("\n=== PRO-FORMA INCOME STATEMENT (US$M) ===")
print("Year             " + "".join(f"{y:>12}" for y in YEARS))
for label in ["Total revenue", "Gross profit", "EBITDA", "EBIT (operating income)",
              "Pre-tax income (loss)", "Net income (loss)"]:
    r = is_row(label)
    print(f"{label:<16} " + "".join(fmt(cell("Income Statement", f"{COLS_ISCF[y]}{r}")) for y in YEARS))


def bs_rownum(label):
    s = wb["Balance Sheet"]
    for r in range(1, s.max_row + 1):
        if s.cell(row=r, column=1).value == label:
            return r


print("\n=== BALANCE SHEET (US$M) ===")
hdr_years = ["2025A", "PF Open", 2026, 2027, 2028, 2029, 2030]
hdr_cols = ["C", "F", "G", "H", "I", "J", "K"]
print("Line                 " + "".join(f"{str(h):>12}" for h in hdr_years))
for label in ["Cash, equivalents & ST investments", "Goodwill", "TOTAL ASSETS",
              "TOTAL LIABILITIES", "TOTAL EQUITY", "TOTAL LIABILITIES & EQUITY"]:
    r = bs_rownum(label)
    print(f"{label[:20]:<20} " + "".join(fmt(cell("Balance Sheet", f"{c}{r}")) for c in hdr_cols))


def cf_rownum(label):
    s = wb["Cash Flow"]
    for r in range(1, s.max_row + 1):
        if s.cell(row=r, column=1).value == label:
            return r


print("\n=== CASH FLOW (US$M) ===")
print("Line                 " + "".join(f"{y:>12}" for y in YEARS[1:]))
for label in ["Cash from operations", "Cash from investing", "Cash from financing",
              "Net change in cash", "Cash, end of period"]:
    r = cf_rownum(label)
    print(f"{label[:20]:<20} " + "".join(fmt(cell("Cash Flow", f"{COLS_ISCF[y]}{r}")) for y in YEARS[1:]))
