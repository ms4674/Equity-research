"""Cross-check the generated workbook's Excel formulas against an independent
Python implementation of the model (build_ai_capex_model.print_expected_values
logic). Exits non-zero on any mismatch > 1e-9."""

import sys

import formulas

from build_ai_capex_model import (CAPEX_BASE_2025_BN, N_SEG, SCENARIOS,
                                  SEG_FIRST_ROW, SEGMENTS,
                                  TRAINING_SHARE_2025)

FILE = "AI_Capex_Cut_Scenario_Model.xlsx"

xl = formulas.ExcelModel().loads(FILE).finish()
sol = xl.calculate()

def cell(sheet, ref):
    key = f"'[{FILE}]{sheet.upper()}'!{ref}"
    return sol[key].value[0, 0]


errors = []


def check(label, sheet, ref, expected):
    got = float(cell(sheet, ref))
    if abs(got - expected) > 1e-9:
        errors.append(f"MISMATCH {label} ({sheet}!{ref}): excel={got} expected={expected}")


t_share = TRAINING_SHARE_2025
base_t, base_i = SCENARIOS[0][1], SCENARIOS[0][2]

# Scenario Assumptions: blended growth and implied capex
for s_idx, (name, g_t, g_i, _) in enumerate(SCENARIOS):
    r = 9 + s_idx
    blended = t_share * g_t + (1 - t_share) * g_i
    check(f"{name} blended", "Scenario Assumptions", f"D{r}", blended)
    check(f"{name} capex $", "Scenario Assumptions", f"E{r}",
          CAPEX_BASE_2025_BN * (1 + blended))

# Per-segment scenario sales growth and EPS growth
sales_growth_cols = ["D", "G", "J", "M"]
eps_growth_cols = ["C", "E", "G", "I"]
for i, seg in enumerate(SEGMENTS):
    name, _, expo, train, lev, g_sales, g_eps, _ = seg
    r = SEG_FIRST_ROW + i
    base_blend = train * base_t + (1 - train) * base_i
    for s_idx, (_sn, g_t, g_i, _n) in enumerate(SCENARIOS):
        blend = train * g_t + (1 - train) * g_i
        d_sales = expo * (blend - base_blend)
        check(f"{name} s{s_idx} sales", "Sales Growth Impact",
              f"{sales_growth_cols[s_idx]}{r}", g_sales + d_sales)
        check(f"{name} s{s_idx} eps", "EPS Growth Impact",
              f"{eps_growth_cols[s_idx]}{r}", g_eps + lev * d_sales)
        # Summary matrix mirrors the two sheets
        check(f"{name} s{s_idx} summary sales", "Summary Matrix",
              f"{chr(ord('B') + s_idx)}{5 + i}", g_sales + d_sales)
        check(f"{name} s{s_idx} summary eps", "Summary Matrix",
              f"{chr(ord('B') + s_idx)}{5 + N_SEG + 4 + i}", g_eps + lev * d_sales)

n_checks = 8 + N_SEG * len(SCENARIOS) * 4
if errors:
    print("\n".join(errors))
    sys.exit(1)
print(f"All {n_checks} formula checks passed.")
