#!/usr/bin/env python3
"""Aggregate global marketplace bookings from data/marketplace_bookings_2025.csv.

Prints per-sector subtotals (split by disclosure basis), the grand total, and an
adjusted online-travel subtotal that strips out B2B bed banks (HBX, WebBeds)
whose volumes are largely resold through the retail OTAs already counted.
"""

import csv
from collections import defaultdict
from pathlib import Path

DATA = Path(__file__).resolve().parent.parent / "data" / "marketplace_bookings_2025.csv"

# B2B wholesalers whose TTV overlaps retail OTA gross bookings.
B2B_BED_BANKS = {"HBX Group (Hotelbeds)", "WEB Travel Group (WebBeds)"}

DISCLOSED_BASES = {"disclosed", "disclosed (stale)", "company-announced", "management-stated approx"}


def main() -> None:
    sectors: dict[str, dict[str, float]] = defaultdict(lambda: {"disclosed": 0.0, "estimate": 0.0})
    bed_bank_total = 0.0

    with open(DATA, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row["in_total"] != "yes" or not row["value_usd_bn"]:
                continue
            value = float(row["value_usd_bn"])
            bucket = "disclosed" if row["basis"] in DISCLOSED_BASES else "estimate"
            sectors[row["sector"]][bucket] += value
            if row["company"] in B2B_BED_BANKS:
                bed_bank_total += value

    grand_total = 0.0
    print(f"{'Sector':<16}{'Disclosed':>12}{'Estimated':>12}{'Subtotal':>12}")
    print("-" * 52)
    for sector, buckets in sectors.items():
        subtotal = buckets["disclosed"] + buckets["estimate"]
        grand_total += subtotal
        print(f"{sector:<16}{buckets['disclosed']:>11.1f} {buckets['estimate']:>11.1f} {subtotal:>11.1f}")
    print("-" * 52)
    print(f"{'Grand total':<16}{'':>12}{'':>12}{grand_total:>11.1f}")
    print()
    travel = sectors["Online travel"]
    travel_adj = travel["disclosed"] + travel["estimate"] - bed_bank_total
    print(f"Online travel ex-B2B bed banks (-{bed_bank_total:.1f}): {travel_adj:.1f}")
    print(f"Grand total ex-B2B bed banks: {grand_total - bed_bank_total:.1f}")


if __name__ == "__main__":
    main()
