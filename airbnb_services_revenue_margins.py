"""
Airbnb Services Revenue & Margins Analysis
Comparing new Services/Experiences business vs Core Stays business
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
import os

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

COLORS = {
    "airbnb_pink": "#FF5A5F",
    "airbnb_dark": "#484848",
    "airbnb_teal": "#00A699",
    "airbnb_orange": "#FC642D",
    "airbnb_purple": "#914669",
    "light_gray": "#F5F5F5",
    "mid_gray": "#C4C4C4",
    "blue": "#3B82F6",
}

plt.rcParams.update({
    "font.family": "sans-serif",
    "font.size": 11,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "figure.facecolor": "white",
    "axes.facecolor": "white",
    "axes.grid": True,
    "grid.alpha": 0.3,
    "grid.linestyle": "--",
})


# =============================================================================
# 1. Consolidated P&L (2019-2025)
# =============================================================================
def chart_consolidated_pl():
    years = [2019, 2020, 2021, 2022, 2023, 2024, 2025]
    revenue = [4805, 3378, 5992, 8399, 9917, 11102, 12241]
    adj_ebitda = [None, None, None, 2900, 3500, 4000, 4300]
    net_income = [None, -4585, -352, 1893, 4792, 2648, 2510]
    fcf = [None, None, None, 3400, 4200, 4500, 4650]

    fig, ax1 = plt.subplots(figsize=(12, 6))

    x = np.arange(len(years))
    width = 0.25

    bars_rev = ax1.bar(x - width, revenue, width, label="Revenue",
                       color=COLORS["airbnb_pink"], zorder=3)

    ebitda_vals = [v if v else 0 for v in adj_ebitda]
    ebitda_mask = [v is not None for v in adj_ebitda]
    bars_ebitda = ax1.bar(x[ebitda_mask], [ebitda_vals[i] for i in range(len(years)) if ebitda_mask[i]],
                          width, label="Adj. EBITDA",
                          color=COLORS["airbnb_teal"], zorder=3)

    fcf_vals = [v if v else 0 for v in fcf]
    fcf_mask = [v is not None for v in fcf]
    bars_fcf = ax1.bar(x[fcf_mask] + width, [fcf_vals[i] for i in range(len(years)) if fcf_mask[i]],
                       width, label="Free Cash Flow",
                       color=COLORS["blue"], zorder=3)

    ax1.set_xlabel("Fiscal Year")
    ax1.set_ylabel("$ Millions")
    ax1.set_title("Airbnb Consolidated Financials (2019-2025)", fontsize=14, fontweight="bold", pad=15)
    ax1.set_xticks(x)
    ax1.set_xticklabels(years)
    ax1.legend(loc="upper left", framealpha=0.9)
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v/1000:.1f}B"))

    for bar in bars_rev:
        ax1.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 100,
                 f"${bar.get_height()/1000:.1f}B", ha="center", va="bottom", fontsize=7.5)

    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/01_consolidated_pl.png", dpi=150)
    plt.close()
    print("  -> 01_consolidated_pl.png")


# =============================================================================
# 2. Margin Trends
# =============================================================================
def chart_margin_trends():
    years = [2022, 2023, 2024, 2025]
    gross_margin = [82.2, 82.8, 83.1, 83.0]
    ebitda_margin = [34.5, 35.3, 36.0, 35.1]
    ni_margin = [22.5, 48.3, 23.9, 20.5]
    fcf_margin = [40.5, 42.3, 40.5, 38.0]

    fig, ax = plt.subplots(figsize=(10, 6))

    ax.plot(years, gross_margin, "o-", color=COLORS["airbnb_pink"], linewidth=2.5,
            markersize=8, label="Gross Margin")
    ax.plot(years, ebitda_margin, "s-", color=COLORS["airbnb_teal"], linewidth=2.5,
            markersize=8, label="Adj. EBITDA Margin")
    ax.plot(years, fcf_margin, "^-", color=COLORS["blue"], linewidth=2.5,
            markersize=8, label="FCF Margin")
    ax.plot(years, ni_margin, "D-", color=COLORS["airbnb_orange"], linewidth=2.5,
            markersize=8, label="Net Income Margin")

    for vals, offset in [(gross_margin, 1.2), (ebitda_margin, -2.5), (fcf_margin, 1.2), (ni_margin, -2.5)]:
        for yr, v in zip(years, vals):
            ax.annotate(f"{v:.1f}%", (yr, v), textcoords="offset points",
                        xytext=(0, offset * 5), ha="center", fontsize=8)

    ax.set_xlabel("Fiscal Year")
    ax.set_ylabel("Margin (%)")
    ax.set_title("Airbnb Margin Trends (2022-2025)", fontsize=14, fontweight="bold", pad=15)
    ax.set_xticks(years)
    ax.set_ylim(10, 95)
    ax.legend(loc="upper right", framealpha=0.9)
    ax.yaxis.set_major_formatter(mticker.PercentFormatter())

    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/02_margin_trends.png", dpi=150)
    plt.close()
    print("  -> 02_margin_trends.png")


# =============================================================================
# 3. Take Rate & GBV Analysis
# =============================================================================
def chart_take_rate_gbv():
    years = [2022, 2023, 2024, 2025]
    gbv = [63.2, 73.3, 81.8, 91.3]
    revenue = [8.40, 9.92, 11.10, 12.24]
    take_rate = [r / g * 100 for r, g in zip(revenue, gbv)]

    fig, ax1 = plt.subplots(figsize=(10, 6))

    ax1.bar(years, gbv, width=0.5, color=COLORS["mid_gray"], alpha=0.7, label="GBV ($B)", zorder=2)
    ax1.set_ylabel("Gross Booking Value ($B)")
    ax1.set_xlabel("Fiscal Year")
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:.0f}B"))

    for yr, v in zip(years, gbv):
        ax1.text(yr, v + 1, f"${v:.1f}B", ha="center", fontsize=9)

    ax2 = ax1.twinx()
    ax2.plot(years, take_rate, "o-", color=COLORS["airbnb_pink"], linewidth=2.5,
             markersize=10, label="Implied Take Rate", zorder=5)
    ax2.set_ylabel("Take Rate (%)")
    ax2.yaxis.set_major_formatter(mticker.PercentFormatter())

    for yr, v in zip(years, take_rate):
        ax2.annotate(f"{v:.1f}%", (yr, v), textcoords="offset points",
                     xytext=(0, 12), ha="center", fontsize=9, fontweight="bold",
                     color=COLORS["airbnb_pink"])

    ax2.set_ylim(10, 16)
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left", framealpha=0.9)

    ax1.set_title("Airbnb GBV & Implied Take Rate", fontsize=14, fontweight="bold", pad=15)
    ax1.set_xticks(years)

    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/03_take_rate_gbv.png", dpi=150)
    plt.close()
    print("  -> 03_take_rate_gbv.png")


# =============================================================================
# 4. Services vs Core Business — Fee Structure & Economics
# =============================================================================
def chart_services_vs_core_fees():
    categories = ["Core Stays\n(Host-only model)", "Core Stays\n(Split-fee model)", "Experiences", "Services"]
    host_fee = [15.5, 3, 20, 15]
    guest_fee = [0, 14.5, 0, 0]
    total_take = [15.5, 17.5, 20, 15]

    x = np.arange(len(categories))
    width = 0.35

    fig, ax = plt.subplots(figsize=(10, 6))

    bars1 = ax.bar(x - width / 2, host_fee, width, label="Host Fee (%)",
                   color=COLORS["airbnb_pink"], zorder=3)
    bars2 = ax.bar(x + width / 2, guest_fee, width, label="Guest Fee (%)",
                   color=COLORS["airbnb_teal"], zorder=3)

    for i, t in enumerate(total_take):
        ax.annotate(f"Total: {t}%", (x[i], max(host_fee[i], guest_fee[i]) + 1.5),
                    ha="center", fontsize=10, fontweight="bold", color=COLORS["airbnb_dark"])

    ax.set_ylabel("Fee Rate (%)")
    ax.set_title("Airbnb Fee Structure: Core Stays vs Services & Experiences",
                 fontsize=14, fontweight="bold", pad=15)
    ax.set_xticks(x)
    ax.set_xticklabels(categories)
    ax.legend(loc="upper right", framealpha=0.9)
    ax.set_ylim(0, 28)
    ax.yaxis.set_major_formatter(mticker.PercentFormatter())

    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/04_fee_structure_comparison.png", dpi=150)
    plt.close()
    print("  -> 04_fee_structure_comparison.png")


# =============================================================================
# 5. Services & Experiences Revenue Trajectory
# =============================================================================
def chart_services_revenue_trajectory():
    periods = ["H2 2025\n(launch period)", "Q1 2026\n(annualized)"]
    revenue_est = [0.6, 1.2]

    pct_total = [revenue_est[0] / 6.2 * 100, 1.2 / 13.5 * 100]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))

    bars = ax1.bar(periods, revenue_est, width=0.4,
                   color=[COLORS["airbnb_orange"], COLORS["airbnb_pink"]], zorder=3)
    ax1.set_ylabel("Revenue ($B, annualized est.)")
    ax1.set_title("Services & Experiences Revenue", fontsize=13, fontweight="bold", pad=10)
    for bar, val in zip(bars, revenue_est):
        ax1.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.03,
                 f"~${val:.1f}B", ha="center", fontsize=11, fontweight="bold")
    ax1.set_ylim(0, 1.8)

    colors_pie = [COLORS["airbnb_pink"], COLORS["airbnb_teal"]]
    services_pct = 8.9
    core_pct = 100 - services_pct
    ax2.pie([core_pct, services_pct], labels=["Core Stays", "Services &\nExperiences"],
            colors=colors_pie, autopct="%1.1f%%", startangle=90,
            textprops={"fontsize": 11}, pctdistance=0.75)
    ax2.set_title("Est. Revenue Mix (Q1 2026 annualized)", fontsize=13, fontweight="bold", pad=10)

    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/05_services_revenue_trajectory.png", dpi=150)
    plt.close()
    print("  -> 05_services_revenue_trajectory.png")


# =============================================================================
# 6. Margin Impact — Investment vs Returns
# =============================================================================
def chart_margin_impact():
    quarters = ["Q1'25", "Q2'25", "Q3'25", "Q4'25"]
    ebitda_margin_2024 = [33.0, 33.0, 37.0, 28.5]
    ebitda_margin_2025 = [34.5, 34.0, 36.0, 28.0]

    fig, ax = plt.subplots(figsize=(10, 6))

    x = np.arange(len(quarters))
    width = 0.3

    bars1 = ax.bar(x - width / 2, ebitda_margin_2024, width, label="FY2024 Adj. EBITDA Margin",
                   color=COLORS["mid_gray"], zorder=3)
    bars2 = ax.bar(x + width / 2, ebitda_margin_2025, width, label="FY2025 Adj. EBITDA Margin",
                   color=COLORS["airbnb_teal"], zorder=3)

    for bar in bars1:
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
                f"{bar.get_height():.1f}%", ha="center", fontsize=8.5, color=COLORS["airbnb_dark"])
    for bar in bars2:
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
                f"{bar.get_height():.1f}%", ha="center", fontsize=8.5, color=COLORS["airbnb_dark"])

    ax.set_ylabel("Adj. EBITDA Margin (%)")
    ax.set_title("Quarterly Adj. EBITDA Margin: Services Investment Impact",
                 fontsize=14, fontweight="bold", pad=15)
    ax.set_xticks(x)
    ax.set_xticklabels(quarters)
    ax.legend(loc="upper left", framealpha=0.9)
    ax.set_ylim(20, 42)
    ax.yaxis.set_major_formatter(mticker.PercentFormatter())

    ax.annotate("~$200M invested in\nServices & Experiences\nin FY2025",
                xy=(2, 36), xytext=(2.5, 40),
                arrowprops=dict(arrowstyle="->", color=COLORS["airbnb_orange"], lw=1.5),
                fontsize=9, ha="center", color=COLORS["airbnb_orange"], fontweight="bold",
                bbox=dict(boxstyle="round,pad=0.4", facecolor="lightyellow", edgecolor=COLORS["airbnb_orange"]))

    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/06_margin_impact.png", dpi=150)
    plt.close()
    print("  -> 06_margin_impact.png")


# =============================================================================
# 7. Services Unit Economics Comparison
# =============================================================================
def chart_unit_economics():
    categories = ["Core Stays\n(Marketplace)", "Experiences\n(Curated)", "Services\n(Managed)"]
    take_rate = [15.5, 20.0, 15.0]
    est_gross_margin = [85, 75, 65]
    est_contribution_margin = [55, 30, 15]

    x = np.arange(len(categories))
    width = 0.22

    fig, ax = plt.subplots(figsize=(10, 6))

    bars1 = ax.bar(x - width, take_rate, width, label="Take Rate",
                   color=COLORS["airbnb_pink"], zorder=3)
    bars2 = ax.bar(x, est_gross_margin, width, label="Est. Gross Margin",
                   color=COLORS["airbnb_teal"], zorder=3)
    bars3 = ax.bar(x + width, est_contribution_margin, width, label="Est. Contribution Margin",
                   color=COLORS["blue"], zorder=3)

    for bars in [bars1, bars2, bars3]:
        for bar in bars:
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.8,
                    f"{bar.get_height():.0f}%", ha="center", fontsize=9, fontweight="bold")

    ax.set_ylabel("Percentage (%)")
    ax.set_title("Estimated Unit Economics: Core vs Services vs Experiences",
                 fontsize=14, fontweight="bold", pad=15)
    ax.set_xticks(x)
    ax.set_xticklabels(categories)
    ax.legend(loc="upper right", framealpha=0.9)
    ax.set_ylim(0, 100)
    ax.yaxis.set_major_formatter(mticker.PercentFormatter())

    ax.text(0.02, 0.02, "Note: Gross and contribution margins for Experiences and Services are estimates\n"
            "based on industry benchmarks; Airbnb does not disclose segment-level margins.",
            transform=ax.transAxes, fontsize=7.5, style="italic", color="gray", va="bottom")

    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/07_unit_economics.png", dpi=150)
    plt.close()
    print("  -> 07_unit_economics.png")


# =============================================================================
# 8. Experiences Growth Metrics
# =============================================================================
def chart_experiences_growth():
    quarters = ["Q2'25\n(Launch)", "Q3'25", "Q4'25", "Q1'26"]
    bookings_m = [15, 25, 35, 45]
    pct_total_bookings = [4, 8, 12, 15]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))

    ax1.bar(quarters, bookings_m, color=COLORS["airbnb_pink"], width=0.5, zorder=3)
    ax1.set_ylabel("Bookings (Millions)")
    ax1.set_title("Experiences & Services Bookings", fontsize=13, fontweight="bold", pad=10)
    for i, v in enumerate(bookings_m):
        ax1.text(i, v + 0.8, f"{v}M", ha="center", fontsize=10, fontweight="bold")
    ax1.set_ylim(0, 55)

    ax2.plot(quarters, pct_total_bookings, "o-", color=COLORS["airbnb_teal"],
             linewidth=2.5, markersize=10, zorder=5)
    ax2.fill_between(range(len(quarters)), pct_total_bookings,
                     alpha=0.15, color=COLORS["airbnb_teal"])
    ax2.set_ylabel("% of Total Platform Bookings")
    ax2.set_title("Services & Experiences as % of Total Bookings",
                  fontsize=13, fontweight="bold", pad=10)
    for i, v in enumerate(pct_total_bookings):
        ax2.annotate(f"{v}%", (i, v), textcoords="offset points",
                     xytext=(0, 10), ha="center", fontsize=10, fontweight="bold")
    ax2.set_ylim(0, 22)
    ax2.yaxis.set_major_formatter(mticker.PercentFormatter())

    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/08_experiences_growth.png", dpi=150)
    plt.close()
    print("  -> 08_experiences_growth.png")


# =============================================================================
# 9. Opex Breakdown
# =============================================================================
def chart_opex_breakdown():
    years = [2023, 2024, 2025]

    revenue = [9917, 11102, 12241]
    cost_rev = [1702, 1878, 2086]
    ops_support = [1186, 1282, 1327]
    product_dev = [1722, 2056, 2354]
    sales_mktg = [1763, 2148, 2588]
    gen_admin = [2025, 1184, 1342]

    fig, ax = plt.subplots(figsize=(12, 6))

    x = np.arange(len(years))
    width = 0.5

    p1 = ax.bar(x, cost_rev, width, label="Cost of Revenue", color=COLORS["airbnb_dark"], zorder=3)
    p2 = ax.bar(x, ops_support, width, bottom=cost_rev, label="Ops & Support",
                color=COLORS["airbnb_purple"], zorder=3)

    bottom2 = [a + b for a, b in zip(cost_rev, ops_support)]
    p3 = ax.bar(x, product_dev, width, bottom=bottom2, label="Product Development",
                color=COLORS["airbnb_teal"], zorder=3)

    bottom3 = [a + b for a, b in zip(bottom2, product_dev)]
    p4 = ax.bar(x, sales_mktg, width, bottom=bottom3, label="Sales & Marketing",
                color=COLORS["airbnb_orange"], zorder=3)

    bottom4 = [a + b for a, b in zip(bottom3, sales_mktg)]
    p5 = ax.bar(x, gen_admin, width, bottom=bottom4, label="General & Administrative",
                color=COLORS["mid_gray"], zorder=3)

    for i, rev in enumerate(revenue):
        total_opex = cost_rev[i] + ops_support[i] + product_dev[i] + sales_mktg[i] + gen_admin[i]
        ax.plot([x[i] - 0.3, x[i] + 0.3], [rev, rev], color=COLORS["airbnb_pink"],
                linewidth=3, zorder=5)
        if i == 0:
            ax.plot([], [], color=COLORS["airbnb_pink"], linewidth=3, label="Revenue")

    ax.set_ylabel("$ Millions")
    ax.set_title("Airbnb Operating Expense Breakdown vs Revenue",
                 fontsize=14, fontweight="bold", pad=15)
    ax.set_xticks(x)
    ax.set_xticklabels(years)
    ax.legend(loc="upper left", framealpha=0.9, fontsize=9)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v/1000:.1f}B"))

    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/09_opex_breakdown.png", dpi=150)
    plt.close()
    print("  -> 09_opex_breakdown.png")


# =============================================================================
# 10. Strategic Revenue Potential (2026E-2028E)
# =============================================================================
def chart_revenue_potential():
    years = [2024, 2025, "2026E", "2027E", "2028E"]
    core_stays = [11.10, 11.64, 12.80, 14.08, 15.20]
    services_exp = [0, 0.60, 1.50, 2.80, 4.50]
    total = [a + b for a, b in zip(core_stays, services_exp)]

    fig, ax = plt.subplots(figsize=(11, 6))

    x = np.arange(len(years))
    width = 0.5

    ax.bar(x, core_stays, width, label="Core Stays", color=COLORS["airbnb_pink"], zorder=3)
    ax.bar(x, services_exp, width, bottom=core_stays, label="Services & Experiences",
           color=COLORS["airbnb_teal"], zorder=3)

    for i in range(len(years)):
        if services_exp[i] > 0:
            pct = services_exp[i] / total[i] * 100
            ax.text(x[i], total[i] + 0.2, f"${total[i]:.1f}B\n({pct:.0f}% S&E)",
                    ha="center", fontsize=8.5, fontweight="bold")
        else:
            ax.text(x[i], total[i] + 0.2, f"${total[i]:.1f}B",
                    ha="center", fontsize=8.5, fontweight="bold")

    ax.axvline(x=1.5, color=COLORS["mid_gray"], linestyle=":", linewidth=1.2, alpha=0.7)
    ax.text(1.7, 18, "Estimates →", fontsize=9, style="italic", color="gray")

    ax.set_ylabel("Revenue ($B)")
    ax.set_title("Airbnb Revenue Potential: Core Stays + Services & Experiences",
                 fontsize=14, fontweight="bold", pad=15)
    ax.set_xticks(x)
    ax.set_xticklabels(years)
    ax.legend(loc="upper left", framealpha=0.9)
    ax.set_ylim(0, 22)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:.0f}B"))

    plt.tight_layout()
    plt.savefig(f"{OUTPUT_DIR}/10_revenue_potential.png", dpi=150)
    plt.close()
    print("  -> 10_revenue_potential.png")


# =============================================================================
# Main
# =============================================================================
if __name__ == "__main__":
    print("Generating Airbnb Services Revenue & Margins Analysis charts...\n")

    chart_consolidated_pl()
    chart_margin_trends()
    chart_take_rate_gbv()
    chart_services_vs_core_fees()
    chart_services_revenue_trajectory()
    chart_margin_impact()
    chart_unit_economics()
    chart_experiences_growth()
    chart_opex_breakdown()
    chart_revenue_potential()

    print("\nAll charts saved to output/")
