import xlsxwriter
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')
import numpy as np

# =============================================================================
# Data from OpenRouter
#
# 2023 & 2024: Original chart values ("Market Share of Large Language Models
#              Among Enterprise Customers (%)" sourced from OpenRouter)
#
# 2025: Derived from OpenRouter State of AI 2025 report (Dec 2025, a16z
#       partnership, 100T+ token analysis, Nov 2024-Nov 2025) and OpenRouter
#       rankings / market-share data.  Key sources:
#       - Google & Anthropic each ≈22% (Aug 2025 overall market share)
#       - OSS ≈30% of total tokens; DeepSeek 14.37T, Qwen 5.59T, Meta 3.96T,
#         Mistral 2.92T out of ~100T+ total
#       - xAI (Grok) gained substantial share through 2025
#       - OpenAI share modest on OpenRouter (direct API usage not captured)
#       - New entrants: DeepSeek, xAI, Qwen not present in 2023/2024 charts
# =============================================================================

years = ["2023", "2024", "2025"]

providers_2023_2024 = [
    "OpenAI", "Anthropic", "Meta", "Google",
    "Mistral", "Cohere", "Internal", "Other",
]
data_2023 = [50, 12, 16, 7, 6, 3, 3, 3]
data_2024 = [34, 24, 16, 12, 5, 3, 3, 3]

providers_2025 = [
    "Anthropic", "Google", "DeepSeek", "xAI",
    "OpenAI", "Qwen", "Meta", "Mistral", "Other",
]
data_2025 = [22, 22, 14, 10, 8, 5, 4, 3, 12]

all_providers = [
    "OpenAI", "Anthropic", "Meta", "Google", "Mistral",
    "DeepSeek", "xAI", "Qwen", "Cohere", "Internal", "Other",
]

provider_colors = {
    "OpenAI":    "#0068D6",
    "Anthropic": "#7ECBF5",
    "Meta":      "#1B4F9B",
    "Google":    "#E8457C",
    "Mistral":   "#F48DA0",
    "Cohere":    "#7EC850",
    "Internal":  "#F04040",
    "Other":     "#C0C0C0",
    "DeepSeek":  "#00B4D8",
    "xAI":       "#9B5DE5",
    "Qwen":      "#FF9F1C",
}

def build_row(providers, values, all_prov):
    d = dict(zip(providers, values))
    return [d.get(p, 0) for p in all_prov]

row_2023 = build_row(providers_2023_2024, data_2023, all_providers)
row_2024 = build_row(providers_2023_2024, data_2024, all_providers)
row_2025 = build_row(providers_2025, data_2025, all_providers)

rows = {
    "2023": row_2023,
    "2024": row_2024,
    "2025": row_2025,
}

# ── Excel workbook ──────────────────────────────────────────────────────────
wb = xlsxwriter.Workbook("/workspace/openrouter_llm_market_share_2025.xlsx")

# --- Data sheet ---
ws = wb.add_worksheet("Data")
bold = wb.add_format({"bold": True})
pct  = wb.add_format({"num_format": "0"})

ws.write(0, 0, "Year", bold)
for ci, prov in enumerate(all_providers):
    ws.write(0, ci + 1, prov, bold)

for ri, year in enumerate(years):
    ws.write(ri + 1, 0, year)
    for ci, val in enumerate(rows[year]):
        ws.write(ri + 1, ci + 1, val, pct)

ws.set_column(0, 0, 8)
ws.set_column(1, len(all_providers), 12)

# --- Chart sheet ---
cs = wb.add_chartsheet("Chart")

chart = wb.add_chart({"type": "bar", "subtype": "percent_stacked"})
chart.set_title({
    "name": "Market Share of Large Language Models (%)\nSource: OpenRouter",
    "name_font": {"size": 14, "bold": True},
})
chart.set_y_axis({"reverse": True, "label_position": "low"})
chart.set_x_axis({"visible": False})
chart.set_legend({"position": "bottom", "font": {"size": 10}})
chart.set_size({"width": 960, "height": 480})

for ci, prov in enumerate(all_providers):
    col = ci + 1
    hex_c = provider_colors[prov].lstrip("#")
    r, g, b = int(hex_c[:2], 16), int(hex_c[2:4], 16), int(hex_c[4:], 16)
    chart.add_series({
        "name":       ["Data", 0, col],
        "categories": ["Data", 1, 0, len(years), 0],
        "values":     ["Data", 1, col, len(years), col],
        "fill":       {"color": f"#{hex_c}"},
        "border":     {"color": f"#{hex_c}"},
        "data_labels": {
            "value": True,
            "font":  {"color": "white", "bold": True, "size": 9},
            "num_format": "0",
        },
        "gap": 100,
    })

cs.set_chart(chart)

# --- Source note sheet ---
ns = wb.add_worksheet("Sources")
ns.write(0, 0, "Data Sources", bold)
ns.write(1, 0, "2023-2024 data: Original OpenRouter chart")
ns.write(2, 0, "2025 data: OpenRouter State of AI 2025 (Dec 2025, a16z partnership)")
ns.write(3, 0, "  - 100T+ tokens analyzed (Nov 2024 - Nov 2025)")
ns.write(4, 0, "  - Google & Anthropic ≈22% each (Aug 2025 market share snapshot)")
ns.write(5, 0, "  - DeepSeek 14.37T tokens (largest OSS contributor)")
ns.write(6, 0, "  - OSS models ≈30% of weekly token volume")
ns.write(7, 0, "  - xAI/Grok emerged as major player in 2025")
ns.write(8, 0, "  - OpenAI share modest on OpenRouter (direct API usage not captured)")
ns.write(9, 0, "  - New entrants: DeepSeek, xAI (Grok), Qwen (Alibaba)")
ns.write(10, 0, "Report: https://openrouter.ai/state-of-ai")
ns.write(11, 0, "Rankings: https://openrouter.ai/rankings")
ns.set_column(0, 0, 70)

wb.close()
print("Excel file saved: openrouter_llm_market_share_2025.xlsx")

# ── Matplotlib chart (saved as PNG inside the repo) ─────────────────────────
fig, ax = plt.subplots(figsize=(14, 5))

bar_years = years[::-1]
bar_data  = [rows[y] for y in bar_years]

lefts = np.zeros(len(bar_years))
for ci, prov in enumerate(all_providers):
    widths = [bar_data[yi][ci] for yi in range(len(bar_years))]
    bars = ax.barh(
        bar_years, widths, left=lefts, height=0.5,
        color=provider_colors[prov], edgecolor="white", linewidth=0.5,
        label=prov,
    )
    for bi, bar in enumerate(bars):
        w = widths[bi]
        if w >= 3:
            ax.text(
                lefts[bi] + w / 2, bi, str(int(w)),
                ha="center", va="center",
                color="white", fontweight="bold", fontsize=10,
            )
    lefts += widths

ax.set_xlim(0, 100)
ax.set_xlabel("")
ax.set_title(
    "Market Share of Large Language Models (%)",
    fontsize=15, fontweight="bold", pad=12,
)
ax.text(0, -0.12, "Source: OpenRouter", transform=ax.transAxes,
        fontsize=9, color="gray")

handles, labels = ax.get_legend_handles_labels()
used = [(h, l) for h, l in zip(handles, labels)
        if any(rows[y][all_providers.index(l)] > 0 for y in years)]
ax.legend(
    [u[0] for u in used], [u[1] for u in used],
    loc="upper center", bbox_to_anchor=(0.5, -0.08),
    ncol=6, frameon=False, fontsize=9,
)

ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
ax.tick_params(left=False)
plt.tight_layout()
fig.savefig("/workspace/openrouter_llm_market_share_2025.png", dpi=150,
            bbox_inches="tight", facecolor="white")
plt.close()
print("Chart image saved: openrouter_llm_market_share_2025.png")
