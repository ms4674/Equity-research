"""Token price history dataset: frontier and open-weight LLMs, 2020-2026.

Each record is one price observation (launch price or subsequent price change)
for one model, at the vendor's standard list rate per 1 million tokens
(base/short-context tier, no cache, batch, or priority discounts).

Blended price is computed downstream as (3 * input + 1 * output) / 4,
the standard 3:1 input:output weighting used by industry trackers
(Artificial Analysis, ModelPriceWatch, AIMultiple).

Fields:
    developer, country, model, model_class, tier,
    date (ISO), date_precision ('exact' | 'approx'),
    event, input_usd_per_m, output_usd_per_m, basis, notes

Compiled 2026-08-29 from vendor pricing pages, launch announcements, and
third-party price trackers. See the Sources sheet of the workbook.
"""

FRONTIER = "Frontier (proprietary)"
OPEN = "Open-weight"

# (developer, country, model, class, tier, date, precision, event, input, output, basis, notes)
RECORDS = [
    # ------------------------------------------------------------------ OpenAI
    ("OpenAI", "USA", "GPT-3 (davinci)", FRONTIER, "Flagship", "2020-09-01", "approx", "Launch", 60.00, 60.00, "OpenAI API", "Original GPT-3 pricing, $0.06/1K tokens flat (no input/output split)"),
    ("OpenAI", "USA", "GPT-3.5 Turbo", FRONTIER, "Budget", "2023-03-01", "exact", "Launch", 2.00, 2.00, "OpenAI API", "$0.002/1K flat at launch"),
    ("OpenAI", "USA", "GPT-3.5 Turbo", FRONTIER, "Budget", "2023-06-13", "exact", "Price cut", 1.50, 2.00, "OpenAI API", "25% input price cut"),
    ("OpenAI", "USA", "GPT-3.5 Turbo", FRONTIER, "Budget", "2023-11-06", "exact", "Price cut", 1.00, 2.00, "OpenAI API", "gpt-3.5-turbo-1106"),
    ("OpenAI", "USA", "GPT-3.5 Turbo", FRONTIER, "Budget", "2024-01-25", "exact", "Price cut", 0.50, 1.50, "OpenAI API", "gpt-3.5-turbo-0125"),
    ("OpenAI", "USA", "GPT-4", FRONTIER, "Flagship", "2023-03-14", "exact", "Launch", 30.00, 60.00, "OpenAI API", "8K context tier; 32K tier was $60/$120"),
    ("OpenAI", "USA", "GPT-4 Turbo", FRONTIER, "Flagship", "2023-11-06", "exact", "Launch", 10.00, 30.00, "OpenAI API", "128K context"),
    ("OpenAI", "USA", "GPT-4o", FRONTIER, "Flagship", "2024-05-13", "exact", "Launch", 5.00, 15.00, "OpenAI API", ""),
    ("OpenAI", "USA", "GPT-4o", FRONTIER, "Flagship", "2024-08-06", "exact", "Price cut", 2.50, 10.00, "OpenAI API", "gpt-4o-2024-08-06"),
    ("OpenAI", "USA", "GPT-4o mini", FRONTIER, "Budget", "2024-07-18", "exact", "Launch", 0.15, 0.60, "OpenAI API", ""),
    ("OpenAI", "USA", "o1-preview", FRONTIER, "Reasoning", "2024-09-12", "exact", "Launch", 15.00, 60.00, "OpenAI API", "Reasoning tokens billed as output"),
    ("OpenAI", "USA", "o1-mini", FRONTIER, "Reasoning budget", "2024-09-12", "exact", "Launch", 3.00, 12.00, "OpenAI API", ""),
    ("OpenAI", "USA", "o1", FRONTIER, "Reasoning", "2024-12-17", "exact", "Launch", 15.00, 60.00, "OpenAI API", ""),
    ("OpenAI", "USA", "o3-mini", FRONTIER, "Reasoning budget", "2025-01-31", "exact", "Launch", 1.10, 4.40, "OpenAI API", ""),
    ("OpenAI", "USA", "GPT-4.5 (preview)", FRONTIER, "Premium", "2025-02-27", "exact", "Launch", 75.00, 150.00, "OpenAI API", "Research preview; retired July 2025"),
    ("OpenAI", "USA", "o1-pro", FRONTIER, "Premium reasoning", "2025-03-19", "exact", "Launch", 150.00, 600.00, "OpenAI API", "Most expensive list price to date"),
    ("OpenAI", "USA", "GPT-4.1", FRONTIER, "Mid", "2025-04-14", "exact", "Launch", 2.00, 8.00, "OpenAI API", "1M context"),
    ("OpenAI", "USA", "GPT-4.1 mini", FRONTIER, "Budget", "2025-04-14", "exact", "Launch", 0.40, 1.60, "OpenAI API", ""),
    ("OpenAI", "USA", "GPT-4.1 nano", FRONTIER, "Tiny", "2025-04-14", "exact", "Launch", 0.10, 0.40, "OpenAI API", ""),
    ("OpenAI", "USA", "o3", FRONTIER, "Reasoning", "2025-04-16", "exact", "Launch", 10.00, 40.00, "OpenAI API", ""),
    ("OpenAI", "USA", "o3", FRONTIER, "Reasoning", "2025-06-10", "exact", "Price cut", 2.00, 8.00, "OpenAI API", "80% cut"),
    ("OpenAI", "USA", "o4-mini", FRONTIER, "Reasoning budget", "2025-04-16", "exact", "Launch", 1.10, 4.40, "OpenAI API", ""),
    ("OpenAI", "USA", "o3-pro", FRONTIER, "Premium reasoning", "2025-06-10", "exact", "Launch", 20.00, 80.00, "OpenAI API", ""),
    ("OpenAI", "USA", "GPT-5", FRONTIER, "Flagship", "2025-08-07", "exact", "Launch", 1.25, 10.00, "OpenAI API", "400K context"),
    ("OpenAI", "USA", "GPT-5 mini", FRONTIER, "Budget", "2025-08-07", "exact", "Launch", 0.25, 2.00, "OpenAI API", ""),
    ("OpenAI", "USA", "GPT-5 nano", FRONTIER, "Tiny", "2025-08-07", "exact", "Launch", 0.05, 0.40, "OpenAI API", ""),
    ("OpenAI", "USA", "GPT-5.1", FRONTIER, "Flagship", "2025-11-12", "exact", "Launch", 1.25, 10.00, "OpenAI API", "Same pricing as GPT-5"),
    ("OpenAI", "USA", "GPT-5.2", FRONTIER, "Flagship", "2025-12-11", "exact", "Launch", 1.75, 14.00, "OpenAI API", ""),
    ("OpenAI", "USA", "GPT-5.2 Pro", FRONTIER, "Premium reasoning", "2025-12-11", "exact", "Launch", 21.00, 168.00, "OpenAI API", ""),
    ("OpenAI", "USA", "GPT-5.4", FRONTIER, "Flagship", "2026-03-05", "exact", "Launch", 2.50, 15.00, "OpenAI API", "First mainline model with native computer use; 1M context option"),
    ("OpenAI", "USA", "GPT-5.4 Pro", FRONTIER, "Premium reasoning", "2026-03-05", "exact", "Launch", 30.00, 180.00, "OpenAI API", ""),
    ("OpenAI", "USA", "GPT-5.4 mini", FRONTIER, "Budget", "2026-03-17", "exact", "Launch", 0.75, 4.50, "OpenAI API", ""),
    ("OpenAI", "USA", "GPT-5.4 nano", FRONTIER, "Tiny", "2026-03-17", "exact", "Launch", 0.20, 1.25, "OpenAI API", ""),
    ("OpenAI", "USA", "GPT-5.5", FRONTIER, "Flagship", "2026-04-24", "exact", "Launch", 5.00, 30.00, "OpenAI API", "API availability 2026-04-24 (announced 04-23)"),
    ("OpenAI", "USA", "GPT-5.5 Pro", FRONTIER, "Premium reasoning", "2026-04-24", "exact", "Launch", 30.00, 180.00, "OpenAI API", ""),
    ("OpenAI", "USA", "GPT-5.6 Sol", FRONTIER, "Flagship", "2026-07-09", "exact", "Launch", 5.00, 30.00, "OpenAI API", "GA 2026-07-09 after limited preview from 06-26; 1.05M context"),
    ("OpenAI", "USA", "GPT-5.6 Sol", FRONTIER, "Flagship", "2026-08-21", "exact", "Promotional cut", 4.00, 20.00, "OpenAI API", "Promo rate through at least 2026-11-21; list remains $5/$30"),
    ("OpenAI", "USA", "GPT-5.6 Terra", FRONTIER, "Mid", "2026-07-09", "exact", "Launch", 2.50, 15.00, "OpenAI API", ""),
    ("OpenAI", "USA", "GPT-5.6 Terra", FRONTIER, "Mid", "2026-07-30", "exact", "Price cut", 2.00, 12.00, "OpenAI API", "20% cut"),
    ("OpenAI", "USA", "GPT-5.6 Luna", FRONTIER, "Budget", "2026-07-09", "exact", "Launch", 1.00, 6.00, "OpenAI API", ""),
    ("OpenAI", "USA", "GPT-5.6 Luna", FRONTIER, "Budget", "2026-07-30", "exact", "Price cut", 0.20, 1.20, "OpenAI API", "80% cut"),
    ("OpenAI", "USA", "gpt-oss-120b", OPEN, "Mid", "2025-08-05", "exact", "Launch", 0.15, 0.60, "Together AI (reference host)", "Apache 2.0; no first-party price - hosted reference rate"),
    ("OpenAI", "USA", "gpt-oss-20b", OPEN, "Small", "2025-08-05", "exact", "Launch", 0.05, 0.20, "Together AI (reference host)", "Apache 2.0"),

    # --------------------------------------------------------------- Anthropic
    ("Anthropic", "USA", "Claude 1", FRONTIER, "Flagship", "2023-03-14", "exact", "Launch", 11.02, 32.68, "Anthropic API", "Token-equivalent of original per-character pricing"),
    ("Anthropic", "USA", "Claude Instant", FRONTIER, "Budget", "2023-03-14", "exact", "Launch", 1.63, 5.51, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude 2", FRONTIER, "Flagship", "2023-07-11", "exact", "Launch", 11.02, 32.68, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude 2.1", FRONTIER, "Flagship", "2023-11-21", "exact", "Launch", 8.00, 24.00, "Anthropic API", "Effective price cut vs Claude 2"),
    ("Anthropic", "USA", "Claude Instant 1.2", FRONTIER, "Budget", "2023-11-21", "exact", "Price cut", 0.80, 2.40, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude 3 Opus", FRONTIER, "Flagship", "2024-03-04", "exact", "Launch", 15.00, 75.00, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude 3 Sonnet", FRONTIER, "Mid", "2024-03-04", "exact", "Launch", 3.00, 15.00, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude 3 Haiku", FRONTIER, "Budget", "2024-03-13", "exact", "Launch", 0.25, 1.25, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude 3.5 Sonnet", FRONTIER, "Mid", "2024-06-20", "exact", "Launch", 3.00, 15.00, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude 3.5 Haiku", FRONTIER, "Budget", "2024-11-04", "exact", "Launch", 1.00, 5.00, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude 3.5 Haiku", FRONTIER, "Budget", "2024-12-01", "approx", "Price cut", 0.80, 4.00, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude 3.7 Sonnet", FRONTIER, "Mid", "2025-02-24", "exact", "Launch", 3.00, 15.00, "Anthropic API", "First hybrid reasoning Claude"),
    ("Anthropic", "USA", "Claude Opus 4", FRONTIER, "Flagship", "2025-05-22", "exact", "Launch", 15.00, 75.00, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude Sonnet 4", FRONTIER, "Mid", "2025-05-22", "exact", "Launch", 3.00, 15.00, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude Opus 4.1", FRONTIER, "Flagship", "2025-08-05", "exact", "Launch", 15.00, 75.00, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude Sonnet 4.5", FRONTIER, "Mid", "2025-09-29", "exact", "Launch", 3.00, 15.00, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude Haiku 4.5", FRONTIER, "Budget", "2025-10-15", "exact", "Launch", 1.00, 5.00, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude Opus 4.5", FRONTIER, "Flagship", "2025-11-24", "exact", "Launch", 5.00, 25.00, "Anthropic API", "67% price cut vs Opus 4.1 ($15/$75)"),
    ("Anthropic", "USA", "Claude Opus 4.6", FRONTIER, "Flagship", "2026-02-05", "exact", "Launch", 5.00, 25.00, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude Sonnet 4.6", FRONTIER, "Mid", "2026-02-05", "approx", "Launch", 3.00, 15.00, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude Opus 4.7", FRONTIER, "Flagship", "2026-04-16", "exact", "Launch", 5.00, 25.00, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude Opus 4.8", FRONTIER, "Flagship", "2026-06-01", "approx", "Launch", 5.00, 25.00, "Anthropic API", ""),
    ("Anthropic", "USA", "Claude Fable 5", FRONTIER, "Premium", "2026-06-09", "exact", "Launch", 10.00, 50.00, "Anthropic API", "Premium agentic tier above Opus; Mythos 5 (limited availability) same price"),
    ("Anthropic", "USA", "Claude Sonnet 5", FRONTIER, "Mid", "2026-06-30", "exact", "Launch", 2.00, 10.00, "Anthropic API", "Intro price made standard on 2026-08-10; scheduled rise to $3/$15 cancelled"),
    ("Anthropic", "USA", "Claude Opus 5", FRONTIER, "Flagship", "2026-07-24", "exact", "Launch", 5.00, 25.00, "Anthropic API", "1M context default"),

    # ------------------------------------------------------------------ Google
    ("Google", "USA", "Gemini 1.0 Pro", FRONTIER, "Mid", "2024-02-15", "exact", "Launch", 0.50, 1.50, "Gemini API", "Paid tier GA"),
    ("Google", "USA", "Gemini 1.5 Pro", FRONTIER, "Flagship", "2024-05-24", "exact", "Launch", 3.50, 10.50, "Gemini API", "<=128K tier"),
    ("Google", "USA", "Gemini 1.5 Pro", FRONTIER, "Flagship", "2024-10-01", "exact", "Price cut", 1.25, 5.00, "Gemini API", "64% cut, <=128K tier"),
    ("Google", "USA", "Gemini 1.5 Flash", FRONTIER, "Budget", "2024-05-24", "exact", "Launch", 0.35, 1.05, "Gemini API", "<=128K tier"),
    ("Google", "USA", "Gemini 1.5 Flash", FRONTIER, "Budget", "2024-08-08", "exact", "Price cut", 0.075, 0.30, "Gemini API", "~78% cut"),
    ("Google", "USA", "Gemini 2.0 Flash", FRONTIER, "Budget", "2025-02-05", "exact", "Launch", 0.10, 0.40, "Gemini API", ""),
    ("Google", "USA", "Gemini 2.0 Flash-Lite", FRONTIER, "Tiny", "2025-02-25", "exact", "Launch", 0.075, 0.30, "Gemini API", ""),
    ("Google", "USA", "Gemini 2.5 Pro", FRONTIER, "Flagship", "2025-06-17", "exact", "Launch", 1.25, 10.00, "Gemini API", "GA; <=200K tier ($2.50/$15 above)"),
    ("Google", "USA", "Gemini 2.5 Flash", FRONTIER, "Budget", "2025-06-17", "exact", "Launch", 0.30, 2.50, "Gemini API", ""),
    ("Google", "USA", "Gemini 2.5 Flash-Lite", FRONTIER, "Tiny", "2025-07-22", "exact", "Launch", 0.10, 0.40, "Gemini API", ""),
    ("Google", "USA", "Gemini 3 Pro", FRONTIER, "Flagship", "2025-11-18", "exact", "Launch", 2.00, 12.00, "Gemini API", "<=200K tier ($4/$18 above)"),
    ("Google", "USA", "Gemini 3 Flash", FRONTIER, "Budget", "2025-12-09", "approx", "Launch", 0.50, 3.00, "Gemini API", ""),
    ("Google", "USA", "Gemini 3.1 Pro", FRONTIER, "Flagship", "2026-03-01", "approx", "Launch", 2.00, 12.00, "Gemini API", "<=200K tier ($4/$18 above); 2M context"),
    ("Google", "USA", "Gemini 3.1 Flash-Lite", FRONTIER, "Tiny", "2026-04-01", "approx", "Launch", 0.25, 1.50, "Gemini API", ""),
    ("Google", "USA", "Gemini 3.5 Flash", FRONTIER, "Budget", "2026-05-19", "exact", "Launch", 1.50, 9.00, "Gemini API", "Price INCREASE vs 3 Flash; positioned above 3.1 Pro on coding"),
    ("Google", "USA", "Gemini 3.5 Flash-Lite", FRONTIER, "Tiny", "2026-06-01", "approx", "Launch", 0.30, 2.50, "Gemini API", ""),
    ("Google", "USA", "Gemini 3.6 Flash", FRONTIER, "Budget", "2026-07-15", "approx", "Launch", 1.50, 7.50, "Gemini API", ""),

    # --------------------------------------------------------------------- xAI
    ("xAI", "USA", "Grok Beta", FRONTIER, "Flagship", "2024-11-04", "approx", "Launch", 5.00, 15.00, "xAI API", ""),
    ("xAI", "USA", "Grok 2", FRONTIER, "Flagship", "2024-12-12", "exact", "Launch", 2.00, 10.00, "xAI API", ""),
    ("xAI", "USA", "Grok 3", FRONTIER, "Flagship", "2025-04-09", "exact", "Launch", 3.00, 15.00, "xAI API", "API launch"),
    ("xAI", "USA", "Grok 3 mini", FRONTIER, "Budget", "2025-04-09", "exact", "Launch", 0.30, 0.50, "xAI API", ""),
    ("xAI", "USA", "Grok 4", FRONTIER, "Flagship", "2025-07-09", "exact", "Launch", 3.00, 15.00, "xAI API", ""),
    ("xAI", "USA", "Grok 4 Fast", FRONTIER, "Budget", "2025-09-19", "exact", "Launch", 0.20, 0.50, "xAI API", "<128K tier"),
    ("xAI", "USA", "Grok 4.1 Fast", FRONTIER, "Budget", "2025-11-19", "approx", "Launch", 0.20, 0.50, "xAI API", ""),
    ("xAI", "USA", "Grok 4.20", FRONTIER, "Flagship", "2026-03-09", "approx", "Launch", 1.25, 2.50, "xAI API", "<200K tier (2x above); snapshot id 0309"),
    ("xAI", "USA", "Grok 4.3", FRONTIER, "Flagship", "2026-04-15", "approx", "Launch", 1.25, 2.50, "xAI API", "<200K tier"),
    ("xAI", "USA", "Grok 4.5", FRONTIER, "Flagship", "2026-05-15", "approx", "Launch", 2.00, 6.00, "xAI API", "<200K tier (2x above)"),
    ("xAI", "USA", "Grok 4.6", FRONTIER, "Flagship", "2026-08-10", "approx", "Launch", 2.00, 6.00, "xAI API", "<200K tier (2x above); 500K context"),
    ("xAI", "USA", "Grok Build 0.1", FRONTIER, "Budget", "2026-07-01", "approx", "Launch", 1.00, 2.00, "xAI API", "Coding/app-building model"),

    # -------------------------------------------------------------------- Meta
    ("Meta", "USA", "Llama 2 70B", OPEN, "Flagship", "2023-07-18", "exact", "Launch", 0.90, 0.90, "Together AI (reference host)", "No first-party API - hosted reference rate"),
    ("Meta", "USA", "Llama 3 70B", OPEN, "Flagship", "2024-04-18", "exact", "Launch", 0.90, 0.90, "Together AI (reference host)", ""),
    ("Meta", "USA", "Llama 3 8B", OPEN, "Small", "2024-04-18", "exact", "Launch", 0.20, 0.20, "Together AI (reference host)", ""),
    ("Meta", "USA", "Llama 3.1 405B", OPEN, "Flagship", "2024-07-23", "exact", "Launch", 3.50, 3.50, "Together AI (reference host)", "Largest open-weight launch of 2024"),
    ("Meta", "USA", "Llama 3.1 70B", OPEN, "Mid", "2024-07-23", "exact", "Launch", 0.88, 0.88, "Together AI (reference host)", ""),
    ("Meta", "USA", "Llama 3.1 8B", OPEN, "Small", "2024-07-23", "exact", "Launch", 0.18, 0.18, "Together AI (reference host)", ""),
    ("Meta", "USA", "Llama 3.3 70B", OPEN, "Mid", "2024-12-06", "exact", "Launch", 0.88, 0.88, "Together AI (reference host)", ""),
    ("Meta", "USA", "Llama 4 Scout", OPEN, "Small", "2025-04-05", "exact", "Launch", 0.18, 0.59, "Together AI (reference host)", "10M context MoE"),
    ("Meta", "USA", "Llama 4 Scout", OPEN, "Small", "2026-08-01", "approx", "Host price decline", 0.08, 0.30, "Cheapest tracked host", "Together/DeepInfra tier, Aug 2026"),
    ("Meta", "USA", "Llama 4 Maverick", OPEN, "Flagship", "2025-04-05", "exact", "Launch", 0.27, 0.85, "Together AI (reference host)", ""),
    ("Meta", "USA", "Llama 4 Maverick", OPEN, "Flagship", "2026-08-01", "approx", "Host price decline", 0.15, 0.60, "Cheapest tracked host (DeepInfra)", ""),
    ("Meta", "USA", "Muse Spark 1.2", FRONTIER, "Mid", "2026-08-01", "approx", "Launch", 1.25, 4.25, "Meta Model API (public preview)", "Meta's proprietary frontier line; Muse Spark 1.0 launched 2026-04-08"),

    # ---------------------------------------------------------------- DeepSeek
    ("DeepSeek", "China", "DeepSeek V2", OPEN, "Flagship", "2024-05-06", "exact", "Launch", 0.14, 0.28, "DeepSeek API", "Triggered China's 2024 LLM price war"),
    ("DeepSeek", "China", "DeepSeek V3", OPEN, "Flagship", "2024-12-26", "exact", "Launch (promo)", 0.14, 0.28, "DeepSeek API", "Promotional launch pricing"),
    ("DeepSeek", "China", "DeepSeek V3", OPEN, "Flagship", "2025-02-08", "exact", "Promo ended", 0.27, 1.10, "DeepSeek API", "Standard rate took effect"),
    ("DeepSeek", "China", "DeepSeek R1", OPEN, "Reasoning", "2025-01-20", "exact", "Launch", 0.55, 2.19, "DeepSeek API", "Open-weight reasoning model"),
    ("DeepSeek", "China", "DeepSeek V3.1", OPEN, "Flagship", "2025-08-21", "exact", "Launch", 0.56, 1.68, "DeepSeek API", "Unified chat+reasoning pricing"),
    ("DeepSeek", "China", "DeepSeek V3.2", OPEN, "Flagship", "2025-09-29", "exact", "Launch", 0.28, 0.42, "DeepSeek API", ">50% price cut vs V3.1 (V3.2-Exp)"),
    ("DeepSeek", "China", "DeepSeek V4 Flash", OPEN, "Mid", "2026-05-01", "approx", "Launch", 0.44, 1.33, "DeepSeek API", "CNY 3/9 per 1M; off-peak 50% off since 2026-08-16 ($0.22/$0.66)"),
    ("DeepSeek", "China", "DeepSeek V4 Pro", OPEN, "Flagship", "2026-08-13", "exact", "Launch", 1.33, 4.00, "DeepSeek API", "CNY 9/27 per 1M; off-peak $0.66/$1.98; snapshot deepseek-v4-pro-0813"),

    # ----------------------------------------------------------- Alibaba (Qwen)
    ("Alibaba", "China", "Qwen2.5-Max", FRONTIER, "Flagship", "2025-01-28", "exact", "Launch", 1.60, 6.40, "Alibaba Model Studio (intl)", "Proprietary Max line"),
    ("Alibaba", "China", "Qwen3 235B-A22B", OPEN, "Flagship", "2025-04-29", "exact", "Launch", 0.70, 2.80, "Alibaba Model Studio (intl)", "Non-thinking rate"),
    ("Alibaba", "China", "Qwen3-Max", FRONTIER, "Flagship", "2025-09-24", "exact", "Launch", 1.20, 6.00, "Alibaba Model Studio (intl)", "<=32K tier; tiered above"),
    ("Alibaba", "China", "Qwen3.5 Flash", OPEN, "Small", "2026-02-15", "approx", "Launch", 0.10, 0.40, "Alibaba Model Studio (intl)", "CN endpoint CNY 0.2/2 (<=128K)"),
    ("Alibaba", "China", "Qwen3.5 397B", OPEN, "Flagship", "2026-02-15", "approx", "Launch", 0.60, 3.60, "Alibaba Model Studio (intl)", "Open-weight flagship"),
    ("Alibaba", "China", "Qwen3.6 Plus", FRONTIER, "Mid", "2026-04-01", "approx", "Launch", 0.325, 1.95, "Alibaba Model Studio (intl)", "Singapore endpoint; CN cheaper"),
    ("Alibaba", "China", "Qwen3.7 Max", FRONTIER, "Flagship", "2026-05-20", "exact", "Launch", 1.78, 5.34, "Alibaba Model Studio (CN, converted)", "CNY 12/36 per 1M at ~6.75 CNY/USD"),
    ("Alibaba", "China", "Qwen3.8 Max", FRONTIER, "Flagship", "2026-08-15", "approx", "Launch", 1.78, 5.34, "Alibaba Model Studio (CN, converted)", "CNY 12/36; Prime speed tier CNY 24/72"),

    # ---------------------------------------------------------- Moonshot (Kimi)
    ("Moonshot AI", "China", "Kimi K2", OPEN, "Flagship", "2025-07-11", "exact", "Launch", 0.60, 2.50, "Moonshot API", "1T-param MoE, open-weight"),
    ("Moonshot AI", "China", "Kimi K2 Thinking", OPEN, "Flagship", "2025-11-06", "exact", "Launch", 0.60, 2.50, "Moonshot API", ""),
    ("Moonshot AI", "China", "Kimi K2.5", OPEN, "Flagship", "2026-01-15", "approx", "Launch", 0.59, 3.11, "Moonshot API (CN, converted)", "CNY 4/21 per 1M"),
    ("Moonshot AI", "China", "Kimi K2.6", OPEN, "Flagship", "2026-03-15", "approx", "Launch", 0.96, 4.00, "Moonshot API (CN, converted)", "CNY 6.5/27 per 1M"),
    ("Moonshot AI", "China", "Kimi K3", OPEN, "Flagship", "2026-07-16", "exact", "Launch", 3.00, 15.00, "Moonshot API", "2.8T-param flagship, 1M context; CNY 20/100; most expensive open-weight launch to date"),

    # ------------------------------------------------------------- Zhipu (GLM)
    ("Zhipu AI", "China", "GLM-4.5", OPEN, "Flagship", "2025-07-28", "exact", "Launch", 0.60, 2.20, "Z.ai API (intl)", ""),
    ("Zhipu AI", "China", "GLM-4.6", OPEN, "Flagship", "2025-09-30", "exact", "Launch", 0.60, 2.20, "Z.ai API (intl)", ""),
    ("Zhipu AI", "China", "GLM-4.7", OPEN, "Flagship", "2025-12-10", "approx", "Launch", 0.60, 2.20, "Z.ai API (intl)", "CN endpoint CNY 2/8 (~$0.30/$1.19)"),
    ("Zhipu AI", "China", "GLM-5", OPEN, "Flagship", "2026-02-15", "approx", "Launch", 1.00, 3.20, "Z.ai API (intl)", ""),
    ("Zhipu AI", "China", "GLM-5.1", OPEN, "Flagship", "2026-04-15", "approx", "Launch", 1.40, 4.40, "Z.ai API (intl)", "CN CNY 6/24"),
    ("Zhipu AI", "China", "GLM-5.2", OPEN, "Flagship", "2026-06-15", "approx", "Launch", 1.40, 4.40, "Z.ai API (intl)", "MIT-licensed open weights; CN CNY 8/28"),
    ("Zhipu AI", "China", "GLM-5.3", OPEN, "Flagship", "2026-08-15", "approx", "Launch", 1.19, 4.15, "Zhipu CN (converted)", "CNY 8/28 per 1M at ~6.75 CNY/USD; intl rate not yet listed"),

    # ----------------------------------------------------------------- MiniMax
    ("MiniMax", "China", "MiniMax M1", OPEN, "Flagship", "2025-06-17", "exact", "Launch", 0.40, 2.20, "MiniMax API", "<=200K tier"),
    ("MiniMax", "China", "MiniMax M2", OPEN, "Flagship", "2025-10-27", "exact", "Launch", 0.30, 1.20, "MiniMax API", ""),
    ("MiniMax", "China", "MiniMax M2.5", OPEN, "Flagship", "2026-02-01", "approx", "Launch", 0.30, 1.20, "MiniMax API", ""),
    ("MiniMax", "China", "MiniMax M3", OPEN, "Flagship", "2026-06-15", "approx", "Launch", 0.30, 1.20, "MiniMax API", "<=512K tier (2x above)"),

    # ----------------------------------------------------------------- Mistral
    ("Mistral AI", "France", "Mixtral 8x7B", OPEN, "Mid", "2023-12-11", "exact", "Launch", 0.70, 0.70, "La Plateforme", "Apache 2.0 MoE"),
    ("Mistral AI", "France", "Mistral Large", FRONTIER, "Flagship", "2024-02-26", "exact", "Launch", 8.00, 24.00, "La Plateforme", "Proprietary"),
    ("Mistral AI", "France", "Mistral Large 2", OPEN, "Flagship", "2024-07-24", "exact", "Launch", 3.00, 9.00, "La Plateforme", "Open weights (research license)"),
    ("Mistral AI", "France", "Mistral Large 2", OPEN, "Flagship", "2024-11-01", "approx", "Price cut", 2.00, 6.00, "La Plateforme", ""),
    ("Mistral AI", "France", "Mistral Small 3", OPEN, "Small", "2025-01-30", "exact", "Launch", 0.10, 0.30, "La Plateforme", "Apache 2.0"),
    ("Mistral AI", "France", "Mistral Medium 3", FRONTIER, "Mid", "2025-05-07", "exact", "Launch", 0.40, 2.00, "La Plateforme", "Proprietary"),
    ("Mistral AI", "France", "Magistral Medium", FRONTIER, "Reasoning", "2025-06-10", "exact", "Launch", 2.00, 5.00, "La Plateforme", "Reasoning model"),
    ("Mistral AI", "France", "Mistral Large 3", OPEN, "Flagship", "2025-12-02", "approx", "Launch", 0.50, 1.50, "La Plateforme", "Open weights (Apache 2.0)"),
    ("Mistral AI", "France", "Mistral Small 4", OPEN, "Small", "2026-03-01", "approx", "Launch", 0.15, 0.60, "La Plateforme", ""),
    ("Mistral AI", "France", "Mistral Medium 3.5", FRONTIER, "Mid", "2026-06-01", "approx", "Launch", 1.50, 7.50, "La Plateforme", ""),

    # ------------------------------------------------------------------ Amazon
    ("Amazon", "USA", "Nova Micro", FRONTIER, "Tiny", "2024-12-03", "exact", "Launch", 0.035, 0.14, "AWS Bedrock", ""),
    ("Amazon", "USA", "Nova Lite", FRONTIER, "Budget", "2024-12-03", "exact", "Launch", 0.06, 0.24, "AWS Bedrock", ""),
    ("Amazon", "USA", "Nova Pro", FRONTIER, "Mid", "2024-12-03", "exact", "Launch", 0.80, 3.20, "AWS Bedrock", ""),
    ("Amazon", "USA", "Nova Premier", FRONTIER, "Flagship", "2025-04-30", "exact", "Launch", 2.50, 12.50, "AWS Bedrock", ""),

    # ------------------------------------------------------------------ Cohere
    ("Cohere", "Canada", "Command R+", FRONTIER, "Flagship", "2024-04-04", "exact", "Launch", 3.00, 15.00, "Cohere API", "Cut to $2.50/$10 with 08-2024 refresh"),
    ("Cohere", "Canada", "Command A", FRONTIER, "Flagship", "2025-03-13", "exact", "Launch", 2.50, 10.00, "Cohere API", ""),
]

FIELDS = [
    "developer", "country", "model", "model_class", "tier",
    "date", "date_precision", "event",
    "input_usd_per_m", "output_usd_per_m", "basis", "notes",
]


def blended(input_price: float, output_price: float) -> float:
    """Blended $/1M tokens at the standard 3:1 input:output usage weighting."""
    return (3.0 * input_price + output_price) / 4.0
