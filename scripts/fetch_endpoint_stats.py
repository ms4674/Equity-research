#!/usr/bin/env python3
"""Snapshot OpenRouter per-provider endpoint stats for the top model-variant
pairs by weekly token volume.

For each (model_permaslug, variant) pair covering the bulk of weekly volume,
fetches https://openrouter.ai/api/frontend/v1/stats/endpoint?permaslug=...
which returns each serving provider's request count over a live 30-minute
window. Also snapshots the provider directory (headquarters, data policy).

Writes:
  - data/raw/openrouter_endpoint_stats_<date>.json
  - data/raw/openrouter_providers_<date>.json
"""

import glob
import json
import os
import time
import urllib.parse
import urllib.request
from collections import defaultdict
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(ROOT, "data", "raw")
TOP_N_PAIRS = 60
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64)"}


def get(url):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.load(resp)


def main():
    today = date.today().isoformat()
    rankings_path = sorted(glob.glob(os.path.join(RAW, "openrouter_rankings_week_*.json")))[-1]
    with open(rankings_path) as f:
        rankings = json.load(f)["data"]

    agg = defaultdict(int)
    for row in rankings:
        agg[(row["model_permaslug"], row["variant"])] += (
            row["total_prompt_tokens"] + row["total_completion_tokens"]
        )
    pairs = sorted(agg.items(), key=lambda kv: -kv[1])[:TOP_N_PAIRS]

    out = []
    for (permaslug, variant), tokens in pairs:
        base = ("https://openrouter.ai/api/frontend/v1/stats/endpoint?permaslug="
                + urllib.parse.quote(permaslug, safe="/"))
        fallback = False
        try:
            eps = get(base + "&variant=" + urllib.parse.quote(variant, safe="/")).get("data", [])
        except Exception:
            eps = []
        if not eps:
            # Some variants (e.g. promo free tiers) expose no live endpoint
            # stats; use the model's other endpoints as a serving proxy.
            try:
                eps = get(base).get("data", [])
                fallback = True
            except Exception as exc:
                print(f"  ERROR {permaslug}: {exc}")
                eps = []
        endpoints = [
            {
                "provider_name": ep.get("provider_name"),
                "variant": ep.get("variant"),
                "quantization": ep.get("quantization"),
                "request_count_30m": (ep.get("stats") or {}).get("request_count") or 0,
            }
            for ep in eps
        ]
        out.append(
            {
                "model_permaslug": permaslug,
                "variant": variant,
                "weekly_tokens": tokens,
                "variant_fallback": fallback,
                "endpoints": endpoints,
            }
        )
        n = len(endpoints)
        print(f"  {permaslug} [{variant}]{' (fallback)' if fallback else ''}: "
              f"{n} endpoints, {tokens / 1e9:.0f}B tokens/wk")
        time.sleep(0.3)

    stats_path = os.path.join(RAW, f"openrouter_endpoint_stats_{today}.json")
    with open(stats_path, "w") as f:
        json.dump({"fetched_at": today, "source_week": rankings_path.split("_")[-1][:-5],
                   "data": out}, f, indent=1)
    print(f"Wrote {stats_path}")

    providers = get("https://openrouter.ai/api/frontend/v1/providers")
    prov_path = os.path.join(RAW, f"openrouter_providers_{today}.json")
    with open(prov_path, "w") as f:
        json.dump(providers, f, indent=1)
    print(f"Wrote {prov_path} ({len(providers)} providers)")


if __name__ == "__main__":
    main()
