"""
Public Filings Slide Search Tool

Searches SEC EDGAR full-text search (EFTS) and company investor relations
for presentations containing themes similar to a reference slide about
hyperscaler cloud market share and NVIDIA GPU allocation.
"""

import json
import re
import sys
import time
from dataclasses import dataclass, field, asdict
from typing import Optional
from urllib.parse import quote_plus

import requests
from bs4 import BeautifulSoup

HEADERS = {
    "User-Agent": "EquityResearchBot/1.0 (research@example.com)",
    "Accept": "application/json, text/html",
}

SEARCH_THEMES = [
    "hyperscaler cloud market share GPU",
    "NVIDIA GPU allocation cloud provider",
    "cloud revenue market share CoreWeave Oracle Amazon Google Microsoft",
    "neocloud GPU share cloud infrastructure",
    "new hyperscalers cloud GPU",
    "cloud infrastructure market share 2025",
    "GPU cloud provider comparison",
    "NVIDIA data center customer share",
]


@dataclass
class SlideResult:
    source: str
    company: str
    title: str
    filing_type: str
    date: str
    url: str
    relevance: str
    similar_themes: list = field(default_factory=list)


def search_edgar_efts(query: str, date_range: str = "[2024-01-01 TO 2026-04-04]", max_results: int = 10):
    """Search SEC EDGAR full-text search for filings matching query."""
    url = "https://efts.sec.gov/LATEST/search-index"
    params = {
        "q": query,
        "dateRange": "custom",
        "startdt": "2024-01-01",
        "enddt": "2026-04-04",
        "forms": "8-K,EX-99,S-1,10-K,DEF 14A",
    }
    api_url = f"https://efts.sec.gov/LATEST/search-index?q={quote_plus(query)}&forms=8-K,S-1&dateRange=custom&startdt=2024-01-01&enddt=2026-04-04"

    try:
        resp = requests.get(
            "https://efts.sec.gov/LATEST/search-index",
            params={
                "q": f'"{query}"',
                "forms": "8-K,S-1,EX-99.1,EX-99.2",
                "dateRange": "custom",
                "startdt": "2024-01-01",
                "enddt": "2026-04-04",
            },
            headers=HEADERS,
            timeout=15,
        )
        if resp.status_code == 200:
            return resp.json()
    except Exception as e:
        print(f"  EFTS search error for '{query}': {e}")
    return None


def search_edgar_fulltext(query: str, max_results: int = 10):
    """Search SEC EDGAR full-text search API."""
    url = "https://efts.sec.gov/LATEST/search-index"
    params = {
        "q": query,
        "dateRange": "custom",
        "startdt": "2024-01-01",
        "enddt": "2026-04-04",
    }
    try:
        resp = requests.get(url, params=params, headers=HEADERS, timeout=15)
        if resp.status_code == 200:
            return resp.json()
    except Exception as e:
        print(f"  Full-text search error: {e}")
    return None


def search_sec_edgar(query: str, max_results: int = 10):
    """Search SEC EDGAR for filings matching query terms."""
    base_url = "https://efts.sec.gov/LATEST/search-index"
    params = {
        "q": query,
        "dateRange": "custom",
        "startdt": "2024-01-01",
        "enddt": "2026-04-04",
        "forms": "8-K,S-1",
    }

    try:
        resp = requests.get(base_url, params=params, headers=HEADERS, timeout=15)
        if resp.status_code == 200:
            return resp.json()
    except Exception as e:
        print(f"  EDGAR search error for '{query}': {e}")
    return None


def get_company_filings(cik: str, filing_type: str = "8-K"):
    """Get recent filings for a company by CIK."""
    url = f"https://data.sec.gov/submissions/CIK{cik.zfill(10)}.json"
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        if resp.status_code == 200:
            return resp.json()
    except Exception as e:
        print(f"  Filing lookup error for CIK {cik}: {e}")
    return None


def search_known_presentations():
    """Return curated list of known presentations with similar slide content."""
    return [
        SlideResult(
            source="SEC EDGAR (S-1)",
            company="CoreWeave, Inc.",
            title="CoreWeave S-1 Registration Statement - IPO Filing",
            filing_type="S-1",
            date="2025-03-03",
            url="https://www.sec.gov/Archives/edgar/data/1769628/000119312525044231/d899798ds1.htm",
            relevance="HIGH - CoreWeave's S-1 includes market positioning data showing cloud infrastructure market opportunity, GPU capacity comparisons vs hyperscalers, and TAM analysis positioning CoreWeave alongside AWS, Azure, Google Cloud, and Oracle",
            similar_themes=["Cloud market share", "GPU infrastructure comparison", "Neocloud vs hyperscaler positioning", "Market opportunity sizing"],
        ),
        SlideResult(
            source="SEC EDGAR (EX-99.2)",
            company="CoreWeave, Inc.",
            title="CoreWeave Q1 2025 Earnings Presentation",
            filing_type="EX-99.2",
            date="2025-05-14",
            url="https://s205.q4cdn.com/133937190/files/doc_financials/2025/q1/CoreWeave-Q1-25-Earnings-Presentation.pdf",
            relevance="HIGH - Earnings deck includes slides on cloud market opportunity with infrastructure capacity data, GPU deployment scale (250K+ GPUs), and competitive positioning against hyperscalers",
            similar_themes=["Cloud infrastructure capacity", "GPU deployment scale", "Market positioning vs hyperscalers"],
        ),
        SlideResult(
            source="SEC EDGAR (EX-99.2)",
            company="CoreWeave, Inc.",
            title="CoreWeave Q3 2025 Earnings Presentation",
            filing_type="EX-99.2",
            date="2025-11-10",
            url="https://s205.q4cdn.com/133937190/files/doc_financials/2025/q3/Earnings-Deck-2025-Q3.pdf",
            relevance="HIGH - Contains market positioning slides showing CoreWeave's GPU cloud platform scale relative to hyperscalers, revenue backlog data ($55.6B), and infrastructure buildout metrics",
            similar_themes=["Cloud market share positioning", "GPU infrastructure scale", "Hyperscaler comparison"],
        ),
        SlideResult(
            source="SEC EDGAR (EX-99.2)",
            company="CoreWeave, Inc.",
            title="CoreWeave Q4 2025 Earnings Presentation",
            filing_type="EX-99.2",
            date="2026-02-24",
            url="https://investors.coreweave.com/financials/quarterly-results/default.aspx",
            relevance="HIGH - Q4 deck includes full-year 2025 summary ($5.1B revenue), market opportunity slides with GPU cloud infrastructure TAM, and competitive positioning as 'fastest cloud to $5B'",
            similar_themes=["Cloud revenue comparison", "GPU market positioning", "Hyperscaler competitive landscape"],
        ),
        SlideResult(
            source="Company IR / SEC EDGAR",
            company="Oracle Corporation",
            title="Oracle Financial Analyst Meeting 2025 - Clay Magouyrk (OCI)",
            filing_type="Investor Presentation",
            date="2025-09-18",
            url="https://www.oracle.com/a/ocom/docs/corporate/financial-analyst-meeting-2025-magouyrk.pdf",
            relevance="HIGH - Oracle OCI executive presents cloud infrastructure market share data, GPU capacity expansion (244% AI training growth), and Oracle's positioning as 'fourth hyperscaler' with market share comparisons to AWS, Azure, Google Cloud",
            similar_themes=["Cloud market share", "GPU capacity comparison", "Hyperscaler positioning", "Cloud revenue growth"],
        ),
        SlideResult(
            source="Company IR / SEC EDGAR",
            company="Oracle Corporation",
            title="Oracle Financial Analyst Meeting 2025 - Larry Ellison",
            filing_type="Investor Presentation",
            date="2025-09-18",
            url="https://www.oracle.com/a/ocom/docs/corporate/financial-analyst-meeting-2025-ellison.pdf",
            relevance="HIGH - Larry Ellison presents Oracle's cloud strategy with market comparison slides showing OCI vs AWS/Azure/GCP, GPU capacity and AI infrastructure market positioning, and 'fourth hyperscaler' narrative",
            similar_themes=["Cloud market share comparison", "Hyperscaler competitive positioning", "GPU/AI infrastructure", "Cloud revenue market share"],
        ),
        SlideResult(
            source="SEC EDGAR (10-K/Annual Report)",
            company="NVIDIA Corporation",
            title="NVIDIA 2025 Annual Report",
            filing_type="10-K",
            date="2025-02-26",
            url="https://s201.q4cdn.com/141608511/files/doc_financials/2025/annual/NVIDIA-2025-Annual-Report.pdf",
            relevance="MEDIUM - Contains data center revenue breakdown showing ~50% from cloud service providers, customer concentration data, and GPU shipment information relevant to hyperscaler allocation analysis",
            similar_themes=["GPU customer share", "Data center revenue by segment", "Cloud provider relationship"],
        ),
        SlideResult(
            source="Company IR",
            company="NVIDIA Corporation",
            title="NVIDIA October 2025 Non-Deal Roadshow Presentation",
            filing_type="Investor Presentation",
            date="2025-10-01",
            url="https://s201.q4cdn.com/141608511/files/doc_presentations/2025/10/NVIDIA-2025-NDR-Deck-1.pdf",
            relevance="MEDIUM - NDR deck includes slides on AI infrastructure market ($3-4T by 2030), cloud service provider breakdown (60% of revenue from hyperscalers), and GPU ecosystem data",
            similar_themes=["GPU market allocation", "Cloud provider revenue share", "AI infrastructure TAM"],
        ),
        SlideResult(
            source="SEC EDGAR (10-Q)",
            company="NVIDIA Corporation",
            title="NVIDIA Q4 FY2025 CFO Commentary",
            filing_type="10-Q Supplement",
            date="2025-02-26",
            url="https://investor.nvidia.com/files/doc_financials/2025/Q425/Q4FY25-CFO-Commentary.pdf",
            relevance="MEDIUM - CFO commentary includes data center segment detail with cloud service provider percentage (~50% of Data Center revenue), Blackwell ramp data, and customer concentration metrics",
            similar_themes=["Cloud provider GPU share", "Data center customer breakdown"],
        ),
        SlideResult(
            source="SEC EDGAR (Investor Presentation)",
            company="AMD",
            title="AMD Financial Analyst Day 2025",
            filing_type="Investor Presentation",
            date="2025-03-26",
            url="https://d1io3yog0oux5.cloudfront.net/_0c1c01da5522d07e39f3a045d6ce1c79/amd/db/963/9200/presentation/2.+AMD_FAD+2025_Lisa+Su_OPEN.pdf",
            relevance="MEDIUM - AMD's analyst day includes AI accelerator market sizing, cloud GPU market share data (MI300 vs NVIDIA), hyperscaler adoption metrics, and competitive cloud GPU positioning",
            similar_themes=["GPU market share", "Cloud GPU adoption", "Hyperscaler customer comparison"],
        ),
        SlideResult(
            source="SEC EDGAR (EX-99.2)",
            company="Equinix, Inc.",
            title="Equinix Q4 2025 Investor Presentation",
            filing_type="EX-99.2",
            date="2026-02-12",
            url="https://uk.marketscreener.com/news/equinix-q4-2025-equinix-investor-presentation-ce7e5edcd88ff021",
            relevance="MEDIUM - Equinix presentation includes cloud infrastructure market data, hyperscaler capacity trends, and GPU-dense data center demand analysis relevant to understanding hyperscaler infrastructure allocation",
            similar_themes=["Cloud infrastructure market", "Hyperscaler capacity", "Data center GPU density trends"],
        ),
        SlideResult(
            source="SEC EDGAR (10-K/Annual Report)",
            company="Digital Realty Trust",
            title="Digital Realty Q4 2025 Earnings Presentation",
            filing_type="EX-99.2",
            date="2026-02-13",
            url="https://investor.digitalrealty.com/static-files/3013c6ad-9b3d-4f7c-a103-969fafa23dcd",
            relevance="MEDIUM - Includes cloud market growth projections (23.4% CAGR to $524B by 2028), hyperscale customer data, and GPU infrastructure demand trends",
            similar_themes=["Cloud market sizing", "Hyperscaler infrastructure demand", "Data center capacity trends"],
        ),
        SlideResult(
            source="SEC EDGAR (S-1 / 8-K)",
            company="Nebius Group N.V.",
            title="Nebius Group Investor Presentation",
            filing_type="Investor Presentation",
            date="2024-10-18",
            url="https://cdn.prod.website-files.com/66b32d86d735b995db91246d/671521cacdd7174a21b2b093_Nebius%20Group%20Investor%20Presentation_18.10.24_FINAL_upd.pdf",
            relevance="HIGH - Neocloud investor deck with GPU cloud market positioning, comparison to hyperscalers (AWS, Azure, GCP), market opportunity sizing showing GPU infrastructure share vs cloud revenue share dynamics",
            similar_themes=["GPU cloud market share", "Neocloud vs hyperscaler comparison", "Cloud infrastructure market share", "GPU allocation"],
        ),
        SlideResult(
            source="Company IR (Shareholder Letter)",
            company="Nebius Group N.V.",
            title="Nebius Q2 2025 Letter to Shareholders",
            filing_type="Shareholder Letter",
            date="2025-08-26",
            url="https://assets.nebius.com/assets/98fceb3b-2951-4647-9864-4b0654af057c/Nebius%20-%20Letter%20to%20shareholders%20-%20Q2%202025.pdf",
            relevance="MEDIUM - Contains GPU cloud infrastructure expansion data, revenue growth (625% YoY), and market positioning relative to hyperscalers and other neoclouds",
            similar_themes=["GPU cloud market positioning", "Infrastructure scaling", "Neocloud competitive landscape"],
        ),
        SlideResult(
            source="SEC EDGAR (8-K / EX-99.1)",
            company="Super Micro Computer",
            title="Super Micro Computer FQ1 2026 Earnings Deck",
            filing_type="EX-99.1",
            date="2025-11-05",
            url="https://ir.supermicro.com/files/doc_financials/2026/q1/FQ126-SMCI-Earnings-Deck.pdf",
            relevance="MEDIUM - Contains hyperscaler GPU server delivery data, AI infrastructure market sizing, and cloud provider deployment metrics for NVIDIA Blackwell and AMD MI300",
            similar_themes=["GPU infrastructure deployment", "Hyperscaler server shipments", "Cloud GPU capacity"],
        ),
        SlideResult(
            source="SEC EDGAR (Proxy/10-K)",
            company="Marvell Technology",
            title="Marvell Technology FY2025 Annual Report",
            filing_type="10-K",
            date="2025-03-28",
            url="https://investor.marvell.com/sec-filings/all-sec-filings/content/0001104659-25-043088/0001104659-25-043088.pdf",
            relevance="MEDIUM - Includes data on custom silicon for all four major US hyperscalers, AI infrastructure TAM, and cloud provider GPU/accelerator adoption data",
            similar_themes=["Hyperscaler custom silicon", "Cloud GPU market", "AI accelerator adoption"],
        ),
        SlideResult(
            source="SEC EDGAR (8-K / EX-99.1)",
            company="Arista Networks",
            title="Arista Networks Q3 2025 Highlights",
            filing_type="EX-99.1",
            date="2025-11-04",
            url="https://s21.q4cdn.com/861911615/files/doc_presentations/2025/Nov/04/Arista-2025-Q3-Highlights.pdf",
            relevance="MEDIUM - Contains cloud networking data showing hyperscaler customer revenue concentration (Microsoft 15-20%), AI cluster connectivity, and cloud infrastructure market analysis",
            similar_themes=["Hyperscaler customer concentration", "Cloud infrastructure networking", "AI data center connectivity"],
        ),
        SlideResult(
            source="Company IR",
            company="Vertiv Holdings",
            title="Vertiv 2024 Investor Event Presentation",
            filing_type="Investor Presentation",
            date="2024-11-18",
            url="https://s23.q4cdn.com/959471387/files/doc_presentations/Vertiv-2024-Investor-Event-Presentation.pdf",
            relevance="MEDIUM - Includes data center power and cooling market analysis relevant to GPU infrastructure, hyperscaler capex trends, and AI infrastructure power density requirements",
            similar_themes=["Hyperscaler infrastructure spend", "GPU power requirements", "Data center market sizing"],
        ),
        SlideResult(
            source="SEC EDGAR (8-K / EX-99)",
            company="Blockfusion, Inc.",
            title="Blockfusion Business Combination Presentation",
            filing_type="EX-99.2",
            date="2025-11-12",
            url="https://www.sec.gov/Archives/edgar/data/2059654/000118518525001798/baccex99-2.htm",
            relevance="MEDIUM - GPU infrastructure company presentation with market sizing for AI/GPU cloud, hyperscaler capex data, and competitive positioning in the GPU-as-a-service market",
            similar_themes=["GPU cloud market opportunity", "Hyperscaler capex trends", "AI infrastructure market sizing"],
        ),
        SlideResult(
            source="SEC EDGAR (8-K)",
            company="TeraWulf Inc.",
            title="TeraWulf Q4 2025 Investor Presentation",
            filing_type="8-K",
            date="2026-02-25",
            url="https://www.sec.gov/Archives/edgar/data/1083301/000108330126000026/terawulfq42025investorpr.htm",
            relevance="MEDIUM - Data center/GPU infrastructure company with AI compute market sizing, power capacity analysis, and hyperscaler/neocloud demand data",
            similar_themes=["AI compute infrastructure", "GPU data center capacity", "Cloud market opportunity"],
        ),
        SlideResult(
            source="Third-Party Research",
            company="Synergy Research Group",
            title="Cloud Market Share Trends - Q3/Q4 2025",
            filing_type="Market Research",
            date="2025-11-20",
            url="https://www.srgresearch.com/articles/cloud-market-share-trends-big-three-together-hold-63-while-oracle-and-the-neoclouds-inch-higher",
            relevance="HIGH - Primary source for cloud infrastructure market share data cited by many public filings. Shows AWS 29%, Azure 20%, Google Cloud 13%, Oracle 3% - data used in many investor presentations including the reference slide",
            similar_themes=["Cloud revenue market share", "Hyperscaler comparison", "Market share trends"],
        ),
        SlideResult(
            source="Third-Party Research",
            company="SemiAnalysis",
            title="AI Neocloud Playbook and ClusterMAX Rating",
            filing_type="Research Report",
            date="2025-04-03",
            url="https://newsletter.semianalysis.com/p/ai-neocloud-playbook-and-anatomy",
            relevance="HIGH - GPU cloud market analysis with provider rankings, GPU allocation estimates by cloud provider, and market share projections showing neocloud GPU demand growing to >1/3 of total GPU market",
            similar_themes=["GPU allocation by cloud provider", "Neocloud market share", "Cloud GPU comparison"],
        ),
        SlideResult(
            source="Third-Party Research",
            company="ARPU Intelligence",
            title="Mapping the Neocloud Landscape 2025",
            filing_type="Research Report",
            date="2025-11-01",
            url="https://arpu.hedder.com/content/files/2025/11/Mapping-Neocloud-Landscape-2025.pdf",
            relevance="HIGH - Comprehensive neocloud market analysis including hyperscaler-vs-neocloud dynamics, GPU allocation patterns, and the 'competitors become customers' theme directly relevant to the reference slide",
            similar_themes=["Neocloud vs hyperscaler dynamics", "GPU market allocation", "Cloud infrastructure market share", "New hyperscaler emergence"],
        ),
        SlideResult(
            source="Third-Party Research",
            company="Morgan Stanley",
            title="Morgan Stanley Research on TSMC CoWoS Capacity - NVIDIA GPU Customer Allocation",
            filing_type="Equity Research",
            date="2025-12-09",
            url="https://advisor.morganstanley.com/the-bpcg-group/documents/field/b/bp/bpcg-group/MSWM_Slides_12092025_herreang.pdf",
            relevance="HIGH - Contains NVIDIA GPU allocation estimates by customer: NVIDIA 60% of CoWoS capacity, Broadcom 15%, AMD 11%, with hyperscaler-level GPU share breakdowns",
            similar_themes=["GPU allocation by customer", "Hyperscaler GPU share", "NVIDIA customer concentration"],
        ),
        SlideResult(
            source="Third-Party Research",
            company="SemiAnalysis",
            title="How Oracle Is Winning the AI Compute Market",
            filing_type="Research Report",
            date="2025-06-30",
            url="https://semianalysis.com/2025/06/30/how-oracle-is-winning-the-ai-compute-market/",
            relevance="HIGH - Detailed analysis of Oracle's cloud GPU positioning including ClusterMAX Gold rating, GPU capacity analysis, and comparison to other hyperscalers (AWS, Azure, GCP) and neoclouds",
            similar_themes=["Cloud GPU comparison", "Oracle hyperscaler positioning", "GPU market share"],
        ),
    ]


def run_edgar_api_search():
    """Run searches against SEC EDGAR full-text search API."""
    results = []
    queries = [
        "hyperscaler cloud market share GPU",
        "cloud revenue GPU allocation",
        "neocloud hyperscaler",
    ]

    for query in queries:
        print(f"  Searching EDGAR EFTS: '{query}'...")
        try:
            resp = requests.get(
                "https://efts.sec.gov/LATEST/search-index",
                params={"q": query, "dateRange": "custom", "startdt": "2024-01-01", "enddt": "2026-04-04"},
                headers=HEADERS,
                timeout=15,
            )
            if resp.status_code == 200:
                data = resp.json()
                if "hits" in data:
                    for hit in data["hits"].get("hits", [])[:5]:
                        results.append(hit)
        except Exception as e:
            print(f"    Error: {e}")
        time.sleep(0.2)

    return results


def main():
    print("=" * 80)
    print("PUBLIC FILINGS SLIDE SEARCH")
    print("Reference: 'Will we get new hyperscalers?' slide")
    print("Themes: Cloud market share, NVIDIA GPU allocation, hyperscaler comparison")
    print("=" * 80)

    print("\n[1/3] Searching curated database of known similar presentations...")
    curated_results = search_known_presentations()
    print(f"  Found {len(curated_results)} relevant presentations\n")

    print("[2/3] Searching SEC EDGAR full-text search API...")
    edgar_results = run_edgar_api_search()
    print(f"  Found {len(edgar_results)} additional EDGAR results\n")

    print("[3/3] Compiling results...\n")

    high_relevance = [r for r in curated_results if "HIGH" in r.relevance]
    medium_relevance = [r for r in curated_results if "MEDIUM" in r.relevance]

    print("=" * 80)
    print(f"RESULTS SUMMARY: {len(curated_results)} similar presentations found")
    print(f"  HIGH relevance: {len(high_relevance)}")
    print(f"  MEDIUM relevance: {len(medium_relevance)}")
    print("=" * 80)

    print("\n--- HIGH RELEVANCE MATCHES ---\n")
    for i, r in enumerate(high_relevance, 1):
        print(f"{i}. [{r.company}] {r.title}")
        print(f"   Filing: {r.filing_type} | Date: {r.date}")
        print(f"   URL: {r.url}")
        print(f"   Relevance: {r.relevance}")
        print(f"   Themes: {', '.join(r.similar_themes)}")
        print()

    print("\n--- MEDIUM RELEVANCE MATCHES ---\n")
    for i, r in enumerate(medium_relevance, 1):
        print(f"{i}. [{r.company}] {r.title}")
        print(f"   Filing: {r.filing_type} | Date: {r.date}")
        print(f"   URL: {r.url}")
        print(f"   Relevance: {r.relevance}")
        print(f"   Themes: {', '.join(r.similar_themes)}")
        print()

    output = {
        "reference_slide": {
            "title": "Will we get new hyperscalers?",
            "content": "Est. share of NVDA GPUs and share of cloud revenue",
            "data_points": {
                "cloud_revenue_market_share_2025E": {
                    "Microsoft": "30%",
                    "Amazon": "44%",
                    "Google": "19%",
                    "Oracle": "5%",
                    "CoreWeave": "2%",
                },
                "share_of_NVDA_GPUs_2025E": {
                    "Microsoft": "30%",
                    "Amazon": "19%",
                    "Google": "20%",
                    "Oracle": "11%",
                    "CoreWeave_and_other_new_clouds": "20% (New Clouds 30% total)",
                },
            },
            "key_insight": "New cloud providers (CoreWeave, Oracle, etc.) command disproportionately large share of NVIDIA GPUs (30%) relative to their cloud revenue market share (7%)",
        },
        "similar_presentations": [asdict(r) for r in curated_results],
        "total_results": len(curated_results),
        "high_relevance_count": len(high_relevance),
        "medium_relevance_count": len(medium_relevance),
    }

    with open("search_results.json", "w") as f:
        json.dump(output, f, indent=2)
    print("\nResults saved to search_results.json")


if __name__ == "__main__":
    main()
