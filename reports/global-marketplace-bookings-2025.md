# Global Marketplace Bookings: Delivery, Ridesharing & Online Travel (2025)

*Compiled September 2026 from Q4/FY2025 earnings releases, annual reports, and (where flagged) third-party estimates. All figures in USD billions of gross transaction value — GOV / GMV / GTV / gross bookings as each company defines it — not revenue.*

## Headline

**Aggregate marketplace bookings across the three sectors were roughly $1.54 trillion in the latest reported full year** (calendar 2025 for most companies; fiscal years ending Feb–Mar 2026 for several Indian, UK and travel names).

| Sector | Company-disclosed | Estimated | Subtotal |
|---|---:|---:|---:|
| Food & grocery delivery | $308B | $343B | **~$651B** |
| Ridesharing / ride-hailing | $216B | $12B | **~$228B** |
| Online travel | $650B | $11B | **~$661B** |
| **Grand total** | | | **~$1,539B** |

Stripping out the B2B travel bed banks (HBX/Hotelbeds, WebBeds), whose volumes are largely resold through the retail OTAs already counted, the consumer-facing total is **~$1,527B**.

Two structural notes on the estimate-heavy delivery column: China's platforms (Meituan, Ele.me/Taobao Instant Commerce, JD) do not disclose absolute transaction value, so ~$275B of the delivery subtotal rests on triangulated estimates of a market distorted by the 2025 subsidy war (~RMB 150B of subsidies burned across the three players). And 2025 was a year of heavy consolidation — DoorDash absorbed Deliveroo (Oct 2025), Prosus absorbed Just Eat Takeaway (Oct 2025) and Despegar (May 2025), Lyft absorbed FreeNow (Jul 2025) — which removed several formerly disclosed figures from public view.

## Methodology

- **Metric.** Each platform's own headline gross transaction metric is used: Marketplace GOV (DoorDash), Gross Bookings (Uber, Lyft, Booking, Expedia, MakeMyTrip), GMV (Delivery Hero, Grab, Meituan), GTV (Instacart, Didi, Just Eat, Klook, Ixigo), GBV (Airbnb, Viator, HomeToGo), TTV (HBX, WebBeds), Gross Travel Value (lastminute.com), net ticket sales (Trainline). Definitions are not perfectly comparable — e.g., DoorDash GOV and Instacart GTV include taxes/tips/fees while Uber's bookings exclude tips; Booking/Expedia/Airbnb are net of cancellations while Viator is gross.
- **Period.** Latest completed full fiscal year: calendar 2025 for most; FY ending March 2026 for Eternal, Swiggy, MakeMyTrip, eDreams, Ixigo, WEB Travel; February 2026 for Trainline; September 2025 for HBX. Tongcheng (FY2024), Despegar (FY2024) and Hopper (CY2024) are the last years those companies disclosed.
- **FX.** Full-period average rates: EUR/USD ~1.13 (1.10 for travel names per HBX's stated average), GBP/USD ~1.29–1.30, CNY 7.15–7.19/USD, RUB 83.8/USD, INR 87–94/USD depending on fiscal period, AUD/USD 0.65, KRW 1,400/USD.
- **Basis.** Every line is tagged `disclosed` or `estimate` in [`data/marketplace_bookings_2025.csv`](../data/marketplace_bookings_2025.csv). Estimates are used only where companies do not publish the figure (mostly China delivery and private companies) and are drawn from company qualitative disclosures, card-panel data, and credible industry research — treat them as directional.
- **Reproducibility.** `python3 scripts/aggregate_bookings.py` recomputes all subtotals from the CSV.

## Sector 1: Food & grocery delivery (~$651B)

| Company | Metric | Period | USD bn | Basis |
|---|---|---|---:|---|
| Meituan (food delivery + Instashopping) | GTV | CY2025 | ~182.5 | Estimate (range $175–190B) |
| DoorDash (incl. Wolt; Deliveroo from Oct 2025) | Marketplace GOV | CY2025 | 102.0 | Disclosed (+27% YoY) |
| Uber Delivery segment | Gross Bookings | CY2025 | 90.9 | Disclosed (+22% YoY) |
| Ele.me / Taobao Instant Commerce (Alibaba) | GMV | CY2025 | ~82.5 | Rough estimate (range $70–95B) |
| Delivery Hero (incl. talabat) | GMV | CY2025 | 55.6 | Disclosed (€49.2B, +9% LFL cc) |
| Instacart | GTV | CY2025 | 37.2 | Disclosed (+11% YoY) |
| Just Eat Takeaway.com (Prosus, private since Oct 2025) | GTV | CY2025 | ~21.0 | Estimate (€18–19B) |
| Grab Deliveries segment | GMV | CY2025 | 14.2 | Disclosed (+21% YoY) |
| Eternal (Zomato food delivery + Blinkit) | GOV (est. from disclosed NOV of ~$11B) | FY Mar-2026 | ~12.2 | Estimate |
| iFood core food delivery (Prosus, private) | GMV | FY Mar-2026 | ~11.5 | Estimate (R$60–65B) |
| JD.com food delivery (launched Feb 2025) | GMV | CY2025 | ~9.5 | Rough estimate |
| Coupang Eats | GMV | CY2025 | ~9.0 | Estimate ($8–10B) |
| Rappi (private) | GMV | CY2025 | ~8.0 | Low-confidence estimate |
| Swiggy (food delivery + Instamart) | GOV | FY Mar-2026 | 7.8 | Disclosed (+45.5% B2C) |
| Deliveroo Jan–Sep 2025 (pre-DoorDash stub) | GTV | 9M 2025 | ~6.7 | Estimate |
| *Memo: talabat (inside Delivery Hero)* | *GMV* | *CY2025* | *9.4* | *Disclosed, not added* |

Sector color: DoorDash ($102.0B) and Uber Delivery ($90.9B) are the two largest disclosed totals outside China, but Meituan remains the world's largest delivery platform by transaction value even though it stopped disclosing absolute GTV. Quick commerce was the growth story of the year — Blinkit NOV +117%, Swiggy Instamart GOV +94%, and Alibaba's Taobao Instant Commerce scaling from near-zero in May 2025 to an estimated $70–95B run. Checked but below the ~$5B inclusion bar: Zepto (~$4–5B annualized GOV), GoTo/GoFood (~$2.6B), Grubhub (sold to Wonder, no longer disclosed).

## Sector 2: Ridesharing / ride-hailing (~$228B)

| Company | Metric | Period | USD bn | Basis |
|---|---|---|---:|---|
| Uber Mobility segment | Gross Bookings | CY2025 | 97.5 | Disclosed (+17.5% YoY) |
| Didi Global (China Mobility $46.4B + International $16.3B) | GTV | CY2025 | 62.7 | Disclosed (RMB 450.8B, +14.8%) |
| Yandex Ride-tech (Yandex Go; incl. carsharing/scooters) | GTV | CY2025 | 19.5 | Disclosed (RUB 1,633B, +11%) |
| Lyft (incl. FreeNow from Aug 2025) | Gross Bookings | CY2025 | 18.5 | Disclosed (+15% YoY) |
| Bolt — ride-hailing portion (private) | GMV | CY2025 | ~11.0 | Estimate (from $14B all-vertical run-rate) |
| Grab Mobility segment | GMV | CY2025 | 7.9 | Disclosed (+19% YoY) |
| inDrive (incl. some deliveries) | Gross Bookings | CY2025 | 6.4 | Disclosed, preliminary (+30%) |
| Cao Cao Mobility | GTV | CY2025 | 3.3 | Disclosed (RMB 23.4B, +38%) |
| Ola / ANI Technologies (private) | GOV | FY Mar-2025 | ~1.3 | Estimate (declining) |
| *Memo: Careem (inside Uber), FreeNow (inside Lyft), Kakao Mobility (GTV not disclosed; ~$0.5B net revenue)* | | | — | Not added |

Sector color: Uber Mobility alone is ~43% of tracked global ride-hailing bookings. Didi's International segment (+28% YoY) is the fastest-growing large piece, though it includes 99Food delivery in LatAm. The tail is thinner than delivery or travel: Kakao Mobility's dominant Korean taxi GTV is simply unpublished, China's aggregators (Amap) would double-count supplier platforms, and India's volume leader Rapido (~$1.5–2B GMV) is below the inclusion bar.

## Sector 3: Online travel (~$661B; ~$648B ex-bed banks)

| Company | Metric | Period | USD bn | Basis |
|---|---|---|---:|---|
| Booking Holdings | Gross bookings | CY2025 | 186.1 | Disclosed (+12.4% YoY) |
| Trip.com Group (core OTA) | Gross bookings | CY2025 | ~157.0 | Management-stated (~RMB 1.1T on Q4 call) |
| Expedia Group (B2C $83.9B + B2B $35.7B) | Gross bookings | CY2025 | 119.6 | Disclosed (+7.8% YoY) |
| Airbnb | Gross Booking Value | CY2025 | 91.3 | Disclosed (+12% YoY) |
| Tongcheng Travel | GMV | CY2024 (last disclosed) | ~36.0 | Disclosed, stale |
| MakeMyTrip | Gross Bookings | FY Mar-2026 | 10.4 | Disclosed (+10.4% cc) |
| HBX Group / Hotelbeds — B2B | TTV | FY Sep-2025 | 9.0 | Disclosed (€8.2B, +7%) |
| Trainline | Net ticket sales | FY Feb-2026 | 8.2 | Disclosed (£6.3B, +7%) |
| Hopper (private; ~40% fintech attach) | Total bookings | CY2024 (latest) | ~7.5 | Third-party estimate |
| Despegar (Prosus, private since May 2025) | Gross Bookings | CY2024 (last public) | 5.5 | Disclosed, stale |
| eDreams ODIGEO | Gross Bookings | FY Mar-2026 | 5.2 | Disclosed (€4.7B, −8.7%) |
| GetYourGuide (private) | GMV | CY2025 | 4.8 | Company-announced (€4B+) |
| Viator + Experiences (Tripadvisor) | GBV | CY2025 | 4.7 | Disclosed (+13% YoY) |
| WEB Travel Group / WebBeds — B2B | TTV | FY Mar-2026 | 3.8 | Disclosed (A$5.8B, +20%) |
| lastminute.com group | Gross Travel Value | CY2025 | 3.8 | Disclosed (€3.5B, +19%) |
| Klook (private; F-1 filed) | GTV | CY2025 | ~3.25 | Annualized from disclosed 9M (+31%) |
| HomeToGo | GBV | CY2025 | 2.5 | Disclosed (€2.2B, +4.6%) |
| Ixigo | GTV | FY Mar-2026 | 2.0 | Disclosed (+25%) |

Sector color: the four global majors (Booking, Trip.com, Expedia, Airbnb) account for ~$554B — 84% of the sector. Experiences was the fastest-growing niche (GetYourGuide GMV +30% in Q3'25, Klook GTV +31%, Viator +13%). Not quantifiable with credible sources: Traveloka (largest SEA OTA; plausibly mid-single-digit billions but no substantiated figure) and Fliggy (Alibaba does not disclose). Below the ~$1B bar: Yatra ($0.86B), Webjet Group ($0.95B).

## Double-counting adjustments applied

1. **talabat is inside Delivery Hero** (majority-owned post-IPO) — listed as memo, not added.
2. **Deliveroo:** DoorDash's $102.0B includes only the Oct–Dec 2025 stub (~$3.3B). A separate line adds Deliveroo's pre-acquisition Jan–Sep bookings (~$6.7B) so 2025 is fully covered without overlap.
3. **Careem is inside Uber Mobility; FreeNow is inside Lyft** (from Aug 2025); Wolt and Trendyol Go are inside DoorDash and Uber Delivery respectively — none added separately.
4. **Uber and Grab span two sectors:** only Uber Mobility / Grab Mobility count in ridesharing and only Uber Delivery / Grab Deliveries in delivery (Uber Freight, ~$5.1B, is excluded entirely).
5. **Bolt's $14B run-rate covers rides + delivery + rentals**; only the estimated ride-hailing portion (~$11B) is counted, in ridesharing only.
6. **B2B bed banks (HBX, WebBeds, and Expedia's $35.7B B2B segment)** power retail sellers, so part of the same hotel stay can appear twice within travel. The adjusted travel subtotal (~$648B) removes the standalone bed banks; Expedia B2B remains in Expedia's reported figure.
7. **Didi International includes 99Food** (delivery) with no split disclosed; counted once, in ridesharing. Similarly inDrive's bookings include some deliveries and Yandex Ride-tech includes carsharing/scooters.

## Key caveats

- **China delivery is the softest part of the aggregate.** Meituan, Alibaba and JD disclosed no absolute transaction values for 2025; the ~$275B combined estimate triangulates market sizing, disclosed order run-rates and average order values in a year when a subsidy war inflated volumes. The plausible range on the delivery sector total is roughly ±$50B.
- **Stale entries:** Tongcheng (~$36B) and Despegar ($5.5B) are 2024 figures; Hopper is a 2024 third-party estimate. Using them slightly understates a true 2025 total.
- **Eternal reports NOV (net of discounts)** as its headline; the GOV estimate applies prior-year GOV/NOV ratios.
- **Metric definitions vary** (tips, taxes, cancellations, fintech attach volume) — cross-company comparisons are approximate by construction.
- Full per-company sourcing (earnings release / filing URLs) is in the `source` column of [`data/marketplace_bookings_2025.csv`](../data/marketplace_bookings_2025.csv).
