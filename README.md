# AI Lab Compute Economics (Quarterly)

Excel workbook aggregating **operational GW capacity**, **revenue**, **compute cost**, **revenue/cost per GW**, and **token volume** for frontier (closed) and open-weight AI labs on a quarterly time series from **Q1 2024 through Q3 2026 (partial)**.

## File

[`ai_lab_compute_economics_quarterly.xlsx`](ai_lab_compute_economics_quarterly.xlsx)

## Labs covered

| Category | Labs |
|---|---|
| Frontier (closed) | OpenAI, Anthropic, Google DeepMind, xAI |
| Open-weight | Meta (Llama), DeepSeek, Alibaba (Qwen), Mistral AI |

## Sheets

1. **Read Me** — how to read the workbook and headline takeaways  
2. **Summary (Q2 2026)** — latest full-quarter snapshot with derived metrics  
3. **GW Capacity** — operational IT power (owned + leased), end of quarter  
4. **Revenue (Quarterly)** — booked revenue, US$B  
5. **Revenue Run Rate** — annualized run rate at quarter end (what labs quote)  
6. **Compute Cost** — estimated quarterly cost from Epoch AI $8.5B/GW/yr TCO  
7. **Rev & Cost per GW** — derived unit economics vs ~$31B/GW/yr break-even  
8. **Token Volume** — tokens processed per quarter (trillions), all surfaces  
9. **Sources & Methodology** — disclosed anchors with URLs + caveats  

Red bold cells are directly disclosed figures; everything else is estimated/interpolated between anchors. See the Sources sheet for every datapoint.

## Rebuild

```bash
python3 scripts/build_workbook.py
```

Requires `openpyxl`. Compiled from public reporting as of Aug 23, 2026.
