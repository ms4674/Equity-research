# Equity-research

## Sources of Alpha & the Bloomberg Terminal (presentation)

A 15-slide PowerPoint deck covering the major sources of alpha in the market
and the Bloomberg Terminal workflow used to discover, validate and monitor
each one.

- **Deck:** [`alpha_sources_bloomberg.pptx`](alpha_sources_bloomberg.pptx)
- **Generator:** [`generate_presentation.py`](generate_presentation.py)

### Contents

1. Title
2. Alpha 101 — definition, Jensen's alpha, hidden beta
3. Framework — four edges (informational, analytical, behavioral, structural), eight alpha sources
4. Factor & style premia (`EQS`, `EQBT`, `PORT`, `FTW`)
5. Fundamental stock selection (`FA`, `MODL`, `EE`, `SURP`, `ANR`, `RV`, `BI`)
6. Event-driven & special situations (`MA`, `MARB`, `EVTS`, `CACS`, `MEMB`, `ECDR`, `DVD`)
7. Macro & cross-asset (`ECO`, `WIRP`/`MIPR`, `ECFC`, `FXFC`, `ECST`)
8. Carry & term structure (`CCRV`, `FWCV`, `FXFA`, `GV`, `SKEW`)
9. Flow, positioning & ownership (`HDS`, `OWN`, `SI`, `OMON`, `CFTC`, `FLNG`)
10. Sentiment, news & alternative data (`NSE`, `TREN`, `SPLC`, `BI`)
11. Statistical arbitrage & quant research (`HRA`, `CORR`, `BQNT`, `BQL`, `FLDS`)
12. Workflow — from idea to alpha, six-step terminal pipeline
13. Bloomberg function cheat sheet
14. Why alpha dies — overfitting, crowding, costs, hidden beta
15. Key takeaways

### Regenerating the deck

```bash
pip install -r requirements.txt
python3 generate_presentation.py
```

The script produces `alpha_sources_bloomberg.pptx` (16:9, dark
terminal-inspired theme with amber function mnemonics).
