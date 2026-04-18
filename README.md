# Equity-research

Models and visualizations supporting equity research on the AI inference market.

## AI Token Consumption Model

`AI_Token_Consumption_Model.xlsx` is a quarterly forecast (2024 – 2028) of
global AI inference token consumption split across **Text**, **Voice** and
**Video** modalities. The workbook contains:

| Sheet | Contents |
|-------|----------|
| `Assumptions` | Editable yellow driver cells: starting MAU, user CAGR, sessions/user/day, tokens/session, days/quarter — for each modality. |
| `Forecast`    | Formula-driven quarterly and annual token consumption (trillions of tokens), with totals. |
| `Charts`      | Time-series **line chart** of quarterly tokens by modality, plus **clustered** and **stacked bar charts** of annual tokens by modality. |

All forecast values are derived from formulas referencing the `Assumptions`
sheet, so flexing any driver instantly updates the model and all charts.

### Default driver assumptions

| Driver                       | Text   | Voice  | Video   |
|------------------------------|-------:|-------:|--------:|
| Starting MAU (M, Q1 2024)    | 500    | 60     | 12      |
| User CAGR (annualized)       | 45%    | 85%    | 120%    |
| Sessions / user / day        | 8      | 2      | 0.5     |
| Tokens / session             | 750    | 6,000  | 60,000  |

These benchmarks reflect typical 2025 multimodal inference profiles
(text ≈ 500–1,000 tokens per round-trip; voice ≈ 3,000 tokens/min via
ASR + LLM + TTS; video ≈ 50,000 tokens/min via patch-token vision models).

### Regenerate the workbook

```bash
pip install -r requirements.txt
python build_model.py
```

This (re)writes `AI_Token_Consumption_Model.xlsx` from scratch using
`build_model.py`.
