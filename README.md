# BRRRR Deal Analyzer Spreadsheet

**Analyze BRRRR, Flip & Buy-and-Hold Deals in 60 Seconds**

A comprehensive Google Sheet (downloadable as `.xlsx`) for real estate investors that runs BRRRR, flip, and buy-and-hold scenarios on any property instantly.

## 🏠 What's Inside

| Tab | Description |
|-----|-------------|
| **Deal Input** | Enter purchase price, rehab budget, ARV, financing terms |
| **BRRRR Analyzer** | Cash-on-cash return, ARV refinance proceeds with seasoning period |
| **Flip Analyzer** | Separate tab for flip scenarios with ROI calculation |
| **Buy & Hold 30yr** | 30-year projection with appreciation and rent-growth assumptions |
| **Deal Score** | 1-page summary with deal-quality score (green/yellow/red) |
| **Refi Scenarios** | Compare refinance outcomes at different LTVs and ARV assumptions |
| **Deal Criteria** | Define your buy box — sheet auto-flags matching deals |

## 🔑 Key Formula

```
ARV × Refi LTV − Purchase × 1.06 (Closing) − Rehab − Holding × 6 months = Cash you actually pull out
```

Most competitors miss the refinance seasoning math. This sheet includes it.

## 📦 Usage

1. Open the `.xlsx` in Google Sheets, Excel, or LibreOffice
2. Enter your property numbers in the **Deal Input** tab (yellow cells)
3. Instantly see results across all analysis tabs
4. Use the **Deal Criteria** tab to match against your buy box

## 🛠 Build

```bash
python build_spreadsheet.py output/BRRRR_Deal_Analyzer.xlsx
python generate_cover.py
```

## License

© Logan Chandra — All rights reserved.
