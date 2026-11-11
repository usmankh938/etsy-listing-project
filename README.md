# Ultimate Annual & Monthly Budget Planner (Excel)

A formula-driven Excel financial planner generated in Python — built as a stronger alternative to typical beige Etsy “ultimate financial planner” templates.

## Deliverables

| File | Description |
|------|-------------|
| [`artifacts/Ultimate_Annual_Monthly_Budget_Planner.xlsx`](artifacts/Ultimate_Annual_Monthly_Budget_Planner.xlsx) | Full annual + monthly budget workbook (27 sheets, sample year pre-filled) |
| [`artifacts/financial_planner_dashboard.png`](artifacts/financial_planner_dashboard.png) | Executive dashboard chart from the analytics script |
| [`generate_budget_planner.py`](generate_budget_planner.py) | Regenerates the Excel workbook |
| [`generate_dashboard.py`](generate_dashboard.py) | Generates the PNG analytics dashboard |

## What’s inside the Excel file

1. **Instructions** — 10-minute setup guide + tab map  
2. **Setup** — year, goals, 50/30/20 targets, category → Needs/Wants/Savings mapping  
3. **Accounts** — bank / investment / credit snapshot with net worth  
4. **Recurring** — set income & bills once; annual totals auto-calculate  
5. **Transactions** — variable income/expense log (bank-export friendly)  
6. **January–December** — 12 monthly Budget vs Actual dashboards  
7. **Annual Dashboard** — year matrix, income vs expenses, savings rate, charts  
8. **Quarterly** — Q1–Q4 rollups  
9. **50-30-20** — Needs / Wants / Savings target vs actual  
10. **Net Worth** — month-end asset/liability tracker  
11. **Debt Tracker** — snowball + avalanche ranks + payoff estimator (NPER)  
12. **Sinking Funds** — goal jars with suggested monthly contributions  
13. **Emergency Fund** — runway in months of essential expenses  
14. **Wealth Projection** — 5-year compound growth from your surplus  
15. **Health Score** — automated 0–100 financial health score  
16. **Bill Calendar** — due-day payment status grid synced to Recurring  

Cream/input cells are editable. All dashboards recalculate from your entries.

## Why this beats a typical Etsy beige planner

- Linked annual ↔ monthly formulas (edit once, see it everywhere)
- Financial health scoring (savings rate, runway, debt burden, cash-flow months)
- Debt snowball **and** avalanche with interest-aware payoff estimate
- Wealth projection tied to your real average monthly surplus
- Sample year included so the file is useful the moment you open it

## Regenerate

```bash
pip install pandas openpyxl matplotlib numpy
python generate_budget_planner.py
python generate_dashboard.py
```

Outputs land in `output/` and `artifacts/`.
