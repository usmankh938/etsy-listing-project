# Premium Household Budget Planner (Excel & Google Sheets)

A formula-driven **household budget book** for Excel and Google Sheets — income & expense tracking, savings, debt, and financial planning — built to outperform typical beige Etsy financial planners.

## Download

| File | Description |
|------|-------------|
| [`artifacts/Premium_Household_Budget_Planner.xlsx`](artifacts/Premium_Household_Budget_Planner.xlsx) | **Primary deliverable** — 35-sheet premium workbook (sample year included) |
| [`artifacts/Ultimate_Annual_Monthly_Budget_Planner.xlsx`](artifacts/Ultimate_Annual_Monthly_Budget_Planner.xlsx) | Same workbook (alternate filename) |
| [`artifacts/financial_planner_dashboard.png`](artifacts/financial_planner_dashboard.png) | Executive analytics dashboard PNG |

## What’s inside (35 sheets)

**Setup & household**
- Instructions · Google Sheets Guide · Setup · Household (dual-income) · Accounts · Recurring

**Income & expense tracking**
- Transactions (expense tracker) · Subscriptions audit · Weekly Spend · Paycheck Budget · Bill Calendar

**12 monthly dashboards** — January–December Budget vs Actual

**Big picture**
- Annual Dashboard (+ charts) · Distribution · Quarterly · 50/30/20 · Health Score

**Savings, debt & wealth**
- Net Worth · Debt Tracker (snowball + avalanche) · Sinking Funds · Emergency Fund · Investments · Wealth Projection · No-Spend Challenge

Cream/input cells are editable. Dashboards recalculate automatically.

## Better than typical Etsy beige planners

- Dual-income household cash flow (Person A / Person B)
- Paycheck budgeting for bi-weekly / irregular income
- Subscriptions + weekly envelopes + full transaction log
- Setup-once recurring bills that feed all 12 months
- Automated financial health score (0–100)
- Debt snowball **and** avalanche with NPER payoff estimates
- Investments forecast + 5-year wealth projection
- Google Sheets–ready `.xlsx` (Drive → Open with Google Sheets)
- Sample year pre-filled so it’s useful immediately

## Regenerate

```bash
pip install pandas openpyxl matplotlib numpy
python generate_budget_planner.py
python generate_dashboard.py
```
