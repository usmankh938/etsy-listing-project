#!/usr/bin/env python3
"""
Premium Household Budget Planner (Excel + Google Sheets ready)
--------------------------------------------------------------
A polished, formula-driven household budget book that beats typical beige
Etsy planners: dual-income household setup, paycheck budgeting, expense
tracker, subscriptions, investments, no-spend challenges, setup-once
recurring income/bills, 12 monthly dashboards, annual + quarterly views,
50/30/20, net worth, debt snowball/avalanche, sinking funds, emergency
fund runway, wealth projection, and an automated financial health score.
"""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# Theme (warm, calm, non-generic — slate + sage + sand)
# ---------------------------------------------------------------------------
COLORS = {
    "bg": "F7F4EF",
    "header": "3D4F5F",
    "accent": "6B8F71",
    "accent2": "C4A484",
    "warn": "C4785A",
    "soft": "EDE6DC",
    "soft2": "D9E3D6",
    "white": "FFFFFF",
    "text": "2A2A2A",
    "muted": "6B6B6B",
    "input": "FFFCF7",
    "good": "E8F2EA",
    "bad": "F8E8E3",
}

THIN = Border(
    left=Side(style="thin", color="D4CBBE"),
    right=Side(style="thin", color="D4CBBE"),
    top=Side(style="thin", color="D4CBBE"),
    bottom=Side(style="thin", color="D4CBBE"),
)

MONTHS = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]
MONTH_SHORT = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
]

EXPENSE_CATS = [
    "Housing",
    "Utilities",
    "Groceries",
    "Transportation",
    "Insurance",
    "Healthcare",
    "Subscriptions",
    "Dining Out",
    "Entertainment",
    "Personal Care",
    "Clothing",
    "Education",
    "Kids / Family",
    "Travel",
    "Gifts / Donations",
    "Debt Payments",
    "Savings Transfers",
    "Miscellaneous",
]

INCOME_CATS = [
    "Salary",
    "Side Hustle",
    "Freelance",
    "Investments",
    "Rental Income",
    "Other Income",
]


def fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def font(size=11, bold=False, color=None, name="Calibri") -> Font:
    return Font(name=name, size=size, bold=bold, color=color or COLORS["text"])


def money_format(cell):
    cell.number_format = '"$"#,##0.00'


def pct_format(cell):
    cell.number_format = '0.0%'


def set_col_widths(ws, widths: dict):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def style_header_row(ws, row: int, start_col: int, end_col: int):
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = fill(COLORS["header"])
        cell.font = font(11, bold=True, color=COLORS["white"])
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN


def title_block(ws, text: str, subtitle: str | None = None, merge_to: str = "H1"):
    ws.merge_cells(f"A1:{merge_to}")
    cell = ws["A1"]
    cell.value = text
    cell.font = font(20, bold=True, color=COLORS["white"])
    cell.fill = fill(COLORS["header"])
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 36
    if subtitle:
        ws.merge_cells(f"A2:{merge_to.replace('1', '2')}")
        c2 = ws["A2"]
        c2.value = subtitle
        c2.font = font(11, color=COLORS["muted"])
        c2.fill = fill(COLORS["soft"])
        c2.alignment = Alignment(vertical="center")
        ws.row_dimensions[2].height = 22


def section_label(ws, cell_ref: str, text: str, color: str | None = None):
    cell = ws[cell_ref]
    cell.value = text
    cell.font = font(13, bold=True, color=COLORS["white"])
    cell.fill = fill(color or COLORS["accent"])
    cell.alignment = Alignment(horizontal="left", vertical="center")


def input_cell(cell, value=None):
    if value is not None:
        cell.value = value
    cell.fill = fill(COLORS["input"])
    cell.border = THIN
    cell.font = font(11)


def label_cell(cell, text: str):
    cell.value = text
    cell.font = font(11, bold=True)
    cell.fill = fill(COLORS["soft"])
    cell.border = THIN


def kpi_box(ws, row: int, col: int, label: str, formula_or_value, is_pct=False):
    label_c = ws.cell(row=row, column=col, value=label)
    label_c.font = font(10, bold=True, color=COLORS["muted"])
    label_c.fill = fill(COLORS["soft"])
    label_c.alignment = Alignment(horizontal="center")
    label_c.border = THIN

    value_c = ws.cell(row=row + 1, column=col, value=formula_or_value)
    value_c.font = font(14, bold=True, color=COLORS["header"])
    value_c.fill = fill(COLORS["white"])
    value_c.alignment = Alignment(horizontal="center", vertical="center")
    value_c.border = THIN
    if is_pct:
        pct_format(value_c)
    else:
        money_format(value_c)
    ws.row_dimensions[row + 1].height = 28


def generate_sample_monthly() -> pd.DataFrame:
    np.random.seed(42)
    rows = []
    for i, month in enumerate(MONTH_SHORT):
        salary = 5500
        side = int(np.random.randint(400, 900))
        housing = 1800
        groceries = float(np.clip(np.random.normal(650, 50), 500, 800))
        utilities = float(np.clip(np.random.normal(250, 30), 200, 350))
        transport = float(np.clip(np.random.normal(300, 40), 220, 400))
        insurance = 220
        healthcare = float(np.clip(np.random.normal(80, 40), 20, 200))
        subscriptions = 95
        dining = float(np.clip(np.random.normal(280, 60), 120, 450))
        entertainment = float(np.clip(np.random.normal(150, 40), 50, 280))
        personal = float(np.clip(np.random.normal(90, 25), 40, 160))
        clothing = float(np.clip(np.random.normal(80, 50), 0, 220))
        education = 50 if i % 3 == 0 else 0
        kids = 0
        travel = 600 if month in ("Jun", "Dec") else (150 if month in ("Mar", "Sep") else 0)
        gifts = 200 if month == "Dec" else (50 if month in ("Feb", "May", "Nov") else 25)
        debt = 400
        savings_xfer = 800
        misc = float(np.clip(np.random.normal(75, 30), 20, 160))
        rows.append(
            {
                "Month": month,
                "Salary": salary,
                "Side_Hustle": side,
                "Housing": round(housing, 2),
                "Utilities": round(utilities, 2),
                "Groceries": round(groceries, 2),
                "Transportation": round(transport, 2),
                "Insurance": insurance,
                "Healthcare": round(healthcare, 2),
                "Subscriptions": subscriptions,
                "Dining_Out": round(dining, 2),
                "Entertainment": round(entertainment, 2),
                "Personal_Care": round(personal, 2),
                "Clothing": round(clothing, 2),
                "Education": education,
                "Kids_Family": kids,
                "Travel": travel,
                "Gifts_Donations": gifts,
                "Debt_Payments": debt,
                "Savings_Transfers": savings_xfer,
                "Miscellaneous": round(misc, 2),
            }
        )
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------


def build_instructions(ws):
    title_block(
        ws,
        "PREMIUM HOUSEHOLD BUDGET PLANNER",
        "Excel & Google Sheets ready · Household budget book · Income, expenses, savings & debt — better than beige Etsy templates.",
        "G1",
    )
    set_col_widths(ws, {"A": 28, "B": 58, "C": 22, "D": 18, "E": 18, "F": 18, "G": 18})

    ws["A4"] = "How to use (10 minutes)"
    ws["A4"].font = font(14, bold=True, color=COLORS["header"])

    steps = [
        ("1. Setup", "Open Setup. Enter year, currency, household name, partners, pay frequency, and savings goals."),
        ("2. Household", "Set Person A / Person B incomes and shared vs personal expense splits."),
        ("3. Accounts", "List every bank, investment, and credit account on Accounts."),
        ("4. Recurring", "Add paychecks & bills once. They feed Annual + Monthly sheets all year."),
        ("5. Track spend", "Log variable expenses in Transactions (or paste a bank CSV). Use Subscriptions + Weekly Spend."),
        ("6. Paycheck", "If you budget by paycheck, use Paycheck Budget for bi-weekly / irregular income."),
        ("7. Debt & savings", "Update Debt Tracker, Sinking Funds, Investments, Emergency Fund."),
        ("8. Review", "Check Annual Dashboard, 50/30/20, Health Score, and Wealth Projection weekly."),
    ]
    for i, (title, body) in enumerate(steps, start=5):
        ws.cell(row=i, column=1, value=title).font = font(11, bold=True, color=COLORS["accent"])
        ws.cell(row=i, column=2, value=body).font = font(11)

    ws["A14"] = "Why this beats typical Etsy beige / household planners"
    ws["A14"].font = font(14, bold=True, color=COLORS["header"])
    upgrades = [
        "True household dual-income planner (Person A / Person B) with combined cash-flow view.",
        "Paycheck Budget for bi-weekly and irregular income — not just calendar months.",
        "Full Expense Tracker (Transactions) + Subscriptions audit + Weekly variable spend.",
        "Setup-once Recurring engine that rolls into all 12 months (no retyping bills).",
        "Annual Dashboard with charts, Quarterly rollups, and Income vs Expense distribution.",
        "Financial Health Score (savings rate, emergency runway, debt burden, cash-flow months).",
        "Debt Snowball AND Avalanche with interest-aware payoff estimates (NPER).",
        "50/30/20 Needs / Wants / Savings with automatic category mapping.",
        "Sinking funds, emergency runway, investments forecast, and 5-year wealth projection.",
        "No-Spend Challenge tracker to cut discretionary leaks fast.",
        "Google Sheets–compatible formulas (open the .xlsx in Drive → Open with Google Sheets).",
        "Sample year pre-filled so the file is useful the moment you open it.",
    ]
    for i, tip in enumerate(upgrades, start=15):
        ws.cell(row=i, column=1, value="•").font = font(12, bold=True, color=COLORS["accent"])
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=6)
        ws.cell(row=i, column=2, value=tip)

    ws["A28"] = "Tab map"
    ws["A28"].font = font(14, bold=True, color=COLORS["header"])
    tabs = [
        ("Instructions", "You are here"),
        ("Google Sheets Guide", "How to open this file in Google Sheets"),
        ("Setup", "Year, currency, goals, category rules"),
        ("Household", "Dual-income household cash flow"),
        ("Accounts", "Bank / investment / credit snapshot"),
        ("Recurring", "Income & bills that auto-fill months"),
        ("Transactions", "Expense & income tracker log"),
        ("Subscriptions", "Recurring software/media/service audit"),
        ("Weekly Spend", "Flexible weekly spending envelopes"),
        ("Paycheck Budget", "Bi-weekly / irregular paycheck planner"),
        ("Jan … Dec", "12 monthly Budget vs Actual dashboards"),
        ("Annual Dashboard", "Year overview + charts"),
        ("Distribution", "Income & expense category mix"),
        ("Quarterly", "Q1–Q4 rollups"),
        ("50-30-20", "Needs / Wants / Savings rule"),
        ("Net Worth", "Assets − liabilities over time"),
        ("Debt Tracker", "Snowball & avalanche payoff tools"),
        ("Sinking Funds", "Goal-based savings jars"),
        ("Emergency Fund", "Runway calculator"),
        ("Investments", "Portfolio contribution forecast"),
        ("Wealth Projection", "5-year compound growth"),
        ("Health Score", "Automated financial health KPIs"),
        ("No-Spend Challenge", "Discretionary spending challenges"),
        ("Bill Calendar", "Due-date payment schedule"),
    ]
    style_header_row(ws, 29, 1, 2)
    ws.cell(row=29, column=1, value="Tab")
    ws.cell(row=29, column=2, value="Purpose")
    for i, (tab, purpose) in enumerate(tabs, start=30):
        ws.cell(row=i, column=1, value=tab).border = THIN
        ws.cell(row=i, column=2, value=purpose).border = THIN
        if i % 2 == 0:
            ws.cell(row=i, column=1).fill = fill(COLORS["soft"])
            ws.cell(row=i, column=2).fill = fill(COLORS["soft"])

    ws["A56"] = "Tip: Cream cells are inputs. Everything else is formula-driven — don’t overwrite formulas."
    ws["A56"].font = font(11, bold=True, color=COLORS["warn"])
    ws.merge_cells("A56:G56")


def build_setup(ws):
    title_block(ws, "SETUP — Do this once", "Cream cells are editable inputs. Linked sheets read from here.", "F1")
    set_col_widths(ws, {"A": 28, "B": 22, "C": 18, "D": 28, "E": 18, "F": 22})

    section_label(ws, "A4", "Profile & Goals", COLORS["header"])
    ws.merge_cells("A4:B4")

    fields = [
        (5, "Budget Year", 2026),
        (6, "Currency Symbol", "$"),
        (7, "Household Name", "Sample Household"),
        (8, "Person A Name", "Alex"),
        (9, "Person B Name", "Jordan"),
        (10, "Pay Frequency", "Bi-weekly"),
        (11, "Annual Savings Goal", 24000),
        (12, "Starting Cash (Jan 1)", 8500),
        (13, "Emergency Fund Target (months)", 6),
        (14, "Expected Investment Return %", 0.075),
        (15, "Starting Investable Nest Egg", 15000),
    ]
    for row, label, value in fields:
        label_cell(ws.cell(row=row, column=1), label)
        c = ws.cell(row=row, column=2)
        input_cell(c, value)
        if row in (11, 12, 15):
            money_format(c)
        if row == 14:
            pct_format(c)

    freq_dv = DataValidation(type="list", formula1='"Monthly,Bi-weekly,Weekly,Semi-monthly"', allow_blank=False)
    ws.add_data_validation(freq_dv)
    freq_dv.add("B10")

    section_label(ws, "A17", "50/30/20 Targets", COLORS["accent"])
    ws.merge_cells("A17:B17")
    for row, label, value in [(18, "Needs %", 0.50), (19, "Wants %", 0.30), (20, "Savings/Debt %", 0.20)]:
        label_cell(ws.cell(row=row, column=1), label)
        c = ws.cell(row=row, column=2)
        input_cell(c, value)
        pct_format(c)

    section_label(ws, "D4", "Category → Bucket Mapping", COLORS["header"])
    ws.merge_cells("D4:E4")
    ws["D5"] = "Expense Category"
    ws["E5"] = "Bucket (Needs/Wants/Savings)"
    style_header_row(ws, 5, 4, 5)

    mapping = {
        "Housing": "Needs",
        "Utilities": "Needs",
        "Groceries": "Needs",
        "Transportation": "Needs",
        "Insurance": "Needs",
        "Healthcare": "Needs",
        "Subscriptions": "Wants",
        "Dining Out": "Wants",
        "Entertainment": "Wants",
        "Personal Care": "Wants",
        "Clothing": "Wants",
        "Education": "Needs",
        "Kids / Family": "Needs",
        "Travel": "Wants",
        "Gifts / Donations": "Wants",
        "Debt Payments": "Savings",
        "Savings Transfers": "Savings",
        "Miscellaneous": "Wants",
    }
    for i, cat in enumerate(EXPENSE_CATS, start=6):
        input_cell(ws.cell(row=i, column=4), cat)
        input_cell(ws.cell(row=i, column=5), mapping[cat])

    dv = DataValidation(type="list", formula1='"Needs,Wants,Savings"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("E6:E23")

    ws["A22"] = "Key cells used across the workbook"
    ws["A22"].font = font(12, bold=True, color=COLORS["header"])
    ws["A23"] = (
        "B5 Year · B8/B9 Partners · B11 Savings Goal · B12 Starting Cash · "
        "B13 EF months · B14 Return · B15 Nest Egg · B18–B20 50/30/20"
    )
    ws["A23"].font = font(10, color=COLORS["muted"])
    ws.merge_cells("A23:F23")


def build_accounts(ws):
    title_block(ws, "ACCOUNTS TRACKER", "Update balances monthly. Net worth sheet reads from here.", "F1")
    set_col_widths(ws, {"A": 26, "B": 16, "C": 16, "D": 16, "E": 18, "F": 22})

    headers = ["Account Name", "Type", "Institution", "Balance", "As Of", "Notes"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, 1, 6)

    sample = [
        ("Everyday Checking", "Checking", "Local Bank", 3200, "2026-01-01", "Primary spending"),
        ("High-Yield Savings", "Savings", "Online Bank", 9800, "2026-01-01", "Emergency + sinking"),
        ("Brokerage", "Investment", "Vanguard", 15000, "2026-01-01", "Index funds"),
        ("401(k)", "Retirement", "Fidelity", 42000, "2026-01-01", "Employer match"),
        ("Visa Credit Card", "Credit Card", "Chase", -1450, "2026-01-01", "Pay in full monthly"),
        ("Auto Loan", "Loan", "Credit Union", -9800, "2026-01-01", "5.2% APR"),
        ("Student Loan", "Loan", "Mohela", -12400, "2026-01-01", "4.8% APR"),
    ]
    for r, row in enumerate(sample, start=5):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            input_cell(cell)
            if c == 4:
                money_format(cell)

    for r in range(12, 25):
        for c in range(1, 7):
            input_cell(ws.cell(row=r, column=c))
            if c == 4:
                money_format(ws.cell(row=r, column=c))

    type_dv = DataValidation(
        type="list",
        formula1='"Checking,Savings,Investment,Retirement,Credit Card,Loan,Cash,Other"',
        allow_blank=True,
    )
    ws.add_data_validation(type_dv)
    type_dv.add("B5:B24")

    ws["A26"] = "Totals"
    ws["A26"].font = font(12, bold=True)
    label_cell(ws["A27"], "Total Assets")
    ws["B27"] = '=SUMIFS(D5:D24,B5:B24,"<>Credit Card",B5:B24,"<>Loan")'
    money_format(ws["B27"])
    label_cell(ws["A28"], "Total Liabilities")
    ws["B28"] = '=ABS(SUMIF(B5:B24,"Credit Card",D5:D24)+SUMIF(B5:B24,"Loan",D5:D24))'
    money_format(ws["B28"])
    label_cell(ws["A29"], "Net Worth (Accounts)")
    ws["B29"] = "=B27-B28"
    money_format(ws["B29"])
    ws["B29"].font = font(13, bold=True, color=COLORS["accent"])


def build_recurring(ws, sample: pd.DataFrame):
    title_block(
        ws,
        "RECURRING INCOME & BILLS",
        "Set once — Annual and Monthly sheets pull these amounts. Toggle Active = Yes/No.",
        "H1",
    )
    set_col_widths(
        ws,
        {"A": 26, "B": 14, "C": 14, "D": 14, "E": 12, "F": 12, "G": 14, "H": 28},
    )

    section_label(ws, "A4", "Recurring Income", COLORS["accent"])
    ws.merge_cells("A4:H4")
    headers = ["Name", "Category", "Amount", "Frequency", "Day of Month", "Active", "Annual Total", "Notes"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=5, column=i, value=h)
    style_header_row(ws, 5, 1, 8)

    income_rows = [
        ("Primary Salary", "Salary", 5500, "Monthly", 1, "Yes", "Direct deposit"),
        ("Side Hustle Avg", "Side Hustle", round(float(sample["Side_Hustle"].mean()), 2), "Monthly", 15, "Yes", "Average of sample year"),
    ]
    for r, row in enumerate(income_rows, start=6):
        for c, val in enumerate(row[:6], start=1):
            input_cell(ws.cell(row=r, column=c), val)
        money_format(ws.cell(row=r, column=3))
        ws.cell(row=r, column=7, value=f'=IF(F{r}="Yes",IF(D{r}="Monthly",C{r}*12,IF(D{r}="Bi-weekly",C{r}*26,IF(D{r}="Weekly",C{r}*52,C{r}))),0)')
        money_format(ws.cell(row=r, column=7))
        input_cell(ws.cell(row=r, column=8), row[6])

    for r in range(8, 14):
        for c in range(1, 9):
            input_cell(ws.cell(row=r, column=c))
            if c in (3, 7):
                money_format(ws.cell(row=r, column=c))
        ws.cell(row=r, column=7, value=f'=IF(OR(C{r}="",F{r}<>"Yes"),0,IF(D{r}="Monthly",C{r}*12,IF(D{r}="Bi-weekly",C{r}*26,IF(D{r}="Weekly",C{r}*52,C{r}))))')

    label_cell(ws["A14"], "Annual Recurring Income")
    ws["B14"] = "=SUM(G6:G13)"
    money_format(ws["B14"])
    ws["B14"].font = font(12, bold=True, color=COLORS["accent"])

    section_label(ws, "A16", "Recurring Bills / Expenses", COLORS["warn"])
    ws.merge_cells("A16:H16")
    for i, h in enumerate(headers, start=1):
        ws.cell(row=17, column=i, value=h)
    style_header_row(ws, 17, 1, 8)

    bill_rows = [
        ("Rent / Mortgage", "Housing", 1800, "Monthly", 1, "Yes", "Due on the 1st"),
        ("Electric + Gas", "Utilities", round(float(sample["Utilities"].mean()), 2), "Monthly", 12, "Yes", ""),
        ("Internet + Phone", "Utilities", 95, "Monthly", 8, "Yes", ""),
        ("Car Insurance", "Insurance", 220, "Monthly", 20, "Yes", ""),
        ("Streaming Bundle", "Subscriptions", 45, "Monthly", 5, "Yes", ""),
        ("Gym", "Subscriptions", 50, "Monthly", 3, "Yes", ""),
        ("Auto Loan Payment", "Debt Payments", 285, "Monthly", 15, "Yes", ""),
        ("Student Loan Payment", "Debt Payments", 115, "Monthly", 22, "Yes", ""),
        ("Emergency Fund Auto", "Savings Transfers", 400, "Monthly", 1, "Yes", "HYSA sweep"),
        ("Investment Auto", "Savings Transfers", 400, "Monthly", 1, "Yes", "Brokerage DCA"),
    ]
    for r, row in enumerate(bill_rows, start=18):
        for c, val in enumerate(row[:6], start=1):
            input_cell(ws.cell(row=r, column=c), val)
        money_format(ws.cell(row=r, column=3))
        ws.cell(row=r, column=7, value=f'=IF(F{r}="Yes",IF(D{r}="Monthly",C{r}*12,IF(D{r}="Bi-weekly",C{r}*26,IF(D{r}="Weekly",C{r}*52,C{r}))),0)')
        money_format(ws.cell(row=r, column=7))
        input_cell(ws.cell(row=r, column=8), row[6])

    for r in range(28, 36):
        for c in range(1, 9):
            input_cell(ws.cell(row=r, column=c))
            if c in (3, 7):
                money_format(ws.cell(row=r, column=c))
        ws.cell(row=r, column=7, value=f'=IF(OR(C{r}="",F{r}<>"Yes"),0,IF(D{r}="Monthly",C{r}*12,IF(D{r}="Bi-weekly",C{r}*26,IF(D{r}="Weekly",C{r}*52,C{r}))))')

    label_cell(ws["A36"], "Annual Recurring Bills")
    ws["B36"] = "=SUM(G18:G35)"
    money_format(ws["B36"])
    ws["B36"].font = font(12, bold=True, color=COLORS["warn"])

    freq_dv = DataValidation(type="list", formula1='"Monthly,Bi-weekly,Weekly,Annual"', allow_blank=True)
    active_dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    cat_dv = DataValidation(type="list", formula1='"' + ",".join(INCOME_CATS + EXPENSE_CATS) + '"', allow_blank=True)
    ws.add_data_validation(freq_dv)
    ws.add_data_validation(active_dv)
    ws.add_data_validation(cat_dv)
    freq_dv.add("D6:D13")
    freq_dv.add("D18:D35")
    active_dv.add("F6:F13")
    active_dv.add("F18:F35")
    cat_dv.add("B6:B13")
    cat_dv.add("B18:B35")


def build_transactions(ws, sample: pd.DataFrame):
    title_block(
        ws,
        "TRANSACTIONS LOG",
        "Log one-off and variable spending here. Monthly sheets aggregate by category & month.",
        "H1",
    )
    set_col_widths(
        ws,
        {"A": 12, "B": 12, "C": 18, "D": 16, "E": 28, "F": 14, "G": 12, "H": 22},
    )
    headers = ["Date", "Month", "Type", "Category", "Description", "Amount", "Account", "Notes"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, 1, 8)

    # Seed variable transactions per month from sample discretionary-ish categories
    row = 5
    for m_i, month in enumerate(MONTHS):
        short = MONTH_SHORT[m_i]
        seeds = [
            (f"{2026}-{m_i+1:02d}-04", "Expense", "Groceries", "Weekly groceries", sample.loc[m_i, "Groceries"] * 0.35),
            (f"{2026}-{m_i+1:02d}-11", "Expense", "Groceries", "Weekly groceries", sample.loc[m_i, "Groceries"] * 0.35),
            (f"{2026}-{m_i+1:02d}-18", "Expense", "Dining Out", "Restaurants / takeout", sample.loc[m_i, "Dining_Out"]),
            (f"{2026}-{m_i+1:02d}-22", "Expense", "Entertainment", "Movies / hobbies", sample.loc[m_i, "Entertainment"]),
            (f"{2026}-{m_i+1:02d}-25", "Expense", "Transportation", "Fuel / transit", sample.loc[m_i, "Transportation"]),
            (f"{2026}-{m_i+1:02d}-28", "Expense", "Miscellaneous", "Odds & ends", sample.loc[m_i, "Miscellaneous"]),
        ]
        if sample.loc[m_i, "Travel"] > 0:
            seeds.append((f"{2026}-{m_i+1:02d}-15", "Expense", "Travel", "Trip / weekend", sample.loc[m_i, "Travel"]))
        if sample.loc[m_i, "Side_Hustle"] > 0:
            seeds.append((f"{2026}-{m_i+1:02d}-16", "Income", "Side Hustle", "Side project payout", sample.loc[m_i, "Side_Hustle"]))
        for date, typ, cat, desc, amt in seeds:
            vals = [date, short, typ, cat, desc, round(float(amt), 2), "Checking", ""]
            for c, v in enumerate(vals, start=1):
                cell = ws.cell(row=row, column=c, value=v)
                input_cell(cell)
                if c == 6:
                    money_format(cell)
            row += 1

    # Extra blank rows for user entry
    blank_end = row + 80
    for r in range(row, blank_end):
        for c in range(1, 9):
            input_cell(ws.cell(row=r, column=c))
            if c == 6:
                money_format(ws.cell(row=r, column=c))

    type_dv = DataValidation(type="list", formula1='"Income,Expense"', allow_blank=True)
    month_dv = DataValidation(type="list", formula1='"' + ",".join(MONTH_SHORT) + '"', allow_blank=True)
    cat_dv = DataValidation(type="list", formula1='"' + ",".join(INCOME_CATS + EXPENSE_CATS) + '"', allow_blank=True)
    ws.add_data_validation(type_dv)
    ws.add_data_validation(month_dv)
    ws.add_data_validation(cat_dv)
    type_dv.add(f"C5:C{blank_end - 1}")
    month_dv.add(f"B5:B{blank_end - 1}")
    cat_dv.add(f"D5:D{blank_end - 1}")

    ws._tx_last_row = blank_end - 1  # noqa: SLF001 — stash for monthly formulas


def build_monthly_sheet(ws, month_name: str, month_idx: int, sample: pd.DataFrame, tx_last: int):
    short = MONTH_SHORT[month_idx]
    title_block(
        ws,
        f"{month_name.upper()} BUDGET",
        "Budget vs Actual. Recurring items prefill Planned; Actual pulls from Transactions + overrides.",
        "G1",
    )
    set_col_widths(ws, {"A": 24, "B": 14, "C": 14, "D": 14, "E": 14, "F": 18, "G": 18})

    # KPI row
    kpi_box(ws, 4, 1, "Planned Income", f"=B20")
    kpi_box(ws, 4, 2, "Actual Income", f"=C20")
    kpi_box(ws, 4, 3, "Planned Expenses", f"=B42")
    kpi_box(ws, 4, 4, "Actual Expenses", f"=C42")
    kpi_box(ws, 4, 5, "Net Cash Flow", f"=C20-C42")
    kpi_box(ws, 4, 6, "Savings Rate", f'=IF(C20=0,0,(C20-C42)/C20)', is_pct=True)

    # Income section
    section_label(ws, "A7", "Income", COLORS["accent"])
    ws.merge_cells("A7:E7")
    for i, h in enumerate(["Category", "Planned", "Actual", "Variance", "Notes"], start=1):
        ws.cell(row=8, column=i, value=h)
    style_header_row(ws, 8, 1, 5)

    # Planned income from sample / recurring defaults
    planned_income = {
        "Salary": 5500,
        "Side Hustle": float(sample.loc[month_idx, "Side_Hustle"]),
        "Freelance": 0,
        "Investments": 0,
        "Rental Income": 0,
        "Other Income": 0,
    }
    for i, cat in enumerate(INCOME_CATS, start=9):
        ws.cell(row=i, column=1, value=cat).border = THIN
        input_cell(ws.cell(row=i, column=2), round(planned_income[cat], 2))
        money_format(ws.cell(row=i, column=2))
        # Actual = SUMIFS from Transactions
        ws.cell(
            row=i,
            column=3,
            value=f'=SUMIFS(Transactions!$F$5:$F${tx_last},Transactions!$B$5:$B${tx_last},"{short}",Transactions!$C$5:$C${tx_last},"Income",Transactions!$D$5:$D${tx_last},A{i})',
        )
        money_format(ws.cell(row=i, column=3))
        ws.cell(row=i, column=4, value=f"=C{i}-B{i}")
        money_format(ws.cell(row=i, column=4))
        input_cell(ws.cell(row=i, column=5))

    # For salary (often not in transactions), seed actual = planned if no tx
    # Add helper note rows
    label_cell(ws["A15"], "Income Subtotal")
    ws["B15"] = "=SUM(B9:B14)"
    ws["C15"] = "=SUM(C9:C14)"
    ws["D15"] = "=C15-B15"
    for col in ("B", "C", "D"):
        money_format(ws[f"{col}15"])
        ws[f"{col}15"].font = font(11, bold=True)

    # Override tip: if salary not logged in Transactions, enter Actual manually in cream cells — but
    # formulas win. So for sample usefulness, add a "Manual Actual Override" column? Simpler approach:
    # Pre-write Actual for salary as planned amount using a max/fallback pattern.
    # Replace salary actual formula with: if sumifs=0 use planned else sumifs
    ws["C9"] = f'=IF(SUMIFS(Transactions!$F$5:$F${tx_last},Transactions!$B$5:$B${tx_last},"{short}",Transactions!$C$5:$C${tx_last},"Income",Transactions!$D$5:$D${tx_last},A9)=0,B9,SUMIFS(Transactions!$F$5:$F${tx_last},Transactions!$B$5:$B${tx_last},"{short}",Transactions!$C$5:$C${tx_last},"Income",Transactions!$D$5:$D${tx_last},A9))'

    # Hidden bridge totals used by KPIs (row 20)
    ws["A20"] = "Income totals (KPI bridge)"
    ws["A20"].font = font(9, color=COLORS["muted"])
    ws["B20"] = "=B15"
    ws["C20"] = "=C15"
    money_format(ws["B20"])
    money_format(ws["C20"])

    # Expenses
    section_label(ws, "A22", "Expenses", COLORS["warn"])
    ws.merge_cells("A22:E22")
    for i, h in enumerate(["Category", "Planned", "Actual", "Variance", "Notes"], start=1):
        ws.cell(row=23, column=i, value=h)
    style_header_row(ws, 23, 1, 5)

    planned_map = {
        "Housing": sample.loc[month_idx, "Housing"],
        "Utilities": sample.loc[month_idx, "Utilities"],
        "Groceries": sample.loc[month_idx, "Groceries"],
        "Transportation": sample.loc[month_idx, "Transportation"],
        "Insurance": sample.loc[month_idx, "Insurance"],
        "Healthcare": sample.loc[month_idx, "Healthcare"],
        "Subscriptions": sample.loc[month_idx, "Subscriptions"],
        "Dining Out": sample.loc[month_idx, "Dining_Out"],
        "Entertainment": sample.loc[month_idx, "Entertainment"],
        "Personal Care": sample.loc[month_idx, "Personal_Care"],
        "Clothing": sample.loc[month_idx, "Clothing"],
        "Education": sample.loc[month_idx, "Education"],
        "Kids / Family": sample.loc[month_idx, "Kids_Family"],
        "Travel": sample.loc[month_idx, "Travel"],
        "Gifts / Donations": sample.loc[month_idx, "Gifts_Donations"],
        "Debt Payments": sample.loc[month_idx, "Debt_Payments"],
        "Savings Transfers": sample.loc[month_idx, "Savings_Transfers"],
        "Miscellaneous": sample.loc[month_idx, "Miscellaneous"],
    }

    # Categories that are fully covered by transaction seeds vs recurring-only
    tx_covered = {
        "Groceries",
        "Dining Out",
        "Entertainment",
        "Transportation",
        "Miscellaneous",
        "Travel",
    }

    for i, cat in enumerate(EXPENSE_CATS, start=24):
        ws.cell(row=i, column=1, value=cat).border = THIN
        planned = round(float(planned_map[cat]), 2)
        input_cell(ws.cell(row=i, column=2), planned)
        money_format(ws.cell(row=i, column=2))
        if cat in tx_covered:
            formula = (
                f'=SUMIFS(Transactions!$F$5:$F${tx_last},Transactions!$B$5:$B${tx_last},"{short}",'
                f'Transactions!$C$5:$C${tx_last},"Expense",Transactions!$D$5:$D${tx_last},A{i})'
            )
            # If no tx, fall back to planned for demo continuity on partial months — actually leave 0/sum
            ws.cell(row=i, column=3, value=formula)
        else:
            # Recurring-style: default actual = planned (user overrides by typing over... but we need cream)
            # Use formula that equals planned so Annual works, user can hard-edit Planned and we sync:
            input_cell(ws.cell(row=i, column=3), planned)
        money_format(ws.cell(row=i, column=3))
        ws.cell(row=i, column=4, value=f"=B{i}-C{i}")
        money_format(ws.cell(row=i, column=4))
        input_cell(ws.cell(row=i, column=5))

    label_cell(ws["A42"], "Expense Subtotal")
    ws["B42"] = "=SUM(B24:B41)"
    ws["C42"] = "=SUM(C24:C41)"
    ws["D42"] = "=B42-C42"
    for col in ("B", "C", "D"):
        money_format(ws[f"{col}42"])
        ws[f"{col}42"].font = font(11, bold=True)

    # Conditional formatting: overspend (variance negative for expenses means over)
    ws.conditional_formatting.add(
        "D24:D41",
        FormulaRule(formula=["D24<0"], fill=fill(COLORS["bad"])),
    )
    ws.conditional_formatting.add(
        "D24:D41",
        FormulaRule(formula=["D24>=0"], fill=fill(COLORS["good"])),
    )

    ws["A44"] = "Quick tip: Variance for expenses = Planned − Actual (green = under budget)."
    ws["A44"].font = font(10, color=COLORS["muted"])
    ws.merge_cells("A44:G44")


def build_annual_dashboard(ws, sample: pd.DataFrame):
    title_block(
        ws,
        "ANNUAL DASHBOARD",
        "Year-at-a-glance. Monthly totals are linked from each month tab.",
        "N1",
    )
    set_col_widths(ws, {get_column_letter(i): 12 for i in range(1, 16)})
    ws.column_dimensions["A"].width = 22

    # Summary KPIs
    kpi_box(ws, 4, 1, "Annual Income", "=N9")
    kpi_box(ws, 4, 3, "Annual Expenses", "=N10")
    kpi_box(ws, 4, 5, "Annual Savings", "=N11")
    kpi_box(ws, 4, 7, "Avg Savings Rate", "=N12", is_pct=True)
    kpi_box(ws, 4, 9, "vs Savings Goal", "=N11-Setup!B11")

    # Matrix header
    ws["A8"] = "Category"
    for i, m in enumerate(MONTH_SHORT, start=2):
        ws.cell(row=8, column=i, value=m)
    ws.cell(row=8, column=14, value="Annual")
    style_header_row(ws, 8, 1, 14)

    # Income row (actual)
    ws["A9"] = "Total Income"
    for i, month in enumerate(MONTHS, start=2):
        # Reference month sheet C20
        ws.cell(row=9, column=i, value=f"='{month}'!C20")
        money_format(ws.cell(row=9, column=i))
    ws["N9"] = "=SUM(B9:M9)"
    money_format(ws["N9"])
    ws["A9"].font = font(11, bold=True, color=COLORS["accent"])

    ws["A10"] = "Total Expenses"
    for i, month in enumerate(MONTHS, start=2):
        ws.cell(row=10, column=i, value=f"='{month}'!C42")
        money_format(ws.cell(row=10, column=i))
    ws["N10"] = "=SUM(B10:M10)"
    money_format(ws["N10"])
    ws["A10"].font = font(11, bold=True, color=COLORS["warn"])

    ws["A11"] = "Net Savings"
    for i in range(2, 14):
        col = get_column_letter(i)
        ws.cell(row=11, column=i, value=f"={col}9-{col}10")
        money_format(ws.cell(row=11, column=i))
    ws["N11"] = "=N9-N10"
    money_format(ws["N11"])
    ws["A11"].font = font(11, bold=True, color=COLORS["header"])

    ws["A12"] = "Savings Rate"
    for i in range(2, 14):
        col = get_column_letter(i)
        ws.cell(row=12, column=i, value=f"=IF({col}9=0,0,{col}11/{col}9)")
        pct_format(ws.cell(row=12, column=i))
    ws["N12"] = "=IF(N9=0,0,N11/N9)"
    pct_format(ws["N12"])

    # Expense category breakdown annual
    section_label(ws, "A14", "Expense Categories (Actual by Month)", COLORS["header"])
    ws.merge_cells("A14:N14")
    ws["A15"] = "Category"
    for i, m in enumerate(MONTH_SHORT, start=2):
        ws.cell(row=15, column=i, value=m)
    ws["N15"] = "Annual"
    style_header_row(ws, 15, 1, 14)

    for r, cat in enumerate(EXPENSE_CATS, start=16):
        ws.cell(row=r, column=1, value=cat).border = THIN
        src_row = 24 + EXPENSE_CATS.index(cat)
        for i, month in enumerate(MONTHS, start=2):
            ws.cell(row=r, column=i, value=f"='{month}'!C{src_row}")
            money_format(ws.cell(row=r, column=i))
        ws.cell(row=r, column=14, value=f"=SUM(B{r}:M{r})")
        money_format(ws.cell(row=r, column=14))

    ws["A34"] = "Expense Total Check"
    ws["N34"] = "=SUM(N16:N33)"
    money_format(ws["N34"])

    # Charts data already in rows 9-11; create bar + line
    chart1 = BarChart()
    chart1.type = "col"
    chart1.grouping = "clustered"
    chart1.title = "Monthly Income vs Expenses"
    chart1.y_axis.title = "USD"
    chart1.x_axis.title = None
    data = Reference(ws, min_col=1, min_row=9, max_col=13, max_row=10)
    cats = Reference(ws, min_col=2, min_row=8, max_col=13)
    chart1.add_data(data, from_rows=True, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 4
    chart1.width = 18
    chart1.height = 8
    ws.add_chart(chart1, "A36")

    chart2 = LineChart()
    chart2.title = "Monthly Net Savings"
    chart2.y_axis.title = "USD"
    data2 = Reference(ws, min_col=1, min_row=11, max_col=13, max_row=11)
    chart2.add_data(data2, from_rows=True, titles_from_data=True)
    chart2.set_categories(cats)
    chart2.width = 18
    chart2.height = 8
    ws.add_chart(chart2, "H36")

    # Pie of annual expense categories
    pie = PieChart()
    pie.title = "Annual Expense Mix"
    labels = Reference(ws, min_col=1, min_row=16, max_row=33)
    pdata = Reference(ws, min_col=14, min_row=15, max_row=33)
    pie.add_data(pdata, titles_from_data=True)
    pie.set_categories(labels)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = False
    pie.dataLabels.showCatName = False
    pie.width = 14
    pie.height = 10
    ws.add_chart(pie, "A52")

    ws["H52"] = "Replace sample figures anytime — every chart and KPI recalculates automatically."
    ws["H52"].font = font(11, color=COLORS["muted"])


def build_quarterly(ws):
    title_block(ws, "QUARTERLY DASHBOARD", "Seasonal patterns from the Annual Dashboard.", "F1")
    set_col_widths(ws, {"A": 22, "B": 16, "C": 16, "D": 16, "E": 16, "F": 16})

    headers = ["Metric", "Q1 (Jan-Mar)", "Q2 (Apr-Jun)", "Q3 (Jul-Sep)", "Q4 (Oct-Dec)", "Year"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, 1, 6)

    income_q = [
        "='Annual Dashboard'!B9+'Annual Dashboard'!C9+'Annual Dashboard'!D9",
        "='Annual Dashboard'!E9+'Annual Dashboard'!F9+'Annual Dashboard'!G9",
        "='Annual Dashboard'!H9+'Annual Dashboard'!I9+'Annual Dashboard'!J9",
        "='Annual Dashboard'!K9+'Annual Dashboard'!L9+'Annual Dashboard'!M9",
        "='Annual Dashboard'!N9",
    ]
    expense_q = [
        "='Annual Dashboard'!B10+'Annual Dashboard'!C10+'Annual Dashboard'!D10",
        "='Annual Dashboard'!E10+'Annual Dashboard'!F10+'Annual Dashboard'!G10",
        "='Annual Dashboard'!H10+'Annual Dashboard'!I10+'Annual Dashboard'!J10",
        "='Annual Dashboard'!K10+'Annual Dashboard'!L10+'Annual Dashboard'!M10",
        "='Annual Dashboard'!N10",
    ]

    ws["A5"] = "Income"
    for i, f in enumerate(income_q, start=2):
        ws.cell(row=5, column=i, value=f)
        money_format(ws.cell(row=5, column=i))
    ws["A6"] = "Expenses"
    for i, f in enumerate(expense_q, start=2):
        ws.cell(row=6, column=i, value=f)
        money_format(ws.cell(row=6, column=i))
    ws["A7"] = "Net Savings"
    for i in range(2, 7):
        col = get_column_letter(i)
        ws.cell(row=7, column=i, value=f"={col}5-{col}6")
        money_format(ws.cell(row=7, column=i))
    ws["A8"] = "Savings Rate"
    for i in range(2, 7):
        col = get_column_letter(i)
        ws.cell(row=8, column=i, value=f"=IF({col}5=0,0,{col}7/{col}5)")
        pct_format(ws.cell(row=8, column=i))

    for r in range(5, 9):
        ws.cell(row=r, column=1).font = font(11, bold=True)
        ws.cell(row=r, column=1).fill = fill(COLORS["soft"])
        ws.cell(row=r, column=1).border = THIN

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "Quarterly Income vs Expenses"
    data = Reference(ws, min_col=1, min_row=5, max_col=5, max_row=6)
    cats = Reference(ws, min_col=2, min_row=4, max_col=5)
    chart.add_data(data, from_rows=True, titles_from_data=True)
    chart.set_categories(cats)
    chart.width = 15
    chart.height = 8
    ws.add_chart(chart, "A11")


def build_502030(ws):
    title_block(ws, "50 / 30 / 20 DASHBOARD", "Needs · Wants · Savings — targets from Setup.", "G1")
    set_col_widths(ws, {"A": 24, "B": 16, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16})

    ws["A4"] = "Annual Income (Actual)"
    ws["B4"] = "='Annual Dashboard'!N9"
    money_format(ws["B4"])

    label_cell(ws["A6"], "Bucket")
    label_cell(ws["B6"], "Target %")
    label_cell(ws["C6"], "Target $")
    label_cell(ws["D6"], "Actual $")
    label_cell(ws["E6"], "Actual %")
    label_cell(ws["F6"], "Variance $")

    buckets = [("Needs", "Setup!B18"), ("Wants", "Setup!B19"), ("Savings", "Setup!B20")]
    # Actual formulas via SUMIF on annual category totals using Setup mapping
    # Needs = sum of categories mapped to Needs on Setup!E6:E23 against Annual N16:N33
    for i, (name, pct_ref) in enumerate(buckets, start=7):
        ws.cell(row=i, column=1, value=name).border = THIN
        ws.cell(row=i, column=2, value=f"={pct_ref}")
        pct_format(ws.cell(row=i, column=2))
        ws.cell(row=i, column=3, value=f"=$B$4*B{i}")
        money_format(ws.cell(row=i, column=3))
        # SUMPRODUCT for mapped categories
        ws.cell(
            row=i,
            column=4,
            value=(
                f'=SUMPRODUCT((Setup!$E$6:$E$23=A{i})*('
                f"'Annual Dashboard'!$N$16:$N$33))"
            ),
        )
        money_format(ws.cell(row=i, column=4))
        ws.cell(row=i, column=5, value=f"=IF($B$4=0,0,D{i}/$B$4)")
        pct_format(ws.cell(row=i, column=5))
        ws.cell(row=i, column=6, value=f"=C{i}-D{i}")
        money_format(ws.cell(row=i, column=6))

    ws["A11"] = "Interpretation: positive Variance $ means you spent less than the target allocation (good for Needs/Wants; for Savings, negative variance means you under-saved)."
    ws["A11"].font = font(10, color=COLORS["muted"])
    ws.merge_cells("A11:G11")

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "50/30/20 Target vs Actual"
    data = Reference(ws, min_col=3, min_row=6, max_col=4, max_row=9)
    cats = Reference(ws, min_col=1, min_row=7, max_row=9)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.width = 12
    chart.height = 8
    ws.add_chart(chart, "A13")


def build_net_worth(ws):
    title_block(ws, "NET WORTH TRACKER", "Log month-end snapshots. Starting point pulls from Accounts.", "H1")
    set_col_widths(ws, {"A": 14, "B": 16, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16, "H": 20})

    headers = ["Month", "Cash & Bank", "Investments", "Other Assets", "Credit Cards", "Loans", "Net Worth", "MoM Change"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, 1, 8)

    # January seeded from Accounts formulas where possible
    for i, month in enumerate(MONTH_SHORT, start=5):
        ws.cell(row=i, column=1, value=month).border = THIN
        if i == 5:
            input_cell(ws.cell(row=i, column=2), 13000)  # checking+savings approx
            input_cell(ws.cell(row=i, column=3), 57000)  # brokerage+401k
            input_cell(ws.cell(row=i, column=4), 0)
            input_cell(ws.cell(row=i, column=5), 1450)
            input_cell(ws.cell(row=i, column=6), 22200)
        else:
            for c in range(2, 7):
                input_cell(ws.cell(row=i, column=c))
                money_format(ws.cell(row=i, column=c))
        for c in range(2, 7):
            money_format(ws.cell(row=i, column=c))
        ws.cell(row=i, column=7, value=f"=B{i}+C{i}+D{i}-E{i}-F{i}")
        money_format(ws.cell(row=i, column=7))
        if i == 5:
            ws.cell(row=i, column=8, value=0)
        else:
            ws.cell(row=i, column=8, value=f"=G{i}-G{i-1}")
        money_format(ws.cell(row=i, column=8))

    chart = LineChart()
    chart.title = "Net Worth Over the Year"
    data = Reference(ws, min_col=7, min_row=4, max_row=16)
    cats = Reference(ws, min_col=1, min_row=5, max_row=16)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.width = 16
    chart.height = 8
    ws.add_chart(chart, "A19")


def build_debt(ws):
    title_block(
        ws,
        "DEBT TRACKER — Snowball & Avalanche",
        "Enter debts once. Extra payment power accelerates both payoff plans.",
        "J1",
    )
    set_col_widths(
        ws,
        {
            "A": 20,
            "B": 14,
            "C": 12,
            "D": 14,
            "E": 14,
            "F": 14,
            "G": 12,
            "H": 14,
            "I": 14,
            "J": 18,
        },
    )

    label_cell(ws["A4"], "Extra Monthly Payment Power")
    input_cell(ws["B4"], 300)
    money_format(ws["B4"])
    ws["C4"] = "Applied to priority debt after minimums"
    ws["C4"].font = font(10, color=COLORS["muted"])

    headers = [
        "Debt Name",
        "Balance",
        "APR",
        "Min Payment",
        "Snowball Rank",
        "Avalanche Rank",
        "Type",
        "Due Day",
        "Status",
        "Notes",
    ]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=6, column=i, value=h)
    style_header_row(ws, 6, 1, 10)

    debts = [
        ("Visa Credit Card", 1450, 0.219, 45, "Credit Card", 18, "Active", "Pay aggressively"),
        ("Auto Loan", 9800, 0.052, 285, "Loan", 15, "Active", ""),
        ("Student Loan", 12400, 0.048, 115, "Loan", 22, "Active", ""),
        ("Store Card", 620, 0.249, 35, "Credit Card", 9, "Active", "Highest APR"),
    ]
    for r, d in enumerate(debts, start=7):
        name, bal, apr, amin, typ, due, status, notes = d
        input_cell(ws.cell(row=r, column=1), name)
        input_cell(ws.cell(row=r, column=2), bal)
        money_format(ws.cell(row=r, column=2))
        input_cell(ws.cell(row=r, column=3), apr)
        pct_format(ws.cell(row=r, column=3))
        input_cell(ws.cell(row=r, column=4), amin)
        money_format(ws.cell(row=r, column=4))
        # Snowball rank: smallest balance first
        ws.cell(row=r, column=5, value=f'=IF(B{r}="","",RANK(B{r},$B$7:$B$16,1))')
        # Avalanche rank: highest APR first
        ws.cell(row=r, column=6, value=f'=IF(C{r}="","",RANK(C{r},$C$7:$C$16,0))')
        input_cell(ws.cell(row=r, column=7), typ)
        input_cell(ws.cell(row=r, column=8), due)
        input_cell(ws.cell(row=r, column=9), status)
        input_cell(ws.cell(row=r, column=10), notes)

    for r in range(11, 17):
        for c in range(1, 11):
            input_cell(ws.cell(row=r, column=c))
            if c in (2, 4):
                money_format(ws.cell(row=r, column=c))
            if c == 3:
                pct_format(ws.cell(row=r, column=c))
        ws.cell(row=r, column=5, value=f'=IF(B{r}="","",RANK(B{r},$B$7:$B$16,1))')
        ws.cell(row=r, column=6, value=f'=IF(C{r}="","",RANK(C{r},$C$7:$C$16,0))')

    label_cell(ws["A18"], "Total Debt")
    ws["B18"] = "=SUM(B7:B16)"
    money_format(ws["B18"])
    label_cell(ws["A19"], "Total Minimums")
    ws["B19"] = "=SUM(D7:D16)"
    money_format(ws["B19"])
    label_cell(ws["A20"], "Weighted Avg APR")
    ws["B20"] = "=IF(B18=0,0,SUMPRODUCT(B7:B16,C7:C16)/B18)"
    pct_format(ws["B20"])

    section_label(ws, "A22", "How to use the ranks", COLORS["header"])
    ws.merge_cells("A22:J22")
    ws["A23"] = (
        "Snowball (col E): attack the lowest balance first for quick wins. "
        "Avalanche (col F): attack the highest APR first to minimize interest. "
        "Pay all minimums, then pour Extra Payment Power into rank #1."
    )
    ws["A23"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A23:J24")
    ws.row_dimensions[23].height = 36

    # Simple payoff estimate helper table
    section_label(ws, "A26", "Payoff Estimator (simple interest approximation)", COLORS["accent"])
    ws.merge_cells("A26:E26")
    ws["A27"] = "Focus Debt (type name)"
    input_cell(ws["B27"], "Store Card")
    ws["A28"] = "Balance Lookup"
    ws["B28"] = '=IFERROR(VLOOKUP(B27,A7:D16,2,FALSE),0)'
    money_format(ws["B28"])
    ws["A29"] = "APR Lookup"
    ws["B29"] = '=IFERROR(VLOOKUP(B27,A7:D16,3,FALSE),0)'
    pct_format(ws["B29"])
    ws["A30"] = "Min Payment Lookup"
    ws["B30"] = '=IFERROR(VLOOKUP(B27,A7:D16,4,FALSE),0)'
    money_format(ws["B30"])
    ws["A31"] = "Payment With Extra"
    ws["B31"] = "=B30+$B$4"
    money_format(ws["B31"])
    ws["A32"] = "Est. Months to Payoff"
    ws["B32"] = '=IF(OR(B28<=0,B31<=0),"—",ROUND(NPER(B29/12,-B31,B28),1))'
    ws["A33"] = "Est. Interest Cost"
    ws["B33"] = '=IF(OR(B28<=0,B31<=0),0,B31*B32-B28)'
    money_format(ws["B33"])


def build_sinking(ws):
    title_block(ws, "SINKING FUNDS & SAVINGS GOALS", "Give every future expense a jar and a monthly contribution.", "H1")
    set_col_widths(ws, {"A": 24, "B": 14, "C": 14, "D": 16, "E": 14, "F": 16, "G": 14, "H": 20})

    headers = [
        "Goal / Fund",
        "Target",
        "Saved So Far",
        "Remaining",
        "Target Date",
        "Monthly Contribution",
        "% Complete",
        "Notes",
    ]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, 1, 8)

    funds = [
        ("Emergency Fund Top-Up", 12000, 6500, "2026-12-31", "Primary safety net"),
        ("Annual Vacation", 3000, 900, "2026-07-01", "Summer trip"),
        ("Holiday Gifts", 800, 150, "2026-11-15", ""),
        ("Car Maintenance", 1200, 400, "2026-12-31", "Tires + service"),
        ("Home / Apartment Projects", 2000, 250, "2026-10-01", ""),
        ("Insurance Deductible Reserve", 1500, 500, "2026-12-31", ""),
    ]
    for r, f in enumerate(funds, start=5):
        name, target, saved, date, notes = f
        input_cell(ws.cell(row=r, column=1), name)
        input_cell(ws.cell(row=r, column=2), target)
        money_format(ws.cell(row=r, column=2))
        input_cell(ws.cell(row=r, column=3), saved)
        money_format(ws.cell(row=r, column=3))
        ws.cell(row=r, column=4, value=f"=MAX(0,B{r}-C{r})")
        money_format(ws.cell(row=r, column=4))
        input_cell(ws.cell(row=r, column=5), date)
        # Monthly contribution suggestion assuming equal months remaining ~ rough: remaining/12
        ws.cell(row=r, column=6, value=f"=IF(D{r}=0,0,ROUND(D{r}/12,2))")
        money_format(ws.cell(row=r, column=6))
        ws.cell(row=r, column=7, value=f"=IF(B{r}=0,0,C{r}/B{r})")
        pct_format(ws.cell(row=r, column=7))
        input_cell(ws.cell(row=r, column=8), notes)

    for r in range(11, 20):
        for c in range(1, 9):
            input_cell(ws.cell(row=r, column=c))
            if c in (2, 3, 4, 6):
                money_format(ws.cell(row=r, column=c))
        ws.cell(row=r, column=4, value=f'=IF(B{r}="","",MAX(0,B{r}-C{r}))')
        ws.cell(row=r, column=6, value=f'=IF(OR(B{r}="",D{r}=0),0,ROUND(D{r}/12,2))')
        ws.cell(row=r, column=7, value=f'=IF(OR(B{r}="",B{r}=0),0,C{r}/B{r})')
        pct_format(ws.cell(row=r, column=7))

    label_cell(ws["A21"], "Total Targets")
    ws["B21"] = "=SUM(B5:B19)"
    money_format(ws["B21"])
    label_cell(ws["A22"], "Total Saved")
    ws["B22"] = "=SUM(C5:C19)"
    money_format(ws["B22"])
    label_cell(ws["A23"], "Suggested Monthly Total")
    ws["B23"] = "=SUM(F5:F19)"
    money_format(ws["B23"])
    ws["B23"].font = font(12, bold=True, color=COLORS["accent"])


def build_emergency(ws):
    title_block(ws, "EMERGENCY FUND RUNWAY", "How many months of expenses can you survive?", "F1")
    set_col_widths(ws, {"A": 36, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18})

    label_cell(ws["A4"], "Current Emergency Savings")
    input_cell(ws["B4"], 6500)
    money_format(ws["B4"])

    label_cell(ws["A5"], "Avg Monthly Essential Expenses")
    # Approximate essentials from annual needs-ish categories / 12
    ws["B5"] = (
        "=('Annual Dashboard'!N16+'Annual Dashboard'!N17+'Annual Dashboard'!N18+"
        "'Annual Dashboard'!N19+'Annual Dashboard'!N20+'Annual Dashboard'!N21)/12"
    )
    money_format(ws["B5"])

    label_cell(ws["A6"], "Target Months (from Setup)")
    ws["B6"] = "=Setup!B13"

    label_cell(ws["A7"], "Target Balance")
    ws["B7"] = "=B5*B6"
    money_format(ws["B7"])

    label_cell(ws["A8"], "Months of Runway Now")
    ws["B8"] = "=IF(B5=0,0,B4/B5)"
    ws["B8"].number_format = "0.00"
    ws["B8"].font = font(14, bold=True, color=COLORS["header"])

    label_cell(ws["A9"], "Gap to Target")
    ws["B9"] = "=MAX(0,B7-B4)"
    money_format(ws["B9"])

    label_cell(ws["A10"], "Months to Target at $400/mo")
    input_cell(ws["C10"], 400)
    money_format(ws["C10"])
    ws["B10"] = "=IF(C10=0,\"—\",ROUNDUP(B9/C10,0))"

    ws["A12"] = "Essentials used above: Housing, Utilities, Groceries, Transportation, Insurance, Healthcare (annual ÷ 12)."
    ws["A12"].font = font(10, color=COLORS["muted"])
    ws.merge_cells("A12:F12")


def build_wealth(ws):
    title_block(
        ws,
        "5-YEAR WEALTH PROJECTION",
        "Uses Setup nest egg, expected return, and your average monthly surplus from the Annual Dashboard.",
        "F1",
    )
    set_col_widths(ws, {"A": 18, "B": 18, "C": 18, "D": 18, "E": 18, "F": 22})

    label_cell(ws["A4"], "Starting Nest Egg")
    ws["B4"] = "=Setup!B15"
    money_format(ws["B4"])
    label_cell(ws["A5"], "Expected Annual Return")
    ws["B5"] = "=Setup!B14"
    pct_format(ws["B5"])
    label_cell(ws["A6"], "Avg Monthly Surplus")
    ws["B6"] = "='Annual Dashboard'!N11/12"
    money_format(ws["B6"])

    ws["A8"] = "Month #"
    ws["B8"] = "Projected Balance"
    style_header_row(ws, 8, 1, 2)

    # Month 1
    ws["A9"] = 1
    ws["B9"] = "=$B$4*(1+$B$5/12)+$B$6"
    money_format(ws["B9"])
    for r in range(10, 69):  # 60 months
        ws.cell(row=r, column=1, value=r - 8)
        ws.cell(row=r, column=2, value=f"=B{r-1}*(1+$B$5/12)+$B$6")
        money_format(ws.cell(row=r, column=2))

    label_cell(ws["D4"], "Balance at 1 Year")
    ws["E4"] = "=B20"
    money_format(ws["E4"])
    label_cell(ws["D5"], "Balance at 3 Years")
    ws["E5"] = "=B44"
    money_format(ws["E5"])
    label_cell(ws["D6"], "Balance at 5 Years")
    ws["E6"] = "=B68"
    money_format(ws["E6"])
    ws["E6"].font = font(14, bold=True, color=COLORS["accent"])

    chart = LineChart()
    chart.title = "Projected Portfolio Growth"
    data = Reference(ws, min_col=2, min_row=8, max_row=68)
    cats = Reference(ws, min_col=1, min_row=9, max_row=68)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.width = 16
    chart.height = 9
    ws.add_chart(chart, "D8")


def build_health(ws):
    title_block(
        ws,
        "FINANCIAL HEALTH SCORE",
        "Automated KPIs that typical beige planners leave as vibes. Score out of 100.",
        "F1",
    )
    set_col_widths(ws, {"A": 36, "B": 18, "C": 14, "D": 40, "E": 14, "F": 14})

    headers = ["Metric", "Value", "Score (0-25)", "Rubric", "Weight"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, 1, 5)

    # Savings rate
    ws["A5"] = "Average Savings Rate"
    ws["B5"] = "='Annual Dashboard'!N12"
    pct_format(ws["B5"])
    ws["C5"] = '=IF(B5>=0.2,25,IF(B5>=0.15,20,IF(B5>=0.1,14,IF(B5>=0.05,8,2))))'
    ws["D5"] = "≥20% = 25 · ≥15% = 20 · ≥10% = 14 · ≥5% = 8 · else 2"
    ws["E5"] = 25

    # Emergency runway
    ws["A6"] = "Emergency Fund Runway (months)"
    ws["B6"] = "='Emergency Fund'!B8"
    ws["B6"].number_format = "0.00"
    ws["C6"] = '=IF(B6>=6,25,IF(B6>=3,18,IF(B6>=1,10,3)))'
    ws["D6"] = "≥6 mo = 25 · ≥3 = 18 · ≥1 = 10 · else 3"
    ws["E6"] = 25

    # Debt burden: annual debt payments / income
    ws["A7"] = "Debt Payment Burden"
    ws["B7"] = "=IF('Annual Dashboard'!N9=0,0,'Annual Dashboard'!N31/'Annual Dashboard'!N9)"
    pct_format(ws["B7"])
    ws["C7"] = '=IF(B7<=0.1,25,IF(B7<=0.2,18,IF(B7<=0.35,10,3)))'
    ws["D7"] = "≤10% income = 25 · ≤20% = 18 · ≤35% = 10 · else 3"
    ws["E7"] = 25

    # Budget adherence: share of months with positive net
    ws["A8"] = "Months Cash-Flow Positive"
    ws["B8"] = (
        "=('Annual Dashboard'!B11>0)+('Annual Dashboard'!C11>0)+('Annual Dashboard'!D11>0)+"
        "('Annual Dashboard'!E11>0)+('Annual Dashboard'!F11>0)+('Annual Dashboard'!G11>0)+"
        "('Annual Dashboard'!H11>0)+('Annual Dashboard'!I11>0)+('Annual Dashboard'!J11>0)+"
        "('Annual Dashboard'!K11>0)+('Annual Dashboard'!L11>0)+('Annual Dashboard'!M11>0)"
    )
    ws["C8"] = '=IF(B8>=11,25,IF(B8>=9,20,IF(B8>=6,12,5)))'
    ws["D8"] = "≥11 months = 25 · ≥9 = 20 · ≥6 = 12 · else 5"
    ws["E8"] = 25

    for r in range(5, 9):
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = THIN

    label_cell(ws["A10"], "TOTAL HEALTH SCORE")
    ws["B10"] = "=SUM(C5:C8)"
    ws["B10"].font = font(18, bold=True, color=COLORS["white"])
    ws["B10"].fill = fill(COLORS["header"])
    ws["B10"].alignment = Alignment(horizontal="center")
    ws["C10"] = '=IF(B10>=85,"EXCELLENT",IF(B10>=70,"GOOD",IF(B10>=50,"FAIR","ATTENTION")))'
    ws["C10"].font = font(14, bold=True, color=COLORS["accent"])

    ws["A12"] = "Compared to static Etsy sheets, this score updates the moment your monthly actuals change — so you always know if the plan is working."
    ws["A12"].font = font(10, color=COLORS["muted"])
    ws.merge_cells("A12:F12")


def build_bill_calendar(ws):
    title_block(ws, "BILL PAYMENT CALENDAR", "Pulled from Recurring bills — mark Paid each month.", "N1")
    set_col_widths(ws, {get_column_letter(i): 11 for i in range(1, 15)})
    ws.column_dimensions["A"].width = 24

    ws["A4"] = "Bill"
    ws["B4"] = "Due Day"
    ws["C4"] = "Amount"
    for i, m in enumerate(MONTH_SHORT, start=4):
        ws.cell(row=4, column=i, value=m)
    style_header_row(ws, 4, 1, 15)

    # Link first 10 recurring bill rows
    for i in range(5, 15):
        src = 13 + i  # recurring bills start row 18 → for i=5, src=18
        src = i + 13
        ws.cell(row=i, column=1, value=f"=IF(Recurring!A{src}=\"\",\"\",Recurring!A{src})")
        ws.cell(row=i, column=2, value=f"=IF(Recurring!A{src}=\"\",\"\",Recurring!E{src})")
        ws.cell(row=i, column=3, value=f"=IF(Recurring!A{src}=\"\",\"\",Recurring!C{src})")
        money_format(ws.cell(row=i, column=3))
        for c in range(4, 16):
            input_cell(ws.cell(row=i, column=c), "Paid" if i < 15 else "")
            ws.cell(row=i, column=c).alignment = Alignment(horizontal="center")

    # Fix: only mark sample as Paid for demo on first few — actually leave blank with dropdown
    paid_dv = DataValidation(type="list", formula1='"Paid,Due,Skipped"', allow_blank=True)
    ws.add_data_validation(paid_dv)
    paid_dv.add("D5:O20")

    for i in range(5, 15):
        for c in range(4, 16):
            ws.cell(row=i, column=c).value = "Paid" if c <= 5 else ""  # Jan-Feb paid sample

    for i in range(15, 21):
        src = i + 13
        ws.cell(row=i, column=1, value=f"=IF(Recurring!A{src}=\"\",\"\",Recurring!A{src})")
        ws.cell(row=i, column=2, value=f"=IF(Recurring!A{src}=\"\",\"\",Recurring!E{src})")
        ws.cell(row=i, column=3, value=f"=IF(Recurring!A{src}=\"\",\"\",Recurring!C{src})")
        money_format(ws.cell(row=i, column=3))
        for c in range(4, 16):
            input_cell(ws.cell(row=i, column=c))

    ws["A22"] = "Update status cells (Paid / Due / Skipped) throughout the year. Amounts stay synced with the Recurring tab."
    ws["A22"].font = font(10, color=COLORS["muted"])
    ws.merge_cells("A22:N22")


def build_google_sheets_guide(ws):
    title_block(
        ws,
        "GOOGLE SHEETS GUIDE",
        "This .xlsx is Excel-native and opens cleanly in Google Sheets — no separate fragile copy required.",
        "F1",
    )
    set_col_widths(ws, {"A": 28, "B": 70, "C": 18, "D": 18, "E": 18, "F": 18})
    steps = [
        ("1. Upload", "Upload Premium_Household_Budget_Planner.xlsx to Google Drive."),
        ("2. Open", "Right-click → Open with → Google Sheets."),
        ("3. Make a copy", "File → Make a copy so you can edit safely."),
        ("4. Keep formulas", "Do not convert charts to images; leave formulas intact."),
        ("5. Currency", "Change Setup!B6 if you need €, £, etc. (display symbol for your notes)."),
        ("6. Share", "Share the Sheet with your household partner (Editor)."),
    ]
    ws["A4"] = "Open in Google Sheets"
    ws["A4"].font = font(14, bold=True, color=COLORS["header"])
    for i, (t, b) in enumerate(steps, start=5):
        ws.cell(row=i, column=1, value=t).font = font(11, bold=True, color=COLORS["accent"])
        ws.cell(row=i, column=2, value=b)
    ws["A12"] = "Compatibility notes"
    ws["A12"].font = font(14, bold=True, color=COLORS["header"])
    notes = [
        "All core formulas use standard Excel/Sheets functions (SUMIFS, RANK, NPER, SUMPRODUCT, IF).",
        "Charts created in Excel appear in Sheets; you can rebuild any chart from the Annual Dashboard data table.",
        "Dropdown validations transfer; if a list is missing, re-add Data → Data validation.",
        "Best on desktop Sheets or Excel 365. Not designed for Apple Numbers.",
    ]
    for i, n in enumerate(notes, start=13):
        ws.cell(row=i, column=1, value="•").font = font(12, bold=True, color=COLORS["accent"])
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=5)
        ws.cell(row=i, column=2, value=n)


def build_household(ws):
    title_block(
        ws,
        "HOUSEHOLD CASH FLOW",
        "Dual-income household book — Person A, Person B, shared bills, leftover to save/invest.",
        "G1",
    )
    set_col_widths(ws, {"A": 28, "B": 16, "C": 16, "D": 16, "E": 16, "F": 18, "G": 22})

    label_cell(ws["A4"], "Person A")
    ws["B4"] = "=Setup!B8"
    label_cell(ws["A5"], "Person B")
    ws["B5"] = "=Setup!B9"

    section_label(ws, "A7", "Monthly Take-Home Income", COLORS["accent"])
    ws.merge_cells("A7:D7")
    for i, h in enumerate(["Source", "Person A", "Person B", "Combined"], start=1):
        ws.cell(row=8, column=i, value=h)
    style_header_row(ws, 8, 1, 4)

    income_rows = [
        ("Primary Job", 3200, 2800),
        ("Side Income", 400, 250),
        ("Other", 0, 0),
    ]
    for r, (name, a, b) in enumerate(income_rows, start=9):
        input_cell(ws.cell(row=r, column=1), name)
        input_cell(ws.cell(row=r, column=2), a)
        money_format(ws.cell(row=r, column=2))
        input_cell(ws.cell(row=r, column=3), b)
        money_format(ws.cell(row=r, column=3))
        ws.cell(row=r, column=4, value=f"=B{r}+C{r}")
        money_format(ws.cell(row=r, column=4))

    label_cell(ws["A12"], "Total Monthly Income")
    ws["B12"] = "=SUM(B9:B11)"
    ws["C12"] = "=SUM(C9:C11)"
    ws["D12"] = "=SUM(D9:D11)"
    for col in ("B", "C", "D"):
        money_format(ws[f"{col}12"])
        ws[f"{col}12"].font = font(11, bold=True)

    section_label(ws, "A14", "Shared Household Expenses (Monthly)", COLORS["warn"])
    ws.merge_cells("A14:D14")
    for i, h in enumerate(["Expense", "Budget", "Actual", "Variance"], start=1):
        ws.cell(row=15, column=i, value=h)
    style_header_row(ws, 15, 1, 4)

    shared = [
        ("Rent / Mortgage", 1800, 1800),
        ("Utilities", 250, 240),
        ("Groceries", 650, 620),
        ("Insurance", 220, 220),
        ("Transport / Fuel", 300, 280),
        ("Childcare / Family", 0, 0),
        ("Shared Subscriptions", 95, 95),
        ("Debt Minimums (shared)", 400, 400),
    ]
    for r, (name, planned, actual) in enumerate(shared, start=16):
        input_cell(ws.cell(row=r, column=1), name)
        input_cell(ws.cell(row=r, column=2), planned)
        money_format(ws.cell(row=r, column=2))
        input_cell(ws.cell(row=r, column=3), actual)
        money_format(ws.cell(row=r, column=3))
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}")
        money_format(ws.cell(row=r, column=4))

    label_cell(ws["A24"], "Shared Expense Total")
    ws["B24"] = "=SUM(B16:B23)"
    ws["C24"] = "=SUM(C16:C23)"
    ws["D24"] = "=B24-C24"
    for col in ("B", "C", "D"):
        money_format(ws[f"{col}24"])

    section_label(ws, "A26", "Personal Allowances", COLORS["header"])
    ws.merge_cells("A26:D26")
    ws["A27"] = "Person A Fun Money"
    input_cell(ws["B27"], 250)
    money_format(ws["B27"])
    ws["A28"] = "Person B Fun Money"
    input_cell(ws["B28"], 250)
    money_format(ws["B28"])

    label_cell(ws["A30"], "Left to Save / Invest (Combined)")
    ws["B30"] = "=D12-C24-B27-B28"
    money_format(ws["B30"])
    ws["B30"].font = font(14, bold=True, color=COLORS["accent"])
    ws["C30"] = '=IF(D12=0,0,B30/D12)'
    pct_format(ws["C30"])
    ws["D30"] = "Household savings rate this month"
    ws["D30"].font = font(10, color=COLORS["muted"])


def build_subscriptions(ws):
    title_block(
        ws,
        "SUBSCRIPTIONS AUDIT",
        "Find money you forgot you were spending — cancel, downgrade, or keep.",
        "H1",
    )
    set_col_widths(ws, {"A": 24, "B": 16, "C": 12, "D": 12, "E": 14, "F": 14, "G": 14, "H": 22})
    headers = ["Service", "Category", "Amount", "Billing", "Owner", "Still Needed?", "Annual Cost", "Action"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, 1, 8)

    subs = [
        ("Netflix", "Entertainment", 15.49, "Monthly", "Shared", "Yes", "Keep"),
        ("Spotify Duo", "Entertainment", 14.99, "Monthly", "Shared", "Yes", "Keep"),
        ("iCloud+", "Utilities", 2.99, "Monthly", "Person A", "Yes", "Keep"),
        ("Adobe CC", "Work", 54.99, "Monthly", "Person B", "Maybe", "Review"),
        ("Gym", "Health", 50.00, "Monthly", "Person A", "Yes", "Keep"),
        ("News App", "Entertainment", 8.00, "Monthly", "Person B", "No", "Cancel"),
        ("Meal Kit leftover", "Food", 0.00, "Monthly", "Shared", "No", "Cancel"),
        ("Domain / Hosting", "Work", 12.00, "Monthly", "Person B", "Yes", "Keep"),
    ]
    for r, row in enumerate(subs, start=5):
        for c, v in enumerate(row[:6], start=1):
            input_cell(ws.cell(row=r, column=c), v)
        money_format(ws.cell(row=r, column=3))
        ws.cell(row=r, column=7, value=f'=IF(D{r}="Monthly",C{r}*12,IF(D{r}="Annual",C{r},IF(D{r}="Weekly",C{r}*52,C{r}*12)))')
        money_format(ws.cell(row=r, column=7))
        input_cell(ws.cell(row=r, column=8), row[6])

    for r in range(13, 25):
        for c in range(1, 9):
            input_cell(ws.cell(row=r, column=c))
            if c in (3, 7):
                money_format(ws.cell(row=r, column=c))
        ws.cell(row=r, column=7, value=f'=IF(C{r}="","",IF(D{r}="Monthly",C{r}*12,IF(D{r}="Annual",C{r},IF(D{r}="Weekly",C{r}*52,C{r}*12))))')

    label_cell(ws["A26"], "Monthly Subscription Burn")
    ws["B26"] = '=SUMIF(D5:D24,"Monthly",C5:C24)'
    money_format(ws["B26"])
    label_cell(ws["A27"], "Annual Subscription Burn")
    ws["B27"] = "=SUM(G5:G24)"
    money_format(ws["B27"])
    ws["B27"].font = font(12, bold=True, color=COLORS["warn"])
    label_cell(ws["A28"], "Cancel Candidates (Annual)")
    ws["B28"] = '=SUMIF(F5:F24,"No",G5:G24)'
    money_format(ws["B28"])
    ws["B28"].font = font(12, bold=True, color=COLORS["accent"])

    for formula1, cells in [
        ('"Monthly,Annual,Weekly"', "D5:D24"),
        ('"Shared,Person A,Person B"', "E5:E24"),
        ('"Yes,No,Maybe"', "F5:F24"),
        ('"Keep,Cancel,Downgrade,Review"', "H5:H24"),
    ]:
        dv = DataValidation(type="list", formula1=formula1, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(cells)


def build_weekly_spend(ws):
    title_block(
        ws,
        "WEEKLY VARIABLE SPEND",
        "Envelope-style tracker for groceries, dining, transport, and fun money.",
        "G1",
    )
    set_col_widths(ws, {"A": 16, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 18})
    label_cell(ws["A4"], "Active Month")
    input_cell(ws["B4"], "January")
    month_dv = DataValidation(type="list", formula1='"' + ",".join(MONTHS) + '"', allow_blank=False)
    ws.add_data_validation(month_dv)
    month_dv.add("B4")

    headers = ["Week", "Groceries", "Dining Out", "Transport", "Fun / Misc", "Week Total", "Notes"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=6, column=i, value=h)
    style_header_row(ws, 6, 1, 7)

    sample_weeks = [
        ("Week 1", 160, 55, 70, 40, ""),
        ("Week 2", 145, 80, 65, 25, ""),
        ("Week 3", 170, 40, 75, 60, ""),
        ("Week 4", 155, 90, 70, 35, ""),
        ("Week 5 (if any)", 0, 0, 0, 0, ""),
    ]
    for r, row in enumerate(sample_weeks, start=7):
        input_cell(ws.cell(row=r, column=1), row[0])
        for c, v in enumerate(row[1:5], start=2):
            input_cell(ws.cell(row=r, column=c), v)
            money_format(ws.cell(row=r, column=c))
        ws.cell(row=r, column=6, value=f"=SUM(B{r}:E{r})")
        money_format(ws.cell(row=r, column=6))
        input_cell(ws.cell(row=r, column=7), row[5])

    label_cell(ws["A13"], "Month Total")
    ws["B13"] = "=SUM(B7:B11)"
    ws["C13"] = "=SUM(C7:C11)"
    ws["D13"] = "=SUM(D7:D11)"
    ws["E13"] = "=SUM(E7:E11)"
    ws["F13"] = "=SUM(F7:F11)"
    for col in ("B", "C", "D", "E", "F"):
        money_format(ws[f"{col}13"])
        ws[f"{col}13"].font = font(11, bold=True)

    label_cell(ws["A15"], "Weekly Average")
    ws["B15"] = "=IF(COUNTIF(F7:F11,\">0\")=0,0,F13/COUNTIF(F7:F11,\">0\"))"
    money_format(ws["B15"])


def build_paycheck_budget(ws):
    title_block(
        ws,
        "PAYCHECK BUDGET",
        "Budget by paycheck — perfect for bi-weekly or irregular household income.",
        "H1",
    )
    set_col_widths(ws, {"A": 24, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 20})

    label_cell(ws["A4"], "Pay Frequency")
    ws["B4"] = "=Setup!B10"
    label_cell(ws["A5"], "Next Payday")
    input_cell(ws["B5"], "2026-01-09")

    headers = [
        "Paycheck #",
        "Pay Date",
        "Person A Net",
        "Person B Net",
        "Total In",
        "Bills Assigned",
        "Left to Assign",
        "Notes",
    ]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=7, column=i, value=h)
    style_header_row(ws, 7, 1, 8)

    pays = [
        (1, "2026-01-09", 1600, 1400, 1850, "Rent half + utilities"),
        (2, "2026-01-23", 1600, 1400, 1200, "Groceries + debt + fun"),
        (3, "2026-02-06", 1600, 1400, 1850, ""),
        (4, "2026-02-20", 1600, 1400, 1200, ""),
        (5, "2026-03-06", 1600, 1400, 1850, ""),
        (6, "2026-03-20", 1600, 1400, 1200, ""),
    ]
    for r, row in enumerate(pays, start=8):
        input_cell(ws.cell(row=r, column=1), row[0])
        input_cell(ws.cell(row=r, column=2), row[1])
        input_cell(ws.cell(row=r, column=3), row[2])
        money_format(ws.cell(row=r, column=3))
        input_cell(ws.cell(row=r, column=4), row[3])
        money_format(ws.cell(row=r, column=4))
        ws.cell(row=r, column=5, value=f"=C{r}+D{r}")
        money_format(ws.cell(row=r, column=5))
        input_cell(ws.cell(row=r, column=6), row[4])
        money_format(ws.cell(row=r, column=6))
        ws.cell(row=r, column=7, value=f"=E{r}-F{r}")
        money_format(ws.cell(row=r, column=7))
        input_cell(ws.cell(row=r, column=8), row[5])

    for r in range(14, 34):
        for c in range(1, 9):
            input_cell(ws.cell(row=r, column=c))
            if c in (3, 4, 5, 6, 7):
                money_format(ws.cell(row=r, column=c))
        ws.cell(row=r, column=5, value=f'=IF(AND(C{r}="",D{r}=""),"",C{r}+D{r})')
        ws.cell(row=r, column=7, value=f'=IF(E{r}="","",E{r}-F{r})')

    label_cell(ws["A35"], "YTD Paycheck Income")
    ws["B35"] = "=SUM(E8:E33)"
    money_format(ws["B35"])
    label_cell(ws["A36"], "YTD Bills Assigned")
    ws["B36"] = "=SUM(F8:F33)"
    money_format(ws["B36"])
    label_cell(ws["A37"], "Unassigned Buffer")
    ws["B37"] = "=B35-B36"
    money_format(ws["B37"])
    ws["B37"].font = font(12, bold=True, color=COLORS["accent"])


def build_distribution(ws):
    title_block(
        ws,
        "INCOME & EXPENSE DISTRIBUTION",
        "See where household money comes from and where it goes — annual mix.",
        "F1",
    )
    set_col_widths(ws, {"A": 24, "B": 16, "C": 14, "D": 18, "E": 16, "F": 16})

    section_label(ws, "A4", "Annual Expense Mix", COLORS["header"])
    ws.merge_cells("A4:C4")
    ws["A5"] = "Category"
    ws["B5"] = "Annual $"
    ws["C5"] = "% of Expenses"
    style_header_row(ws, 5, 1, 3)

    for i, cat in enumerate(EXPENSE_CATS, start=6):
        src = 16 + i - 6
        ws.cell(row=i, column=1, value=cat).border = THIN
        ws.cell(row=i, column=2, value=f"='Annual Dashboard'!N{src}")
        money_format(ws.cell(row=i, column=2))
        ws.cell(row=i, column=3, value=f"=IF('Annual Dashboard'!N10=0,0,B{i}/'Annual Dashboard'!N10)")
        pct_format(ws.cell(row=i, column=3))

    label_cell(ws["A24"], "Expense Total")
    ws["B24"] = "=SUM(B6:B23)"
    money_format(ws["B24"])

    pie = PieChart()
    pie.title = "Where Expenses Go"
    labels = Reference(ws, min_col=1, min_row=6, max_row=23)
    data = Reference(ws, min_col=2, min_row=5, max_row=23)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.width = 14
    pie.height = 10
    ws.add_chart(pie, "E4")

    section_label(ws, "A26", "Cash Flow Snapshot", COLORS["accent"])
    ws.merge_cells("A26:C26")
    label_cell(ws["A27"], "Annual Income")
    ws["B27"] = "='Annual Dashboard'!N9"
    money_format(ws["B27"])
    label_cell(ws["A28"], "Annual Expenses")
    ws["B28"] = "='Annual Dashboard'!N10"
    money_format(ws["B28"])
    label_cell(ws["A29"], "Annual Savings")
    ws["B29"] = "='Annual Dashboard'!N11"
    money_format(ws["B29"])
    label_cell(ws["A30"], "Savings Rate")
    ws["B30"] = "='Annual Dashboard'!N12"
    pct_format(ws["B30"])


def build_investments(ws):
    title_block(
        ws,
        "INVESTMENTS FORECAST",
        "Track contributions and project portfolio growth by account.",
        "H1",
    )
    set_col_widths(ws, {"A": 22, "B": 14, "C": 14, "D": 14, "E": 14, "F": 16, "G": 14, "H": 20})
    headers = [
        "Account",
        "Balance Now",
        "Monthly Add",
        "Expected Return",
        "Years",
        "Future Value",
        "Owner",
        "Notes",
    ]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, 1, 8)

    inv = [
        ("Brokerage", 15000, 400, 0.075, 5, "Shared", "Index funds"),
        ("401(k) Person A", 28000, 500, 0.07, 5, "Person A", "With match"),
        ("401(k) Person B", 14000, 350, 0.07, 5, "Person B", ""),
        ("Roth IRA", 9000, 200, 0.075, 5, "Person A", ""),
        ("Kids 529 (optional)", 2500, 100, 0.06, 5, "Shared", ""),
    ]
    for r, row in enumerate(inv, start=5):
        input_cell(ws.cell(row=r, column=1), row[0])
        input_cell(ws.cell(row=r, column=2), row[1])
        money_format(ws.cell(row=r, column=2))
        input_cell(ws.cell(row=r, column=3), row[2])
        money_format(ws.cell(row=r, column=3))
        input_cell(ws.cell(row=r, column=4), row[3])
        pct_format(ws.cell(row=r, column=4))
        input_cell(ws.cell(row=r, column=5), row[4])
        # FV of current balance + annuity of monthly adds
        ws.cell(
            row=r,
            column=6,
            value=f'=IF(B{r}="","",FV(D{r}/12,E{r}*12,-C{r},-B{r}))',
        )
        money_format(ws.cell(row=r, column=6))
        input_cell(ws.cell(row=r, column=7), row[5])
        input_cell(ws.cell(row=r, column=8), row[6])

    for r in range(10, 20):
        for c in range(1, 9):
            input_cell(ws.cell(row=r, column=c))
            if c in (2, 3, 6):
                money_format(ws.cell(row=r, column=c))
            if c == 4:
                pct_format(ws.cell(row=r, column=c))
        ws.cell(row=r, column=6, value=f'=IF(OR(B{r}="",E{r}=""),"",FV(D{r}/12,E{r}*12,-C{r},-B{r}))')

    label_cell(ws["A21"], "Total Balance Now")
    ws["B21"] = "=SUM(B5:B19)"
    money_format(ws["B21"])
    label_cell(ws["A22"], "Total Monthly Adds")
    ws["B22"] = "=SUM(C5:C19)"
    money_format(ws["B22"])
    label_cell(ws["A23"], "Projected Future Value")
    ws["B23"] = "=SUM(F5:F19)"
    money_format(ws["B23"])
    ws["B23"].font = font(13, bold=True, color=COLORS["accent"])


def build_no_spend(ws):
    title_block(
        ws,
        "NO-SPEND CHALLENGE",
        "Cut discretionary leaks for 7 / 14 / 30 days — track wins and money saved.",
        "G1",
    )
    set_col_widths(ws, {"A": 22, "B": 14, "C": 14, "D": 16, "E": 16, "F": 16, "G": 28})
    headers = [
        "Challenge Name",
        "Start Date",
        "Days",
        "Categories Blocked",
        "Money Saved",
        "Completed?",
        "Notes",
    ]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, 1, 7)

    challenges = [
        ("January No-Dining", "2026-01-06", 14, "Dining Out", 180, "Yes", "Cooked at home"),
        ("Subscription Freeze", "2026-02-01", 30, "Subscriptions", 45, "In Progress", "Paused one app"),
        ("Coffee Shop Pause", "2026-03-01", 7, "Dining Out", 28, "No", ""),
        ("Impulse Buy Ban", "2026-04-01", 21, "Clothing, Misc", 0, "Planned", ""),
        ("Entertainment Reset", "2026-05-01", 14, "Entertainment", 0, "Planned", ""),
    ]
    for r, row in enumerate(challenges, start=5):
        for c, v in enumerate(row, start=1):
            input_cell(ws.cell(row=r, column=c), v)
            if c == 5:
                money_format(ws.cell(row=r, column=c))

    for r in range(10, 16):
        for c in range(1, 8):
            input_cell(ws.cell(row=r, column=c))
            if c == 5:
                money_format(ws.cell(row=r, column=c))

    status_dv = DataValidation(type="list", formula1='"Yes,No,In Progress,Planned"', allow_blank=True)
    ws.add_data_validation(status_dv)
    status_dv.add("F5:F15")

    label_cell(ws["A17"], "Total Saved from Challenges")
    ws["B17"] = "=SUM(E5:E15)"
    money_format(ws["B17"])
    ws["B17"].font = font(13, bold=True, color=COLORS["accent"])
    label_cell(ws["A18"], "Completed Challenges")
    ws["B18"] = '=COUNTIF(F5:F15,"Yes")'


def style_sheet_tabs(wb: Workbook):
    tab_colors = {
        "Instructions": "3D4F5F",
        "Google Sheets Guide": "3D4F5F",
        "Setup": "6B8F71",
        "Household": "6B8F71",
        "Accounts": "6B8F71",
        "Recurring": "C4A484",
        "Transactions": "C4A484",
        "Subscriptions": "C4A484",
        "Weekly Spend": "C4A484",
        "Paycheck Budget": "C4A484",
        "Annual Dashboard": "3D4F5F",
        "Distribution": "3D4F5F",
        "Quarterly": "3D4F5F",
        "50-30-20": "6B8F71",
        "Net Worth": "6B8F71",
        "Debt Tracker": "C4785A",
        "Sinking Funds": "6B8F71",
        "Emergency Fund": "6B8F71",
        "Investments": "6B8F71",
        "Wealth Projection": "3D4F5F",
        "Health Score": "3D4F5F",
        "No-Spend Challenge": "C4785A",
        "Bill Calendar": "C4A484",
    }
    for name, color in tab_colors.items():
        if name in wb.sheetnames:
            wb[name].sheet_properties.tabColor = color
    for month in MONTHS:
        if month in wb.sheetnames:
            wb[month].sheet_properties.tabColor = "A3B18A"


def paint_background(ws, max_row=80, max_col=16):
    bg = fill(COLORS["bg"])
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            if cell.fill.fgColor is None or cell.fill.fgColor.rgb in ("00000000", None):
                # Don't stomp intentional fills — openpyxl default is often '00000000'
                if cell.fill.fill_type is None:
                    cell.fill = bg


def build_workbook(output_path: Path) -> Path:
    sample = generate_sample_monthly()
    wb = Workbook()

    ws = wb.active
    ws.title = "Instructions"
    build_instructions(ws)

    build_google_sheets_guide(wb.create_sheet("Google Sheets Guide"))
    build_setup(wb.create_sheet("Setup"))
    build_household(wb.create_sheet("Household"))
    build_accounts(wb.create_sheet("Accounts"))
    build_recurring(wb.create_sheet("Recurring"), sample)

    ws_tx = wb.create_sheet("Transactions")
    build_transactions(ws_tx, sample)
    tx_last = getattr(ws_tx, "_tx_last_row", 200)

    build_subscriptions(wb.create_sheet("Subscriptions"))
    build_weekly_spend(wb.create_sheet("Weekly Spend"))
    build_paycheck_budget(wb.create_sheet("Paycheck Budget"))

    for idx, month in enumerate(MONTHS):
        build_monthly_sheet(wb.create_sheet(month), month, idx, sample, tx_last)

    build_annual_dashboard(wb.create_sheet("Annual Dashboard"), sample)
    build_distribution(wb.create_sheet("Distribution"))
    build_quarterly(wb.create_sheet("Quarterly"))
    build_502030(wb.create_sheet("50-30-20"))
    build_net_worth(wb.create_sheet("Net Worth"))
    build_debt(wb.create_sheet("Debt Tracker"))
    build_sinking(wb.create_sheet("Sinking Funds"))
    build_emergency(wb.create_sheet("Emergency Fund"))
    build_investments(wb.create_sheet("Investments"))
    build_wealth(wb.create_sheet("Wealth Projection"))
    build_health(wb.create_sheet("Health Score"))
    build_no_spend(wb.create_sheet("No-Spend Challenge"))
    build_bill_calendar(wb.create_sheet("Bill Calendar"))

    style_sheet_tabs(wb)

    for name in wb.sheetnames:
        if name in MONTHS or name in (
            "Transactions",
            "Annual Dashboard",
            "Bill Calendar",
            "Subscriptions",
            "Paycheck Budget",
        ):
            wb[name].freeze_panes = "A5"
        elif name not in ("Instructions", "Google Sheets Guide"):
            wb[name].freeze_panes = "A3"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def main():
    root = Path(__file__).resolve().parent
    filenames = [
        "Premium_Household_Budget_Planner.xlsx",
        "Ultimate_Annual_Monthly_Budget_Planner.xlsx",
    ]
    primary = root / "output" / filenames[0]
    path = build_workbook(primary)

    artifacts = root / "artifacts"
    artifacts.mkdir(exist_ok=True)
    for name in filenames:
        (root / "output" / name).write_bytes(path.read_bytes())
        (artifacts / name).write_bytes(path.read_bytes())

    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=False)
    print("=" * 60)
    print("PREMIUM HOUSEHOLD BUDGET PLANNER GENERATED")
    print("=" * 60)
    print(f"File: {path}")
    print(f"Also: {artifacts / filenames[0]}")
    print(f"Sheets ({len(wb.sheetnames)}):")
    for name in wb.sheetnames:
        print(f"  - {name}")
    print("=" * 60)


if __name__ == "__main__":
    main()
