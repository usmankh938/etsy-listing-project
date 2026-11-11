"""
Advanced Financial Planner & Analytics Dashboard
------------------------------------------------
A robust Python replacement for static spreadsheet templates featuring 
automated financial analytics, health scoring, and multi-panel charts.
"""

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# Set global plotting style with warm aesthetic tones
plt.style.use("seaborn-v0_8-whitegrid")
plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["text.color"] = "#2C2C2C"
plt.rcParams["axes.labelcolor"] = "#2C2C2C"
plt.rcParams["xtick.color"] = "#2C2C2C"
plt.rcParams["ytick.color"] = "#2C2C2C"

# -------------------------------------------------------------------------
# 1. DATA ENGINE & SIMULATED TRANSACTIONS
# -------------------------------------------------------------------------


def generate_financial_data():
  """Generates comprehensive mock financial transactions and monthly summaries."""
  np.random.seed(42)
  months = [
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

  # Monthly Income (Base salary + side hustles)
  income_data = {
      "Month": months,
      "Salary": [5500] * 12,
      "Side_Hustle": np.random.randint(400, 900, size=12),
  }
  df_income = pd.DataFrame(income_data)
  df_income["Total_Income"] = df_income["Salary"] + df_income["Side_Hustle"]

  # Monthly Expenses across key categories
  expense_data = {
      "Month": months,
      "Housing": [1800] * 12,
      "Groceries": np.random.normal(650, 50, size=12).clip(500, 800),
      "Utilities": np.random.normal(250, 30, size=12).clip(200, 350),
      "Transportation": np.random.normal(300, 40, size=12).clip(220, 400),
      "Discretionary": np.random.normal(600, 90, size=12).clip(450, 850),
      "Debt_Payments": [400] * 12,
  }
  df_expenses = pd.DataFrame(expense_data)
  df_expenses["Total_Expenses"] = df_expenses[
      [
          "Housing",
          "Groceries",
          "Utilities",
          "Transportation",
          "Discretionary",
          "Debt_Payments",
      ]
  ].sum(axis=1)

  # Merge and compute Net Savings
  df_master = pd.merge(df_income, df_expenses, on="Month")
  df_master["Net_Savings"] = (
      df_master["Total_Income"] - df_master["Total_Expenses"]
  )
  df_master["Savings_Rate_Pct"] = (
      df_master["Net_Savings"] / df_master["Total_Income"]
  ) * 100

  return df_master


# -------------------------------------------------------------------------
# 2. FINANCIAL HEALTH & METRICS CALCULATOR
# -------------------------------------------------------------------------


def print_financial_health_report(df):
  """Computes and prints professional financial health KPIs."""
  total_inc = df["Total_Income"].sum()
  total_exp = df["Total_Expenses"].sum()
  total_sav = df["Net_Savings"].sum()
  avg_savings_rate = df["Savings_Rate_Pct"].mean()

  print("=" * 60)
  print("         ULTIMATE FINANCIAL PLANNER - ANNUAL SUMMARY           ")
  print("=" * 60)
  print(f"Total Annual Income    : ${total_inc:,.2f}")
  print(f"Total Annual Expenses  : ${total_exp:,.2f}")
  print(f"Total Annual Savings   : ${total_sav:,.2f}")
  print(f"Average Savings Rate   : {avg_savings_rate:.1f}%")
  print("-" * 60)

  # Automated Health Check
  if avg_savings_rate >= 20:
    print("Health Status: 🌟 EXCELLENT (Exceeds 20% savings benchmark)")
  elif avg_savings_rate >= 10:
    print("Health Status: ✅ GOOD (Meeting baseline savings goals)")
  else:
    print("Health Status: ⚠️ ATTENTION REQUIRED (Savings rate below 10%)")
  print("=" * 60)


# -------------------------------------------------------------------------
# 3. ADVANCED VISUALIZATION DASHBOARD
# -------------------------------------------------------------------------


def render_dashboard(df):
  """Generates a multi-panel publication-ready financial dashboard."""
  fig = plt.figure(figsize=(16, 10), facecolor="#FDFBF7")

  # Define custom warm beige / terracotta / sage color palette
  c_income = "#8A9A86"  # Sage green
  c_expense = "#D9A066"  # Warm terracotta / beige
  c_savings = "#5C768D"  # Muted slate blue
  c_accent = "#4A403A"  # Deep charcoal

  # Grid specification for subplots
  gs = fig.add_gridspec(2, 2, hspace=0.35, wspace=0.25)

  # --- Plot 1: Income vs. Expenses Bar Chart ---
  ax1 = fig.add_subplot(gs[0, 0])
  x = np.arange(len(df["Month"]))
  width = 0.35

  rects1 = ax1.bar(
      x - width / 2,
      df["Total_Income"],
      width,
      label="Income",
      color=c_income,
      edgecolor="none",
      alpha=0.9,
  )
  rects2 = ax1.bar(
      x + width / 2,
      df["Total_Expenses"],
      width,
      label="Expenses",
      color=c_expense,
      edgecolor="none",
      alpha=0.9,
  )

  ax1.set_title(
      "Monthly Cash Flow: Income vs. Expenses",
      fontsize=13,
      fontweight="bold",
      pad=12,
  )
  ax1.set_xticks(x)
  ax1.set_xticklabels(df["Month"])
  ax1.legend(frameon=True, facecolor="#FFFFFF", edgecolor="none")
  ax1.set_ylabel("USD ($)")

  # --- Plot 2: Expense Category Breakdown (Pie Chart) ---
  ax2 = fig.add_subplot(gs[0, 1])
  categories = [
      "Housing",
      "Groceries",
      "Utilities",
      "Transportation",
      "Discretionary",
      "Debt_Payments",
  ]
  category_totals = [df[cat].sum() for cat in categories]
  pie_colors = ["#D9A066", "#E3BC8E", "#C48CB3", "#7FA6C1", "#A3C1AD", "#D68C8C"]

  wedges, texts, autotexts = ax2.pie(
      category_totals,
      labels=categories,
      autopct="%1.1f%%",
      startangle=140,
      colors=pie_colors,
      wedgeprops={"edgecolor": "#FDFBF7", "linewidth": 2},
  )
  for autotext in autotexts:
    autotext.set_color("#FFFFFF")
    autotext.set_fontweight("bold")
    autotext.set_fontsize(9)

  ax2.set_title(
      "Annual Expense Distribution by Category",
      fontsize=13,
      fontweight="bold",
      pad=12,
  )

  # --- Plot 3: Net Savings Trajectory ---
  ax3 = fig.add_subplot(gs[1, 0])
  ax3.plot(
      df["Month"],
      df["Net_Savings"],
      marker="o",
      color=c_savings,
      linewidth=2.5,
      markersize=6,
      label="Net Savings",
  )
  ax3.axhline(
      0, color="#999999", linestyle="--", alpha=0.7, label="Break-Even Line"
  )
  ax3.fill_between(
      df["Month"],
      df["Net_Savings"],
      0,
      color=c_savings,
      alpha=0.15,
  )

  ax3.set_title(
      "Monthly Net Savings Trajectory", fontsize=13, fontweight="bold", pad=12
  )
  ax3.set_ylabel("USD ($)")
  ax3.legend(frameon=True, facecolor="#FFFFFF", edgecolor="none")

  # --- Plot 4: Wealth Projection (Compound Growth Over 5 Years) ---
  ax4 = fig.add_subplot(gs[1, 1])
  avg_monthly_savings = df["Net_Savings"].mean()
  annual_rate_return = 0.075  # 7.5% expected investment return
  months_proj = np.arange(1, 61)

  wealth_accumulated = []
  current_wealth = 15000  # Starting emergency fund / seed capital
  for m in months_proj:
    current_wealth = (
        current_wealth * (1 + annual_rate_return / 12) + avg_monthly_savings
    )
    wealth_accumulated.append(current_wealth)

  ax4.plot(
      months_proj,
      wealth_accumulated,
      color="#4A7065",
      linewidth=2.5,
      label="Projected Portfolio",
  )
  ax4.set_title(
      "5-Year Wealth Growth Projection (7.5% Annual Return)",
      fontsize=13,
      fontweight="bold",
      pad=12,
  )
  ax4.set_xlabel("Months Ahead")
  ax4.set_ylabel("Total Net Worth ($)")
  ax4.legend(frameon=True, facecolor="#FFFFFF", edgecolor="none")

  # Final polish & layout adjustments
  fig.suptitle(
      "EXECUTIVE FINANCIAL DASHBOARD",
      fontsize=18,
      fontweight="bold",
      y=0.96,
      color="#333333",
  )

  plt.savefig(
      "financial_planner_dashboard.png", dpi=300, bbox_inches="tight"
  )
  print("\n[+] Dashboard successfully compiled and saved as 'financial_planner_dashboard.png'")
  plt.show()


# -------------------------------------------------------------------------
# EXECUTION ENTRY POINT
# -------------------------------------------------------------------------
if __name__ == "__main__":
  financial_data = generate_financial_data()
  print_financial_health_report(financial_data)
  render_dashboard(financial_data)