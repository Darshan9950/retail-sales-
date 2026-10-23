"""
Retail Sales Analytics - Python version of the Excel project
=============================================================
Reproduces the Excel work (data cleaning, validation, KPIs, pivot-style
summaries, charts and insights) using pandas + matplotlib.

Usage:
    pip install pandas openpyxl matplotlib
    python retail_sales_analysis.py first_project_retail_sales_dataset.xlsx

Outputs (in ./output):
    charts/*.png            - one chart per analysis
    summary_tables.xlsx     - every pivot-style table, one sheet each
    insights.txt            - auto-generated business insights
"""

import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")  # no display needed
import matplotlib.pyplot as plt
import pandas as pd

# --------------------------------------------------------------------------
# 0. Settings
# --------------------------------------------------------------------------
INPUT_FILE = sys.argv[1] if len(sys.argv) > 1 else "first_project_retail_sales_dataset.xlsx"
SHEET = "clean_data"
OUT = Path("output")
CHARTS = OUT / "charts"
CHARTS.mkdir(parents=True, exist_ok=True)

CORE_COLS = [
    "transaction_id", "transaction_date", "customer_id", "customer_gender",
    "customer_age_group", "customer_segment", "product_id", "product_name",
    "category", "brand", "quantity", "unit_price", "discount_pct",
    "sales_amount", "payment_method", "sales_channel", "region",
]

plt.rcParams.update({"figure.dpi": 110, "axes.spines.top": False,
                     "axes.spines.right": False, "font.size": 10})


def inr(x):
    """Format a number as Indian-style rupees (lakh/crore grouping)."""
    s = f"{x:,.2f}"
    whole, dec = s.split(".")
    whole = whole.replace(",", "")
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        parts = []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        whole = ",".join(parts + [tail])
    return f"₹{whole}.{dec}"


# --------------------------------------------------------------------------
# 1. Load + clean
# --------------------------------------------------------------------------
print("1. Loading data ...")
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET)

# The sheet has stray helper columns (count, 'Unnamed: 18', 'basic statistical', 0)
df = df[CORE_COLS].copy()

df["transaction_date"] = pd.to_datetime(df["transaction_date"])
for c in ["customer_gender", "customer_age_group", "customer_segment", "category",
          "brand", "payment_method", "sales_channel", "region", "product_name"]:
    df[c] = df[c].astype(str).str.strip()

before = len(df)
df = df.drop_duplicates(subset="transaction_id")
df = df.dropna(subset=["transaction_id", "transaction_date", "sales_amount"])
print(f"   rows: {before:,} -> {len(df):,} after removing duplicates / blanks")

# --------------------------------------------------------------------------
# 2. Validation
# --------------------------------------------------------------------------
print("2. Validating ...")
expected = (df["quantity"] * df["unit_price"] * (1 - df["discount_pct"] / 100)).round(2)
mismatch = ((expected - df["sales_amount"]).abs() > 0.05).sum()
print(f"   sales_amount != qty * price * (1-discount): {mismatch} rows")
print(f"   missing values: {int(df.isna().sum().sum())}")
print(f"   date range: {df.transaction_date.min().date()} to {df.transaction_date.max().date()}")

df["year"] = df["transaction_date"].dt.year
df["month"] = df["transaction_date"].dt.to_period("M").dt.to_timestamp()
df["month_name"] = df["transaction_date"].dt.month_name()

# --------------------------------------------------------------------------
# 3. KPIs  (matches the 'Analysis' sheet)
# --------------------------------------------------------------------------
print("3. KPIs")
kpis = pd.Series({
    "Total Sales": df.sales_amount.sum(),
    "Total Transactions": len(df),
    "Unique Customers": df.customer_id.nunique(),
    "Average Order Value": df.sales_amount.mean(),
    "Average Discount (%)": df.discount_pct.mean(),
    "Total Quantity Sold": df.quantity.sum(),
    "Avg Items per Transaction": df.quantity.mean(),
})
print(kpis.round(2).to_string())

# --------------------------------------------------------------------------
# 4. Pivot-style summaries
# --------------------------------------------------------------------------
print("4. Building summary tables ...")


def summary(by):
    t = df.groupby(by).agg(
        total_sales=("sales_amount", "sum"),
        transactions=("transaction_id", "count"),
        avg_sales=("sales_amount", "mean"),
        quantity=("quantity", "sum"),
        avg_unit_price=("unit_price", "mean"),
    ).sort_values("total_sales", ascending=False)
    t["share_%"] = t.total_sales / t.total_sales.sum() * 100
    return t.round(2)


tables = {
    "Channel": summary("sales_channel"),
    "Category": summary("category"),
    "Region": summary("region"),
    "Segment": summary("customer_segment"),
    "Payment": summary("payment_method"),
    "Age Group": summary("customer_age_group").sort_index(),
    "Gender": summary("customer_gender"),
    "Brand": summary("brand"),
    "Discount": summary("discount_pct").sort_index(),
    "Year": summary("year").sort_index(),
    "Top Products": summary("product_name").head(10),
}
monthly = df.groupby("month").agg(total_sales=("sales_amount", "sum"),
                                  transactions=("transaction_id", "count")).round(2)
tables["Monthly Trend"] = monthly
tables["Category x Region"] = df.pivot_table(index="category", columns="region",
                                             values="sales_amount", aggfunc="sum").round(2)
tables["Channel x Category"] = df.pivot_table(index="category", columns="sales_channel",
                                              values="sales_amount", aggfunc="sum").round(2)

with pd.ExcelWriter(OUT / "summary_tables.xlsx") as xw:
    kpis.round(2).to_frame("Value").to_excel(xw, sheet_name="KPIs")
    for name, t in tables.items():
        t.to_excel(xw, sheet_name=name[:31])

# --------------------------------------------------------------------------
# 5. Charts
# --------------------------------------------------------------------------
print("5. Drawing charts ...")
BLUE, ORANGE = "#2F5D8C", "#E08E45"


def bar(table, col, title, fname, horizontal=False, color=BLUE, ylabel="Sales (₹ lakh)"):
    s = table[col] / 1e5 if col == "total_sales" else table[col]
    fig, ax = plt.subplots(figsize=(7, 4))
    if horizontal:
        s.sort_values().plot.barh(ax=ax, color=color)
        ax.set_xlabel(ylabel)
    else:
        s.plot.bar(ax=ax, color=color, rot=0)
        ax.set_ylabel(ylabel)
    ax.set_title(title, fontweight="bold")
    ax.set_ylabel(ax.get_ylabel()) if not horizontal else ax.set_ylabel("")
    ax.set_xlabel("") if not horizontal else None
    fig.tight_layout()
    fig.savefig(CHARTS / fname)
    plt.close(fig)


bar(tables["Channel"], "total_sales", "Sales by Channel", "01_sales_by_channel.png")
bar(tables["Channel"], "transactions", "Transactions by Channel",
    "02_transactions_by_channel.png", ylabel="Transactions", color=ORANGE)
bar(tables["Category"], "total_sales", "Sales by Category", "03_sales_by_category.png", horizontal=True)
bar(tables["Region"], "total_sales", "Sales by Region", "04_sales_by_region.png")
bar(tables["Segment"], "total_sales", "Sales by Customer Segment", "05_sales_by_segment.png")
bar(tables["Payment"], "total_sales", "Sales by Payment Method", "06_sales_by_payment.png")
bar(tables["Age Group"], "total_sales", "Sales by Age Group", "07_sales_by_age.png")
bar(tables["Discount"], "total_sales", "Sales by Discount Level (%)", "08_sales_by_discount.png")
bar(tables["Year"], "total_sales", "Sales by Year", "09_sales_by_year.png")

# Quantity vs price per category (insight: volume != revenue)
fig, ax = plt.subplots(figsize=(7, 4))
cat = tables["Category"]
ax.scatter(cat.quantity, cat.avg_unit_price, s=cat.total_sales / 4e4, color=BLUE, alpha=0.7)
for name, r in cat.iterrows():
    ax.annotate(name, (r.quantity, r.avg_unit_price), xytext=(5, 5), textcoords="offset points")
ax.set_xlabel("Units sold")
ax.set_ylabel("Average unit price (₹)")
ax.set_title("Quantity vs Price by Category (bubble = sales)", fontweight="bold")
fig.tight_layout()
fig.savefig(CHARTS / "10_quantity_vs_price.png")
plt.close(fig)

# Monthly trend
fig, ax = plt.subplots(figsize=(9, 4))
(monthly.total_sales / 1e5).plot(ax=ax, color=BLUE, marker="o")
ax.set_ylabel("Sales (₹ lakh)")
ax.set_xlabel("")
ax.set_title("Monthly Sales Trend", fontweight="bold")
fig.tight_layout()
fig.savefig(CHARTS / "11_monthly_trend.png")
plt.close(fig)

# Category x Region heatmap
fig, ax = plt.subplots(figsize=(7, 4.5))
hm = tables["Category x Region"] / 1e5
im = ax.imshow(hm.values, cmap="Blues", aspect="auto")
ax.set_xticks(range(len(hm.columns)), hm.columns)
ax.set_yticks(range(len(hm.index)), hm.index)
for i in range(hm.shape[0]):
    for j in range(hm.shape[1]):
        ax.text(j, i, f"{hm.iat[i, j]:.1f}", ha="center", va="center", fontsize=8)
fig.colorbar(im, label="₹ lakh")
ax.set_title("Sales Heatmap: Category x Region", fontweight="bold")
fig.tight_layout()
fig.savefig(CHARTS / "12_heatmap_category_region.png")
plt.close(fig)

# --------------------------------------------------------------------------
# 6. Insights (generated from the data, not hard-coded)
# --------------------------------------------------------------------------
print("6. Writing insights ...")
ch, ca, rg = tables["Channel"], tables["Category"], tables["Region"]
yr, dc = tables["Year"], tables["Discount"]
sg, pm = tables["Segment"], tables["Payment"]
y0, y1 = yr.index.min(), yr.index.max()
chg = (yr.loc[y1, "total_sales"] / yr.loc[y0, "total_sales"] - 1) * 100

lines = [
    "FINAL BUSINESS INSIGHTS",
    "=" * 40,
    f"1. Overall: total sales {inr(kpis['Total Sales'])} from {int(kpis['Total Transactions']):,} "
    f"transactions and {int(kpis['Unique Customers']):,} customers; AOV {inr(kpis['Average Order Value'])}.",
    f"2. Channel: {ch.total_sales.idxmax()} has the highest sales; "
    f"{ch.transactions.idxmax()} has the most transactions "
    f"({int(ch.transactions.max()):,}); {ch.avg_sales.idxmin()} has the lowest average order value "
    f"({inr(ch.avg_sales.min())}) -> opportunity for bundles / cross-selling.",
    f"3. Category: {ca.total_sales.idxmax()} earns the most revenue, {ca.total_sales.idxmin()} the least. "
    f"{ca.quantity.idxmax()} sells the most units ({int(ca.quantity.max()):,}) at an average price of "
    f"{inr(ca.loc[ca.quantity.idxmax(), 'avg_unit_price'])} -> volume != revenue.",
    f"4. Region: {rg.total_sales.idxmax()} is best ({inr(rg.total_sales.max())}), "
    f"{rg.total_sales.idxmin()} is lowest ({inr(rg.total_sales.min())}); the gap is only "
    f"{(rg.total_sales.max() / rg.total_sales.min() - 1) * 100:.1f}%.",
    f"5. Trend: {y0} = {inr(yr.loc[y0, 'total_sales'])}, {y1} = {inr(yr.loc[y1, 'total_sales'])} "
    f"({chg:+.2f}%).",
    f"6. Discounts: {int(dc.total_sales.idxmax())}% discount level gives the highest sales "
    f"({inr(dc.total_sales.max())}); {int(dc.total_sales.idxmin())}% the lowest "
    f"({inr(dc.total_sales.min())}). Deeper discounts do not automatically lift sales.",
    f"7. Segments: {sg.total_sales.idxmax()} is the top segment ({inr(sg.total_sales.max())}); "
    f"VIP = {inr(sg.loc['VIP', 'total_sales'])}.",
    f"8. Payments: {pm.total_sales.idxmax()} is the most used/valuable method; "
    f"{pm.total_sales.idxmin()} is the lowest.",
]
text = "\n".join(lines)
(OUT / "insights.txt").write_text(text, encoding="utf-8")
print("\n" + text)
print(f"\nDone. See the '{OUT}' folder.")
