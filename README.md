# Retail Sales Analytics

An end-to-end analysis of 1,20,000 retail transactions (Jan 2024 – Dec 2025), originally built in Microsoft Excel (Pivot Tables, KPIs, dashboard) and reproduced here in Python with pandas and matplotlib.

## Project Overview

The goal is to turn raw transaction data into business insights: which channels, categories, regions, customer segments and discount levels drive sales, and where the opportunities are.

**Skills shown:** data cleaning, validation, KPI calculation, pivot-style aggregation, data visualisation, insight writing.

## Dataset

File: `first_project_retail_sales_dataset.xlsx` (sheet `clean_data`)

| Column | Description |
|---|---|
| transaction_id | Unique transaction ID |
| transaction_date | Date of purchase |
| customer_id, customer_gender, customer_age_group, customer_segment | Customer details (segments: New, Returning, Loyal, VIP) |
| product_id, product_name, category, brand | Product details (8 categories, 3 brands) |
| quantity, unit_price, discount_pct | Order details |
| sales_amount | quantity × unit_price × (1 − discount) |
| payment_method, sales_channel, region | How and where the sale happened |

The workbook also has `Analysis`, `pivot_table`, `Dashboard`, `insights` and `conclusion` sheets from the Excel version of the project.

## Project Structure

```
.
├── retail_sales_analysis.py                  # main script
├── requirements.txt                          # dependencies
├── first_project_retail_sales_dataset.xlsx   # input data (place here)
├── README.md
└── output/                                   # created when the script runs
    ├── charts/                               # 12 PNG charts
    ├── summary_tables.xlsx                   # all pivot-style tables
    └── insights.txt                          # auto-generated insights
```

## How to Run

```bash
pip install -r requirements.txt
python retail_sales_analysis.py first_project_retail_sales_dataset.xlsx
```

Requires Python 3.9 or newer. If you don't pass a file name, the script looks for `first_project_retail_sales_dataset.xlsx` in the current folder.

## What the Script Does

1. **Load and clean** – reads `clean_data`, removes helper columns, trims text, drops duplicates and blank rows.
2. **Validate** – checks missing values, the date range, and that `sales_amount` matches quantity × price × (1 − discount).
3. **KPIs** – total sales, transactions, unique customers, average order value, average discount, total quantity, items per transaction.
4. **Summaries** – sales, transactions, average sale and share by channel, category, region, segment, payment method, age group, gender, brand, discount level and year, plus a monthly trend and cross-tabs.
5. **Charts** – 12 charts saved to `output/charts/`, including a monthly trend and a category × region heatmap.
6. **Insights** – written to `output/insights.txt` from the data, not hard-coded.

## Key Results

| KPI | Value |
|---|---|
| Total Sales | ₹4,53,57,054 (₹4.54 Cr) |
| Transactions | 1,20,000 |
| Unique Customers | 24,784 |
| Average Order Value | ₹377.98 |
| Average Discount | 5.50% |
| Total Quantity Sold | 1,99,549 |

**Findings**

- **In-Store** has the highest sales. **Mobile App** has the most transactions (40,204) but the lowest average order value (₹374.74), so it is an opportunity for bundles and cross-selling.
- **Beauty** earns the most revenue and **Clothing** the least. Clothing sells the most units (25,223) at the lowest average price (₹194.05), so volume does not equal revenue.
- **Central** is the top region (₹92.4 lakh) and **East** the lowest (₹89.7 lakh); the gap is only about 3%.
- Sales fell slightly from 2024 (₹2.28 Cr) to 2025 (₹2.26 Cr), about −0.95%.
- The **0% discount** level gives the highest sales; deeper discounts did not lift total sales.
- **Loyal** is the top customer segment (₹1.14 Cr), with **VIP** close behind (₹1.13 Cr).
- **Cash** is the most used payment method and **Debit Card** the lowest.

## Tools Used

Python, pandas, matplotlib, openpyxl, Microsoft Excel (original version)

## Author

**Darshan** – B.Tech CSE student, aspiring Data Analyst
