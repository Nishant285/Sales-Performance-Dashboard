# Sales Performance Dashboard

**Tools:** SQL (SQLite) · Excel (openpyxl, live formulas) · Power BI (build guide included) · 2,823 order lines, 307 orders, 2003–2005

**Dashboard:** 
A Power BI version of this dashboard is also included in powerbi/Sales-Dashboard.pbix, with DAX measures and a department slicer, alongside the original interactive web dashboard above.

## Business Question

Which regions, product lines, and customers actually drive profit — not just
sales — and where is risk concentrated (large deals, specific product lines)?

## 1. Data

A classic B2B distributor dataset: scale-model vehicles (motorcycles, classic
cars, planes, ships, trains) sold to retailers across 19 countries, 4
territories, 2003–2005.

### A data bug worth knowing about
Loading this CSV with default settings **silently destroys the region data**:
pandas treats the literal string `"NA"` — which this dataset uses as the code
for the **North America** territory — as a missing value. A naive
`pd.read_csv()` wipes out 1,074 of 2,823 rows (38%) of territory data before
any analysis even starts. Fixed in `scripts/01_clean_data.py` with
`keep_default_na=False`. This is the kind of bug that produces a dashboard
that *looks* fine (no errors, charts render) while being quietly wrong —
worth mentioning in an interview if asked about a time you caught a data
quality issue.

### Profit estimation
This dataset has no cost/profit column — only `PRICEEACH` (actual sale
price) and `MSRP` (list price). Profit is estimated as
`Sales − (MSRP × 0.70 × Quantity)`, using 70% of MSRP as an assumed unit
cost. This is a standard distributor-margin assumption, documented here and
in the cleaning script rather than presented as exact.

## 2. SQL Analysis

Full queries in [`sql/analysis_queries.sql`](sql/analysis_queries.sql) — CTEs
and window functions (`RANK`, `SUM OVER`, `LAG`, conditional aggregation)
answering 7 questions.

### Key findings

**Region-wise:** EMEA leads on both sales ($4.98M) and profit ($1.50M, 49%
of company profit), North America is a close second (39% of profit). Profit
margin is nearly flat across all four regions (30.0–30.5%) — the real
differences between regions are in **volume**, not profitability.

**Product line:** Classic Cars drives the most revenue ($3.92M) but has the
**lowest margin of any line (27.1%)**, below the 32.2% company average.
Trains are the opposite — smallest volume ($226K) but **highest margin
(39.0%)**. This is the same "high sales, low margin vs. low sales, high
margin" pattern as the Superstore project, in a completely different
industry.

**Customer concentration:** the top 2 customers (Euro Shopping Channel,
Mini Gifts Distributors) alone account for **15.6% of all revenue** across
307 orders — a tighter concentration than the Superstore dataset's top
decile.

**Large deals carry more risk:** Large deal-size orders have a **3.18%
dispute rate**, roughly 10× the dispute rate of Medium (0.36%) and Small
(0.31%) deals. Worth flagging to account management — bigger deals aren't
just bigger, they're measurably riskier to fulfill cleanly.

**Seasonality:** sales spike sharply every Q4 (October in particular) across
all years in the dataset — consistent with gift-buying seasonality for a
collectibles/model distributor.

## 3. Excel Workbook

[`excel/Sales_Performance_Dashboard.xlsx`](excel/Sales_Performance_Dashboard.xlsx) —
built with `openpyxl`, every number is a **live formula** (`SUMIFS`,
`COUNTIFS`, `SUMPRODUCT`), not a pasted value, so it recalculates if the
underlying data changes. Verified against the SQL results with zero
discrepancies and zero formula errors (LibreOffice recalc check).

**One bug worth knowing about, caught before shipping:** since each order
spans multiple product lines (one row per line item), a naive
`COUNTIFS`/`COUNTA` on Order Number counts *line items* (2,823), not actual
*orders* (307) — off by over 9×. Fixed using a `SUMPRODUCT`/`COUNTIFS`
distinct-count formula pattern everywhere an order count appears (KPI
Summary, Region Analysis, Top Customers). This is the same class of bug as
the territory "NA" issue above — passes silently, looks plausible, and is
simply wrong until checked against a second source (SQL, in this case).

**Sheets:**
- **Raw Data** — full 2,823-row dataset as an Excel Table
- **KPI Summary** — 8 headline metrics (orders, sales, profit, margin, AOV, customers, territories, product lines)
- **Region Analysis** — territory breakdown + bar chart
- **Product Line Analysis** — margin-vs-company-average comparison + chart
- **Monthly Trend** — sales & profit over time + line chart
- **Top Customers** — top 10 by revenue, live-formula totals

## 4. Power BI

This project is built to carry straight into Power BI: import
`excel/Sales_Performance_Dashboard.xlsx`'s **Raw Data** table (or connect to
`sales_project/sales.db` directly), then recreate the same KPIs as DAX
measures:

```dax
Total Sales = SUM('Raw Data'[Sales])
Total Profit = SUM('Raw Data'[Est. Profit])
Overall Margin = DIVIDE([Total Profit], [Total Sales])
Distinct Orders = DISTINCTCOUNT('Raw Data'[Order Number])
```

Note `Distinct Orders` — Power BI's `DISTINCTCOUNT()` handles the
line-items-vs-orders distinction natively, unlike Excel, which needed the
manual `SUMPRODUCT` workaround above. Worth mentioning if asked why the same
metric needed different approaches in each tool.

Recreate the Region Analysis, Product Line, and Monthly Trend charts as
clustered bar / line visuals with `Territory`, `Product Line`, and
`Order Month` on the axis respectively, and add a Territory slicer for
interactivity — same pattern as the HR Analytics dashboard build.

## 5. Recommendations

1. **Investigate Classic Cars' margin specifically**, not just its sales —
   it's the single largest revenue line and also the weakest performer on
   profitability. Even a small margin improvement here has outsized dollar
   impact given the volume.
2. **Protect the top 2 customers** (15.6% of revenue combined) with
   dedicated account management — concentration risk is real at this level.
3. **Add review/approval steps for Large deals** given the 10x higher
   dispute rate — this is a process signal, not a sales-volume problem.
4. **Plan inventory and staffing around the consistent Q4 spike.**
5. **Promote Trains and Vintage Cars** — both carry above-average margins
   but comparatively low volume; a targeted push here is higher-leverage
   than growing already-high-volume, lower-margin lines.

## Project Structure

```
sales_project/
├── data/
│   ├── sales_raw.csv
│   └── sales_clean.csv
├── sql/
│   ├── analysis_queries.sql
│   └── results/
├── excel/
│   └── Sales_Performance_Dashboard.xlsx
├── scripts/
│   ├── 01_clean_data.py
│   ├── 02_load_to_sql.py
│   ├── 03_run_sql_analysis.py
│   └── 04_build_excel.py
├── sales.db
└── README.md
```

---
*Dataset: classic "Sample Sales Data" scale-model distributor dataset
(public, educational use).*
