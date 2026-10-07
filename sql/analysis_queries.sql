-- ============================================================
-- Sales Performance Analysis — SQL Queries
-- Table: sales (classic scale-model distributor data, 2003-2005)
-- ============================================================

-- ------------------------------------------------------------
-- Q1. Region-wise (Territory) performance, ranked by profit
--     (window function: RANK, SUM OVER for company-wide share)
-- ------------------------------------------------------------
SELECT
    Territory,
    COUNT(DISTINCT "Order Number")                          AS orders,
    ROUND(SUM(Sales), 2)                                     AS total_sales,
    ROUND(SUM("Est. Profit"), 2)                             AS total_profit,
    ROUND(SUM("Est. Profit") * 100.0 / SUM(Sales), 2)        AS profit_margin_pct,
    RANK() OVER (ORDER BY SUM("Est. Profit") DESC)           AS profit_rank,
    ROUND(SUM("Est. Profit") * 100.0
          / SUM(SUM("Est. Profit")) OVER (), 2)              AS pct_of_total_profit
FROM sales
GROUP BY Territory
ORDER BY profit_rank;


-- ------------------------------------------------------------
-- Q2. Monthly sales & profit trend with month-over-month growth
--     (CTE + window function: LAG)
-- ------------------------------------------------------------
WITH monthly AS (
    SELECT
        "Order Month"             AS month,
        SUM(Sales)                 AS total_sales,
        SUM("Est. Profit")         AS total_profit
    FROM sales
    GROUP BY "Order Month"
)
SELECT
    month,
    ROUND(total_sales, 2)                                              AS total_sales,
    ROUND(total_profit, 2)                                             AS total_profit,
    ROUND((total_sales - LAG(total_sales) OVER (ORDER BY month)) * 100.0
          / LAG(total_sales) OVER (ORDER BY month), 2)                 AS sales_growth_pct
FROM monthly
ORDER BY month;


-- ------------------------------------------------------------
-- Q3. Top 10 customers by total sales, with their rank and
--     share of company revenue (window functions: RANK, SUM OVER)
-- ------------------------------------------------------------
WITH customer_rev AS (
    SELECT
        "Customer Name"                AS customer,
        Territory,
        ROUND(SUM(Sales), 2)           AS total_sales,
        COUNT(DISTINCT "Order Number") AS orders
    FROM sales
    GROUP BY "Customer Name", Territory
)
SELECT
    customer,
    Territory,
    total_sales,
    orders,
    RANK() OVER (ORDER BY total_sales DESC)                  AS sales_rank,
    ROUND(total_sales * 100.0 / SUM(total_sales) OVER (), 2) AS pct_of_total_revenue
FROM customer_rev
ORDER BY sales_rank
LIMIT 10;


-- ------------------------------------------------------------
-- Q4. Product line performance — sales, profit margin, and
--     how each line's margin compares to the company average
--     (window function: AVG OVER)
-- ------------------------------------------------------------
SELECT
    "Product Line"                                            AS product_line,
    ROUND(SUM(Sales), 2)                                      AS total_sales,
    ROUND(SUM("Est. Profit"), 2)                              AS total_profit,
    ROUND(SUM("Est. Profit") * 100.0 / SUM(Sales), 2)         AS profit_margin_pct,
    ROUND(AVG(SUM("Est. Profit") * 100.0 / SUM(Sales))
          OVER (), 2)                                         AS company_avg_margin_pct
FROM sales
GROUP BY "Product Line"
ORDER BY total_sales DESC;


-- ------------------------------------------------------------
-- Q5. Deal Size vs. order fulfillment — are larger deals more
--     likely to be cancelled/disputed?
-- ------------------------------------------------------------
SELECT
    "Deal Size"                                               AS deal_size,
    Status,
    COUNT(*)                                                  AS order_count,
    ROUND(COUNT(*) * 100.0
          / SUM(COUNT(*)) OVER (PARTITION BY "Deal Size"), 2) AS pct_within_deal_size
FROM sales
GROUP BY "Deal Size", Status
ORDER BY deal_size,
    CASE Status WHEN 'Shipped' THEN 1 WHEN 'Resolved' THEN 2 WHEN 'In Process' THEN 3
                WHEN 'On Hold' THEN 4 WHEN 'Disputed' THEN 5 ELSE 6 END;


-- ------------------------------------------------------------
-- Q6. Year-over-year performance by territory (CTE + pivot-style
--     comparison using conditional aggregation)
-- ------------------------------------------------------------
SELECT
    Territory,
    ROUND(SUM(CASE WHEN Year = 2003 THEN Sales ELSE 0 END), 2) AS sales_2003,
    ROUND(SUM(CASE WHEN Year = 2004 THEN Sales ELSE 0 END), 2) AS sales_2004,
    ROUND(SUM(CASE WHEN Year = 2005 THEN Sales ELSE 0 END), 2) AS sales_2005_partial
FROM sales
GROUP BY Territory
ORDER BY sales_2004 DESC;


-- ------------------------------------------------------------
-- Q7. Core KPI summary (single-row scorecard, for the Excel/
--     Power BI KPI cards)
-- ------------------------------------------------------------
SELECT
    COUNT(DISTINCT "Order Number")                     AS total_orders,
    ROUND(SUM(Sales), 2)                                AS total_sales,
    ROUND(SUM("Est. Profit"), 2)                        AS total_profit,
    ROUND(SUM("Est. Profit") * 100.0 / SUM(Sales), 2)   AS overall_margin_pct,
    ROUND(SUM(Sales) * 1.0
          / COUNT(DISTINCT "Order Number"), 2)          AS avg_order_value,
    COUNT(DISTINCT "Customer Name")                     AS total_customers,
    COUNT(DISTINCT Territory)                           AS total_territories
FROM sales;
