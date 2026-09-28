-- ============================================================
-- Supply chain analysis queries (SQLite)
-- Each query starts with "-- name: <query_name>" so Python can
-- load and run them by name (see src/database.py).
-- ============================================================

-- name: supplier_delivery_performance
-- On-time delivery (OTD), actual lead time and fill rate per supplier.
SELECT
    s.supplier_id,
    s.supplier_name,
    s.country,
    s.promised_lead_time_wks,
    COUNT(*)                                                         AS po_count,
    ROUND(SUM(po.po_value_cad), 2)                                   AS total_spend_cad,
    ROUND(AVG(CASE WHEN po.received_date <= po.promised_date
                   THEN 1.0 ELSE 0.0 END), 4)                        AS on_time_rate,
    ROUND(AVG(julianday(po.received_date) - julianday(po.order_date)) / 7.0, 2)
                                                                     AS avg_lead_time_wks,
    ROUND(AVG(MAX(julianday(po.received_date) - julianday(po.promised_date), 0)), 2)
                                                                     AS avg_days_late,
    ROUND(1.0 * SUM(po.qty_received) / SUM(po.qty_ordered), 4)       AS fill_rate,
    ROUND(1.0 * SUM(po.qty_defective) / SUM(po.qty_received), 4)     AS defect_rate
FROM purchase_orders po
JOIN suppliers s ON s.supplier_id = po.supplier_id
GROUP BY s.supplier_id
ORDER BY on_time_rate DESC;

-- name: supplier_lead_times
-- Every individual lead time in weeks (used to measure variability in Python,
-- because SQLite has no built-in standard deviation function).
SELECT
    supplier_id,
    sku,
    (julianday(received_date) - julianday(order_date)) / 7.0 AS lead_time_wks
FROM purchase_orders;

-- name: annual_spend_by_category
-- Last 52 weeks of demand valued at unit cost, by category.
SELECT
    p.category,
    COUNT(DISTINCT p.sku)                       AS sku_count,
    SUM(d.units)                                AS units_52w,
    ROUND(SUM(d.units * p.unit_cost_cad), 2)    AS spend_52w_cad
FROM weekly_demand d
JOIN products p ON p.sku = d.sku
WHERE d.week_start > (SELECT date(MAX(week_start), '-364 days') FROM weekly_demand)
GROUP BY p.category
ORDER BY spend_52w_cad DESC;

-- name: monthly_late_trend
-- Share of late purchase orders per month - is supplier reliability improving?
SELECT
    strftime('%Y-%m', promised_date)                                    AS month,
    COUNT(*)                                                            AS po_count,
    ROUND(AVG(CASE WHEN received_date > promised_date THEN 1.0 ELSE 0.0 END), 4)
                                                                        AS late_rate
FROM purchase_orders
GROUP BY month
ORDER BY month;

-- name: single_source_exposure
-- How much of total spend depends on each supplier (concentration risk).
SELECT
    supplier_id,
    COUNT(DISTINCT sku)                                                     AS skus_supplied,
    ROUND(SUM(po_value_cad), 2)                                             AS spend_cad,
    ROUND(SUM(po_value_cad) / (SELECT SUM(po_value_cad) FROM purchase_orders), 4)
                                                                            AS share_of_spend
FROM purchase_orders
GROUP BY supplier_id
ORDER BY share_of_spend DESC;
