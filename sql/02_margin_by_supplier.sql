-- 02_margin_by_supplier.sql
-- Monthly margin rate by supplier, with the supplier type from the suppliers table (JOIN).
-- Use it to spot a supplier whose margin falls while its volume holds (net-rate increase).

SELECT
    STRFTIME(b.booking_date, '%Y-%m')           AS year_month,
    s.supplier_type,
    b.supplier_id,
    s.supplier_name,
    COUNT(*)                                     AS bookings,
    ROUND(SUM(b.sales_eur), 0)                   AS sales_eur,
    ROUND(SUM(b.margin_eur) / SUM(b.sales_eur), 4) AS margin_rate,
    SUM(CASE WHEN b.cost_imputed THEN 1 ELSE 0 END) AS rows_with_imputed_cost  -- data-quality flag
FROM bookings AS b
JOIN suppliers AS s
  ON s.supplier_id = b.supplier_id
WHERE b.line_of_business = 'Hotels'
  AND b.destination_country = 'ES'                -- example filter: Spain
GROUP BY ALL                                      -- groups by every non-aggregated column
ORDER BY b.supplier_id, year_month;
