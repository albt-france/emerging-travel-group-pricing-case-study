-- 04_top_partner_variations.sql
-- Which partners explain a drop? Last 14 days vs the 14 days before, by partner.
-- Joins the partners table for type, country and tier. Top 10 drops in sales.

WITH periods AS (
    SELECT
        b.partner_id,
        SUM(CASE WHEN b.booking_date BETWEEN DATE '2026-09-15' AND DATE '2026-09-28' THEN b.sales_eur ELSE 0 END) AS sales_now,
        SUM(CASE WHEN b.booking_date BETWEEN DATE '2026-09-01' AND DATE '2026-09-14' THEN b.sales_eur ELSE 0 END) AS sales_before
    FROM bookings AS b
    WHERE b.booking_date BETWEEN DATE '2026-09-01' AND DATE '2026-09-28'
    GROUP BY b.partner_id
)
SELECT
    p.partner_id,
    pa.partner_type,
    pa.country,
    pa.tier,
    ROUND(p.sales_before, 0)                      AS sales_before_eur,
    ROUND(p.sales_now, 0)                         AS sales_now_eur,
    ROUND(p.sales_now - p.sales_before, 0)        AS change_eur,
    ROUND(p.sales_now / NULLIF(p.sales_before, 0) - 1, 3) AS change_pct
FROM periods AS p
JOIN partners AS pa ON pa.partner_id = p.partner_id
ORDER BY change_eur ASC                           -- biggest drops first
LIMIT 10;
