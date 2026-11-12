-- 01_daily_sales_by_segment.sql
-- Daily sales, margin and margin rate by brand and line of business.
-- Grain: one row per booking date x brand x line of business.
-- Sales and margin are counted on the booking date (bookings made that day, before cancellations).

SELECT
    b.booking_date,
    b.brand,
    b.line_of_business,
    COUNT(*)                                   AS bookings,
    SUM(b.room_nights)                         AS room_nights,
    ROUND(SUM(b.sales_eur), 2)                 AS sales_eur,
    ROUND(SUM(b.margin_eur), 2)                AS margin_eur,
    ROUND(SUM(b.margin_eur) / SUM(b.sales_eur), 4) AS margin_rate,   -- margin / sales
    ROUND(SUM(b.sales_eur) / COUNT(*), 2)      AS avg_booking_value
FROM bookings AS b
WHERE b.booking_date BETWEEN DATE '2026-10-01' AND DATE '2026-10-31'  -- change the period here
GROUP BY b.booking_date, b.brand, b.line_of_business
ORDER BY b.booking_date, sales_eur DESC;
