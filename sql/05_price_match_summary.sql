-- 05_price_match_summary.sql
-- Price-match requests: volume, decisions, discount cost and handling time by partner type.
-- Joins the partners table to get the partner type.

SELECT
    pa.partner_type,
    r.decision,
    COUNT(*)                                                     AS requests,
    ROUND(AVG(1 - r.competitor_price_eur / r.etg_price_eur), 3)  AS avg_claimed_gap_pct,
    ROUND(SUM(CASE WHEN r.booked = 'Yes' THEN r.discount_eur ELSE 0 END), 0) AS discount_cost_eur,
    ROUND(MEDIAN(r.handling_hours), 1)                           AS median_handling_hours,
    ROUND(AVG(CASE WHEN r.booked = 'Yes' THEN 1.0 ELSE 0 END), 3) AS booked_rate
FROM price_match_requests AS r
JOIN partners AS pa ON pa.partner_id = r.partner_id
GROUP BY pa.partner_type, r.decision
ORDER BY pa.partner_type, requests DESC;
