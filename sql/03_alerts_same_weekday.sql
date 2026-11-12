-- 03_alerts_same_weekday.sql
-- Daily volume alerts by client country (source market). Same rule as src/analysis.py
-- and the Excel "Alerts" sheet:
--   expected  = average of the same weekday over the previous 4 weeks
--   deviation = bookings / expected - 1
--   noise     = standard deviation of the deviation over the previous 56 days
--   alert     = |deviation / noise| >= 3.5, for markets expected at 20+ bookings a day

WITH daily AS (                                     -- step 1: one row per day and market
    SELECT booking_date, client_country, COUNT(*) AS bookings
    FROM bookings
    GROUP BY booking_date, client_country
),
expected AS (                                       -- step 2: same weekday, previous 4 weeks
    SELECT *,
           AVG(bookings) OVER (
               PARTITION BY client_country, DAYOFWEEK(booking_date)
               ORDER BY booking_date
               ROWS BETWEEN 4 PRECEDING AND 1 PRECEDING) AS expected_bookings
    FROM daily
),
deviation AS (                                      -- step 3: relative deviation of each day
    SELECT *, bookings / expected_bookings - 1 AS deviation
    FROM expected
),
noise AS (                                          -- step 4: normal noise over the previous 56 days
    SELECT *,
           STDDEV_SAMP(deviation) OVER w AS noise_sd,
           COUNT(deviation)       OVER w AS days_in_noise
    FROM deviation
    WINDOW w AS (PARTITION BY client_country ORDER BY booking_date
                 ROWS BETWEEN 56 PRECEDING AND 1 PRECEDING)
)
SELECT                                              -- step 5: keep the alerts
    booking_date,
    client_country,
    bookings,
    ROUND(expected_bookings, 1)        AS expected_bookings,
    ROUND(deviation, 3)                AS deviation_pct,
    ROUND(deviation / noise_sd, 2)     AS z_score
FROM noise
WHERE booking_date >= (SELECT MIN(booking_date) FROM daily) + INTERVAL 56 DAY
  AND days_in_noise >= 28
  AND expected_bookings >= 20
  AND ABS(deviation / noise_sd) >= 3.5
ORDER BY booking_date, client_country;
