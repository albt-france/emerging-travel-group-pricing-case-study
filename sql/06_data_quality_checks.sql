-- 06_data_quality_checks.sql
-- Checks to run on the RAW ledger before reading any number. Each block returns
-- the rows that need attention; an empty result means the check passed.

-- 6a. Duplicated booking ids, by booking date and channel
SELECT 'duplicate booking_id' AS check_name,
       CAST(booking_date AS DATE) AS booking_date, channel,
       COUNT(*) - COUNT(DISTINCT booking_id) AS extra_rows
FROM raw_bookings
GROUP BY booking_date, channel
HAVING COUNT(*) > COUNT(DISTINCT booking_id)

UNION ALL

-- 6b. Missing net cost (margin would be overstated if read as zero), by supplier and day
SELECT 'missing net cost', CAST(booking_date AS DATE), supplier_id, COUNT(*)
FROM raw_bookings
WHERE net_cost_eur IS NULL
GROUP BY booking_date, supplier_id

UNION ALL

-- 6c. Feed time-zone shift: days where the average booking hour of a busy channel
--     (50+ bookings a day) moves by 3 hours or more from its usual average hour
SELECT 'booking hour shifted', booking_date, channel, CAST(ROUND(avg_hour - usual_hour) AS INTEGER)
FROM (
    SELECT CAST(booking_date AS DATE) AS booking_date, channel,
           COUNT(*)                                   AS bookings,
           AVG(HOUR(booking_datetime_utc))            AS avg_hour,
           AVG(AVG(HOUR(booking_datetime_utc))) OVER (PARTITION BY channel) AS usual_hour
    FROM raw_bookings
    GROUP BY booking_date, channel
) AS t
WHERE bookings >= 50
  AND ABS(avg_hour - usual_hour) >= 3

ORDER BY check_name, booking_date;
