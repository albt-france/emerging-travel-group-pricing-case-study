# SQL queries (DuckDB)

Plain, commented queries at the level of the job ("read and modify queries, simple queries on several tables").
DuckDB reads the CSV files directly, no database server needed.

```bash
python src/run_sql.py        # runs 00_views.sql, then every query; results in data/output/sql/
```

Or interactively: `duckdb` → `.read sql/00_views.sql` → paste a query.

All booking data is **SIMULATED** (see `data/ASSUMPTIONS.md`).

| File | Question it answers | SQL used |
|---|---|---|
| `00_views.sql` | Defines the tables: raw ledger, clean ledger, suppliers, partners, price-match requests, calendar | `CREATE VIEW`, `read_csv_auto` |
| `01_daily_sales_by_segment.sql` | Daily bookings, sales, margin and margin rate by brand and line of business | `GROUP BY`, `SUM`, `COUNT`, date filter `BETWEEN` |
| `02_margin_by_supplier.sql` | Monthly margin rate by supplier, with the supplier type; flags imputed costs | `JOIN` suppliers, `STRFTIME`, `CASE WHEN`, `GROUP BY ALL` |
| `03_alerts_same_weekday.sql` | Which source markets had an abnormal day? Same rule as Python and Excel | CTEs (`WITH`), window functions `AVG() OVER`, `STDDEV_SAMP() OVER`, `DAYOFWEEK` |
| `04_top_partner_variations.sql` | Which partners explain a drop? Last 14 days vs the 14 before, top 10 drops | conditional `SUM(CASE ...)`, `JOIN` partners, `NULLIF`, `ORDER BY ... LIMIT` |
| `05_price_match_summary.sql` | Price-match requests by partner type and decision: claimed gap, discount cost, handling time | `JOIN`, `MEDIAN`, `AVG(CASE ...)` |
| `06_data_quality_checks.sql` | Duplicates, missing net costs, time-zone shift in a feed (run on the raw ledger) | `HAVING`, `UNION ALL`, `IS NULL`, window over an aggregate |

## How to read each query

**01** - one row per day × brand × line of business. `margin_rate = SUM(margin) / SUM(sales)`: a ratio of sums, not an
average of ratios (an average of daily ratios would give a small day the same weight as a big one). Change the dates
in the `WHERE` clause to look at another period.

**02** - the `JOIN` brings `supplier_type` and `supplier_name` from the suppliers table. Filtered on Spain to show the
Bedbank C net-rate increase (incident E1): its margin rate falls from about 8-9% to about 2% from June 2026 while its
bookings stay flat. `rows_with_imputed_cost` reminds the reader which months contain repaired data.

**03** - four steps: (1) bookings per day and market; (2) `expected` = average of the same weekday over the previous
4 weeks (`PARTITION BY client_country, DAYOFWEEK(...)` then `ROWS BETWEEN 4 PRECEDING AND 1 PRECEDING`);
(3) relative deviation; (4) normal noise = standard deviation of the deviation over the previous 56 days. An alert is
|deviation / noise| ≥ 3.5. It returns exactly the same 9 market alerts as `src/analysis.py`.

**04** - the "weekly drill-down" layer. A partner of 10 bookings a day that stops is invisible in a daily total but
comes first here (API-002, incident E2).

**05** - the price-match KPIs by partner type. `booked_rate` after a refusal shows whether refusals are well argued.

**06** - run before reading any number. Each block returns the rows to look at; an empty result means the check
passed. The time-zone check is a simple version (average booking hour moves by 3 h or more on a busy channel);
`src/analysis.py` uses a stricter profile-matching version.
