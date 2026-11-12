-- 00_views.sql
-- Creates the tables used by the other queries (DuckDB reads the CSV files directly).
-- Run from the repository root:  python src/run_sql.py
-- All booking data is SIMULATED (see data/ASSUMPTIONS.md).

-- Raw ledger as received (contains the injected data-quality issues).
CREATE OR REPLACE VIEW raw_bookings AS
SELECT * FROM read_csv_auto('data/simulated/bookings.csv');

-- Clean ledger produced by src/analysis.py after the data-quality checks:
-- duplicates removed, missing costs imputed (cost_imputed = true), time-zone shift corrected.
-- It adds two columns: sales_eur = sell_price_eur - discount_eur, margin_eur = sales_eur - net_cost_eur.
CREATE OR REPLACE VIEW bookings AS
SELECT * FROM read_csv_auto('data/output/sim/bookings_clean.csv');

CREATE OR REPLACE VIEW suppliers AS
SELECT * FROM read_csv_auto('data/simulated/suppliers.csv');

CREATE OR REPLACE VIEW partners AS
SELECT * FROM read_csv_auto('data/simulated/partners.csv');

CREATE OR REPLACE VIEW price_match_requests AS
SELECT * FROM read_csv_auto('data/simulated/price_match_requests.csv');

CREATE OR REPLACE VIEW calendar AS
SELECT * FROM read_csv_auto('data/simulated/calendar.csv');
