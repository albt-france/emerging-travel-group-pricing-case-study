"""Run every query in sql/ with DuckDB and save the results to data/output/sql/.

Usage (from the repository root):  python src/run_sql.py
"""
from pathlib import Path
import os

import duckdb

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "output" / "sql"


def main() -> None:
    os.chdir(ROOT)                      # the queries use paths relative to the repository root
    OUT.mkdir(parents=True, exist_ok=True)
    con = duckdb.connect()              # in-memory database, nothing written to disk
    con.execute((ROOT / "sql" / "00_views.sql").read_text(encoding="utf-8"))
    for path in sorted((ROOT / "sql").glob("[0-9][1-9]_*.sql")):
        df = con.execute(path.read_text(encoding="utf-8")).df()
        df.to_csv(OUT / f"{path.stem}.csv", index=False)
        print(f"{path.name:40s} {len(df):>6} rows")


if __name__ == "__main__":
    main()
