"""Download real daily euro foreign exchange reference rates from the ECB Data Portal.

Series key EXR.D.<CUR>.EUR.SP00.A = units of <CUR> for 1 EUR (daily, business days).
Weekends/holidays are forward-filled in data/public/ecb_fx_daily.csv so every calendar day
has a rate (the usual convention for converting daily sales).
Licence: ECB statistics may be reused with source attribution.
"""
from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw" / "ecb_fx"
PUBLIC = ROOT / "data" / "public"
CURRENCIES = ["USD", "GBP", "TRY", "PLN", "BRL", "MXN", "INR", "CHF"]
START, END = "2024-04-01", "2026-10-31"
UA = {"User-Agent": "Mozilla/5.0 (compatible; case-study-research/1.0)"}


def main() -> None:
    RAW.mkdir(parents=True, exist_ok=True)
    url = ("https://data-api.ecb.europa.eu/service/data/EXR/D." + "+".join(CURRENCIES)
           + ".EUR.SP00.A")
    resp = requests.get(url, params={"startPeriod": START, "endPeriod": END, "format": "csvdata"},
                        headers=UA, timeout=120)
    resp.raise_for_status()
    (RAW / f"ecb_exr_daily_{START}_{END}.csv").write_text(resp.text, encoding="utf-8")

    raw = pd.read_csv(io.StringIO(resp.text))
    wide = raw.pivot(index="TIME_PERIOD", columns="CURRENCY", values="OBS_VALUE")
    wide.index = pd.to_datetime(wide.index)
    days = pd.date_range(START, END, freq="D")
    wide = wide.reindex(days).ffill().bfill()  # carry Friday's rate over the weekend
    wide["EUR"] = 1.0
    tidy = (wide.rename_axis("date").reset_index()
            .melt(id_vars="date", var_name="currency", value_name="units_per_eur"))
    tidy.to_csv(PUBLIC / "ecb_fx_daily.csv", index=False, date_format="%Y-%m-%d")
    print(tidy.groupby("currency").units_per_eur.agg(["first", "last", "min", "max"]))


if __name__ == "__main__":
    main()
