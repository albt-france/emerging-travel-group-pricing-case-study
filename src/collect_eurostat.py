"""Download Eurostat tourism statistics (open data, Eurostat API, JSON-stat format).

Datasets (licence: Eurostat re-use policy, free re-use with attribution):
- tour_occ_ninraw : nights spent in EU tourist accommodation, by country of residence
                    of the guest (annual)  -> which source markets fill EU hotels.
- tour_dem_tnmd   : nights spent by residents of each EU country, by month of departure
                    (annual survey)          -> the seasonality of each source market.
- tour_occ_nim    : nights spent in tourist accommodation, monthly, by destination
                    country, residents vs non-residents -> recent monthly trend.

Raw JSON is saved to data/raw/eurostat/, tidy CSVs to data/public/.
Requests are spaced by 2 seconds.
"""
from __future__ import annotations

import itertools
import json
import time
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw" / "eurostat"
PUBLIC = ROOT / "data" / "public"
BASE = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/"
UA = {"User-Agent": "Mozilla/5.0 (compatible; case-study-research/1.0)"}

# Source markets shown in the report (EU residents only for the month-of-departure survey).
SEASONALITY_MARKETS = ["DE", "FR", "IT", "ES", "NL", "PL", "BE", "AT", "SE", "CZ"]
DESTINATIONS = ["EU27_2020", "ES", "IT", "FR", "DE", "AT", "EL", "PT", "NL", "HR"]


def fetch(dataset: str, params: list[tuple[str, str]], name: str) -> dict:
    """Download one Eurostat query (cached on disk) and return the JSON-stat dict."""
    path = RAW / f"{name}.json"
    if not path.exists():
        resp = requests.get(BASE + dataset, params=[("format", "JSON"), ("lang", "EN")] + params,
                            headers=UA, timeout=180)
        resp.raise_for_status()
        path.write_text(resp.text, encoding="utf-8")
        time.sleep(2)
    return json.loads(path.read_text(encoding="utf-8"))


def jsonstat_to_frame(js: dict) -> pd.DataFrame:
    """Turn a JSON-stat 2.0 response into a long table (one row per cell)."""
    dims = js["id"]
    codes = []
    for d in dims:
        index = js["dimension"][d]["category"]["index"]
        codes.append(sorted(index, key=index.get))
    rows = []
    for flat, combo in enumerate(itertools.product(*codes)):
        value = js["value"].get(str(flat))
        if value is not None:
            rows.append(dict(zip(dims, combo), value=value))
    frame = pd.DataFrame(rows)
    # Add readable labels for every coded dimension.
    for d in dims:
        labels = js["dimension"][d]["category"].get("label", {})
        frame[d + "_label"] = frame[d].map(labels)
    return frame


def main() -> None:
    RAW.mkdir(parents=True, exist_ok=True)
    PUBLIC.mkdir(parents=True, exist_ok=True)

    # A. Source markets of EU accommodation (annual nights by country of residence).
    js = fetch("tour_occ_ninraw",
               [("unit", "NR"), ("nace_r2", "I551-I553"), ("geo", "EU27_2020")]
               + [("time", y) for y in ["2019", "2023", "2024", "2025"]],
               "tour_occ_ninraw_EU27")
    a = jsonstat_to_frame(js)
    a = a.rename(columns={"time": "year", "c_resid": "residence_code",
                          "c_resid_label": "residence", "value": "nights"})
    a[["year", "residence_code", "residence", "nights"]].to_csv(
        PUBLIC / "eurostat_nights_by_residence.csv", index=False)

    # B. Seasonality of each source market (nights by month of departure, trips abroad).
    js = fetch("tour_dem_tnmd",
               [("unit", "NR"), ("purpose", "TOTAL"), ("duration", "N_GE1"), ("c_dest", "FOR")]
               + [("geo", g) for g in SEASONALITY_MARKETS]
               + [("time", y) for y in ["2023", "2024", "2025"]],
               "tour_dem_tnmd_markets")
    b = jsonstat_to_frame(js)
    b = b[b["month"].str.fullmatch(r"M\d\d")]  # keep the 12 calendar months
    b = b.rename(columns={"geo": "market_code", "geo_label": "market", "time": "year",
                          "value": "nights"})
    b["month"] = b["month"].str[1:].astype(int)
    b[["year", "market_code", "market", "month", "nights"]].to_csv(
        PUBLIC / "eurostat_market_seasonality.csv", index=False)

    # C. Monthly nights by destination, residents (DOM) vs non-residents (FOR).
    js = fetch("tour_occ_nim",
               [("unit", "NR"), ("nace_r2", "I551-I553"), ("c_resid", "DOM"), ("c_resid", "FOR")]
               + [("geo", g) for g in DESTINATIONS] + [("sinceTimePeriod", "2019-01")],
               "tour_occ_nim_destinations")
    c = jsonstat_to_frame(js)
    c = c.rename(columns={"geo": "destination_code", "geo_label": "destination",
                          "time": "month", "c_resid": "residence", "value": "nights"})
    c[["month", "destination_code", "destination", "residence", "nights"]].to_csv(
        PUBLIC / "eurostat_monthly_nights_by_destination.csv", index=False)

    print("A", a.shape, "B", b.shape, "C", c.shape)


if __name__ == "__main__":
    main()
