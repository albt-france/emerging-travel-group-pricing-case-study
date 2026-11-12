"""Commercial analyses on a REAL hotel booking dataset.

Dataset: "Hotel booking demand datasets", Antonio, Almeida & Nunes, Data in Brief 22 (2019)
41-49, doi:10.1016/j.dib.2018.11.126, licence CC BY 4.0. Two hotels in Portugal (a resort
hotel and a city hotel), arrivals July 2015 - August 2017. It is used to demonstrate the
analysis on real bookings; it is NOT Emerging Travel Group data.

Outputs (data/output/real/):
  dq_checks.csv               data-quality checks run before any conclusion
  channel_kpis.csv            B2B (TA/TO) vs direct vs corporate: price, cancellations, lead time
  segment_kpis.csv            same by market segment
  deposit_cancellation.csv    cancellation rate by deposit type (a suspicious pattern to verify)
  agent_pareto.csv            revenue concentration on travel agents
  monthly.csv                 monthly room nights, revenue and ADR by hotel
  bridge.csv                  revenue bridge (volume / mix / price) between two real months
  bridge_drilldown.csv        the same change split by hotel x market segment
"""
from __future__ import annotations

import pandas as pd

from common import OUT, RAW

OUT_REAL = OUT / "real"
MONTH_FROM, MONTH_TO = "2016-08", "2016-09"   # the revenue swing explained on page 4


def load() -> pd.DataFrame:
    d = pd.read_csv(RAW / "hotel_bookings" / "hotels.csv")
    d["nights"] = d["stays_in_weekend_nights"] + d["stays_in_week_nights"]
    d["arrival_month"] = pd.to_datetime(
        d["arrival_date_year"].astype(str) + "-" + d["arrival_date_month"] + "-01",
        format="%Y-%B-%d").dt.to_period("M")
    d["guests"] = d["adults"] + d["children"].fillna(0) + d["babies"]
    # Channel groups used in the report: TA/TO = travel agents and tour operators (B2B).
    d["channel"] = d["distribution_channel"].replace({"Undefined": "Other", "GDS": "GDS"})
    return d


def dq_checks(d: pd.DataFrame) -> pd.DataFrame:
    checks = [
        ("Rows in file", len(d), "Matches the 119,390 bookings described in the paper"),
        ("Fully duplicated rows", int(d.duplicated().sum()),
         "No booking ID in the file: identical rows can be real group bookings. Kept, flagged."),
        ("Negative ADR", int((d["adr"] < 0).sum()), "Excluded from price statistics"),
        ("ADR above 1,000 EUR", int((d["adr"] > 1000).sum()), "Outlier, excluded from price statistics"),
        ("Zero ADR, not complimentary", int(((d["adr"] == 0) & (d["market_segment"] != "Complementary")).sum()),
         "Kept in volumes, excluded from price statistics"),
        ("Zero nights", int((d["nights"] == 0).sum()), "Day use or error: excluded from nights and ADR"),
        ("Zero guests", int((d["guests"] == 0).sum()), "Probable data-entry error: excluded"),
        ("Missing country", int(d["country"].isna().sum()), "Shown as 'Unknown'"),
        ("Undefined channel", int((d["distribution_channel"] == "Undefined").sum()), "Grouped as 'Other'"),
    ]
    return pd.DataFrame(checks, columns=["check", "rows", "decision"])


def clean_for_prices(d: pd.DataFrame) -> pd.Series:
    """Rows usable for price statistics (ADR, revenue)."""
    return (d["adr"] > 0) & (d["adr"] <= 1000) & (d["nights"] > 0) & (d["guests"] > 0)


def kpis_by(d: pd.DataFrame, key: str) -> pd.DataFrame:
    ok = clean_for_prices(d)
    stayed = d[ok & (d["is_canceled"] == 0)]
    g = d.groupby(key)
    out = pd.DataFrame({
        "bookings": g.size(),
        "cancel_rate": g["is_canceled"].mean(),
        "median_lead_time_days": g["lead_time"].median(),
        "avg_nights": g["nights"].mean(),
    })
    s = stayed.groupby(key)
    out["room_nights_stayed"] = s["nights"].sum()
    out["revenue_eur"] = (stayed["adr"] * stayed["nights"]).groupby(stayed[key]).sum()
    out["adr_eur"] = out["revenue_eur"] / out["room_nights_stayed"]
    out["share_of_bookings"] = out["bookings"] / out["bookings"].sum()
    out["share_of_revenue"] = out["revenue_eur"] / out["revenue_eur"].sum()
    return out.reset_index().sort_values("bookings", ascending=False)


def agent_pareto(d: pd.DataFrame) -> pd.DataFrame:
    ok = clean_for_prices(d) & (d["is_canceled"] == 0) & d["agent"].notna()
    a = d[ok].assign(revenue=lambda x: x["adr"] * x["nights"])
    rev = a.groupby("agent")["revenue"].sum().sort_values(ascending=False)
    out = rev.reset_index()
    out["rank"] = range(1, len(out) + 1)
    out["cum_share"] = out["revenue"].cumsum() / out["revenue"].sum()
    out["agent"] = "Agent " + out["agent"].astype(int).astype(str)   # IDs are anonymised in the source
    return out


def revenue_bridge(d: pd.DataFrame, m0: str, m1: str, seg_cols: list[str]):
    """Split the revenue change between two months into volume, mix and price effects.

    Revenue R = sum over segments of N_s * P_s  (N = room nights, P = ADR).
      Volume = (N1 - N0) * P0_avg               more or fewer nights at the old average price
      Mix    = N1 * sum((w1_s - w0_s) * P0_s)   nights moved between cheap and expensive segments
      Price  = N1 * sum(w1_s * (P1_s - P0_s))   price change inside each segment
    The three effects add up exactly to R1 - R0.
    """
    ok = clean_for_prices(d) & (d["is_canceled"] == 0)
    x = d[ok].assign(revenue=lambda t: t["adr"] * t["nights"])
    x = x[x["arrival_month"].astype(str).isin([m0, m1])]
    seg = x.groupby(["arrival_month"] + seg_cols).agg(N=("nights", "sum"), R=("revenue", "sum")).reset_index()
    seg["period"] = seg["arrival_month"].astype(str).map({m0: "0", m1: "1"})
    w = seg.pivot_table(index=seg_cols, columns="period", values=["N", "R"], aggfunc="sum").fillna(0)
    w.columns = [f"{a}{b}" for a, b in w.columns]
    N0, N1, R0, R1 = w["N0"].sum(), w["N1"].sum(), w["R0"].sum(), w["R1"].sum()
    p0_avg = R0 / N0
    w["P0"] = (w["R0"] / w["N0"]).where(w["N0"] > 0)
    w["P1"] = (w["R1"] / w["N1"]).where(w["N1"] > 0)
    w["P0"] = w["P0"].fillna(w["P1"])      # segment absent in month 0: no price effect
    w["P1"] = w["P1"].fillna(w["P0"])
    w["w0"], w["w1"] = w["N0"] / N0, w["N1"] / N1
    w["volume"] = (w["N1"] - w["N0"]) * p0_avg          # allocated by segment for the drill-down
    w["mix"] = N1 * (w["w1"] - w["w0"]) * (w["P0"] - p0_avg)
    w["price"] = N1 * w["w1"] * (w["P1"] - w["P0"])
    # NB: (P0 - p0_avg) in the mix term gives the same total (sum of w1-w0 is zero) and a
    # cleaner per-segment reading: moving nights to a segment priced above average is positive.
    w["total_change"] = w["R1"] - w["R0"]
    bridge = pd.DataFrame({
        "step": [f"Revenue {m0}", "Volume", "Mix", "Price", f"Revenue {m1}"],
        "eur": [R0, w["volume"].sum(), w["mix"].sum(), w["price"].sum(), R1],
    })
    bridge.attrs["check"] = R0 + w["volume"].sum() + w["mix"].sum() + w["price"].sum() - R1
    return bridge, w.reset_index()


def main() -> None:
    OUT_REAL.mkdir(parents=True, exist_ok=True)
    d = load()
    dq_checks(d).to_csv(OUT_REAL / "dq_checks.csv", index=False)
    kpis_by(d, "channel").to_csv(OUT_REAL / "channel_kpis.csv", index=False)
    kpis_by(d, "market_segment").to_csv(OUT_REAL / "segment_kpis.csv", index=False)

    dep = d.groupby(["deposit_type"]).agg(bookings=("is_canceled", "size"),
                                          cancel_rate=("is_canceled", "mean")).reset_index()
    nr = d[d["deposit_type"] == "Non Refund"]
    dep["share_from_PRT"] = dep["deposit_type"].map(
        {"Non Refund": (nr["country"] == "PRT").mean()})
    dep["share_TA_TO"] = dep["deposit_type"].map(
        {"Non Refund": (nr["channel"] == "TA/TO").mean()})
    dep.to_csv(OUT_REAL / "deposit_cancellation.csv", index=False)

    agent_pareto(d).to_csv(OUT_REAL / "agent_pareto.csv", index=False)

    ok = clean_for_prices(d) & (d["is_canceled"] == 0)
    m = (d[ok].assign(revenue=lambda t: t["adr"] * t["nights"])
         .groupby(["arrival_month", "hotel"]).agg(room_nights=("nights", "sum"),
                                                  revenue_eur=("revenue", "sum")).reset_index())
    m["adr_eur"] = m["revenue_eur"] / m["room_nights"]
    canc = d.groupby(["arrival_month", "hotel"])["is_canceled"].mean().rename("cancel_rate").reset_index()
    m = m.merge(canc, on=["arrival_month", "hotel"])
    m.to_csv(OUT_REAL / "monthly.csv", index=False)

    bridge, drill = revenue_bridge(d, MONTH_FROM, MONTH_TO, ["hotel", "market_segment"])
    bridge.to_csv(OUT_REAL / "bridge.csv", index=False)
    drill.to_csv(OUT_REAL / "bridge_drilldown.csv", index=False)

    # Lead time vs cancellation (B2B vs direct), for the cancellation discussion.
    bins = [-1, 7, 30, 90, 180, 365, 1000]
    labels = ["0-7", "8-30", "31-90", "91-180", "181-365", "365+"]
    lt = (d.assign(lead_bucket=pd.cut(d["lead_time"], bins=bins, labels=labels))
          .groupby(["channel", "lead_bucket"], observed=True)["is_canceled"].agg(["size", "mean"])
          .reset_index().rename(columns={"size": "bookings", "mean": "cancel_rate"}))
    lt.to_csv(OUT_REAL / "lead_time_cancellation.csv", index=False)

    print(bridge.to_string(index=False), "\ncheck (should be 0):", round(bridge.attrs["check"], 6))


if __name__ == "__main__":
    main()
