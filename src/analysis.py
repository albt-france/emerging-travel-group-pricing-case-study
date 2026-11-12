"""Daily commercial monitoring on the SIMULATED ledger: data quality first, then cockpit,
alerts, root causes, price-match statistics and pricing recommendations.

Order matters: the data-quality checks run BEFORE any business conclusion, and the cockpit
uses the cleaned table. Every number in the report pages 5-9 comes from this script
(outputs in data/output/sim/), from sql/*.sql or from the Excel workbook formulas.
"""
from __future__ import annotations

import numpy as np
import pandas as pd

from common import OUT, SIM

OUT_SIM = OUT / "sim"
MARGIN_FLOOR = 0.04
Z_ALERT = 3.5             # alert threshold in standard deviations of the normal daily noise
BASELINE_WEEKS = 8         # the normal noise is measured over the previous 8 weeks
MIN_SEGMENT_BOOKINGS = 20  # no alert on segments expected below 20 bookings a day

# Event windows used ONLY to score the alerts afterwards (did the cockpit find them?).
EVENT_KEYS = [
    ("E1 supplier net rate", "2026-06-15", "2026-06-28", {"SUP-BB-C", "ES"}),
    ("E2 API partner lost", "2026-09-07", "2026-09-28", {"API-002", "API", "GB"}),
    ("E4 UAE holiday", "2026-05-11", "2026-05-25", {"AE"}),
    ("E4 UAE holiday", "2025-05-22", "2025-06-05", {"AE"}),
    ("E6 Greece cancellations", "2026-08-08", "2026-08-14", {"EL"}),
]
ELASTICITY_RANGE = (-1.0, -3.0)   # assumption for B2B hotel demand (prudent range)


# --------------------------------------------------------------------------- loading
def load() -> pd.DataFrame:
    b = pd.read_csv(SIM / "bookings.csv", parse_dates=["booking_date", "checkin_date", "cancel_date",
                                                       "booking_datetime_utc"])
    return b


# --------------------------------------------------------------------------- data quality
def dq_checks(b: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return (issues table, clean bookings). Each check says what it found and what we did."""
    issues = []

    # 1. duplicated booking ids
    dup = b[b.duplicated("booking_id", keep="first")]
    if len(dup):
        by = dup.groupby([dup["booking_date"].dt.date, "channel"]).size().reset_index(name="rows")
        for _, r in by.iterrows():
            day = b[b["booking_date"].dt.date == r["booking_date"]]
            raw_sales = day["sell_price_eur"].sum()
            clean_sales = day.drop_duplicates("booking_id")["sell_price_eur"].sum()
            issues.append(dict(issue="Duplicated bookings", date_from=str(r["booking_date"]), date_to=str(r["booking_date"]),
                               scope=f"Channel {r['channel']}", rows=int(r["rows"]),
                               effect=f"Day sales overstated by {raw_sales / clean_sales - 1:+.0%} "
                                      f"(EUR {raw_sales - clean_sales:,.0f})",
                               action="Removed duplicates on booking_id; ask data engineering to fix the load"))
    c = b.drop_duplicates("booking_id").copy()

    # 2. missing net cost
    miss = c[c["net_cost_eur"].isna()]
    if len(miss):
        for sup, g in miss.groupby("supplier_id"):
            d0, d1 = g["booking_date"].min().date(), g["booking_date"].max().date()
            week = c[c["booking_date"].between(pd.Timestamp(d0), pd.Timestamp(d1))]
            wrong = (week["sell_price_eur"] - week["net_cost_eur"].fillna(0)).sum() / week["sell_price_eur"].sum()
            ok = week.dropna(subset=["net_cost_eur"])
            right = (ok["sell_price_eur"] - ok["net_cost_eur"]).sum() / ok["sell_price_eur"].sum()
            issues.append(dict(issue="Missing net cost", date_from=str(d0), date_to=str(d1), scope=f"Supplier {sup}",
                               rows=len(g),
                               effect=f"If read as zero, total margin rate shows {wrong:.1%} instead of {right:.1%}",
                               action="Cost imputed with the supplier's markup of the previous 28 days, flagged; "
                                      "supplier invoice to be reconciled"))
    # impute: cost = sell / (1 + median markup of same supplier x partner type over previous 28 days)
    c["cost_imputed"] = c["net_cost_eur"].isna()
    if c["cost_imputed"].any():
        start = c.loc[c["cost_imputed"], "booking_date"].min()
        ref = c[(c["booking_date"] < start) & (c["booking_date"] >= start - pd.Timedelta(days=28))].dropna(
            subset=["net_cost_eur"])
        mk = (ref["sell_price_eur"] / ref["net_cost_eur"] - 1).groupby([ref["supplier_id"], ref["partner_type"]]).median()
        keys = list(zip(c.loc[c["cost_imputed"], "supplier_id"], c.loc[c["cost_imputed"], "partner_type"]))
        c.loc[c["cost_imputed"], "net_cost_eur"] = (
            c.loc[c["cost_imputed"], "sell_price_eur"].to_numpy() / (1 + mk.reindex(keys).to_numpy())).round(2)

    # 3. timestamp shift in a feed. For each channel and day we take the hourly profile of the
    #    bookings of that day and the two neighbouring days, and find the shift (0-8 h) that best
    #    matches the channel's usual profile. 3+ consecutive days with a shift >= 3 h = issue.
    c["hour"] = c["booking_datetime_utc"].dt.hour
    rows = []
    for ch, g in c.groupby("channel"):
        hist = pd.crosstab(g["booking_date"], g["hour"]).reindex(columns=range(24), fill_value=0)
        usual = hist.sum().to_numpy(float)
        usual /= usual.sum()
        pooled = hist.rolling(3, center=True, min_periods=1).sum()
        for day, h in pooled.iterrows():
            v = h.to_numpy(float) / max(h.sum(), 1)
            errors = [np.abs(v - np.roll(usual, k)).sum() for k in range(9)]
            rows.append((ch, day, int(np.argmin(errors))))
    med = pd.DataFrame(rows, columns=["channel", "booking_date", "shift"]).sort_values(["channel", "booking_date"])
    med["flag"] = med["shift"] >= 3        # (4+ consecutive days required below)
    med["run_id"] = (med["flag"] != med.groupby("channel")["flag"].shift()).cumsum()
    runs = (med[med["flag"]].groupby(["channel", "run_id"])
            .agg(d0=("booking_date", "min"), d1=("booking_date", "max"), days=("booking_date", "size"),
                 shift=("shift", "median")).reset_index())
    for _, run in runs[runs["days"] >= 4].iterrows():
        ch, shift_h = run["channel"], int(round(run["shift"]))
        d0, d1 = run["d0"], run["d1"]
        moved = c[(c["channel"] == ch) & c["booking_date"].between(d0, d1 + pd.Timedelta(days=1))
                  & (c["booking_datetime_utc"].dt.hour < shift_h)]
        if len(moved) < 10:      # too small to matter: noise in a low-volume channel
            continue
        issues.append(dict(issue="Time-zone shift in a feed", date_from=str(d0.date()), date_to=str(d1.date()),
                           scope=f"Channel {ch}", rows=int(len(moved)),
                           effect=f"Timestamps {shift_h:+d} h: {len(moved)} bookings counted on the wrong day",
                           action=f"Shifted back {shift_h} h for the window; feed owner to send UTC"))
        fix = (c["channel"] == ch) & c["booking_date"].between(d0, d1 + pd.Timedelta(days=1))
        # only rows of the faulty window (dates computed from shifted timestamps)
        # approximate correction: a clean re-load of the feed would be the real fix
        fix &= (c["booking_datetime_utc"] - pd.Timedelta(hours=shift_h)).dt.normalize().between(d0, d1)
        c.loc[fix, "booking_datetime_utc"] -= pd.Timedelta(hours=shift_h)
        c.loc[fix, "booking_date"] = c.loc[fix, "booking_datetime_utc"].dt.normalize()

    # 4. sanity checks that should return zero rows
    neg = (c["sell_price_eur"] - c["discount_eur"]) < c["net_cost_eur"]
    top_sup = c[neg].groupby("supplier_id").size().sort_values(ascending=False)
    issues.append(dict(issue="Sold below net cost (business alert, not a data error)", date_from=str(c[neg]["booking_date"].min().date()),
                       date_to=str(c[neg]["booking_date"].max().date()), scope="All bookings", rows=int(neg.sum()),
                       effect=f"{top_sup.iloc[0] / neg.sum():.0%} of them on {top_sup.index[0]}",
                       action="Sent to pricing: selling price below net cost after a supplier rate change or a discount"))
    for name, mask in [("Selling price <= 0", c["sell_price_eur"] <= 0),
                       ("Check-in before booking date", c["checkin_date"] < c["booking_date"]),
                       ("Cancellation before booking date", c["cancel_date"] < c["booking_date"]),
                       ("Missing FX rate", c["fx_rate_to_eur"].isna())]:
        issues.append(dict(issue=name, date_from="", date_to="", scope="All bookings", rows=int(mask.sum()),
                           effect="None" if mask.sum() == 0 else "Review", action="Check passed" if mask.sum() == 0
                           else "Listed for review (expected for some price-match discounts)"))
    c = c.drop(columns="hour")
    c["sales_eur"] = c["sell_price_eur"] - c["discount_eur"]
    c["margin_eur"] = c["sales_eur"] - c["net_cost_eur"]
    return pd.DataFrame(issues), c


# --------------------------------------------------------------------------- daily cockpit
def daily_table(c: pd.DataFrame, by: str | None = None) -> pd.DataFrame:
    keys = ["booking_date"] + ([by] if by else [])
    g = c.groupby(keys)
    d = pd.DataFrame({"bookings": g.size(), "sales_eur": g["sales_eur"].sum(), "margin_eur": g["margin_eur"].sum(),
                      "room_nights": g["room_nights"].sum()})
    canc = c[c["status"] == "Cancelled"].groupby(["cancel_date"] + ([by] if by else [])).size()
    canc.index = canc.index.set_names(keys)
    d["cancellations"] = canc.reindex(d.index).fillna(0).astype(int)
    d = d.reset_index()
    d["margin_rate"] = d["margin_eur"] / d["sales_eur"]
    d["avg_booking_value"] = d["sales_eur"] / d["bookings"]
    return d


def add_comparisons(d: pd.DataFrame, by: str | None = None) -> pd.DataFrame:
    """D-1, same weekday last year (364 days), 7-day average, and the alert band."""
    d = d.sort_values(([by] if by else []) + ["booking_date"]).copy()
    groups = d.groupby(by) if by else [(None, d)]
    out = []
    for _, g in groups:
        g = g.set_index("booking_date").asfreq("D") if by is None else g.set_index("booking_date")
        full = g.reindex(pd.date_range(g.index.min(), g.index.max()))
        if by:
            full[by] = full[by].ffill()
        # a day without any booking is a real zero (e.g. a partner who stopped), not missing data
        full[["bookings", "sales_eur", "margin_eur", "room_nights", "cancellations"]] = \
            full[["bookings", "sales_eur", "margin_eur", "room_nights", "cancellations"]].fillna(0)
        for m in ["bookings", "sales_eur", "margin_rate", "cancellations"]:
            s = full[m]
            full[f"{m}_d1"] = s.shift(1)
            full[f"{m}_ly"] = s.shift(364)
            full[f"{m}_avg7"] = s.shift(1).rolling(7, min_periods=5).mean()
            # Expected value: average of the same weekday over the previous 4 weeks
            # (captures the weekly pattern and follows the seasonal ramp).
            same_wd = pd.concat([s.shift(7 * k) for k in range(1, 5)], axis=1)
            full[f"{m}_expected"] = same_wd.mean(axis=1)
            # Normal noise: standard deviation of the daily deviation over the previous 8 weeks
            # (relative deviation for volumes, deviation in points for the margin rate).
            dev = (s / full[f"{m}_expected"] - 1) if m != "margin_rate" else (s - full[f"{m}_expected"])
            full[f"{m}_sd"] = dev.shift(1).rolling(7 * BASELINE_WEEKS, min_periods=28).std()
            full[f"{m}_z"] = dev / full[f"{m}_sd"]
        out.append(full.rename_axis("booking_date").reset_index())
    return pd.concat(out)


def alerts(c: pd.DataFrame) -> pd.DataFrame:
    dims = ["brand", "line_of_business", "client_country", "destination_country", "supplier_id", "channel"]
    top_partners = c.groupby("partner_id")["sales_eur"].sum().nlargest(25).index
    rows = []
    for dim in dims + ["partner_id"]:
        sub = c if dim != "partner_id" else c[c["partner_id"].isin(top_partners)]
        d = add_comparisons(daily_table(sub, dim), dim)
        d = d[d["booking_date"] >= d["booking_date"].min() + pd.Timedelta(days=7 * BASELINE_WEEKS)]
        min_size = MIN_SEGMENT_BOOKINGS if dim != "partner_id" else 8   # partners are smaller
        for m in ["bookings", "sales_eur", "margin_rate", "cancellations"]:
            hit = d[(d[f"{m}_z"].abs() >= Z_ALERT) & (d[f"{m}_expected"].notna())]
            # ignore small segments, and margin moves below 1.5 points
            hit = hit[hit["bookings_expected"] >= min_size]
            if m == "margin_rate":
                hit = hit[(hit[m] - hit[f"{m}_expected"]).abs() >= 0.015]
            if m == "cancellations":   # small counts are too noisy to alert on
                hit = hit[hit["cancellations_expected"] >= 8]
            for _, r in hit.iterrows():
                rows.append(dict(date=r["booking_date"].date(), dimension=dim, member=r[dim], metric=m,
                                 actual=r[m], expected=r[f"{m}_expected"], z=r[f"{m}_z"]))
    out = pd.DataFrame(rows).sort_values(["date", "dimension"])
    out["matched_event"] = ""
    for name, d0, d1, members in EVENT_KEYS:
        hit = (pd.to_datetime(out["date"]).between(d0, d1)) & out["member"].isin(members)
        out.loc[hit, "matched_event"] = name
    return out


# --------------------------------------------------------------------------- variance bridge
def margin_bridge(c: pd.DataFrame, p0: str, p1: str, seg: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Margin change between two months split into volume, mix, price, FX and cost.

    Per segment s (N = bookings, P = sales per booking in EUR, C = cost per booking, m = P - C):
      Volume = (N1 - N0) * m0_avg
      Mix    = N1 * sum((w1 - w0) * (m0_s - m0_avg))
      Price  = N1 * sum(w1 * (P1_s at last year's FX - P0_s))    local price change
      FX     = N1 * sum(w1 * (P1_s - P1_s at last year's FX))    currency translation
      Cost   = -N1 * sum(w1 * (C1_s - C0_s))
    The five effects add up exactly to M1 - M0.
    """
    c = c.assign(ym=c["booking_date"].dt.strftime("%Y-%m"),
                 sales_local=c["sell_price_local"] - c["discount_eur"] * c["fx_rate_to_eur"])
    x = c[c["ym"].isin([p0, p1])]
    g = x.groupby(seg + ["ym"]).agg(N=("booking_id", "size"), S=("sales_eur", "sum"), L=("sales_local", "sum"),
                                     C=("net_cost_eur", "sum")).reset_index()
    w = g.pivot_table(index=seg, columns="ym", values=["N", "S", "L", "C"], aggfunc="sum")
    w.columns = [f"{a}{'0' if b == p0 else '1'}" for a, b in w.columns]
    w = w.dropna()     # segments present in both months (others are reported apart)
    N0, N1 = w["N0"].sum(), w["N1"].sum()
    P0, P1, C0, C1 = w["S0"] / w["N0"], w["S1"] / w["N1"], w["C0"] / w["N0"], w["C1"] / w["N1"]
    fx0 = w["L0"] / w["S0"]                       # average local units per EUR in month 0
    P1_at_fx0 = (w["L1"] / w["N1"]) / fx0
    m0 = P0 - C0
    m0_avg = (w["S0"].sum() - w["C0"].sum()) / N0
    w1, w0 = w["N1"] / N1, w["N0"] / N0
    w["volume"] = (w["N1"] - w["N0"]) * m0_avg
    w["mix"] = N1 * (w1 - w0) * (m0 - m0_avg)
    w["price"] = N1 * w1 * (P1_at_fx0 - P0)
    w["fx"] = N1 * w1 * (P1 - P1_at_fx0)
    w["cost"] = -N1 * w1 * (C1 - C0)
    M0, M1 = w["S0"].sum() - w["C0"].sum(), w["S1"].sum() - w["C1"].sum()
    bridge = pd.DataFrame({"step": [f"Margin {p0}", "Volume", "Mix", "Price", "FX", "Cost", f"Margin {p1}"],
                           "eur": [M0, w["volume"].sum(), w["mix"].sum(), w["price"].sum(), w["fx"].sum(),
                                   w["cost"].sum(), M1]})
    assert abs(M0 + w[["volume", "mix", "price", "fx", "cost"]].sum().sum() - M1) < 1e-6
    return bridge, w.reset_index()


# --------------------------------------------------------------------------- incidents
def window(c, start, end, **filters):
    m = c["booking_date"].between(pd.Timestamp(start), pd.Timestamp(end))
    for k, v in filters.items():
        m &= c[k] == v
    return c[m]


def incident_numbers(c: pd.DataFrame, pm: pd.DataFrame) -> pd.DataFrame:
    rows = []
    end = c["booking_date"].max()

    # E1 supplier net-rate increase (Bedbank C, Spain)
    before = window(c, "2026-05-18", "2026-06-14", supplier_id="SUP-BB-C", destination_country="ES")
    after = window(c, "2026-06-15", end, supplier_id="SUP-BB-C", destination_country="ES")
    mr0 = before["margin_eur"].sum() / before["sales_eur"].sum()
    mr1 = after["margin_eur"].sum() / after["sales_eur"].sum()
    vol0 = len(before) / 28
    vol1 = len(after) / ((end - pd.Timestamp("2026-06-15")).days + 1)
    lost = (mr0 - mr1) * after["sales_eur"].sum()
    rows.append(dict(event="E1", title="Supplier net-rate increase",
                     what="Bedbank C raised its net rates in Spain from 15 Jun 2026; selling prices did not move",
                     evidence=f"Margin rate on Bedbank C x Spain {mr0:.1%} -> {mr1:.1%}; bookings/day {vol0:.1f} -> {vol1:.1f}",
                     impact_eur=-lost,
                     recommendation="Re-route Spain demand to Direct contracts / Bedbank B where cheaper; "
                                    "renegotiate or raise sell rates +5% on affected hotels",
                     owner="Supply / contracting", follow_up="Daily margin rate on Bedbank C x Spain back above 7%"))

    # E2 API partner lost
    p_before = window(c, "2026-08-10", "2026-09-06", partner_id="API-002")
    p_after = window(c, "2026-09-15", end, partner_id="API-002")
    api_b = window(c, "2026-08-10", "2026-09-06", channel="API")
    api_a = window(c, "2026-09-15", end, channel="API")
    days_after = (end - pd.Timestamp("2026-09-15")).days + 1
    lost_sales_day = p_before["sales_eur"].sum() / 28 - p_after["sales_eur"].sum() / days_after
    lost_margin_day = p_before["margin_eur"].sum() / 28 - p_after["margin_eur"].sum() / days_after
    api_drop = len(api_b) / 28 - len(api_a) / days_after
    share = (len(p_before) / 28 - len(p_after) / days_after) / api_drop if api_drop > 0 else np.nan
    rows.append(dict(event="E2", title="Large API partner moved to a competitor",
                     what="API-002 (top partner) ramped down from 7 Sep and stopped on 14 Sep 2026",
                     evidence=f"API-002 bookings/day {len(p_before) / 28:.0f} -> {len(p_after) / days_after:.0f}; "
                              f"it explains {share:.0%} of the API channel drop",
                     impact_eur=-lost_margin_day * days_after,
                     recommendation="Account review with the partner: compare our rates on its top 50 hotels, "
                                    "offer a targeted tier rather than a blanket discount",
                     owner="Key account manager + pricing", follow_up="Weekly volume of API-002; win-back offer within 2 weeks"))

    # E3 FX: Turkey (TRY), Sep-Oct 2026 vs Sep-Oct 2025
    t1 = window(c, "2026-09-01", "2026-10-31", client_country="TR")
    t0 = window(c, "2025-09-01", "2025-10-31", client_country="TR")
    loc_growth = t1["sell_price_local"].sum() / t0["sell_price_local"].sum() - 1
    eur_growth = t1["sales_eur"].sum() / t0["sales_eur"].sum() - 1
    vol_growth = len(t1) / len(t0) - 1
    # Local price lists are set at the rate of the 1st of the previous month: when the currency
    # falls, the euro value collected is lower than the list was built for (the "lag cost").
    fx = pd.read_csv(SIM / "fx_rates.csv", parse_dates=["date"]).set_index(["date", "currency"])["units_per_eur"]
    tr = c[c["client_country"] == "TR"]
    list_date = (tr["booking_date"].dt.to_period("M") - 1).dt.to_timestamp()
    fx_list = fx.reindex(pd.MultiIndex.from_arrays([list_date, tr["currency"]])).to_numpy()
    lag_cost = (tr["sell_price_local"] / fx_list - tr["sell_price_eur"]).sum()
    fxm = fx.xs("TRY", level="currency")
    try_fall = fxm["2025-10"].mean() / fxm["2026-10"].mean() - 1      # value of 1 TRY in EUR, YoY
    mr_tr = tr["margin_eur"].sum() / tr["sales_eur"].sum()
    eur_mk = c[c["currency"] == "EUR"]
    mr_eur = eur_mk["margin_eur"].sum() / eur_mk["sales_eur"].sum()
    rows.append(dict(event="E3", title="Currency effect (real TRY depreciation, ECB rates)",
                     what="Turkish partners pay in TRY; the lira lost value against the euro (ECB reference rates)",
                     evidence=f"TRY {try_fall:+.0%} vs EUR (Oct 26 vs Oct 25). Sep-Oct YoY: bookings {vol_growth:+.0%}, "
                              f"sales in TRY {loc_growth:+.0%}, in EUR {eur_growth:+.0%}; margin rate {mr_tr:.1%} "
                              f"vs {mr_eur:.1%} in euro markets",
                     impact_eur=-lag_cost,
                     recommendation="Refresh TRY price lists weekly (not monthly) or quote in EUR/USD; "
                                    "report the market at constant FX",
                     owner="Pricing + finance (treasury)", follow_up="Constant-FX sales line in the daily report"))

    # E4 holiday in a source market (UAE)
    h1 = window(c, "2026-05-11", "2026-05-25", client_country="AE")
    n1 = window(c, "2026-04-13", "2026-05-10", client_country="AE")
    h0 = window(c, "2025-05-22", "2025-06-05", client_country="AE")
    n0 = window(c, "2025-04-24", "2025-05-21", client_country="AE") if c["booking_date"].min() <= pd.Timestamp("2025-04-24") \
        else window(c, "2025-05-01", "2025-05-21", client_country="AE")
    up1 = (len(h1) / 15) / (len(n1) / 28) - 1
    up0 = (len(h0) / 15) / (len(n0) / max((n0["booking_date"].max() - n0["booking_date"].min()).days + 1, 1)) - 1
    rows.append(dict(event="E4", title="Source-market holiday (UAE)",
                     what="Bookings from the UAE jump before the Eid al-Adha holiday (late May 2026)",
                     evidence=f"Bookings/day +{up1:.0%} vs the previous 4 weeks; same holiday in 2025 (early June): +{up0:.0%}",
                     impact_eur=h1["margin_eur"].sum() - (len(n1) / 28) * 15 * (n1["margin_eur"].sum() / len(n1)),
                     recommendation="Normal seasonality, not an anomaly: align comparisons on the holiday calendar; "
                                    "secure allotments in Turkey/UAE destinations 6 weeks ahead",
                     owner="Commercial team (market manager UAE)", follow_up="Holiday calendar added to the alert baseline"))

    # E6 cancellations (Greece)
    ce = c[(c["status"] == "Cancelled") & (c["destination_country"] == "EL")]
    spike = ce[ce["cancel_date"].between("2026-08-08", "2026-08-14")]
    base = ce[ce["cancel_date"].between("2026-07-11", "2026-08-07")]
    rows.append(dict(event="E6", title="Cancellation spike at one destination (Greece)",
                     what="A local event (e.g. severe weather) triggered cancellations for mid-August stays",
                     evidence=f"Cancellations/day for Greece {len(base) / 28:.0f} -> {len(spike) / 7:.0f} (8-14 Aug 2026)",
                     impact_eur=-(spike["margin_eur"].sum() - len(base) / 28 * 7 * base["margin_eur"].mean()),
                     recommendation="Offer re-booking to other islands/dates before refund; check supplier "
                                    "cancellation terms (who bears the cost); no price action",
                     owner="Customer operations + supply", follow_up="Re-booking rate of cancelled Greek stays"))

    # E7 price-match requests
    granted = pm[pm["decision"] == "Discount"]
    saved = pm[pm["decision"].isin(["Discount", "Switch supplier"]) & (pm["booked"] == "Yes")]
    rows.append(dict(event="E7", title="Price-match requests",
                     what="Partners report lower competitor prices; each request is checked like-for-like",
                     evidence=f"{len(pm):,} requests; {len(granted) / len(pm):.0%} discounted, "
                              f"{(pm['decision'] == 'Refuse').mean():.0%} refused with arguments",
                     impact_eur=-granted[granted["booked"] == "Yes"]["discount_eur"].sum(),
                     recommendation="Keep the margin floor; publish the checklist to partners to cut invalid requests",
                     owner="Pricing analyst", follow_up=f"Sales saved EUR {saved['etg_price_eur'].sum():,.0f}"))
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- price match
def price_match_kpis(pm: pd.DataFrame) -> dict[str, pd.DataFrame]:
    pm = pm.copy()
    by_dec = pm.groupby("decision").agg(requests=("request_id", "size"), avg_handling_h=("handling_hours", "mean"),
                                        booked_rate=("booked", lambda s: (s == "Yes").mean()),
                                        discount_cost_eur=("discount_eur", "sum")).reset_index()
    by_dec["share"] = by_dec["requests"] / by_dec["requests"].sum()
    by_case = pd.crosstab(pm["case_type_truth"], pm["decision"])
    pm["month"] = pd.to_datetime(pm["request_date"]).dt.strftime("%Y-%m")
    monthly = pm.groupby("month").agg(requests=("request_id", "size"),
                                      discounts=("decision", lambda s: (s == "Discount").sum()),
                                      discount_cost_eur=("discount_eur", "sum")).reset_index()
    pm["gap_pct"] = 1 - pm["competitor_price_eur"] / pm["etg_price_eur"]
    pm["saved_sales_eur"] = np.where(pm["decision"].isin(["Discount", "Switch supplier"]) & (pm["booked"] == "Yes"),
                                     pm["etg_price_eur"] - pm["discount_eur"], 0)
    summary = pd.DataFrame([dict(
        requests=len(pm), valid_like_for_like=(pm["valid_gap_eur"] > 0).sum(),
        valid_share=(pm["valid_gap_eur"] > 0).mean(), discount_share=(pm["decision"] == "Discount").mean(),
        refuse_share=(pm["decision"] == "Refuse").mean(), switch_share=(pm["decision"] == "Switch supplier").mean(),
        ask_share=(pm["decision"] == "Ask for details").mean(),
        discount_cost_eur=pm.loc[(pm["decision"] == "Discount") & (pm["booked"] == "Yes"), "discount_eur"].sum(),
        saved_sales_eur=pm["saved_sales_eur"].sum(), median_handling_h=pm["handling_hours"].median(),
        avg_claimed_gap=pm["gap_pct"].mean())])
    return {"pm_by_decision": by_dec, "pm_by_case": by_case.reset_index(), "pm_monthly": monthly,
            "pm_summary": summary}


# --------------------------------------------------------------------------- pricing
def pricing_recommendations(c: pd.DataFrame, pm: pd.DataFrame) -> pd.DataFrame:
    # B2B partners only (the price-match signal does not exist for B2C)
    h = c[(c["line_of_business"] == "Hotels") & c["partner_type"].isin(["Agency", "API partner", "Corporate"])]
    last = h[h["booking_date"] >= h["booking_date"].max() - pd.Timedelta(days=89)]
    ly = h[h["booking_date"].between(h["booking_date"].max() - pd.Timedelta(days=364 + 89),
                                     h["booking_date"].max() - pd.Timedelta(days=364))]
    seg = ["partner_type", "destination_country"]
    g = last.groupby(seg).agg(bookings=("booking_id", "size"), sales_eur=("sales_eur", "sum"),
                              margin_eur=("margin_eur", "sum"), cost_eur=("net_cost_eur", "sum"))
    g["bookings_ly"] = ly.groupby(seg).size().reindex(g.index)
    g["yoy_volume"] = g["bookings"] / g["bookings_ly"] - 1
    g["markup"] = g["sales_eur"] / g["cost_eur"] - 1
    g["margin_rate"] = g["margin_eur"] / g["sales_eur"]
    pm2 = pm.merge(c[["booking_id", "partner_type"]].drop_duplicates(), on="booking_id", how="left")
    pm2 = pm2[pd.to_datetime(pm2["request_date"]) >= last["booking_date"].min()]
    req = pm2.groupby(["partner_type", "destination_country"]).agg(
        requests=("request_id", "size"), valid=("valid_gap_eur", lambda s: (s > 0).sum()))
    g = g.join(req).fillna({"requests": 0, "valid": 0})
    g["valid_requests_per_1000"] = g["valid"] / g["bookings"] * 1000
    g = g[g["bookings"] >= 300].reset_index()
    pressure_hi = g["valid_requests_per_1000"].quantile(0.75)
    pressure_lo = g["valid_requests_per_1000"].quantile(0.40)

    def rule(r):
        if r["margin_rate"] < MARGIN_FLOOR + 0.01:
            return 0.0, "Hold price; margin close to floor -> renegotiate net rates"
        if r["valid_requests_per_1000"] >= pressure_hi and r["yoy_volume"] < 0.10:
            return -0.005, "Price pressure: -0.5 pt markup on this segment, or re-source cheaper supply"
        if r["valid_requests_per_1000"] <= pressure_lo and r["yoy_volume"] >= 0.10:
            return 0.005, "Low price pressure and strong growth: test +0.5 pt markup"
        return 0.0, "Keep current markup; monitor"
    rr = g.apply(lambda r: pd.Series(rule(r), index=["markup_change", "recommendation"]), axis=1)
    g = pd.concat([g, rr], axis=1)
    # Expected effect over 90 days, for a prudent elasticity range.
    for name, eps in zip(["low", "high"], ELASTICITY_RANGE):
        dp = g["markup_change"] / (1 + g["markup"])            # % change in selling price
        dv = eps * dp                                         # % change in volume
        new_sales = g["sales_eur"] * (1 + dp) * (1 + dv)
        new_cost = g["cost_eur"] * (1 + dv)
        g[f"margin_effect_eur_{name}"] = (new_sales - new_cost) - g["margin_eur"]
        g[f"volume_effect_pct_{name}"] = dv
    return g.sort_values(["partner_type", "destination_country"])


# --------------------------------------------------------------------------- main
def main() -> None:
    OUT_SIM.mkdir(parents=True, exist_ok=True)
    raw = load()
    issues, c = dq_checks(raw)
    issues.to_csv(OUT_SIM / "dq_issues.csv", index=False)
    c.to_csv(OUT_SIM / "bookings_clean.csv", index=False, date_format="%Y-%m-%d")

    # Data-quality effect on the daily total (raw vs clean) for page 7
    raw2 = raw.assign(sales_eur=raw["sell_price_eur"] - raw["discount_eur"],
                      margin_eur=raw["sell_price_eur"] - raw["discount_eur"] - raw["net_cost_eur"].fillna(0))
    dr = daily_table(raw2).set_index("booking_date")[["bookings", "sales_eur", "margin_rate"]]
    dc = daily_table(c).set_index("booking_date")[["bookings", "sales_eur", "margin_rate"]]
    dr.add_suffix("_raw").join(dc.add_suffix("_clean")).reset_index().to_csv(OUT_SIM / "daily_raw_vs_clean.csv", index=False)

    daily = add_comparisons(daily_table(c))
    daily.to_csv(OUT_SIM / "daily_total.csv", index=False)
    for dim in ["brand", "line_of_business", "client_country", "destination_country", "supplier_id",
                "partner_type", "channel"]:
        daily_table(c, dim).to_csv(OUT_SIM / f"daily_by_{dim}.csv", index=False)
    al = alerts(c)
    al.to_csv(OUT_SIM / "alerts.csv", index=False)

    bridge, seg = margin_bridge(c, "2025-09", "2026-09", ["brand", "line_of_business", "client_country"])
    bridge.to_csv(OUT_SIM / "margin_bridge_sep.csv", index=False)
    seg.to_csv(OUT_SIM / "margin_bridge_sep_segments.csv", index=False)

    pm = pd.read_csv(SIM / "price_match_requests.csv")
    for k, v in price_match_kpis(pm).items():
        v.to_csv(OUT_SIM / f"{k}.csv", index=False)
    inc = incident_numbers(c, pm)
    inc.to_csv(OUT_SIM / "incidents.csv", index=False)
    rec = pricing_recommendations(c, pm)
    rec.to_csv(OUT_SIM / "pricing_recommendations.csv", index=False)

    print(issues[["issue", "date_from", "scope", "rows", "effect"]].to_string(index=False))
    print(bridge.round(0).to_string(index=False))
    print(inc[["event", "evidence", "impact_eur"]].to_string(index=False))
    print("alerts:", len(al))


if __name__ == "__main__":
    main()
