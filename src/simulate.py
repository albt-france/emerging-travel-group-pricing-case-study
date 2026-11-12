"""SIMULATED sales ledger for a multi-brand B2B/B2C travel platform.

Everything produced here is SIMULATED. It is calibrated on a few public orders of magnitude
(see data/ASSUMPTIONS.md) but it is NOT Emerging Travel Group data: no real sales, margins,
suppliers or partners. Supplier, partner and competitor names are fictional codes.

Real inputs used:
- ECB daily FX reference rates (data/public/ecb_fx_daily.csv) to convert local prices to EUR;
- Eurostat seasonality profiles (data/public/eurostat_*.csv) to shape demand by month.

Period: 2025-05-01 -> 2026-10-31 (18 months, daily). Fixed random seed 42.
Seven events are programmed so that the daily monitoring has something to find:
  E1 supplier net-rate increase, E2 API partner lost, E3 FX (real lira depreciation),
  E4 source-market holiday, E5 data-quality issues, E6 cancellation spike, E7 price-match requests.

Outputs (data/simulated/): bookings.csv, suppliers.csv, partners.csv, price_match_requests.csv,
fx_rates.csv, calendar.csv, events_truth.csv (the answer key, used only to check detection).
"""
from __future__ import annotations

import numpy as np
import pandas as pd

from common import PUBLIC, SIM

SEED = 42
START, END = pd.Timestamp("2025-05-01"), pd.Timestamp("2026-10-31")
EXTRACT_DATE = END                    # status of bookings as seen on this date
BASE_BOOKINGS_PER_DAY = 300           # ~1:50 of the public "130k+ bookings weekly"
ANNUAL_GROWTH = 0.18                  # assumption, consistent with growing public figures
MARGIN_FLOOR = 0.04                   # minimum margin rate on selling price (price-match rule)

rng = np.random.default_rng(SEED)

# --------------------------------------------------------------------------- dimensions
MARKETS = pd.DataFrame([
    # code, name, currency, share of bookings, profile used for booking seasonality
    ("DE", "Germany", "EUR", 0.13, "DE"), ("GB", "United Kingdom", "GBP", 0.09, "NL"),
    ("FR", "France", "EUR", 0.08, "FR"), ("IT", "Italy", "EUR", 0.08, "IT"),
    ("ES", "Spain", "EUR", 0.07, "ES"), ("NL", "Netherlands", "EUR", 0.04, "NL"),
    ("PL", "Poland", "PLN", 0.05, "PL"), ("TR", "Turkey", "TRY", 0.08, "TR"),
    ("AE", "United Arab Emirates", "USD", 0.09, "AE"), ("US", "United States", "USD", 0.07, "US"),
    ("BR", "Brazil", "BRL", 0.06, "BR"), ("MX", "Mexico", "MXN", 0.05, "MX"),
    ("IN", "India", "INR", 0.06, "IN"),
], columns=["code", "name", "currency", "share", "profile"])
MARKETS["share"] /= MARKETS["share"].sum()

DESTS = pd.DataFrame([
    # code, name, base net rate per room-night (EUR), seasonality profile
    ("ES", "Spain", 95, "ES"), ("IT", "Italy", 115, "IT"), ("FR", "France", 120, "FR"),
    ("EL", "Greece", 105, "EL"), ("PT", "Portugal", 85, "PT"), ("DE", "Germany", 90, "DE"),
    ("HR", "Croatia", 95, "HR"), ("TR", "Turkey", 70, "TRD"), ("AE", "United Arab Emirates", 125, "AED"),
    ("TH", "Thailand", 50, "THD"),
], columns=["code", "name", "base_rate", "profile"])

# Non-EU seasonality profiles (12 monthly weights, assumption) where Eurostat has no series.
GENERIC_PROFILES = {
    "TR": [5, 5, 6, 7, 9, 12, 14, 14, 10, 8, 5, 5], "AE": [8, 7, 8, 8, 9, 10, 11, 11, 7, 7, 6, 8],
    "US": [6, 6, 8, 8, 9, 11, 12, 11, 8, 8, 6, 7], "BR": [12, 9, 7, 6, 6, 7, 11, 8, 7, 7, 8, 12],
    "MX": [8, 6, 9, 9, 7, 8, 12, 10, 6, 7, 8, 10], "IN": [7, 6, 7, 9, 12, 11, 7, 6, 7, 10, 9, 9],
    "TRD": [3, 3, 5, 7, 10, 13, 16, 16, 12, 8, 4, 3], "AED": [12, 11, 11, 9, 6, 5, 5, 5, 6, 9, 10, 11],
    "THD": [12, 11, 10, 9, 7, 6, 7, 7, 6, 7, 8, 10],
}

# Preference of each source market for each destination (assumption).
DEST_PREF = {
    "DE": dict(ES=20, IT=14, FR=6, EL=12, PT=6, DE=14, HR=8, TR=14, AE=4, TH=2),
    "GB": dict(ES=24, IT=10, FR=10, EL=12, PT=12, DE=4, HR=5, TR=12, AE=8, TH=3),
    "FR": dict(ES=18, IT=16, FR=30, EL=10, PT=8, DE=4, HR=4, TR=4, AE=4, TH=2),
    "IT": dict(ES=16, IT=36, FR=12, EL=12, PT=5, DE=6, HR=6, TR=3, AE=3, TH=1),
    "ES": dict(ES=40, IT=16, FR=14, EL=8, PT=12, DE=4, HR=2, TR=2, AE=1, TH=1),
    "NL": dict(ES=18, IT=12, FR=16, EL=12, PT=8, DE=14, HR=6, TR=10, AE=3, TH=1),
    "PL": dict(ES=14, IT=12, FR=6, EL=16, PT=4, DE=12, HR=14, TR=18, AE=3, TH=1),
    "TR": dict(ES=8, IT=10, FR=8, EL=10, PT=2, DE=12, HR=2, TR=40, AE=7, TH=1),
    "AE": dict(ES=8, IT=10, FR=10, EL=6, PT=2, DE=8, HR=2, TR=24, AE=26, TH=4),
    "US": dict(ES=16, IT=24, FR=22, EL=12, PT=8, DE=8, HR=3, TR=3, AE=3, TH=1),
    "BR": dict(ES=16, IT=22, FR=20, EL=6, PT=24, DE=6, HR=1, TR=2, AE=2, TH=1),
    "MX": dict(ES=34, IT=20, FR=20, EL=6, PT=8, DE=6, HR=1, TR=2, AE=2, TH=1),
    "IN": dict(ES=8, IT=10, FR=12, EL=4, PT=2, DE=10, HR=2, TR=12, AE=30, TH=10),
}

PARTNER_TYPES = pd.DataFrame([
    # type, brand, channel, share of bookings, markup on net (hotels), cancel base, lead-time mean (days)
    ("Agency", "RateHawk", "Website", 0.45, 0.095, 0.20, 38),
    ("API partner", "RateHawk", "API", 0.30, 0.065, 0.25, 30),
    ("White label", "RateHawk", "White label", 0.05, 0.080, 0.22, 34),
    ("Corporate", "Roundtrip", "Website", 0.06, 0.100, 0.15, 14),
    ("B2C", "ZenHotels", "Website", 0.14, 0.140, 0.24, 26),
], columns=["type", "brand", "channel", "share", "markup", "cancel_base", "lead_mean"])

LOB_BY_BRAND = {   # line-of-business mix per brand (assumption)
    "RateHawk": {"Hotels": 0.85, "Transfers": 0.06, "Car rentals": 0.03, "Flights": 0.06},
    "ZenHotels": {"Hotels": 1.0},
    "Roundtrip": {"Hotels": 0.70, "Flights": 0.25, "Transfers": 0.03, "Car rentals": 0.02},
}
LOB_PRICING = {   # net cost per booking (lognormal median, EUR) and markup for non-hotel lines
    "Transfers": (45, 0.20), "Car rentals": (180, 0.12), "Flights": (280, 0.035),
}

SUPPLIERS = pd.DataFrame(
    [(f"SUP-D-{d}", f"Direct contracts {d}", "Direct", d, 30, 0.96, "Hotels") for d in DESTS["code"]]
    + [(f"SUP-BB-{c}", f"Bedbank {c}", "Bedbank", "Global", t, f, "Hotels")
       for c, t, f in [("A", 15, 1.00), ("B", 15, 0.99), ("C", 20, 1.00), ("D", 30, 1.01),
                       ("E", 15, 1.02), ("F", 30, 0.995)]]
    + [(f"SUP-WS-{c}", f"Wholesaler {c}", "Wholesaler", r, t, f, "Hotels")
       for c, r, t, f in [("A", "Southern Europe", 45, 1.01), ("B", "Middle East", 30, 1.00),
                          ("C", "Asia", 30, 1.00), ("D", "Western Europe", 45, 1.02),
                          ("E", "Eastern Mediterranean", 30, 0.99)]]
    + [("SUP-TR-A", "Transfer supplier A", "Ground", "Global", 30, 1.0, "Transfers"),
       ("SUP-TR-B", "Transfer supplier B", "Ground", "Global", 30, 1.0, "Transfers"),
       ("SUP-CR-A", "Car rental supplier A", "Ground", "Global", 30, 1.0, "Car rentals"),
       ("SUP-CR-B", "Car rental supplier B", "Ground", "Global", 30, 1.0, "Car rentals"),
       ("SUP-FL-A", "Flight consolidator A", "Air", "Global", 7, 1.0, "Flights"),
       ("SUP-FL-B", "Flight consolidator B", "Air", "Global", 7, 1.0, "Flights")],
    columns=["supplier_id", "supplier_name", "supplier_type", "coverage", "payment_terms_days",
             "cost_factor", "line_of_business"])
WHOLESALER_DESTS = {"SUP-WS-A": ["ES", "IT", "PT", "HR"], "SUP-WS-B": ["AE", "TR"],
                    "SUP-WS-C": ["TH"], "SUP-WS-D": ["FR", "DE"], "SUP-WS-E": ["EL", "TR"]}

# --------------------------------------------------------------------------- events
EV = {
    "E1_supplier": dict(supplier="SUP-BB-C", dest="ES", start="2026-06-15", net_increase=0.07),
    "E2_partner_lost": dict(ramp_start="2026-09-07", end="2026-09-14"),
    "E4_holiday": dict(market="AE", windows=[("2025-05-22", "2025-06-05"), ("2026-05-11", "2026-05-25")],
                       uplift=0.7),
    "E5_duplicates": dict(date="2026-07-21", channel="API"),
    "E5_missing_cost": dict(supplier="SUP-WS-A", start="2026-08-03", end="2026-08-09"),
    "E5_timezone": dict(channel="API", start="2026-10-12", end="2026-10-18", shift_hours=5),
    "E6_cancellations": dict(dest="EL", checkin_start="2026-08-12", checkin_end="2026-08-25",
                             cancel_start="2026-08-08", cancel_end="2026-08-14", extra_rate=0.30),
}


def load_profiles() -> dict[str, np.ndarray]:
    """Monthly weight profiles (index 0 = January), mostly from Eurostat."""
    prof = {k: np.array(v, float) for k, v in GENERIC_PROFILES.items()}
    s = pd.read_csv(PUBLIC / "eurostat_market_seasonality.csv")
    s = s[s["year"] == 2024].pivot_table(index="month", columns="market_code", values="nights")
    for col in s.columns:
        if s[col].notna().all():
            prof[col] = s[col].to_numpy()
    d = pd.read_csv(PUBLIC / "eurostat_monthly_nights_by_destination.csv")
    d = d[(d["residence"] == "FOR") & d["month"].str.startswith("2025")]
    d["m"] = d["month"].str[5:].astype(int)
    d = d.pivot_table(index="m", columns="destination_code", values="nights")
    for col in d.columns:
        prof[col] = d[col].to_numpy()
    return {k: v / v.mean() for k, v in prof.items()}


def booking_profile(stay_profile: np.ndarray) -> np.ndarray:
    """Bookings are made ~1 month before the stay and are smoother than stays."""
    shifted = np.roll(stay_profile, -1)
    return 0.5 + 0.5 * shifted


def make_partners() -> pd.DataFrame:
    rows = []
    for m in MARKETS["code"]:
        for i in range(1, 151):
            rows.append((f"AG-{m}-{i:04d}", "Agency", m))
        for i in range(1, 11):
            rows.append((f"CORP-{m}-{i:03d}", "Corporate", m))
        rows.append((f"B2C-{m}", "B2C", m))
    api_home = list(MARKETS["code"]) + list(rng.choice(MARKETS["code"], 17, p=MARKETS["share"]))
    for i, m in enumerate(api_home, start=1):
        rows.append((f"API-{i:03d}", "API partner", m))
    wl_home = list(rng.choice(MARKETS["code"], 15, p=MARKETS["share"]))
    for i, m in enumerate(wl_home, start=1):
        rows.append((f"WL-{i:03d}", "White label", m))
    p = pd.DataFrame(rows, columns=["partner_id", "partner_type", "country"])
    # Size of each partner inside its (market, type) cell: long-tailed, like real B2B portfolios.
    p["weight"] = rng.lognormal(0, 1.3, len(p))
    p.loc[p["partner_type"] == "B2C", "weight"] = 1.0
    # One very large API partner (the one lost in event E2).
    p.loc[p["partner_id"] == "API-002", "weight"] = 40.0
    terms = {"Agency": "Net rates + markup, prepayment or credit line",
             "API partner": "API net rates + markup, 15-day credit, monthly volume review",
             "White label": "Partner-branded site, net rates + markup, revenue share",
             "Corporate": "Corporate tariff, monthly invoicing, 30-day terms",
             "B2C": "Public selling price (ZenHotels), card payment"}
    p["commercial_terms"] = p["partner_type"].map(terms)
    p["partner_name"] = p["partner_id"].str.replace("-", " ")
    return p


def main() -> None:
    SIM.mkdir(parents=True, exist_ok=True)
    prof = load_profiles()
    partners = make_partners()
    fx = pd.read_csv(PUBLIC / "ecb_fx_daily.csv", parse_dates=["date"])
    fx_w = fx.pivot(index="date", columns="currency", values="units_per_eur")

    days = pd.date_range(START, END, freq="D")
    weekday_factor = np.array([1.12, 1.14, 1.10, 1.05, 0.96, 0.78, 0.85])

    # 1) number of bookings per day x market x partner type (Poisson around an expected value)
    cells = []
    for _, mk in MARKETS.iterrows():
        bprof = booking_profile(prof[mk["profile"]])
        for _, pt in PARTNER_TYPES.iterrows():
            lam = (BASE_BOOKINGS_PER_DAY * mk["share"] * pt["share"]
                   * (1 + ANNUAL_GROWTH) ** ((days - START).days / 365)
                   * weekday_factor[days.day_of_week] * bprof[days.month - 1])
            if mk["code"] == EV["E4_holiday"]["market"]:
                for a, b in EV["E4_holiday"]["windows"]:
                    lam = np.where((days >= a) & (days <= b), lam * (1 + EV["E4_holiday"]["uplift"]), lam)
            n = rng.poisson(lam)
            cells.append(pd.DataFrame({"booking_date": days, "client_country": mk["code"],
                                       "partner_type": pt["type"], "n": n}))
    cells = pd.concat(cells, ignore_index=True)
    # one row per booking: repeat each (day, market, type) cell n times
    b = cells.loc[cells.index.repeat(cells["n"])].drop(columns="n").reset_index(drop=True)
    N = len(b)

    # 2) partner: drawn inside its (market, type) cell, proportionally to partner weight
    b["partner_id"] = None
    for (m, t), idx in b.groupby(["client_country", "partner_type"]).groups.items():
        cand = partners[(partners["country"] == m) & (partners["partner_type"] == t)]
        if cand.empty:   # API / white label partners do not exist in every market
            cand = partners[partners["partner_type"] == t]
        w = cand["weight"].to_numpy() / cand["weight"].sum()
        b.loc[idx, "partner_id"] = rng.choice(cand["partner_id"].to_numpy(), len(idx), p=w)

    # E2: the large API partner moves to a competitor (ramp-down then stop)
    lost = b["partner_id"].eq("API-002")
    ramp = lost & b["booking_date"].between(EV["E2_partner_lost"]["ramp_start"], EV["E2_partner_lost"]["end"])
    gone = lost & (b["booking_date"] > EV["E2_partner_lost"]["end"])
    drop = gone | (ramp & (rng.random(N) < 0.6))
    b = b[~drop].reset_index(drop=True)
    N = len(b)

    pt = PARTNER_TYPES.set_index("type")
    b["brand"] = b["partner_type"].map(pt["brand"])
    b["channel"] = b["partner_type"].map(pt["channel"])
    b.loc[(b["partner_type"] == "B2C") & (rng.random(N) < 0.45), "channel"] = "Mobile app"

    # 3) line of business by brand
    b["line_of_business"] = "Hotels"
    for brand, mix in LOB_BY_BRAND.items():
        idx = b.index[b["brand"] == brand]
        b.loc[idx, "line_of_business"] = rng.choice(list(mix), len(idx), p=list(mix.values()))

    # 4) lead time and check-in date
    lead_mean = b["partner_type"].map(pt["lead_mean"]).to_numpy()
    lead = np.clip(rng.gamma(1.3, lead_mean / 1.3), 0, 330).round().astype(int)
    b["checkin_date"] = b["booking_date"] + pd.to_timedelta(lead, unit="D")

    # 5) destination: market preference x destination seasonality in the check-in month
    b["destination_country"] = None
    cmonth = b["checkin_date"].dt.month.to_numpy()
    for m in MARKETS["code"]:
        for mo in range(1, 13):
            idx = b.index[(b["client_country"] == m) & (cmonth == mo)]
            if len(idx) == 0:
                continue
            w = np.array([DEST_PREF[m][d] * prof[p][mo - 1] for d, p in zip(DESTS["code"], DESTS["profile"])])
            b.loc[idx, "destination_country"] = rng.choice(DESTS["code"], len(idx), p=w / w.sum())

    # 6) supplier
    hotel = b["line_of_business"].eq("Hotels")
    b["supplier_id"] = None
    for d in DESTS["code"]:
        idx = b.index[hotel & (b["destination_country"] == d)]
        cands = [f"SUP-D-{d}"] + [s for s in SUPPLIERS["supplier_id"] if s.startswith("SUP-BB")] \
            + [s for s, ds in WHOLESALER_DESTS.items() if d in ds]
        w = np.array([3.0] + [1.0, 1.4, 1.8, 0.8, 0.6, 0.9] + [1.2] * (len(cands) - 7))
        b.loc[idx, "supplier_id"] = rng.choice(cands, len(idx), p=w / w.sum())
    for lob, sups in [("Transfers", ["SUP-TR-A", "SUP-TR-B"]), ("Car rentals", ["SUP-CR-A", "SUP-CR-B"]),
                      ("Flights", ["SUP-FL-A", "SUP-FL-B"])]:
        idx = b.index[b["line_of_business"] == lob]
        b.loc[idx, "supplier_id"] = rng.choice(sups, len(idx), p=[0.6, 0.4])

    # 7) nights, net cost and selling price
    nights = np.where(hotel, 1 + rng.poisson(2.2, N), 0)
    nights = np.where(hotel & b["partner_type"].eq("Corporate"), 1 + rng.poisson(1.0, N), nights)
    b["room_nights"] = nights
    dest = DESTS.set_index("code")
    season_matrix = np.vstack([prof[p] for p in DESTS["profile"]])          # destination x month
    dest_idx = b["destination_country"].map({c: i for i, c in enumerate(DESTS["code"])}).to_numpy()
    season = season_matrix[dest_idx, cmonth - 1]
    base = b["destination_country"].map(dest["base_rate"]).to_numpy()
    cost_factor = b["supplier_id"].map(SUPPLIERS.set_index("supplier_id")["cost_factor"]).to_numpy()
    hotel_cost = base * (0.75 + 0.25 * season) * cost_factor * rng.lognormal(0, 0.25, N) * nights
    other_cost = np.zeros(N)
    other_markup = np.zeros(N)
    for lob, (median, mu) in LOB_PRICING.items():
        idx = (b["line_of_business"] == lob).to_numpy()
        other_cost[idx] = median * rng.lognormal(0, 0.35, idx.sum())
        other_markup[idx] = mu
    net_cost = np.where(hotel, hotel_cost, other_cost)
    markup = np.where(hotel, b["partner_type"].map(pt["markup"]).to_numpy(), other_markup)
    markup = markup + rng.normal(0, 0.012, N)
    sell_eur_list = net_cost * (1 + markup)       # price list in EUR, based on the usual net rate

    # E1: a bedbank raises its net rates in Spain; selling prices are set by the market and stay put
    e1 = EV["E1_supplier"]
    e1_mask = ((b["supplier_id"] == e1["supplier"]) & (b["destination_country"] == e1["dest"])
               & (b["booking_date"] >= e1["start"])).to_numpy()
    net_cost = np.where(e1_mask, net_cost * (1 + e1["net_increase"]), net_cost)

    # 8) currency: local price lists are refreshed monthly at the FX rate of the 1st of the previous
    #    month; the EUR value of a sale uses the ECB rate of the booking day (real data).
    b["currency"] = b["client_country"].map(MARKETS.set_index("code")["currency"])
    list_date = (b["booking_date"].dt.to_period("M") - 1).dt.to_timestamp()
    fx_long = fx_w.stack()                       # (date, currency) -> units per EUR
    fx_list = fx_long.reindex(pd.MultiIndex.from_arrays([list_date, b["currency"]])).to_numpy()
    fx_day = fx_long.reindex(pd.MultiIndex.from_arrays([b["booking_date"], b["currency"]])).to_numpy()
    b["sell_price_local"] = np.round(sell_eur_list * fx_list, 2)
    b["fx_rate_to_eur"] = fx_day
    b["sell_price_eur"] = np.round(b["sell_price_local"] / fx_day, 2)
    b["net_cost_eur"] = np.round(net_cost, 2)
    b["discount_eur"] = 0.0

    # 9) booking time (UTC) - daytime profile
    hour_w = np.array([1, 1, 1, 1, 2, 3, 5, 7, 9, 10, 10, 9, 8, 8, 8, 7, 6, 5, 4, 3, 3, 2, 2, 1], float)
    hours = rng.choice(24, N, p=hour_w / hour_w.sum())
    minutes = rng.integers(0, 60, N)
    b["booking_datetime_utc"] = b["booking_date"] + pd.to_timedelta(hours, unit="h") + pd.to_timedelta(minutes, unit="m")

    # 10) cancellations
    p_cancel = b["partner_type"].map(pt["cancel_base"]).to_numpy() + np.where(lead > 60, 0.08, 0)
    p_cancel = np.where(hotel, p_cancel, 0.08)
    cancelled = rng.random(N) < p_cancel
    frac = rng.beta(2.5, 1.5, N)   # cancellations cluster closer to check-in
    cancel_date = b["booking_date"] + pd.to_timedelta(np.floor(frac * lead), unit="D")
    # E6: local event at one destination -> extra cancellations during a few days
    e6 = EV["E6_cancellations"]
    at_risk = (hotel & (b["destination_country"] == e6["dest"])
               & b["checkin_date"].between(e6["checkin_start"], e6["checkin_end"])
               & (b["booking_date"] < e6["cancel_start"]) & ~pd.Series(cancelled)).to_numpy()
    extra = at_risk & (rng.random(N) < e6["extra_rate"])
    span = (pd.Timestamp(e6["cancel_end"]) - pd.Timestamp(e6["cancel_start"])).days + 1
    cancel_date = cancel_date.where(~extra, pd.Timestamp(e6["cancel_start"])
                                    + pd.to_timedelta(rng.integers(0, span, N), unit="D"))
    cancelled = cancelled | extra
    seen = cancelled & (cancel_date <= EXTRACT_DATE).to_numpy()
    b["status"] = np.where(seen, "Cancelled", "Confirmed")
    b["cancel_date"] = cancel_date.where(seen)

    # 11) ids, in booking order
    b = b.sort_values("booking_datetime_utc").reset_index(drop=True)
    b.insert(0, "booking_id", [f"BK{i:07d}" for i in range(1, len(b) + 1)])

    # 12) price-match requests (E7) -> apply granted discounts to the matching bookings
    pm = make_price_match(b, partners)
    disc = pm[(pm["decision"] == "Discount") & (pm["booked"] == "Yes")].set_index("booking_id")["discount_eur"]
    b.loc[b["booking_id"].isin(disc.index), "discount_eur"] = b["booking_id"].map(disc)

    # 13) data-quality issues (E5) - injected AFTER everything else, like a faulty load would
    e5 = EV["E5_missing_cost"]
    miss = (b["supplier_id"].eq(e5["supplier"]) & b["booking_date"].between(e5["start"], e5["end"]))
    b.loc[miss, "net_cost_eur"] = np.nan
    tz = EV["E5_timezone"]
    shift = b["channel"].eq(tz["channel"]) & b["booking_date"].between(tz["start"], tz["end"])
    b.loc[shift, "booking_datetime_utc"] += pd.Timedelta(hours=tz["shift_hours"])
    b["booking_date"] = b["booking_datetime_utc"].dt.normalize()
    dup = b[b["booking_date"].eq(EV["E5_duplicates"]["date"]) & b["channel"].eq(EV["E5_duplicates"]["channel"])]
    b = pd.concat([b, dup]).sort_values(["booking_datetime_utc", "booking_id"]).reset_index(drop=True)
    b = b[b["booking_date"] <= END]

    cols = ["booking_id", "booking_date", "booking_datetime_utc", "checkin_date", "brand", "line_of_business",
            "client_country", "destination_country", "supplier_id", "partner_id", "partner_type", "channel",
            "room_nights", "currency", "sell_price_local", "fx_rate_to_eur", "sell_price_eur", "net_cost_eur",
            "discount_eur", "status", "cancel_date"]
    b = b[cols]
    b.to_csv(SIM / "bookings.csv", index=False, date_format="%Y-%m-%d")
    # keep the time in the datetime column
    b.assign(booking_datetime_utc=b["booking_datetime_utc"].dt.strftime("%Y-%m-%d %H:%M")) \
     .to_csv(SIM / "bookings.csv", index=False, date_format="%Y-%m-%d")

    # dimension tables
    sup = SUPPLIERS.drop(columns=["cost_factor"]).copy()
    sup.to_csv(SIM / "suppliers.csv", index=False)
    clean = b.drop_duplicates("booking_id")
    sales = (clean[clean["status"] == "Confirmed"].groupby("partner_id")["sell_price_eur"].sum())
    partners["sales_eur_18m"] = partners["partner_id"].map(sales).fillna(0).round(0)
    top = partners[partners["partner_type"] != "B2C"]["sales_eur_18m"].quantile(0.95)
    partners["tier"] = np.where((partners["sales_eur_18m"] >= top) & (partners["partner_type"] != "B2C"),
                                "Top partner", "Standard")
    partners["brand"] = partners["partner_type"].map(pt["brand"])
    partners[["partner_id", "partner_name", "partner_type", "brand", "country", "tier",
              "commercial_terms"]].to_csv(SIM / "partners.csv", index=False)
    pm.to_csv(SIM / "price_match_requests.csv", index=False, date_format="%Y-%m-%d")

    fxo = fx[(fx["date"] >= START - pd.Timedelta(days=62)) & (fx["date"] <= END)
             & fx["currency"].isin(MARKETS["currency"].unique())].copy()
    fxo["fx_key"] = fxo["date"].dt.strftime("%Y-%m-%d") + "|" + fxo["currency"]
    fxo.to_csv(SIM / "fx_rates.csv", index=False, date_format="%Y-%m-%d")

    cal = pd.DataFrame({"date": pd.date_range(START - pd.Timedelta(days=0), END + pd.Timedelta(days=330))})
    cal["date_key"] = cal["date"].dt.strftime("%Y-%m-%d")
    cal["year"] = cal["date"].dt.year
    cal["month_num"] = cal["date"].dt.month
    cal["year_month"] = cal["date"].dt.strftime("%Y-%m")
    cal["month_name"] = cal["date"].dt.strftime("%b")
    cal["iso_week"] = cal["date"].dt.isocalendar().week.astype(int)
    cal["weekday"] = cal["date"].dt.strftime("%a")
    cal["weekday_num"] = cal["date"].dt.day_of_week + 1
    cal["is_weekend"] = cal["weekday_num"] >= 6
    cal.to_csv(SIM / "calendar.csv", index=False, date_format="%Y-%m-%d")

    truth = pd.DataFrame([(k, str(v)) for k, v in EV.items()], columns=["event", "parameters"])
    truth.to_csv(SIM / "events_truth.csv", index=False)
    print(f"bookings rows: {len(b):,}  unique ids: {b['booking_id'].nunique():,}  "
          f"price-match requests: {len(pm):,}")


def make_price_match(b: pd.DataFrame, partners: pd.DataFrame) -> pd.DataFrame:
    """E7: internal requests from partners claiming a lower competitor price.

    Inputs are random; the DECISION is not random: it applies the written rules
    (playbook/price_match_rules.md), the same rules as the Excel calculator.
    """
    pool = b[b["line_of_business"].eq("Hotels") & b["partner_type"].isin(["Agency", "API partner", "Corporate"])]
    n = int(len(b) * 0.007)
    s = pool.sample(n, random_state=SEED).reset_index(drop=True)
    r = np.random.default_rng(SEED + 7)
    case = r.choice(["like_for_like", "non_refundable", "no_breakfast", "taxes_excluded",
                     "missing_evidence", "old_screenshot"], n, p=[0.42, 0.16, 0.12, 0.12, 0.10, 0.08])
    etg_price = s["sell_price_eur"].to_numpy()
    gap = np.where(case == "like_for_like", r.uniform(0.01, 0.12, n), r.uniform(0.03, 0.15, n))
    comp_price = np.round(etg_price * (1 - gap), 2)
    pm = pd.DataFrame({
        "request_id": [f"PM{i:05d}" for i in range(1, n + 1)],
        "request_date": s["booking_date"] - pd.to_timedelta(r.integers(0, 2, n), unit="D"),
        "partner_id": s["partner_id"], "booking_id": s["booking_id"],
        "destination_country": s["destination_country"], "supplier_id": s["supplier_id"],
        "checkin_date": s["checkin_date"], "room_nights": s["room_nights"],
        "competitor": r.choice(["Competitor A (B2B bedbank)", "Competitor B (OTA)", "Competitor C (wholesaler)"],
                               n, p=[0.45, 0.35, 0.20]),
        "etg_price_eur": etg_price, "competitor_price_eur": comp_price,
        "same_room_type": "Yes", "same_dates": "Yes",
        "etg_meal": r.choice(["Room only", "Breakfast"], n, p=[0.55, 0.45]),
        "etg_cancellation": r.choice(["Free cancellation", "Non-refundable"], n, p=[0.8, 0.2]),
        "competitor_taxes_included": "Yes",
        "screenshot": "Yes", "screenshot_age_hours": r.integers(1, 20, n),
        "net_cost_eur": s["net_cost_eur"].to_numpy(),
    })
    pm["competitor_meal"] = pm["etg_meal"]
    pm["competitor_cancellation"] = pm["etg_cancellation"]
    nr = case == "non_refundable"
    pm.loc[nr, "etg_cancellation"] = "Free cancellation"
    pm.loc[nr, "competitor_cancellation"] = "Non-refundable"
    nb = case == "no_breakfast"
    pm.loc[nb, "etg_meal"] = "Breakfast"
    pm.loc[nb, "competitor_meal"] = "Room only"
    pm.loc[case == "taxes_excluded", "competitor_taxes_included"] = "No"
    me = case == "missing_evidence"
    pm.loc[me, "screenshot"] = "No"
    pm.loc[me & (r.random(n) < 0.4), "same_dates"] = "No"
    pm.loc[case == "old_screenshot", "screenshot_age_hours"] = r.integers(49, 160, (case == "old_screenshot").sum())
    # A cheaper alternative supplier exists for some hotels (found by re-shopping our own suppliers).
    alt = r.random(n) < 0.35
    pm["alt_supplier_net_eur"] = np.where(alt, np.round(pm["net_cost_eur"] * r.uniform(0.86, 0.97, n), 2), np.nan)

    dec = pm.apply(decide, axis=1, result_type="expand")
    pm = pd.concat([pm, dec], axis=1)
    pm["handling_hours"] = np.round(np.where(pm["decision"] == "Ask for details", r.uniform(0.3, 2, n),
                                             r.gamma(2.0, 1.6, n)), 1)
    p_book = pm["decision"].map({"Discount": 0.82, "Switch supplier": 0.80, "Refuse": 0.35,
                                 "Ask for details": 0.45})
    pm["booked"] = np.where(r.random(n) < p_book, "Yes", "No")
    pm["case_type_truth"] = case
    return pm


def decide(row: pd.Series) -> pd.Series:
    """Price-match decision rules (documented in playbook/price_match_rules.md)."""
    # Step 1 - evidence
    if row["screenshot"] != "Yes" or row["screenshot_age_hours"] > 48 or row["same_dates"] != "Yes":
        return pd.Series({"adjusted_competitor_eur": np.nan, "valid_gap_eur": np.nan, "max_discount_eur": np.nan,
                          "decision": "Ask for details", "discount_eur": 0.0,
                          "reason": "Evidence incomplete (screenshot, age or dates)"})
    # Step 2 - like-for-like adjustment of the competitor price
    adj = row["competitor_price_eur"]
    reasons = []
    nights = max(int(row["room_nights"]), 1)
    if row["competitor_taxes_included"] == "No":
        adj += 0.10 * row["competitor_price_eur"]
        reasons.append("competitor price excludes ~10% taxes")
    if row["etg_meal"] == "Breakfast" and row["competitor_meal"] == "Room only":
        adj += 12.0 * nights * 2       # 2 guests x EUR 12 per breakfast
        reasons.append("our rate includes breakfast")
    if row["etg_cancellation"] == "Free cancellation" and row["competitor_cancellation"] == "Non-refundable":
        adj += 0.08 * row["etg_price_eur"]
        reasons.append("our rate is refundable (flexibility worth ~8%)")
    gap = row["etg_price_eur"] - adj
    floor_price = row["net_cost_eur"] / (1 - MARGIN_FLOOR)
    max_disc = max(row["etg_price_eur"] - floor_price, 0.0)
    out = {"adjusted_competitor_eur": round(adj, 2), "valid_gap_eur": round(gap, 2),
           "max_discount_eur": round(max_disc, 2)}
    # Step 3 - decision
    if gap <= 0:
        out.update(decision="Refuse", discount_eur=0.0,
                   reason="Not like-for-like: " + "; ".join(reasons) if reasons else "No real gap")
    elif gap <= max_disc:
        out.update(decision="Discount", discount_eur=round(gap, 2), reason="Valid gap, margin stays above floor")
    elif pd.notna(row["alt_supplier_net_eur"]) and row["alt_supplier_net_eur"] / (1 - MARGIN_FLOOR) <= adj:
        out.update(decision="Switch supplier", discount_eur=0.0,
                   reason="Cheaper supplier found: matched price keeps margin above floor")
    else:
        out.update(decision="Refuse", discount_eur=0.0, reason="Valid gap but matching breaks the margin floor")
    return pd.Series(out)


if __name__ == "__main__":
    main()
