"""Build the case-study report: charts (matplotlib) -> HTML pages -> PDF (Edge/Chrome headless).

Every number shown comes from a file written by another script (src/analysis_real.py,
src/analysis.py, src/collect_*.py) or from the Excel workbook screenshots.
Output: docs/ETG_CommercialPricing_CaseStudy.pdf and docs/img/page_XX.png
Run from the repository root:  python src/build_report.py
"""
from __future__ import annotations

import base64
import html
import shutil
import subprocess
from pathlib import Path

import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.ticker import FuncFormatter, PercentFormatter

from common import (BLUE, BLUE_DARK, BLUE_LIGHT, BLUE_PALE, BORDER, DOCS, GREY, GREY_LIGHT, IMG, INK, INK_SOFT,
                    OUT, PUBLIC, RED, ROOT)

matplotlib.use("Agg")
plt.rcParams.update({
    "font.family": "Segoe UI", "font.size": 10, "axes.edgecolor": GREY_LIGHT, "axes.labelcolor": INK_SOFT,
    "xtick.color": GREY, "ytick.color": GREY, "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "axes.grid.axis": "y", "grid.color": "#ECEBE6", "grid.linewidth": 0.8,
    "axes.titlesize": 11, "axes.titleweight": "bold", "axes.titlecolor": INK, "axes.titlelocation": "left",
    "legend.frameon": False, "figure.facecolor": "white", "axes.facecolor": "white", "savefig.dpi": 200,
})
REAL, SIMO = OUT / "real", OUT / "sim"
BUILD = ROOT / "docs" / "_build"          # temporary HTML (deleted at the end)
PDF = DOCS / "ETG_CommercialPricing_CaseStudy.pdf"
FOOT = "Independent case study based on public data · not affiliated with or endorsed by Emerging Travel Group"
def ai_note() -> str:
    """The AI-assistance note is written once, in README.md (section 'AI note'), and copied into the report."""
    text = (ROOT / "README.md").read_text(encoding="utf-8")
    return text.split("## AI note", 1)[1].strip().split("\n\n")[0].replace("\n", " ")


AI_NOTE = ai_note()
eur = lambda x: f"€{x:,.0f}"
keur = lambda x: f"€{x / 1000:,.0f}k"
pct = lambda x, d=0: f"{x * 100:.{d}f}%"


def date_axis(ax):
    """Readable date ticks (no overlapping labels)."""
    import matplotlib.dates as mdates
    loc = mdates.AutoDateLocator(minticks=3, maxticks=6)
    ax.xaxis.set_major_locator(loc)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(loc))


def signed_eur(x: float) -> str:
    return ("+" if x >= 0 else "−") + f"€{abs(x):,.0f}"


def save(fig, name: str) -> Path:
    BUILD.mkdir(parents=True, exist_ok=True)
    p = BUILD / f"{name}.png"
    fig.savefig(p, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return p


def img(path: Path, style: str = "") -> str:
    data = base64.b64encode(Path(path).read_bytes()).decode()
    return f'<img src="data:image/png;base64,{data}" style="{style}">'


# =========================================================================== charts & numbers
def page1():
    f = pd.read_csv(PUBLIC / "etg_facts.csv").set_index("fact_id")
    h = pd.read_csv(PUBLIC / "etg_headline_history.csv")
    h["snapshot"] = pd.to_datetime(h["snapshot"])
    fig, ax = plt.subplots(figsize=(7.2, 3.6))
    series = [("direct_hotel_contracts", "Direct hotel contracts", BLUE_DARK),
              ("b2b_agents_or_partners", "B2B agents / partners", BLUE),
              ("suppliers", "Suppliers", BLUE_LIGHT), ("accommodation_options", "Accommodation options", GREY)]
    for col, lab, colr in series:
        s = h.set_index("snapshot")[col].dropna()
        idx = s / s.iloc[0] * 100
        ax.plot(idx.index, idx.values, color=colr, lw=2, marker="o", ms=4)
        ax.annotate(f"{lab}  ×{s.iloc[-1] / s.iloc[0]:.1f}", (idx.index[-1], idx.values[-1]), xytext=(6, 0),
                    textcoords="offset points", va="center", fontsize=9, color=INK_SOFT)
    ax.set_title("Headline figures on emergingtravel.com, index July 2021 = 100")
    ax.set_xlim(h["snapshot"].min() - pd.Timedelta(days=60), h["snapshot"].max() + pd.Timedelta(days=560))
    date_axis(ax)
    ax.set_ylabel("Index")
    p = save(fig, "p1_history")
    growth = h["direct_hotel_contracts"].iloc[-1] / h["direct_hotel_contracts"].iloc[0]
    return {"chart": p, "facts": f, "growth_direct": growth, "h": h}


def page2():
    a = pd.read_csv(PUBLIC / "eurostat_nights_by_residence.csv")
    aggregates = {"WORLD", "EUR", "INT_EU27_2020", "INT_EU28", "INT_EU27_2007", "DOM", "EU27_2020_FOR", "EFTA",
                  "EUR_OTH", "AFR", "AFR_OTH", "AME", "AME_N", "AME_N_OTH", "AME_C_S", "AME_C_S_OTH", "ASI",
                  "ASI_OTH", "OCE", "OCE_OTH", "WRL_NAL", "CH_LI"}
    y = a[(a["year"] == 2024) & ~a["residence_code"].isin(aggregates)].copy()
    world = a[(a["year"] == 2024) & (a["residence_code"] == "WORLD")]["nights"].iloc[0]
    y["share"] = y["nights"] / world
    top = y.sort_values("nights", ascending=False).head(12).iloc[::-1]
    fig, ax = plt.subplots(figsize=(5.2, 4.0))
    ax.barh(top["residence"].str.replace(" (from 2020)", ""), top["nights"] / 1e6, color=BLUE, height=0.62)
    for yy, (v, s) in enumerate(zip(top["nights"] / 1e6, top["share"])):
        ax.text(v + 8, yy, f"{v:,.0f}M · {s:.0%}", va="center", fontsize=8.5, color=INK_SOFT)
    ax.grid(axis="y", visible=False)
    ax.grid(axis="x", visible=True)
    ax.set_xlim(0, top["nights"].max() / 1e6 * 1.3)
    ax.set_title("Nights in EU accommodation by country of residence, 2024 (millions)")
    p_markets = save(fig, "p2_markets")

    s = pd.read_csv(PUBLIC / "eurostat_market_seasonality.csv")
    s = s[(s["year"] == 2024) & (s["market_code"] != "SE")]   # SE 2024 incomplete in the source
    piv = s.pivot_table(index="market", columns="month", values="nights")
    piv = piv.div(piv.sum(axis=1), axis=0)
    piv = piv.loc[piv[[7, 8]].sum(axis=1).sort_values().index]
    fig, ax = plt.subplots(figsize=(6.6, 3.9))
    im = ax.imshow(piv.values, cmap=matplotlib.colors.LinearSegmentedColormap.from_list("b", ["#F7FAFE", BLUE_PALE, BLUE, BLUE_DARK]),
                   aspect="auto", vmin=0, vmax=0.35)
    ax.set_xticks(range(12), ["J", "F", "M", "A", "M", "J", "J", "A", "S", "O", "N", "D"])
    ax.set_yticks(range(len(piv)), piv.index)
    ax.grid(False)
    for i in range(piv.shape[0]):
        for j in range(12):
            v = piv.values[i, j]
            ax.text(j, i, f"{v * 100:.0f}", ha="center", va="center", fontsize=7.5,
                    color="white" if v > 0.17 else INK_SOFT)
    ax.set_title("Share of a market's nights on trips abroad, by month of departure, 2024 (%)")
    for sp in ax.spines.values():
        sp.set_visible(False)
    p_season = save(fig, "p2_season")
    it_aug = piv.loc["Italy", 8]
    de_peak = piv.loc["Germany"].max()
    de_months = (piv.loc["Germany"] >= 0.08).sum()
    return {"markets": p_markets, "season": p_season, "it_aug": it_aug, "de_peak": de_peak, "de_months": de_months,
            "top": y.sort_values("nights", ascending=False).head(5), "piv": piv}


def page3():
    ch = pd.read_csv(REAL / "channel_kpis.csv")
    ch = ch[ch["channel"].isin(["TA/TO", "Direct", "Corporate"])].set_index("channel").loc[["TA/TO", "Direct", "Corporate"]]
    fig, axes = plt.subplots(1, 3, figsize=(8.6, 2.7))
    specs = [("adr_eur", "ADR (EUR per night)", lambda v: f"€{v:.0f}"),
             ("cancel_rate", "Cancellation rate", lambda v: f"{v:.0%}"),
             ("median_lead_time_days", "Median lead time (days)", lambda v: f"{v:.0f}")]
    for ax, (col, t, fmt) in zip(axes, specs):
        colors = [BLUE, GREY_LIGHT, GREY_LIGHT]
        ax.bar(ch.index, ch[col], color=colors, width=0.6)
        for i, v in enumerate(ch[col]):
            ax.text(i, v * 1.02, fmt(v), ha="center", va="bottom", fontsize=9, color=INK_SOFT)
        ax.set_title(t, fontsize=10)
        ax.set_yticks([])
        ax.spines["left"].set_visible(False)
        ax.grid(False)
        ax.set_ylim(0, ch[col].max() * 1.25)
    p_ch = save(fig, "p3_channels")

    ag = pd.read_csv(REAL / "agent_pareto.csv")
    fig, ax = plt.subplots(figsize=(4.6, 3.0))
    ax.plot(ag["rank"], ag["cum_share"], color=BLUE, lw=2)
    ax.axhline(0.8, color=GREY, ls="--", lw=1)
    n80 = int((ag["cum_share"] < 0.8).sum() + 1)
    ax.axvline(n80, color=GREY, ls="--", lw=1)
    ax.annotate(f"{n80} of {len(ag)} agents = 80% of revenue", (n80, 0.8), xytext=(30, -40),
                textcoords="offset points", fontsize=9, color=INK_SOFT)
    ax.yaxis.set_major_formatter(PercentFormatter(1))
    ax.set_xscale("log")
    ax.set_xlabel("Agents ranked by realised room revenue (log scale)")
    ax.set_title("Revenue concentration on travel agents")
    p_par = save(fig, "p3_pareto")
    lt = pd.read_csv(REAL / "lead_time_cancellation.csv")
    lt = lt[lt["channel"].isin(["TA/TO", "Direct"]) & (lt["lead_bucket"] != "365+")]   # 365+: too few direct bookings
    fig, ax = plt.subplots(figsize=(4.6, 3.0))
    for chn, col in (("TA/TO", BLUE), ("Direct", GREY)):
        x = lt[lt["channel"] == chn]
        ax.plot(x["lead_bucket"], x["cancel_rate"], color=col, lw=2, marker="o", ms=4, label=chn)
        ax.annotate(chn, (len(x) - 1, x["cancel_rate"].iloc[-1]), xytext=(6, 0), textcoords="offset points",
                    color=INK_SOFT, fontsize=9, va="center")
    ax.yaxis.set_major_formatter(PercentFormatter(1))
    ax.set_xlabel("Lead time (days)")
    ax.set_title("Cancellation rate by lead time")
    ax.legend(loc="upper left", fontsize=8)
    p_lt = save(fig, "p3_leadtime")
    dep = pd.read_csv(REAL / "deposit_cancellation.csv").set_index("deposit_type")
    dq = pd.read_csv(REAL / "dq_checks.csv").set_index("check")
    ta = ch.loc["TA/TO"]
    return {"ch": p_ch, "pareto": p_par, "lt": p_lt, "n80": n80, "agents": len(ag), "top1": ag["cum_share"].iloc[0],
            "top10": ag["cum_share"].iloc[9], "ta": ta, "di": ch.loc["Direct"], "dep": dep, "dq": dq,
            "kpis": pd.read_csv(REAL / "channel_kpis.csv")}


def waterfall(ax, labels, values, colors_pos=BLUE, colors_neg=RED, fmt=keur):
    cum = 0
    for i, (lab, v) in enumerate(zip(labels, values)):
        if i in (0, len(values) - 1):
            ax.bar(i, v, color=GREY if i == 0 else BLUE_DARK, width=0.6)
            ax.text(i, v, fmt(v), ha="center", va="bottom", fontsize=9, color=INK_SOFT)
            cum = v
        else:
            bottom = cum if v >= 0 else cum + v
            ax.bar(i, abs(v), bottom=bottom, color=colors_pos if v >= 0 else colors_neg, width=0.6)
            ax.text(i, bottom + abs(v), ("+" if v >= 0 else "−") + fmt(abs(v)), ha="center", va="bottom",
                    fontsize=9, color=INK_SOFT)
            cum += v
    ax.set_xticks(range(len(labels)), labels, fontsize=9)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: keur(v)))


def page4():
    b = pd.read_csv(REAL / "bridge.csv")
    dd = pd.read_csv(REAL / "bridge_drilldown.csv")
    fig, ax = plt.subplots(figsize=(6.0, 3.7))
    labs = ["Aug 2016", "Volume", "Mix", "Price", "Sep 2016"]
    waterfall(ax, labs, b["eur"].tolist())
    ax.set_ylim(0, b["eur"].iloc[0] * 1.12)
    ax.set_title("Room revenue, August → September 2016 (arrival month, stays not cancelled)")
    p_w = save(fig, "p4_bridge")
    m = pd.read_csv(REAL / "monthly.csv")
    m["arrival_month"] = pd.PeriodIndex(m["arrival_month"], freq="M").to_timestamp()
    fig, ax = plt.subplots(figsize=(5.4, 3.0))
    for hotel, col in (("Resort Hotel", BLUE), ("City Hotel", GREY)):
        x = m[m["hotel"] == hotel]
        ax.plot(x["arrival_month"], x["adr_eur"], color=col, lw=2, label=hotel)
        ax.annotate(hotel, (x["arrival_month"].iloc[-1], x["adr_eur"].iloc[-1]), xytext=(5, 0),
                    textcoords="offset points", fontsize=8.5, color=INK_SOFT, va="center")
    ax.axvspan(pd.Timestamp("2016-07-20"), pd.Timestamp("2016-09-15"), color=BLUE_PALE, alpha=0.6, lw=0)
    ax.set_title("ADR by arrival month (EUR per night)")
    ax.legend(loc="upper left", fontsize=8)
    date_axis(ax)
    p_adr = save(fig, "p4_adr")
    seg = dd.groupby(["hotel", "market_segment"])[["N0", "N1", "R0", "R1", "volume", "mix", "price", "total_change"]].sum()
    seg = seg.reindex(seg["total_change"].abs().sort_values(ascending=False).index).head(6).reset_index()
    seg["P0"] = seg["R0"] / seg["N0"].replace(0, np.nan)
    seg["P1"] = seg["R1"] / seg["N1"].replace(0, np.nan)
    tot = dd[["N0", "N1"]].sum()
    resort = dd.groupby("hotel")["total_change"].sum()["Resort Hotel"]
    return {"w": p_w, "adr": p_adr, "b": b, "seg": seg, "N0": tot["N0"], "N1": tot["N1"], "resort": resort}


def page5():
    d = pd.read_csv(SIMO / "daily_total.csv", parse_dates=["booking_date"])
    al = pd.read_csv(SIMO / "alerts.csv")
    br = pd.read_csv(SIMO / "margin_bridge_sep.csv")
    fig, ax = plt.subplots(figsize=(5.6, 3.3))
    labs = ["Sep 2025", "Volume", "Mix", "Price", "FX", "Cost", "Sep 2026"]
    waterfall(ax, labs, br["eur"].tolist())
    ax.set_ylim(0, br["eur"].max() * 1.35)
    ax.set_title("Monthly layer: margin bridge, Sep 2025 → Sep 2026 (EUR)")
    p_b = save(fig, "p5_bridge")
    matched = al["matched_event"].notna().sum()
    months = (pd.to_datetime(al["date"]).max() - pd.to_datetime(al["date"]).min()).days / 30.4
    excel = {k: IMG / "excel" / f"{f}.png" for k, f in (("cockpit", "cockpit_kpis"), ("alerts", "alerts"),
                                                        ("pivots", "pivots"), ("bridge", "bridge"))}
    return {"bridge": p_b, "br": br, "alerts": al, "matched": matched, "per_month": len(al) / months,
            "excel": {k: v for k, v in excel.items() if v.exists()}, "day": d[d["booking_date"] == "2026-09-15"].iloc[0]}


def page6():
    c = pd.read_csv(SIMO / "bookings_clean.csv", parse_dates=["booking_date"])
    inc = pd.read_csv(SIMO / "incidents.csv")
    x = c[(c["supplier_id"] == "SUP-BB-C") & (c["destination_country"] == "ES")]
    w = x.groupby(pd.Grouper(key="booking_date", freq="W-MON")).agg(m=("margin_eur", "sum"), s=("sales_eur", "sum"),
                                                                    n=("booking_id", "size"))
    w = w[(w.index >= "2026-03-01")]
    fig, ax = plt.subplots(figsize=(3.9, 2.4))
    ax.plot(w.index, w["m"] / w["s"], color=RED, lw=2)
    ax.axhline(0.04, color=GREY, ls="--", lw=1)
    ax.text(w.index[0], 0.043, "margin floor 4%", fontsize=8, color=GREY)
    ax.axvline(pd.Timestamp("2026-06-15"), color=GREY_LIGHT, lw=1)
    ax.yaxis.set_major_formatter(PercentFormatter(1, decimals=0))
    ax.set_title("E1 · Bedbank C × Spain, weekly margin rate", fontsize=9.5)
    ax.tick_params(labelsize=8)
    date_axis(ax)
    p1 = save(fig, "p6_e1")
    y = c[c["partner_id"] == "API-002"].groupby("booking_date").size()
    y = y.reindex(pd.date_range("2026-07-01", "2026-10-31"), fill_value=0).rolling(7).mean()
    fig, ax = plt.subplots(figsize=(3.9, 2.4))
    ax.plot(y.index, y.values, color=BLUE, lw=2)
    ax.set_title("E2 · API-002 bookings per day (7-day avg)", fontsize=9.5)
    ax.tick_params(labelsize=8)
    date_axis(ax)
    p2 = save(fig, "p6_e2")
    return {"inc": inc, "e1": p1, "e2": p2}


def page7():
    rc = pd.read_csv(SIMO / "daily_raw_vs_clean.csv", parse_dates=["booking_date"])
    dq = pd.read_csv(SIMO / "dq_issues.csv")
    a = rc[rc["booking_date"].between("2026-07-07", "2026-08-16")]
    fig, axes = plt.subplots(1, 2, figsize=(9.2, 2.8))
    ax = axes[0]
    ax.plot(a["booking_date"], a["sales_eur_raw"] / 1000, color=RED, lw=1.6, label="Raw")
    ax.plot(a["booking_date"], a["sales_eur_clean"] / 1000, color=BLUE, lw=2, label="Clean")
    ax.set_title("Daily sales (EUR k): duplicated load on 21 Jul", fontsize=10)
    ax.legend(fontsize=8, loc="lower left")
    ax.tick_params(labelsize=8)
    date_axis(ax)
    ax = axes[1]
    ax.plot(a["booking_date"], a["margin_rate_raw"], color=RED, lw=1.6, label="Raw (missing cost read as 0)")
    ax.plot(a["booking_date"], a["margin_rate_clean"], color=BLUE, lw=2, label="Clean (cost imputed)")
    ax.yaxis.set_major_formatter(PercentFormatter(1, decimals=0))
    ax.set_title("Daily margin rate: missing supplier costs, 3-9 Aug", fontsize=10)
    ax.legend(fontsize=8, loc="upper left")
    ax.tick_params(labelsize=8)
    date_axis(ax)
    p1 = save(fig, "p7_dq")
    b = pd.read_csv(ROOT / "data" / "simulated" / "bookings.csv", usecols=["booking_datetime_utc", "channel", "booking_date"],
                    parse_dates=["booking_datetime_utc", "booking_date"])
    api = b[b["channel"] == "API"]
    normal = api[api["booking_date"].between("2026-09-14", "2026-10-04")]["booking_datetime_utc"].dt.hour.value_counts(normalize=True).sort_index()
    shifted = api[api["booking_date"].between("2026-10-12", "2026-10-18")]["booking_datetime_utc"].dt.hour.value_counts(normalize=True).sort_index()
    fig, ax = plt.subplots(figsize=(4.4, 2.8))
    ax.bar(normal.index - 0.2, normal.values, width=0.4, color=GREY, label="Usual weeks")
    ax.bar(shifted.index + 0.2, shifted.values, width=0.4, color=RED, label="12-18 Oct")
    ax.yaxis.set_major_formatter(PercentFormatter(1, decimals=0))
    ax.set_xlabel("Booking hour (UTC)")
    ax.set_title("API feed: booking hours moved +5 h", fontsize=10)
    ax.legend(fontsize=8)
    ax.tick_params(labelsize=8)
    p2 = save(fig, "p7_tz")
    return {"chart": p1, "tz": p2, "dq": dq, "rc": rc}


def page8():
    s = pd.read_csv(SIMO / "pm_summary.csv").iloc[0]
    by = pd.read_csv(SIMO / "pm_by_decision.csv").set_index("decision")
    cs = pd.read_csv(SIMO / "pm_by_case.csv").set_index("case_type_truth")
    order = ["Discount", "Switch supplier", "Refuse", "Ask for details"]
    cols = {"Discount": BLUE, "Switch supplier": BLUE_LIGHT, "Refuse": RED, "Ask for details": GREY}
    names = {"like_for_like": "Same product", "non_refundable": "Competitor non-refundable",
             "no_breakfast": "Competitor without breakfast", "taxes_excluded": "Competitor price without taxes",
             "missing_evidence": "No screenshot / other dates", "old_screenshot": "Screenshot > 48 h old"}
    cs = cs.reindex(list(names)).fillna(0)
    fig, ax = plt.subplots(figsize=(6.4, 3.1))
    left = np.zeros(len(cs))
    for dname in order:
        if dname in cs.columns:
            v = cs[dname].to_numpy()
            ax.barh([names[i] for i in cs.index], v, left=left, color=cols[dname], label=dname, height=0.62,
                    edgecolor="white", linewidth=1)
            left += v
    ax.invert_yaxis()
    ax.grid(axis="y", visible=False)
    ax.grid(axis="x", visible=True)
    ax.legend(ncol=4, fontsize=8, loc="upper center", bbox_to_anchor=(0.45, -0.12))
    ax.set_title("Decision by type of request (number of requests)")
    ax.tick_params(labelsize=8.5)
    p = save(fig, "p8_pm")
    return {"s": s, "by": by, "chart": p}


def page9():
    r = pd.read_csv(SIMO / "pricing_recommendations.csv")
    r["segment"] = r["partner_type"] + " · " + r["destination_country"]
    fig, ax = plt.subplots(figsize=(6.2, 3.8))
    rr = r.sort_values("margin_rate")
    colors = [RED if m < 0.05 else (BLUE if ch > 0 else (BLUE_LIGHT if ch < 0 else GREY_LIGHT))
              for m, ch in zip(rr["margin_rate"], rr["markup_change"])]
    ax.barh(rr["segment"], rr["margin_rate"], color=colors, height=0.65)
    ax.axvline(0.04, color=GREY, ls="--", lw=1)
    ax.text(0.0405, -0.9, "floor 4%", fontsize=8, color=GREY)
    ax.xaxis.set_major_formatter(PercentFormatter(1, decimals=0))
    ax.grid(axis="y", visible=False)
    ax.grid(axis="x", visible=True)
    ax.tick_params(labelsize=8)
    ax.set_title("Margin rate by B2B hotel segment, last 90 days")
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color=BLUE, label="+0.5 pt markup"), Patch(color=BLUE_LIGHT, label="−0.5 pt markup"),
                       Patch(color=RED, label="hold: < 1 pt above floor"), Patch(color=GREY_LIGHT, label="keep")],
              fontsize=8, loc="lower right")
    p = save(fig, "p9_segments")
    ch = r[r["markup_change"] != 0]
    return {"r": r, "chart": p, "changes": ch, "low": ch["margin_effect_eur_low"].sum(),
            "high": ch["margin_effect_eur_high"].sum()}


# =========================================================================== HTML
CSS = """
@page { size: 1600px 900px; margin: 0; }
* { box-sizing: border-box; }
body { margin: 0; font-family: 'Segoe UI', Arial, sans-serif; color: #0B0B0B; background: #F4F4F2; }
.page { width: 1600px; height: 900px; padding: 34px 44px 26px 44px; background: #F4F4F2; position: relative;
        page-break-after: always; overflow: hidden; display: flex; flex-direction: column; }
.page:last-child { page-break-after: auto; }
.head { display: flex; justify-content: space-between; align-items: flex-start; }
h1 { font-size: 29px; margin: 0 0 6px 0; font-weight: 700; max-width: 1330px; line-height: 1.18; }
.sub { color: #52514E; font-size: 16.5px; margin: 0 0 16px 0; max-width: 1350px; line-height: 1.4; }
.tag { font-size: 12px; font-weight: 700; color: #52514E; letter-spacing: .06em; white-space: nowrap; padding-top: 8px; }
.tag.sim { color: #1C5CAB; } .tag.real { color: #0B6B3A; }
.grid { display: grid; gap: 14px; flex: 1; min-height: 0; }
.card { background: #FCFCFB; border: 1px solid #E1E0D9; border-radius: 8px; padding: 14px 16px; overflow: hidden; }
.card h3 { margin: 0 0 8px 0; font-size: 17px; }
.card img { max-width: 100%; max-height: 100%; display: block; }
.kpi .lab { font-size: 11px; font-weight: 700; color: #52514E; letter-spacing: .05em; }
.kpi .val { font-size: 44px; font-weight: 700; line-height: 1.1; margin: 4px 0; }
.kpi .txt { font-size: 15.5px; color: #52514E; line-height: 1.35; }
.small { font-size: 15px; color: #3E3D3A; line-height: 1.5; }
.foot { font-size: 11.5px; color: #898781; margin-top: 10px; }
table { border-collapse: collapse; width: 100%; font-size: 14.5px; }
th { text-align: left; color: #52514E; font-weight: 600; border-bottom: 1px solid #898781; padding: 4px 6px; }
td { padding: 4px 6px; border-bottom: 1px solid #EFEEE9; vertical-align: top; }
td.num, th.num { text-align: right; white-space: nowrap; }
.red { color: #D03B3B; font-weight: 600; } .blue { color: #1C5CAB; font-weight: 600; }
ul { margin: 4px 0; padding-left: 18px; } li { margin: 3px 0; }
.q { font-size: 15px; line-height: 1.55; } .q b { color: #2A78D6; }
.inc { font-size: 13.6px; line-height: 1.42; } .inc .t { font-weight: 700; font-size: 15px; margin-bottom: 4px; }
.inc .imp { font-size: 20px; font-weight: 700; margin: 2px 0 4px 0; }
.flow { display: flex; align-items: center; gap: 10px; font-size: 14.5px; }
.box { background: #fff; border: 1px solid #E1E0D9; border-radius: 6px; padding: 10px 12px; text-align: center; flex: 1; }
.arrow { color: #898781; font-size: 20px; }
.quote { background: #F7F7F4; border-left: 3px solid #2A78D6; padding: 9px 11px; font-size: 14.5px; line-height: 1.45; margin: 6px 0; }
"""


def page_html(title: str, sub: str, tag: str, body: str, foot: str, grid_style: str) -> str:
    cls = {"PUBLIC DATA": "", "REAL DATASET": "real", "SIMULATED": "sim"}.get(tag, "")
    return f"""<section class="page"><div class="head"><div><h1>{title}</h1><p class="sub">{sub}</p></div>
<div class="tag {cls}">{tag}</div></div><div class="grid" style="{grid_style}">{body}</div>
<div class="foot">{foot} · {FOOT}</div></section>"""


def build_pages() -> list[str]:
    p1, p2, p3, p4, p5, p6, p7, p8, p9 = page1(), page2(), page3(), page4(), page5(), page6(), page7(), page8(), page9()
    pages = []
    inc = p6["inc"].set_index("event")
    s = p8["s"]
    ta, di = p3["ta"], p3["di"]
    br = p4["b"]["eur"].tolist()
    drop = br[-1] - br[0]

    # ------------------------------------------------------------ 0 overview
    k = [("REAL DATASET · PAGE 3", f"{pct(ta['cancel_rate'])} vs {pct(di['cancel_rate'])}",
          "cancellation rate of travel-agent bookings vs direct bookings, on 119,390 real hotel bookings"),
         ("REAL DATASET · PAGE 4", f"{pct((br[2] + br[3]) / drop)}",
          f"of a real {keur(-drop)} monthly revenue drop came from mix and price, not from fewer room nights"),
         ("SIMULATED · PAGE 6", keur(-inc.loc['E3', 'impact_eur']),
          "margin lost to monthly lira price lists lagging the real ECB rate: the costliest, and the slowest, event"),
         ("SIMULATED · PAGE 8", f"{pct(s['valid_share'])}",
          f"of price-match claims survive a like-for-like check; matching within the floor saved {keur(s['saved_sales_eur'])} of sales")]
    kp = "".join(f'<div class="card kpi"><div class="lab">{a}</div><div class="val">{b}</div><div class="txt">{c}</div></div>'
                 for a, b, c in k)
    body = f"""{kp}
<div class="card" style="grid-column: span 2"><h3 style="font-size:20px">Business questions</h3><div class="q">
<b>1 · Context</b>&nbsp; What does ETG sell, at what scale, and where is the margin made? <span class="small">(p.1)</span><br>
<b>2 · Demand</b>&nbsp; Which source markets matter and what is "normal" seasonality for each? <span class="small">(p.2)</span><br>
<b>3 · Channels</b>&nbsp; How do B2B bookings differ from direct ones on real data? <span class="small">(p.3-4)</span><br>
<b>4 · Monitoring</b>&nbsp; Is today normal, and if not, why? <span class="small">(p.5-7)</span><br>
<b>5 · Pricing</b>&nbsp; When to grant a price match, and where to move markups? <span class="small">(p.8-9)</span></div>
<p class="small" style="font-size:14px;margin-top:12px"><b style="color:#0B0B0B">Recommendations:</b> check data quality before any
conclusion; read every move against the same weekday and the market's own calendar; refresh local-currency price lists
weekly; re-route Spanish demand away from the supplier whose rates rose; decide price matches with a written like-for-like
checklist and a {pct(0.04)} margin floor; test +0.5 pt markups only where price pressure is low.</p></div>
<div class="card" style="grid-column: span 2"><h3 style="font-size:20px">Data &amp; methods</h3><div class="small" style="font-size:14px">
<b>PUBLIC DATA</b>&nbsp; ETG web pages (live and archived), Eurostat tourism statistics, ECB exchange rates.<br>
<b>REAL DATASET</b>&nbsp; Hotel booking demand (Antonio, Almeida &amp; Nunes, 2019, CC BY 4.0): two hotels in Portugal,
2015-2017, used to demonstrate the analysis, not ETG data.<br>
<b>SIMULATED</b>&nbsp; An 18-month daily sales ledger (192k bookings, seed 42) calibrated on public orders of magnitude,
with seven programmed events. Not ETG data.<br><br>
<b>Excel</b>&nbsp; Power Pivot model, ~50 DAX measures, CUBE formulas, PivotTables with slicers, dynamic arrays, price-match calculator.<br>
<b>SQL</b>&nbsp; DuckDB queries: segments, joins, alert band, top variations, data checks.<br>
<b>Python</b>&nbsp; collection, simulation, variance bridges, alert scoring.</div></div>
<div class="card" style="grid-column: span 4"><h3>How to read this report</h3><div class="small" style="font-size:13.5px">
<b>Titles</b> state the conclusion of the page. &nbsp;&nbsp;<b>Tags</b> PUBLIC DATA / REAL DATASET / SIMULATED mark where
the data comes from. &nbsp;&nbsp;<b>Blue</b> = data, <b>grey</b> = comparison, <b style="color:#D03B3B">red</b> = alert or
below the margin floor, <b>dashed lines</b> = thresholds. &nbsp;&nbsp;<b>Code</b> scripts, SQL, Excel workbook and sources are in the repository.</div></div>"""
    pages.append(page_html("Emerging Travel Group: a daily commercial &amp; pricing cockpit for a multi-brand B2B travel platform",
                           "Public market data, a real hotel-booking dataset and a simulated sales ledger, turned into daily "
                           "monitoring, root-cause analysis and a price-match decision process.", "OVERVIEW", body,
                           "Sources and collection dates in data/SOURCES.md", "grid-template-columns: repeat(4, 1fr); grid-template-rows: 190px 1fr 92px"))

    # ------------------------------------------------------------ 1 ETG public
    f = p1["facts"]
    v = lambda i: f.loc[i, "value"]
    tiles = [("Travel agents (RateHawk)", f"{v('F06') / 1000:.0f}K+"), ("Suppliers", f"{v('F03'):.0f}+"),
             ("API partners", f"{v('F10'):,.0f}+"), ("Direct properties", f"{v('F09') / 1000:.0f}K+"),
             ("Bookings per week", f"{v('F15') / 1000:.0f}K+"), ("Annual turnover (USD)", f"{v('F25') / 1e9:.1f}BN+")]
    tl = "".join(f'<div class="card kpi"><div class="lab">{a.upper()}</div><div class="val" style="font-size:34px">{b}</div></div>' for a, b in tiles)
    body = f"""<div style="grid-column: span 3; display:grid; grid-template-columns: repeat(6,1fr); gap:14px">{tl}</div>
<div class="card" style="grid-row: span 1"><h3>Three brands, one inventory</h3><div class="small">
<b>RateHawk</b> · B2B booking tool for travel professionals: hotels, flights, transfers, car rentals ({v('F06') / 1000:.0f}K+ agents).<br>
<b>ZenHotels</b> · B2C hotel booking site for individual travellers ({v('F08') / 1e6:.1f}M+ travellers).<br>
<b>Roundtrip</b> · corporate travel platform, est. 2022 ({v('F07') / 1000:.0f}K+ corporate clients).<br>
<b>Extranet</b> · hotels distribute directly ({v('F09') / 1000:.0f}K+ direct properties, {v('F05'):.0f}+ channel managers).<br>
Also: {v('F10'):,.0f}+ API partners, white label and affiliate programmes; {v('F02') / 1e6:.1f}M+ accommodation options in
{v('F01'):.0f}+ source markets; search answers in {v('F12'):.0f} ms over {v('F11') / 1000:.0f}K rates per second.</div></div>
<div class="card"><h3>Where the margin is made</h3><div class="flow" style="margin:8px 0 12px 0">
<div class="box"><b>Suppliers</b><br>hotels direct, bedbanks, wholesalers<br><span class="small">sell a <b>net rate</b></span></div><div class="arrow">→</div>
<div class="box"><b>ETG</b><br>adds a <b>markup</b><br><span class="small">by brand, segment, market</span></div><div class="arrow">→</div>
<div class="box"><b>Clients</b><br>agents, API partners, companies, travellers</div></div>
<div class="small"><b>Margin = selling price − supplier net cost.</b> It moves with the net rate (supplier), the markup
(pricing), the currency the client pays in (FX), and the mix of brands, markets and partners. A <b>price-match request</b>
arrives when a partner sees the same room cheaper elsewhere: the analyst checks it is the same product, then grants a
discount, switches supplier, or refuses with arguments.</div></div>
<div class="card">{img(p1['chart'], 'height:370px')}<div class="small">Direct hotel contracts grew ×{p1['growth_direct']:.1f}
between the July 2021 and November 2026 versions of the home page. Public figures differ between pages (190+ vs 220+ source
markets; 330+ vs 350+ suppliers; 110K+ agents vs 220K+ clients): each figure is quoted with its page and date.</div></div>"""
    pages.append(page_html(f"ETG connects {v('F03'):.0f}+ suppliers to {v('F06') / 1000:.0f}K+ agents and {v('F10'):,.0f}+ API partners: "
                           "its margin is the gap between net rates and B2B selling prices",
                           "What the company publishes about itself (emergingtravel.com, live and archived, collected 12 Nov 2026). "
                           "No financial detail is published beyond these headline figures.", "PUBLIC DATA", body,
                           "Sources: emergingtravel.com, ratehawk.com, web.archive.org (6 yearly snapshots) · data/public/etg_facts.csv",
                           "grid-template-columns: 1fr 1.05fr 1.15fr; grid-template-rows: 112px 1fr"))

    # ------------------------------------------------------------ 2 demand
    top5 = p2["top"]
    body = f"""<div class="card">{img(p2['markets'], 'height:520px')}</div>
<div class="card">{img(p2['season'], 'height:520px')}</div>
<div class="card" style="grid-column: span 2"><div class="small" style="font-size:13.5px"><b>What is "normal" per market</b>
 &nbsp;Residents of {top5['residence'].iloc[0]} spend the most nights in EU accommodation (domestic stays included), then
{top5['residence'].iloc[1]}, {top5['residence'].iloc[2]} and {top5['residence'].iloc[3]}. Italians take {pct(p2['it_aug'])} of
their nights abroad on trips starting in August; Germans spread theirs (peak month {pct(p2['de_peak'])}, {p2['de_months']} months
above 8%). A 30% jump in Italian bookings for August stays is seasonality, not an anomaly: the cockpit compares each day
with the same weekday and last year, and a market calendar explains peaks before anyone calls them a problem.
Sweden is left out: its 2024 monthly series is incomplete in the source.</div></div>"""
    pages.append(page_html(f"What is normal differs by market: Italians take {pct(p2['it_aug'])} of their trips abroad in August, "
                           "Germans spread theirs over half the year",
                           "Eurostat tourism statistics: where hotel nights in the EU come from, and when each source market travels.",
                           "PUBLIC DATA", body, "Source: Eurostat tour_occ_ninraw, tour_dem_tnmd (extracted 12 Nov 2026), re-use with attribution",
                           "grid-template-columns: 1fr 1.25fr; grid-template-rows: 1fr 92px"))

    # ------------------------------------------------------------ 3 B2B vs direct
    dep = p3["dep"]
    body = f"""<div class="card" style="grid-column: span 2">{img(p3['ch'], 'height:300px')}<div class="small">
Travel agents and tour operators (TA/TO, the B2B channel) bring {pct(ta['share_of_bookings'])} of bookings and
{pct(ta['share_of_revenue'])} of realised room revenue, at a lower nightly rate than direct guests (€{ta['adr_eur']:.0f}
vs €{di['adr_eur']:.0f}), book {ta['median_lead_time_days'] / di['median_lead_time_days']:.0f}× earlier and cancel
{ta['cancel_rate'] / di['cancel_rate']:.1f}× more often. For a B2B platform this means: price on net rates, watch the
cancellation exposure of long lead times, and look at margin after cancellations, not at gross bookings.</div></div>
<div class="card">{img(p3['pareto'], 'height:300px')}<div class="small">The largest agent alone weighs
{pct(p3['top1'])} of revenue, the top 10 {pct(p3['top10'])}. Losing one key partner moves the whole month (see E2, p.6).</div></div>
<div class="card">{img(p3['lt'], 'height:300px')}<div class="small">Cancellations rise with lead time in both channels;
B2B sits higher at every lead time.</div></div>
<div class="card" style="grid-column: span 4"><h3>Check before concluding: a suspicious pattern in real data</h3><div class="small">
"Non-refundable" bookings show a <span class="red">{pct(dep.loc['Non Refund', 'cancel_rate'], 1)}</span> cancellation rate,
far above refundable ones ({pct(dep.loc['No Deposit', 'cancel_rate'])}). {pct(dep.loc['Non Refund', 'share_from_PRT'])}
of them come from one country and {pct(dep.loc['Non Refund', 'share_TA_TO'])} from TA/TO: most likely a coding convention
(e.g. group blocks released by agents), not guests paying for rooms they never use. Treated as a question for the data
owner, not as a finding. Other checks: {p3['dq'].loc['Fully duplicated rows', 'rows']:,} identical rows (no booking ID,
can be real group bookings: kept and flagged), {p3['dq'].loc['Zero nights', 'rows']} zero-night stays and
{p3['dq'].loc['Negative ADR', 'rows']} negative rate excluded from price statistics.</div></div>"""
    pages.append(page_html(f"Travel agents bring {pct(ta['share_of_bookings'])} of bookings but cancel "
                           f"{ta['cancel_rate'] / di['cancel_rate']:.1f}× more often than direct guests, and ten agents make "
                           f"{pct(p3['top10'])} of agent revenue",
                           "119,390 real bookings by distribution channel. Two hotels in Portugal, 2015-2017, used to demonstrate the "
                           "analysis, not ETG data.", "REAL DATASET", body,
                           "Data: Antonio, Almeida & Nunes (2019), Hotel booking demand datasets, Data in Brief 22:41-49, CC BY 4.0 · src/analysis_real.py",
                           "grid-template-columns: 1fr 1fr 1fr 1fr; grid-template-rows: 1fr 150px"))

    # ------------------------------------------------------------ 4 bridge real
    seg = p4["seg"]
    rows = "".join(f"<tr><td>{r.hotel} · {r.market_segment}</td><td class='num'>{r.N0:,.0f} → {r.N1:,.0f}</td>"
                   f"<td class='num'>{'' if pd.isna(r.P0) else f'€{r.P0:.0f}'} → {'' if pd.isna(r.P1) else f'€{r.P1:.0f}'}</td>"
                   f"<td class='num {'red' if r.total_change < 0 else 'blue'}'>{'+' if r.total_change >= 0 else '−'}{keur(abs(r.total_change))}</td></tr>"
                   for r in seg.itertuples())
    body = f"""<div class="card">{img(p4['w'], 'height:420px')}<div class="small">Volume = change in room nights at August's
average rate. Mix = nights moving between cheaper and dearer segments. Price = rate change inside each segment.
The three add up exactly to the change.</div></div>
<div class="card"><h3>Drill-down: hotel × market segment, largest contributors</h3><table><tr><th>Segment</th>
<th class="num">Room nights</th><th class="num">ADR</th><th class="num">Change</th></tr>{rows}</table>
<div style="margin-top:10px">{img(p4['adr'], 'height:250px')}</div></div>
<div class="card" style="grid-column: span 2"><div class="small" style="font-size:13.5px"><b>Causes, top-down.</b>
Total {keur(-drop)} (−{pct(-drop / br[0])}) → the resort hotel explains {pct(p4['resort'] / drop)} of it → room nights only fell from
{p4['N0']:,.0f} to {p4['N1']:,.0f} (volume {keur(br[1])}), because a group block refilled the resort at about €70 a night
while online-agency and direct guests left after the summer peak (mix {keur(br[2])}); the resort's rates also dropped from
peak to shoulder season (price {keur(br[3])}). <b>Reading:</b> a normal end of season, made worse by low-rate group business.
<b>Action for a platform:</b> keep group rates above a floor in shoulder months, and push September offers to online-agency
partners in August.</div></div>"""
    pages.append(page_html(f"September 2016 revenue fell {keur(-drop)}, yet room nights barely moved: cheaper group business "
                           "and the end of summer prices explain most of the drop",
                           "A month-on-month revenue bridge (volume / mix / price) on real bookings, then a top-down search for the cause.",
                           "REAL DATASET", body, "Data: Hotel booking demand datasets (CC BY 4.0), arrival months, stays not cancelled · src/analysis_real.py",
                           "grid-template-columns: 1fr 1fr; grid-template-rows: 1fr 96px"))

    # ------------------------------------------------------------ 5 cockpit
    ex = p5["excel"]
    day = p5["day"]
    if "cockpit" in ex:
        left = (f'<div class="card" style="grid-row: span 3; padding:8px"><h3>Cockpit sheet (CUBE formulas), '
                f'15 Sep 2026</h3>{img(ex["cockpit"], "width:100%")}</div>')
    else:
        left = f"""<div class="card" style="grid-row: span 2"><h3>Cockpit, 15 Sep 2026 (values from the data model)</h3><table>
<tr><th>Metric</th><th class="num">Day</th><th class="num">Expected*</th><th class="num">vs exp.</th><th class="num">Same day LY</th></tr>
<tr><td>Bookings</td><td class="num">{day['bookings']:,.0f}</td><td class="num">{day['bookings_expected']:,.0f}</td><td class="num">{day['bookings'] / day['bookings_expected'] - 1:+.0%}</td><td class="num">{day['bookings_ly']:,.0f}</td></tr>
<tr><td>Sales (EUR)</td><td class="num">{day['sales_eur']:,.0f}</td><td class="num">{day['sales_eur_expected']:,.0f}</td><td class="num">{day['sales_eur'] / day['sales_eur_expected'] - 1:+.0%}</td><td class="num">{day['sales_eur_ly']:,.0f}</td></tr>
<tr><td>Margin rate</td><td class="num">{day['margin_rate']:.1%}</td><td class="num">{day['margin_rate_expected']:.1%}</td><td class="num">{(day['margin_rate'] - day['margin_rate_expected']) * 100:+.1f} pt</td><td class="num">{day['margin_rate_ly']:.1%}</td></tr>
</table><p class="small">* same weekday, previous 4 weeks.</p></div>"""
    al = p5["alerts"]
    months = (pd.to_datetime(al["date"]).max() - pd.to_datetime(al["date"]).min()).days / 30.4
    alerts_img = img(ex["alerts"], "width:100%") if "alerts" in ex else ""
    bridge_img = img(ex["bridge"], "width:100%") if "bridge" in ex else img(p5["bridge"], "height:250px")
    body = f"""{left}
<div class="card"><h3>Three layers, because one alert rule cannot catch everything</h3><div class="small" style="font-size:14px">
<b>1 · Daily band</b> (Cockpit, Alerts): each segment vs the same weekday of the last 4 weeks, alert at 3.5 standard
deviations of its normal noise. {len(al)} alerts in {months:.0f} months (~{p5['per_month']:.0f} a month); it flagged the
supplier margin drop (E1), the UAE holiday (E4) and the Greek cancellations (E6) on day one.<br>
<b>2 · Weekly drill-down</b> (Pivots, sql/04): a partner of 10 bookings a day leaving (E2) stays inside daily noise
but tops the weekly list of drops.<br>
<b>3 · Monthly bridge</b> (Bridge): a currency drift of 1-2% a month (E3) never fires a daily alert; it shows as the FX bar.</div></div>
<div class="card" style="padding:8px"><h3>Alerts sheet: UAE bookings, alert band and the two holiday alerts</h3>{alerts_img}</div>
<div class="card" style="padding:8px"><h3>Bridge sheet: Sep 2025 → Sep 2026 margin (check cell = 0)</h3>{bridge_img}</div>"""
    pages.append(page_html("One screen answers \"is today normal?\" by brand, market, supplier and partner; the weekly and monthly "
                           "layers catch what a daily alert cannot",
                           "Excel workbook: Power Pivot model with DAX measures, a cockpit built with CUBE formulas, PivotTables with "
                           "slicers, a statistical alert band and a variance bridge. Screens from the workbook.", "SIMULATED", body,
                           "Simulated ledger (seed 42), not ETG data · excel/ETG_Commercial_Pricing_Cockpit.xlsx",
                           "grid-template-columns: 1fr 1.05fr; grid-template-rows: auto 1fr 1fr"))

    # ------------------------------------------------------------ 6 root causes
    cards = []
    for e in ["E3", "E1", "E2", "E4", "E6"]:
        r = inc.loc[e]
        imp = r["impact_eur"]
        cards.append(f"""<div class="card inc"><div class="t">{e} · {html.escape(r['title'])}</div>
<div class="imp {'red' if imp < 0 else 'blue'}">{'+' if imp >= 0 else '−'}{eur(abs(imp))}</div>
<b>What:</b> {html.escape(r['what'])}<br><b>Evidence:</b> {html.escape(r['evidence'])}<br>
<b>Do:</b> {html.escape(r['recommendation'])}<br><b>Owner:</b> {html.escape(r['owner'])} · <b>Follow-up:</b> {html.escape(r['follow_up'])}</div>""")
    cards.append("""<div class="card inc"><div class="t">E5 · Data-quality issues (page 7)</div>
<div class="imp" style="color:#52514E">checked first</div>Duplicated load, missing supplier costs and a time-zone shift
were found and fixed <b>before</b> any of the analyses above. Owner: data engineering · Follow-up: automated checks
(sql/06) run before the daily report.</div>""")
    body = "".join(cards) + f"""<div class="card">{img(p6['e1'], 'width:100%')}</div><div class="card">{img(p6['e2'], 'width:100%')}</div>"""
    e3 = -inc.loc["E3", "impact_eur"]
    pages.append(page_html(f"Six events traced to their cause: the costliest was the quietest, lira price lists lagging the "
                           f"exchange rate ({keur(e3)} of margin)",
                           "Each incident card: what happened, the evidence, the impact in euros (margin), the recommendation, "
                           "the owner and the follow-up. Impacts measured from the event start to 31 Oct 2026.", "SIMULATED", body,
                           "Simulated ledger with programmed events; FX effect uses real ECB rates · src/analysis.py incident_numbers()",
                           "grid-template-columns: repeat(4, 1fr); grid-template-rows: 1fr 1fr"))

    # ------------------------------------------------------------ 7 DQ
    dq = p7["dq"]
    rows = "".join(f"<tr><td>{html.escape(str(r.issue))}</td><td style='white-space:nowrap'>{'' if pd.isna(r.date_from) else r.date_from}</td>"
                   f"<td>{html.escape(str(r.scope))}</td><td class='num'>{r.rows:,}</td><td>{html.escape(str(r.effect))}</td>"
                   f"<td>{html.escape(str(r.action))}</td></tr>" for r in dq.head(4).itertuples())
    rc = p7["rc"].set_index("booking_date")
    dup = rc.loc["2026-07-21"]
    body = f"""<div class="card" style="grid-column: span 2">{img(p7['chart'], 'height:285px')}</div>
<div class="card">{img(p7['tz'], 'height:285px')}</div>
<div class="card" style="grid-column: span 3"><table><tr><th>Issue</th><th>From</th><th>Scope</th><th class="num">Rows</th>
<th>Effect if not caught</th><th>Action</th></tr>{rows}</table><p class="small" style="margin-top:8px">Checks that returned
zero rows: selling price ≤ 0, check-in before booking, cancellation before booking, missing FX rate. Without the checks:
21 July would have shown a {pct(dup['sales_eur_raw'] / dup['sales_eur_clean'] - 1)} sales "record" and the first week of
August a margin rate near 12%; the time shift moves late bookings to the next day and breaks any hour-of-day or
month-end cut-off analysis. Sold-below-cost bookings are a <b>business</b> alert (flight fares in falling currencies,
the Spanish supplier rate rise), sent to pricing, not deleted.</p></div>"""
    pages.append(page_html(f"Three data issues would have produced three wrong conclusions: a {pct(dup['sales_eur_raw'] / dup['sales_eur_clean'] - 1)} "
                           "sales spike, an inflated margin and bookings on the wrong day",
                           "Checks run on the raw ledger before any number is read (src/analysis.py dq_checks, sql/06_data_quality_checks.sql, "
                           "Excel Data_Quality sheet).", "SIMULATED", body, "Simulated ledger with injected data issues · data/output/sim/dq_issues.csv",
                           "grid-template-columns: 1fr 1fr 1fr; grid-template-rows: 315px 1fr"))

    # ------------------------------------------------------------ 8 price match
    by = p8["by"]
    body = f"""<div class="card">{img(p8['chart'], 'height:400px')}<div class="small">{s['requests']:,.0f} requests in 18 months
(~0.7% of bookings). Evidence first, then like-for-like: taxes added back, breakfast valued at €12 per guest-night,
free cancellation worth 8% of the price. Then the margin floor decides.</div></div>
<div class="card"><h3>Decision rules (playbook/price_match_rules.md, Excel PriceMatch_Calc)</h3><div class="small">
<b>1.</b> Evidence complete? Same hotel, room, dates, screenshot under 48 h. If not → <b>ask for details</b>.<br>
<b>2.</b> Make it like-for-like: competitor price + missing taxes + our breakfast + our flexibility.<br>
<b>3.</b> Real gap ≤ 0 → <b>refuse</b>, with the reasons. <b>4.</b> Gap ≤ maximum discount (price − net cost / (1 − 4%)) →
<b>discount</b> the gap. <b>5.</b> Else, a cheaper supplier keeps the floor → <b>switch supplier</b>. <b>6.</b> Else → <b>refuse</b>:
the floor is not negotiable, offer alternatives.</div>
<table style="margin-top:8px"><tr><th>Decision</th><th class="num">Share</th><th class="num">Booked after</th><th class="num">Median handling</th></tr>
{''.join(f"<tr><td>{d}</td><td class='num'>{by.loc[d, 'share']:.0%}</td><td class='num'>{by.loc[d, 'booked_rate']:.0%}</td><td class='num'>{by.loc[d, 'avg_handling_h']:.1f} h</td></tr>" for d in ['Discount', 'Switch supplier', 'Refuse', 'Ask for details'])}
</table><p class="small">Discounts granted on booked requests cost {eur(s['discount_cost_eur'])}; discount + switch kept
{eur(s['saved_sales_eur'])} of sales.</p></div>
<div class="card"><h3>Saying no, with arguments (templates)</h3>
<div class="quote"><b>Not the same product.</b> "Thank you for flagging this. We checked the offer: the competitor rate is
non-refundable and excludes taxes, ours can be cancelled free of charge and includes taxes. Compared like-for-like,
our price is the better offer, so we keep it."</div>
<div class="quote"><b>Margin floor.</b> "We checked carefully: the difference comes from the supplier cost for these dates and
we cannot go lower without selling at a loss. We can offer two similar hotels nearby at a better price."</div>
<div class="quote"><b>Missing evidence.</b> "To review the price, please send a screenshot from the last 48 hours with
hotel, room type, dates, meal plan, cancellation policy and total with taxes."</div></div>"""
    pages.append(page_html(f"{pct(s['valid_share'])} of price-match claims survive a like-for-like check; the margin floor turns "
                           f"half of all requests into a reasoned no",
                           "Internal requests from partners who saw a lower competitor price: check, decide, reply and track.",
                           "SIMULATED", body, "Simulated requests (competitor names are generic) · Excel PriceMatch_Log rule check: every row matches",
                           "grid-template-columns: 1.15fr 1fr 1fr; grid-template-rows: 1fr"))

    # ------------------------------------------------------------ 9 pricing
    ch = p9["changes"]
    rows = "".join(f"<tr><td>{r.partner_type} · {r.destination_country}</td><td class='num'>{r.margin_rate:.1%}</td>"
                   f"<td class='num'>{r.valid_requests_per_1000:.1f}</td><td class='num'>{r.yoy_volume:+.0%}</td>"
                   f"<td class='num {'blue' if r.markup_change > 0 else 'red'}'>{r.markup_change * 100:+.1f} pt</td>"
                   f"<td class='num'>{signed_eur(r.margin_effect_eur_high)} to {signed_eur(r.margin_effect_eur_low)}</td></tr>"
                   for r in ch.itertuples())
    hold = p9["r"][p9["r"]["recommendation"].str.startswith("Hold")]
    body = f"""<div class="card">{img(p9['chart'], 'height:600px')}</div>
<div class="card"><h3>Recommended moves (90-day effect, elasticity −1 to −3)</h3><table><tr><th>Segment</th>
<th class="num">Margin</th><th class="num">Valid PM / 1,000</th><th class="num">Volume YoY</th><th class="num">Markup</th>
<th class="num">Margin effect</th></tr>{rows}</table>
<p class="small"><b>Hold and renegotiate:</b> {', '.join(hold['partner_type'] + ' · ' + hold['destination_country'])}
(margin {', '.join(f'{m:.1%}' for m in hold['margin_rate'])}, within 1 pt of the floor after the supplier rate rise).<br>
<b>Read the −0.5 pt carefully:</b> it loses margin at any plausible elasticity. With a
{ch[ch['markup_change'] < 0]['markup'].iloc[0]:.1%} markup, a 0.5 pt cut needs
+{ch[ch['markup_change'] < 0]['markup'].iloc[0] / (ch[ch['markup_change'] < 0]['markup'].iloc[0] - 0.005) - 1:.1%}
volume just to break even. It is worth it only to keep partners from leaving (the E2 risk); otherwise re-source cheaper
supply first.<br>
<b>Rules:</b> −0.5 pt where valid price-match pressure is in the top quartile and growth below 10%; +0.5 pt where pressure
is low and growth above 10%; never below a 4% margin floor. Elasticity is an assumption: the effect is a range, total
{signed_eur(p9['high'])} to {signed_eur(p9['low'])} over 90 days. Small on purpose: markups move by half points, and only where the data
supports it.</p>
<h3 style="margin-top:12px">Follow-up plan</h3><div class="small">Week 0: apply changes on 5% of traffic (A/B by partner) ·
Weeks 1-4: daily cockpit on test vs control (bookings, margin, price-match requests) · Week 4: keep, extend or roll back;
re-estimate elasticity from the test instead of the assumption.</div></div>"""
    pages.append(page_html("Move markups by half points where the data supports it, and hold the segments squeezed by the supplier "
                           "rate rise",
                           "B2B hotel segments (partner type × destination), last 90 days: margin, valid price-match pressure and growth "
                           "decide the move; the expected effect is a range, not a point.", "SIMULATED", body,
                           "Simulated ledger; elasticities are assumptions to be validated by a price test · src/analysis.py pricing_recommendations()",
                           "grid-template-columns: 1fr 1.1fr; grid-template-rows: 1fr"))

    # ------------------------------------------------------------ 10 methods
    body = f"""<div class="card"><h3>Sources and licences</h3><div class="small">
<b>ETG public pages</b> emergingtravel.com, ratehawk.com, zenhotels.com (live, 12 Nov 2026) and six archived home pages
(web.archive.org, 2021-2026). robots.txt checked, one request every 2.5 s, no login, no prices collected.<br>
<b>Eurostat</b> tour_occ_ninraw, tour_dem_tnmd, tour_occ_nim, Eurostat re-use policy (attribution).<br>
<b>ECB</b> euro foreign exchange reference rates (EXR, daily), ECB Data Portal, re-use with attribution.<br>
<b>Hotel booking demand datasets</b> N. Antonio, A. de Almeida, L. Nunes, Data in Brief 22 (2019) 41-49,
doi:10.1016/j.dib.2018.11.126, CC BY 4.0 (TidyTuesday mirror, 2020-02-11).<br>
<b>Competitor prices</b> are never collected: all price-match comparisons are simulated.<br>
Full list with dates: data/SOURCES.md.</div></div>
<div class="card"><h3>Simulation and assumptions</h3><div class="small">
18 months (May 2025 - Oct 2026), daily, ~350 bookings a day (about 1:50 of ETG's public "130k+ bookings weekly"), seed 42.
Brands, lines of business, 13 source markets, 10 destinations, 27 fictional suppliers, 2,138 fictional partners.
Markups 6.5-14% by partner type (assumption); local price lists refreshed monthly; real ECB rates; Eurostat seasonality.
Seven programmed events (supplier rate, lost partner, FX, holiday, data issues, cancellations, price matches), all found
by the monitoring. All assumptions: data/ASSUMPTIONS.md.</div>
<h3 style="margin-top:12px">Reproduce</h3><div class="small">python src/collect_*.py → simulate.py → analysis_real.py →
analysis.py → run_sql.py → build_excel.py (Windows + Excel 365) → build_report.py. README.md lists every step.</div></div>
<div class="card"><h3>Limits</h3><div class="small">
The real booking dataset is old (2015-2017) and covers two hotels in one country: it shows methods, not today's market.
The ledger is simulated: patterns are realistic by construction, so the findings show the process, not ETG's results.
Public ETG figures are marketing headlines ("+") that differ between pages. Elasticities are assumptions. Eurostat covers
EU destinations and EU residents' trips only.</div>
<h3 style="margin-top:12px">AI note</h3><div class="small">{AI_NOTE}</div></div>"""
    pages.append(page_html("Methods, limits &amp; AI note", "How the data was collected, what is assumed, and what this study cannot say.",
                           "METHODS", body, "Sources and collection dates in data/SOURCES.md", "grid-template-columns: 1fr 1fr 1fr; grid-template-rows: 1fr"))
    return pages


def find_browser() -> str:
    for p in [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
              r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
              r"C:\Program Files\Google\Chrome\Application\chrome.exe"]:
        if Path(p).exists():
            return p
    raise FileNotFoundError("Edge or Chrome is required to print the PDF")


def main() -> None:
    IMG.mkdir(parents=True, exist_ok=True)
    pages = build_pages()
    head = f'<!doctype html><html lang="en"><head><meta charset="utf-8"><title>ETG Commercial Pricing Case Study</title><style>{CSS}</style></head><body>'
    full = BUILD / "report.html"
    full.write_text(head + "".join(pages) + "</body></html>", encoding="utf-8")
    browser = find_browser()
    tmp_pdf = BUILD / "report.pdf"
    subprocess.run([browser, "--headless", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={tmp_pdf}",
                    full.as_uri()], check=True, capture_output=True, timeout=180)
    # Clean PDF metadata (no author, neutral producer)
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(str(tmp_pdf))
    writer = PdfWriter()
    for pg in reader.pages:
        writer.add_page(pg)
    writer.add_metadata({"/Title": "Emerging Travel Group - commercial & pricing case study", "/Author": "",
                         "/Creator": "", "/Producer": "", "/Subject": "Independent case study based on public data"})
    with open(PDF, "wb") as fh:
        writer.write(fh)
    print(f"PDF: {PDF} ({len(reader.pages)} pages)")
    # One PNG per page for the README
    for i, pg in enumerate(pages):
        f = BUILD / f"page_{i:02d}.html"
        f.write_text(head + pg + "</body></html>", encoding="utf-8")
        subprocess.run([browser, "--headless", "--disable-gpu", "--hide-scrollbars", "--window-size=1600,900",
                        f"--screenshot={IMG / f'page_{i:02d}.png'}", f.as_uri()], check=True, capture_output=True, timeout=120)
    shutil.rmtree(BUILD, ignore_errors=True)
    print("PNG pages:", len(pages))


if __name__ == "__main__":
    main()
