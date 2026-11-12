# Assumptions of the simulated ledger

Everything in `data/simulated/` is **SIMULATED** by `src/simulate.py` (seed 42). It is **not Emerging Travel Group
data**. The goal is a ledger realistic enough to practise the job: daily monitoring by segment, root causes, data
quality, price-match requests and pricing. Values below are assumptions unless marked *public* or *real*.

## Scale and period

| Item | Value | Basis |
|---|---|---|
| Period | 1 May 2025 - 31 Oct 2026 (18 months), daily | so that the last 6 months have a "same day last year" |
| Volume | 300 bookings a day at the start, +18% a year → ~350 a day on average, 192,489 unique bookings | about 1:50 of the *public* "130k+ bookings weekly" (emergingtravel.com/partners) |
| Average booking value | ~€350 | consistent with *public* $2.6BN+ annual turnover / ~6.8M bookings a year (order of magnitude only) |
| Weekly pattern | Mon-Thu above average, Saturday lowest | B2B is a weekday business |
| Booking seasonality | per source market, from *public* Eurostat month-of-departure profiles (shifted one month earlier, smoothed) | `tour_dem_tnmd`; generic profiles for non-EU markets |
| Destination seasonality | from *public* Eurostat non-resident nights by month (2025) | `tour_occ_nim`; generic profiles for TR, AE, TH |

## Dimensions

| Dimension | Values |
|---|---|
| Brands | RateHawk (agencies, API partners, white label), ZenHotels (B2C), Roundtrip (corporate) |
| Partner types (share of bookings) | Agency 45%, API partner 30%, B2C 14%, Corporate 6%, White label 5% |
| Lines of business | RateHawk: hotels 85%, transfers 6%, flights 6%, car rentals 3%; ZenHotels: hotels 100%; Roundtrip: hotels 70%, flights 25%, transfers 3%, cars 2% |
| Source markets (13) and currency | DE, FR, IT, ES, NL (EUR), GB (GBP), PL (PLN), TR (TRY), AE and US (USD - Gulf B2B often settles in USD; the ECB publishes no AED rate), BR (BRL), MX (MXN), IN (INR) |
| Destinations (10) | ES, IT, FR, EL, PT, DE, HR, TR, AE, TH; choice = market preference × destination seasonality at check-in |
| Suppliers (27, fictional) | 10 "Direct contracts" (one per destination, 30-day terms, net rates 4% below market), 6 bedbanks, 5 regional wholesalers, 2 transfer, 2 car rental, 2 flight consolidators |
| Partners (2,138, fictional) | 150 agencies and 10 corporate clients per market, 30 API partners, 15 white-label partners, one B2C account per market; size of each partner drawn from a long-tailed (log-normal) distribution; "Top partner" = top 5% by sales |

## Prices, costs and margins

| Item | Assumption |
|---|---|
| Hotel net cost | base net rate per room-night by destination (€50 Thailand to €125 UAE) × seasonal factor (0.75-1.25) × supplier factor × random noise (log-normal, sd 0.25) × room nights (1 + Poisson(2.2); corporate 1 + Poisson(1.0)) |
| Markup on net (hotels) | Agency 9.5%, White label 8%, API partner 6.5%, Corporate 10%, B2C (ZenHotels) 14%, ± 1.2 pt noise |
| Other lines | transfers €45 net, +20%; car rentals €180, +12%; flights €280, +3.5% |
| Currency | suppliers paid in EUR; partners pay in their currency; **local price lists are refreshed monthly at the ECB rate of the 1st of the previous month**; the EUR value of a sale uses the *real* ECB rate of the booking day |
| Margin | selling price − discount − supplier net cost (EUR). Margin floor used in the rules: **4% of the selling price** |
| Lead time | gamma distribution, mean 14 days (corporate) to 38 days (agencies) |
| Cancellations | base rate by partner type (15-25% for hotels, 8% for other lines), +8 pt when booked more than 60 days ahead; cancellation dates cluster near check-in; status as seen on 31 Oct 2026 |

## Programmed events (what the monitoring must find)

| # | Event | Parameters |
|---|---|---|
| E1 | Supplier net-rate increase | Bedbank C raises net rates **+7% in Spain** from 15 Jun 2026; selling prices unchanged (set by the market) |
| E2 | Large API partner lost | API-002 (largest API partner, UK) ramps down 7-14 Sep 2026 (60% of its bookings lost), then stops |
| E3 | Currency effect | not programmed: the **real** depreciation of the Turkish lira (ECB rates) against monthly local price lists |
| E4 | Source-market holiday | UAE bookings +70% during the two weeks before Eid al-Adha, in 2025 (22 May - 5 Jun) and 2026 (11-25 May) |
| E5 | Data-quality issues | (a) all API bookings of 21 Jul 2026 loaded twice; (b) net cost missing for Wholesaler A, 3-9 Aug 2026; (c) API feed timestamps shifted +5 h, 12-18 Oct 2026 |
| E6 | Cancellation spike | Greece, stays 12-25 Aug 2026: 30% of still-active bookings cancelled between 8 and 14 Aug 2026 (a local event such as severe weather) |
| E7 | Price-match requests | 1,347 requests (~0.7% of bookings) on hotel bookings of agencies, API partners and corporate clients: 42% same product, 16% competitor non-refundable, 12% competitor without breakfast, 12% competitor price without taxes, 10% missing evidence, 8% screenshot older than 48 h; claimed gap 1-15%; a cheaper alternative supplier exists in 35% of cases. Decisions follow `playbook/price_match_rules.md`; probability of booking after the decision: discount 82%, switch 80%, ask 45%, refuse 35% |

`data/simulated/events_truth.csv` stores the event parameters (the answer key). It is used **only** to score the
alerts afterwards (`matched_event` column of `data/output/sim/alerts.csv`), never by the detection itself.

## What the simulation does not try to reproduce

ETG's real brand mix, margins, supplier names, partner concentration, competitor prices, or cancellation policies.
Price elasticities used on the Pricing page (−1 to −3) are assumptions to be replaced by a price test.
