# Price-match requests: checklist and decision rules

A **price-match request** arrives when a partner (travel agent, API partner, corporate client) reports that a
competitor sells the same room or service for less. The job: check the claim, decide quickly, defend the margin,
and keep the partner. The same rules run in three places, so they always give the same answer:

- `src/simulate.py` → `decide()` (the 1,347 simulated requests),
- Excel `PriceMatch_Calc` (one request at a time) and `PriceMatch_Log` (the "Rule check" column re-applies them to every row),
- this page.

> Competitor prices in this project are **simulated**. No competitor website was scraped.

## 1. Evidence checklist (all must be "yes", otherwise ask for details)

| # | Check | Why |
|---|---|---|
| 1 | Same hotel, same room type (and bed type) | A different room is a different product |
| 2 | Same check-in and check-out dates, same number of guests | Prices change by night and occupancy |
| 3 | Screenshot provided, taken less than **48 hours** ago | Rates move fast; old screenshots are not evidence |
| 4 | Total price visible, currency visible | Avoid comparing a nightly rate with a total, or EUR with USD |

## 2. Make the offers like-for-like

Start from the competitor price and add what our offer includes and theirs does not:

| Difference | Adjustment added to the competitor price |
|---|---|
| Competitor price **excludes taxes and fees** | + 10% of the competitor price (city tax, VAT, resort fee; use the real amount when known) |
| Our rate **includes breakfast**, theirs is room only | + EUR 12 per guest per night (2 guests by default) |
| Our rate has **free cancellation**, theirs is **non-refundable** | + 8% of our price (value of the flexibility) |
| Different currency | Convert at the ECB reference rate of the screenshot date |

`Real gap = our price − like-for-like competitor price`

## 3. Decision rules

```
Floor price     = net cost / (1 − margin floor)          margin floor = 4% of the selling price
Max discount    = our price − floor price                 (never negative)

if evidence incomplete            -> ASK FOR DETAILS
elif real gap <= 0                -> REFUSE  (not the same product: explain the differences)
elif real gap <= max discount     -> DISCOUNT by the real gap (not more)
elif alt. supplier net / (1 − floor) <= like-for-like price -> SWITCH SUPPLIER and match
else                              -> REFUSE  (margin floor) and offer alternatives
```

Worked example (request PM00009, the Excel `PriceMatch_Calc` default): our price EUR 462.37 with free cancellation,
competitor EUR 444.34 non-refundable (same room, same breakfast, taxes included). Like-for-like competitor price =
444.34 + 8% × 462.37 = EUR 481.33, above our price: the real gap is −18.96, so **refuse** and say why (their rate cannot
be cancelled). Note that a discount was impossible anyway: net cost 445.35 means the floor price is 445.35 / 0.96 =
EUR 463.91, already above our selling price.

### Saying no under pressure

- The **margin floor is not negotiable** in a single request. Exceptions (strategic partner, volume deal) go to the
  commercial manager with numbers: expected volume, margin at the requested price, duration.
- Always give the **reason** (taxes, breakfast, flexibility, supplier cost) and an **alternative** (similar hotel,
  other dates, other room type).
- Never discount more than the real gap; never match a screenshot older than 48 hours.

## 4. What we track (Excel `PriceMatch_Log`, `sql/05_price_match_summary.sql`)

| KPI | Definition |
|---|---|
| Requests | count, by partner type and competitor |
| Valid share | requests with complete evidence and a positive like-for-like gap |
| Decision mix | discount / switch supplier / refuse / ask for details |
| Handling time | hours from request to answer (target: median under 4 h) |
| Discount cost | sum of discounts on requests that were booked |
| Sales saved | value of bookings made after a discount or a supplier switch |
| Booked after refusal | share of refused requests that still booked (shows if refusals are well argued) |

A segment with many **valid** requests is a pricing signal (see `Pricing` sheet): our rates are really above the
market there, so we look at the supplier or the markup instead of granting discounts one by one.
