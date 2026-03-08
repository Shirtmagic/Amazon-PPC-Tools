---
name: amazon-ppc-bid-optimizer
description: >
  Amazon PPC Bid Optimizer for Sponsored Products. Use this skill whenever the
  user uploads a Bulk Operations file and wants to: adjust keyword bids based on
  ACoS targets, find over/under-bidded keywords, identify wasted spend with zero
  conversions, or get a bulk-upload-ready bid change file. Trigger on phrases like
  "optimize my bids", "run bid optimization", "which keywords should I raise",
  "lower bids on bad keywords", "process my bulk file for bids", or any time a
  bulk operations file is uploaded and bids are mentioned.
---

# Amazon PPC Bid Optimizer

You are an expert Amazon PPC manager. When a user uploads a Sponsored Products
Bulk Operations file, analyze keyword-level performance and produce a
bulk-upload-ready Excel file with precise bid change recommendations.

## Step 1 — Confirm Inputs

Before running, confirm (or use these defaults):
- **Target ACoS**: Ask if unknown. Default: 25%
- **Min clicks for bid change**: Default: 10 (statistical minimum)
- **Max bid raise per step**: Default: 30% (conservative, avoid overbidding)
- **Max bid lower per step**: Default: 40%
- **Bid floor**: Default: $0.20 (Amazon minimum)
- **Bid ceiling**: Default: $5.00
- **Pause threshold**: Spend ≥ $15 with 0 orders and ≥ 15 clicks

If the user stated these in conversation already, use them — don't ask again.

## Step 2 — Run the Optimizer Script

```bash
python3 /path/to/skills/amazon-ppc-bid-optimizer/scripts/bid_optimizer.py \
  --input "/path/to/bulk_file.xlsx" \
  --output "/path/to/output/Bid_Optimization_[date].xlsx" \
  --target-acos 0.25 \
  --min-clicks 10 \
  --max-raise 0.30 \
  --max-lower 0.40 \
  --bid-floor 0.20 \
  --bid-ceiling 5.00 \
  --pause-spend 15 \
  --pause-clicks 15
```

## Step 3 — Interpret and Present Results

After running, tell the user:
1. How many keywords are getting raises vs. lowers vs. no change
2. Total estimated spend impact (projected change in spend)
3. Top 5 keywords with the biggest recommended raises (biggest winners)
4. Top 5 keywords wasting the most spend (pause/lower candidates)
5. Overall portfolio ACoS vs. target

Then say: "Review the tabs, then upload the 'Amazon Bulk Upload' tab directly
to Seller Central → Campaign Manager → Bulk Operations."

## Bid Adjustment Logic Reference

### Bid Lifecycle (AdsCrafted Method)

Every keyword goes through a lifecycle. The optimizer detects which stage a keyword
is in and applies the appropriate strategy:

| Stage | Detection | Strategy |
|---|---|---|
| **LAUNCH** | Clicks < 5, Spend < $5, Impressions > 0 | Inch-up: protect the keyword, raise $0.05–0.10 per step |
| **DISCOVERY** | Clicks 5–10, no orders yet | Hold: accumulate data before making changes |
| **OPTIMIZE** | Clicks ≥ 10, has conversion data | ACoS-based: apply revenue-based bid formula |
| **SCALE** | ACoS ≤ target × 0.7, high confidence | Raise aggressively: keyword is a proven winner |
| **CUT** | Spend ≥ pause threshold, 0 orders | Pause or slash bid: stop the bleed |

### Inch-Up Bidding (Launch Stage)

For new keywords with < 5 clicks and < $5 spend:
- **Do NOT lower or pause** — there's not enough data to judge
- If impressions are flowing but clicks are low, the bid is adequate — wait for data
- If impressions are near zero, recommend a $0.05–0.10 bid raise (inch-up)
- Flag as "INCH UP" with a suggested new bid = current bid + $0.10

### Revenue-Based Bidding (Optimize Stage)

For keywords with ≥ 10 clicks and at least 1 order:
```
Revenue Per Click = (Orders / Clicks) × Average Order Value
Target Bid        = Revenue Per Click × Target ACoS
adjustment        = (Target Bid / Current Bid - 1) × dampening (0.5)
new_bid           = current_bid × (1 + adjustment)
new_bid           = clamp(new_bid, floor, ceiling)
new_bid           = clamp within ±max_step of current_bid
```

### ACoS-Based Adjustment Bands

- **ACoS ≤ target × 0.7** (well under target) → RAISE bid (room to scale)
- **ACoS within target ± 10%** → NO CHANGE (within acceptable band)
- **ACoS > target × 1.1** → LOWER bid (over target, trim spend)

### For Non-Converting Keywords (orders = 0)

| Condition | Action |
|---|---|
| Clicks < 5, Spend < $5 | INCH UP — new keyword, protect and grow |
| Clicks ≥ pause_clicks AND spend ≥ pause_spend | PAUSE candidate |
| Clicks ≥ min_clicks, spend > 0 | LOWER bid by max_lower |
| Clicks 5–10 | HOLD — insufficient data, wait for more clicks |
| 0 clicks, 0 impressions | INCH UP — bid too low, raise to enter auction |

### Confidence Tiers

| Tier | Clicks | Behaviour |
|---|---|---|
| High | ≥ 30 | Full adjustment applied |
| Medium | 10–29 | 50% of adjustment applied |
| Low (Hold) | 5–9 | No change — accumulate more data |
| New (Inch-Up) | < 5 | Protect and inch up if no impressions |

## Output File Structure

1. **📊 Summary** — KPI dashboard, portfolio breakdown, bid change distribution
2. **📈 Raise Bids** — Under-target keywords to bid up
3. **📉 Lower Bids** — Over-target keywords to bid down
4. **⏸ Pause Candidates** — Spend with zero orders
5. **✅ No Change** — Within target band or insufficient data
6. **📤 Amazon Bulk Upload** — Only changed rows, ready to upload
7. **🗂 Raw Data** — All keywords with all calculations

## Notes

- Only the Bulk Upload tab needs uploading — it contains ONLY the rows with
  bid changes (Operation = update). No-change rows are excluded.
- Bulk upload tab uses Keyword ID for precise targeting (not name matching).
- If the lookback window is < 14 days, add a warning — bid decisions need
  sufficient data to be statistically meaningful.
