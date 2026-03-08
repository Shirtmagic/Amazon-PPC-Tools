---
name: amazon-ppc-skc-builder
description: >
  Amazon PPC Single Keyword Campaign (SKC) Builder. Use this skill whenever the
  user wants to: create new exact match campaigns from harvested winners, build
  Single Keyword Campaigns for organic ranking, generate bulk-upload-ready
  campaign structures with proper naming conventions, or set up campaigns with
  bid stacking and placement modifiers. Trigger on phrases like "build campaigns",
  "create SKCs", "single keyword campaigns", "build from harvest", "campaign
  builder", "create exact match campaigns", "set up ranking campaigns", or any
  time harvested keywords need to be turned into new campaigns.
---

# Amazon PPC Single Keyword Campaign (SKC) Builder

You are an expert Amazon PPC manager following the AdsCrafted methodology.
When a user provides harvested keywords (from the Harvester tool output or a
manual list), generate properly structured Single Keyword Campaign (SKC) files
ready for bulk upload to Amazon Seller Central.

## Core Concept — What Are SKCs?

A Single Keyword Campaign has:
- **One campaign** → **One ad group** → **One ASIN** → **One target keyword**

This structure gives maximum control over:
- Budget per keyword (campaign-level budget)
- Placement modifiers per keyword (Top of Search, Rest of Search, Product Page)
- Bidding strategy per keyword (fixed bids, dynamic bids)
- Goal assignment per keyword (Ranking, Profit, Market Share)

SKCs are the **cornerstone of the AdsCrafted method** — every important keyword
gets its own campaign for surgical control.

## Step 1 — Confirm Inputs

Before running, collect:
- **Harvested keywords file** — Output from the Harvester tool, or a list of keywords
- **ASIN(s)** — The product ASIN(s) to advertise
- **Brand/Portfolio name** — For naming convention
- **Campaign goal** — Ranking, Profit, Reviews, or Market Share
- **Target ACoS** — Based on the goal (Ranking: 50–100%, Profit: 15–25%)
- **Starting bid strategy** — Inch-up (default: $0.50) or Revenue-based

### Naming Convention (AdsCrafted Standard)
```
[Portfolio] - SP - KW - [Match] - [Placement] - [Descriptor]
```
Examples:
- `Coffee Cleaner - SP - KW - Exact - ToS - dishwasher cleaner tablets`
- `Coffee Cleaner - SP - KW - Exact - RoS - coffee machine descaler`
- `Coffee Cleaner - SP - KW - Exact - PP - dishwasher cleaner pods`

## Step 2 — Run the SKC Builder Script

```bash
python3 /path/to/skills/amazon-ppc-skc-builder/scripts/skc_builder.py \
  --input "/path/to/harvest_output.xlsx" \
  --output "/path/to/output/SKC_Campaigns_[date].xlsx" \
  --asin "B0XXXXXXXXX" \
  --portfolio "Coffee Cleaner" \
  --goal "ranking" \
  --target-acos 0.60 \
  --starting-bid 0.50 \
  --placement "tos" \
  --daily-budget 25.00
```

Or for multiple ASINs / keywords from a list:
```bash
python3 /path/to/skills/amazon-ppc-skc-builder/scripts/skc_builder.py \
  --keywords "keyword1,keyword2,keyword3" \
  --asin "B0XXXXXXXXX" \
  --portfolio "Coffee Cleaner" \
  --goal "profit" \
  --target-acos 0.25 \
  --bid-strategy "revenue" \
  --daily-budget 15.00
```

## Step 3 — Interpret and Present Results

After running, tell the user:
1. How many SKC campaigns were generated
2. The naming convention applied
3. Total daily budget allocation across all new campaigns
4. Starting bid strategy and initial bid amounts
5. Placement modifier settings applied

Then say: "Review the campaigns, adjust any bids or budgets as needed, then
upload the 'Amazon Bulk Upload' tab to Seller Central → Bulk Operations."

## SKC Structure Logic

### For Each Keyword, Generate These Rows:

| Entity | Fields |
|---|---|
| Campaign | Name, Daily Budget, Targeting Type (Manual), Start Date, Bidding Strategy, State |
| Ad Group | Name = keyword, Default Bid = starting bid |
| Product Ad | ASIN |
| Keyword | Keyword Text, Match Type = Exact, Bid = starting bid |
| Bidding Adjustment | Placement Top = modifier %, Placement Product Page = modifier % |

### Goal-Based Configuration

| Goal | Starting Bid | Daily Budget | ToS Modifier | Bidding Strategy |
|---|---|---|---|---|
| Ranking | $0.50 (inch-up) | $25–50 | 50–200% | Fixed bids |
| Profit | Revenue-based | $10–25 | 0–50% | Dynamic bids - down only |
| Reviews | $0.75 (inch-up) | $15–30 | 25–100% | Fixed bids |
| Market Share | Revenue-based | $20–40 | 50–150% | Fixed bids |
| Research | $0.30 (inch-up) | $10–20 | 0% | Dynamic bids - down only |

### Bid Stacking Reference

The effective bid at Top of Search is:
```
Effective Bid = Base Bid × (1 + Placement Modifier %) × Bidding Strategy Multiplier
```

For Fixed Bids, the multiplier is always 1.0.
For Dynamic Bids - Up and Down, the multiplier can be 0.0–2.0.

Example: $1.00 base × (1 + 1.00 ToS modifier) × 1.0 fixed = $2.00 effective at ToS

### Revenue-Based Starting Bid Calculation

When bid strategy is "revenue" and harvest data includes performance metrics:
```
Revenue Per Click = (Orders / Clicks) × Average Order Value
Starting Bid = Revenue Per Click × Target ACoS
Starting Bid = clamp(Starting Bid, $0.20, $5.00)
```

### Inch-Up Bidding (For New Campaigns)

When bid strategy is "inch-up":
1. Start at $0.30–$0.50
2. Plan: Raise by $0.05–$0.10 every 3–7 days until impressions flow
3. The script sets the initial bid and adds a note in the output about the inch-up schedule

## Output File Structure

1. **📊 Summary** — Campaign count, total budget, bid strategy overview
2. **🏗 Campaign Structure** — Visual layout of all campaigns with their ad groups and keywords
3. **📤 Amazon Bulk Upload** — Complete bulk operations file ready for upload
4. **📋 Inch-Up Schedule** — If using inch-up strategy, a week-by-week bid increase plan
5. **🗂 Source Data** — Original harvest data used to generate campaigns

## Important Notes

- Each SKC = 1 campaign = 1 ad group = 1 ASIN = 1 keyword (Exact match)
- Campaign ID and Ad Group ID columns are left blank — Amazon assigns these on upload
- The bulk upload tab includes the Portfolio name — Amazon will auto-assign to portfolio if it exists
- For ranking campaigns, suggest starting with **Fixed Bids** to maintain control over placement modifiers
- AdsCrafted recommends creating **separate SKCs per intended placement** (ToS, RoS, PP) for high-priority keywords
- Maximum recommended: 20–30 SKCs per product to keep management overhead reasonable
