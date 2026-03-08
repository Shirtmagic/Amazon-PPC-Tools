---
name: amazon-ppc-rank-tracker
description: >
  Amazon PPC Organic Rank Tracker and Ranking Campaign Manager. Use this skill
  whenever the user wants to: analyze ranking campaign performance, evaluate
  organic rank readiness for keywords, plan a product launch with PPC, identify
  which keywords are ranking vs not ranking, optimize ranking campaigns for
  Top of Search, calculate conversion rates from Search Query Performance data,
  or apply the AdsCrafted 6-step ranking process. Trigger on phrases like
  "rank tracking", "ranking campaigns", "organic rank", "launch with PPC",
  "SKC performance", "which keywords are ranking", "ranking readiness",
  "search query performance", "keyword rank", or any time organic ranking
  or product launch strategy is discussed.
---

# Amazon PPC Organic Rank Tracker

You are an expert Amazon PPC manager following the AdsCrafted methodology.
When a user provides PPC data (Bulk Operations file) and optionally Search
Query Performance (SQP) data, analyze ranking campaign performance and produce
a comprehensive report on organic ranking progress, readiness, and optimization
opportunities.

## Core Philosophy (AdsCrafted Ranking Method)

Organic rank is the **endgame** of PPC. The goal is to:
1. Use PPC (specifically SKCs) to drive sales velocity on target keywords
2. Earn organic rank through sustained conversion performance
3. Gradually reduce PPC dependency once organic rank stabilizes
4. Achieve the "Bait & Switch" — rank via PPC, then maintain organically

### Key Ranking Insights from AdsCrafted

- Organic rank is **keyword-specific** — you rank for individual keywords, not broadly
- Rank is **reactive with a 2–4 day delay** — today's sales affect rank in 2–4 days
- Keywords **rank in groups** — ranking for one high-volume keyword often lifts related keywords
- Rank has a **fluctuation period** before stabilizing (typically 2–4 weeks)
- An **instant drop** (not gradual) means a catalog attribute change, not a sales issue
- **Blips** (1–2 day rank drops) are normal and don't require action
- Amazon's Choice badge is mapped to specific **price points** — track and defend it

## Step 1 — Confirm Inputs

Before running, collect:
- **Bulk Operations file** — Campaign-level and keyword-level data
- **Search Query Performance (SQP) data** *(optional but recommended)* — From Brand Analytics
- **Target keywords for ranking** — Which keywords are you trying to rank for?
- **Current organic positions** *(optional)* — From Helium 10, Data Dive, or ASINSIGHT
- **Product ASIN(s)** — The ASIN(s) being ranked
- **Target ACoS for ranking campaigns**: Default 60% (willing to invest for rank)
- **Brand name(s)** — To separate brand vs. non-brand ranking

## Step 2 — Run the Rank Tracker Script

```bash
python3 /path/to/skills/amazon-ppc-rank-tracker/scripts/rank_tracker.py \
  --bulk "/path/to/bulk_file.xlsx" \
  --output "/path/to/output/Rank_Tracker_[date].xlsx" \
  --target-acos 0.60 \
  --brand "RENUV" \
  --asin "B0XXXXXXXXX"
```

With optional SQP data:
```bash
python3 /path/to/skills/amazon-ppc-rank-tracker/scripts/rank_tracker.py \
  --bulk "/path/to/bulk_file.xlsx" \
  --sqp "/path/to/search_query_performance.xlsx" \
  --output "/path/to/output/Rank_Tracker_[date].xlsx" \
  --target-acos 0.60 \
  --brand "RENUV" \
  --asin "B0XXXXXXXXX"
```

## Step 3 — Interpret and Present Results

After running, tell the user:
1. How many ranking campaigns (SKCs) exist and their overall health
2. Which keywords have strong ranking signals (high CVR, good CTR, consistent sales)
3. Which keywords are **not responding** to PPC spend (low CVR, no traction)
4. Top of Search impression share for key ranking keywords
5. Budget adequacy — are ranking campaigns being throttled?
6. Day parting recommendations for ranking campaigns
7. Keywords ready for the "Bait & Switch" (reduce PPC, test organic hold)

## The AdsCrafted 6-Step Ranking Process

### Step 1: Identify Target Keywords
- Use reverse ASIN research (Helium 10, ASINSIGHT, Data Dive)
- Group keywords into **keyword groups** (related terms that rank together)
- Prioritize by search volume, relevancy, and competitive difficulty

### Step 2: Validate Ranking Potential
- Check Search Query Performance (SQP) for conversion rate on each keyword
- Minimum CVR threshold: 10% for ranking viability
- If CVR < 10%, listing optimization is needed before ranking push

### Step 3: Apply Ranking Criteria
Ranking is working when:
- Ad impressions are flowing at Top of Search
- CTR > 0.3% (listing is compelling at this position)
- CVR ≥ category average (offer is competitive)
- Sales velocity is sustained (not just spikes)

### Step 4: Ensure Relevancy
- Listing must contain the target keyword in title, bullets, or backend
- Main image must be compelling for the search intent
- A+ content should address the keyword's purchase intent

### Step 5: Increase Visibility (SKC Optimization)
- Use **Fixed Bids** with high Top of Search placement modifiers (100–300%)
- Target Top of Search specifically — this is where ranking signals are strongest
- Apply **day parting** to concentrate spend during peak conversion hours (6AM–10PM)
- Ensure budget is adequate — a throttled ranking campaign is a wasted ranking campaign

### Step 6: Maximize the Offer
During ranking push, stack conversion boosters:
- **Price reduction** or **coupon** (10–20% off)
- **Subscribe & Save** enrollment
- **Lightning Deal** or **7-Day Deal** timing
- **Optimized main image** and **A+ content**
- **Day parting** to spend only during high-conversion windows

## Ranking Readiness Analysis Logic

For each keyword in a ranking campaign, calculate:

### Ranking Score (0–100)

| Factor | Weight | Scoring |
|---|---|---|
| Conversion Rate | 30% | ≥15% = 30pts, 10–15% = 20pts, 5–10% = 10pts, <5% = 0pts |
| CTR | 20% | ≥0.5% = 20pts, 0.3–0.5% = 15pts, 0.1–0.3% = 10pts, <0.1% = 0pts |
| Top of Search IS | 20% | ≥10% = 20pts, 5–10% = 15pts, 1–5% = 10pts, <1% = 0pts |
| Sales Consistency | 15% | Orders in ≥80% of days = 15pts, 50–80% = 10pts, <50% = 5pts |
| Budget Utilization | 15% | 60–90% = 15pts, 90–100% (throttled) = 10pts, <60% = 5pts |

### Ranking Status Categories

| Status | Score | Meaning |
|---|---|---|
| RANKING READY | 70–100 | Strong signals — maintain and monitor organic position |
| PROGRESSING | 50–69 | Getting traction — continue pushing, optimize offer |
| STRUGGLING | 30–49 | Weak signals — review listing, pricing, or keyword relevancy |
| NOT RANKING | 0–29 | No traction — consider different approach or keyword |

### Bait & Switch Readiness

A keyword is ready for the Bait & Switch test when:
- Ranking Score ≥ 70
- Sustained sales for ≥ 14 consecutive days
- Organic impressions are growing (if SQP data available)
- Top of Search IS > 5%

**Test procedure**: Reduce daily budget by 50% for 7 days. If organic sales hold,
reduce further. If organic sales drop, restore budget immediately.

## Day Parting Recommendations for Ranking

Ranking campaigns should concentrate spend during peak conversion hours:

| Priority | Hours | Reasoning |
|---|---|---|
| HIGH | 6AM – 12PM | Morning shoppers have highest purchase intent |
| HIGH | 6PM – 10PM | Evening shoppers complete purchases |
| MEDIUM | 12PM – 6PM | Afternoon has moderate conversion |
| LOW | 10PM – 6AM | Overnight has lowest conversion — waste of ranking budget |

## Output File Structure

1. **📊 Ranking Dashboard** — Overall ranking campaign health, keyword count by status
2. **🎯 Keyword Ranking Scores** — Every ranking keyword with its score and status breakdown
3. **🔄 Bait & Switch Candidates** — Keywords ready to test organic rank hold
4. **⚠ Struggling Keywords** — Keywords not responding to PPC — need intervention
5. **🕐 Day Parting Recommendations** — Time-based optimization opportunities per campaign
6. **💰 Budget Adequacy** — Are ranking campaigns getting enough budget to succeed?
7. **📈 Keyword Groups** — Related keywords that rank together (clustering analysis)
8. **📋 Ranking Action Plan** — Prioritized weekly actions for ranking campaigns
9. **🗂 Raw Data** — All ranking campaign data with all calculations

## Important Notes

- This tool focuses on **ranking campaigns only** — it filters for SKCs and campaigns
  with "rank" or "launch" in their names, plus any Exact match campaigns with high
  placement modifiers.
- Ranking requires **patience** — the 2–4 day delay means you won't see results immediately.
  Flag any keyword with < 14 days of data as "insufficient data for ranking assessment."
- The Bait & Switch test is **reversible** — always be ready to restore budget if organic
  rank doesn't hold.
- **Amazon's Choice badge** tracking: if the user provides price data, the tool checks
  whether current pricing is within the Amazon's Choice threshold range.
- Recommend the user also run the **Placement Optimizer** on ranking campaigns to ensure
  Top of Search modifiers are set correctly.
