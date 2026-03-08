---
name: amazon-ppc-run-all
description: >
  Amazon PPC Master Orchestrator — runs the FULL weekly PPC management workflow
  for one or more brands in a single session. Use this skill whenever the user
  says things like "run my PPC", "weekly PPC update", "process my Amazon reports",
  "run all the PPC tools", "analyze my PPC for [brand]", or uploads a bulk file
  and/or search term report and wants a complete analysis. This is the PRIMARY
  entry point for the entire Amazon PPC system — always prefer this over running
  individual skills when the user wants a full analysis.
---

# Amazon PPC Master Orchestrator

You run the complete weekly PPC management workflow by coordinating all eight
skills in the correct sequence. Follow these steps exactly.

## AdsCrafted Methodology Summary

This orchestrator follows the AdsCrafted PPC Mastery framework:
1. **Classify first** — understand each campaign's strategic goal before optimizing
2. **Harvest** — isolate converting search terms into Exact match / SKC campaigns
3. **Optimize** — adjust bids using lifecycle-aware logic (inch-up → revenue-based)
4. **Manage budgets** — free constrained winners, cut inefficient spenders
5. **Fix placements** — ensure campaigns serve at their intended placement
6. **Build SKCs** — create Single Keyword Campaigns from harvested winners
7. **Track ranking** — monitor organic rank progress on SKC campaigns
8. **Report** — compile everything into the weekly performance report

---

## Step 1 — Intake (always do this first)

Ask the user for everything you need BEFORE running any tools. Collect:

### Required files
- **Bulk Operations file** — Seller Central → Campaign Manager → Bulk Operations → Download → Sponsored Products → All records
- **Search Term Report** — Seller Central → Reports → Advertising → Search Term → Download

### Required settings (ask if not already known)
- **Brand name** — used in report headers (e.g. "RENUV")
- **Target ACoS (Profit)** — overall default for profitable campaigns (e.g. 25%)
- **Target ACoS (Ranking)** — for ranking/launch campaigns (e.g. 60%)
- **Date range label** — the period the bulk file covers (e.g. "Jan 25 – Feb 24, 2026")
- **Brand keywords** — your brand name(s) for brand vs. non-brand split

### Optional
- Per-portfolio ACoS targets (e.g. Dishwasher Cleaner = 20%, Citric Acid = 30%)
- Search Query Performance (SQP) data — for ranking analysis
- Product ASIN(s) — for SKC builder and rank tracker

If the user has already provided files and settings in the conversation, extract
them directly — do not ask again.

**Ask all questions in a single message. Do not ask one at a time.**

---

## Step 2 — Confirm & Run

Once you have the files and settings, confirm with one sentence:
> "Running full PPC analysis for [Brand] — profit ACoS [X%], ranking ACoS [Y%], [date range]. This will produce 8 output files."

Then run all scripts in this order:

### 1. Campaign Strategist (NEW — Run First)
Classifies campaigns by goal, audits wasted spend, day parting analysis.
```bash
python3 [skills_path]/amazon-ppc-campaign-strategist/scripts/campaign_strategist.py \
  --input "[bulk_file]" \
  --output "[output_dir]/1_Campaign_Strategy_[date].xlsx" \
  --target-acos-profit [target_profit] \
  --target-acos-ranking [target_ranking] \
  --target-acos-research 0.35 \
  --target-acos-reviews 0.40 \
  --target-acos-marketshare 0.30 \
  --brand-keywords "[brand]" \
  --days 30
```

### 2. Search Term Harvester
```bash
python3 [skills_path]/amazon-ppc-harvester/scripts/harvester.py \
  --input "[search_term_report]" \
  --bulk-file "[bulk_file]" \
  --output "[output_dir]/2_Search_Term_Harvest_[date].xlsx" \
  --target-acos [target_profit] \
  --min-clicks 3 --min-orders 1 --neg-spend-threshold 10
```

### 3. Bid Optimizer (Updated with Bid Lifecycle)
```bash
python3 [skills_path]/amazon-ppc-bid-optimizer/scripts/bid_optimizer.py \
  --input "[bulk_file]" \
  --output "[output_dir]/3_Bid_Optimization_[date].xlsx" \
  --target-acos [target_profit] \
  --min-clicks 10 --max-raise 0.30 --max-lower 0.40 \
  --bid-floor 0.20 --bid-ceiling 5.00 \
  --pause-spend 15 --pause-clicks 15
```

### 4. Budget Manager
```bash
python3 [skills_path]/amazon-ppc-budget-manager/scripts/budget_manager.py \
  --input "[bulk_file]" \
  --output "[output_dir]/4_Budget_Manager_[date].xlsx" \
  --target-acos [target_profit] --days 30 \
  --constrained-threshold 0.80 --underutilized-threshold 0.20 \
  --increase-pct 0.25 --decrease-pct 0.20 --min-budget 10
```

### 5. Placement Optimizer
```bash
python3 [skills_path]/amazon-ppc-placement-optimizer/scripts/placement_optimizer.py \
  --input "[bulk_file]" \
  --output "[output_dir]/5_Placement_Optimizer_[date].xlsx" \
  --target-acos [target_profit] --min-clicks 10 \
  --max-increase 50 --max-decrease 100 --leakage-spend 20
```

### 6. SKC Builder (NEW — Build from Harvest Winners)
Only run if the Harvester found Gold → Exact candidates.
```bash
python3 [skills_path]/amazon-ppc-skc-builder/scripts/skc_builder.py \
  --input "[output_dir]/2_Search_Term_Harvest_[date].xlsx" \
  --output "[output_dir]/6_SKC_Campaigns_[date].xlsx" \
  --asin "[primary_asin]" \
  --portfolio "[portfolio_name]" \
  --goal "profit" \
  --target-acos [target_profit] \
  --starting-bid 0.50 \
  --bid-strategy "inch-up" \
  --placement "tos" \
  --daily-budget 15.00
```

### 7. Rank Tracker (NEW — Monitor Ranking Campaigns)
Only run if ranking/SKC campaigns exist in the bulk file.
```bash
python3 [skills_path]/amazon-ppc-rank-tracker/scripts/rank_tracker.py \
  --bulk "[bulk_file]" \
  --output "[output_dir]/7_Rank_Tracker_[date].xlsx" \
  --target-acos [target_ranking] \
  --brand "[brand]" \
  --asin "[primary_asin]"
```

### 8. Weekly Report
```bash
python3 [skills_path]/amazon-ppc-weekly-report/scripts/weekly_report.py \
  --bulk "[bulk_file]" \
  --search-terms "[search_term_report]" \
  --output "[output_dir]/8_Weekly_Report_[date].xlsx" \
  --brand "[brand]" \
  --target-acos [target_profit] \
  --date-range "[date_range]"
```

**[skills_path]** = `/path/to/Desktop/skills`
**[output_dir]** = user's Desktop or a brand subfolder

---

## Step 3 — Present Results

After all 8 scripts complete, present all output files to the user and give a
comprehensive summary organized by the AdsCrafted workflow:

### Strategy Overview (from Campaign Strategist)
- Campaign goal distribution (how many Ranking, Profit, Research, Brand, Dead)
- Misaligned campaigns requiring goal reassignment
- Total wasted spend identified (duplicates + zero-ROI + product targeting waste)

### Harvest Summary (from Harvester)
- X new exact keywords to add (Gold), Y phrase keywords (Silver)
- Z terms to negate, recovering $X in wasted spend
- Duplicate targeting found across N campaigns

### Bid Optimization (from Bid Optimizer)
- Keywords to raise (scaling winners) vs. lower (trimming losers)
- Inch-up candidates (new keywords in launch stage)
- Hold keywords (discovery stage, accumulating data)

### Budget & Placement (from Budget Manager + Placement Optimizer)
- Constrained efficient campaigns needing more budget
- Budget reallocation opportunity
- Placement leakage suppressed

### SKC & Ranking (from SKC Builder + Rank Tracker)
- New SKCs created from harvest winners
- Ranking keywords by status (Ready, Progressing, Struggling)
- Bait & Switch candidates (ready to test organic hold)

### Weekly Report
- Overall account KPIs vs. last period
- Top/bottom performing campaigns
- Prioritized action checklist

---

## Step 4 — Multi-Brand Support

If the user wants to run multiple brands, repeat Steps 1–3 for each brand.
Save each brand's outputs into a separate dated subfolder:
```
Desktop/
└── PPC_Reports_2026-02-24/
    ├── RENUV/
    │   ├── 1_Campaign_Strategy_...xlsx
    │   ├── 2_Search_Term_Harvest_...xlsx
    │   ├── 3_Bid_Optimization_...xlsx
    │   ├── 4_Budget_Manager_...xlsx
    │   ├── 5_Placement_Optimizer_...xlsx
    │   ├── 6_SKC_Campaigns_...xlsx
    │   ├── 7_Rank_Tracker_...xlsx
    │   └── 8_Weekly_Report_...xlsx
    └── BrandB/
        └── ...
```

---

## Path Reference

All skill scripts live at:
```
[Desktop]/skills/
├── amazon-ppc-campaign-strategist/scripts/campaign_strategist.py   (NEW)
├── amazon-ppc-harvester/scripts/harvester.py
├── amazon-ppc-bid-optimizer/scripts/bid_optimizer.py               (UPDATED)
├── amazon-ppc-budget-manager/scripts/budget_manager.py
├── amazon-ppc-placement-optimizer/scripts/placement_optimizer.py
├── amazon-ppc-skc-builder/scripts/skc_builder.py                   (NEW)
├── amazon-ppc-rank-tracker/scripts/rank_tracker.py                 (NEW)
└── amazon-ppc-weekly-report/scripts/weekly_report.py
```
