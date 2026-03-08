---
name: amazon-ppc-budget-manager
description: >
  Amazon PPC Budget Manager for Sponsored Products. Use this skill whenever
  the user uploads a Bulk Operations file and wants to: review campaign budget
  utilization, find campaigns being throttled by budget caps, identify wasteful
  overspending campaigns, reallocate budget across portfolios, or get a
  bulk-upload-ready budget change file. Trigger on phrases like "manage my
  budgets", "which campaigns are budget-limited", "budget optimization",
  "find campaigns hitting budget cap", "reallocate budget", or any time a bulk
  file is uploaded and budgets are mentioned.
---

# Amazon PPC Budget Manager

You are an expert Amazon PPC manager. When a user uploads a Sponsored Products
Bulk Operations file, analyze campaign-level budget utilization and produce a
bulk-upload-ready Excel file with budget change recommendations.

## Step 1 — Confirm Inputs

Before running, confirm (or use these defaults):
- **Target ACoS**: Default 25%
- **Lookback days**: Default 30 (matches the bulk file date range)
- **Constrained threshold**: Default 80% — campaigns spending ≥80% of daily budget
- **Underutilized threshold**: Default 20% — campaigns spending ≤20% of daily budget
- **Budget increase %**: Default 25% for constrained efficient campaigns
- **Budget decrease %**: Default 20% for constrained inefficient campaigns
- **Min daily budget**: Default $10

If the user already stated these, don't ask again.

## Step 2 — Run the Script

```bash
python3 /path/to/skills/amazon-ppc-budget-manager/scripts/budget_manager.py \
  --input "/path/to/bulk_file.xlsx" \
  --output "/path/to/output/Budget_Manager_[date].xlsx" \
  --target-acos 0.25 \
  --days 30 \
  --constrained-threshold 0.80 \
  --underutilized-threshold 0.20 \
  --increase-pct 0.25 \
  --decrease-pct 0.20 \
  --min-budget 10
```

## Step 3 — Interpret and Present Results

After running, tell the user:
1. How many campaigns are budget-constrained (hitting their cap)
2. Which constrained campaigns have good ACoS (increase budget → more profit)
3. Which constrained campaigns have bad ACoS (decrease budget → stop bleeding)
4. How much budget is sitting unused across underutilized campaigns
5. Total budget reallocation opportunity ($ that could shift from bad to good)

## Budget Logic Reference

| Condition | Action | Reason |
|---|---|---|
| Util ≥ constrained_threshold AND ACoS ≤ target×1.5 | INCREASE | Profitable campaign being throttled — give it more room |
| Util ≥ constrained_threshold AND ACoS > target×1.5 | DECREASE | Over-budget AND inefficient — cut losses |
| Util ≤ underutilized_threshold AND spend > 0 | INVESTIGATE | Budget set too high or campaign underperforming |
| Spend = 0, State = enabled | ALERT | Campaign not serving at all — check bids/targeting |
| Otherwise | NO CHANGE | Within normal operating range |

### Budget Increase Formula
```
new_budget = current_budget × (1 + increase_pct)
```
Only applies when campaign is both constrained AND efficient.

### Budget Decrease Formula
```
new_budget = max(current_budget × (1 - decrease_pct), min_budget)
```
Applies when constrained AND inefficient (ACoS > target × 1.5).

### Underutilized Right-Size Formula
```
suggested_budget = max(daily_spend × 1.5, min_budget)
```
Suggests a tighter budget to free up unused allocation.

## Output File Structure

1. **📊 Summary** — Portfolio budget health dashboard
2. **📈 Increase Budget** — Constrained + efficient campaigns
3. **📉 Decrease Budget** — Constrained + inefficient campaigns
4. **🔍 Investigate** — Severely underutilized campaigns
5. **🚨 Alerts** — Zero-spend enabled campaigns
6. **♻ Reallocation Map** — Budget to move from bad to good campaigns
7. **📤 Amazon Bulk Upload** — Ready to upload (Campaign budget updates)
8. **🗂 Raw Data** — All campaigns with all calculations

## Notes

- The bulk upload tab uses Campaign ID for precise updates.
- Only INCREASE and DECREASE rows appear in the bulk upload tab.
  INVESTIGATE and ALERT rows require manual review first.
- Budget changes take effect immediately on upload — review carefully.
