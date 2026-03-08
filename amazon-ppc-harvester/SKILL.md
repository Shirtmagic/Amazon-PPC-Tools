---
name: amazon-ppc-harvester
description: >
  Amazon PPC Search Term Harvester for Sponsored Products. Use this skill whenever
  the user uploads a Search Term Report (xlsx/csv) and wants to: find converting
  search terms to add as new keywords, identify wasted spend terms to negate,
  get a bulk-upload-ready action file, or run a weekly PPC harvest analysis.
  Trigger on phrases like "harvest search terms", "run PPC harvest", "find new
  keywords from search terms", "what should I negate", "process my search term
  report", or any time a file named with "search_term" or "search term" is uploaded.
---

# Amazon PPC Search Term Harvester

You are an expert Amazon PPC manager. When a user uploads a Sponsored Products
Search Term Report, your job is to analyze it and produce a clear, actionable
Excel output with harvesting recommendations and a bulk-upload-ready tab.

## Step 1 — Confirm Inputs

Before running, confirm (or use these defaults):
- **Target ACoS**: Ask the user if unknown. Default: 25%
- **Min clicks to harvest**: Default: 3
- **Min orders to harvest**: Default: 1
- **Negative spend threshold**: Default: $10 (spend with 0 orders)
- **Lookback window**: Use all data in the file unless user specifies

If the user has already stated these in the conversation, use them — do not ask again.

## Step 2 — Run the Harvester Script

Run the Python script at `scripts/harvester.py` with the uploaded file:

```bash
python3 /path/to/skills/amazon-ppc-harvester/scripts/harvester.py \
  --input "/path/to/search_term_report.xlsx" \
  --output "/path/to/output/PPC_Harvest_[date].xlsx" \
  --target-acos 0.25 \
  --min-clicks 3 \
  --min-orders 1 \
  --neg-spend-threshold 10
```

Replace paths with actual file locations. The output path should go to the
user's Desktop or outputs folder.

## Step 3 — Interpret and Present Results

After the script runs, read the Summary tab and tell the user:
1. How many harvest candidates were found (Gold → Exact, Silver → Phrase)
2. How much wasted spend the negatives would recover
3. Top 5 best converting search terms found
4. Top 5 biggest spend-with-no-orders terms to negate
5. Any portfolios or campaigns with unusually high ACoS

Then say: "The full action file is ready. Review the tabs, then upload the
'Amazon Bulk Upload' tab as a bulk operations file in Seller Central."

## Step 4 — Answer Follow-Up Questions

The user may ask things like:
- "Why is [term] flagged as a negative?" → Explain the rule triggered (spend, clicks, ACoS)
- "Change the ACoS target to 20%" → Re-run the script with updated params
- "Show me only the Dishwasher Cleaner portfolio" → Filter and re-summarize

## Harvesting Logic Reference (AdsCrafted Search Term Isolation)

The AdsCrafted method emphasizes **Search Term Isolation** — every converting search
term should end up in its own Exact match campaign (SKC) for maximum control. The
harvest funnel flows: Auto → Broad → Phrase → Exact (SKC).

| Category | Rule | Action |
|---|---|---|
| Gold → Exact | Orders ≥ 1, ACoS ≤ target, Clicks ≥ 3, source is Auto/Broad/Phrase | Add as EXACT keyword to manual campaign (ideally SKC) |
| Silver → Phrase | Orders ≥ 1, ACoS ≤ target×1.5, Clicks ≥ 2, source is Auto (-) only | Add as PHRASE keyword to manual campaign |
| Negative Exact | Spend ≥ threshold, Orders = 0, Clicks ≥ 5 | Add as NEGATIVE EXACT to source campaign |
| Negative Phrase | Spend ≥ threshold×0.5, Orders = 0, Clicks ≥ 3, long-tail term (3+ words) | Add as NEGATIVE PHRASE to source campaign |
| Review | Orders ≥ 1, ACoS > target×2, Spend ≥ $5 | Flag for manual review |
| Duplicate | Same search term converting in 2+ campaigns | Flag — consolidate into best-performing campaign |

Auto campaigns have Match Type = "-" in Amazon's report.

## Wasted Spend Detection (AdsCrafted Method)

Beyond simple negation, the harvester identifies three types of waste:

### 1. Zero-Order Spend
Keywords with significant spend but zero orders. This is the most straightforward
waste — negate or pause these terms.

### 2. Duplicate Targeting (Cannibalization)
The same search term appearing in multiple campaigns means you're bidding against
yourself in the auction. The harvester detects when a search term converts in
Campaign A but is also being bid on (with or without conversions) in Campaign B.
- **Recommendation**: Keep the term in the best-performing campaign, negate everywhere else
- **Cannibalization cost**: Sum of spend on the duplicated term across non-primary campaigns

### 3. Product Targeting Waste (AdsCrafted Insight)
The AdsCrafted method recommends **shifting away from product targeting** over time.
Product targeting (ASIN targeting) typically has lower CVR and higher ACoS than
keyword targeting. The harvester flags:
- Product targeting search terms with ACoS > target × 2
- Total spend on product targeting vs. keyword targeting
- Recommendation to phase out high-ACoS product targeting

## Output File Structure

The script produces an Excel with these sheets:
1. **Summary** — KPI dashboard with portfolio breakdown
2. **Harvest → Exact** — Gold terms ready to add as exact keywords
3. **Harvest → Phrase** — Silver terms ready to add as phrase keywords
4. **Negatives** — Terms to add as negative exact/phrase
5. **Review** — High-ACoS converters needing manual attention
6. **Amazon Bulk Upload** — Pre-formatted for direct upload to Seller Central bulk ops
7. **Raw Data** — Original aggregated data with all flags appended

## Important Notes

- The bulk upload tab produces KEYWORD and NEGATIVE KEYWORD rows only.
  The user must fill in Campaign ID and Ad Group ID columns before uploading,
  as these require their specific account IDs.
- Suggest bid for new exact keywords = (Avg CPC from data) × 1.1, capped at $3.00
- All monetary values in USD
- If data covers fewer than 7 days, add a warning in the summary
