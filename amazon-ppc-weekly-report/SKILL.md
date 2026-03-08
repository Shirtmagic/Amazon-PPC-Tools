---
name: amazon-ppc-weekly-report
description: >
  Generates a full Amazon PPC Weekly Performance Report from a Bulk Operations
  file + Search Term Report. Creates a branded Excel dashboard with Executive
  Summary, Portfolio Performance, Keyword Analysis, Search Term Insights, and
  Action Checklist sheets. Use when the user says "weekly report", "PPC report",
  "performance report", "run the report", or uploads PPC data files.
---

# Amazon PPC Weekly Report Skill

## What It Does
Combines your Amazon Bulk Operations file and Search Term Report into a
single polished Excel dashboard. Five sheets are produced automatically:

| Sheet | Contents |
|---|---|
| 📋 Executive Summary | 8 KPI cards, portfolio performance table, top/bottom 5 campaigns |
| 📂 Portfolio Detail | Every campaign with budget utilisation, ACoS vs target, status |
| 🔍 Keyword Analysis | All active keywords colour-coded vs target ACoS |
| 🔎 Search Term Insights | Top converters + wasted spend (zero-order terms) |
| ✅ Action Checklist | Auto-generated prioritised weekly actions |

---

## Step 1 — Intake (ask in ONE message)

> "To generate your weekly PPC report I need a few things:
>
> 1. **Bulk Operations file** — the `.xlsx` from Amazon Seller Central
>    (Advertising → Bulk Operations → Download). Use the file covering your
>    desired date window (typically last 30 days).
> 2. **Search Term Report** — the `.xlsx` from Reports → Advertising Reports
>    → Sponsored Products → Search Term (date range matching the bulk file).
> 3. **Brand name** — e.g. `RENUV` (used in the report header)
> 4. **Target ACoS** — e.g. `25%` (default 25%)
> 5. **Date range label** — e.g. `Jan 25 – Feb 24, 2026` (appears in header)
> 6. **Per-portfolio targets** _(optional)_ — if certain portfolios have
>    different ACoS targets, list them (e.g. `Coffee Machine Cleaner: 20%`)
>
> Please upload the two files and answer the questions above, and I'll
> generate your report immediately."

---

## Step 2 — Run the Script

Once files are uploaded, run:

```bash
python3 /sessions/optimistic-clever-carson/mnt/Desktop/skills/amazon-ppc-weekly-report/scripts/weekly_report.py \
  --bulk      "<path to bulk file>" \
  --search-terms "<path to search term report>" \
  --output    "/sessions/optimistic-clever-carson/mnt/Desktop/Weekly_Report_<BRAND>_<DATE>.xlsx" \
  --brand     "<BRAND NAME>" \
  --target-acos <0.20> \
  --date-range  "<Date Range>"
```

**Path template** (Desktop output folder):
```
/sessions/optimistic-clever-carson/mnt/Desktop/Weekly_Report_RENUV_2026-02-24.xlsx
```

---

## Step 3 — Present Results

After the script runs, share the file link with a brief 5-bullet summary:

- **Total Spend / Sales / ACoS** (overall account)
- **Portfolio count and status** (how many ✅ / ⚠ / 🔴)
- **Top performer** (lowest ACoS campaign with real spend)
- **Biggest concern** (highest ACoS campaign with real spend)
- **Action count** (number of items on the Action Checklist)

---

## Multi-Brand Support

For multiple brands, run the script once per brand, saving outputs to dated
subfolders:

```
Desktop/PPC_Reports_2026-02-24/RENUV/Weekly_Report_RENUV_2026-02-24.xlsx
Desktop/PPC_Reports_2026-02-24/OtherBrand/Weekly_Report_OtherBrand_2026-02-24.xlsx
```

---

## Reference — RENUV Portfolios & Targets

| Portfolio | Target ACoS |
|---|---|
| Coffee Machine Cleaner | 25% |
| Dishwasher Cleaner | 25% |
| Citric Acid Powder | 25% |
| Washing Machine Cleaner | 25% |
| Garbage Disposal Cleaner | 25% |
| Multi ASIN (Combined PPC) | 25% |
| Laundry Sheets | 25% |

---

## Files

| File | Purpose |
|---|---|
| `scripts/weekly_report.py` | Main report generator |
| `SKILL.md` | This file — skill instructions |
