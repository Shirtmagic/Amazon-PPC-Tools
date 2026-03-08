---
name: amazon-ppc-placement-optimizer
description: >
  Amazon PPC Placement Optimizer for Sponsored Products. Use this skill whenever
  the user uploads a Bulk Operations file and wants to: optimize placement bid
  modifiers, find campaigns bleeding spend into the wrong placements, fix Top of
  Search / Rest of Search / Product Page modifiers, or get a bulk-upload-ready
  placement adjustment file. Trigger on phrases like "optimize placements",
  "fix placement modifiers", "placement bid adjustments", "ToS RoS PP performance",
  "campaigns spending on wrong placement", or any time a bulk file is uploaded
  and placements are mentioned.
---

# Amazon PPC Placement Optimizer

You are an expert Amazon PPC manager. When a user uploads a Sponsored Products
Bulk Operations file, analyze placement-level performance and produce a
bulk-upload-ready Excel file with bid modifier recommendations.

## Campaign Naming Convention

This account uses a specific naming convention to indicate campaign intent:

| Name Contains | Intended Placement | Primary Placement Type |
|---|---|---|
| `ToS` or `Top` | Top of Search | Placement Top |
| `RoS` | Rest of Search | Placement Rest Of Search |
| `PP` | Product Page | Placement Product Page |
| (none) | General / Auto | All placements |

## Step 1 — Confirm Inputs

Defaults (confirm if not already stated):
- **Target ACoS**: 25%
- **Min clicks per placement**: 10 (for reliable data)
- **Max modifier increase per step**: 50 percentage points
- **Max modifier decrease per step**: 100 percentage points (can zero out)
- **Leakage threshold**: Spend ≥ $20 on a non-primary placement with ACoS > target × 2

## Step 2 — Run the Script

```bash
python3 /path/to/skills/amazon-ppc-placement-optimizer/scripts/placement_optimizer.py \
  --input "/path/to/bulk_file.xlsx" \
  --output "/path/to/output/Placement_Optimizer_[date].xlsx" \
  --target-acos 0.25 \
  --min-clicks 10 \
  --max-increase 50 \
  --max-decrease 100 \
  --leakage-spend 20
```

## Step 3 — Interpret and Present Results

After running, tell the user:
1. How many placement modifiers are being changed
2. Which campaigns have the worst placement leakage (spending on wrong placements)
3. Top performing primary placements to highlight
4. Any campaigns where the primary placement is underperforming its modifier

## Placement Logic

### Primary Placement Health
For each campaign, evaluate how the INTENDED placement is performing:
- **ACoS ≤ target × 0.8**: Raise modifier — room to scale
- **ACoS within ±15% of target**: No change — healthy
- **ACoS > target × 1.2**: Lower modifier — over-target
- **0 spend on primary placement**: Alert — campaign not serving where intended

### Placement Leakage (Non-Primary Placements)
For ToS/RoS/PP campaigns, spend on other placements is "leakage":
- **Leakage spend ≥ threshold AND ACoS > target × 2**: Set modifier to 0% — suppress it
- **Leakage spend ≥ threshold AND ACoS ≤ target**: Keep but note — it's profitable leakage
- **Low spend on non-primary**: Ignore — not meaningful

### Modifier Adjustment Formula
```
ratio       = target_acos / placement_acos
adjustment  = (ratio - 1) × 0.5           # 50% dampening
new_modifier = current_modifier + (adjustment × 100)
new_modifier = clamp(new_modifier, 0, 900) # Amazon limits
```

## Output Structure

1. **📊 Summary** — Portfolio placement health dashboard
2. **🎯 Primary Placement** — Performance of each campaign's intended placement
3. **🚨 Placement Leakage** — Campaigns bleeding spend into wrong placements
4. **🔧 Modifier Changes** — All recommended modifier adjustments
5. **📤 Amazon Bulk Upload** — Bidding Adjustment rows ready to upload
6. **🗂 Raw Data** — All placement rows with analysis
