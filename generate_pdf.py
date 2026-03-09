#!/usr/bin/env python3
"""Generate a comprehensive PDF summary of the Amazon PPC Optimization System."""

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether
)

OUTPUT = "/Users/toddkriney/Desktop/Amazon_PPC_System_Summary.pdf"

ORANGE = HexColor("#FF9900")
DARK = HexColor("#232F3E")
BLUE = HexColor("#146EB4")
GREEN = HexColor("#067D62")
RED = HexColor("#D13212")
GRAY = HexColor("#666666")
LIGHT = HexColor("#F5F5F5")
WHITE = HexColor("#FFFFFF")


def build_pdf():
    doc = SimpleDocTemplate(
        OUTPUT, pagesize=letter,
        topMargin=0.5 * inch, bottomMargin=0.5 * inch,
        leftMargin=0.65 * inch, rightMargin=0.65 * inch,
    )
    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle("CTitle", fontSize=26, textColor=DARK, alignment=TA_CENTER, spaceAfter=4, fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle("CSub", fontSize=12, textColor=GRAY, alignment=TA_CENTER, spaceAfter=2))
    styles.add(ParagraphStyle("Sec", fontSize=16, textColor=DARK, spaceBefore=16, spaceAfter=6, fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle("Sub2", fontSize=13, textColor=BLUE, spaceBefore=12, spaceAfter=4, fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle("Sub3", fontSize=11, textColor=DARK, spaceBefore=10, spaceAfter=6, fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle("Body", fontSize=10, leading=14, spaceAfter=6, alignment=TA_JUSTIFY))
    styles.add(ParagraphStyle("BL", fontSize=10, leading=14, leftIndent=18, bulletIndent=6, spaceAfter=3))
    styles.add(ParagraphStyle("BL2", fontSize=10, leading=14, leftIndent=36, bulletIndent=24, spaceAfter=2))
    styles.add(ParagraphStyle("Foot", fontSize=7, textColor=GRAY, alignment=TA_CENTER))
    styles.add(ParagraphStyle("CallOut", fontSize=10, leading=14, spaceBefore=4, spaceAfter=8, backColor=LIGHT, borderColor=ORANGE, borderWidth=1, borderPadding=8, leftIndent=12, rightIndent=12))

    W = 7.0 * inch  # usable width

    def sec(title):
        return [
            HRFlowable(width="100%", thickness=2, color=ORANGE, spaceBefore=10, spaceAfter=0),
            Paragraph(title, styles["Sec"]),
        ]

    # Paragraph styles for table cells
    styles.add(ParagraphStyle("TH", fontSize=8.5, leading=12, textColor=WHITE, fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle("TC", fontSize=8.5, leading=12, textColor=DARK))

    def tbl(data, widths, extra=None):
        # Wrap all cell text in Paragraph objects so word-wrapping works
        wrapped = []
        for r, row in enumerate(data):
            new_row = []
            for cell in row:
                if isinstance(cell, str):
                    cell_text = cell.replace("\n", "<br/>")
                    style = styles["TH"] if r == 0 else styles["TC"]
                    new_row.append(Paragraph(cell_text, style))
                else:
                    new_row.append(cell)
            wrapped.append(new_row)
        t = Table(wrapped, colWidths=widths, hAlign="LEFT")
        s = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), DARK),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.5, HexColor("#CCCCCC")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, LIGHT]),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ])
        if extra:
            for e in extra:
                s.add(*e)
        t.setStyle(s)
        return t

    def bullet(text):
        return Paragraph(f"<bullet>&bull;</bullet> {text}", styles["BL"])

    def bullet2(text):
        return Paragraph(f"<bullet>-</bullet> {text}", styles["BL2"])

    story = []

    # ══════════════════════════════════════════════════════════════════════════
    # COVER
    # ══════════════════════════════════════════════════════════════════════════
    story.append(Spacer(1, 2 * inch))
    story.append(Paragraph("Amazon PPC Optimization System", styles["CTitle"]))
    story.append(Spacer(1, 6))
    story.append(Paragraph("AdsCrafted PPC Mastery Methodology", styles["CSub"]))
    story.append(Paragraph("Complete Strategy, Decision Framework & Tool Reference", styles["CSub"]))
    story.append(Spacer(1, 0.4 * inch))
    story.append(HRFlowable(width="50%", thickness=3, color=ORANGE, spaceAfter=10, spaceBefore=10))
    story.append(Spacer(1, 0.3 * inch))
    story.append(Paragraph("8 Integrated Tools  |  Weekly Optimization Cadence", styles["CSub"]))
    story.append(Paragraph("Goal-Based Strategy  |  Search Term Isolation  |  Revenue-Based Bidding", styles["CSub"]))
    story.append(Paragraph("Organic Ranking Engine  |  Bait & Switch Framework", styles["CSub"]))
    story.append(Spacer(1, 1.5 * inch))
    story.append(Paragraph("Confidential  |  Internal Use Only", styles["Foot"]))
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # 1. THE BIG PICTURE
    # ══════════════════════════════════════════════════════════════════════════
    story.extend(sec("1. The Big Picture"))

    story.append(Paragraph(
        "Most sellers treat Amazon PPC as a cost center — something you spend money on to get sales. "
        "The AdsCrafted methodology flips this completely. <b>PPC is an organic ranking engine first, "
        "and a profitability engine second.</b>", styles["Body"]))

    story.append(Paragraph(
        "The core idea: Amazon's A9 algorithm ranks products based on sales velocity, conversion rate, "
        "and relevance. When you drive consistent PPC sales on a specific keyword, Amazon starts ranking "
        "you organically for that keyword. Once organic rank holds, you reduce PPC spend — keeping the "
        "rank but eliminating the ad cost. This is called the <b>Bait & Switch</b>.", styles["Body"]))

    story.append(Paragraph("<b>The system operates on five core principles:</b>", styles["Sub3"]))

    story.append(bullet("<b>Goal-Based Campaigns:</b> Every campaign exists for a specific reason — ranking, profit, "
        "research, reviews, market share, or brand defense. The goal determines how much ACoS you're willing to accept."))
    story.append(bullet("<b>Search Term Isolation:</b> Search terms start in broad Auto campaigns and get promoted through "
        "a funnel (Auto → Broad → Phrase → Exact) as they prove themselves. Each promotion gives you more control and prevents wasted spend."))
    story.append(bullet("<b>Bid Lifecycle Management:</b> New keywords start with small 'inch-up' bids that gradually increase "
        "until impressions flow. As data accumulates, bids shift to revenue-based calculations. Losers get cut."))
    story.append(bullet("<b>Single Keyword Campaigns (SKCs):</b> Your best-performing keywords graduate into their own campaign "
        "with one ad group, one keyword, one ASIN. This gives maximum bid and budget control per keyword."))
    story.append(bullet("<b>Weekly Cadence:</b> The system runs weekly. Download your data, run the 8 tools, review "
        "recommendations, apply changes. Consistency beats perfection."))

    story.append(Spacer(1, 6))
    story.append(Paragraph(
        "Think of it as a machine: research campaigns discover new keywords, proven winners get promoted "
        "into exact match SKCs, bids adjust automatically based on performance, and the rank tracker tells "
        "you when a keyword is ready to go organic. The 8 tools automate this entire process.", styles["CallOut"]))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # 2. THE 8 TOOLS
    # ══════════════════════════════════════════════════════════════════════════
    story.extend(sec("2. The 8 Tools — What Each One Does"))

    story.append(Paragraph(
        "The tools run in a specific order because each one feeds into the next. The Campaign Strategist "
        "classifies everything first, which informs how the Harvester and Bid Optimizer make decisions.", styles["Body"]))

    story.append(tbl([
        ["#", "Tool", "What It Does", "Key Outputs"],
        ["1", "Campaign\nStrategist", "Classifies every campaign by goal (Ranking, Profit, Research, etc.) and checks if its "
            "actual performance matches that goal. Finds wasted spend from duplicate targeting, zero-ROI keywords, "
            "and product targeting overspend. Recommends day parting.", "Strategy Dashboard\nGoal Classification\nMisaligned Campaigns\nWasted Spend Audit\nDay Parting Analysis"],
        ["2", "Search Term\nHarvester", "Reads your Search Term Report and decides what to do with every search term. Winning "
            "terms get promoted up the funnel (Auto→Broad→Phrase→Exact). Losing terms get negated. Catches "
            "duplicate targeting across campaigns.", "Gold Promotions (→Exact)\nSilver Promotions (→Phrase)\nNegative Keywords\nDuplicate Alerts\nWasted Spend Terms"],
        ["3", "Bid\nOptimizer", "Adjusts every keyword bid based on its lifecycle stage. New keywords get inch-up bids. "
            "Keywords with enough data get revenue-based bids. High performers scale up. Non-converters get paused.", "Bid Change Sheet\nInch-Up Schedule\nHold List\nPause Recommendations\nBulk Upload Ready"],
        ["4", "Budget\nManager", "Analyzes how much of each campaign's daily budget is actually being spent. Constrained "
            "campaigns (spending 80%+) that are efficient get more budget. Inefficient ones get cut. "
            "Underutilized campaigns get investigated.", "Budget Increases\nBudget Decreases\nUnderutilized Alerts\nZero-Spend Alerts"],
        ["5", "Placement\nOptimizer", "Looks at performance by placement — Top of Search, Rest of Search, Product Pages. "
            "Adjusts bid modifiers to push spend toward whichever placement converts best. Catches spend "
            "leaking to bad placements.", "Modifier Adjustments\nLeakage Detection\nPlacement Performance"],
        ["6", "SKC\nBuilder", "Takes harvested winners and generates full Single Keyword Campaign structures. Outputs in "
            "Amazon's Bulk Upload format — ready to paste directly into Campaign Manager. Sets bids, budgets, "
            "and placement modifiers based on the keyword's goal.", "Campaign Structure\nBulk Upload Sheet\nInch-Up Schedule\nBid Stacking Calc"],
        ["7", "Rank\nTracker", "Scores every keyword on a 0-100 Ranking Score based on CVR, CTR, Top of Search impression "
            "share, sales consistency, and budget utilization. Identifies Bait & Switch candidates — keywords "
            "ready to go organic.", "Ranking Scores\nBait & Switch Candidates\nStruggling Keywords\nDay Parting Recs"],
        ["8", "Weekly\nReport", "Executive summary rolling up all metrics: total spend, sales, ACoS, TACoS, top/bottom "
            "campaigns, week-over-week trends. Designed to be the review document for team meetings or "
            "client handoffs.", "Executive Summary\nPortfolio Detail\nKeyword Analysis\nAction Checklist"],
    ], [0.3 * inch, 0.75 * inch, 3.75 * inch, 2.2 * inch]))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # 3. CAMPAIGN GOALS
    # ══════════════════════════════════════════════════════════════════════════
    story.extend(sec("3. Campaign Goal Classification"))

    story.append(Paragraph(
        "This is the foundation of the entire system. Every campaign must have a clear goal, because the goal "
        "determines what ACoS is acceptable. A ranking campaign at 70% ACoS is doing its job. A profit campaign "
        "at 70% ACoS is a disaster. Without goal classification, you can't make good decisions.", styles["Body"]))

    story.append(tbl([
        ["Goal", "Target ACoS", "When to Use", "Typical Campaign Type"],
        ["Ranking", "50–100%", "You want a keyword on page 1 organically. You're willing to lose money "
            "short-term because organic rank will pay off long-term.", "SKC with Fixed Bids, heavy Top of Search modifier"],
        ["Profit", "15–25%", "The keyword already ranks and converts. You want to extract margin.", "Exact match, Dynamic Bids Down Only"],
        ["Research", "30–50%", "You're discovering what customers actually search for. Auto and Broad campaigns "
            "that cast a wide net.", "Auto / Broad / Phrase campaigns"],
        ["Reviews", "40–60%", "You need sales velocity to generate reviews. Higher ACoS is acceptable because "
            "reviews drive long-term organic conversion.", "SKC with Fixed Bids"],
        ["Market Share", "25–40%", "Defending or expanding your position in the category. Aggressive but not reckless.", "Exact match with Fixed Bids"],
        ["Brand Defense", "15–25%", "Competitors are bidding on your brand name. You need to own that real estate cheaply.", "Exact match, Dynamic Bids Down Only"],
    ], [0.85 * inch, 0.75 * inch, 3.0 * inch, 2.4 * inch]))

    story.append(Spacer(1, 8))
    story.append(Paragraph("<b>How Campaigns Get Classified:</b>", styles["Sub3"]))
    story.append(Paragraph(
        "The Campaign Strategist auto-classifies based on naming conventions and match types. If a campaign name "
        "contains 'rank', 'launch', or 'SKC' it's classified as Ranking. If it contains 'brand' it's Brand Defense. "
        "Auto/Broad/Phrase match types default to Research. Exact match with good ACoS defaults to Profit.", styles["Body"]))

    story.append(Paragraph("<b>Campaign Health Check:</b>", styles["Sub3"]))
    story.append(Paragraph(
        "Once classified, each campaign gets a health status by comparing its actual ACoS to its goal's target:", styles["Body"]))

    story.append(tbl([
        ["Status", "Rule", "What It Means", "Action"],
        ["HEALTHY", "ACoS within ±20% of goal target", "Campaign is performing as intended", "Maintain current settings"],
        ["OVER-PERFORMING", "ACoS < target × 0.6", "Making way more profit than needed — could capture more volume", "Raise bids or budget to grow"],
        ["UNDER-PERFORMING", "ACoS > target × 1.5", "Spending too much for what you're getting", "Lower bids, check listing quality"],
        ["MISALIGNED", "Goal doesn't match behavior", "e.g., a 'ranking' campaign running at profit ACoS targets", "Reclassify or restructure"],
        ["DEAD", "0 impressions, but enabled", "Something is wrong — bids too low or targeting dead", "Inch up bids or kill campaign"],
    ], [1.1 * inch, 1.5 * inch, 2.4 * inch, 2.0 * inch]))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # 4. SEARCH TERM FUNNEL
    # ══════════════════════════════════════════════════════════════════════════
    story.extend(sec("4. Search Term Harvesting Funnel"))

    story.append(Paragraph(
        "This is how you discover winning keywords and eliminate losers. Search terms start in broad, cheap "
        "Auto campaigns. As they prove they convert, they get promoted into more controlled campaign types. "
        "At each level, you negate the term in the source campaign so spend doesn't leak backward.", styles["Body"]))

    story.append(Paragraph("<b>The Funnel Flow:</b>", styles["Sub3"]))
    story.append(Paragraph(
        "<b>AUTO</b> (discover new terms) → <b>BROAD</b> (validate with more data) → <b>PHRASE</b> (confirm the pattern) → "
        "<b>EXACT / SKC</b> (maximum control and efficiency)", styles["CallOut"]))

    story.append(Paragraph("<b>Harvesting Decision Table:</b>", styles["Sub3"]))
    story.append(tbl([
        ["Category", "Conditions", "What Happens"],
        ["GOLD → Exact", "Orders ≥ 1, ACoS ≤ target, Clicks ≥ 3\nSource: Auto, Broad, or Phrase", "This term is a proven winner. Promote it to an Exact match SKC for maximum "
            "control. Add it as a negative in the source campaign so you stop paying for it there."],
        ["SILVER → Phrase", "Orders ≥ 1, ACoS ≤ target × 1.5, Clicks ≥ 2\nSource: Auto only", "Promising but not proven enough for Exact yet. Promote to Phrase match "
            "to validate with more data. Negate in Auto."],
        ["NEGATIVE EXACT", "Spend ≥ $10, Orders = 0, Clicks ≥ 5", "This term is eating your budget with zero return. Add it as a negative exact "
            "match so you never pay for it again."],
        ["NEGATIVE PHRASE", "Spend ≥ $5, Orders = 0, Clicks ≥ 3\n3+ word term", "A long-tail term that's not converting. Negate as phrase match to block "
            "the entire pattern, not just the exact term."],
        ["REVIEW", "Orders ≥ 1, ACoS > target × 2, Spend ≥ $5", "Converting but very expensive. Don't negate — it might improve. Keep watching. "
            "May need listing optimization or lower bid."],
        ["DUPLICATE", "Same search term converting in 2+ campaigns", "You're competing against yourself. Consolidate to whichever campaign "
            "performs best, negate in the others."],
    ], [1.1 * inch, 2.3 * inch, 3.6 * inch]))

    story.append(Spacer(1, 8))
    story.append(Paragraph(
        "The suggested starting bid for new exact keywords is calculated as the average CPC from the data × 1.1, "
        "capped at $3.00. This gives the keyword a fighting chance without overpaying.", styles["Body"]))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # 5. BID LIFECYCLE
    # ══════════════════════════════════════════════════════════════════════════
    story.extend(sec("5. Bid Lifecycle — How Every Bid Gets Set"))

    story.append(Paragraph(
        "This is the most critical piece of the system. Every keyword moves through a lifecycle, and the stage "
        "it's in determines how you manage its bid. The biggest mistake sellers make is applying the same bidding "
        "logic to a brand-new keyword (2 clicks) as a proven performer (200 clicks). This system prevents that.", styles["Body"]))

    story.append(tbl([
        ["Stage", "When", "Strategy", "Detailed Logic"],
        ["1. LAUNCH", "Clicks < 5\nSpend < $5\nImpressions may be 0", "INCH UP",
            "Start the bid at $0.30–$0.50. Every 3–7 days, raise by $0.05–$0.10 until impressions start flowing. "
            "NEVER lower or pause a keyword at this stage — you don't have enough data to judge. If impressions "
            "are zero, the bid is simply too low for the auction. Keep inching up."],
        ["2. DISCOVERY", "Clicks 5–9\nNo orders yet", "HOLD",
            "You have some data but not enough to make bid decisions. Hold the bid steady and let clicks "
            "accumulate. It takes roughly 10-15 clicks to have any statistical confidence. Patience here "
            "prevents killing good keywords too early."],
        ["3. OPTIMIZE", "Clicks ≥ 10\nHas conversion data", "REVENUE-BASED",
            "Now you have real data. Calculate the optimal bid using revenue:\n"
            "Revenue Per Click = (Orders ÷ Clicks) × Avg Order Value\n"
            "Target Bid = Revenue Per Click × Target ACoS\n"
            "Apply 50% dampening — only move halfway to target to avoid overcorrection.\n"
            "Clamp between bid floor ($0.20) and ceiling ($5.00)."],
        ["4. SCALE", "ACoS ≤ target × 0.7", "AGGRESSIVE RAISE",
            "This keyword is a winner — ACoS is well below target, meaning every dollar spent returns strong "
            "profit. Raise the bid to capture more impression share and sales volume. Don't leave money on the table."],
        ["5. CUT", "Spend ≥ $15\nClicks ≥ 15\n0 orders", "PAUSE",
            "You've given this keyword enough budget and clicks. It's not converting. Pause it or add as "
            "a negative. Don't keep spending on hope."],
    ], [0.7 * inch, 1.0 * inch, 0.9 * inch, 4.4 * inch]))

    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>Confidence Tiers — How Much of the Adjustment to Apply:</b>", styles["Sub3"]))
    story.append(Paragraph(
        "Even in the Optimize stage, not all data is equal. More clicks = more confidence:", styles["Body"]))

    story.append(tbl([
        ["Clicks", "Confidence", "% of Adjustment Applied", "Why"],
        ["≥ 30", "High", "100%", "Strong data — trust the calculation fully"],
        ["10–29", "Medium", "50%", "Decent data — move cautiously, only half the adjustment"],
        ["5–9", "Hold", "0%", "Not enough data — hold position, gather more clicks"],
        ["< 5", "New / Inch Up", "0%", "Way too early — only inch up bid if no impressions"],
    ], [0.7 * inch, 0.9 * inch, 1.5 * inch, 3.9 * inch]))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # 6. BUDGET + PLACEMENT
    # ══════════════════════════════════════════════════════════════════════════
    story.extend(sec("6. Budget Management"))

    story.append(Paragraph(
        "Budget management answers a simple question: is each campaign getting the right amount of daily budget? "
        "A campaign that's spending 95% of its budget every day and maintaining good ACoS is being <b>throttled</b> — "
        "it could generate more sales if you gave it more budget. Conversely, a campaign spending 5% of its budget "
        "is either broken or the budget is way too high.", styles["Body"]))

    story.append(tbl([
        ["Scenario", "Decision", "Formula", "Reasoning"],
        ["Spending ≥ 80% of budget\nACoS ≤ target × 1.5", "INCREASE", "Budget × 1.25\n(25% increase)", "This campaign is efficient and running out of budget. "
            "Give it room to grow. You're leaving sales on the table."],
        ["Spending ≥ 80% of budget\nACoS > target × 1.5", "DECREASE", "Budget × 0.80\n(min $10)", "Burning through budget but not efficiently. "
            "Reduce budget while you fix bids or targeting."],
        ["Spending ≤ 20% of budget", "INVESTIGATE", "Right-size to:\navg daily spend × 1.5", "Something is off. Bids may be too low, targeting "
            "may be too narrow, or the campaign may need restructuring."],
        ["$0 spend, state = enabled", "ALERT", "Manual review needed", "Campaign is live but getting zero impressions. "
            "Likely a bid or targeting issue."],
        ["20–80% utilization", "NO CHANGE", "Leave as-is", "Budget is appropriately sized for current performance."],
    ], [1.6 * inch, 0.85 * inch, 1.35 * inch, 3.2 * inch]))

    story.append(Spacer(1, 12))
    story.extend(sec("7. Placement Optimization"))

    story.append(Paragraph(
        "Amazon shows your ad in three placements: Top of Search (first row of results), Rest of Search (lower positions), "
        "and Product Pages (competitor listings). Each placement has different conversion rates and costs. Placement optimization "
        "pushes your spend toward whichever placement converts best for each campaign.", styles["Body"]))

    story.append(tbl([
        ["Situation", "Decision", "Why"],
        ["Primary placement ACoS ≤ target × 0.8", "RAISE modifier", "This placement converts well — push more spend here"],
        ["Primary placement ACoS within ±15% of target", "NO CHANGE", "Performing as expected"],
        ["Primary placement ACoS > target × 1.2", "LOWER modifier", "Too expensive for this placement"],
        ["$0 spend on primary placement", "ALERT", "Modifier may be too low, or competition is too high"],
        ["Non-primary spend ≥ $20 AND ACoS > target × 2", "SET TO 0% (Leakage)", "Money is leaking to a bad placement — cut it off"],
        ["Non-primary spend ≥ $20 AND ACoS ≤ target", "KEEP", "Profitable even in the 'wrong' placement — let it run"],
    ], [2.4 * inch, 1.4 * inch, 3.2 * inch]))

    story.append(Spacer(1, 6))
    story.append(Paragraph(
        "<b>Modifier Formula:</b> new_modifier = current + ((target_acos ÷ placement_acos - 1) × 0.5 × 100). "
        "Modifiers are clamped between 0% and 900%.", styles["Body"]))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # 8. SKC BUILDER
    # ══════════════════════════════════════════════════════════════════════════
    story.extend(sec("8. Single Keyword Campaigns (SKCs)"))

    story.append(Paragraph(
        "SKCs are the pinnacle of campaign control. One campaign, one ad group, one ASIN, one keyword in Exact match. "
        "This structure lets you set a unique bid, budget, and placement modifier for every individual keyword. "
        "No other keyword in the same ad group can steal your budget or muddy your data.", styles["Body"]))

    story.append(Paragraph(
        "<b>Naming Convention:</b> [Portfolio] - SP - KW - Exact - [Placement] - [keyword]<br/>"
        "Example: <i>Coffee Cleaner - SP - KW - Exact - ToS - dishwasher cleaner tablets</i>", styles["Body"]))

    story.append(Paragraph("<b>Goal-Based SKC Configuration:</b>", styles["Sub3"]))
    story.append(tbl([
        ["Goal", "Starting Bid", "Daily Budget", "Top of Search Mod", "Bid Strategy", "Philosophy"],
        ["Ranking", "$0.50 (inch-up)", "$25–50", "50–200%", "Fixed Bids", "Maximum ToS visibility. Willing to overpay for rank."],
        ["Profit", "Revenue-based calc", "$10–25", "0–50%", "Dynamic Down", "Let Amazon optimize for conversions. Protect margin."],
        ["Reviews", "$0.75 (inch-up)", "$15–30", "25–100%", "Fixed Bids", "Drive consistent sales velocity for review accumulation."],
        ["Market Share", "Revenue-based calc", "$20–40", "50–150%", "Fixed Bids", "Aggressive visibility across all placements."],
        ["Research", "$0.30 (inch-up)", "$10–20", "0%", "Dynamic Down", "Low risk exploration. Let Amazon find what works."],
    ], [0.75 * inch, 1.0 * inch, 0.75 * inch, 1.0 * inch, 0.9 * inch, 2.6 * inch]))

    story.append(Spacer(1, 6))
    story.append(Paragraph(
        "<b>Effective Bid Calculation:</b> Base Bid × (1 + Placement Modifier%) × Bidding Strategy Multiplier. "
        "Example: $0.50 bid × (1 + 100% ToS modifier) × 1.0 (Fixed) = $1.00 effective bid at Top of Search.", styles["Body"]))
    story.append(Paragraph("Maximum recommended: <b>20–30 SKCs per product</b>. More than that becomes unmanageable.", styles["Body"]))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # 9. RANKING
    # ══════════════════════════════════════════════════════════════════════════
    story.extend(sec("9. Organic Ranking & the Bait-and-Switch"))

    story.append(Paragraph(
        "This is the endgame. You're using PPC to build organic rank, then turning off PPC once rank holds. "
        "The Rank Tracker scores every keyword on a 0–100 scale to tell you how close it is to ranking organically.", styles["Body"]))

    story.append(Paragraph("<b>Ranking Score Calculation (0–100 points):</b>", styles["Sub3"]))
    story.append(tbl([
        ["Dimension", "Weight", "Full Score", "Medium Score", "Low Score"],
        ["Conversion Rate", "30%", "≥ 15% → 30 pts", "10–15% → 20 pts", "< 5% → 0 pts"],
        ["Click-Through Rate", "20%", "≥ 0.5% → 20 pts", "0.3–0.5% → 15 pts", "< 0.1% → 0 pts"],
        ["Top of Search Imp Share", "20%", "≥ 10% → 20 pts", "5–10% → 15 pts", "< 1% → 0 pts"],
        ["Sales Consistency", "15%", "≥ 80% of days → 15 pts", "50–80% → 10 pts", "< 50% → 5 pts"],
        ["Budget Utilization", "15%", "60–90% → 15 pts", "90–100% → 10 pts", "< 60% → 5 pts"],
    ], [1.3 * inch, 0.6 * inch, 1.4 * inch, 1.35 * inch, 1.35 * inch]))

    story.append(Spacer(1, 8))
    story.append(Paragraph("<b>What the Score Means:</b>", styles["Sub3"]))
    story.append(tbl([
        ["Score Range", "Status", "What to Do"],
        ["70–100", "RANKING READY", "This keyword is organically viable. Consider the Bait & Switch test. Monitor organic position."],
        ["50–69", "PROGRESSING", "Getting there. Keep the PPC pressure on. Optimize your listing, pricing, and images to boost CVR."],
        ["30–49", "STRUGGLING", "Something is off. Check listing relevance for this keyword. Is it in your title? Are images compelling?"],
        ["0–29", "NOT RANKING", "This keyword may not be a fit for your product. Consider abandoning it and redirecting budget elsewhere."],
    ], [0.9 * inch, 1.2 * inch, 4.9 * inch]))

    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>The Bait & Switch Test — When to Turn Off PPC:</b>", styles["Sub3"]))
    story.append(Paragraph("ALL 5 conditions must be true before testing:", styles["Body"]))
    story.append(bullet("Ranking Score ≥ 70"))
    story.append(bullet("Sustained sales for ≥ 14 consecutive days"))
    story.append(bullet("Organic impressions are growing (not flat or declining)"))
    story.append(bullet("Top of Search Impression Share > 5%"))
    story.append(bullet("<b>The Test:</b> Cut the campaign's daily budget by 50% for 7 days. If organic rank and sales hold, "
        "you can safely reduce PPC further or turn it off entirely for that keyword."))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # 10. WASTED SPEND + DAY PARTING
    # ══════════════════════════════════════════════════════════════════════════
    story.extend(sec("10. Wasted Spend Detection"))

    story.append(Paragraph(
        "Before optimizing for growth, eliminate waste. These three categories typically account for 15–30% of "
        "total PPC spend in unoptimized accounts:", styles["Body"]))

    story.append(tbl([
        ["Waste Type", "How to Detect", "How to Fix", "Typical Impact"],
        ["Zero-ROI\nKeywords", "Spend > $10, Clicks > 10, Orders = 0\n(for campaigns: Spend > $50, 0 orders)", "Negate the keyword or pause the campaign. "
            "These terms are consuming budget that could go to winners.", "5–15% of total spend"],
        ["Duplicate\nTargeting", "Same keyword active and enabled in 2+ campaigns simultaneously", "Keep the keyword in whichever campaign performs best. "
            "Negate in all others. You're bidding against yourself.", "3–8% of total spend"],
        ["Product\nTargeting Waste", "ASIN targeting with ACoS > target × 2", "Phase out underperforming ASIN targets. Product targeting "
            "tends to have lower conversion rates than keyword targeting.", "2–5% of total spend"],
    ], [1.0 * inch, 2.1 * inch, 2.4 * inch, 1.5 * inch]))

    story.append(Spacer(1, 12))
    story.extend(sec("11. Day Parting Strategy"))

    story.append(Paragraph(
        "Amazon doesn't natively support day parting (scheduling ads by time of day), but you can manually adjust "
        "budgets or use third-party tools. The data consistently shows that conversion rates vary significantly by time block:", styles["Body"]))

    story.append(tbl([
        ["Time Block", "Priority", "Conversion Pattern", "Budget Strategy"],
        ["6 AM – 12 PM", "HIGH", "Morning shoppers with high purchase intent. Researched overnight, buying now.", "Ensure budget is available. Don't exhaust budget before this window."],
        ["6 PM – 10 PM", "HIGH", "Evening shoppers. Second peak after work. Often mobile browsing → purchase.", "This is where most budgets should be flowing."],
        ["12 PM – 6 PM", "MEDIUM", "Moderate activity. Mix of browsing and buying.", "Acceptable spend — don't cut here unless budget is very tight."],
        ["10 PM – 6 AM", "LOW", "Mostly browsing, low conversion. Cheap clicks but poor ROI.", "Reduce spend here if campaigns are burning through budget before evening peak."],
    ], [1.0 * inch, 0.7 * inch, 2.5 * inch, 2.8 * inch]))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # 11. WEEKLY WORKFLOW + INPUTS
    # ══════════════════════════════════════════════════════════════════════════
    story.extend(sec("12. Weekly Workflow"))

    story.append(Paragraph(
        "Run this every week, ideally on the same day. Consistency is more important than perfection.", styles["Body"]))

    steps = [
        ("Download your data", "Go to Amazon Ads and download (1) Bulk Operations file (.xlsx) from Campaign Manager → Bulk Operations, "
            "and (2) Search Term Report (.xlsx) from Measurement & Reporting → Sponsored Ads Reports. Optionally download "
            "Brand Analytics Search Query Performance (.csv) from Seller Central."),
        ("Upload to the dashboard", "Open the Streamlit dashboard, select your brand, and upload all files in the file upload section."),
        ("Run All 8 Tools", "Click the 'Run All 8 Tools' button. The system processes everything in order and generates a Summary Report "
            "combining all outputs into one file."),
        ("Review recommendations", "Start with the Summary Report. Look at: (1) Misaligned campaigns — are goals set correctly? "
            "(2) Wasted spend — what can you cut immediately? (3) Harvesting — what new keywords should be promoted? "
            "(4) Bid changes — are the revenue-based calculations reasonable? (5) Ranking scores — any Bait & Switch candidates?"),
        ("Apply changes", "Download the bid change and budget change files. Upload bulk changes to Amazon Campaign Manager. "
            "Create new SKCs from the SKC Builder output. Add negative keywords from the Harvester."),
        ("Repeat next week", "The system tracks week-over-week trends. Consistency compounds. Most accounts see significant "
            "improvement within 4–6 weeks of disciplined weekly optimization."),
    ]
    for i, (title, desc) in enumerate(steps, 1):
        story.append(Paragraph(f"<b>Step {i}: {title}</b>", styles["Sub3"]))
        story.append(Paragraph(desc, styles["Body"]))

    story.append(Spacer(1, 12))
    story.extend(sec("13. Required Inputs & Configuration"))

    story.append(tbl([
        ["Input", "Required?", "Where to Get It", "Format"],
        ["Bulk Operations", "YES", "Amazon Ads → Campaign Manager → Bulk Operations → Create spreadsheet", ".xlsx"],
        ["Search Term Report", "YES", "Amazon Ads → Measurement & Reporting → Sponsored Ads Reports → Search Term", ".xlsx"],
        ["Brand Analytics SQP", "Optional", "Seller Central → Brands → Brand Analytics → Search Query Performance (ASIN view)", ".csv"],
    ], [1.3 * inch, 0.7 * inch, 4.0 * inch, 0.6 * inch]))

    story.append(Spacer(1, 8))
    story.append(Paragraph("<b>Brand Configuration (set once per brand in the dashboard):</b>", styles["Sub3"]))
    story.append(bullet("Brand name and brand keywords (for brand vs non-brand segmentation)"))
    story.append(bullet("Primary ASIN (for rank tracking)"))
    story.append(bullet("Target ACoS by goal — Profit (default 25%), Ranking (60%), Research (35%), Reviews (40%), Market Share (30%)"))
    story.append(bullet("Portfolio names (for campaign organization)"))

    story.append(Spacer(1, 8))
    story.append(Paragraph("<b>Global Defaults (adjustable in Settings page):</b>", styles["Sub3"]))

    story.append(tbl([
        ["Setting", "Default", "Setting", "Default", "Setting", "Default"],
        ["Bid Floor", "$0.20", "Max Bid Raise", "30%", "Pause Spend", "$15"],
        ["Bid Ceiling", "$5.00", "Max Bid Lower", "40%", "Pause Clicks", "15"],
        ["Min Clicks (Bid)", "10", "Min Clicks (Harvest)", "3", "Neg Spend Threshold", "$10"],
        ["Budget Constrained", "80%", "Budget Underutil", "20%", "Min Daily Budget", "$10"],
        ["Budget Increase", "25%", "Budget Decrease", "20%", "Lookback Days", "30"],
        ["Placement Leakage", "$20", "Max Modifier Up", "50pp", "Max Modifier Down", "100pp"],
    ], [1.2 * inch, 0.6 * inch, 1.3 * inch, 0.6 * inch, 1.3 * inch, 0.6 * inch]))

    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # 12. MASTER DECISION TREE
    # ══════════════════════════════════════════════════════════════════════════
    story.extend(sec("14. Master Decision Tree"))

    story.append(Paragraph(
        "This is the complete decision flow for every keyword, every week. If you only read one page of this "
        "document, make it this one.", styles["Body"]))

    story.append(tbl([
        ["Step", "Question", "Decision Logic"],
        ["1", "What GOAL is this\nkeyword serving?", "Determines the ACoS target. Ranking = 50-100%, Profit = 15-25%, Research = 30-50%, "
            "Reviews = 40-60%, Market Share = 25-40%, Brand Defense = 15-25%. If goal is unclear, classify based on "
            "campaign name and match type."],
        ["2", "How many CLICKS\ndoes it have?", "< 5 clicks: INCH UP the bid $0.05-$0.10. Don't make any other changes. Way too early.\n"
            "5-9 clicks: HOLD. Gather more data before making decisions.\n"
            "10+ clicks: Proceed to Step 3 — you have enough data to act."],
        ["3", "Is it CONVERTING?", "YES, ACoS ≤ target: RAISE bid (revenue-based formula), INCREASE budget if constrained.\n"
            "YES, ACoS > target: LOWER bid, check placement leakage, review listing quality.\n"
            "NO orders, ≥ 15 clicks, ≥ $15 spend: PAUSE this keyword — it's not going to convert."],
        ["4", "Is it in the right\nCAMPAIGN TYPE?", "Is this winner sitting in an Auto or Broad campaign? HARVEST it to Exact/SKC.\n"
            "Is the same term active in multiple campaigns? CONSOLIDATE to best performer, negate in others.\n"
            "Is the campaign goal misaligned with its behavior? RECLASSIFY."],
        ["5", "Is the PLACEMENT\nworking?", "Is the primary placement (e.g., Top of Search) profitable? RAISE the modifier.\n"
            "Is spend leaking to a non-primary placement with bad ACoS? CUT the modifier to 0%.\n"
            "Is the campaign burning budget before evening peak hours? Consider day parting."],
        ["6", "Is it RANKING\norganically?", "Ranking Score ≥ 70: Candidate for Bait & Switch. Test by cutting budget 50% for 7 days.\n"
            "Score 50-69: Progressing. Keep the PPC pressure on and optimize the listing.\n"
            "Score 30-49: Struggling. Something is off — check relevance, images, price.\n"
            "Score < 30: Not viable. Redirect budget to keywords with more potential."],
    ], [0.5 * inch, 1.3 * inch, 5.2 * inch]))

    story.append(Spacer(1, 0.3 * inch))
    story.append(HRFlowable(width="100%", thickness=3, color=ORANGE, spaceAfter=8, spaceBefore=0))

    story.append(Paragraph(
        "This system is designed to be run weekly. The 8 tools automate the analysis — your job is to review the "
        "recommendations, apply the changes, and stay consistent. Most accounts see meaningful improvement within "
        "4–6 weeks. The compounding effect of weekly harvesting, bid optimization, and waste elimination is "
        "significant over time.", styles["Body"]))

    story.append(Spacer(1, 12))
    story.append(Paragraph(
        "Dashboard: amazon-ppc-tools-hfntjz43nzejvfjfautqyr.streamlit.app  |  GitHub: github.com/Shirtmagic/Amazon-PPC-Tools",
        styles["Foot"]))

    doc.build(story)
    print(f"PDF generated: {OUTPUT}")


if __name__ == "__main__":
    build_pdf()
