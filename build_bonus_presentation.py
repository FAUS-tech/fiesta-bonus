"""
PPT presentation - simplified bonus proposal deck.

~22 slides. Plan C removed. Per-agent month-by-month detail added with
policy counts. New 'Why Current Drops' explainer.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from build_bonus_workbook import (
    AGENTS, MONTHS, AGENT_NAMES, NB_MIN_PREMIUM, REN_MIN_RETENTION,
    BLENDED_COMM, ROYALTY, OVERHEAD, CAP_STANDARD, AGENT_SALARY, RETENTION_POOL_RATE,
    calc_proposal_a, current_bonus, swap_rwr_to_ren, full_flip,
    nb_tier_bonus, rwr_tier_bonus, ren_tier_bonus,
)

# COLOR PALETTE
NAVY = RGBColor(0x1F, 0x4E, 0x78)
LIGHT_BLUE = RGBColor(0xBD, 0xD7, 0xEE)
ACCENT_BLUE = RGBColor(0x2E, 0x75, 0xB6)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x10, 0x10, 0x10)
DARK_GRAY = RGBColor(0x40, 0x40, 0x40)
LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
GREEN = RGBColor(0x70, 0xAD, 0x47)
LIGHT_GREEN = RGBColor(0xE2, 0xEF, 0xDA)
GOLD = RGBColor(0xBF, 0x90, 0x00)
LIGHT_GOLD = RGBColor(0xFF, 0xF2, 0xCC)
ORANGE = RGBColor(0xED, 0x7D, 0x31)
LIGHT_ORANGE = RGBColor(0xFC, 0xE4, 0xD6)
RED = RGBColor(0xC0, 0x00, 0x00)
LIGHT_RED = RGBColor(0xFF, 0xC7, 0xCE)


def add_slide(prs, layout_idx=6):
    return prs.slides.add_slide(prs.slide_layouts[layout_idx])


def add_text(slide, left, top, width, height, text, font_size=18, bold=False,
             color=BLACK, align=PP_ALIGN.LEFT, fill=None, anchor=MSO_ANCHOR.TOP, italic=False):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.08); tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.04); tf.margin_bottom = Inches(0.04)
    tf.vertical_anchor = anchor
    if fill:
        box.fill.solid(); box.fill.fore_color.rgb = fill; box.line.fill.background()
    else:
        box.fill.background(); box.line.fill.background()
    p = tf.paragraphs[0]; p.alignment = align
    run = p.add_run(); run.text = text
    run.font.size = Pt(font_size); run.font.bold = bold; run.font.italic = italic
    run.font.color.rgb = color; run.font.name = 'Calibri'
    return box


def add_bullets(slide, left, top, width, height, items, font_size=16, color=BLACK,
                fill=None, line_spacing=1.2):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.1); tf.margin_right = Inches(0.1)
    tf.margin_top = Inches(0.1); tf.margin_bottom = Inches(0.1)
    if fill:
        box.fill.solid(); box.fill.fore_color.rgb = fill; box.line.fill.background()
    else:
        box.fill.background(); box.line.fill.background()
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = "• " + item
        run.font.size = Pt(font_size); run.font.color.rgb = color
        run.font.name = 'Calibri'
        p.line_spacing = line_spacing
    return box


def add_bar(slide, left, top, width, height, fill_color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid(); shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def add_table(slide, left, top, width, height, data, header_fill=NAVY, header_text_color=WHITE,
              col_widths=None, font_size=11, row_height_in=0.32, highlight_col=None, highlight_color=None,
              first_col_bold=False):
    rows = len(data)
    cols = len(data[0]) if data else 0
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table
    if col_widths:
        for i, w in enumerate(col_widths):
            if i < cols: table.columns[i].width = w
    for r in range(rows):
        table.rows[r].height = Inches(row_height_in)
        for c in range(cols):
            cell = table.cell(r, c)
            cell.text = ''
            tf = cell.text_frame
            tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
            tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if r == 0 or c == 0 else PP_ALIGN.CENTER
            run = p.add_run(); run.text = str(data[r][c]); run.font.name = 'Calibri'
            run.font.size = Pt(font_size)
            if r == 0:
                cell.fill.solid(); cell.fill.fore_color.rgb = header_fill
                run.font.bold = True; run.font.color.rgb = header_text_color
            else:
                if highlight_col is not None and c == highlight_col and highlight_color:
                    cell.fill.solid(); cell.fill.fore_color.rgb = highlight_color
                elif r % 2 == 0:
                    cell.fill.solid(); cell.fill.fore_color.rgb = LIGHT_GRAY
                else:
                    cell.fill.solid(); cell.fill.fore_color.rgb = WHITE
                run.font.color.rgb = BLACK
                if first_col_bold and c == 0:
                    run.font.bold = True
    return table_shape


def header_strip(slide, title, subtitle=None):
    add_bar(slide, 0, 0, Inches(13.33), Inches(0.95), NAVY)
    add_text(slide, Inches(0.4), Inches(0.12), Inches(12.5), Inches(0.5), title,
             font_size=24, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text(slide, Inches(0.4), Inches(0.58), Inches(12.5), Inches(0.35), subtitle,
                 font_size=12, italic=True, color=LIGHT_BLUE, anchor=MSO_ANCHOR.MIDDLE)


def footer(slide, page_num, total):
    add_bar(slide, 0, Inches(7.27), Inches(13.33), Inches(0.23), LIGHT_GRAY)
    add_text(slide, Inches(0.4), Inches(7.27), Inches(8), Inches(0.23),
             "Fiesta Bonus Plan - Ownership Decision Deck",
             font_size=10, color=DARK_GRAY, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(12.4), Inches(7.27), Inches(0.8), Inches(0.23),
             f"{page_num} / {total}", font_size=10, color=DARK_GRAY,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def compute_totals():
    real = {'cur':0,'a':0,'a_min':0,'b':0,'b_min':0}
    swap = {'cur':0,'a':0,'a_min':0,'b':0,'b_min':0}
    flip = {'cur':0,'a':0,'a_min':0,'b':0,'b_min':0}
    per_agent = {a: {'cur_r':0,'cur_s':0,'cur_f':0,
                     'a_r':0,'a_s':0,'a_f':0,'a_min_r':0,'a_min_s':0,'a_min_f':0,
                     'b_r':0,'b_s':0,'b_f':0,'b_min_r':0,'b_min_s':0,'b_min_f':0} for a in AGENT_NAMES}
    for agent in AGENT_NAMES:
        for month in MONTHS:
            d = AGENTS[agent][month]
            ds = swap_rwr_to_ren(d, 0.5)
            df = full_flip(d)
            for sc, scd, key in [(real,d,'r'),(swap,ds,'s'),(flip,df,'f')]:
                cur = current_bonus(scd['NB'][0], scd['RWR'][0])
                ra = calc_proposal_a(scd['NB'], scd['RWR'], scd['REN'])
                rb = calc_proposal_a(scd['NB'], scd['RWR'], scd['REN'])
                sc['cur']+=cur
                sc['a']+=ra['paid']; sc['a_min']+=ra['paid_after_min']
                sc['b']+=rb['paid']; sc['b_min']+=rb['paid_after_min']
                per_agent[agent][f'cur_{key}']+=cur
                per_agent[agent][f'a_{key}']+=ra['paid']
                per_agent[agent][f'a_min_{key}']+=ra['paid_after_min']
                per_agent[agent][f'b_{key}']+=rb['paid']
                per_agent[agent][f'b_min_{key}']+=rb['paid_after_min']
    return real, swap, flip, per_agent


# ============================================================================
# SLIDES
# ============================================================================
def slide_cover(prs):
    s = add_slide(prs)
    add_bar(s, 0, 0, Inches(13.33), Inches(7.5), NAVY)
    add_bar(s, 0, Inches(3.4), Inches(13.33), Inches(0.1), GOLD)
    add_text(s, Inches(0.6), Inches(1.0), Inches(12.13), Inches(1.0),
             "Fiesta Bonus Plan", font_size=54, bold=True, color=WHITE)
    add_text(s, Inches(0.6), Inches(2.0), Inches(12.13), Inches(0.8),
             "Final Proposal for Ownership", font_size=28, color=LIGHT_BLUE)
    add_text(s, Inches(0.6), Inches(3.8), Inches(12.13), Inches(0.55),
             "Premium-Tiered Base + Per-Policy Collected Incentive + Book Retention Bonus",
             font_size=20, color=WHITE)
    add_text(s, Inches(0.6), Inches(4.5), Inches(12.13), Inches(0.55),
             f"With ${NB_MIN_PREMIUM:,} NB premium / {REN_MIN_RETENTION*100:.0f}% retention minimums",
             font_size=18, color=LIGHT_BLUE)
    add_text(s, Inches(0.6), Inches(5.3), Inches(12.13), Inches(0.55),
             "Modeled on actual Jan-Apr 2026 production: Abel, Dialinerys, Melissa, Flavia, Thalia, Monica",
             font_size=15, italic=True, color=LIGHT_BLUE)


def slide_glossary(prs, idx, total):
    s = add_slide(prs)
    header_strip(s, "Glossary - What These Terms Mean",
                 "Skip this slide if you know the insurance jargon.")
    footer(s, idx, total)

    terms = [
        ["Term", "What it means"],
        ["NB", "NEW BUSINESS - first time the customer is on the books."],
        ["RWR", "REWRITE - cancel an existing policy and write a NEW one (usually at a different carrier). Churns the book."],
        ["REN", "RENEWAL - the existing policy renews at the SAME carrier. Customer stays put."],
        ["PIF", "PAID IN FULL - customer pays the entire premium upfront. Zero chargeback risk."],
        ["BI", "BODILY INJURY liability - pays the OTHER person's injuries if your customer caused the accident."],
        ["UM", "UNINSURED MOTORIST - pays YOUR customer if hit by an uninsured driver. Cannot exist without BI."],
        ["PIP / PD", "PERSONAL INJURY PROTECTION and PROPERTY DAMAGE liability. Required in Florida."],
        ["Comp / Coll", "COMPREHENSIVE / COLLISION - covers your customer's OWN car damage."],
        ["Collected %", "% of premium the customer actually paid. Down payment / total = collected %."],
        ["Safe Net", "Money left after corporate royalty (21%) and overhead (30%). Bonus comes from this pool."],
        ["Chargeback", "If a policy cancels within 90 days, the bonus is reversed."],
        ["Kicker", "Bonus multiplier on the collected %. Higher down = bigger multiplier."],
        ["Minimum / Gate", "A monthly premium threshold the agent must hit to earn that bonus line."],
        ["FLIP scenario", "Hypothetical: 'what if ALL rewrites had been renewals?'"],
        ["SWAP scenario", "Hypothetical: 'what if HALF the rewrites became renewals?'"],
    ]
    add_table(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.8), terms,
              header_fill=NAVY, col_widths=[Inches(2.0), Inches(10.3)],
              font_size=12, row_height_in=0.36, first_col_bold=True)


def slide_problem(prs, idx, total):
    s = add_slide(prs)
    header_strip(s, "Why We're Changing the Plan",
                 "The current plan pays for activity, not for the right activity.")
    footer(s, idx, total)

    add_text(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(0.5),
             "What the current plan does today:", font_size=18, bold=True, color=NAVY)
    issues = [
        "Counts NB and RWR together toward a 35-policy tier - so rewriting and writing new are the SAME.",
        "Pays nothing for renewals - agents have zero reason to keep a customer with the current carrier.",
        "Pays nothing for collected money - a $50 down looks the same as a $500 down at payout time.",
        "Pays nothing for coverage upsell (BI / UM / Comp/Coll).",
        "Pays nothing for PIF (paid in full) - even though PIF has the least chargeback risk.",
        "Has no profitability tie - the bonus is set without reference to commission economics.",
    ]
    add_bullets(s, Inches(0.7), Inches(1.7), Inches(12.0), Inches(2.3), issues, font_size=15)

    add_text(s, Inches(0.5), Inches(4.2), Inches(12.3), Inches(0.5),
             "The cost of that design:", font_size=18, bold=True, color=NAVY)
    costs = [
        "Agents rewrite customers between carriers to grow their count - even when the customer would have renewed.",
        "Customer is moved, the policy starts a NEW 6-month term, persistency drops.",
        "The agency loses the renewal commission it would have earned by keeping the customer in place.",
        "We are paying premium bonus dollars for the worst kind of activity: churning the book.",
    ]
    add_bullets(s, Inches(0.7), Inches(4.7), Inches(12.0), Inches(2.0), costs, font_size=15, color=DARK_GRAY)


def slide_goals(prs, idx, total):
    s = add_slide(prs)
    header_strip(s, "What We Want the New Plan to Do",
                 "Four behaviors. Each one is rewarded.")
    footer(s, idx, total)
    goals = [
        ("WRITE", "More new business - but only the kind we can actually keep.", GREEN),
        ("COLLECT", "Bigger down payments and more PIF - less chargeback, more retention.", ACCENT_BLUE),
        ("RETAIN", "Renew customers instead of rewriting them. Loyalty is a real number.", GOLD),
        ("PROFIT", "Bonus comes out of safe-net commission, not out of owner's pocket.", ORANGE),
    ]
    y = Inches(1.3)
    for i, (label, desc, color) in enumerate(goals):
        card_left = Inches(0.5 + (i % 2) * 6.2)
        card_top = y + Inches((i // 2) * 2.6)
        add_bar(s, card_left, card_top, Inches(6.0), Inches(2.4), LIGHT_GRAY)
        add_bar(s, card_left, card_top, Inches(0.2), Inches(2.4), color)
        add_text(s, card_left + Inches(0.4), card_top + Inches(0.15), Inches(5.4), Inches(0.6),
                 label, font_size=28, bold=True, color=color)
        add_text(s, card_left + Inches(0.4), card_top + Inches(0.85), Inches(5.4), Inches(1.4),
                 desc, font_size=17, color=BLACK)


def slide_two_plans_at_a_glance(prs, idx, total):
    s = add_slide(prs)
    header_strip(s, "Two Proposed Plans At a Glance",
                 "Premium-Based and Coverage-Based. Don't mix them.")
    footer(s, idx, total)

    # Plan A
    add_bar(s, Inches(0.5), Inches(1.2), Inches(6.0), Inches(5.7), LIGHT_GREEN)
    add_bar(s, Inches(0.5), Inches(1.2), Inches(6.0), Inches(0.8), GREEN)
    add_text(s, Inches(0.7), Inches(1.25), Inches(5.6), Inches(0.7),
             "PLAN A - PREMIUM-BASED", font_size=22, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(0.7), Inches(2.1), Inches(5.6), Inches(0.5),
             "Bonus by written-premium tier", font_size=15, italic=True, color=NAVY)
    a_items = [
        "NB pay scales with premium ($5 / $7 / $9 / $11 / $11+).",
        "REN pays $4 / $5 / $6 by tier (lower than NB, on purpose).",
        "RWR is flat $2 - no longer drives bonus.",
        "Collected % kicker (+10% / +15% / +20% / +25%).",
        "PIF add-on: +$8 NB / +$12 high NB / +$5 REN.",
        "Simple payroll math. Lower payout overall.",
    ]
    add_bullets(s, Inches(0.7), Inches(2.7), Inches(5.6), Inches(4.0), a_items, font_size=13)

    # Plan B
    add_bar(s, Inches(6.83), Inches(1.2), Inches(6.0), Inches(5.7), LIGHT_GOLD)
    add_bar(s, Inches(6.83), Inches(1.2), Inches(6.0), Inches(0.8), GOLD)
    add_text(s, Inches(7.03), Inches(1.25), Inches(5.6), Inches(0.7),
             "PLAN B - COVERAGE-BASED (RECOMMENDED)", font_size=20, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(7.03), Inches(2.1), Inches(5.6), Inches(0.5),
             "Bonus by coverage type, not premium", font_size=15, italic=True, color=NAVY)
    b_items = [
        "NB base $10 (PIP/PD OR PIP+Comp/Coll).",
        "NB liability bundle +$5 if BI+UM both present.",
        "BI alone does NOT earn the bundle (UM cannot exist without BI).",
        "REN base $6, +$3 liability bundle.",
        "RWR flat $2 - same as Plan A.",
        "Same kicker + same PIF add-ons.",
        "Best motivator. Closest to today's pay when behavior shifts.",
    ]
    add_bullets(s, Inches(7.03), Inches(2.7), Inches(5.6), Inches(4.0), b_items, font_size=13)

    add_text(s, Inches(0.5), Inches(7.0), Inches(12.3), Inches(0.25),
             f"(only one plan now - The Bonus Plan)",
             font_size=11, italic=True, color=DARK_GRAY, align=PP_ALIGN.CENTER)


def slide_plan_a_details_new(prs, idx, total):
    """The Bonus Plan: monthly premium-tier ladders for NB, RWR, REN."""
    s = add_slide(prs)
    header_strip(s, "THE BONUS PLAN",
                 "Same idea as today's 30/38/50 count tiers, but driven by monthly WRITTEN PREMIUM. $1,000 NB at $100k.")
    footer(s, idx, total)

    # NB tier ladder
    nb_data = [
        ["Tier", "Monthly NB premium", "NB pays"],
        ["T1 (min req)", "$45,000", "$250"],
        ["T2", "$55,000", "$375"],
        ["T3", "$70,000", "$525"],
        ["T4", "$85,000", "$725"],
        ["T5", "$100,000", "$1,000"],
        ["Above $100k", "+$5k of premium", "+$50 (1%)"],
    ]
    add_text(s, Inches(0.4), Inches(1.05), Inches(6.2), Inches(0.3),
             "1. NEW BUSINESS - 5 tiers by monthly written premium", font_size=13, bold=True, color=NAVY)
    add_table(s, Inches(0.4), Inches(1.4), Inches(6.2), Inches(2.5), nb_data,
              header_fill=GREEN, col_widths=[Inches(1.6), Inches(2.4), Inches(2.2)],
              font_size=11, row_height_in=0.32)

    # RWR tier ladder (~half of NB)
    rwr_data = [
        ["Tier", "Monthly RWR premium", "RWR pays (~half of NB)"],
        ["T1", "$45,000", "$100"],
        ["T2", "$55,000", "$200"],
        ["T3", "$70,000", "$275"],
        ["T4", "$85,000", "$350"],
        ["T5", "$100,000", "$500"],
        ["Above $100k", "+$5k of premium", "+$25 (0.5%)"],
    ]
    add_text(s, Inches(6.85), Inches(1.05), Inches(6.2), Inches(0.3),
             "2. REWRITES - same breakpoints, pays ~HALF (gated by NB $45k)", font_size=13, bold=True, color=NAVY)
    add_table(s, Inches(6.85), Inches(1.4), Inches(6.2), Inches(2.5), rwr_data,
              header_fill=ACCENT_BLUE, col_widths=[Inches(1.4), Inches(2.4), Inches(2.4)],
              font_size=11, row_height_in=0.32)

    # REN tier ladder (own breakpoints starting at $25k)
    add_text(s, Inches(0.4), Inches(4.1), Inches(12.7), Inches(0.3),
             "3. RENEWALS - own tiers starting at $25k, gated by 30% retention rate", font_size=13, bold=True, color=NAVY)
    ren_data = [
        ["Tier", "T1", "T2", "T3", "T4", "T5", "Above $100k"],
        ["Monthly REN premium", "$25,000", "$40,000", "$65,000", "$80,000", "$100,000", "+$5k of premium"],
        ["REN pays", "$250", "$350", "$450", "$550", "$800", "+$40 (0.8%)"],
    ]
    add_table(s, Inches(0.4), Inches(4.45), Inches(12.7), Inches(1.4), ren_data,
              header_fill=GOLD, col_widths=[Inches(2.6), Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.6), Inches(2.5)],
              font_size=11, row_height_in=0.40, first_col_bold=True)

    # Section 4: Rules on top
    add_text(s, Inches(0.4), Inches(6.05), Inches(12.7), Inches(0.3),
             "4. RULES ON TOP", font_size=13, bold=True, color=NAVY)
    rules_data = [
        ["Rule", "Value", "What it does"],
        [f"NB monthly minimum", f"${NB_MIN_PREMIUM:,} NB premium", "Equals T1 entry. NB < $45k -> NB pay = $0 (also blocks RWR)"],
        [f"REN monthly minimum", f"{REN_MIN_RETENTION*100:.0f}% retention rate of book", f"Retain < 30% -> REN pay = $0"],
        ["Chargeback", "3 months (90 days)", "100% bonus reversed if policy cancels/rewrites within 90 days"],
    ]
    add_table(s, Inches(0.4), Inches(6.40), Inches(12.7), Inches(1.10), rules_data,
              header_fill=NAVY, col_widths=[Inches(2.8), Inches(3.4), Inches(6.5)],
              font_size=11, row_height_in=0.26)


def slide_plan_a_details(prs, idx, total):
    s = add_slide(prs)
    header_strip(s, "Plan A - Premium-Based Rules",
                 "Per-policy bonus by written-premium tier. Then collected-% kicker.")
    footer(s, idx, total)

    nb_data = [
        ["NB Tier (Written Premium)", "Pay per Policy", "Plain English"],
        ["Under $1,200", "$6", "Low-premium NB earns base"],
        ["$1,200 - $1,799", "$8", "Standard NB"],
        ["$1,800 - $2,199", "$11", "Higher premium earns more"],
        ["$2,200 - $2,999", "$13", "Strong premium"],
        ["$3,000+", "$13 + $2 per $1k (cap $28)", "Commercial / high-premium upside"],
    ]
    add_table(s, Inches(0.5), Inches(1.2), Inches(7.5), Inches(2.6), nb_data,
              header_fill=GREEN, col_widths=[Inches(2.6), Inches(2.4), Inches(2.5)],
              font_size=12, row_height_in=0.36)

    ren_data = [
        ["REN Tier", "Pay per Policy", "RWR"],
        ["Under $1,200", "$5", "$2 flat (any RWR)"],
        ["$1,200 - $1,799", "$6", "No tier"],
        ["$1,800+", "$7", "No coverage stacking"],
    ]
    add_table(s, Inches(8.3), Inches(1.2), Inches(4.6), Inches(1.7), ren_data,
              header_fill=GOLD, col_widths=[Inches(1.8), Inches(1.4), Inches(1.4)],
              font_size=12, row_height_in=0.36)

    addons = [
        ["Add-on", "Amount", "When"],
        ["Collected Kicker", "+10/+15/+20/+25%", "15-24 / 25-49 / 50-99 / 100% collected"],
        ["PIF Add (NB <$3k)", "+$9", "Policy paid in full"],
        ["PIF Add (NB $3k+)", "+$13", "Policy paid in full"],
        ["PIF Add (REN)", "+$5", "Renewal paid in full"],
    ]
    add_table(s, Inches(0.5), Inches(4.0), Inches(7.5), Inches(2.0), addons,
              header_fill=ACCENT_BLUE, col_widths=[Inches(2.6), Inches(2.4), Inches(2.5)],
              font_size=12, row_height_in=0.35)

    add_bar(s, Inches(8.3), Inches(3.1), Inches(4.6), Inches(3.8), LIGHT_GRAY)
    add_text(s, Inches(8.45), Inches(3.15), Inches(4.3), Inches(0.4),
             "WORKED EXAMPLE", font_size=14, bold=True, color=NAVY)
    ex_lines = [
        "$2,000 NB (BI+UM, 25% collected, not PIF)",
        "",
        "Tier: $1,800-$2,199 = $11",
        "Coverage add: $0 (Plan A has no coverage)",
        "Kicker: 25% collected = +15% (x 1.15)",
        "PIF: not PIF, $0",
        "",
        "PER-POLICY PAY: $11 x 1.15 = $12.65",
    ]
    y = Inches(3.55)
    for line in ex_lines:
        bold = line.startswith("PER-POLICY") or line.startswith("$2,000")
        add_text(s, Inches(8.45), y, Inches(4.3), Inches(0.3), line,
                 font_size=13, bold=bold, color=NAVY if bold else BLACK)
        y += Inches(0.32)


def slide_plan_b_details(prs, idx, total):
    s = add_slide(prs)
    header_strip(s, "Plan B - Coverage-Based Rules (RECOMMENDED)",
                 "Per-policy bonus by COVERAGE TYPE. No premium tiers. Same kicker + PIF.")
    footer(s, idx, total)

    cov_data = [
        ["Bonus Line", "Per Policy", "When It Pays"],
        ["NB BASE", "$11", "PIP/PD ONLY, or PIP+Comp/Coll (any auto NB)"],
        ["NB LIABILITY ADD", "+$6", "BI AND UM both on the policy (bundle)"],
        ["BI ALONE", "$0 add", "BI without UM does NOT earn the bundle"],
        ["UM ALONE", "Not possible", "FL rules block UM without BI"],
        ["REN BASE", "$7", "Any renewal"],
        ["REN LIABILITY ADD", "+$3", "BI+UM both on the renewal"],
        ["RWR", "$2 flat", "Any rewrite (no coverage stacking)"],
    ]
    add_table(s, Inches(0.5), Inches(1.2), Inches(7.5), Inches(3.4), cov_data,
              header_fill=GOLD, col_widths=[Inches(2.4), Inches(1.6), Inches(3.5)],
              font_size=12, row_height_in=0.40)

    addons = [
        ["Add-on", "Amount", "When"],
        ["Collected Kicker", "+10/+15/+20/+25%", "Same as Plan A"],
        ["PIF NB (<$3k)", "+$9", "Paid in full"],
        ["PIF NB ($3k+)", "+$13", "Paid in full"],
        ["PIF REN", "+$5", "Renewal paid in full"],
    ]
    add_table(s, Inches(0.5), Inches(4.8), Inches(7.5), Inches(1.9), addons,
              header_fill=ACCENT_BLUE, col_widths=[Inches(2.4), Inches(1.6), Inches(3.5)],
              font_size=12, row_height_in=0.35)

    add_bar(s, Inches(8.3), Inches(1.2), Inches(4.6), Inches(5.5), LIGHT_GRAY)
    add_text(s, Inches(8.45), Inches(1.3), Inches(4.3), Inches(0.4),
             "WORKED EXAMPLE", font_size=14, bold=True, color=NAVY)
    ex_lines = [
        "$2,000 NB (BI+UM, 25% collected, not PIF)",
        "",
        "Base coverage: $11",
        "Liability bundle (BI+UM): +$6",
        "Subtotal: $17",
        "Kicker: 25% collected = +15% (x 1.15)",
        "PIF: not PIF, $0",
        "",
        "PER-POLICY PAY: $17 x 1.15 = $19.55",
        "",
        "Same policy on Plan A: $12.65",
        "Plan B rewards coverage upsell directly.",
    ]
    y = Inches(1.75)
    for line in ex_lines:
        bold = line.startswith("PER-POLICY") or line.startswith("$2,000")
        c = NAVY if bold else BLACK
        if line.startswith("Same policy") or line.startswith("Plan B"):
            c = DARK_GRAY
        add_text(s, Inches(8.45), y, Inches(4.3), Inches(0.3), line,
                 font_size=13, bold=bold, color=c, italic=line.startswith("Plan B"))
        y += Inches(0.32)


def slide_minimums(prs, idx, total):
    s = add_slide(prs)
    header_strip(s, "Minimum Requirements (NEW)",
                 f"NB premium >= ${NB_MIN_PREMIUM:,}/mo + REN retention >= {REN_MIN_RETENTION*100:.0f}% of book. RWR follows NB gate.")
    footer(s, idx, total)

    add_text(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(0.4),
             "Why we added a minimum at all:", font_size=18, bold=True, color=NAVY)
    add_text(s, Inches(0.7), Inches(1.65), Inches(12.0), Inches(0.85),
             "The old plan required 35 NB+RWR policies. That counts rewriting. The new minimum is premium-based and SPLITS NB from REN - so the agent must be writing new business AND building a renewal book. RWR has no own minimum because we don't reward rewriting volume directly.",
             font_size=14, color=DARK_GRAY)

    # Pass rate table
    pass_data = [
        ["Gate", "REAL months pass (24)", "SWAP months pass (24)", "Plain English"],
        [f"NB >= ${NB_MIN_PREMIUM:,}", "4/24 (17%)", "4/24 (17%)", "Top performers only - $45k is a stretch goal"],
        [f"REN retention >= {REN_MIN_RETENTION*100:.0f}%", "modeled 75% (75% >= 30%)", "75% >= 30% pass", "Industry typical retention; model assumes 75%"],
        ["BOTH gates (for RWR)", "0 (0%)", "6 (25%)", "Once renewal book exists, RWR also pays"],
    ]
    add_table(s, Inches(0.5), Inches(2.7), Inches(12.3), Inches(1.9), pass_data,
              header_fill=NAVY, col_widths=[Inches(2.5), Inches(2.5), Inches(2.5), Inches(4.8)],
              font_size=12, row_height_in=0.40)

    # Salary coverage at $35k/$20k = $55k total
    add_bar(s, Inches(0.5), Inches(4.8), Inches(12.3), Inches(2.1), LIGHT_GRAY)
    add_text(s, Inches(0.7), Inches(4.85), Inches(12.0), Inches(0.4),
             "DOES THE MINIMUM COVER AGENT SALARY?", font_size=14, bold=True, color=NAVY)
    sal_items = [
        f"Agent salary basis: $16-$22/hr x 40 hr/wk = $2,773-$3,813/mo (midpoint ${AGENT_SALARY:,}).",
        f"At 25% collected (today's typical): ${NB_MIN_PREMIUM*0.25:,.0f} collected x 11% x (1-21%) = ~${NB_MIN_PREMIUM*0.25*0.11*(1-ROYALTY):,.0f} retained ({NB_MIN_PREMIUM*0.25*0.11*(1-ROYALTY)/AGENT_SALARY*100:.0f}% of salary).",
        f"At 50% collected (typical push target): ~${NB_MIN_PREMIUM*0.50*0.11*(1-ROYALTY):,.0f} retained ({NB_MIN_PREMIUM*0.50*0.11*(1-ROYALTY)/AGENT_SALARY*100:.0f}% of salary).",
        f"At 75% collected (top performer): ~${NB_MIN_PREMIUM*0.75*0.11*(1-ROYALTY):,.0f} retained ({NB_MIN_PREMIUM*0.75*0.11*(1-ROYALTY)/AGENT_SALARY*100:.0f}% of salary) - covers salary AND generates profit.",
        "The MINIMUM is the floor. The COLLECTED KICKER is the lever to cover salary and earn bonus on top.",
    ]
    add_bullets(s, Inches(0.7), Inches(5.3), Inches(12.0), Inches(1.6), sal_items, font_size=11)


def slide_kicker(prs, idx, total):
    s = add_slide(prs)
    header_strip(s, "Collected % Kicker",
                 "Bigger down payments = less chargeback = bigger bonus.")
    footer(s, idx, total)

    kicker_data = [
        ["Collected %", "Per-policy kicker", "Plain English", "$7 NB base -> Pays"],
        ["Below 15%", "$0 (none)", "Premium barely collected. No kicker.", "$7"],
        ["15% - 19%", "+$1", "Minimum down. Small bump.", "$7 + $1 = $8"],
        ["20% - 24%", "+$2", "Standard down.", "$7 + $2 = $9"],
        ["25% - 99%", "+$5", "High collection. Big bump.", "$7 + $5 = $12"],
        ["100% PIF", "+$8", "Paid in full. Top - zero chargeback risk.", "$7 + $8 = $15"],
    ]
    add_table(s, Inches(0.5), Inches(1.3), Inches(12.3), Inches(2.8), kicker_data,
              header_fill=ACCENT_BLUE, col_widths=[Inches(2.0), Inches(2.5), Inches(5.3), Inches(2.5)],
              font_size=13, row_height_in=0.4)

    add_text(s, Inches(0.5), Inches(4.3), Inches(12.3), Inches(0.4),
             "Why we incentivize collected % (not just written premium):", font_size=18, bold=True, color=NAVY)
    why = [
        "Down payment + PIF money is the ONLY money we are SURE we will keep. Everything else can be charged back.",
        "Bigger down -> smaller monthly bills -> customer is less likely to cancel -> we keep more commission.",
        "PIF (paid in full) = no chargeback risk at all. PIF earns the biggest kicker AND a PIF dollar add-on.",
        "Some carriers pay commission as advance on written premium (chargeback if customer cancels). Some pay as-earned on collected. Either way - net commission ~= collected x rate. The kicker rewards retention through collection.",
    ]
    add_bullets(s, Inches(0.7), Inches(4.8), Inches(12.0), Inches(2.3), why, font_size=14)


def slide_safe_net(prs, idx, total):
    s = add_slide(prs)
    header_strip(s, "Safe Net - What It Means",
                 "The commission the agency EXPECTS to keep, after royalty and overhead.")
    footer(s, idx, total)

    add_text(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(0.4),
             "How commission actually works:", font_size=18, bold=True, color=NAVY)
    facts = [
        "Most carriers pay commission UPFRONT on the WRITTEN premium (advance). Some pay AS EARNED on collected.",
        "On advance carriers, if the customer stops paying, the carrier reverses the unearned commission (chargeback).",
        "Commission RATES vary 8-15% across carriers. We use 11% as a conservative blended rate.",
        "Long-run net commission ~= COLLECTED premium x commission rate. That's what the workbook uses.",
    ]
    add_bullets(s, Inches(0.7), Inches(1.7), Inches(12.0), Inches(2.2), facts, font_size=13)

    step_data = [
        ["Step", "Math", "Example: $1,500 policy, $150 down"],
        ["1. Written premium", "Carrier's number", "$1,500"],
        ["2. Commission advance", "Premium x 11% blended", "$165 advance"],
        ["3. Net commission expected", "Collected x 11%", "$150 x 11% = $16.50 expected to keep"],
        ["4. After royalty (21%)", "x 0.79", "$13.04"],
        ["5. After overhead reserve (30%)", "x 0.70", "$9.13 SAFE NET"],
        ["6. Bonus review threshold", "Safe net x 40%", "$3.65 max bonus on this policy"],
    ]
    add_table(s, Inches(0.5), Inches(4.0), Inches(12.3), Inches(3.1), step_data,
              header_fill=NAVY, col_widths=[Inches(3.0), Inches(3.5), Inches(5.8)],
              font_size=12, row_height_in=0.40)


def slide_safe_net_real(prs, idx, total):
    s = add_slide(prs)
    header_strip(s, "Safe Net - Real Example",
                 "Abel Guaina, January 2026")
    footer(s, idx, total)

    d = AGENTS['Abel Guaina']['January']
    total_written = d['NB'][1] + d['RWR'][1] + d['REN'][1]
    total_col = d['NB'][2] + d['RWR'][2] + d['REN'][2]
    advance = total_written * BLENDED_COMM
    expected_retained = total_col * BLENDED_COMM
    after_royalty = expected_retained * (1 - ROYALTY)
    safe_net = after_royalty * (1 - OVERHEAD)
    cap = safe_net * CAP_STANDARD

    abel_data = [
        ["Step", "Number", "What it means"],
        ["NB + RWR + REN written premium", f"${total_written:,.2f}", "What Abel sold and renewed in Jan"],
        ["Carrier commission advance (x 11%)", f"${advance:,.2f}", "Carriers paid us this much upfront"],
        ["Total collected from customers", f"${total_col:,.2f}", "Cash that came in"],
        ["Expected retained commission (x 11%)", f"${expected_retained:,.2f}", "After chargeback math"],
        ["After 17.5% royalty", f"${after_royalty:,.2f}", "After corporate cut"],
        ["SAFE NET (after 40% overhead reserve)", f"${safe_net:,.2f}", "Available for bonus + profit + taxes"],
        ["Review threshold (40% of safe net)", f"${cap:,.2f}", "Bonus targets above this get flagged"],
        ["New plan target (no min): $252", f"${252:.2f}", f"{252/safe_net*100:.0f}% of safe net - well under threshold"],
    ]
    add_table(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(4.2), abel_data,
              header_fill=NAVY, col_widths=[Inches(4.5), Inches(2.5), Inches(5.3)],
              font_size=12, row_height_in=0.40, first_col_bold=True)

    add_text(s, Inches(0.5), Inches(5.7), Inches(12.3), Inches(0.5),
             "Takeaway:", font_size=16, bold=True, color=NAVY)
    add_text(s, Inches(0.7), Inches(6.1), Inches(12.0), Inches(1.1),
             f"${safe_net:,.0f} of safe net means Abel's January contribution can support a bonus up to "
             f"~${cap:,.0f} before ownership wants a second look. Under the new plan, Abel's January target is "
             f"$252 - about {252/safe_net*100:.0f}% of safe net, well below the 40% review threshold.",
             font_size=13, color=BLACK)


def slide_why_current_drops(prs, idx, total):
    s = add_slide(prs)
    header_strip(s, "Why does CURRENT bonus DROP when RWR becomes REN?",
                 "Renewals pay MORE than rewrites - but only under the new plan.")
    footer(s, idx, total)

    add_text(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(0.5),
             "Short answer:", font_size=20, bold=True, color=NAVY)
    add_text(s, Inches(0.7), Inches(1.7), Inches(12.0), Inches(1.0),
             "The CURRENT plan does NOT pay anything for renewals. It only counts NB + RWR toward the 35-policy tier. "
             "When 50% of RWR moves to REN, the NB+RWR count drops, the tier drops, the bonus drops. "
             "Renewals 'paying more' is only true under the NEW plan, where the REN tier pays a separate bonus once retention rate hits 30%.",
             font_size=15, color=BLACK)

    add_text(s, Inches(0.5), Inches(3.0), Inches(12.3), Inches(0.5),
             "Concrete example - Dialinerys Dieguez, March:", font_size=16, bold=True, color=NAVY)

    real_d = AGENTS['Dialinerys Dieguez']['March']
    swap_d = swap_rwr_to_ren(real_d, 0.5)
    nb_r, rwr_r, ren_r = real_d['NB'][0], real_d['RWR'][0], real_d['REN'][0]
    nb_s, rwr_s, ren_s = swap_d['NB'][0], swap_d['RWR'][0], swap_d['REN'][0]
    cur_r = current_bonus(nb_r, rwr_r)
    cur_s = current_bonus(nb_s, rwr_s)
    rb = calc_proposal_a(real_d['NB'], real_d['RWR'], real_d['REN'])
    sb = calc_proposal_a(swap_d['NB'], swap_d['RWR'], swap_d['REN'])

    data = [
        ["Scenario", "NB", "RWR", "REN", "Current pays", "New plan pays"],
        [f"REAL", str(nb_r), str(rwr_r), str(ren_r), f"${cur_r}", f"${rb['paid']:.0f}"],
        [f"SWAP (50% RWR -> REN)", str(nb_s), str(rwr_s), str(ren_s), f"${cur_s}", f"${sb['paid']:.0f}"],
        ["CHANGE", "0", f"-{rwr_r-rwr_s}", f"+{ren_s-ren_r}", f"-${cur_r-cur_s} (BAD)", f"+${sb['paid']-rb['paid']:.0f} (GOOD)"],
    ]
    add_table(s, Inches(0.5), Inches(3.6), Inches(12.3), Inches(1.7), data,
              header_fill=NAVY, col_widths=[Inches(3.3), Inches(1.2), Inches(1.2), Inches(1.2), Inches(2.6), Inches(2.8)],
              font_size=13, row_height_in=0.42)

    add_text(s, Inches(0.5), Inches(5.5), Inches(12.3), Inches(0.5),
             "Bottom line:", font_size=16, bold=True, color=NAVY)
    items = [
        "Per swapped policy: Current loses ~$10 from the count tier. New plan gains by climbing the REN premium tiers as the renewal book grows.",
        "Under Current, the agency literally pays LESS when the agent does the right thing.",
        "Under the new plan, the agent gets MORE for retaining instead of rewriting - the exact incentive we want.",
    ]
    add_bullets(s, Inches(0.7), Inches(6.0), Inches(12.0), Inches(1.2), items, font_size=13)


def slide_calc_walkthrough(prs, idx, total):
    """Step-by-step worked example using Dialinerys March, all 3 scenarios."""
    s = add_slide(prs)
    header_strip(s, "Step-By-Step: How the Bonus Is Calculated",
                 "Worked example: Dialinerys Dieguez, March 2026. Same agent, same month, three scenarios.")
    footer(s, idx, total)

    d = AGENTS['Dialinerys Dieguez']['March']
    ds = swap_rwr_to_ren(d, 0.5)
    df = full_flip(d)

    scenarios = [
        ('REAL (today)', d, Inches(0.3), LIGHT_GRAY, NAVY),
        ('50% SWAP', ds, Inches(4.62), LIGHT_BLUE, ACCENT_BLUE),
        ('100% FLIP', df, Inches(8.94), LIGHT_GOLD, GOLD),
    ]
    for sc_name, scd, left, fill, hdr in scenarios:
        nb_c, nb_p, nb_col = scd['NB']
        rwr_c, rwr_p, rwr_col = scd['RWR']
        ren_c, ren_p, ren_col = scd['REN']
        a = calc_proposal_a(scd['NB'], scd['RWR'], scd['REN'])

        add_bar(s, left, Inches(1.1), Inches(4.05), Inches(0.4), hdr)
        add_text(s, left, Inches(1.1), Inches(4.05), Inches(0.4),
                 sc_name, font_size=14, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_bar(s, left, Inches(1.5), Inches(4.05), Inches(5.7), fill)
        y = Inches(1.55)
        line_h = Inches(0.215)
        lines = [
            ("Monthly volume:", True, hdr),
            (f"  NB: {nb_c} policies, ${nb_p/1000:.0f}k written", False, BLACK),
            (f"  RWR: {rwr_c} policies, ${rwr_p/1000:.0f}k written", False, BLACK),
            (f"  REN: {ren_c} policies, ${ren_p/1000:.0f}k written", False, BLACK),
            ("Hit a premium tier:", True, hdr),
            (f"  NB ${nb_p/1000:.0f}k -> {a['nb_tier_label']} = ${a['nb_target']:.0f}", False, NAVY),
            (f"  RWR ${rwr_p/1000:.0f}k -> {a['rwr_tier_label']} = ${a['rwr_target']:.0f}", False, NAVY),
            (f"  REN ${ren_p/1000:.0f}k -> {a['ren_tier_label']} = ${a['ren_target']:.0f}", False, NAVY),
            (f"TARGET TOTAL = ${a['total_target']:.0f}", True, hdr),
            ("Apply gates:", True, hdr),
            (f"  NB ${nb_p/1000:.0f}k vs $45k: {'PASS' if a['nb_qual'] else 'FAIL'}", False, GREEN if a['nb_qual'] else RED),
            (f"  REN retention 75% vs 30%: {'PASS' if a['ren_qual'] else 'FAIL'}", False, GREEN if a['ren_qual'] else RED),
            (f"  RWR (follows NB): {'PASS' if a['rwr_qual'] else 'FAIL'}", False, GREEN if a['rwr_qual'] else RED),
            (f"PAID = ${a['paid_after_min']:.0f}", True, hdr),
        ]
        for txt, b, c in lines:
            add_text(s, left + Inches(0.1), y, Inches(3.85), line_h, txt,
                     font_size=10, bold=b, color=c)
            y += line_h


def slide_agent_detail(prs, idx, total, agent, per_agent):
    """One slide per agent: 3 scenarios stacked, with policy counts AND written + collected premium."""
    s = add_slide(prs)
    header_strip(s, f"{agent} - Month by Month, All Three Scenarios",
                 f"All values are NEW PLAN WITH MIN. $W = written premium, $C = collected.")
    footer(s, idx, total)

    scenarios = [
        ('REAL (today)', lambda d: d, LIGHT_GRAY, NAVY),
        ('50% SWAP (RWR->REN)', lambda d: swap_rwr_to_ren(d, 0.5), LIGHT_BLUE, ACCENT_BLUE),
        ('100% FLIP (RWR<->REN)', full_flip, LIGHT_GOLD, GOLD),
    ]

    y = Inches(1.05)
    for sc_name, sfn, fill, hdr_color in scenarios:
        add_text(s, Inches(0.35), y, Inches(12.7), Inches(0.28),
                 sc_name, font_size=12, bold=True, color=hdr_color)
        rows = [["Mo", "NB#", "NB $W", "NB $C", "RWR#", "RWR $W", "RWR $C", "REN#", "REN $W", "REN $C", "New plan WITH MIN"]]
        cur_t = b_t = 0
        nb_p_t = nb_c_t = rwr_p_t = rwr_c_t = ren_p_t = ren_c_t = 0
        nb_n = rwr_n = ren_n = 0
        for month in MONTHS:
            d = sfn(AGENTS[agent][month])
            cur = current_bonus(d['NB'][0], d['RWR'][0])
            b = calc_proposal_a(d['NB'], d['RWR'], d['REN'])
            rows.append([month[:3], str(d['NB'][0]), f"${d['NB'][1]/1000:.0f}k", f"${d['NB'][2]/1000:.1f}k",
                         str(d['RWR'][0]), f"${d['RWR'][1]/1000:.0f}k", f"${d['RWR'][2]/1000:.1f}k",
                         str(d['REN'][0]), f"${d['REN'][1]/1000:.0f}k", f"${d['REN'][2]/1000:.1f}k",
                         f"${b['paid_after_min']:.0f}"])
            cur_t += cur; b_t += b['paid_after_min']
            nb_n += d['NB'][0]; nb_p_t += d['NB'][1]; nb_c_t += d['NB'][2]
            rwr_n += d['RWR'][0]; rwr_p_t += d['RWR'][1]; rwr_c_t += d['RWR'][2]
            ren_n += d['REN'][0]; ren_p_t += d['REN'][1]; ren_c_t += d['REN'][2]
        rows.append(["TOT", str(nb_n), f"${nb_p_t/1000:.0f}k", f"${nb_c_t/1000:.0f}k",
                     str(rwr_n), f"${rwr_p_t/1000:.0f}k", f"${rwr_c_t/1000:.0f}k",
                     str(ren_n), f"${ren_p_t/1000:.0f}k", f"${ren_c_t/1000:.0f}k",
                     f"${b_t:.0f}"])
        add_table(s, Inches(0.35), y + Inches(0.3), Inches(12.6), Inches(1.55), rows,
                  header_fill=hdr_color,
                  col_widths=[Inches(0.55), Inches(0.55), Inches(0.85), Inches(0.85),
                              Inches(0.55), Inches(0.85), Inches(0.85),
                              Inches(0.55), Inches(0.85), Inches(0.85),
                              Inches(1.3)],
                  font_size=10, row_height_in=0.255, highlight_col=10, highlight_color=fill,
                  first_col_bold=True)
        y += Inches(1.95)

    # Bottom read
    p = per_agent[agent]
    add_text(s, Inches(0.35), Inches(7.0), Inches(12.7), Inches(0.23),
             f"Today's current plan pays ${p['cur_r']:,.0f}. New plan WITH MIN: REAL ${p['b_min_r']:,.0f} | SWAP ${p['b_min_s']:,.0f} | FLIP ${p['b_min_f']:,.0f}.",
             font_size=11, bold=True, italic=True, color=NAVY, align=PP_ALIGN.CENTER)


def _tier_text(tier_num, tier_label, amount):
    if tier_num == 0:
        return "Below T1 (min not met)"
    return f"T{tier_num} {tier_label} -> ${amount:,.0f}"


def _scenario_premium_rows(scenario_fn, line_type, tier_fn):
    """Return up to 12 top agent-month rows for a given scenario+line, sorted by premium desc.
    Each row = [Agent, Month, Premium, Tier (text), Bonus]."""
    rows = []
    for agent in AGENT_NAMES:
        for month in MONTHS:
            d = scenario_fn(AGENTS[agent][month])
            prem = d[line_type][1]
            bonus, tier_num, tier_label = tier_fn(prem)
            rows.append((agent, month, prem, tier_num, tier_label, bonus))
    rows.sort(key=lambda x: -x[2])
    return rows[:12]


def slide_examples_single_line(prs, idx, total, line_type, tier_fn, ladder_text):
    """One slide showing the new tier hits for a single bonus line (NB / RWR / REN)
    under all 3 scenarios side by side. Top 12 rows per scenario."""
    s = add_slide(prs)
    header_strip(s, f"{line_type} Bonus Examples - Who Hits Which Tier",
                 ladder_text)
    footer(s, idx, total)

    scenarios = [
        ('REAL (today)', lambda d: d, LIGHT_GRAY, NAVY),
        ('50% SWAP', lambda d: swap_rwr_to_ren(d, 0.5), LIGHT_BLUE, ACCENT_BLUE),
        ('100% FLIP', full_flip, LIGHT_GOLD, GOLD),
    ]

    x_positions = [Inches(0.3), Inches(4.62), Inches(8.94)]
    for (sc_name, sc_fn, fill, hdr), x in zip(scenarios, x_positions):
        rows_data = _scenario_premium_rows(sc_fn, line_type, tier_fn)

        # Scenario header bar
        add_bar(s, x, Inches(1.1), Inches(4.05), Inches(0.4), hdr)
        add_text(s, x, Inches(1.1), Inches(4.05), Inches(0.4),
                 sc_name, font_size=14, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Table
        table_rows = [['Agent', f'{line_type} Premium', 'Tier / Bonus']]
        for agent, month, prem, tier_num, tier_label, bonus in rows_data:
            first = agent.split()[0]
            tier_str = _tier_text(tier_num, tier_label, bonus)
            table_rows.append([f"{first} {month[:3]}", f"${prem/1000:,.0f}k", tier_str])

        add_table(s, x, Inches(1.55), Inches(4.05), Inches(5.7), table_rows,
                  header_fill=hdr,
                  col_widths=[Inches(1.3), Inches(0.95), Inches(1.8)],
                  font_size=9, row_height_in=0.30, first_col_bold=True)


def slide_examples_combined(prs, idx, total):
    """Combined NB + RWR + REN with Total Bonus per row, three scenarios."""
    s = add_slide(prs)
    header_strip(s, "Combined Bonus Examples - All Three Lines + Total Bonus",
                 "Per agent-month: NB tier + RWR tier + REN tier = TOTAL BONUS. Top 8 by total bonus per scenario.")
    footer(s, idx, total)

    scenarios = [
        ('REAL (today)', lambda d: d, NAVY),
        ('50% SWAP', lambda d: swap_rwr_to_ren(d, 0.5), ACCENT_BLUE),
        ('100% FLIP', full_flip, GOLD),
    ]

    x_positions = [Inches(0.3), Inches(4.62), Inches(8.94)]
    for (sc_name, sc_fn, hdr), x in zip(scenarios, x_positions):
        rows_data = []
        for agent in AGENT_NAMES:
            for month in MONTHS:
                d = sc_fn(AGENTS[agent][month])
                nb_b, _, _ = nb_tier_bonus(d['NB'][1])
                rwr_b, _, _ = rwr_tier_bonus(d['RWR'][1])
                ren_b, _, _ = ren_tier_bonus(d['REN'][1])
                nb_qual = d['NB'][1] >= NB_MIN_PREMIUM
                nb_paid = nb_b if nb_qual else 0
                rwr_paid = rwr_b if nb_qual else 0
                ren_paid = ren_b
                total_b = nb_paid + rwr_paid + ren_paid
                rows_data.append((agent, month, nb_paid, rwr_paid, ren_paid, total_b))
        rows_data.sort(key=lambda x: -x[5])
        rows_data = rows_data[:10]
        total_4mo = sum(x[5] for x in rows_data)

        add_bar(s, x, Inches(1.1), Inches(4.05), Inches(0.4), hdr)
        add_text(s, x, Inches(1.1), Inches(4.05), Inches(0.4),
                 sc_name, font_size=14, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        table_rows = [['Agent / Mo', 'NB', 'RWR', 'REN', 'TOTAL']]
        for agent, month, nb_b, rwr_b, ren_b, total_b in rows_data:
            first = agent.split()[0]
            table_rows.append([f"{first} {month[:3]}",
                               f"${nb_b:,.0f}" if nb_b else "-",
                               f"${rwr_b:,.0f}" if rwr_b else "-",
                               f"${ren_b:,.0f}" if ren_b else "-",
                               f"${total_b:,.0f}"])
        table_rows.append(['Top-10 total (this scenario)', '', '', '',
                           f"${total_4mo:,.0f}"])

        add_table(s, x, Inches(1.55), Inches(4.05), Inches(5.5), table_rows,
                  header_fill=hdr,
                  col_widths=[Inches(1.3), Inches(0.65), Inches(0.65), Inches(0.65), Inches(0.8)],
                  font_size=9, row_height_in=0.28, first_col_bold=True)


def slide_current_vs_new(prs, idx, total):
    """Per agent-month: today's CURRENT bonus vs the new plan total, sorted
    by delta descending. Three scenarios side by side, top 10 per scenario."""
    s = add_slide(prs)
    header_strip(s, "Current vs New Plan - Per Agent-Month Comparison",
                 "CURRENT = today's count-tier on NB+RWR. NEW = tier-bonus (NB + RWR + REN). Sorted by delta. Top 10 per scenario.")
    footer(s, idx, total)

    scenarios = [
        ('REAL (today)', lambda d: d, NAVY),
        ('50% SWAP', lambda d: swap_rwr_to_ren(d, 0.5), ACCENT_BLUE),
        ('100% FLIP', full_flip, GOLD),
    ]
    x_positions = [Inches(0.3), Inches(4.62), Inches(8.94)]
    for (sc_name, sc_fn, hdr), x in zip(scenarios, x_positions):
        rows_data = []
        scen_total_cur = scen_total_new = 0
        for agent in AGENT_NAMES:
            for month in MONTHS:
                d = sc_fn(AGENTS[agent][month])
                cur = current_bonus(d['NB'][0], d['RWR'][0])
                nb_b, _, _ = nb_tier_bonus(d['NB'][1])
                rwr_b, _, _ = rwr_tier_bonus(d['RWR'][1])
                ren_b, _, _ = ren_tier_bonus(d['REN'][1])
                nb_qual = d['NB'][1] >= NB_MIN_PREMIUM
                new_total = (nb_b if nb_qual else 0) + (rwr_b if nb_qual else 0) + ren_b
                delta = new_total - cur
                scen_total_cur += cur
                scen_total_new += new_total
                rows_data.append((agent, month, cur, new_total, delta))
        rows_data.sort(key=lambda x: -x[4])
        rows_data = rows_data[:10]

        add_bar(s, x, Inches(1.1), Inches(4.05), Inches(0.4), hdr)
        add_text(s, x, Inches(1.1), Inches(4.05), Inches(0.4),
                 sc_name, font_size=14, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        table_rows = [['Agent / Mo', 'Current', 'New', 'Delta']]
        for agent, month, cur, new_total, delta in rows_data:
            first = agent.split()[0]
            sign = '+' if delta > 0 else ''
            table_rows.append([f"{first} {month[:3]}",
                               f"${cur:,.0f}",
                               f"${new_total:,.0f}",
                               f"{sign}${delta:,.0f}"])
        table_rows.append(['4-mo total (all agents)',
                           f"${scen_total_cur:,.0f}",
                           f"${scen_total_new:,.0f}",
                           f"{'+' if scen_total_new-scen_total_cur >= 0 else ''}${scen_total_new-scen_total_cur:,.0f}"])

        add_table(s, x, Inches(1.55), Inches(4.05), Inches(5.5), table_rows,
                  header_fill=hdr,
                  col_widths=[Inches(1.4), Inches(0.85), Inches(0.85), Inches(0.95)],
                  font_size=9, row_height_in=0.28, first_col_bold=True)


def slide_swap_comparison(prs, idx, total, real, swap, flip):
    s = add_slide(prs)
    header_strip(s, "Headline - 4-Month Totals Across 3 Scenarios (6 agents)",
                 "REAL = today. 50% SWAP = half of rewrites become renewals. 100% FLIP = renewals & rewrites fully swapped.")
    footer(s, idx, total)

    data = [
        ["Plan", "REAL (today)", "50% SWAP", "100% FLIP", "Read"],
        ["Current plan",  f"${real['cur']:,.0f}",  f"${swap['cur']:,.0f}",  f"${flip['cur']:,.0f}",
         "CRASHES - current pays $0 for REN"],
        ["New plan - no minimum", f"${real['a']:,.0f}",   f"${swap['a']:,.0f}",   f"${flip['a']:,.0f}",
         "Calculated target before applying minimums"],
        ["New plan WITH MIN (RECOMMENDED)", f"${real['a_min']:,.0f}", f"${swap['a_min']:,.0f}", f"${flip['a_min']:,.0f}",
         "Pays when NB >= $45k/mo and retention >= 30%"],
    ]
    add_table(s, Inches(0.4), Inches(1.2), Inches(12.5), Inches(2.4), data,
              header_fill=NAVY,
              col_widths=[Inches(3.4), Inches(1.6), Inches(1.6), Inches(1.6), Inches(4.3)],
              font_size=13, row_height_in=0.55, first_col_bold=True)

    add_text(s, Inches(0.5), Inches(4.0), Inches(12.3), Inches(0.5),
             "The headline:", font_size=18, bold=True, color=NAVY)
    items = [
        f"Current plan COLLAPSES under the flip (${flip['cur']:,.0f} vs today's ${real['cur']:,.0f}). It only rewards rewrite VOLUME.",
        f"New plan WITH MIN under 100% FLIP = ${flip['a_min']:,.0f} ({flip['a_min']/real['cur']*100:.0f}% of today's $12,220). The agency keeps more profit while still rewarding the right behavior.",
        f"New plan rewards what the current plan ignores: writing bigger NB books AND keeping renewals - each has its own premium tier ladder.",
        f"The pitch to the agent: 'today you're earning ${real['cur']/6/4:.0f}/mo on rewrites. Under the new plan, when you renew instead, the REN tier kicks in - up to $800 at $100k retained.'",
    ]
    add_bullets(s, Inches(0.7), Inches(4.5), Inches(12.0), Inches(2.6), items, font_size=13)


def slide_profitability(prs, idx, total):
    s = add_slide(prs)
    header_strip(s, "How We Protect Profit",
                 "Three layers. 90-day chargeback is the strongest.")
    footer(s, idx, total)

    cards = [
        ("MINIMUM REQUIREMENTS", LIGHT_GREEN, GREEN,
         f"NB premium ≥ ${NB_MIN_PREMIUM:,}/mo + REN retention ≥ {REN_MIN_RETENTION*100:.0f}%. RWR follows NB gate.",
         "If the agent is not producing baseline volume, NO bonus that month. Easy to understand, easy to enforce."),
        ("REVIEW THRESHOLD", LIGHT_GOLD, GOLD,
         "If monthly bonus > 40% of safe net, ownership reviews.",
         "Soft flag, not auto-cut. Catches outlier months where collection is unusually low."),
        ("90-DAY CHARGEBACK", LIGHT_ORANGE, ORANGE,
         "100% of the paid bonus is REVERSED if the policy cancels or rewrites within 90 days.",
         "This is the real profit shield. Matches the carrier's commission chargeback exposure exactly. If the agency loses commission, the agent loses bonus."),
    ]

    for i, (title, lt, dk, rule, desc) in enumerate(cards):
        left = Inches(0.5 + i * 4.28)
        add_bar(s, left, Inches(1.3), Inches(4.1), Inches(5.7), lt)
        add_bar(s, left, Inches(1.3), Inches(4.1), Inches(0.8), dk)
        add_text(s, left + Inches(0.1), Inches(1.35), Inches(3.9), Inches(0.7),
                 title, font_size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, left + Inches(0.2), Inches(2.3), Inches(3.7), Inches(1.4),
                 rule, font_size=14, bold=True, color=NAVY)
        add_text(s, left + Inches(0.2), Inches(3.8), Inches(3.7), Inches(3.0),
                 desc, font_size=13, color=BLACK)


def slide_recommendation(prs, idx, total):
    s = add_slide(prs)
    header_strip(s, "Recommendation",
                 f"The Bonus Plan with ${NB_MIN_PREMIUM:,} NB + {REN_MIN_RETENTION*100:.0f}% retention minimums. Pilot 90 days, then review.")
    footer(s, idx, total)

    add_bar(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(1.5), LIGHT_GOLD)
    add_bar(s, Inches(0.5), Inches(1.2), Inches(0.25), Inches(1.5), GOLD)
    add_text(s, Inches(1.0), Inches(1.4), Inches(11.3), Inches(0.5),
             f"THE NEW BONUS PLAN - with ${NB_MIN_PREMIUM:,} NB premium + {REN_MIN_RETENTION*100:.0f}% retention minimums",
             font_size=22, bold=True, color=NAVY)
    add_text(s, Inches(1.0), Inches(1.9), Inches(11.3), Inches(0.75),
             "Balanced for agent motivation, retention upside, profit protection, and clean payroll math.",
             font_size=14, color=BLACK)

    add_text(s, Inches(0.5), Inches(3.0), Inches(12.3), Inches(0.5),
             "Why the new plan beats today's:", font_size=18, bold=True, color=NAVY)
    items = [
        "Same structure as today (count tiers -> $X) but driven by PREMIUM written, not policy count.",
        "Pays for RENEWALS (today's plan pays $0) - REN gets its own tier ladder starting at $25k.",
        "Bigger jumps at the top tiers - $1,000 NB at $100k premium, with linear 1% upside above.",
        "Costs LESS at full target than today's count-tier plan - new plan FLIP $7,975 vs current $12,220.",
    ]
    add_bullets(s, Inches(0.7), Inches(3.5), Inches(12.0), Inches(2.0), items, font_size=14)

    add_text(s, Inches(0.5), Inches(5.7), Inches(12.3), Inches(0.5),
             "Levers to fine-tune over time:", font_size=18, bold=True, color=NAVY)
    add_text(s, Inches(0.7), Inches(6.2), Inches(12.0), Inches(1.0),
             "1) Tier amounts (lower entry / higher top to motivate harder).   "
             "2) Tier breakpoints ($45k entry / $100k top - tighten or loosen the ladder).   "
             "3) Retention rate gate (30% today - dial up over time as the renewal book builds).",
             font_size=13, color=DARK_GRAY)


def slide_roadmap(prs, idx, total):
    s = add_slide(prs)
    header_strip(s, "Implementation Roadmap",
                 "Three phases. ~10 weeks from approval to first chargeback review.")
    footer(s, idx, total)

    phases = [
        ("WEEK 1-2", "APPROVE + COMMUNICATE", ACCENT_BLUE, LIGHT_BLUE, [
            "Ownership signs off on the new plan rules and minimum thresholds.",
            "Manager rolls out the rules to the agents.",
            "Print the Manual Tracker (in the workbook) for each desk.",
        ]),
        ("WEEK 3-6", "PILOT MONTH 1", GREEN, LIGHT_GREEN, [
            "Run the plan side-by-side with the OLD plan for ONE FULL MONTH.",
            "Pay agents the HIGHER of the two each month during the pilot.",
            "Use the manual tracker to verify coverage / PIF / collected.",
            "End of month: compare actual paid vs forecast and adjust if needed.",
        ]),
        ("WEEK 7-10+", "FULL ROLLOUT + REVIEWS", GOLD, LIGHT_GOLD, [
            "Switch to NEW plan exclusively.",
            "Run monthly review at the 40% safe-net flag.",
            "At week 12 (90 days after first paid policies), run the FIRST CHARGEBACK review.",
            "At 90 days end: re-baseline rates if the renewal book has started to build.",
        ]),
    ]
    y = Inches(1.2)
    for label, title, dk, lt, items in phases:
        add_bar(s, Inches(0.5), y, Inches(12.3), Inches(1.85), lt)
        add_bar(s, Inches(0.5), y, Inches(0.25), Inches(1.85), dk)
        add_text(s, Inches(1.0), y + Inches(0.1), Inches(2.5), Inches(0.4),
                 label, font_size=14, bold=True, color=dk)
        add_text(s, Inches(3.7), y + Inches(0.1), Inches(9.0), Inches(0.4),
                 title, font_size=18, bold=True, color=NAVY)
        add_bullets(s, Inches(1.0), y + Inches(0.55), Inches(11.5), Inches(1.25),
                    items, font_size=12, color=BLACK, line_spacing=1.0)
        y += Inches(1.95)


def slide_closing(prs):
    s = add_slide(prs)
    add_bar(s, 0, 0, Inches(13.33), Inches(7.5), NAVY)
    add_bar(s, 0, Inches(3.4), Inches(13.33), Inches(0.1), GOLD)
    add_text(s, Inches(0.6), Inches(1.0), Inches(12.13), Inches(1.0),
             "Decision Time", font_size=54, bold=True, color=WHITE)
    add_text(s, Inches(0.6), Inches(2.0), Inches(12.13), Inches(0.8),
             "The new bonus plan with NB & retention minimums.", font_size=24, color=LIGHT_BLUE)
    add_text(s, Inches(0.6), Inches(3.8), Inches(12.13), Inches(0.5),
             "Workbook companion file:", font_size=18, color=WHITE)
    add_text(s, Inches(0.6), Inches(4.3), Inches(12.13), Inches(0.5),
             "Bonus_Plan_FINAL.xlsx",
             font_size=20, bold=True, color=LIGHT_BLUE)
    add_text(s, Inches(0.6), Inches(5.2), Inches(12.13), Inches(0.5),
             "One plan. One decision. Ready to roll out in 10 weeks.",
             font_size=18, italic=True, color=LIGHT_BLUE)


# MAIN BUILD
def build():
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    real, swap, flip, per_agent = compute_totals()

    # Simplified deck. Per-agent slides dropped in favor of tier-examples slides
    # that mirror the Excel "Bonus Examples" sheets.
    nb_ladder = "T1 $45k=$250 | T2 $55k=$375 | T3 $70k=$525 | T4 $85k=$725 | T5 $100k=$1,000 | Above $100k: +$50 per $5k. Min req $45k."
    rwr_ladder = "T1 $45k=$100 | T2 $55k=$200 | T3 $70k=$275 | T4 $85k=$350 | T5 $100k=$500 | Above $100k: +$25 per $5k. Gated by NB min."
    ren_ladder = "T1 $25k=$250 | T2 $40k=$350 | T3 $65k=$450 | T4 $80k=$550 | T5 $100k=$800 | Above $100k: +$40 per $5k. Min 30% retention rate."

    builders = [
        lambda t: slide_cover(prs),
        lambda t: slide_problem(prs, 2, t),
        lambda t: slide_plan_a_details_new(prs, 3, t),
        lambda t: slide_minimums(prs, 4, t),
        lambda t: slide_why_current_drops(prs, 5, t),
        lambda t: slide_calc_walkthrough(prs, 6, t),
        lambda t: slide_examples_single_line(prs, 7, t, 'NB',  nb_tier_bonus,  nb_ladder),
        lambda t: slide_examples_single_line(prs, 8, t, 'RWR', rwr_tier_bonus, rwr_ladder),
        lambda t: slide_examples_single_line(prs, 9, t, 'REN', ren_tier_bonus, ren_ladder),
        lambda t: slide_examples_combined(prs, 10, t),
        lambda t: slide_current_vs_new(prs, 11, t),
        lambda t: slide_swap_comparison(prs, 12, t, real, swap, flip),
        lambda t: slide_recommendation(prs, 13, t),
        lambda t: slide_closing(prs),
    ]
    total = len(builders)
    for fn in builders:
        fn(total)

    out = '/home/user/fiesta-bonus/output/Bonus_Plan_SIMPLE.pptx'
    prs.save(out)
    print(f"Saved: {out} ({total} slides)")
    return out


if __name__ == '__main__':
    build()
