"""
Builds the comprehensive bonus proposal workbook.

Two separated proposals (no mixing):
  - Proposal A: PREMIUM-BASED (NB/REN tiered by written premium + collected kicker)
  - Proposal B: COVERAGE-BASED (flat $10 NB base + $5 BI+UM liability bundle, no premium tiers)
  - Proposal C (bonus idea): PERSISTENCY-FOCUSED (flat per-policy, REN >= NB to actively reward retention)

Examples are run month-by-month for 4 months (Jan-Apr 2026) for 6 agents:
  Abel Guaina, Dialinerys Dieguez, Melissa Hernandez, Flavia Parra, Thalia Rojas, Monica Valdes

Two scenarios for each proposal:
  - REAL data (as written)
  - 50% RWR -> REN swap (showing the incentive direction)
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from copy import copy

# ============================================================================
# AGENT DATA (from agents_nbrwr_pivots_good_.xlsx, source of truth)
# Format: (count, written_premium, collected_premium)
# ============================================================================
AGENTS = {
    'Abel Guaina': {
        'January':  {'NB': (25, 38611.42, 4938.26),  'RWR': (17, 27518.50, 6186.16),  'REN': (9, 7605.00, 1734.23)},
        'February': {'NB': (13, 25367.00, 3258.77),  'RWR': (19, 30643.00, 4268.29),  'REN': (7, 7517.00, 1422.16)},
        'March':    {'NB': (30, 46523.37, 6523.33),  'RWR': (17, 25227.60, 5844.85),  'REN': (5, 3901.00, 1274.08)},
        'April':    {'NB': (27, 37825.93, 8163.72),  'RWR': (15, 15361.95, 3967.81),  'REN': (13, 10935.99, 5439.19)},
    },
    'Dialinerys Dieguez': {
        'January':  {'NB': (23, 40249.26, 4925.00),  'RWR': (48, 65675.38, 10936.32), 'REN': (11, 12530.47, 4951.59)},
        'February': {'NB': (23, 30057.89, 5766.72),  'RWR': (56, 85544.60, 13931.43), 'REN': (9, 11629.68, 1937.60)},
        'March':    {'NB': (38, 52434.90, 10546.59), 'RWR': (69, 96228.35, 19700.38), 'REN': (7, 5962.50, 3080.00)},
        'April':    {'NB': (33, 49689.35, 8831.44),  'RWR': (44, 64165.43, 16668.05), 'REN': (2, 19575.00, 2939.65)},
    },
    'Melissa Hernandez': {
        'January':  {'NB': (31, 37039.00, 7651.14),  'RWR': (48, 63374.50, 11897.57), 'REN': (0, 0,        0)},
        'February': {'NB': (13, 7659.50,  2095.02),  'RWR': (37, 49990.70, 6802.73),  'REN': (1, 1063.00,  91.75)},
        'March':    {'NB': (28, 39759.70, 5669.41),  'RWR': (48, 56860.60, 8620.47),  'REN': (1, 1034.00,  180.45)},
        'April':    {'NB': (15, 19288.35, 4344.30),  'RWR': (37, 43396.20, 7660.70),  'REN': (0, 0,        0)},
    },
    'Flavia Parra': {
        'January':  {'NB': (19, 28525.05, 5657.40),  'RWR': (35, 39558.05, 8795.29),  'REN': (1, 1464.35,  302.12)},
        'February': {'NB': (23, 22247.70, 3832.59),  'RWR': (42, 43687.45, 13820.09), 'REN': (2, 2099.00,  321.01)},
        'March':    {'NB': (27, 31269.00, 8642.02),  'RWR': (27, 23526.70, 9060.63),  'REN': (0, 0,        0)},
        'April':    {'NB': (23, 29581.05, 6600.51),  'RWR': (31, 30832.76, 7710.62),  'REN': (1, 885.00,   885.00)},
    },
    'Thalia Rojas': {
        'January':  {'NB': (10, 15503.90, 2238.28),  'RWR': (28, 33083.93, 6461.48),  'REN': (4, 3690.00,  642.04)},
        'February': {'NB': (14, 19414.25, 3640.91),  'RWR': (24, 28298.16, 5276.12),  'REN': (5, 7446.00,  3293.14)},
        'March':    {'NB': (19, 17843.69, 4965.76),  'RWR': (19, 19840.85, 3981.64),  'REN': (6, 5643.90,  1575.86)},
        'April':    {'NB': (9,  14099.00, 3941.94),  'RWR': (29, 39794.40, 7674.90),  'REN': (3, 3617.00,  3096.58)},
    },
    'Monica Valdes': {
        'January':  {'NB': (23, 28605.29, 6551.70),  'RWR': (37, 48612.96, 8203.88),  'REN': (6, 6502.91,  1439.30)},
        'February': {'NB': (21, 31516.95, 6512.87),  'RWR': (29, 47228.18, 7155.23),  'REN': (2, 3092.00,  272.87)},
        'March':    {'NB': (36, 68546.22, 17019.79), 'RWR': (28, 39254.65, 5300.53),  'REN': (0, 0,        0)},
        'April':    {'NB': (22, 29755.49, 6197.26),  'RWR': (22, 35279.90, 7305.79),  'REN': (2, 2514.00,  291.12)},
    },
}

MONTHS = ['January', 'February', 'March', 'April']
AGENT_NAMES = list(AGENTS.keys())

# ============================================================================
# ECONOMIC ASSUMPTIONS
# ============================================================================
BLENDED_COMM = 0.11      # blended carrier commission rate (8-15% across the book)
ROYALTY = 0.175          # franchise royalty
OVERHEAD = 0.40          # overhead reserve for rent/payroll/tech/etc
CAP_STANDARD = 0.40      # ownership REVIEW threshold (not auto-cap): if target > 40% safe net, flag for review
CAP_PIF = 0.55           # PIF review threshold

# Estimated coverage adoption for Proposal B (since coverage data is not in source pivots)
LIAB_ADOPTION_NB = 0.35  # 35% of NB have BI+UM bundle (manager verifies)
LIAB_ADOPTION_REN = 0.30 # 30% of REN have BI+UM bundle (already in book)

# NEW: Minimum-requirement thresholds (monthly, by written premium)
NB_MIN_PREMIUM = 50000   # Must write at least $50,000 NB premium in the month to earn the NB bonus
REN_MIN_PREMIUM = 50000  # Must keep at least $50,000 REN premium in the month to earn the REN bonus
# RWR bonus is paid ONLY if BOTH NB and REN minimums are met. RWR has no own threshold.

# ============================================================================
# BONUS ENGINES
# ============================================================================
def kicker(collected_pct):
    """Collected % kicker - same for both proposals."""
    if collected_pct >= 1.00: return 1.25  # 100% (PIF)
    if collected_pct >= 0.50: return 1.20  # 50-99%
    if collected_pct >= 0.25: return 1.15  # 25-49%
    if collected_pct >= 0.15: return 1.10  # 15-24%
    return 1.00                              # below 15%, no kicker


def apply_minimums(nb_premium, ren_premium, nb_target, ren_target, rwr_target,
                   nb_min=NB_MIN_PREMIUM, ren_min=REN_MIN_PREMIUM):
    """Gate the bonus lines by the minimum-premium thresholds.

    Rules (per the agent's request):
      - NB line pays only if NB premium >= NB_MIN_PREMIUM that month.
      - REN line pays only if REN premium >= REN_MIN_PREMIUM that month.
      - RWR line pays only if BOTH NB and REN minimums are met.
    """
    nb_qual = nb_premium >= nb_min
    ren_qual = ren_premium >= ren_min
    rwr_qual = nb_qual and ren_qual
    paid_nb = nb_target if nb_qual else 0.0
    paid_ren = ren_target if ren_qual else 0.0
    paid_rwr = rwr_target if rwr_qual else 0.0
    return {
        'nb_qual': nb_qual, 'ren_qual': ren_qual, 'rwr_qual': rwr_qual,
        'paid_nb_after_min': paid_nb,
        'paid_ren_after_min': paid_ren,
        'paid_rwr_after_min': paid_rwr,
        'paid_after_min': paid_nb + paid_ren + paid_rwr,
    }


def kicker_label(collected_pct):
    if collected_pct >= 1.00: return "PIF (+25%)"
    if collected_pct >= 0.50: return "50%+ (+20%)"
    if collected_pct >= 0.25: return "25-49% (+15%)"
    if collected_pct >= 0.15: return "15-24% (+10%)"
    return "<15% (none)"


def nb_premium_tier(avg_prem):
    """Proposal A NB tier."""
    if avg_prem < 1200: return 5,  "Under $1,200"
    if avg_prem < 1800: return 7,  "$1,200-$1,799"
    if avg_prem < 2200: return 9,  "$1,800-$2,199"
    if avg_prem < 3000: return 11, "$2,200-$2,999"
    # $3,000+: $11 + $2 per $1k, cap $25
    val = min(11 + (avg_prem - 3000) / 1000 * 2, 25)
    return val, "$3,000+"


def ren_premium_tier(avg_prem):
    """Proposal A REN tier."""
    if avg_prem < 1200: return 4, "Under $1,200"
    if avg_prem < 1800: return 5, "$1,200-$1,799"
    return 6, "$1,800+"


def calc_proposal_a(nb, rwr, ren):
    """Proposal A: PREMIUM-BASED with collected kicker.
       Returns BOTH the uncapped target AND the cap-limited paid amount.
       In practice, pay the target. The cap is an ownership safety check at month-end."""
    nb_c, nb_p, nb_col = nb
    rwr_c, rwr_p, rwr_col = rwr
    ren_c, ren_p, ren_col = ren

    avg_nb = nb_p / nb_c if nb_c else 0
    avg_ren = ren_p / ren_c if ren_c else 0

    nb_per, nb_tier_label = nb_premium_tier(avg_nb)
    ren_per, ren_tier_label = ren_premium_tier(avg_ren)

    nb_col_pct = nb_col / nb_p if nb_p else 0
    ren_col_pct = ren_col / ren_p if ren_p else 0
    rwr_col_pct = rwr_col / rwr_p if rwr_p else 0

    nb_target = nb_c * nb_per * kicker(nb_col_pct)
    ren_target = ren_c * ren_per * kicker(ren_col_pct)
    rwr_target = rwr_c * 2 * kicker(rwr_col_pct)
    total_target = nb_target + ren_target + rwr_target

    total_collected = nb_col + rwr_col + ren_col
    gross_comm = total_collected * BLENDED_COMM
    safe_net = gross_comm * (1 - ROYALTY) * (1 - OVERHEAD)
    cap = safe_net * CAP_STANDARD
    paid = total_target  # Pay the calculated target. Cap is a review threshold, not auto-cut.
    capped_paid = min(total_target, cap)  # What it would be if 30% cap was enforced
    bonus_pct_of_safe_net = paid / safe_net if safe_net else 0
    review_flag = (paid > cap)  # Soft flag: ownership reviews months that exceed 40% safe net

    mins = apply_minimums(nb_p, ren_p, nb_target, ren_target, rwr_target)

    return {
        'avg_nb': avg_nb, 'avg_ren': avg_ren,
        'nb_per': nb_per, 'nb_tier_label': nb_tier_label,
        'ren_per': ren_per, 'ren_tier_label': ren_tier_label,
        'nb_target': nb_target, 'ren_target': ren_target, 'rwr_target': rwr_target,
        'total_target': total_target,
        'nb_col_pct': nb_col_pct, 'ren_col_pct': ren_col_pct, 'rwr_col_pct': rwr_col_pct,
        'gross_comm': gross_comm, 'safe_net': safe_net, 'cap': cap,
        'paid': paid, 'capped_paid': capped_paid,
        'bonus_pct_of_safe_net': bonus_pct_of_safe_net, 'review_flag': review_flag,
        **mins,
    }


def calc_proposal_b(nb, rwr, ren):
    """Proposal B: COVERAGE-BASED. No premium tiers.
       NB base = $10/policy (FL minimum or full coverage).
       NB liability add = +$5/policy IF BI+UM bundle present (UM cannot exist without BI).
       REN base = $6/policy. REN liability add = +$3/policy.
       RWR = $2 flat.
       Estimated adoption: 35% of NB / 30% of REN carry BI+UM.
    """
    nb_c, nb_p, nb_col = nb
    rwr_c, rwr_p, rwr_col = rwr
    ren_c, ren_p, ren_col = ren

    nb_col_pct = nb_col / nb_p if nb_p else 0
    ren_col_pct = ren_col / ren_p if ren_p else 0
    rwr_col_pct = rwr_col / rwr_p if rwr_p else 0

    nb_base = nb_c * 10
    nb_liability = nb_c * LIAB_ADOPTION_NB * 5
    nb_target = (nb_base + nb_liability) * kicker(nb_col_pct)

    ren_base = ren_c * 6
    ren_liability = ren_c * LIAB_ADOPTION_REN * 3
    ren_target = (ren_base + ren_liability) * kicker(ren_col_pct)

    rwr_target = rwr_c * 2 * kicker(rwr_col_pct)

    total_target = nb_target + ren_target + rwr_target

    total_collected = nb_col + rwr_col + ren_col
    gross_comm = total_collected * BLENDED_COMM
    safe_net = gross_comm * (1 - ROYALTY) * (1 - OVERHEAD)
    cap = safe_net * CAP_STANDARD
    paid = total_target  # pay the target; cap is a review threshold, not auto-cut
    capped_paid = min(total_target, cap)
    bonus_pct_of_safe_net = paid / safe_net if safe_net else 0
    review_flag = (paid > cap)

    mins = apply_minimums(nb_p, ren_p, nb_target, ren_target, rwr_target)

    return {
        'nb_base': nb_base, 'nb_liability': nb_liability, 'nb_target': nb_target,
        'ren_base': ren_base, 'ren_liability': ren_liability, 'ren_target': ren_target,
        'rwr_target': rwr_target, 'total_target': total_target,
        'nb_col_pct': nb_col_pct, 'ren_col_pct': ren_col_pct, 'rwr_col_pct': rwr_col_pct,
        'gross_comm': gross_comm, 'safe_net': safe_net, 'cap': cap,
        'paid': paid, 'capped_paid': capped_paid,
        'bonus_pct_of_safe_net': bonus_pct_of_safe_net, 'review_flag': review_flag,
        **mins,
    }


def calc_proposal_c(nb, rwr, ren):
    """Proposal C (BONUS IDEA): PERSISTENCY-FOCUSED.
       Pays flat per-policy. REN equals NB on purpose to incentivize retention.
       RWR drops to $1 to actively discourage churning.
       Uses same collected kicker + safe net cap.

       NB = $8 flat. REN = $8 flat (parity). RWR = $1 flat.
    """
    nb_c, nb_p, nb_col = nb
    rwr_c, rwr_p, rwr_col = rwr
    ren_c, ren_p, ren_col = ren

    nb_col_pct = nb_col / nb_p if nb_p else 0
    ren_col_pct = ren_col / ren_p if ren_p else 0
    rwr_col_pct = rwr_col / rwr_p if rwr_p else 0

    nb_target = nb_c * 8 * kicker(nb_col_pct)
    ren_target = ren_c * 8 * kicker(ren_col_pct)
    rwr_target = rwr_c * 1 * kicker(rwr_col_pct)
    total_target = nb_target + ren_target + rwr_target

    total_collected = nb_col + rwr_col + ren_col
    gross_comm = total_collected * BLENDED_COMM
    safe_net = gross_comm * (1 - ROYALTY) * (1 - OVERHEAD)
    cap = safe_net * CAP_STANDARD
    paid = total_target  # pay the target; cap is a review threshold, not auto-cut
    capped_paid = min(total_target, cap)
    bonus_pct_of_safe_net = paid / safe_net if safe_net else 0
    review_flag = (paid > cap)

    mins = apply_minimums(nb_p, ren_p, nb_target, ren_target, rwr_target)

    return {
        'nb_target': nb_target, 'ren_target': ren_target, 'rwr_target': rwr_target,
        'total_target': total_target,
        'nb_col_pct': nb_col_pct, 'ren_col_pct': ren_col_pct, 'rwr_col_pct': rwr_col_pct,
        'gross_comm': gross_comm, 'safe_net': safe_net, 'cap': cap,
        'paid': paid, 'capped_paid': capped_paid,
        'bonus_pct_of_safe_net': bonus_pct_of_safe_net, 'review_flag': review_flag,
        **mins,
    }


def current_bonus(nb_c, rwr_c):
    """Reverse-engineered current formula."""
    total = nb_c + rwr_c
    if total < 30: return 0
    if total < 35: return 250
    if total < 50: return 350
    return 450 + 10 * (total - 50)


def swap_rwr_to_ren(month_data, pct=0.5):
    """Move pct of RWR (count, premium, collected) to REN. Used for the 50% swap scenario."""
    nb = month_data['NB']
    rwr_c, rwr_p, rwr_col = month_data['RWR']
    ren_c, ren_p, ren_col = month_data['REN']

    swap_c = round(rwr_c * pct)
    swap_p = rwr_p * pct
    swap_col = rwr_col * pct

    return {
        'NB': nb,
        'RWR': (rwr_c - swap_c, rwr_p - swap_p, rwr_col - swap_col),
        'REN': (ren_c + swap_c, ren_p + swap_p, ren_col + swap_col),
    }

# ============================================================================
# STYLING
# ============================================================================
HEADER_FILL = PatternFill('solid', fgColor='1F4E78')
HEADER_FONT = Font(bold=True, color='FFFFFF', size=11)
SUB_FILL = PatternFill('solid', fgColor='BDD7EE')
SUB_FONT = Font(bold=True, color='1F4E78', size=11)
SECTION_FILL = PatternFill('solid', fgColor='2E75B6')
SECTION_FONT = Font(bold=True, color='FFFFFF', size=12)
TITLE_FONT = Font(bold=True, color='1F4E78', size=16)
PROP_A_FILL = PatternFill('solid', fgColor='E2EFDA')   # green-ish (premium)
PROP_B_FILL = PatternFill('solid', fgColor='FFF2CC')   # gold (coverage)
PROP_C_FILL = PatternFill('solid', fgColor='FCE4D6')   # peach (persistency)
CURRENT_FILL = PatternFill('solid', fgColor='F4B084')  # orange (current/legacy)
SWAP_FILL = PatternFill('solid', fgColor='D9E1F2')     # light blue (swap)
WARN_FILL = PatternFill('solid', fgColor='FFC7CE')
BORDER_THIN = Side(border_style='thin', color='808080')
ALL_BORDERS = Border(top=BORDER_THIN, bottom=BORDER_THIN, left=BORDER_THIN, right=BORDER_THIN)


def style_header(cell):
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = ALL_BORDERS


def style_section(cell):
    cell.fill = SECTION_FILL
    cell.font = SECTION_FONT
    cell.alignment = Alignment(horizontal='left', vertical='center')


def style_sub(cell):
    cell.fill = SUB_FILL
    cell.font = SUB_FONT
    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    cell.border = ALL_BORDERS


def style_data(cell):
    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    cell.border = ALL_BORDERS


def style_number(cell):
    cell.alignment = Alignment(horizontal='right', vertical='center')
    cell.border = ALL_BORDERS
    cell.number_format = '#,##0.00'


def style_dollar(cell):
    cell.alignment = Alignment(horizontal='right', vertical='center')
    cell.border = ALL_BORDERS
    cell.number_format = '"$"#,##0.00'


def style_pct(cell):
    cell.alignment = Alignment(horizontal='right', vertical='center')
    cell.border = ALL_BORDERS
    cell.number_format = '0.0%'


def style_int(cell):
    cell.alignment = Alignment(horizontal='right', vertical='center')
    cell.border = ALL_BORDERS
    cell.number_format = '#,##0'


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ============================================================================
# SHEET BUILDERS
# ============================================================================
def build_readme(wb):
    ws = wb.create_sheet('README')
    ws['A1'] = 'Bonus Proposal Workbook - How to Use'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:F1')

    rows = [
        ('', ''),
        ('Purpose', 'Two clean, separated bonus proposals for ownership decision. Plus a third optional idea.'),
        ('', ''),
        ('Proposal A - Premium-Based', 'Bonus is tiered by WRITTEN PREMIUM, then a COLLECTED-% kicker is applied. No coverage add-ons. Clean and self-balancing on premium quality.'),
        ('Proposal B - Coverage-Based', 'Bonus is per-policy by COVERAGE TYPE (PIP/PD or PIP+Comp/Coll = $10 NB base; BI+UM liability bundle = +$5). No premium tiers. Same collected kicker. UM cannot be added without BI - so the bundle is paid as one $5 add-on.'),
        ('Proposal C - Persistency (Bonus Idea)', 'Flat $8 per NB and $8 per REN (parity), $1 per RWR. Strongest message that rewriting is no longer rewarded. Same collected kicker and safe-net cap.'),
        ('', ''),
        ('Important Constraint', 'Do NOT mix Proposal A and Proposal B. They are separate models. Coverage bonuses do not exist in A. Premium tiers do not exist in B.'),
        ('', ''),
        ('Why Renewals Now Pay Separately', 'The current plan counts NB and RWR together toward a single threshold and pays nothing for renewals. That pushes agents to REWRITE rather than RENEW, which costs the customer time and the agency loyalty/persistency. The new plans pay renewals separately and keep rewrites flat at $2 so the math no longer rewards churning the book.'),
        ('', ''),
        ('Why the 50% Swap Examples', 'Today the agents are mostly rewriting. To prove that the new plan REWARDS the shift to renewals, every agent example is also run with 50% of their rewrites converted into renewals (same premium, same collected). That is the realistic mid-term target.'),
        ('', ''),
        ('Profitability Protection (ALL proposals)', 'Bonus PAID = the calculated TARGET (no automatic hard cap). Safe net = collected commission x (1 - 0.175 royalty) x (1 - 0.40 overhead). Ownership REVIEWS any month where target exceeds 40% of safe net (soft flag). Current plan pays ~57% of safe net - all three proposals are cheaper than that.'),
        ('Chargeback (ALL proposals)', '100% of the paid bonus is reversed if the policy cancels or rewrites within 90 days of effective date. THIS is the primary profit protection.'),
        ('', ''),
        ('Minimum Requirement (NEW)', f'Bonus is GATED by monthly premium: NB premium >= ${NB_MIN_PREMIUM:,} AND REN premium >= ${REN_MIN_PREMIUM:,}. Pass the NB gate to earn NB bonus. Pass the REN gate to earn REN bonus. Pass BOTH to earn the RWR bonus (RWR has no own minimum). This replaces today\'s 35-policy NB+RWR floor and points agents at the renewal book.'),
        ('', ''),
        ('Sheet Map', ''),
        ('  1. Executive Summary', 'Bottom-line comparison: Current vs A vs B vs C for all 6 agents across 4 months. Real + 50% swap. Now shows both no-min and with-min totals.'),
        ('  2. Safe Net Explained', 'PLAIN-LANGUAGE walk-through of what safe net is, with single-policy and full-month worked numbers.'),
        ('  3. Minimum Requirements', 'How the $50k/$50k gate works, pass/fail per agent-month, plus $25k/$35k/$75k sensitivity.'),
        ('  4. Previous vs New Detailed', 'Agent-by-agent dollar comparison: today vs no-min vs with-min, by month.'),
        ('  5. Rules Side by Side', 'One-row-per-rule comparison of all four plans.'),
        ('  6. Proposal A Rules', 'Full payout tables for the PREMIUM-based plan.'),
        ('  7. Proposal B Rules', 'Full payout tables for the COVERAGE-based plan.'),
        ('  8. Proposal C Rules', 'Bonus idea: persistency-focused flat plan.'),
        ('  9. Worked Examples', 'Hand-built single-policy walk-throughs for every plan.'),
        ('  10. Agent Examples - Real', 'Month-by-month bonus for the 6 agents on actual Jan-Apr 2026 data, both proposals.'),
        ('  11. Agent Examples - 50% Swap', 'Same 6 agents, 50% of RWR moved to REN. Shows future-state earnings.'),
        ('  12. Real vs Swap Comparison', 'Side-by-side 4-month total delta showing the renewal-shift upside per agent.'),
        ('  13. Coverage Bundle Logic', 'Why BI+UM is bundled and why BI alone does not qualify.'),
        ('  14. Collected Kicker Logic', 'How the collected % multiplier works.'),
        ('  15. Profitability Cap', 'How the safe-net review threshold protects ownership.'),
        ('  16. Chargeback Process', '90-day cancellation/rewrite reversal procedure.'),
        ('  17. Agent Quick Reference', 'One-page printable for the agent floor.'),
        ('  18. Manual Tracker Template', 'Monthly template for verified bonus tracking.'),
        ('  19. Assumptions', 'All economic and modeling assumptions in one place.'),
    ]
    for r, (a, b) in enumerate(rows, 2):
        ws.cell(row=r, column=1, value=a)
        ws.cell(row=r, column=2, value=b)
        ws.cell(row=r, column=1).font = Font(bold=True, size=11)
        ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical='top')
        ws.row_dimensions[r].height = 30

    set_col_widths(ws, [38, 110])


def build_executive_summary(wb):
    ws = wb.create_sheet('Executive Summary')
    ws['A1'] = 'Executive Summary - Current vs Proposal A vs Proposal B vs Proposal C'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:M1')

    ws['A2'] = (f'Six agents, four months. Each proposal is shown two ways: WITHOUT minimums (pure target) and WITH '
                f'the proposed minimums (NB premium >= ${NB_MIN_PREMIUM:,}/mo AND REN premium >= ${REN_MIN_PREMIUM:,}/mo; RWR paid only when BOTH gates pass).')
    ws['A2'].font = Font(italic=True, size=10, color='666666')
    ws.merge_cells('A2:M2')

    # Two scenario blocks: REAL, then 50% SWAP
    headers = ['Month', 'Agent', 'NB', 'RWR', 'REN', 'Current',
               'A no-min', 'A with min', 'B no-min', 'B with min', 'C no-min', 'C with min',
               'A-min vs Current']
    r = 4
    ws.cell(row=r, column=1, value='SCENARIO 1: REAL DATA (Jan-Apr 2026)').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
    r += 1
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=r, column=i, value=h)
        style_header(c)
    r += 1

    real_totals = {'current': 0, 'a': 0, 'a_min': 0, 'b': 0, 'b_min': 0, 'c': 0, 'c_min': 0}
    for month in MONTHS:
        for agent in AGENT_NAMES:
            d = AGENTS[agent][month]
            nb_c, rwr_c, ren_c = d['NB'][0], d['RWR'][0], d['REN'][0]
            cur = current_bonus(nb_c, rwr_c)
            a = calc_proposal_a(d['NB'], d['RWR'], d['REN'])
            b = calc_proposal_b(d['NB'], d['RWR'], d['REN'])
            cp = calc_proposal_c(d['NB'], d['RWR'], d['REN'])
            row_vals = [month, agent, nb_c, rwr_c, ren_c, cur,
                        a['paid'], a['paid_after_min'],
                        b['paid'], b['paid_after_min'],
                        cp['paid'], cp['paid_after_min'],
                        a['paid_after_min'] - cur]
            for i, v in enumerate(row_vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (3,4,5):
                    style_int(c)
                elif i >= 6:
                    style_dollar(c)
                else:
                    style_data(c)
                if i == 7: c.fill = PROP_A_FILL
                if i == 8: c.fill = PROP_A_FILL; c.font = Font(bold=True)
                if i == 9: c.fill = PROP_B_FILL
                if i == 10: c.fill = PROP_B_FILL; c.font = Font(bold=True)
                if i == 11: c.fill = PROP_C_FILL
                if i == 12: c.fill = PROP_C_FILL; c.font = Font(bold=True)
            real_totals['current'] += cur
            real_totals['a'] += a['paid']
            real_totals['a_min'] += a['paid_after_min']
            real_totals['b'] += b['paid']
            real_totals['b_min'] += b['paid_after_min']
            real_totals['c'] += cp['paid']
            real_totals['c_min'] += cp['paid_after_min']
            r += 1

    ws.cell(row=r, column=1, value='REAL 4-MONTH TOTAL (6 agents)').font = Font(bold=True)
    ws.cell(row=r, column=6, value=real_totals['current'])
    ws.cell(row=r, column=7, value=real_totals['a'])
    ws.cell(row=r, column=8, value=real_totals['a_min'])
    ws.cell(row=r, column=9, value=real_totals['b'])
    ws.cell(row=r, column=10, value=real_totals['b_min'])
    ws.cell(row=r, column=11, value=real_totals['c'])
    ws.cell(row=r, column=12, value=real_totals['c_min'])
    ws.cell(row=r, column=13, value=real_totals['a_min'] - real_totals['current'])
    for i in range(6, 14):
        c = ws.cell(row=r, column=i)
        style_dollar(c)
        c.font = Font(bold=True)
        c.fill = SUB_FILL
    ws.cell(row=r, column=1).fill = SUB_FILL
    r += 2

    # SCENARIO 2: 50% SWAP
    ws.cell(row=r, column=1, value='SCENARIO 2: 50% OF RWR CONVERTED TO REN (future-state target)').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
    r += 1
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=r, column=i, value=h)
        style_header(c)
    r += 1

    swap_totals = {'current': 0, 'a': 0, 'a_min': 0, 'b': 0, 'b_min': 0, 'c': 0, 'c_min': 0}
    for month in MONTHS:
        for agent in AGENT_NAMES:
            d = swap_rwr_to_ren(AGENTS[agent][month], 0.5)
            nb_c, rwr_c, ren_c = d['NB'][0], d['RWR'][0], d['REN'][0]
            cur = current_bonus(nb_c, rwr_c)
            a = calc_proposal_a(d['NB'], d['RWR'], d['REN'])
            b = calc_proposal_b(d['NB'], d['RWR'], d['REN'])
            cp = calc_proposal_c(d['NB'], d['RWR'], d['REN'])
            row_vals = [month, agent, nb_c, rwr_c, ren_c, cur,
                        a['paid'], a['paid_after_min'],
                        b['paid'], b['paid_after_min'],
                        cp['paid'], cp['paid_after_min'],
                        a['paid_after_min'] - cur]
            for i, v in enumerate(row_vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (3,4,5):
                    style_int(c)
                elif i >= 6:
                    style_dollar(c)
                else:
                    style_data(c)
                if i == 7: c.fill = PROP_A_FILL
                if i == 8: c.fill = PROP_A_FILL; c.font = Font(bold=True)
                if i == 9: c.fill = PROP_B_FILL
                if i == 10: c.fill = PROP_B_FILL; c.font = Font(bold=True)
                if i == 11: c.fill = PROP_C_FILL
                if i == 12: c.fill = PROP_C_FILL; c.font = Font(bold=True)
            swap_totals['current'] += cur
            swap_totals['a'] += a['paid']
            swap_totals['a_min'] += a['paid_after_min']
            swap_totals['b'] += b['paid']
            swap_totals['b_min'] += b['paid_after_min']
            swap_totals['c'] += cp['paid']
            swap_totals['c_min'] += cp['paid_after_min']
            r += 1

    ws.cell(row=r, column=1, value='SWAP 4-MONTH TOTAL (6 agents)').font = Font(bold=True)
    ws.cell(row=r, column=6, value=swap_totals['current'])
    ws.cell(row=r, column=7, value=swap_totals['a'])
    ws.cell(row=r, column=8, value=swap_totals['a_min'])
    ws.cell(row=r, column=9, value=swap_totals['b'])
    ws.cell(row=r, column=10, value=swap_totals['b_min'])
    ws.cell(row=r, column=11, value=swap_totals['c'])
    ws.cell(row=r, column=12, value=swap_totals['c_min'])
    ws.cell(row=r, column=13, value=swap_totals['a_min'] - swap_totals['current'])
    for i in range(6, 14):
        c = ws.cell(row=r, column=i)
        style_dollar(c)
        c.font = Font(bold=True)
        c.fill = SUB_FILL
    ws.cell(row=r, column=1).fill = SUB_FILL
    r += 3

    # Final read
    ws.cell(row=r, column=1, value='OWNERSHIP READ').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
    r += 1
    reads = [
        f"REAL TOTAL (no-min): Current ${real_totals['current']:,.0f} | A ${real_totals['a']:,.0f} | B ${real_totals['b']:,.0f} | C ${real_totals['c']:,.0f}. All three new plans are cheaper than the current ${real_totals['current']:,.0f}.",
        f"REAL TOTAL (with ${NB_MIN_PREMIUM:,}/${REN_MIN_PREMIUM:,} mins): A ${real_totals['a_min']:,.0f} | B ${real_totals['b_min']:,.0f} | C ${real_totals['c_min']:,.0f}. The minimums hammer the REAL payout because no agent currently has $50k of REN premium in a single month - so NO REN bonus and NO RWR bonus is earned. NB bonus is only earned when NB premium clears $50k, which happens twice in the four months.",
        f"SWAP TOTAL (no-min): Current drops to ${swap_totals['current']:,.0f} (-${real_totals['current']-swap_totals['current']:,.0f}). A ${swap_totals['a']:,.0f} | B ${swap_totals['b']:,.0f} | C ${swap_totals['c']:,.0f}. New plans GROW because renewals now pay.",
        f"SWAP TOTAL (with mins): A ${swap_totals['a_min']:,.0f} | B ${swap_totals['b_min']:,.0f} | C ${swap_totals['c_min']:,.0f}. Even with the swap, only Dialinerys crosses both gates in some months. The $50k REN minimum is aspirational - it forces agents to build the renewal book to unlock the bonus.",
        "WHY THE WITH-MIN NUMBERS ARE SO LOW: today's agents barely have renewals. The $50k REN gate is intentional - it tells the team 'no bonus until you have a renewal book.' If ownership wants softer activation, see the Minimum Requirements sheet for $25k and $35k variants.",
        "CURRENT BONUS GOES DOWN in the swap because RWR count drops, and the current plan only counts NB+RWR. That is the broken incentive we are fixing.",
        "PROTECTION: Pay = calculated target. Review threshold flags months above 40% of safe net (informational). 90-day chargeback reverses any bonus on a policy that cancels/rewrites - that is the real profit shield. Minimums add a second layer: no premium volume, no bonus.",
        "RECOMMENDATION: Proposal A or B with the $50k/$50k minimums starting in Q3, after 90 days of awareness so agents can build renewals. Proposal C only if the goal is simplest payroll.",
    ]
    for txt in reads:
        c = ws.cell(row=r, column=1, value=txt)
        c.font = Font(size=11)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
        ws.row_dimensions[r].height = 44
        r += 1

    set_col_widths(ws, [12, 22, 7, 7, 7, 11, 11, 12, 11, 12, 11, 12, 14])
    ws.freeze_panes = 'C6'


def build_proposal_a(wb):
    ws = wb.create_sheet('Proposal A - Premium')
    ws['A1'] = 'Proposal A: PREMIUM-BASED Plan'
    ws['A1'].font = TITLE_FONT
    ws['A1'].fill = PROP_A_FILL
    ws.merge_cells('A1:G1')

    ws['A2'] = 'How it works: Each policy is paid by WRITTEN PREMIUM tier. Then the COLLECTED % kicker is applied. Then total is capped at 15% of safe net.'
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:G2')

    r = 4
    headers = ['Component', 'Rule', 'Per-Policy Pay', 'Plain English', 'Example', 'Example Pay', 'Notes']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    rows = [
        ('NB', 'Written premium under $1,200', '$5', 'Low-premium NB earns base.', '$1,000 NB', '$5', 'Before kicker/cap'),
        ('NB', '$1,200-$1,799', '$7', 'Medium NB.', '$1,500 NB', '$7', 'Before kicker/cap'),
        ('NB', '$1,800-$2,199', '$9', 'Better premium NB.', '$2,000 NB', '$9', 'Before kicker/cap'),
        ('NB', '$2,200-$2,999', '$11', 'Strong premium NB.', '$2,500 NB', '$11', 'Before kicker/cap'),
        ('NB', '$3,000+', '$11 + $2 per $1k over $3k, cap $25', 'Commercial / high-premium upside.', '$4,500 NB', '$11 + $3 = $14', 'Before kicker/cap'),
        ('REN', 'Under $1,200', '$4', 'Renewals are PAID separately. Below NB on purpose.', '$1,000 REN', '$4', 'Before kicker/cap'),
        ('REN', '$1,200-$1,799', '$5', 'Medium REN.', '$1,500 REN', '$5', 'Before kicker/cap'),
        ('REN', '$1,800+', '$6', 'High REN.', '$2,000 REN', '$6', 'Before kicker/cap'),
        ('RWR', 'Any rewrite', '$2 flat', 'Rewrites are paid but DO NOT drive the bonus.', 'Any RWR', '$2', 'No kicker stacking by tier'),
        ('Collected Kicker', 'Applies to the per-policy pay', '+10% / +15% / +20% / +25%', '15-24% / 25-49% / 50-99% / 100% (PIF)', '$10 target x 25% collected', '$10 x 1.15 = $11.50', 'Same logic both A and B'),
        ('PIF Add (when paid in full)', 'NB under $3,000 = +$8 / NB $3,000+ = +$12 / REN PIF = +$5', 'Per-policy add-on', 'PIF means cash collected upfront - lower risk.', '$2,500 NB PIF', '$11 base + $8 PIF = $19', 'Manager verifies PIF'),
        ('Review Threshold (soft cap)', 'Target > 40% of safe net', 'Flag for ownership review', 'Catches low-collection months. Does NOT auto-reduce pay.', 'Safe net $200, target $80 (40%)', 'Pay $80, flag for review', '40% review / 55% PIF review'),
        ('Chargeback (PRIMARY PROTECTION)', '90 days', '100% reversal', 'If a paid policy cancels or rewrites within 90 days, full bonus is reversed.', '$12 paid Jan, cancels Mar', '-$12 in next payroll', 'Required'),
    ]
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            c.fill = PROP_A_FILL if r % 2 == 0 else PatternFill('solid', fgColor='F4FFF4')
        r += 1

    r += 1
    ws.cell(row=r, column=1, value='Why this plan').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    r += 1
    reasons = [
        '* Simple payroll math: count policies x tier x kicker = target.',
        '* Self-balancing: higher-premium business naturally earns more. No coverage paperwork required.',
        '* Renewals matter for the first time, but stay below NB so the agent still hunts new business.',
        '* Rewrites stay flat at $2 - they are no longer the path to a bigger bonus.',
        '* Profitability cap means a high-target month does not become a payroll problem.',
    ]
    for t in reasons:
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=11)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        r += 1

    set_col_widths(ws, [18, 30, 22, 36, 22, 18, 25])


def build_proposal_b(wb):
    ws = wb.create_sheet('Proposal B - Coverage')
    ws['A1'] = 'Proposal B: COVERAGE-BASED Plan'
    ws['A1'].font = TITLE_FONT
    ws['A1'].fill = PROP_B_FILL
    ws.merge_cells('A1:G1')

    ws['A2'] = 'How it works: Each policy is paid by COVERAGE TYPE. No premium tiers. The COLLECTED % kicker is applied. Then total is capped at 15% of safe net.'
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:G2')

    r = 4
    headers = ['Component', 'Rule', 'Per-Policy Pay', 'Plain English', 'Example', 'Example Pay', 'Notes']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    rows = [
        ('NB Base Coverage', 'PIP/PD only OR PIP + Comp/Coll', '$10 flat', 'Either FL minimum or full coverage qualifies for the base.', 'PIP+PD NB or PIP+Comp+Coll NB', '$10', 'Per policy, before kicker'),
        ('NB Liability Add', 'BI + UM bundled together', '+$5', 'Adds liability. Bundled because UM cannot exist without BI in FL.', 'NB with BI + UM', '$10 + $5 = $15', 'Bundle only - see below'),
        ('NB - BI ALONE', 'BI present but NO UM', '$0 add (base only)', 'Does not earn the liability bundle. Customer is still partially exposed.', 'NB with BI no UM', '$10 base only', 'Manager verifies'),
        ('NB - UM ALONE', 'Not possible', 'Not applicable', 'UM cannot be issued without BI per FL rules. Bundle exists for this reason.', '-', '-', 'Carrier system blocks this'),
        ('REN Base Coverage', 'Same coverage logic as NB', '$6 flat', 'Renewals paid for the first time, but below NB.', 'Any REN', '$6', 'Before kicker'),
        ('REN Liability Add', 'BI + UM bundled', '+$3', 'Same bundle logic, smaller dollar amount since the coverage already existed.', 'REN with BI + UM', '$6 + $3 = $9', 'Manager verifies'),
        ('RWR', 'Any rewrite', '$2 flat', 'No coverage stacking. Rewrites are intentionally small.', 'Any RWR', '$2', 'No bundle on rewrites'),
        ('PIF Add', 'Policy paid in full', '+$8 NB / +$12 high NB / +$5 REN', 'Cash upfront earns extra.', '$2,500 NB PIF + BI/UM', '$10 + $5 + $8 = $23', 'Manager verifies'),
        ('Collected Kicker', 'Same as Proposal A', '+10% / +15% / +20% / +25%', '15-24% / 25-49% / 50-99% / 100%', '$15 target at 25% collected', '$15 x 1.15 = $17.25', 'Identical kicker logic'),
        ('Review Threshold (soft cap)', 'Target > 40% of safe net', 'Flag for ownership review', 'Same as Proposal A. Does NOT auto-reduce.', 'Safe net $200, target $80 (40%)', 'Pay $80, flag for review', '40% review / 55% PIF review'),
        ('Chargeback (PRIMARY PROTECTION)', '90 days', '100% reversal', 'Cancel or rewrite inside 90 days = full bonus reversed.', '$15 paid Jan, cancels Mar', '-$15 in next payroll', 'Required'),
    ]
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            c.fill = PROP_B_FILL if r % 2 == 0 else PatternFill('solid', fgColor='FFFCEA')
        r += 1

    r += 1
    ws.cell(row=r, column=1, value='Why this plan').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    r += 1
    reasons = [
        '* Coverage = money. Agents who upsell BI/UM earn more without depending on premium size.',
        '* BI+UM bundled per Florida coverage logic: UM cannot exist without BI, so paying them as one $5 add-on prevents gaming.',
        '* Renewals receive a separate, smaller payout that creates retention incentive.',
        '* Rewrites stay flat at $2 - same as Proposal A.',
        '* Requires manager verification of coverages on the manual tracker. No verification = no add-on.',
        '* Premium-quality is rewarded indirectly: bigger policies are typically the ones that buy BI+UM.',
    ]
    for t in reasons:
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=11)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        r += 1

    set_col_widths(ws, [22, 34, 28, 40, 28, 22, 25])


def build_proposal_c(wb):
    ws = wb.create_sheet('Proposal C - Persistency')
    ws['A1'] = 'Proposal C: PERSISTENCY-FOCUSED Plan (Bonus Idea)'
    ws['A1'].font = TITLE_FONT
    ws['A1'].fill = PROP_C_FILL
    ws.merge_cells('A1:G1')

    ws['A2'] = 'How it works: Flat per-policy pay. REN equals NB intentionally. RWR drops to $1 to actively discourage churning the book. Same kicker + cap.'
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:G2')

    r = 4
    headers = ['Component', 'Rule', 'Per-Policy Pay', 'Plain English', 'Example', 'Example Pay', 'Notes']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    rows = [
        ('NB', 'Any new business', '$8 flat', 'No premium tier, no coverage detail. Simplest possible.', 'Any NB', '$8', 'Before kicker/cap'),
        ('REN', 'Any renewal', '$8 flat (PARITY with NB)', 'Renewing pays exactly the same as writing new. Strong retention signal.', 'Any REN', '$8', 'Before kicker/cap'),
        ('RWR', 'Any rewrite', '$1 flat (LOWER than A and B)', 'Active disincentive: rewriting costs the agent earnings.', 'Any RWR', '$1', 'Before kicker'),
        ('Collected Kicker', 'Same as Proposals A and B', '+10% / +15% / +20% / +25%', 'Standard kicker.', '$8 target at 25%', '$8 x 1.15 = $9.20', 'Same'),
        ('Review Threshold (soft cap)', 'Target > 40% of safe net', 'Flag for ownership review', 'Same as A and B.', '-', '-', '40% review threshold'),
        ('Chargeback (PRIMARY PROTECTION)', '90 days', '100% reversal', 'Same', '-', '-', 'Required'),
    ]
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            c.fill = PROP_C_FILL if r % 2 == 0 else PatternFill('solid', fgColor='FFF6F0')
        r += 1

    r += 1
    ws.cell(row=r, column=1, value='When to consider this plan').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    r += 1
    reasons = [
        '* Use when the priority is to BREAK the rewrite habit fastest, even at the cost of premium-quality signal.',
        '* Simplest plan: no premium lookup, no coverage verification. Easiest to administer.',
        '* REN = NB parity is unusual and will be a strong cultural message.',
        '* Weakness: does NOT reward upselling premium or coverage. Best paired with a separate quarterly contest if the agency also wants premium growth.',
        '* Recommended only if Proposal A or B is rejected as too complex.',
    ]
    for t in reasons:
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=11)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        r += 1

    set_col_widths(ws, [22, 30, 28, 40, 22, 18, 22])


def build_worked_examples(wb):
    ws = wb.create_sheet('Worked Examples')
    ws['A1'] = 'Worked Examples (one policy at a time)'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:J1')

    ws['A2'] = 'These hand-built examples show the per-policy target math for every plan. The review threshold (40% of safe net) is monthly, not per-policy. Pay = target.'
    ws['A2'].font = Font(italic=True, size=10, color='666666')
    ws.merge_cells('A2:J2')

    r = 4
    headers = ['Scenario', 'Type', 'Premium', 'Coverages', 'Collected %', 'Current', 'Proposal A (Premium)', 'Proposal B (Coverage)', 'Proposal C (Persistency)', 'Plain-English Read']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    examples = [
        # (Scenario, Type, Premium, Coverages, Coll%, Current logic, A, B, C, Read)
        ('Basic FL NB - PIP/PD only',          'NB', 1000, 'PIP+PD',           '25%', 'Counts toward tier',  '$5 x 1.15 = $5.75',              '$10 x 1.15 = $11.50',                '$8 x 1.15 = $9.20',  'B pays the most because coverage = money, even at low premium.'),
        ('Full Coverage NB - no liability',    'NB', 1500, 'PIP+PD+Comp+Coll', '25%', 'Counts toward tier',  '$7 x 1.15 = $8.05',              '$10 x 1.15 = $11.50',                '$8 x 1.15 = $9.20',  'A pays more because premium is higher. B pays base only (no BI+UM).'),
        ('Full Coverage + Liability NB',       'NB', 2000, 'PIP+PD+CC+BI+UM',  '25%', 'Counts toward tier',  '$9 x 1.15 = $10.35',             '$10 + $5 = $15 x 1.15 = $17.25',     '$8 x 1.15 = $9.20',  'B clearly wins because liability bundle adds $5.'),
        ('NB with BI ALONE (no UM)',           'NB', 2000, 'PIP+PD+BI no UM',  '25%', 'Counts toward tier',  '$9 x 1.15 = $10.35',             '$10 x 1.15 = $11.50 (base only)',    '$8 x 1.15 = $9.20',  'B does NOT pay the bundle because BI alone does not qualify.'),
        ('High-premium PIF NB',                'NB', 4500, 'PIP+PD+CC+BI+UM',  '100% (PIF)', 'Counts toward tier','($11+$3 over-$3k+$12 PIF) x 1.25 = $32.50','($10+$5+$12) x 1.25 = $33.75',       '$8 x 1.25 = $10.00', 'PIF adds the biggest dollars on A and B; C stays simple.'),
        ('Standard Renewal',                   'REN',1500, 'PIP+PD+CC',        '25%', 'Pays $0 today',       '$5 x 1.15 = $5.75',              '$6 x 1.15 = $6.90',                  '$8 x 1.15 = $9.20',  'C pays the most because REN = NB parity is the whole point of C.'),
        ('Renewal + Liability + PIF',          'REN',2000, 'Full + BI+UM PIF', '100% (PIF)', 'Pays $0 today','($6+$5 PIF) x 1.25 = $13.75',     '($6+$3+$5) x 1.25 = $17.50',         '$8 x 1.25 = $10.00', 'B and A both reward renewal PIF; C does not differentiate.'),
        ('Rewrite (any premium)',              'RWR',1200, 'Any',              '25%', 'Counts toward tier (BAD)','$2 x 1.15 = $2.30',           '$2 x 1.15 = $2.30',                  '$1 x 1.15 = $1.15',  'All three new plans deprioritize RWR. C goes further with $1.'),
    ]

    for ex in examples:
        for i, v in enumerate(ex, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if i == 7: c.fill = PROP_A_FILL
            elif i == 8: c.fill = PROP_B_FILL
            elif i == 9: c.fill = PROP_C_FILL
            elif i == 6: c.fill = CURRENT_FILL
        ws.row_dimensions[r].height = 32
        r += 1

    set_col_widths(ws, [32, 7, 11, 18, 13, 22, 26, 30, 25, 50])


def build_agent_examples(wb, swap=False):
    """Detailed month-by-month per-agent worked examples for BOTH proposals (A and B)."""
    name = 'Agent Examples - 50% Swap' if swap else 'Agent Examples - Real'
    ws = wb.create_sheet(name)
    ws['A1'] = f"Agent Examples - {'50% RWR converted to REN' if swap else 'REAL Jan-Apr 2026 data'}"
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:R1')

    subtitle = ('Same premiums, same collected $, but 50% of each agent\'s rewrites are reclassified as renewals. Shows future-state earnings.'
                if swap else
                'Actual NB/RWR/REN volumes from the source data, calculated for every proposal.')
    ws['A2'] = subtitle
    ws['A2'].font = Font(italic=True, size=10, color='666666')
    ws.merge_cells('A2:R2')

    r = 4
    for agent in AGENT_NAMES:
        # Agent header
        ws.cell(row=r, column=1, value=agent).font = SECTION_FONT
        ws.cell(row=r, column=1).fill = SECTION_FILL
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=18)
        ws.row_dimensions[r].height = 22
        r += 1

        # Volume table header
        vol_headers = ['Month', 'NB Count', 'NB Premium', 'NB Collected', 'NB Coll %',
                       'RWR Count', 'RWR Premium', 'RWR Collected', 'RWR Coll %',
                       'REN Count', 'REN Premium', 'REN Collected', 'REN Coll %',
                       'Avg NB Prem', 'Avg REN Prem', 'Total Collected', 'Safe Net', 'Cap']
        for i, h in enumerate(vol_headers, 1):
            style_header(ws.cell(row=r, column=i, value=h))
        r += 1

        agent_data = {}
        for month in MONTHS:
            d = swap_rwr_to_ren(AGENTS[agent][month], 0.5) if swap else AGENTS[agent][month]
            nb_c, nb_p, nb_col = d['NB']
            rwr_c, rwr_p, rwr_col = d['RWR']
            ren_c, ren_p, ren_col = d['REN']
            agent_data[month] = d

            total_col = nb_col + rwr_col + ren_col
            safe_net = total_col * BLENDED_COMM * (1 - ROYALTY) * (1 - OVERHEAD)
            cap = safe_net * CAP_STANDARD

            vals = [month,
                    nb_c, nb_p, nb_col, (nb_col / nb_p if nb_p else 0),
                    rwr_c, rwr_p, rwr_col, (rwr_col / rwr_p if rwr_p else 0),
                    ren_c, ren_p, ren_col, (ren_col / ren_p if ren_p else 0),
                    nb_p / nb_c if nb_c else 0, ren_p / ren_c if ren_c else 0,
                    total_col, safe_net, cap]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i == 1: style_data(c)
                elif i in (2, 6, 10): style_int(c)
                elif i in (5, 9, 13): style_pct(c)
                else: style_dollar(c)
            r += 1
        r += 1

        # Proposal A detail
        ws.cell(row=r, column=1, value='PROPOSAL A - PREMIUM-BASED').font = SUB_FONT
        ws.cell(row=r, column=1).fill = PROP_A_FILL
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=18)
        r += 1

        a_headers = ['Month', 'NB Tier', 'NB Per-Policy $', 'NB Target (w/ kicker)', 'REN Tier', 'REN Per-Policy $', 'REN Target (w/ kicker)', 'RWR Target (w/ kicker)', 'Total Target (no min)', 'PAID after MIN', 'Gates (NB/REN/RWR)', 'Safe Net', 'Bonus % SN', 'vs Current', 'Current Bonus', '', '', '']
        for i, h in enumerate(a_headers, 1):
            if h:
                style_header(ws.cell(row=r, column=i, value=h))
        r += 1

        a_totals = {'paid': 0, 'paid_min': 0, 'cur': 0}
        for month in MONTHS:
            d = agent_data[month]
            a = calc_proposal_a(d['NB'], d['RWR'], d['REN'])
            cur = current_bonus(d['NB'][0], d['RWR'][0])
            gates = f"{'Y' if a['nb_qual'] else 'n'}/{'Y' if a['ren_qual'] else 'n'}/{'Y' if a['rwr_qual'] else 'n'}"
            vals = [month, a['nb_tier_label'], a['nb_per'], a['nb_target'],
                    a['ren_tier_label'], a['ren_per'], a['ren_target'],
                    a['rwr_target'], a['paid'], a['paid_after_min'], gates,
                    a['safe_net'], a['bonus_pct_of_safe_net'],
                    a['paid_after_min'] - cur, cur]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (1, 2, 5, 11): style_data(c)
                elif i == 13: style_pct(c)
                elif i in (3, 6): style_dollar(c)
                else: style_dollar(c)
                c.fill = PROP_A_FILL
                if i == 10: c.font = Font(bold=True)
                if i == 11:
                    if a['nb_qual'] and a['ren_qual']: c.fill = PatternFill('solid', fgColor='C6EFCE')
                    elif a['nb_qual'] or a['ren_qual']: c.fill = PatternFill('solid', fgColor='FFEB9C')
                    else: c.fill = WARN_FILL
            a_totals['paid'] += a['paid']
            a_totals['paid_min'] += a['paid_after_min']
            a_totals['cur'] += cur
            r += 1

        ws.cell(row=r, column=1, value='4-Month Total').font = Font(bold=True)
        ws.cell(row=r, column=9, value=a_totals['paid'])
        ws.cell(row=r, column=10, value=a_totals['paid_min'])
        ws.cell(row=r, column=14, value=a_totals['paid_min'] - a_totals['cur'])
        ws.cell(row=r, column=15, value=a_totals['cur'])
        for i in range(1, 16):
            c = ws.cell(row=r, column=i)
            c.fill = SUB_FILL
            if i in (9, 10, 14, 15):
                style_dollar(c)
                c.font = Font(bold=True)
        r += 2

        # Proposal B detail
        ws.cell(row=r, column=1, value='PROPOSAL B - COVERAGE-BASED').font = SUB_FONT
        ws.cell(row=r, column=1).fill = PROP_B_FILL
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=18)
        r += 1

        b_headers = ['Month', 'NB Base ($10 x ct)', 'NB Liab Add ($5 x 35%)', 'NB Target (w/ kicker)', 'REN Base ($6 x ct)', 'REN Liab Add ($3 x 30%)', 'REN Target (w/ kicker)', 'RWR Target (w/ kicker)', 'Total Target (no min)', 'PAID after MIN', 'Gates (NB/REN/RWR)', 'Safe Net', 'Bonus % SN', 'vs Current', 'Current Bonus']
        for i, h in enumerate(b_headers, 1):
            style_header(ws.cell(row=r, column=i, value=h))
        r += 1

        b_totals = {'paid': 0, 'paid_min': 0, 'cur': 0}
        for month in MONTHS:
            d = agent_data[month]
            b = calc_proposal_b(d['NB'], d['RWR'], d['REN'])
            cur = current_bonus(d['NB'][0], d['RWR'][0])
            gates = f"{'Y' if b['nb_qual'] else 'n'}/{'Y' if b['ren_qual'] else 'n'}/{'Y' if b['rwr_qual'] else 'n'}"
            vals = [month, b['nb_base'], b['nb_liability'], b['nb_target'],
                    b['ren_base'], b['ren_liability'], b['ren_target'],
                    b['rwr_target'], b['paid'], b['paid_after_min'], gates,
                    b['safe_net'], b['bonus_pct_of_safe_net'],
                    b['paid_after_min'] - cur, cur]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (1, 11): style_data(c)
                elif i == 13: style_pct(c)
                else: style_dollar(c)
                c.fill = PROP_B_FILL
                if i == 10: c.font = Font(bold=True)
                if i == 11:
                    if b['nb_qual'] and b['ren_qual']: c.fill = PatternFill('solid', fgColor='C6EFCE')
                    elif b['nb_qual'] or b['ren_qual']: c.fill = PatternFill('solid', fgColor='FFEB9C')
                    else: c.fill = WARN_FILL
            b_totals['paid'] += b['paid']
            b_totals['paid_min'] += b['paid_after_min']
            b_totals['cur'] += cur
            r += 1

        ws.cell(row=r, column=1, value='4-Month Total').font = Font(bold=True)
        ws.cell(row=r, column=9, value=b_totals['paid'])
        ws.cell(row=r, column=10, value=b_totals['paid_min'])
        ws.cell(row=r, column=14, value=b_totals['paid_min'] - b_totals['cur'])
        ws.cell(row=r, column=15, value=b_totals['cur'])
        for i in range(1, 16):
            c = ws.cell(row=r, column=i)
            c.fill = SUB_FILL
            if i in (9, 10, 14, 15):
                style_dollar(c)
                c.font = Font(bold=True)
        r += 2

        # Proposal C detail
        ws.cell(row=r, column=1, value='PROPOSAL C - PERSISTENCY (BONUS IDEA)').font = SUB_FONT
        ws.cell(row=r, column=1).fill = PROP_C_FILL
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=18)
        r += 1

        c_headers = ['Month', 'NB ($8 x ct)', 'REN ($8 x ct)', 'RWR ($1 x ct)', 'NB Target', 'REN Target', 'RWR Target', 'Total Target (no min)', 'PAID after MIN', 'Gates (NB/REN/RWR)', 'Safe Net', 'Bonus % SN', 'vs Current', 'Current Bonus']
        for i, h in enumerate(c_headers, 1):
            style_header(ws.cell(row=r, column=i, value=h))
        r += 1

        c_totals = {'paid': 0, 'paid_min': 0, 'cur': 0}
        for month in MONTHS:
            d = agent_data[month]
            cp = calc_proposal_c(d['NB'], d['RWR'], d['REN'])
            cur = current_bonus(d['NB'][0], d['RWR'][0])
            gates = f"{'Y' if cp['nb_qual'] else 'n'}/{'Y' if cp['ren_qual'] else 'n'}/{'Y' if cp['rwr_qual'] else 'n'}"
            vals = [month, d['NB'][0]*8, d['REN'][0]*8, d['RWR'][0]*1,
                    cp['nb_target'], cp['ren_target'], cp['rwr_target'],
                    cp['paid'], cp['paid_after_min'], gates,
                    cp['safe_net'], cp['bonus_pct_of_safe_net'],
                    cp['paid_after_min'] - cur, cur]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (1, 10): style_data(c)
                elif i == 12: style_pct(c)
                else: style_dollar(c)
                c.fill = PROP_C_FILL
                if i == 9: c.font = Font(bold=True)
                if i == 10:
                    if cp['nb_qual'] and cp['ren_qual']: c.fill = PatternFill('solid', fgColor='C6EFCE')
                    elif cp['nb_qual'] or cp['ren_qual']: c.fill = PatternFill('solid', fgColor='FFEB9C')
                    else: c.fill = WARN_FILL
            c_totals['paid'] += cp['paid']
            c_totals['paid_min'] += cp['paid_after_min']
            c_totals['cur'] += cur
            r += 1

        ws.cell(row=r, column=1, value='4-Month Total').font = Font(bold=True)
        ws.cell(row=r, column=8, value=c_totals['paid'])
        ws.cell(row=r, column=9, value=c_totals['paid_min'])
        ws.cell(row=r, column=13, value=c_totals['paid_min'] - c_totals['cur'])
        ws.cell(row=r, column=14, value=c_totals['cur'])
        for i in range(1, 15):
            c = ws.cell(row=r, column=i)
            c.fill = SUB_FILL
            if i in (8, 9, 13, 14):
                style_dollar(c)
                c.font = Font(bold=True)
        r += 3

    set_col_widths(ws, [12, 13, 14, 14, 11, 13, 14, 14, 11, 13, 14, 14, 11, 13, 13, 15, 13, 13])


def build_comparison(wb):
    ws = wb.create_sheet('Real vs Swap Comparison')
    ws['A1'] = 'Real vs 50%-Swap Comparison (4-month totals per agent)'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:M1')

    ws['A2'] = ('Same agents, same months. NO-MIN columns show target pay. WITH-MIN columns apply the '
                f'${NB_MIN_PREMIUM:,}/{REN_MIN_PREMIUM:,} gates. The renewal-shift upside is the main message: '
                'today the current plan punishes the swap; the new plans reward it.')
    ws['A2'].font = Font(italic=True, size=10, color='666666')
    ws.merge_cells('A2:M2')

    r = 4
    headers = ['Agent', 'Current (Real)', 'Current (Swap)',
               'A no-min (Real)', 'A no-min (Swap)',
               'A WITH MIN (Real)', 'A WITH MIN (Swap)',
               'B WITH MIN (Real)', 'B WITH MIN (Swap)',
               'C WITH MIN (Real)', 'C WITH MIN (Swap)',
               'A-min Swap-Real', 'B-min Swap-Real']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    grand = {k: 0 for k in ['cur_r','cur_s','a_r','a_s','am_r','am_s','b_r','b_s','bm_r','bm_s','c_r','c_s','cm_r','cm_s']}
    for agent in AGENT_NAMES:
        a = {k: 0 for k in grand}
        for month in MONTHS:
            real = AGENTS[agent][month]
            swap = swap_rwr_to_ren(real, 0.5)
            a['cur_r'] += current_bonus(real['NB'][0], real['RWR'][0])
            a['cur_s'] += current_bonus(swap['NB'][0], swap['RWR'][0])
            ra = calc_proposal_a(real['NB'], real['RWR'], real['REN'])
            sa = calc_proposal_a(swap['NB'], swap['RWR'], swap['REN'])
            rb = calc_proposal_b(real['NB'], real['RWR'], real['REN'])
            sb = calc_proposal_b(swap['NB'], swap['RWR'], swap['REN'])
            rc = calc_proposal_c(real['NB'], real['RWR'], real['REN'])
            sc = calc_proposal_c(swap['NB'], swap['RWR'], swap['REN'])
            a['a_r']  += ra['paid'];  a['a_s']  += sa['paid']
            a['am_r'] += ra['paid_after_min']; a['am_s'] += sa['paid_after_min']
            a['b_r']  += rb['paid'];  a['b_s']  += sb['paid']
            a['bm_r'] += rb['paid_after_min']; a['bm_s'] += sb['paid_after_min']
            a['c_r']  += rc['paid'];  a['c_s']  += sc['paid']
            a['cm_r'] += rc['paid_after_min']; a['cm_s'] += sc['paid_after_min']
        row_vals = [agent, a['cur_r'], a['cur_s'],
                    a['a_r'], a['a_s'], a['am_r'], a['am_s'],
                    a['bm_r'], a['bm_s'], a['cm_r'], a['cm_s'],
                    a['am_s'] - a['am_r'], a['bm_s'] - a['bm_r']]
        for col, val in enumerate(row_vals, 1):
            c = ws.cell(row=r, column=col, value=val)
            if col == 1: style_data(c)
            else: style_dollar(c)
            if col in (6, 7): c.fill = PROP_A_FILL
            if col in (8, 9): c.fill = PROP_B_FILL
            if col in (10, 11): c.fill = PROP_C_FILL
        for k in grand: grand[k] += a[k]
        r += 1

    ws.cell(row=r, column=1, value='GRAND TOTAL (6 agents, 4 months)').font = Font(bold=True)
    grand_vals = ['GRAND TOTAL (6 agents, 4 months)',
                  grand['cur_r'], grand['cur_s'],
                  grand['a_r'], grand['a_s'], grand['am_r'], grand['am_s'],
                  grand['bm_r'], grand['bm_s'], grand['cm_r'], grand['cm_s'],
                  grand['am_s'] - grand['am_r'], grand['bm_s'] - grand['bm_r']]
    for col, val in enumerate(grand_vals, 1):
        c = ws.cell(row=r, column=col, value=val)
        if col == 1: style_data(c)
        else: style_dollar(c)
        c.font = Font(bold=True)
        c.fill = SUB_FILL
    r += 2

    ws.cell(row=r, column=1, value='What this shows').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
    r += 1
    reads = [
        '"Real" columns are todays numbers. "Swap" columns are what each agent would earn if 50% of their rewrites became renewals (same dollars, just reclassified).',
        'Under the CURRENT plan, the Swap actually DROPS pay - because the current plan rewards rewrites in the count tier. Broken incentive.',
        'Under Proposals A, B, C no-min, the Swap INCREASES pay - the behavior change ownership wants is rewarded.',
        'Under WITH-MIN, the Swap also increases pay (more agents clear the REN gate). Tighter and harder, but the direction is right.',
        'Use this tab to defend the rollout: "Here is the dollar reward each agent gets for stopping the rewrite habit."',
    ]
    for t in reads:
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=11)
        c.alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
        ws.row_dimensions[r].height = 30
        r += 1

    set_col_widths(ws, [25, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 14, 14])


def build_coverage_logic(wb):
    ws = wb.create_sheet('Coverage Bundle Logic')
    ws['A1'] = 'Coverage Bundle Logic (Proposal B)'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:E1')

    ws['A2'] = "Why BI and UM are paid as ONE $5 bundle on Proposal B."
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:E2')

    r = 4
    headers = ['Scenario', 'BI Present?', 'UM Present?', 'Liability Add Paid?', 'Why']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    rows = [
        ('Full liability (the goal)', 'YES', 'YES', '+$5 NB / +$3 REN', 'Customer is properly covered. Agent gets the bundle pay.'),
        ('BI alone (partial coverage)', 'YES', 'NO', 'NO add - base only', 'Customer is exposed if the at-fault driver has no insurance. Bundle is the standard sell - we will not pay half of it.'),
        ('UM alone (not possible)', 'NO', 'YES', 'N/A', 'Florida (and most carriers) require BI to be on the policy before UM can be issued. The bundle exists for this reason.'),
        ('Neither', 'NO', 'NO', 'NO add', 'Base $10 NB / $6 REN only.'),
    ]
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = PROP_B_FILL
        ws.row_dimensions[r].height = 28
        r += 1

    r += 2
    ws.cell(row=r, column=1, value='What about Comp/Collision?').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    r += 1

    cc = [
        'Comp/Collision is included in the BASE $10 NB pay. PIP/PD alone, OR PIP+PD+Comp+Coll, both qualify for the base.',
        'This means an agent who sells full coverage gets the $10 base, the same as someone who sells FL minimum. The DIFFERENTIATION comes from BI+UM (the liability bundle).',
        'Rationale: in Florida, comp/coll is often required by lienholders, so paying extra for it is rewarding what the carrier/lender already mandates. Liability (BI+UM) is the agent\'s real upsell job - that is what earns the +$5.',
    ]
    for t in cc:
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=11)
        c.alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 32
        r += 1

    set_col_widths(ws, [26, 14, 14, 22, 50])


def build_kicker_logic(wb):
    ws = wb.create_sheet('Collected Kicker Logic')
    ws['A1'] = 'Collected % Kicker (applies to BOTH proposals)'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:E1')

    ws['A2'] = 'The kicker is applied to the per-policy target BEFORE the safe-net cap. It rewards collecting cash, not just writing premium.'
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:E2')

    r = 4
    headers = ['Collected % of Premium', 'Multiplier', 'Plain English', 'Example: $10 target', 'Example after Kicker']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    rows = [
        ('Below 15%', 'x 1.00 (none)', 'Premium barely collected. No kicker.', '$10', '$10.00'),
        ('15% - 24%', 'x 1.10 (+10%)', 'Minimum down. Small kicker.', '$10', '$11.00'),
        ('25% - 49%', 'x 1.15 (+15%)', 'Standard down. Standard kicker.', '$10', '$11.50'),
        ('50% - 99%', 'x 1.20 (+20%)', 'High collection. Stronger kicker.', '$10', '$12.00'),
        ('100% (PIF)', 'x 1.25 (+25%)', 'Paid in full. Top kicker.', '$10', '$12.50'),
    ]
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = SUB_FILL
        r += 1

    r += 2
    ws.cell(row=r, column=1, value='Why use the collected % instead of the written premium').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    r += 1
    why = [
        'Written premium is the carrier number. Collected premium is the agency number - it is the cash that actually pays commission, rent, and payroll.',
        'Without the kicker, a policy with a $50 down looks the same as one with a $500 down. With the kicker, the agent who pushes for a bigger down earns more.',
        'The kicker is INSIDE the cap. The cap is still 15% of safe net. So a high kicker on a low-collected policy cannot break ownership economics.',
    ]
    for t in why:
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=11)
        c.alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 30
        r += 1

    set_col_widths(ws, [22, 18, 38, 18, 22])


def build_profitability(wb):
    ws = wb.create_sheet('Profitability Cap')
    ws['A1'] = 'Profitability Review (applies to ALL proposals)'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:H1')

    ws['A2'] = 'Policy: PAY THE TARGET. The "cap" is a SOFT REVIEW THRESHOLD - if monthly bonus target exceeds 40% of safe net, ownership reviews that agent-month. Today the current plan pays ~57% of safe net so all three proposals are cheaper.'
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:H2')

    r = 4
    headers = ['Step', 'Formula', 'Example Input', 'Example Output', 'Why']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    rows = [
        ('1. Gross commission', 'Collected $ x Blended Carrier Comm %', '$5,000 collected x 11%', '$550.00', 'What the agency actually earns from the carrier on the collected portion.'),
        ('2. After royalty', 'Gross Comm x (1 - 17.5%)', '$550 x 0.825', '$453.75', 'Franchise royalty removed first.'),
        ('3. Safe net', 'After-royalty x (1 - 40% overhead)', '$453.75 x 0.60', '$272.25', 'Cash available after rent/payroll/tech/admin.'),
        ('4. Bonus target (proposal)', 'Per-policy rules + collected kicker', 'Per the proposal sheet', 'e.g., $70', 'Calculated by the plan rules.'),
        ('5. Bonus % of safe net', 'Target / safe net', '$70 / $272.25', '25.7%', 'How much of safe net is going to bonus this month for this agent.'),
        ('6. Review threshold', '40% of safe net', '$272.25 x 0.40 = $108.90', 'Compare', 'If target <= $108.90, auto-pay. If target > $108.90, ownership reviews.'),
        ('7. What ownership reviews', 'High-bonus % months', 'Why is collection so low?', 'Coach or override', 'Catches the case where an agent writes premium without collecting.'),
        ('8. PRIMARY PROTECTION', '90-day chargeback', '100% reversal', 'On cancel/rewrite', 'This is the real profit shield - paid bonus is reversed if the policy does not stick.'),
    ]
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = SUB_FILL
        ws.row_dimensions[r].height = 26
        r += 1

    r += 2
    ws.cell(row=r, column=1, value='Why we are NOT using a hard automatic cap').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    r += 1
    why = [
        'A hard cap on every month at a low % (e.g., 15%) would squash the proposals so much they all pay the same amount - making A, B, and C indistinguishable. That defeats the purpose of having choices.',
        'The 40% REVIEW threshold flags outlier months (low collection, high written premium) so ownership can act, without making the everyday bonus arbitrary.',
        'The 90-day CHARGEBACK is the actual profit protection. If a policy cancels in 90 days, the bonus is reversed. That matches carrier commission chargeback exposure exactly.',
        'The 4-month modeled cost: Current $12,220, Proposal A $6,219, Proposal B $9,459, Proposal C $6,481. All three proposals are CHEAPER than the current plan even at full target.',
    ]
    for t in why:
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=11)
        c.alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 38
        r += 1

    set_col_widths(ws, [22, 36, 26, 18, 60])


def build_chargeback(wb):
    ws = wb.create_sheet('Chargeback Process')
    ws['A1'] = '90-Day Cancellation and Rewrite Chargeback'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:G1')

    ws['A2'] = 'Every bonused policy is reviewed 90 days after effective date. If it cancels or rewrites, the paid bonus is reversed 100%.'
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:G2')

    r = 4
    headers = ['Step', 'Owner', 'Action', 'Timing', 'Required Fields', 'Payroll Result', 'Example']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    rows = [
        ('1', 'Agent',    'Enter policy in monthly tracker', 'Same month written', 'Customer, policy, company, type, premium, collected, PIF, coverages', 'Eligible for review',  'NB policy written Jan 12'),
        ('2', 'Manager',  'Verify premium, collected, PIF, coverages', 'Before payroll',     'Receipts, dec page, MVR, carrier proof',                                'Approved or denied',    'BI/UM proof attached'),
        ('3', 'Payroll',  'Pay approved bonus',                'Monthly payroll',     'Approved tracker row',                                                  'Bonus paid',           '$15 paid in January'),
        ('4', 'Manager',  'Review policy 90 days after effective','Month 3',         'Carrier status, cancel/rewrite report',                                'Keep or chargeback',   'Jan 12 reviewed Apr 12'),
        ('5', 'Payroll',  'If cancel/rewrite, reverse 100%',   'Next payroll',        'Cancel/rewrite flag',                                                  'Negative bonus line',  '-$15 chargeback'),
        ('6', 'Manager',  'Document exception if ownership approves', 'As needed',   'Written approval',                                                       'Exception logged',     'Carrier error fix'),
        ('7', 'Ownership','Review monthly chargeback totals',  'Monthly',             'Chargeback report',                                                    'Coach or adjust plan', 'High chargeback agent flagged'),
    ]
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = SUB_FILL
        ws.row_dimensions[r].height = 30
        r += 1

    set_col_widths(ws, [6, 12, 30, 18, 30, 22, 28])


def build_manual_tracker(wb):
    ws = wb.create_sheet('Manual Tracker Template')
    ws['A1'] = 'Monthly Manual Bonus Tracker - Template'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:T1')

    ws['A2'] = 'One copy per month per agent. Only verified rows get paid. Use this whether running Proposal A, B, or C.'
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:T2')

    r = 4
    headers = ['Agent', 'Customer', 'Policy #', 'Type', 'Carrier', 'Eff Date',
               'Written Premium', 'Carrier Comm %', 'Collected $',
               'BI?', 'UM?', 'BI+UM Bundle?', 'Comp?', 'Coll?', 'PIF?',
               '90-Day Review Date', 'Target Bonus', 'Cap', 'Final Paid', 'Chargeback?']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    for _ in range(30):
        for i in range(1, 21):
            c = ws.cell(row=r, column=i, value='')
            style_data(c)
        r += 1

    set_col_widths(ws, [16, 16, 12, 8, 14, 10, 12, 11, 12, 6, 6, 12, 6, 6, 6, 14, 11, 9, 11, 12])


def build_safe_net_simple(wb):
    """Plain-language explanation of what 'safe net' means and where it comes from."""
    ws = wb.create_sheet('Safe Net Explained')
    ws['A1'] = 'Safe Net Explained Simply'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:F1')

    ws['A2'] = "Safe net = the money the AGENCY actually has left after paying corporate and overhead. The bonus comes out of safe net, so it's the number we use to size what we can afford to pay."
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:F2')

    r = 4
    # Walk-through with a single $1,000 policy at 25% collected
    ws.cell(row=r, column=1, value='STEP-BY-STEP: A $1,000 policy, customer pays $250 down (25%), carrier commission 10%').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1

    headers = ['Step', 'What happens', 'Math', 'Dollar amount', 'Running total', 'Plain English']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    steps = [
        ('1', 'Customer pays down payment', '$250 collected', '$250.00', '$250.00 in our hands', 'This is the money the agency actually has. The other $750 is on a payment plan.'),
        ('2', 'Carrier pays commission on what was collected', '$250 x 10% comm', '$25.00', '$25.00 of commission earned', 'Carriers pay 8-15% on collected premium. We use 11% as the blended rate.'),
        ('3', 'Royalty goes to corporate', '$25 x 17.5% royalty', '-$4.38', '$20.63 left after royalty', 'Every dollar of commission gives up 17.5 cents to the franchise.'),
        ('4', 'Overhead reserve (rent, payroll, tech, admin, bank fees)', '$20.63 x 40% overhead', '-$8.25', '$12.38 SAFE NET', 'This is the conservative estimate of what every commission dollar costs to keep the doors open.'),
        ('5', 'Bonus review threshold', '$12.38 x 40% (soft cap)', '$4.95 cap', '$4.95 is the most we want to spend on bonus for this one policy in a month', "If the bonus target on this policy is above $4.95, ownership flags it for review. It does NOT auto-cut."),
        ('6', 'What the agency keeps', '$12.38 - actual bonus paid', 'Variable', 'Profit + taxes', "Whatever is left of safe net after bonus is real profit. That's why ownership cares about the safe net number."),
    ]
    for row in steps:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = SUB_FILL
        ws.row_dimensions[r].height = 36
        r += 1

    r += 1
    ws.cell(row=r, column=1, value='SAME MATH AT THE MONTHLY LEVEL').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1
    ws.cell(row=r, column=1, value="In the workbook the SAME formula is applied at the agent-month level. Add up everything the agent collected that month (NB + RWR + REN), multiply by 11% carrier commission, subtract 17.5% royalty, subtract 40% overhead. The result is that agent's monthly safe net. The bonus is judged against THAT number.").alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 50
    r += 2

    # Worked example: Abel Guaina January
    ws.cell(row=r, column=1, value="WORKED EXAMPLE - Abel Guaina, January 2026").font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1

    headers2 = ['Step', 'Formula', 'Plug-in', 'Output', 'What it means', '']
    for i, h in enumerate(headers2, 1):
        if h: style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    # Compute Abel Jan numbers fresh
    d = AGENTS['Abel Guaina']['January']
    total_col = d['NB'][2] + d['RWR'][2] + d['REN'][2]
    gross = total_col * BLENDED_COMM
    after_royalty = gross * (1 - ROYALTY)
    safe_net = after_royalty * (1 - OVERHEAD)
    cap = safe_net * CAP_STANDARD

    abel_steps = [
        ('1. Total collected', 'NB col + RWR col + REN col', f"${d['NB'][2]:,.0f} + ${d['RWR'][2]:,.0f} + ${d['REN'][2]:,.0f}", f"${total_col:,.2f}", "Money Abel's customers actually paid in January"),
        ('2. Gross commission', 'Collected x 11% blended comm', f"${total_col:,.2f} x 0.11", f"${gross:,.2f}", "Commission earned from carriers"),
        ('3. After royalty', 'Gross x (1 - 17.5%)', f"${gross:,.2f} x 0.825", f"${after_royalty:,.2f}", "After corporate royalty"),
        ('4. SAFE NET', 'After-royalty x (1 - 40% overhead)', f"${after_royalty:,.2f} x 0.60", f"${safe_net:,.2f}", "What the agency truly has available for bonus + profit + taxes from Abel's book"),
        ('5. Review threshold (40%)', 'Safe net x 40%', f"${safe_net:,.2f} x 0.40", f"${cap:,.2f}", "Bonus targets above this number get flagged for ownership review"),
        ('6. Proposal A target paid', "(from rules)", "Abel's January NB+REN+RWR targets", "$252.00", "$252 is 36% of safe net - under the 40% threshold, no flag."),
        ('7. Proposal B target paid', "(from rules)", "Bigger because $10 NB base", "$399.46", "$399 is 57% of safe net - ABOVE 40%, flagged for review."),
    ]
    for row in abel_steps:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = SUB_FILL
        ws.row_dimensions[r].height = 32
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="KEY POINT").font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1
    keys = [
        "Safe net is NOT the agency's revenue. Revenue is the total commission ($25 in the example). Safe net is what's left after royalty and overhead reserve - that's where the bonus has to come from.",
        "The bigger the COLLECTED amount, the bigger the safe net, the more room for bonus. That is why the collected-% kicker exists - it pushes agents toward bigger down payments and PIF, which makes the safe net bigger.",
        "Under TODAY'S plan, bonuses paid total roughly 57% of safe net. That is too high. All three proposals come in below that. The plain-English message: 'today we are giving away more than half of what is left after rent and royalty. The new plans correct that.'",
    ]
    for t in keys:
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=11)
        c.alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 40
        r += 1

    set_col_widths(ws, [22, 32, 26, 22, 38, 14])


def build_minimum_requirements(wb):
    """Sheet explaining the $50k/$50k min requirement and showing pass/fail per agent-month."""
    ws = wb.create_sheet('Minimum Requirements')
    ws['A1'] = 'Minimum Requirements (NEW)'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:J1')

    ws['A2'] = (f"Replaces the current 35-policy NB+RWR minimum with a PREMIUM-based gate: "
                f"NB premium >= ${NB_MIN_PREMIUM:,}/month AND REN premium >= ${REN_MIN_PREMIUM:,}/month. "
                f"RWR has no own minimum, but RWR bonus only pays when BOTH NB and REN gates pass that month.")
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:J2')

    r = 4
    ws.cell(row=r, column=1, value='RULES').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    r += 1
    headers = ['Bonus line', 'Gate', 'If gate passes', 'If gate fails']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1
    rows = [
        ('NB bonus line', f"NB written premium >= ${NB_MIN_PREMIUM:,} this month", 'Pay the full NB target (from proposal rules + kicker)', 'NB bonus = $0 this month'),
        ('REN bonus line', f"REN written premium >= ${REN_MIN_PREMIUM:,} this month", 'Pay the full REN target', 'REN bonus = $0 this month'),
        ('RWR bonus line', 'BOTH gates pass', 'Pay the full RWR target ($2/policy x kicker)', 'RWR bonus = $0 this month - even with 50 rewrites'),
    ]
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            c.fill = SUB_FILL if r % 2 == 0 else PatternFill('solid', fgColor='FFFFFF')
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1

    # Pass/fail table per agent-month (REAL data)
    ws.cell(row=r, column=1, value='WHO QUALIFIES IN REAL DATA (Jan-Apr 2026)').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    r += 1

    qual_headers = ['Agent', 'Month', 'NB Premium', 'REN Premium', 'NB Gate (>=$50k)', 'REN Gate (>=$50k)', 'RWR Gate (both)', 'A bonus paid', 'B bonus paid', 'C bonus paid']
    for i, h in enumerate(qual_headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    for agent in AGENT_NAMES:
        for month in MONTHS:
            d = AGENTS[agent][month]
            a = calc_proposal_a(d['NB'], d['RWR'], d['REN'])
            b = calc_proposal_b(d['NB'], d['RWR'], d['REN'])
            cp = calc_proposal_c(d['NB'], d['RWR'], d['REN'])
            vals = [agent, month, d['NB'][1], d['REN'][1],
                    'PASS' if a['nb_qual'] else 'fail',
                    'PASS' if a['ren_qual'] else 'fail',
                    'PASS' if a['rwr_qual'] else 'fail',
                    a['paid_after_min'], b['paid_after_min'], cp['paid_after_min']]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (3, 4, 8, 9, 10): style_dollar(c)
                else: style_data(c)
                if i in (5, 6, 7):
                    if v == 'PASS':
                        c.fill = PatternFill('solid', fgColor='C6EFCE')
                        c.font = Font(bold=True, color='006100')
                    else:
                        c.fill = WARN_FILL
            r += 1

    # Now SWAP
    r += 2
    ws.cell(row=r, column=1, value='WHO QUALIFIES IF 50% OF RWR BECOMES REN').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    r += 1
    for i, h in enumerate(qual_headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    for agent in AGENT_NAMES:
        for month in MONTHS:
            d = swap_rwr_to_ren(AGENTS[agent][month], 0.5)
            a = calc_proposal_a(d['NB'], d['RWR'], d['REN'])
            b = calc_proposal_b(d['NB'], d['RWR'], d['REN'])
            cp = calc_proposal_c(d['NB'], d['RWR'], d['REN'])
            vals = [agent, month, d['NB'][1], d['REN'][1],
                    'PASS' if a['nb_qual'] else 'fail',
                    'PASS' if a['ren_qual'] else 'fail',
                    'PASS' if a['rwr_qual'] else 'fail',
                    a['paid_after_min'], b['paid_after_min'], cp['paid_after_min']]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (3, 4, 8, 9, 10): style_dollar(c)
                else: style_data(c)
                if i in (5, 6, 7):
                    if v == 'PASS':
                        c.fill = PatternFill('solid', fgColor='C6EFCE')
                        c.font = Font(bold=True, color='006100')
                    else:
                        c.fill = WARN_FILL
            r += 1

    r += 2
    ws.cell(row=r, column=1, value='WHAT IF WE LOWER THE THRESHOLD?').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    r += 1

    sens_headers = ['Threshold', 'NB qual months (REAL)', 'REN qual months (REAL)', 'RWR qual months (REAL)', 'NB qual months (SWAP)', 'REN qual months (SWAP)', 'RWR qual months (SWAP)', 'Comment', '', '']
    for i, h in enumerate(sens_headers, 1):
        if h: style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    for label, min_v in [('$25,000', 25000), ('$35,000', 35000), ('$50,000 (proposed)', 50000), ('$75,000', 75000)]:
        nb_real_q = ren_real_q = rwr_real_q = 0
        nb_swap_q = ren_swap_q = rwr_swap_q = 0
        for agent in AGENT_NAMES:
            for month in MONTHS:
                d = AGENTS[agent][month]
                ds = swap_rwr_to_ren(d, 0.5)
                if d['NB'][1] >= min_v: nb_real_q += 1
                if d['REN'][1] >= min_v: ren_real_q += 1
                if d['NB'][1] >= min_v and d['REN'][1] >= min_v: rwr_real_q += 1
                if ds['NB'][1] >= min_v: nb_swap_q += 1
                if ds['REN'][1] >= min_v: ren_swap_q += 1
                if ds['NB'][1] >= min_v and ds['REN'][1] >= min_v: rwr_swap_q += 1
        if min_v == 25000: comment = "Soft - most agents qualify on NB; REN still tight in REAL but loosens in SWAP."
        elif min_v == 35000: comment = "Roughly matches the old 35-policy bar. Achievable today on NB, REN remains the stretch."
        elif min_v == 50000: comment = "Proposed. NB gate clears for top performers; REN gate forces book-building."
        else: comment = "Very tight. Only one or two months qualify even in SWAP. Use only for tenured agents."
        vals = [label, f"{nb_real_q}/24", f"{ren_real_q}/24", f"{rwr_real_q}/24",
                f"{nb_swap_q}/24", f"{ren_swap_q}/24", f"{rwr_swap_q}/24", comment]
        for i, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = SUB_FILL
            if 'proposed' in str(vals[0]).lower():
                c.fill = PROP_A_FILL
        ws.row_dimensions[r].height = 30
        r += 1

    set_col_widths(ws, [24, 12, 14, 14, 16, 16, 16, 14, 14, 14])


def build_previous_vs_new(wb):
    """Agent-by-agent dollar comparison: Previous structure (current) vs Proposal A WITH minimum."""
    ws = wb.create_sheet('Previous vs New Detailed')
    ws['A1'] = 'Previous Structure vs New Structure - What Each Agent Would Have Earned'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:M1')

    ws['A2'] = ("Side-by-side dollar comparison for each agent, each month. 'PREVIOUS' is today's count-tier plan. "
                "'NEW (no min)' is the proposal target. 'NEW (with min)' applies the $50k NB / $50k REN gates.")
    ws['A2'].font = Font(italic=True, size=10, color='666666')
    ws.merge_cells('A2:M2')

    r = 4
    for agent in AGENT_NAMES:
        ws.cell(row=r, column=1, value=agent).font = SECTION_FONT
        ws.cell(row=r, column=1).fill = SECTION_FILL
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
        r += 1

        headers = ['Month', 'NB cnt', 'NB Prem', 'REN cnt', 'REN Prem', 'RWR cnt',
                   'Previous (current)', 'A no-min', 'A WITH MIN', 'A vs Previous',
                   'B WITH MIN', 'C WITH MIN', 'Gates']
        for i, h in enumerate(headers, 1):
            style_header(ws.cell(row=r, column=i, value=h))
        r += 1

        prev_total = a_nomin = a_min = b_min = c_min = 0
        for month in MONTHS:
            d = AGENTS[agent][month]
            cur = current_bonus(d['NB'][0], d['RWR'][0])
            a = calc_proposal_a(d['NB'], d['RWR'], d['REN'])
            b = calc_proposal_b(d['NB'], d['RWR'], d['REN'])
            cp = calc_proposal_c(d['NB'], d['RWR'], d['REN'])
            gates = f"NB:{'Y' if a['nb_qual'] else 'n'} REN:{'Y' if a['ren_qual'] else 'n'} RWR:{'Y' if a['rwr_qual'] else 'n'}"
            vals = [month, d['NB'][0], d['NB'][1], d['REN'][0], d['REN'][1], d['RWR'][0],
                    cur, a['paid'], a['paid_after_min'], a['paid_after_min'] - cur,
                    b['paid_after_min'], cp['paid_after_min'], gates]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (2, 4, 6): style_int(c)
                elif i in (3, 5): style_dollar(c)
                elif i == 13: style_data(c)
                else: style_dollar(c)
                if i == 7: c.fill = CURRENT_FILL
                if i in (8, 9): c.fill = PROP_A_FILL
                if i == 9: c.font = Font(bold=True)
                if i == 11: c.fill = PROP_B_FILL; c.font = Font(bold=True)
                if i == 12: c.fill = PROP_C_FILL; c.font = Font(bold=True)
            prev_total += cur
            a_nomin += a['paid']
            a_min += a['paid_after_min']
            b_min += b['paid_after_min']
            c_min += cp['paid_after_min']
            r += 1

        ws.cell(row=r, column=1, value='4-Month Total').font = Font(bold=True)
        ws.cell(row=r, column=7, value=prev_total)
        ws.cell(row=r, column=8, value=a_nomin)
        ws.cell(row=r, column=9, value=a_min)
        ws.cell(row=r, column=10, value=a_min - prev_total)
        ws.cell(row=r, column=11, value=b_min)
        ws.cell(row=r, column=12, value=c_min)
        for i in range(1, 14):
            c = ws.cell(row=r, column=i)
            c.fill = SUB_FILL
            if i in (7, 8, 9, 10, 11, 12):
                style_dollar(c)
                c.font = Font(bold=True)
        r += 1

        # Plain-English read for this agent
        ws.cell(row=r, column=1, value=f"Read: under PREVIOUS plan {agent.split()[0]} earned ${prev_total:,.0f} in 4 months. Under NEW Proposal A WITHOUT minimums it would be ${a_nomin:,.0f}. WITH $50k/$50k minimums it drops to ${a_min:,.0f} because the REN gate barely opens.").alignment = Alignment(wrap_text=True)
        ws.cell(row=r, column=1).font = Font(italic=True, size=10, color='666666')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
        ws.row_dimensions[r].height = 30
        r += 2

    set_col_widths(ws, [12, 8, 11, 8, 11, 8, 13, 12, 13, 14, 13, 13, 20])


def build_rules_side_by_side(wb):
    ws = wb.create_sheet('Rules Side by Side')
    ws['A1'] = 'Rules: Current vs Proposal A vs Proposal B vs Proposal C'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:E1')

    ws['A2'] = 'One-row-per-rule view of all four plans. Use this to compare. A and B are the main proposals (do NOT mix them).'
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:E2')

    r = 4
    headers = ['Rule', 'Current', 'Proposal A (Premium)', 'Proposal B (Coverage)', 'Proposal C (Persistency)']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    # Recompute totals so this table always matches the engine
    real = {'cur': 0, 'a': 0, 'a_min': 0, 'b': 0, 'b_min': 0, 'c': 0, 'c_min': 0}
    swap = {'cur': 0, 'a': 0, 'a_min': 0, 'b': 0, 'b_min': 0, 'c': 0, 'c_min': 0}
    for agent in AGENT_NAMES:
        for month in MONTHS:
            d = AGENTS[agent][month]
            ds = swap_rwr_to_ren(d, 0.5)
            real['cur'] += current_bonus(d['NB'][0], d['RWR'][0])
            swap['cur'] += current_bonus(ds['NB'][0], ds['RWR'][0])
            a = calc_proposal_a(d['NB'], d['RWR'], d['REN'])
            real['a'] += a['paid']; real['a_min'] += a['paid_after_min']
            as_ = calc_proposal_a(ds['NB'], ds['RWR'], ds['REN'])
            swap['a'] += as_['paid']; swap['a_min'] += as_['paid_after_min']
            b = calc_proposal_b(d['NB'], d['RWR'], d['REN'])
            real['b'] += b['paid']; real['b_min'] += b['paid_after_min']
            bs = calc_proposal_b(ds['NB'], ds['RWR'], ds['REN'])
            swap['b'] += bs['paid']; swap['b_min'] += bs['paid_after_min']
            cp = calc_proposal_c(d['NB'], d['RWR'], d['REN'])
            real['c'] += cp['paid']; real['c_min'] += cp['paid_after_min']
            cs = calc_proposal_c(ds['NB'], ds['RWR'], ds['REN'])
            swap['c'] += cs['paid']; swap['c_min'] += cs['paid_after_min']

    rows = [
        ('Bonus base', 'Count tier (NB+RWR)', 'Written premium tier', 'Coverage type', 'Flat per policy'),
        ('NB - low premium (< $1,200)', 'Count toward tier', '$5', '$10 (base coverage)', '$8'),
        ('NB - medium premium', 'Count toward tier', '$7-$11 (tier)', '$10 (base) + $5 if BI+UM', '$8'),
        ('NB - high premium ($3,000+)', 'Count toward tier', '$11 + $2/$1k, cap $25', '$10 (base) + $5 if BI+UM', '$8'),
        ('REN - any', 'PAYS NOTHING', '$4 / $5 / $6 by premium tier', '$6 (base) + $3 if BI+UM', '$8 (PARITY with NB)'),
        ('RWR - any', 'Count toward tier (BAD)', '$2 flat', '$2 flat', '$1 flat (LOWER)'),
        ('Coverage detail', 'Not tracked', 'Not paid separately', 'BI+UM bundle = +$5 NB / +$3 REN', 'Not tracked'),
        ('PIF add', 'Not paid', '+$8 NB / +$12 high NB / +$5 REN', '+$8 NB / +$12 high NB / +$5 REN', 'No (kicker only)'),
        ('Collected % kicker', 'NONE', '+10% / +15% / +20% / +25%', 'Same as A', 'Same as A'),
        ('Monthly minimum (NEW)', '35 policies NB+RWR', f'NB >= ${NB_MIN_PREMIUM/1000:.0f}k AND REN >= ${REN_MIN_PREMIUM/1000:.0f}k', 'Same as A', 'Same as A'),
        ('RWR gate', 'No - RWR counts toward 35', 'RWR paid only if BOTH NB+REN gates pass', 'Same as A', 'Same as A'),
        ('Profit protection', 'None per-policy', '40% safe net review + 90-day chargeback + minimums', 'Same as A', 'Same as A'),
        ('Best for', 'Status quo only', 'Simple payroll, premium-focused', 'Coverage upsell, agent motivation', 'Killing the rewrite habit'),
        ('Tradeoff', 'Rewards churning', 'Less generous than B', 'More generous, requires coverage tracking', 'No premium/coverage signal'),
        ('Recommended decision', 'Phase out', 'Conservative choice', 'PILOT (recommended)', 'If A or B rejected'),
        ('4-month cost - REAL, no min', f"${real['cur']:,.0f}", f"${real['a']:,.0f}", f"${real['b']:,.0f}", f"${real['c']:,.0f}"),
        ('4-month cost - REAL, WITH min', f"${real['cur']:,.0f}", f"${real['a_min']:,.0f}", f"${real['b_min']:,.0f}", f"${real['c_min']:,.0f}"),
        ('4-month cost - SWAP, no min', f"${swap['cur']:,.0f}", f"${swap['a']:,.0f}", f"${swap['b']:,.0f}", f"${swap['c']:,.0f}"),
        ('4-month cost - SWAP, WITH min', f"${swap['cur']:,.0f}", f"${swap['a_min']:,.0f}", f"${swap['b_min']:,.0f}", f"${swap['c_min']:,.0f}"),
    ]
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if i == 2: c.fill = CURRENT_FILL
            elif i == 3: c.fill = PROP_A_FILL
            elif i == 4: c.fill = PROP_B_FILL
            elif i == 5: c.fill = PROP_C_FILL
        ws.row_dimensions[r].height = 30
        r += 1

    set_col_widths(ws, [32, 26, 30, 32, 28])


def build_agent_quick_ref(wb):
    ws = wb.create_sheet('Agent Quick Reference')
    ws['A1'] = 'Agent Quick Reference - How You Earn Bonus'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:D1')

    ws['A2'] = 'Pin this to your desk. The plan ownership picks (A, B, or C) will be highlighted on the day the new plan goes live.'
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:D2')

    r = 4
    ws.cell(row=r, column=1, value='WHAT GETS BONUSED').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    r += 1
    rules = [
        ('1. Renewals NOW PAY.', 'Under all three new plans, every renewal you keep on the books earns bonus. It used to pay $0.'),
        ('2. Rewrites pay LESS.', 'Rewrites pay $1-$2 flat. They do NOT push you into a higher bonus tier anymore.'),
        ('3. Collected money matters.', 'The more of the down payment you collect, the bigger your kicker (+10% to +25%).'),
        ('4. PIF policies pay extra.', 'If the customer pays in full, you get +$8 NB / +$12 high NB / +$5 REN on top.'),
        ('5. BI+UM bundle pays extra (Proposal B).', 'If your NB has BI AND UM together, +$5 added. UM alone is not possible; BI alone earns base only.'),
        ('6. 90-day chargeback.', 'If a policy you got paid on cancels or rewrites within 90 days, the bonus is reversed.'),
    ]
    for label, txt in rules:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True, size=11)
        c = ws.cell(row=r, column=2, value=txt)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    ws.cell(row=r, column=1, value='HOW TO MAXIMIZE YOUR BONUS').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    r += 1
    maxi = [
        ('Push for the bigger down payment.', 'Going from 15% to 25% collected adds +5% to your kicker. Going to PIF adds +25%.'),
        ('Save the customer instead of rewriting them.', 'A renewal pays $4-$8. A rewrite pays $1-$2. Big difference over a year.'),
        ('Sell the BI+UM bundle on NB (Proposal B).', 'Extra +$5 per policy. Track it on the monthly tracker - if it is not documented, you do not get paid for it.'),
        ('Mention PIF.', 'Customers who pay in full earn YOU more bonus AND give the agency more cash up front. Win-win.'),
        ('Avoid same-90-day rewrites.', 'They reverse the bonus you already earned. Fix the policy without rewriting if you can.'),
    ]
    for label, txt in maxi:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True, size=11)
        c = ws.cell(row=r, column=2, value=txt)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    ws.cell(row=r, column=1, value='QUICK EXAMPLES').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    r += 1
    ex_headers = ['Scenario', 'Proposal A pays', 'Proposal B pays', 'Proposal C pays']
    for i, h in enumerate(ex_headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    examples = [
        ('NB $1,000 PIP+PD, 25% down', '$5 x 1.15 = $5.75', '$10 x 1.15 = $11.50', '$8 x 1.15 = $9.20'),
        ('NB $2,000 with BI+UM, 25% down', '$9 x 1.15 = $10.35', '($10 + $5) x 1.15 = $17.25', '$8 x 1.15 = $9.20'),
        ('NB $2,500 PIF with BI+UM', '($11 + $8 PIF) x 1.25 = $23.75', '($10 + $5 + $8 PIF) x 1.25 = $28.75', '$8 x 1.25 = $10.00'),
        ('Renewal $1,500 with BI+UM, 25%', '$5 x 1.15 = $5.75', '($6 + $3) x 1.15 = $10.35', '$8 x 1.15 = $9.20'),
        ('Rewrite $1,200', '$2 x 1.15 = $2.30', '$2 x 1.15 = $2.30', '$1 x 1.15 = $1.15'),
    ]
    for ex in examples:
        for i, v in enumerate(ex, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if i == 1: c.fill = SUB_FILL
            elif i == 2: c.fill = PROP_A_FILL
            elif i == 3: c.fill = PROP_B_FILL
            elif i == 4: c.fill = PROP_C_FILL
        ws.row_dimensions[r].height = 28
        r += 1

    set_col_widths(ws, [34, 30, 32, 28])


def build_assumptions(wb):
    ws = wb.create_sheet('Assumptions')
    ws['A1'] = 'Assumptions and Manual Tracking'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:D1')

    ws['A2'] = 'All economic inputs used in this workbook. Changing any of these requires re-running the examples.'
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:D2')

    r = 4
    headers = ['Assumption', 'Value', 'Meaning', 'Source / Note']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    rows = [
        ('Royalty rate', '17.5%', 'Franchise royalty removed before safe net.', 'Standard Fiesta franchise royalty.'),
        ('Overhead reserve', '40%', 'Rent, payroll, tech, admin, bank fees, etc.', 'Conservative monthly fixed-cost reserve.'),
        ('Blended carrier commission', '11%', 'Used for the safe-net calculation in the aggregate model.', 'Range across carriers: 8% (Bristol West) to 15% (Geico/Kemper). 11% is the book-weighted blend.'),
        ('Bonus paid', 'TARGET (no automatic cap)', 'Pay equals the calculated target. Cap is a soft review threshold only.', 'Same for A, B, C.'),
        ('Review threshold (standard)', '40% of safe net', 'If monthly target > 40% safe net, ownership reviews. Most months stay below.', 'Same for A, B, C.'),
        ('Review threshold (PIF)', '50% of safe net', 'Higher because PIF policies collect 100% upfront - less risk.', 'Same for A, B, C.'),
        ('Chargeback period (PRIMARY PROTECTION)', '90 days', 'Cancel/rewrite reversal window. This is the actual profit shield.', 'Matches carrier commission chargeback exposure.'),
        ('Proposal A NB tiers', '$5 / $7 / $9 / $11 / $11+$2/$1k cap $25', 'NB pay by written premium.', 'See Proposal A sheet for full table.'),
        ('Proposal A REN tiers', '$4 / $5 / $6', 'REN pay by written premium.', 'See Proposal A sheet.'),
        ('Proposal A RWR', '$2 flat', 'No tiers, no kicker stacking.', 'Same as Proposal B.'),
        ('Proposal B NB base', '$10 per policy', 'PIP/PD or PIP+Comp/Coll.', 'See Coverage Bundle Logic.'),
        ('Proposal B NB liability add', '+$5 per policy with BI+UM', 'Bundled - UM cannot exist without BI.', 'BI alone does not qualify.'),
        ('Proposal B REN base', '$6 per policy', 'Same coverage logic, smaller dollar.', '-'),
        ('Proposal B REN liability add', '+$3 per policy with BI+UM', 'Same bundle logic on REN.', '-'),
        ('Proposal C NB', '$8 flat', 'No tiers, no coverage detail.', 'Simplest plan.'),
        ('Proposal C REN', '$8 flat (= NB)', 'REN parity is the design choice.', '-'),
        ('Proposal C RWR', '$1 flat', 'Active disincentive to rewriting.', '-'),
        ('Collected kicker tiers', '15-24% / 25-49% / 50-99% / 100%', '+10% / +15% / +20% / +25%', 'Same across A, B, C.'),
        ('Liability adoption (NB)', '35% of NB', 'Estimated share of NB policies that carry BI+UM.', 'Used in Proposal B aggregate modeling only. Actual bonus is paid per verified policy.'),
        ('Liability adoption (REN)', '30% of REN', 'Estimated share of REN policies that carry BI+UM.', 'Same as above.'),
        ('Source data', 'agents_nbrwr_pivots_good_.xlsx', 'NB / RWR / REN counts and premiums for Jan-Apr 2026.', '6 agents modeled: Abel, Dialinerys, Melissa, Flavia, Thalia, Monica.'),
        ('Swap scenario', '50% RWR to REN', 'Reclassifies half of each agent\'s rewrites as renewals.', 'Premium and collected $ stay the same. Shows future-state earnings.'),
    ]
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = SUB_FILL
        ws.row_dimensions[r].height = 26
        r += 1

    set_col_widths(ws, [28, 24, 50, 60])


# ============================================================================
# BUILD WORKBOOK
# ============================================================================
def build():
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    build_readme(wb)
    build_executive_summary(wb)
    build_safe_net_simple(wb)
    build_minimum_requirements(wb)
    build_previous_vs_new(wb)
    build_rules_side_by_side(wb)
    build_proposal_a(wb)
    build_proposal_b(wb)
    build_proposal_c(wb)
    build_worked_examples(wb)
    build_agent_examples(wb, swap=False)
    build_agent_examples(wb, swap=True)
    build_comparison(wb)
    build_coverage_logic(wb)
    build_kicker_logic(wb)
    build_profitability(wb)
    build_chargeback(wb)
    build_agent_quick_ref(wb)
    build_manual_tracker(wb)
    build_assumptions(wb)

    out = '/home/user/fiesta-bonus/output/Bonus_Proposal_FIXED_Premium_vs_Coverage.xlsx'
    wb.save(out)
    print(f"Saved: {out}")
    return out


if __name__ == '__main__':
    build()
