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
from openpyxl.formatting.rule import CellIsRule, DataBarRule
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
BLENDED_COMM = 0.11      # blended carrier commission rate (8-15% across carriers)
ROYALTY = 0.21           # franchise royalty (corrected from 17.5%)
OVERHEAD = 0.30          # non-salary overhead reserve (rent, tech, admin, bank fees)
                         # Salaries are tracked separately via the minimum requirement.
CAP_STANDARD = 0.40      # ownership REVIEW threshold (informational, not auto-cap)
CAP_PIF = 0.55           # PIF review threshold

# Estimated coverage adoption for Proposal B (since coverage data is not in source pivots)
LIAB_ADOPTION_NB = 0.35
LIAB_ADOPTION_REN = 0.30

# Salary basis: $16-22/hour x 40 hours/week = $2,773-$3,813/mo. Midpoint ~$3,300.
AGENT_SALARY = 3300

# Minimum-requirement thresholds (monthly)
NB_MIN_PREMIUM = 45000      # NB minimum: $45,000 written premium per month
REN_MIN_RETENTION = 0.30    # REN minimum: retain 30% of agent's book (retention-based, not $)
                            # RWR minimum: meet NB min (RWR is paid alongside NB)
REN_MIN_PREMIUM = 0         # deprecated - kept for backward-compat with old text references

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


def col_inc_per_policy(pct):
    """Flat dollar per-policy collected incentive (only for policies > $1,200 premium).
    Tiers (FINAL):
       15-19% collected: +$1
       20-24% collected: +$2
       25-99% collected: +$5
       100% PIF:         +$8
    """
    if pct >= 1.00: return 8   # PIF
    if pct >= 0.25: return 5   # 25%+ (and not PIF)
    if pct >= 0.20: return 2   # 20-24%
    if pct >= 0.15: return 1   # 15-19%
    return 0                    # below 15%


def pct_above_1200(count, prem):
    """Estimate share of policies with premium > $1,200 (used since we don't have per-policy data)."""
    if count == 0: return 0
    avg = prem / count
    if avg < 600: return 0.0
    if avg > 2400: return 1.0
    return (avg - 600) / 1800


# Retention bonus parameters
RETENTION_POOL_RATE = 0.005  # 0.5% of REN written premium - the ONLY renewal pay
                              # = $5 per $1,000 retained
                              # Half the rate of the highest per-policy NB tier so renewals
                              # don't out-pay new business.
RETENTION_RATE_DEFAULT = 0.75  # placeholder retention rate (actual computed from agency history)


def apply_minimums(nb_premium, ren_premium, nb_target, ren_target, rwr_target,
                   nb_min=NB_MIN_PREMIUM, retention_rate=RETENTION_RATE_DEFAULT):
    """Gate the bonus lines by the minimum thresholds.

    FINAL RULES:
      - NB minimum: NB written premium >= $45,000/month -> NB bonus paid
      - REN minimum: retain >= 30% of agent's book (RETENTION metric, not $) -> REN bonus paid
      - RWR minimum: NB minimum met (RWR is paid alongside NB).
        "Meet NB min or both" - meeting NB alone is enough; meeting both also works.

    For modeling we use the assumed retention_rate vs the 30% threshold.
    """
    nb_qual = nb_premium >= nb_min
    ren_qual = retention_rate >= REN_MIN_RETENTION
    rwr_qual = nb_qual  # "meet NB min or both" - NB alone is enough
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
    """Per-policy NB base bonus by written premium tier."""
    if avg_prem < 1200: return 5,  "Under $1,200"
    if avg_prem < 1800: return 7,  "$1,200-$1,799"
    if avg_prem < 2500: return 10, "$1,800-$2,499"
    if avg_prem < 3000: return 13, "$2,500-$2,999"
    val = min(13 + (avg_prem - 3000) / 1000 * 2, 25)
    return val, "$3,000+"


def ren_premium_tier(avg_prem):
    """Per-policy REN base bonus by written premium tier."""
    if avg_prem < 1200: return 4, "Under $1,200"
    if avg_prem < 1800: return 6, "$1,200-$1,799"
    if avg_prem < 2500: return 8, "$1,800-$2,499"
    return 10, "$2,500+"


def rwr_premium_tier(avg_prem):
    """Per-policy RWR base bonus by written premium tier (low - rewrites are deprioritized)."""
    if avg_prem < 1200: return 2, "Under $1,200"
    if avg_prem < 1800: return 3, "$1,200-$1,799"
    return 4, "$1,800+"


# ---------------------------------------------------------------------------
# NEW TIERED-PREMIUM STRUCTURE (replaces per-policy base + collected kicker)
# ---------------------------------------------------------------------------
# Mirrors today's count-tier idea (30 policies = $250, 38 = $350, 50 = $450,
# $10/policy after) but uses monthly WRITTEN PREMIUM instead of policy count
# and scales the reward toward $1,000 at $100k written premium.
#
# NB:   T1 $45k = $250 | T2 $55k = $375 | T3 $70k = $525 | T4 $85k = $725 | T5 $100k = $1,000
#       Above $100k: +$50 per $5k (1% rate)
# RWR:  same breakpoints, pays HALF of NB amounts
#       Above $100k: +$25 per $5k (0.5% rate)
# REN:  own breakpoints starting at $25k, pays 0.80x of NB amounts
#       T1 $25k = $200 | T2 $45k = $300 | T3 $65k = $420 | T4 $85k = $580 | T5 $100k = $800
#       Above $100k: +$40 per $5k (0.8% rate)

NB_TIERS = [
    (45000, 250),
    (55000, 375),
    (70000, 525),
    (85000, 725),
    (100000, 1000),
]
NB_OVERAGE_PER_5K = 50    # +$50 per $5k written premium above $100k

RWR_TIERS = [
    (45000, 100),
    (55000, 200),
    (70000, 275),
    (85000, 350),
    (100000, 500),
]
RWR_OVERAGE_PER_5K = 25   # +$25 per $5k above $100k

REN_TIERS = [
    (25000, 250),
    (40000, 350),
    (65000, 450),
    (80000, 550),
    (100000, 800),
]
REN_OVERAGE_PER_5K = 40   # +$40 per $5k above $100k


def tier_bonus(premium, tiers, overage_per_5k):
    """Find the bonus for a given monthly premium under a tier ladder.

    Below the lowest breakpoint = $0.
    At or above a breakpoint = that tier's amount.
    Above the top breakpoint = top amount + linear overage per $5k.
    """
    if premium < tiers[0][0]:
        return 0.0, 0, "Below T1"
    bonus = 0.0
    tier_num = 0
    label = "Below T1"
    for i, (breakpoint, amount) in enumerate(tiers, 1):
        if premium >= breakpoint:
            bonus = amount
            tier_num = i
            label = f"T{i} ${breakpoint/1000:.0f}k"
    top_break = tiers[-1][0]
    if premium > top_break:
        extra_5k_blocks = (premium - top_break) / 5000
        bonus += extra_5k_blocks * overage_per_5k
        label = f"T5+ (${premium/1000:.0f}k)"
    return bonus, tier_num, label


def nb_tier_bonus(nb_premium):
    return tier_bonus(nb_premium, NB_TIERS, NB_OVERAGE_PER_5K)


def rwr_tier_bonus(rwr_premium):
    return tier_bonus(rwr_premium, RWR_TIERS, RWR_OVERAGE_PER_5K)


def ren_tier_bonus(ren_premium):
    return tier_bonus(ren_premium, REN_TIERS, REN_OVERAGE_PER_5K)


def calc_proposal_a(nb, rwr, ren, retention_rate=RETENTION_RATE_DEFAULT):
    """THE BONUS PLAN - monthly premium-tier structure (replaces count tiers).

    NB:   5-tier ladder by monthly NB written premium ($45k-$100k), $250-$1,000.
          Above $100k: +$50 per $5k.
          Gate: NB premium >= $45k.
    RWR:  Same tier ladder, pays HALF of NB amounts ($125-$500).
          Above $100k: +$25 per $5k.
          Gate: NB premium >= $45k (RWR has no separate floor).
    REN:  Own ladder starting at $25k ($200-$800), pays ~0.80x of NB.
          Above $100k: +$40 per $5k.
          Gate: retention rate >= 30% of book.

    Chargeback: 3 months.
    Manual tracker entry required for every policy.
    """
    nb_c, nb_p, nb_col = nb
    rwr_c, rwr_p, rwr_col = rwr
    ren_c, ren_p, ren_col = ren

    nb_bonus, nb_tier_num, nb_tier_label = nb_tier_bonus(nb_p)
    rwr_bonus, rwr_tier_num, rwr_tier_label = rwr_tier_bonus(rwr_p)
    ren_bonus, ren_tier_num, ren_tier_label = ren_tier_bonus(ren_p)

    nb_target = nb_bonus
    rwr_target = rwr_bonus
    ren_target = ren_bonus

    total_target = nb_target + rwr_target + ren_target

    total_collected = nb_col + rwr_col + ren_col
    gross_comm = total_collected * BLENDED_COMM
    safe_net = gross_comm * (1 - ROYALTY) * (1 - OVERHEAD)
    cap = safe_net * CAP_STANDARD
    paid = total_target
    capped_paid = min(total_target, cap)
    bonus_pct_of_safe_net = paid / safe_net if safe_net else 0
    review_flag = (paid > cap)

    # Gates
    nb_qual = nb_p >= NB_MIN_PREMIUM
    ren_qual = retention_rate >= REN_MIN_RETENTION
    rwr_qual = nb_qual  # RWR follows NB gate

    paid_nb = nb_target if nb_qual else 0.0
    paid_ren = ren_target if ren_qual else 0.0
    paid_rwr = rwr_target if rwr_qual else 0.0
    paid_after_min = paid_nb + paid_ren + paid_rwr

    avg_nb = nb_p / nb_c if nb_c else 0
    avg_rwr = rwr_p / rwr_c if rwr_c else 0
    avg_ren = ren_p / ren_c if ren_c else 0
    nb_col_pct = nb_col / nb_p if nb_p else 0
    rwr_col_pct = rwr_col / rwr_p if rwr_p else 0
    ren_col_pct = ren_col / ren_p if ren_p else 0

    return {
        'nb_target': nb_target, 'ren_target': ren_target, 'rwr_target': rwr_target,
        # Legacy keys kept at 0 so old display code doesn't break.
        'nb_base_pay': nb_bonus, 'nb_col_pay': 0.0,
        'rwr_base_pay': rwr_bonus, 'rwr_col_pay': 0.0,
        'ren_base_pay': ren_bonus, 'ren_col_pay': 0.0,
        'nb_per': nb_bonus, 'nb_tier_label': nb_tier_label,
        'rwr_per': rwr_bonus, 'rwr_tier_label': rwr_tier_label,
        'ren_per': ren_bonus, 'ren_tier_label': ren_tier_label,
        'nb_tier_num': nb_tier_num, 'rwr_tier_num': rwr_tier_num, 'ren_tier_num': ren_tier_num,
        'avg_nb': avg_nb, 'avg_rwr': avg_rwr, 'avg_ren': avg_ren,
        'nb_above_1200': 0.0, 'rwr_above_1200': 0.0, 'ren_above_1200': 0.0,
        'nb_inc_per_policy': 0.0, 'rwr_inc_per_policy': 0.0, 'ren_inc_per_policy': 0.0,
        'retention_bonus': ren_bonus, 'retention_rate': retention_rate,
        'total_target': total_target,
        'nb_col_pct': nb_col_pct, 'ren_col_pct': ren_col_pct, 'rwr_col_pct': rwr_col_pct,
        'gross_comm': gross_comm, 'safe_net': safe_net, 'cap': cap,
        'paid': paid, 'capped_paid': capped_paid, 'paid_after_min': paid_after_min,
        'bonus_pct_of_safe_net': bonus_pct_of_safe_net, 'review_flag': review_flag,
        'nb_qual': nb_qual, 'ren_qual': ren_qual, 'rwr_qual': rwr_qual,
        'paid_nb_after_min': paid_nb,
        'paid_ren_after_min': paid_ren,
        'paid_rwr_after_min': paid_rwr,
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

    nb_base = nb_c * 11
    nb_liability = nb_c * LIAB_ADOPTION_NB * 6
    nb_target = (nb_base + nb_liability) * kicker(nb_col_pct)

    ren_base = ren_c * 7
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


def full_flip(month_data):
    """Flip RWR and REN entirely - the 'if all your rewrites had been renewals' scenario."""
    return {
        'NB': month_data['NB'],
        'RWR': month_data['REN'],  # what used to be REN is now RWR
        'REN': month_data['RWR'],  # what used to be RWR is now REN
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
    ws = wb.create_sheet('README - START HERE')
    ws['A1'] = 'Fiesta Bonus Plan Proposal - START HERE'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:F1')

    rows = [
        ('', ''),
        ('===== TL;DR =====', ''),
        ('The problem', "Today's bonus rewards REWRITING customers (counts NB+RWR policies). It pays $0 for renewals - so there's no incentive to keep the book."),
        ('The new plan', 'Same shape as today (5 tiers) but driven by monthly WRITTEN PREMIUM, not policy count. Each line - NB, RWR, REN - has its own tier ladder.'),
        ('Bottom line (4-mo, 6 agents)', 'TODAY: $12,220. NEW PLAN: $1,675 if behavior stays the same, $3,875 with 50% renewal shift, $7,975 if all rewrites had been renewals. New plan costs LESS while paying for the right behavior.'),
        ('Recommendation', f'Adopt the new plan with the ${NB_MIN_PREMIUM:,} NB minimum + {REN_MIN_RETENTION*100:.0f}% retention minimum + 90-day chargeback. Pilot 90 days side-by-side, pay the higher of the two.'),
        ('', ''),
        ('===== THE PLAN AT A GLANCE =====', ''),
        ('NB', 'T1 $45k=$250 | T2 $55k=$375 | T3 $70k=$525 | T4 $85k=$725 | T5 $100k=$1,000. Above $100k: +$50 per $5k.'),
        ('RWR', 'T1 $45k=$100 | T2 $55k=$200 | T3 $70k=$275 | T4 $85k=$350 | T5 $100k=$500. Above $100k: +$25 per $5k.'),
        ('REN', 'T1 $25k=$250 | T2 $40k=$350 | T3 $65k=$450 | T4 $80k=$550 | T5 $100k=$800. Above $100k: +$40 per $5k.'),
        ('Minimum requirements', f'NB premium >= ${NB_MIN_PREMIUM:,} | Retention rate >= {REN_MIN_RETENTION*100:.0f}% | RWR follows NB gate.'),
        ('90-day chargeback', '100% of paid bonus is reversed if the policy cancels or rewrites within 90 days of effective date.'),
        ('Manual tracker', 'Required for every policy. No tracker entry = no bonus on that policy.'),
        ('', ''),
        ('===== GLOSSARY =====', ''),
        ('NB', 'NEW BUSINESS - first time the customer is on the books.'),
        ('RWR', 'REWRITE - cancel an existing policy, write a new one. Churns the book.'),
        ('REN', 'RENEWAL - existing policy renews at the same carrier. Customer stays.'),
        ('Retention rate', 'Of the policies up for renewal this month, the % that renewed.'),
        ('Written premium', 'Total cost of the policy for the 6-month term.'),
        ('FLIP scenario', '"What if all rewrites had been renewals?" Strongest behavior-change test.'),
        ('SWAP scenario', '"What if half the rewrites became renewals?" Realistic transition target.'),
        ('', ''),
        ('===== SHEET MAP =====', ''),
        ('  1. Executive Summary', 'Bottom-line numbers and per-agent rows.'),
        ('  2. The Bonus Plan', 'Full plan rules: tier ladders, minimums, chargeback, worked example.'),
        ('  3. NB Bonus Examples', 'Who hits which NB tier under each scenario. Sorted by premium with a progress bar.'),
        ('  4. RWR Bonus Examples', 'Same view for the RWR ladder.'),
        ('  5. REN Bonus Examples', 'Same view for the REN ladder.'),
        ('  6. Agent Examples - 100% Flip', 'Per-agent month-by-month under the strongest behavior shift.'),
        ('  7. Chargeback Process', 'Step-by-step 90-day reversal procedure.'),
        ('  8. Manual Tracker Template', 'Monthly template for verified bonus tracking.'),
        ('  9. Source Data', 'Raw NB/RWR/REN volumes from Jan-Apr 2026.'),
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
    ws['A1'] = 'Executive Summary - Current Plan vs The Bonus Plan'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:M1')

    ws['A2'] = (f'Six agents (Abel, Dialinerys, Melissa, Flavia, Thalia, Monica) over four months (Jan-Apr 2026). '
                f'Each plan is shown two ways: WITHOUT minimums (calculated target) and WITH the recommended minimums '
                f'(NB premium >= ${NB_MIN_PREMIUM:,}/mo; REN gated by retention (model assumes 75% >= 30% threshold); RWR gated by NB).')
    ws['A2'].font = Font(italic=True, size=10, color='666666')
    ws.merge_cells('A2:M2')

    # Pre-compute totals for headline - REAL, 50% SWAP, 100% FLIP
    head_real = {'cur': 0, 'a': 0, 'a_min': 0}
    head_swap = {'cur': 0, 'a': 0, 'a_min': 0}
    head_flip = {'cur': 0, 'a': 0, 'a_min': 0}
    for agent in AGENT_NAMES:
        for month in MONTHS:
            real = AGENTS[agent][month]
            swap = swap_rwr_to_ren(real, 0.5)
            flip = full_flip(real)
            head_real['cur'] += current_bonus(real['NB'][0], real['RWR'][0])
            head_swap['cur'] += current_bonus(swap['NB'][0], swap['RWR'][0])
            head_flip['cur'] += current_bonus(flip['NB'][0], flip['RWR'][0])
            ra = calc_proposal_a(real['NB'], real['RWR'], real['REN'])
            sa = calc_proposal_a(swap['NB'], swap['RWR'], swap['REN'])
            fa = calc_proposal_a(flip['NB'], flip['RWR'], flip['REN'])
            head_real['a']  += ra['paid'];  head_real['a_min']  += ra['paid_after_min']
            head_swap['a']  += sa['paid'];  head_swap['a_min']  += sa['paid_after_min']
            head_flip['a']  += fa['paid'];  head_flip['a_min']  += fa['paid_after_min']

    # HEADLINE block - three scenarios side-by-side
    r = 4
    ws.cell(row=r, column=1, value='HEADLINE - 4-MONTH TOTALS ACROSS 3 SCENARIOS (6 agents)').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    r += 1

    hh = ['Plan', 'REAL (today)', '50% SWAP (RWR->REN)', '100% FLIP (RWR<->REN)', 'Read']
    for i, h in enumerate(hh, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    head_rows = [
        ('Current plan (today)', head_real['cur'], head_swap['cur'], head_flip['cur'], 'DROPS hard - the current plan only counts NB+RWR, so any shift toward REN cuts the bonus.'),
        ('The Bonus Plan - no minimum', head_real['a'], head_swap['a'], head_flip['a'], 'The calculated target before applying minimums.'),
        (f'The Bonus Plan WITH MIN (RECOMMENDED)', head_real['a_min'], head_swap['a_min'], head_flip['a_min'], f'${NB_MIN_PREMIUM:,} NB premium + {REN_MIN_RETENTION*100:.0f}% retention gates. RWR follows NB gate.'),
    ]
    for row in head_rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            if i == 1:
                style_data(c)
                c.font = Font(bold=True)
            elif i in (2, 3, 4):
                style_dollar(c)
            else:
                style_data(c)
            if 'Current' in row[0]: c.fill = CURRENT_FILL
            else: c.fill = PROP_A_FILL
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    ws.cell(row=r, column=1, value=(
        "READ: Current plan drops in swap because it only rewards NB+RWR count. New plans grow in swap because they pay for renewals. "
        "WITH MIN totals are smaller because no agent has $50k of REN premium today - until the renewal book is built, the gate stays closed. "
        "That is the design: 'no bonus until you cover yourself.'"
    )).alignment = Alignment(wrap_text=True)
    ws.cell(row=r, column=1).font = Font(size=11, italic=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    ws.row_dimensions[r].height = 38
    r += 2

    # Two scenario blocks: REAL, then 50% SWAP
    # Headers include premium $ next to counts so the bonus math is readable
    # (NB tier and REN tier are driven by premium, not count).
    headers = ['Month', 'Agent',
               'NB #', 'NB $', 'RWR #', 'RWR $', 'REN #', 'REN $',
               'Current', 'Plan no-min', 'Plan WITH MIN', 'vs Current']
    ws.cell(row=r, column=1, value='SCENARIO 1: REAL DATA (Jan-Apr 2026)').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    r += 1
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=r, column=i, value=h)
        style_header(c)
    r += 1

    real_totals = {'current': 0, 'a': 0, 'a_min': 0}
    for month in MONTHS:
        for agent in AGENT_NAMES:
            d = AGENTS[agent][month]
            nb_c, nb_p = d['NB'][0], d['NB'][1]
            rwr_c, rwr_p = d['RWR'][0], d['RWR'][1]
            ren_c, ren_p = d['REN'][0], d['REN'][1]
            cur = current_bonus(nb_c, rwr_c)
            a = calc_proposal_a(d['NB'], d['RWR'], d['REN'])
            row_vals = [month, agent,
                        nb_c, nb_p, rwr_c, rwr_p, ren_c, ren_p,
                        cur, a['paid'], a['paid_after_min'],
                        a['paid_after_min'] - cur]
            for i, v in enumerate(row_vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (3,5,7): style_int(c)
                elif i in (4,6,8) or i >= 9: style_dollar(c)
                else: style_data(c)
                if i == 10: c.fill = PROP_A_FILL
                if i == 11: c.fill = PROP_A_FILL; c.font = Font(bold=True)
            real_totals['current'] += cur
            real_totals['a'] += a['paid']
            real_totals['a_min'] += a['paid_after_min']
            r += 1

    ws.cell(row=r, column=1, value='REAL 4-MONTH TOTAL (6 agents)').font = Font(bold=True)
    ws.cell(row=r, column=9, value=real_totals['current'])
    ws.cell(row=r, column=10, value=real_totals['a'])
    ws.cell(row=r, column=11, value=real_totals['a_min'])
    ws.cell(row=r, column=12, value=real_totals['a_min'] - real_totals['current'])
    for i in range(9, 13):
        c = ws.cell(row=r, column=i)
        style_dollar(c)
        c.font = Font(bold=True)
        c.fill = SUB_FILL
    ws.cell(row=r, column=1).fill = SUB_FILL
    r += 2

    # SCENARIO 2: 50% SWAP
    ws.cell(row=r, column=1, value='SCENARIO 2: 50% OF RWR CONVERTED TO REN (future-state target)').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    r += 1
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=r, column=i, value=h)
        style_header(c)
    r += 1

    swap_totals = {'current': 0, 'a': 0, 'a_min': 0}
    for month in MONTHS:
        for agent in AGENT_NAMES:
            d = swap_rwr_to_ren(AGENTS[agent][month], 0.5)
            nb_c, nb_p = d['NB'][0], d['NB'][1]
            rwr_c, rwr_p = d['RWR'][0], d['RWR'][1]
            ren_c, ren_p = d['REN'][0], d['REN'][1]
            cur = current_bonus(nb_c, rwr_c)
            a = calc_proposal_a(d['NB'], d['RWR'], d['REN'])
            row_vals = [month, agent,
                        nb_c, nb_p, rwr_c, rwr_p, ren_c, ren_p,
                        cur, a['paid'], a['paid_after_min'],
                        a['paid_after_min'] - cur]
            for i, v in enumerate(row_vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (3,5,7): style_int(c)
                elif i in (4,6,8): style_dollar(c)
                elif i >= 9: style_dollar(c)
                else: style_data(c)
                if i == 10: c.fill = PROP_A_FILL
                if i == 11: c.fill = PROP_A_FILL; c.font = Font(bold=True)
            swap_totals['current'] += cur
            swap_totals['a'] += a['paid']
            swap_totals['a_min'] += a['paid_after_min']
            r += 1

    ws.cell(row=r, column=1, value='SWAP 4-MONTH TOTAL (6 agents)').font = Font(bold=True)
    ws.cell(row=r, column=9, value=swap_totals['current'])
    ws.cell(row=r, column=10, value=swap_totals['a'])
    ws.cell(row=r, column=11, value=swap_totals['a_min'])
    ws.cell(row=r, column=12, value=swap_totals['a_min'] - swap_totals['current'])
    for i in range(9, 13):
        c = ws.cell(row=r, column=i)
        style_dollar(c)
        c.font = Font(bold=True)
        c.fill = SUB_FILL
    ws.cell(row=r, column=1).fill = SUB_FILL
    r += 3

    # Final read
    ws.cell(row=r, column=1, value='OWNERSHIP READ').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    r += 1
    reads = [
        f"REAL DATA TODAY: Current pays ${real_totals['current']:,.0f} (broken plan, rewards rewriting). The Bonus Plan WITH MIN pays ${real_totals['a_min']:,.0f} (less, because few agents clear the $45k NB gate today).",
        f"100% FLIP - 'if you had been renewing instead of rewriting': The Bonus Plan WITH MIN pays ${head_flip['a_min']:,.0f} ({head_flip['a_min']/real_totals['current']*100:.0f}% of today's $12,220 current). Agent earns close to today for doing the RIGHT behavior.",
        f"WHY CURRENT DROPS WHEN BEHAVIOR SHIFTS: today's plan pays $0 for renewals - it only counts NB+RWR toward the 35-policy tier. When rewrites convert to renewals, the count drops, the bonus drops. Under FLIP, current crashes to ${head_flip['cur']:,.0f}. The new plan pays for renewals, so behavior shift GROWS pay.",
        f"PROFIT PROTECTION: Pay = target. Review threshold flags months above 40% of safe net. 90-day chargeback reverses any bonus on a policy that cancels/rewrites within 90 days. Minimums add a third layer: NB ${NB_MIN_PREMIUM:,}/mo + retention >= {REN_MIN_RETENTION*100:.0f}% of book or no bonus that month.",
        f"RECOMMENDATION: The Bonus Plan with ${NB_MIN_PREMIUM:,} NB + {REN_MIN_RETENTION*100:.0f}% retention minimums. Pilot 90 days side-by-side with the current plan (pay the higher of the two), then switch over.",
    ]
    for txt in reads:
        c = ws.cell(row=r, column=1, value=txt)
        c.font = Font(size=11)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
        ws.row_dimensions[r].height = 44
        r += 1

    # Cols: Month, Agent, NB#, NB$, RWR#, RWR$, REN#, REN$, Current, no-min, WITH MIN, vs Cur
    set_col_widths(ws, [10, 20, 6, 10, 7, 10, 7, 10, 10, 11, 12, 10])
    ws.freeze_panes = 'C6'


def build_proposal_a(wb):
    ws = wb.create_sheet('The Bonus Plan')
    ws['A1'] = 'THE BONUS PLAN'
    ws['A1'].font = TITLE_FONT
    ws['A1'].fill = PROP_A_FILL
    ws.merge_cells('A1:G1')

    ws['A2'] = ('Monthly written-premium TIERS for NB, RWR, and REN. '
                'Same structure as today (30/38/50 count tiers) but driven by PREMIUM, not policy count. '
                'Retention rate gates REN; NB premium gate ($45k) also gates RWR. '
                '3-month chargeback. Excel tracker required for every policy.')
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:G2')

    r = 4
    # SECTION 1 - NB tier ladder
    ws.cell(row=r, column=1, value='1. NEW BUSINESS (NB) - 5 tiers by monthly written premium, $1,000 at $100k').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    r += 1
    headers = ['Tier', 'Monthly NB written premium', 'NB bonus', 'Jump from prior tier', 'Plain English', '', '']
    for i, h in enumerate(headers, 1):
        if h: style_header(ws.cell(row=r, column=i, value=h))
    r += 1
    nb_rows = [
        ('Below T1', 'Under $45,000', '$0', '-', 'Below minimum requirement.'),
        ('T1 (min)', '$45,000', '$250', 'entry', 'Minimum reached.'),
        ('T2', '$55,000', '$375', '+$125', ''),
        ('T3', '$70,000', '$525', '+$150', ''),
        ('T4', '$85,000', '$725', '+$200', ''),
        ('T5', '$100,000', '$1,000', '+$275', 'Top tier - biggest jump.'),
        ('Above $100k', '+$5k extra', '+$50', '1% rate', 'No cap. $150k -> $1,500.'),
    ]
    for row in nb_rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = PROP_A_FILL
        r += 1
    r += 1

    # SECTION 2 - RWR tier ladder (half of NB)
    ws.cell(row=r, column=1, value='2. REWRITES (RWR) - same tier breakpoints as NB, but pays HALF').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    r += 1
    headers2 = ['Tier', 'Monthly RWR written premium', 'RWR bonus (1/2 of NB)', 'Plain English', '', '', '']
    for i, h in enumerate(headers2, 1):
        if h: style_header(ws.cell(row=r, column=i, value=h))
    r += 1
    rwr_rows = [
        ('Below T1', 'Under $45,000', '$0', ''),
        ('T1', '$45,000', '$100', ''),
        ('T2', '$55,000', '$200', ''),
        ('T3', '$70,000', '$275', ''),
        ('T4', '$85,000', '$350', ''),
        ('T5', '$100,000', '$500', 'Top tier.'),
        ('Above $100k', '+$5k extra', '+$25', '0.5% rate.'),
        ('Gate', 'NB premium >= $45k', '', 'No separate RWR min. NB gate unlocks RWR.'),
    ]
    for row in rwr_rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = PROP_A_FILL
        r += 1
    r += 1

    # SECTION 3 - REN tier ladder
    ws.cell(row=r, column=1, value='3. RENEWALS (REN) - own tiers starting at $25k, gated by retention rate').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    r += 1
    headers3 = ['Tier', 'Monthly REN written premium', 'REN bonus (~0.80x of NB)', 'Plain English', '', '', '']
    for i, h in enumerate(headers3, 1):
        if h: style_header(ws.cell(row=r, column=i, value=h))
    r += 1
    ren_rows = [
        ('Below T1', 'Under $25,000', '$0', ''),
        ('T1', '$25,000', '$250', 'Entry - lower bar (smaller books).'),
        ('T2', '$40,000', '$350', ''),
        ('T3', '$65,000', '$450', ''),
        ('T4', '$80,000', '$550', ''),
        ('T5', '$100,000', '$800', 'Top tier - biggest jump.'),
        ('Above $100k', '+$5k extra', '+$40', '0.8% rate. No cap.'),
        ('Gate', 'Retention rate >= 30%', '', 'Must retain >= 30% of assigned renewal book.'),
    ]
    for row in ren_rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = PROP_A_FILL
        r += 1
    r += 1

    # SECTION 4 - minimums + chargeback + tracking
    ws.cell(row=r, column=1, value='4. RULES THAT APPLY ON TOP').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    r += 1
    headers4 = ['Rule', 'Value', 'How it works', '', '', '', '']
    for i, h in enumerate(headers4, 1):
        if h: style_header(ws.cell(row=r, column=i, value=h))
    r += 1
    extra_rows = [
        ('NB minimum', f'${NB_MIN_PREMIUM:,} monthly written premium', 'NB premium below $45k -> NB bonus = $0 (also blocks RWR).'),
        ('REN minimum', f'>= {REN_MIN_RETENTION*100:.0f}% retention rate of book', 'Retention < 30% -> REN bonus = $0.'),
        ('RWR gate', 'NB minimum met', 'No separate RWR floor.'),
        ('90-day chargeback', '100% reversal', 'Bonus on a policy is REVERSED if it cancels/rewrites within 90 days of effective date.'),
        ('Manual tracker', 'Required per policy', 'No entry = no bonus on that policy.'),
        ('Renewal book', 'Assigned per agent', 'Newer agents inherit a book from former employees.'),
    ]
    for row in extra_rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = PROP_A_FILL
        r += 1
    r += 1

    # Worked example
    ws.cell(row=r, column=1, value='5. WORKED EXAMPLE - Dialinerys Dieguez, March 2026').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    r += 1
    ex_headers = ['Step', 'Math', 'Result', '', '', '', '']
    for i, h in enumerate(ex_headers, 1):
        if h: style_header(ws.cell(row=r, column=i, value=h))
    r += 1
    d = AGENTS['Dialinerys Dieguez']['March']
    a = calc_proposal_a(d['NB'], d['RWR'], d['REN'])
    ex_rows = [
        ('Volumes', f'NB {d["NB"][0]} policies = ${d["NB"][1]:,.0f} written premium ({d["NB"][2]/d["NB"][1]*100:.1f}% collected)', ''),
        ('Volumes', f'RWR {d["RWR"][0]} policies = ${d["RWR"][1]:,.0f} written premium ({d["RWR"][2]/d["RWR"][1]*100:.1f}% collected)', ''),
        ('Volumes', f'REN {d["REN"][0]} policies = ${d["REN"][1]:,.0f} written premium', ''),
        ('NB tier', f'${d["NB"][1]/1000:,.0f}k NB written premium lands in {a["nb_tier_label"]}', f'${a["nb_target"]:.0f}'),
        ('RWR tier', f'${d["RWR"][1]/1000:,.0f}k RWR written premium lands in {a["rwr_tier_label"]}', f'${a["rwr_target"]:.0f}'),
        ('REN tier', f'${d["REN"][1]/1000:,.0f}k REN written premium lands in {a["ren_tier_label"]}', f'${a["ren_target"]:.0f}'),
        ('Subtotal target', '(NB + RWR + REN, before gates)', f'${a["total_target"]:.0f}'),
        (f'NB gate (${NB_MIN_PREMIUM:,})', f'NB written ${d["NB"][1]:,.0f} vs ${NB_MIN_PREMIUM:,}', 'PASS' if a['nb_qual'] else 'FAIL'),
        (f'REN gate (>= {REN_MIN_RETENTION*100:.0f}% retention)', f'Assumed retention 75% vs {REN_MIN_RETENTION*100:.0f}%', 'PASS' if a['ren_qual'] else 'FAIL'),
        (f'RWR gate (NB met)', 'RWR pays only if NB gate passed?', 'PASS' if a['rwr_qual'] else 'FAIL'),
        ('FINAL PAID', 'Pay only the lines whose gates passed', f'${a["paid_after_min"]:.2f}'),
    ]
    for row in ex_rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if 'FAIL' in str(v): c.fill = WARN_FILL
            elif 'PASS' in str(v): c.fill = PatternFill('solid', fgColor='C6EFCE')
            elif row[0] == 'FINAL PAID':
                c.fill = SUB_FILL
                c.font = Font(bold=True)
            elif r % 2 == 0: c.fill = PROP_A_FILL
        ws.row_dimensions[r].height = 24
        r += 1

    set_col_widths(ws, [26, 56, 22, 8, 8, 8, 8])


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
        ('NB Base Coverage', 'PIP/PD only OR PIP + Comp/Coll', '$11 flat', 'Either FL minimum or full coverage qualifies for the base.', 'PIP+PD NB or PIP+Comp+Coll NB', '$11', 'Per policy, before kicker'),
        ('NB Liability Add', 'BI + UM bundled together', '+$6', 'Adds liability. Bundled because UM cannot exist without BI in FL.', 'NB with BI + UM', '$11 + $6 = $17', 'Bundle only - see below'),
        ('NB - BI ALONE', 'BI present but NO UM', '$0 add (base only)', 'Does not earn the liability bundle. Customer is still partially exposed.', 'NB with BI no UM', '$11 base only', 'Manager verifies'),
        ('NB - UM ALONE', 'Not possible', 'Not applicable', 'UM cannot be issued without BI per FL rules. Bundle exists for this reason.', '-', '-', 'Carrier system blocks this'),
        ('REN Base Coverage', 'Same coverage logic as NB', '$7 flat', 'Renewals paid for the first time, but below NB.', 'Any REN', '$7', 'Before kicker'),
        ('REN Liability Add', 'BI + UM bundled', '+$3', 'Same bundle logic, smaller dollar amount since the coverage already existed.', 'REN with BI + UM', '$7 + $3 = $10', 'Manager verifies'),
        ('RWR', 'Any rewrite', '$2 flat', 'No coverage stacking. Rewrites are intentionally small.', 'Any RWR', '$2', 'No bundle on rewrites'),
        ('PIF Add', 'Policy paid in full', '+$9 NB / +$13 high NB / +$5 REN', 'Cash upfront earns extra.', '$2,500 NB PIF + BI/UM', '$11 + $6 + $9 = $26', 'Manager verifies'),
        ('Collected Kicker', 'Same as Proposal A', '+10% / +15% / +20% / +25%', '15-24% / 25-49% / 50-99% / 100%', '$17 target at 25% collected', '$17 x 1.15 = $19.55', 'Identical kicker logic'),
        ('Minimum Requirements', f'NB >= ${NB_MIN_PREMIUM:,}/mo + retention >= {REN_MIN_RETENTION*100:.0f}% of book', 'Gates each bonus line', 'Below NB gate, no NB or RWR bonus. Below REN retention gate, no REN bonus.', 'NB $30k written', 'NB bonus = $0', 'RWR follows NB gate'),
        ('Review Threshold (soft cap)', 'Target > 40% of safe net', 'Flag for ownership review', 'Same as Proposal A. Does NOT auto-reduce.', 'Safe net $200, target $80 (40%)', 'Pay $80, flag for review', '40% / 55% PIF'),
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
    headers = ['Scenario', 'Type', 'Premium', 'Coverages', 'Collected %', 'Current', 'Proposal A (Premium)', 'Proposal B (Coverage)', 'Plain-English Read']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    examples = [
        ('Basic FL NB - PIP/PD only',          'NB', 1000, 'PIP+PD',           '25%', 'Counts toward tier',  '$6 x 1.15 = $6.90',              '$11 x 1.15 = $12.65',                'B pays more - coverage = money, even at low premium.'),
        ('Full Coverage NB - no liability',    'NB', 1500, 'PIP+PD+Comp+Coll', '25%', 'Counts toward tier',  '$8 x 1.15 = $9.20',              '$11 x 1.15 = $12.65',                'B pays base only (no BI+UM).'),
        ('Full Coverage + Liability NB',       'NB', 2000, 'PIP+PD+CC+BI+UM',  '25%', 'Counts toward tier',  '$11 x 1.15 = $12.65',            '($11+$6) x 1.15 = $19.55',           'B clearly wins because liability bundle adds $6.'),
        ('NB with BI ALONE (no UM)',           'NB', 2000, 'PIP+PD+BI no UM',  '25%', 'Counts toward tier',  '$11 x 1.15 = $12.65',            '$11 x 1.15 = $12.65 (base only)',    'B does NOT pay the bundle because BI alone does not qualify.'),
        ('High-premium PIF NB',                'NB', 4500, 'PIP+PD+CC+BI+UM',  '100% (PIF)', 'Counts toward tier','($13+$3 over-$3k+$13 PIF) x 1.25 = $36.25','($11+$6+$13) x 1.25 = $37.50',       'PIF adds the biggest dollars on both A and B.'),
        ('Standard Renewal',                   'REN',1500, 'PIP+PD+CC',        '25%', 'Pays $0 today',       '$6 x 1.15 = $6.90',              '$7 x 1.15 = $8.05',                  'Renewals pay - this is new (current = $0).'),
        ('Renewal + Liability + PIF',          'REN',2000, 'Full + BI+UM PIF', '100% (PIF)', 'Pays $0 today','($7+$5 PIF) x 1.25 = $15.00',     '($7+$3+$5) x 1.25 = $18.75',         'B rewards renewal liability bundle + PIF.'),
        ('Rewrite (any premium)',              'RWR',1200, 'Any',              '25%', 'Counts toward tier (BAD)','$2 x 1.15 = $2.30',           '$2 x 1.15 = $2.30',                  'Both new plans deprioritize RWR.'),
    ]

    for ex in examples:
        for i, v in enumerate(ex, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if i == 7: c.fill = PROP_A_FILL
            elif i == 8: c.fill = PROP_B_FILL
            elif i == 6: c.fill = CURRENT_FILL
        ws.row_dimensions[r].height = 32
        r += 1

    set_col_widths(ws, [32, 7, 11, 18, 13, 22, 28, 32, 50])


def build_agent_examples(wb, swap=False):
    """Detailed month-by-month per-agent worked examples under the new bonus plan."""
    name = 'Agent Examples - 50% Swap' if swap else 'Agent Examples - Real'
    ws = wb.create_sheet(name)
    ws['A1'] = f"Agent Examples - {'50% RWR converted to REN' if swap else 'REAL Jan-Apr 2026 data'}"
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:N1')

    subtitle = ('Same premiums, same collected $, but 50% of each agent\'s rewrites are reclassified as renewals. Shows future-state earnings.'
                if swap else
                'Actual NB/RWR/REN volumes from the source data, calculated under the new bonus plan.')
    ws['A2'] = subtitle
    ws['A2'].font = Font(italic=True, size=10, color='666666')
    ws.merge_cells('A2:N2')

    r = 4
    for agent in AGENT_NAMES:
        # Agent header
        ws.cell(row=r, column=1, value=agent).font = SECTION_FONT
        ws.cell(row=r, column=1).fill = SECTION_FILL
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
        ws.row_dimensions[r].height = 22
        r += 1

        # Volume table header
        vol_headers = ['Month', 'NB Count', 'NB Premium', 'NB Collected', 'NB Coll %',
                       'RWR Count', 'RWR Premium', 'RWR Collected', 'RWR Coll %',
                       'REN Count', 'REN Premium', 'REN Collected', 'REN Coll %',
                       'Avg NB Prem']
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

            vals = [month,
                    nb_c, nb_p, nb_col, (nb_col / nb_p if nb_p else 0),
                    rwr_c, rwr_p, rwr_col, (rwr_col / rwr_p if rwr_p else 0),
                    ren_c, ren_p, ren_col, (ren_col / ren_p if ren_p else 0),
                    nb_p / nb_c if nb_c else 0]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i == 1: style_data(c)
                elif i in (2, 6, 10): style_int(c)
                elif i in (5, 9, 13): style_pct(c)
                else: style_dollar(c)
            r += 1
        r += 1

        # THE BONUS PLAN calculation
        ws.cell(row=r, column=1, value='THE BONUS PLAN - bonus calculation').font = SUB_FONT
        ws.cell(row=r, column=1).fill = PROP_A_FILL
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
        r += 1

        a_headers = ['Month',
                     'NB premium', 'NB tier', 'NB bonus',
                     'RWR premium', 'RWR tier', 'RWR bonus',
                     'REN premium', 'REN tier', 'REN bonus',
                     'Total target', 'PAID after MIN', 'Gates',
                     'vs Current']
        for i, h in enumerate(a_headers, 1):
            style_header(ws.cell(row=r, column=i, value=h))
        r += 1

        a_totals = {'paid': 0, 'paid_min': 0, 'cur': 0}
        for month in MONTHS:
            d = agent_data[month]
            a = calc_proposal_a(d['NB'], d['RWR'], d['REN'])
            cur = current_bonus(d['NB'][0], d['RWR'][0])
            gates = f"{'Y' if a['nb_qual'] else 'n'}/{'Y' if a['ren_qual'] else 'n'}/{'Y' if a['rwr_qual'] else 'n'}"
            vals = [month,
                    d['NB'][1], a['nb_tier_label'], a['nb_target'],
                    d['RWR'][1], a['rwr_tier_label'], a['rwr_target'],
                    d['REN'][1], a['ren_tier_label'], a['ren_target'],
                    a['total_target'], a['paid_after_min'], gates,
                    a['paid_after_min'] - cur]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (1, 3, 6, 9, 13): style_data(c)  # text columns
                else: style_dollar(c)
                c.fill = PROP_A_FILL
                if i == 12: c.font = Font(bold=True)
                if i == 13:
                    if a['nb_qual'] and a['ren_qual']: c.fill = PatternFill('solid', fgColor='C6EFCE')
                    elif a['nb_qual'] or a['ren_qual']: c.fill = PatternFill('solid', fgColor='FFEB9C')
                    else: c.fill = WARN_FILL
            a_totals['paid'] += a['paid']
            a_totals['paid_min'] += a['paid_after_min']
            a_totals['cur'] += cur
            r += 1

        ws.cell(row=r, column=1, value='4-Month Total').font = Font(bold=True)
        ws.cell(row=r, column=11, value=a_totals['paid'])
        ws.cell(row=r, column=12, value=a_totals['paid_min'])
        ws.cell(row=r, column=14, value=a_totals['paid_min'] - a_totals['cur'])
        for i in range(1, 15):
            c = ws.cell(row=r, column=i)
            c.fill = SUB_FILL
            if i in (11, 12, 14):
                style_dollar(c)
                c.font = Font(bold=True)
        r += 3  # spacing between agents

    # Column widths sized for both volume and tier tables.
    set_col_widths(ws, [10, 12, 12, 11, 12, 12, 11, 12, 12, 11, 12, 13, 9, 11])


def build_comparison(wb):
    ws = wb.create_sheet('Real vs Swap Comparison')
    ws['A1'] = 'Real vs 50%-Swap Comparison (4-month totals per agent)'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:K1')

    ws['A2'] = ('Same agents, same months. NO-MIN columns show target pay. WITH-MIN columns apply the '
                f'${NB_MIN_PREMIUM:,}/${REN_MIN_PREMIUM:,} gates. The renewal-shift upside is the main message: '
                'today the current plan punishes the swap; the new plans reward it.')
    ws['A2'].font = Font(italic=True, size=10, color='666666')
    ws.merge_cells('A2:K2')

    r = 4
    headers = ['Agent', 'Current (Real)', 'Current (Swap)',
               'A no-min (Real)', 'A no-min (Swap)',
               'A WITH MIN (Real)', 'A WITH MIN (Swap)',
               'B WITH MIN (Real)', 'B WITH MIN (Swap)',
               'A-min Swap-Real', 'B-min Swap-Real']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    grand = {k: 0 for k in ['cur_r','cur_s','a_r','a_s','am_r','am_s','b_r','b_s','bm_r','bm_s']}
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
            a['a_r']  += ra['paid'];  a['a_s']  += sa['paid']
            a['am_r'] += ra['paid_after_min']; a['am_s'] += sa['paid_after_min']
            a['b_r']  += rb['paid'];  a['b_s']  += sb['paid']
            a['bm_r'] += rb['paid_after_min']; a['bm_s'] += sb['paid_after_min']
        row_vals = [agent, a['cur_r'], a['cur_s'],
                    a['a_r'], a['a_s'], a['am_r'], a['am_s'],
                    a['bm_r'], a['bm_s'],
                    a['am_s'] - a['am_r'], a['bm_s'] - a['bm_r']]
        for col, val in enumerate(row_vals, 1):
            c = ws.cell(row=r, column=col, value=val)
            if col == 1: style_data(c)
            else: style_dollar(c)
            if col in (6, 7): c.fill = PROP_A_FILL
            if col in (8, 9): c.fill = PROP_B_FILL
        for k in grand: grand[k] += a[k]
        r += 1

    ws.cell(row=r, column=1, value='GRAND TOTAL (6 agents, 4 months)').font = Font(bold=True)
    grand_vals = ['GRAND TOTAL (6 agents, 4 months)',
                  grand['cur_r'], grand['cur_s'],
                  grand['a_r'], grand['a_s'], grand['am_r'], grand['am_s'],
                  grand['bm_r'], grand['bm_s'],
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
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
    r += 1
    reads = [
        '"Real" columns are today\'s numbers. "Swap" columns are what each agent would earn if 50% of their rewrites became renewals (same dollars, just reclassified).',
        'Under the CURRENT plan, the Swap actually DROPS pay - because the current plan rewards rewrites in the 35-policy NB+RWR count tier. Broken incentive. (See "Why Current Drops" sheet.)',
        'Under Proposals A and B no-min, the Swap INCREASES pay - the behavior change ownership wants is rewarded.',
        'Under WITH-MIN, the Swap also increases pay because more agents clear the REN gate.',
        'Use this tab to defend the rollout: "Here is the dollar reward each agent gets for stopping the rewrite habit."',
    ]
    for t in reads:
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=11)
        c.alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
        ws.row_dimensions[r].height = 30
        r += 1

    set_col_widths(ws, [25, 13, 13, 13, 13, 13, 13, 13, 13, 14, 14])


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
    ws['A1'] = 'Collected Incentive - per policy, flat dollars'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:E1')

    ws['A2'] = ('Per-policy flat dollar amount added to the base bonus. ONLY applies to policies with WRITTEN PREMIUM > $1,200. '
                'A $800 policy does not earn the kicker, regardless of how much was collected.')
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:E2')

    r = 4
    headers = ['Collected % of premium', 'Per-policy kicker', 'Plain English', 'Example: $1,500 NB base $7', 'Final per-policy pay']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    rows = [
        ('Below 15%', '$0 (none)', 'Premium barely collected. No kicker.', '$1,500 NB, 10% down', '$7 (base only)'),
        ('15% - 19%', '+$1', 'Minimum down. Small bump.', '$1,500 NB, 17% down', '$7 + $1 = $8'),
        ('20% - 24%', '+$2', 'Standard down. Standard bump.', '$1,500 NB, 22% down', '$7 + $2 = $9'),
        ('25% - 99%', '+$5', 'High collection. Big bump.', '$1,500 NB, 30% down', '$7 + $5 = $12'),
        ('100% PIF', '+$8', 'Paid in full. Top kicker - zero chargeback risk.', '$1,500 NB PIF', '$7 + $8 = $15'),
        ('Premium <= $1,200', '$0 (no kicker)', 'Low-premium policies do not earn the incentive, no matter the collected %.', '$800 NB, 50% down', '$5 (base only)'),
    ]
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = SUB_FILL
        ws.row_dimensions[r].height = 26
        r += 1

    r += 2
    ws.cell(row=r, column=1, value='Why a per-policy flat kicker').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    r += 1
    why = [
        'Each policy earns its own kicker based on its own collected % - not the agent\'s total collected %. A $50 down on one policy and a PIF on another are rewarded separately.',
        'The $1,200 premium floor excludes the smallest policies. A $700 policy at 25% collected is still only $175 in cash - not enough to justify a $5 bump.',
        'Flat dollars (not a multiplier) make the math easy to audit. Every policy line on the tracker shows: base + kicker = pay.',
        'PIF is the top tier at +$8 because PIF policies have zero chargeback risk - the carrier already has all the money.',
    ]
    for t in why:
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=11)
        c.alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 30
        r += 1

    set_col_widths(ws, [22, 18, 40, 26, 22])


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
        ('2. After royalty', 'Gross Comm x (1 - 21%)', '$550 x 0.79', '$434.50', 'Franchise royalty removed first.'),
        ('3. Safe net', 'After-royalty x (1 - 30% overhead)', '$434.50 x 0.70', '$304.15', 'Cash available after rent/tech/admin (salaries tracked via the minimum, not here).'),
        ('4. Bonus target (proposal)', 'NB + RWR + REN tier amounts', 'Per the bonus plan sheet', 'e.g., $250', 'Calculated by the plan rules.'),
        ('5. Bonus % of safe net', 'Target / safe net', '$70 / $304.15', '23%', 'How much of safe net is going to bonus this month for this agent.'),
        ('6. Review threshold', '40% of safe net', '$304.15 x 0.40 = $121.66', 'Compare', 'If target <= $121.66, auto-pay. If target > $121.66, ownership reviews.'),
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
        'A hard cap on every month at a low % (e.g., 15%) would squash the plan so much that it pays the same amount regardless of agent performance - defeating the incentive entirely.',
        'The 40% REVIEW threshold flags outlier months (low collection, high written premium) so ownership can act, without making the everyday bonus arbitrary.',
        'The 90-day CHARGEBACK is the actual profit protection. If a policy cancels in 90 days, the bonus is reversed. That matches carrier commission chargeback exposure exactly.',
        'The 4-month modeled cost: Current $12,220, New plan $7,975 at 100% FLIP (still cheaper than today AND aligns pay with the right behavior).',
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


def build_agent_examples_flip(wb):
    """100% FLIP scenario agent examples - what if every rewrite had been a renewal."""
    ws = wb.create_sheet('Agent Examples - 100% Flip')
    ws['A1'] = "Agent Examples - 100% FLIP (if rewrites had been renewals)"
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:N1')
    ws['A2'] = "Same premiums, same collected $, but rewrites and renewals are FULLY SWAPPED. Strongest behavior-change test."
    ws['A2'].font = Font(italic=True, size=10, color='666666')
    ws.merge_cells('A2:N2')

    r = 4
    for agent in AGENT_NAMES:
        ws.cell(row=r, column=1, value=agent).font = SECTION_FONT
        ws.cell(row=r, column=1).fill = SECTION_FILL
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
        r += 1

        # Volume table
        vol_headers = ['Month', 'NB Count', 'NB Premium', 'NB Collected', 'NB Coll %',
                       'RWR Count', 'RWR Premium', 'RWR Collected', 'RWR Coll %',
                       'REN Count', 'REN Premium', 'REN Collected', 'REN Coll %']
        for i, h in enumerate(vol_headers, 1):
            style_header(ws.cell(row=r, column=i, value=h))
        r += 1

        for month in MONTHS:
            d = full_flip(AGENTS[agent][month])
            nb_c, nb_p, nb_col = d['NB']
            rwr_c, rwr_p, rwr_col = d['RWR']
            ren_c, ren_p, ren_col = d['REN']
            vals = [month, nb_c, nb_p, nb_col, (nb_col/nb_p if nb_p else 0),
                    rwr_c, rwr_p, rwr_col, (rwr_col/rwr_p if rwr_p else 0),
                    ren_c, ren_p, ren_col, (ren_col/ren_p if ren_p else 0)]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i == 1: style_data(c)
                elif i in (2, 6, 10): style_int(c)
                elif i in (5, 9, 13): style_pct(c)
                else: style_dollar(c)
            r += 1
        r += 1

        # THE BONUS PLAN calculation
        ws.cell(row=r, column=1, value='THE BONUS PLAN - bonus calculation under 100% FLIP').font = SUB_FONT
        ws.cell(row=r, column=1).fill = PROP_A_FILL
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
        r += 1

        a_headers = ['Month',
                     'NB premium', 'NB tier', 'NB bonus',
                     'RWR premium', 'RWR tier', 'RWR bonus',
                     'REN premium', 'REN tier', 'REN bonus',
                     'Total target', 'PAID after MIN', 'Gates',
                     'vs Current']
        for i, h in enumerate(a_headers, 1):
            style_header(ws.cell(row=r, column=i, value=h))
        r += 1

        a_totals = {'paid_min': 0, 'cur': 0}
        for month in MONTHS:
            d = full_flip(AGENTS[agent][month])
            a = calc_proposal_a(d['NB'], d['RWR'], d['REN'])
            cur = current_bonus(d['NB'][0], d['RWR'][0])
            gates = f"{'Y' if a['nb_qual'] else 'n'}/{'Y' if a['ren_qual'] else 'n'}/{'Y' if a['rwr_qual'] else 'n'}"
            vals = [month,
                    d['NB'][1], a['nb_tier_label'], a['nb_target'],
                    d['RWR'][1], a['rwr_tier_label'], a['rwr_target'],
                    d['REN'][1], a['ren_tier_label'], a['ren_target'],
                    a['total_target'], a['paid_after_min'], gates,
                    a['paid_after_min'] - cur]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (1, 3, 6, 9, 13): style_data(c)
                else: style_dollar(c)
                c.fill = PROP_A_FILL
                if i == 12: c.font = Font(bold=True)
                if i == 13:
                    if a['nb_qual'] and a['ren_qual']: c.fill = PatternFill('solid', fgColor='C6EFCE')
                    elif a['nb_qual'] or a['ren_qual']: c.fill = PatternFill('solid', fgColor='FFEB9C')
                    else: c.fill = WARN_FILL
            a_totals['paid_min'] += a['paid_after_min']
            a_totals['cur'] += cur
            r += 1

        ws.cell(row=r, column=1, value='4-Month Total').font = Font(bold=True)
        ws.cell(row=r, column=12, value=a_totals['paid_min'])
        ws.cell(row=r, column=14, value=a_totals['paid_min'] - a_totals['cur'])
        for i in range(1, 15):
            c = ws.cell(row=r, column=i)
            c.fill = SUB_FILL
            if i in (12, 14):
                style_dollar(c)
                c.font = Font(bold=True)
        r += 2

    set_col_widths(ws, [10, 12, 12, 11, 12, 12, 11, 12, 12, 11, 12, 13, 9, 11])


# ============================================================================
# BONUS EXAMPLES TABLES (one sheet per type: NB, RWR, REN)
# Each sheet shows REAL, SWAP, FLIP scenarios stacked. Rows = 6 agents x
# 4 months, sorted by premium descending. Includes a data-bar progress
# column and color-coded tier hits.
# ============================================================================

TIER_FILL_T5  = PatternFill('solid', fgColor='00B050')  # bright green
TIER_FILL_T4  = PatternFill('solid', fgColor='92D050')  # light green
TIER_FILL_T3  = PatternFill('solid', fgColor='FFFF00')  # yellow
TIER_FILL_T2  = PatternFill('solid', fgColor='FFC000')  # orange
TIER_FILL_T1  = PatternFill('solid', fgColor='F4B084')  # peach
TIER_FILL_NIL = PatternFill('solid', fgColor='C00000')  # red (below min)
TIER_FILL_OK  = PatternFill('solid', fgColor='00B050')  # green for met


def _tier_fill_and_label(tier_num, tier_label, tier_amount, line_type):
    """Return the cell fill and a descriptive label for the tier column."""
    if tier_num == 0:
        if line_type == 'REN':
            return TIER_FILL_NIL, 'Below T1 (need $25k REN premium)'
        return TIER_FILL_NIL, 'Below T1 - Minimum Requirement NOT met'
    color_map = {1: TIER_FILL_T1, 2: TIER_FILL_T2, 3: TIER_FILL_T3,
                 4: TIER_FILL_T4, 5: TIER_FILL_T5}
    fill = color_map.get(min(tier_num, 5), TIER_FILL_T5)
    return fill, f"Tier {tier_num} ({tier_label}) - ${tier_amount:,.0f}"


def _build_one_examples_table(ws, r, scenario_label, scenario_fn, line_type,
                              tier_bonus_fn, line_min_text):
    """Build one scenario table inside a Bonus Examples sheet.

    line_type: 'NB' | 'RWR' | 'REN'
    scenario_fn: callable that takes the original agent-month dict and
                 returns a possibly-modified one (for SWAP / FLIP).
    """
    ws.cell(row=r, column=1, value=scenario_label).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 22
    r += 1

    headers = ['Agent', 'Month', f'{line_type} Premium', 'Progress Bar',
               'Tier', 'Total Bonus $']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    ws.row_dimensions[r].height = 26
    r += 1

    rows = []
    for agent in AGENT_NAMES:
        for month in MONTHS:
            d = scenario_fn(AGENTS[agent][month])
            premium = d[line_type][1]
            bonus_amt, tier_num, tier_label = tier_bonus_fn(premium)
            rows.append((agent, month, premium, bonus_amt, tier_num, tier_label))

    # Sort by premium descending so top earners surface first (matches the
    # screenshot from the boss).
    rows.sort(key=lambda x: -x[2])

    data_start_row = r
    for agent, month, premium, bonus_amt, tier_num, tier_label in rows:
        ws.cell(row=r, column=1, value=agent)
        ws.cell(row=r, column=2, value=month)
        c_prem = ws.cell(row=r, column=3, value=premium)
        c_prog = ws.cell(row=r, column=4, value=premium)  # value drives the data bar
        c_tier = ws.cell(row=r, column=5)
        c_bonus = ws.cell(row=r, column=6, value=bonus_amt)

        tier_fill, tier_text = _tier_fill_and_label(tier_num, tier_label, bonus_amt, line_type)
        c_tier.value = tier_text
        c_tier.fill = tier_fill
        c_tier.font = Font(bold=True, size=11, color='FFFFFF' if tier_num == 0 else '000000')
        c_tier.alignment = Alignment(horizontal='center', vertical='center')

        style_data(ws.cell(row=r, column=1))
        ws.cell(row=r, column=1).font = Font(bold=True)
        style_data(ws.cell(row=r, column=2))
        style_dollar(c_prem)
        style_dollar(c_prog)
        c_prog.font = Font(color='FFFFFF', size=9)  # hide the duplicated number
        style_dollar(c_bonus)
        c_bonus.font = Font(bold=True, size=12, color='006100' if bonus_amt > 0 else '9C0006')

        ws.row_dimensions[r].height = 22
        r += 1

    # Data bar on the Progress column - max value = $120k so we can see >$100k.
    progress_range = f'D{data_start_row}:D{r-1}'
    bar_rule = DataBarRule(start_type='num', start_value=0,
                           end_type='num', end_value=120000,
                           color='63BE7B', showValue=False)
    ws.conditional_formatting.add(progress_range, bar_rule)

    # Footer note: tier ladder reminder.
    ws.cell(row=r, column=1, value=line_min_text).font = Font(italic=True, size=10, color='666666')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 20
    r += 2
    return r


def build_bonus_examples(wb):
    """One sheet per bonus line (NB / RWR / REN). Each has 3 scenarios
    stacked: REAL, 50% SWAP, 100% FLIP. Within each scenario rows are
    sorted by premium descending (top earners on top), with a data-bar
    progress column and color-coded tier hits."""

    configs = [
        ('NB Bonus Examples', 'NB', nb_tier_bonus,
         'NB tiers: T1 $45k=$250 | T2 $55k=$375 | T3 $70k=$525 | T4 $85k=$725 | T5 $100k=$1,000 | Above $100k: +$50 per $5k. Minimum requirement = $45k NB premium.'),
        ('RWR Bonus Examples', 'RWR', rwr_tier_bonus,
         'RWR tiers: T1 $45k=$100 | T2 $55k=$200 | T3 $70k=$275 | T4 $85k=$350 | T5 $100k=$500 | Above $100k: +$25 per $5k. RWR gated by NB minimum.'),
        ('REN Bonus Examples', 'REN', ren_tier_bonus,
         'REN tiers: T1 $25k=$250 | T2 $40k=$350 | T3 $65k=$450 | T4 $80k=$550 | T5 $100k=$800 | Above $100k: +$40 per $5k. REN gated by 30% retention rate.'),
    ]
    scenarios = [
        ('SCENARIO 1: REAL DATA (Jan-Apr 2026)',          lambda d: d),
        ('SCENARIO 2: 50% SWAP (half of RWR -> REN)',     lambda d: swap_rwr_to_ren(d, 0.5)),
        ('SCENARIO 3: 100% FLIP (RWR fully -> REN)',      full_flip),
    ]

    for sheet_name, line_type, tier_fn, ladder_text in configs:
        ws = wb.create_sheet(sheet_name)
        ws['A1'] = f"{line_type} BONUS EXAMPLES - 6 agents x 4 months, three scenarios"
        ws['A1'].font = TITLE_FONT
        ws.merge_cells('A1:F1')
        ws['A2'] = (f"Each row = one agent-month under {line_type}. Sorted by {line_type} premium descending. "
                    f"Progress bar = {line_type} premium / $120k. Tier shows the bonus formula.")
        ws['A2'].font = Font(italic=True, size=10, color='666666')
        ws.merge_cells('A2:F2')

        r = 4
        for scenario_label, scenario_fn in scenarios:
            r = _build_one_examples_table(ws, r, scenario_label, scenario_fn,
                                          line_type, tier_fn, ladder_text)

        # Column widths: Agent | Month | Premium | Progress | Tier | Bonus
        set_col_widths(ws, [24, 12, 14, 24, 42, 14])
        ws.freeze_panes = 'A4'


def build_source_data(wb):
    """Raw NB/RWR/REN source data for all 6 agents x 4 months. The data behind every calculation."""
    ws = wb.create_sheet('Source Data')
    ws['A1'] = "Source Data - 6 Agents, January-April 2026"
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:N1')
    ws['A2'] = "This is the raw NB / RWR / REN production data behind every calculation in the workbook. From agents_nbrwr_pivots_good_.xlsx."
    ws['A2'].font = Font(italic=True, size=10, color='666666')
    ws.merge_cells('A2:N2')

    r = 4
    headers = ['Agent', 'Month', 'NB Count', 'NB Written', 'NB Collected', 'NB Coll %',
               'RWR Count', 'RWR Written', 'RWR Collected', 'RWR Coll %',
               'REN Count', 'REN Written', 'REN Collected', 'REN Coll %']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    for agent in AGENT_NAMES:
        for month in MONTHS:
            d = AGENTS[agent][month]
            nb_c, nb_p, nb_col = d['NB']
            rwr_c, rwr_p, rwr_col = d['RWR']
            ren_c, ren_p, ren_col = d['REN']
            vals = [agent, month, nb_c, nb_p, nb_col, (nb_col/nb_p if nb_p else 0),
                    rwr_c, rwr_p, rwr_col, (rwr_col/rwr_p if rwr_p else 0),
                    ren_c, ren_p, ren_col, (ren_col/ren_p if ren_p else 0)]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (1, 2): style_data(c)
                elif i in (3, 7, 11): style_int(c)
                elif i in (6, 10, 14): style_pct(c)
                else: style_dollar(c)
            r += 1
        # Agent totals row
        d_tot_nb_c = sum(AGENTS[agent][m]['NB'][0] for m in MONTHS)
        d_tot_nb_p = sum(AGENTS[agent][m]['NB'][1] for m in MONTHS)
        d_tot_nb_col = sum(AGENTS[agent][m]['NB'][2] for m in MONTHS)
        d_tot_rwr_c = sum(AGENTS[agent][m]['RWR'][0] for m in MONTHS)
        d_tot_rwr_p = sum(AGENTS[agent][m]['RWR'][1] for m in MONTHS)
        d_tot_rwr_col = sum(AGENTS[agent][m]['RWR'][2] for m in MONTHS)
        d_tot_ren_c = sum(AGENTS[agent][m]['REN'][0] for m in MONTHS)
        d_tot_ren_p = sum(AGENTS[agent][m]['REN'][1] for m in MONTHS)
        d_tot_ren_col = sum(AGENTS[agent][m]['REN'][2] for m in MONTHS)
        vals = ['', '4-mo TOTAL', d_tot_nb_c, d_tot_nb_p, d_tot_nb_col,
                d_tot_nb_col/d_tot_nb_p if d_tot_nb_p else 0,
                d_tot_rwr_c, d_tot_rwr_p, d_tot_rwr_col,
                d_tot_rwr_col/d_tot_rwr_p if d_tot_rwr_p else 0,
                d_tot_ren_c, d_tot_ren_p, d_tot_ren_col,
                d_tot_ren_col/d_tot_ren_p if d_tot_ren_p else 0]
        for i, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=i, value=v)
            c.fill = SUB_FILL
            c.font = Font(bold=True)
            if i in (1, 2): style_data(c)
            elif i in (3, 7, 11): style_int(c)
            elif i in (6, 10, 14): style_pct(c)
            else: style_dollar(c)
        r += 2

    set_col_widths(ws, [22, 12, 10, 12, 12, 10, 10, 12, 12, 10, 10, 12, 12, 10])


def build_manual_tracker(wb):
    ws = wb.create_sheet('Manual Tracker Template')
    ws['A1'] = 'Monthly Manual Bonus Tracker - The Bonus Plan'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:O1')

    ws['A2'] = ('Agents enter ONE ROW PER POLICY. Without an entry the policy does NOT get bonused. '
                'Manager verifies fields. Per-policy bonus columns auto-calculate from the formulas below.')
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:O2')

    # Reference panel with the bonus formulas
    r = 3
    ws.cell(row=r, column=1, value='THE BONUS PLAN FORMULAS (for reference)').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=15)
    r += 1
    formula_rows = [
        ['NB tiered bonus', 'Monthly NB written premium: T1 $45k=$250 | T2 $55k=$375 | T3 $70k=$525 | T4 $85k=$725 | T5 $100k=$1,000. Above $100k: +$50 per $5k.', '', '', ''],
        ['RWR tiered bonus', 'Same breakpoints as NB, pays ~half: T1=$100 | T2=$200 | T3=$275 | T4=$350 | T5=$500. Above $100k: +$25 per $5k.', '', '', ''],
        ['REN tiered bonus', 'Own breakpoints: T1 $25k=$250 | T2 $40k=$350 | T3 $65k=$450 | T4 $80k=$550 | T5 $100k=$800. Above $100k: +$40 per $5k.', '', '', ''],
        ['Minimums (monthly)', 'NB written premium >= $45,000 | REN retention rate >= 30% of book | RWR follows NB gate (no separate RWR floor)', '', '', ''],
        ['Chargeback', '3 months (90 days) - bonus reversed if policy cancels/rewrites within 90 days', '', '', ''],
    ]
    for row in formula_rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            c.fill = PatternFill('solid', fgColor='F2F2F2')
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=15)
        ws.row_dimensions[r].height = 22
        r += 1
    r += 1

    # Tracker table headers
    ws.cell(row=r, column=1, value='ONE ROW PER POLICY').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=15)
    r += 1

    headers = ['Agent', 'Customer', 'Policy #', 'Carrier',
               'Type (NB/RWR/REN)',
               'Eff Date', 'Exp Date', '90-Day Review Date',
               'Written Premium', 'Down Payment $', 'Collected % (calc)',
               'PIF?', 'Per-Policy Base $', 'Collected Incentive $', 'Total Bonus $']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    # Empty rows for data entry
    for _ in range(30):
        for i in range(1, 16):
            c = ws.cell(row=r, column=i, value='')
            style_data(c)
        r += 1

    set_col_widths(ws, [16, 16, 12, 14, 14, 11, 11, 14, 13, 13, 14, 7, 14, 16, 13])


def build_safe_net_simple(wb):
    """Plain-language explanation of what 'safe net' means and where it comes from."""
    ws = wb.create_sheet('Safe Net Explained')
    ws['A1'] = 'Safe Net Explained Simply'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:F1')

    ws['A2'] = ("Safe net = the commission the agency expects to KEEP after corporate royalty and overhead. "
                "Bonus comes out of safe net, so the size of safe net controls how much bonus the agency can afford.")
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:F2')

    r = 4
    ws.cell(row=r, column=1, value='HOW COMMISSION ACTUALLY WORKS (this is the part most people get wrong)').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1

    facts = [
        ('Most carriers pay commission on the WRITTEN premium upfront (advance commission).', 'Example: $1,500 policy at 10% advance commission. Agency receives $150 upfront, even if the customer only paid $150 down.'),
        ('Some carriers pay AS EARNED (only on the premium actually collected from the customer).', 'On those carriers there is no chargeback risk because nothing was advanced in the first place. Net result is similar.'),
        ('On advance-commission carriers, if the customer stops paying, the carrier reverses the unearned commission.', 'Same example: customer pays only $150, then walks. Carrier reverses ~$135 of the $150 advance. Net kept: ~$15.'),
        ('Commission RATES vary 8%-15% by carrier.', 'Bristol West 8%, NATIONAL GENERAL 10%, United Auto 12%, Sterling MGA 13%, Geico/Kemper 15%. The workbook uses an 11% blended rate - conservative.'),
        ('Either way, the commission we actually KEEP nets out to roughly: collected x commission rate.', 'That is the formula in the workbook. The advance + chargeback model and the as-earned model produce the same long-run number.'),
        ('Higher down payment / PIF reduces chargeback risk and lifts the safe net.', 'PIF (100% collected) = no chargeback at all. Bigger down payments raise the safe net the agent contributes to, which raises the effective margin behind every bonus dollar.'),
    ]
    for h, t in facts:
        ws.cell(row=r, column=1, value=h).font = Font(bold=True, size=11)
        c = ws.cell(row=r, column=2, value=t)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 38
        r += 1

    r += 1
    ws.cell(row=r, column=1, value='STEP-BY-STEP: $1,500 policy, customer pays $150 down (10% collected), 10% carrier commission').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1

    headers = ['Step', 'What happens', 'Math', 'Dollar amount', 'Running total', 'Plain English']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    steps = [
        ('1', 'Policy is written', 'Premium $1,500', '$1,500', '$0 in our pocket yet', 'The customer is on the hook for $1,500. Most of it is owed on payments.'),
        ('2', 'Carrier pays commission on WRITTEN premium', '$1,500 x 10%', '$150', '$150 advance commission', 'This is an ADVANCE. The carrier expects to collect the full $1,500.'),
        ('3', 'Customer pays $150 down (10% collected)', '$150 cash in', '$150', '$150 cash + $150 commission', 'Down payment is just the customer paying their bill. Not commission.'),
        ('4', 'If customer pays the rest, no chargeback', 'Best case', 'Keep $150', '$150 commission kept', 'Carrier earned the full premium, we keep the full commission. Renewal still pays.'),
        ('5', 'If customer cancels after only paying $150', 'Carrier reverses unearned commission', '-$135', '$15 commission kept', 'Carrier earned 10% of the $150 collected, the other $135 of commission is reversed.'),
        ('6', 'EXPECTED commission ~= collected x rate', '$150 collected x 10%', '$15 expected to keep', 'Use this for safe net', 'That is why the workbook uses collected x rate. It is the cash we EXPECT to retain.'),
        ('7', 'Royalty to corporate (21%)', '$15 x 21%', '-$3.15', '$11.85 retained', 'Franchise royalty comes off the top.'),
        ('8', 'Overhead reserve (30%)', '$11.85 x 30%', '-$3.56', '$8.30 SAFE NET', 'Rent, tech, admin, bank fees. Salaries tracked separately via the minimum.'),
        ('9', 'Bonus review threshold (40% of safe net)', '$8.30 x 40%', '$3.32 review cap', 'Bonus targets above $2.97 on this policy get flagged for ownership review.', "Soft cap, not auto-cut."),
    ]
    for row in steps:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = SUB_FILL
        ws.row_dimensions[r].height = 36
        r += 1

    r += 1
    ws.cell(row=r, column=1, value='WORKED EXAMPLE - Abel Guaina, January 2026 (full month)').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1

    headers2 = ['Step', 'Formula', 'Plug-in', 'Output', 'What it means', '']
    for i, h in enumerate(headers2, 1):
        if h: style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    d = AGENTS['Abel Guaina']['January']
    total_written = d['NB'][1] + d['RWR'][1] + d['REN'][1]
    total_col = d['NB'][2] + d['RWR'][2] + d['REN'][2]
    advance = total_written * BLENDED_COMM
    expected_retained = total_col * BLENDED_COMM
    after_royalty = expected_retained * (1 - ROYALTY)
    safe_net = after_royalty * (1 - OVERHEAD)
    cap = safe_net * CAP_STANDARD

    abel_steps = [
        ('1. Total written premium', 'NB + RWR + REN written', f"${d['NB'][1]:,.0f} + ${d['RWR'][1]:,.0f} + ${d['REN'][1]:,.0f}", f"${total_written:,.2f}", "What Abel sold and renewed this month"),
        ('2. Carrier commission ADVANCE', 'Written x 11% blended', f"${total_written:,.2f} x 0.11", f"${advance:,.2f}", "Carriers paid us this much commission UPFRONT"),
        ('3. Total collected', 'Down payments + installments collected', f"${d['NB'][2]:,.0f} + ${d['RWR'][2]:,.0f} + ${d['REN'][2]:,.0f}", f"${total_col:,.2f}", "Cash that actually came in from customers"),
        ('4. Expected retained commission', 'Collected x 11% (chargeback math)', f"${total_col:,.2f} x 0.11", f"${expected_retained:,.2f}", "Net of expected chargebacks. Lower than the advance."),
        ('5. After royalty', 'Retained x (1 - 21%)', f"${expected_retained:,.2f} x 0.79", f"${after_royalty:,.2f}", "After 21% corporate royalty"),
        ('6. SAFE NET', 'After-royalty x (1 - 30% overhead)', f"${after_royalty:,.2f} x 0.70", f"${safe_net:,.2f}", "Available for bonus + profit + taxes. Overhead is rent/tech/admin only - salaries tracked via the minimum."),
        ('7. Review threshold', 'Safe net x 40%', f"${safe_net:,.2f} x 0.40", f"${cap:,.2f}", "Bonus targets above this trigger ownership review (soft flag)"),
        ('8. New plan target paid', '(from plan rules)', "Abel's January NB/REN/RWR under the new bonus plan", "$252.00", f"That is {252/safe_net*100:.0f}% of safe net - under threshold, no flag"),
    ]
    for row in abel_steps:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = SUB_FILL
        ws.row_dimensions[r].height = 36
        r += 1

    r += 1
    ws.cell(row=r, column=1, value='KEY TAKEAWAYS').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1
    keys = [
        "Commission is an ADVANCE on written premium. We keep ~commission-on-collected once chargebacks settle out.",
        "Down payment / PIF do NOT change the commission RATE. They reduce CHARGEBACK risk, which makes more of the advance permanent.",
        "Safe net is what is LEFT after royalty and overhead reserve. It is the pool for bonus + agency profit + taxes.",
        f"Under TODAY'S plan, the 6 agents collectively receive ${12220:,.0f} of bonus on ${21285:,.0f} of safe net = ~57% of safe net. That is heavy. All three new proposals come in well below 57%.",
        "The 40% review threshold is a soft flag, not an auto-cut. The 90-day chargeback on the bonus itself is the real profit shield.",
    ]
    for t in keys:
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=11)
        c.alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 40
        r += 1

    set_col_widths(ws, [26, 32, 28, 22, 38, 14])


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
    ws.cell(row=r, column=1, value='HOW THESE NUMBERS WERE PICKED - DATA-DRIVEN').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    r += 1

    ws.cell(row=r, column=1, value=(
        f"Looking at the 24 agent-months in the source data (6 agents x 4 months): "
        f"NB premium per month ranges $7.7k-$68.5k (median $30k). REN premium per month ranges $0-$19.6k (median $3.7k). "
        f"In the 50%-swap scenario, REN climbs to a $11.8k-$54.4k range (median $22.8k). "
        f"The proposed gates are NB ${NB_MIN_PREMIUM:,} and REN ${REN_MIN_PREMIUM:,} - chosen so the NB gate is achievable "
        f"on today's behavior (71% of real months pass) and the REN gate is achievable once behavior shifts (96% of swap months pass)."
    )).alignment = Alignment(wrap_text=True, vertical='top')
    ws.cell(row=r, column=1).font = Font(size=11)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    ws.row_dimensions[r].height = 60
    r += 2

    sal_headers = ['Step', 'Math', 'Number', 'What it means', '', '', '', '', '', '']
    for i, h in enumerate(sal_headers, 1):
        if h: style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    written_floor = NB_MIN_PREMIUM + REN_MIN_PREMIUM
    collected_25 = written_floor * 0.25
    collected_50 = written_floor * 0.50
    collected_75 = written_floor * 0.75
    comm_25 = collected_25 * BLENDED_COMM * (1 - ROYALTY)
    comm_50 = collected_50 * BLENDED_COMM * (1 - ROYALTY)
    comm_75 = collected_75 * BLENDED_COMM * (1 - ROYALTY)
    typical_salary = AGENT_SALARY  # $3,300/mo midpoint of $16-$22/hr range

    sal_steps = [
        ('1', 'Agent salary basis', f"$16-$22/hr x 40 hr/wk = $2,773-$3,813/mo", f"Midpoint ${typical_salary:,}/mo - this is what the agent must generate to cover their own cost."),
        ('2', 'NB minimum + REN minimum', f"${NB_MIN_PREMIUM:,} + ${REN_MIN_PREMIUM:,} = ${written_floor:,} written", "Total written premium needed to clear both gates each month."),
        ('3', 'At 25% collected (today\'s typical)', f"${written_floor:,} x 25% x 11% x (1-21%) = ${comm_25:,.0f}", f"~{comm_25/typical_salary*100:.0f}% of salary. NOT enough on its own. Production needs collection to grow."),
        ('4', 'At 50% collected (push the down payment)', f"${written_floor:,} x 50% x 11% x (1-21%) = ${comm_50:,.0f}", f"~{comm_50/typical_salary*100:.0f}% of salary. Getting close."),
        ('5', 'At 75% collected (high performers)', f"${written_floor:,} x 75% x 11% x (1-21%) = ${comm_75:,.0f}", f"~{comm_75/typical_salary*100:.0f}% of salary. Agent covers themselves AND generates profit."),
        ('6', 'Plus lifetime renewal value', '~30% of NB renews -> repeat commission next year', "A $40k well-collected NB book this year becomes ~$12k of guaranteed renewal commission next year. Book-building is the long-term math."),
        ('7', 'Therefore', '-', "The minimum is the PRODUCTION FLOOR. The KICKER (collected % bonus) is the PROFITABILITY LEVER. Hitting both unlocks the bonus AND makes the agent self-funding."),
    ]
    for row in sal_steps:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = SUB_FILL
        ws.row_dimensions[r].height = 36
        r += 1
    r += 1

    # Pass-rate sensitivity at the chosen threshold
    ws.cell(row=r, column=1, value=f'PASS RATES AT THE PROPOSED MINIMUMS (${NB_MIN_PREMIUM:,} NB / ${REN_MIN_PREMIUM:,} REN)').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    r += 1
    pass_headers = ['Gate', 'REAL data', 'SWAP data (50% RWR→REN)', 'Read', '', '', '', '', '', '']
    for i, h in enumerate(pass_headers, 1):
        if h: style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    # Compute pass rates
    nb_real_pass = sum(1 for a in AGENT_NAMES for m in MONTHS if AGENTS[a][m]['NB'][1] >= NB_MIN_PREMIUM)
    ren_real_pass = sum(1 for a in AGENT_NAMES for m in MONTHS if AGENTS[a][m]['REN'][1] >= REN_MIN_PREMIUM)
    rwr_real_pass = sum(1 for a in AGENT_NAMES for m in MONTHS if AGENTS[a][m]['NB'][1] >= NB_MIN_PREMIUM and AGENTS[a][m]['REN'][1] >= REN_MIN_PREMIUM)
    nb_swap_pass = sum(1 for a in AGENT_NAMES for m in MONTHS if swap_rwr_to_ren(AGENTS[a][m], 0.5)['NB'][1] >= NB_MIN_PREMIUM)
    ren_swap_pass = sum(1 for a in AGENT_NAMES for m in MONTHS if swap_rwr_to_ren(AGENTS[a][m], 0.5)['REN'][1] >= REN_MIN_PREMIUM)
    rwr_swap_pass = sum(1 for a in AGENT_NAMES for m in MONTHS if swap_rwr_to_ren(AGENTS[a][m], 0.5)['NB'][1] >= NB_MIN_PREMIUM and swap_rwr_to_ren(AGENTS[a][m], 0.5)['REN'][1] >= REN_MIN_PREMIUM)

    pass_rows = [
        ('NB gate', f"{nb_real_pass}/24 ({nb_real_pass/24*100:.0f}%)", f"{nb_swap_pass}/24 ({nb_swap_pass/24*100:.0f}%)", "NB gate is achievable on today's behavior. Most agents earn NB bonus most months."),
        ('REN gate', f"{ren_real_pass}/24 ({ren_real_pass/24*100:.0f}%)", f"{ren_swap_pass}/24 ({ren_swap_pass/24*100:.0f}%)", "REN gate is gated by today's almost-zero renewals. Once agents shift behavior, almost all months pass."),
        ('RWR gate (needs both)', f"{rwr_real_pass}/24 ({rwr_real_pass/24*100:.0f}%)", f"{rwr_swap_pass}/24 ({rwr_swap_pass/24*100:.0f}%)", "RWR pay unlocks alongside REN. The lever is the behavior change."),
    ]
    for row in pass_rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = SUB_FILL
        ws.row_dimensions[r].height = 36
        r += 1
    r += 1

    ws.cell(row=r, column=1, value='RULES').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    r += 1
    headers = ['Bonus line', 'Gate', 'If gate passes', 'If gate fails']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1
    rows = [
        ('NB bonus line', f"NB written premium >= ${NB_MIN_PREMIUM:,} this month", 'Pay the NB tier bonus ($250-$1,000+)', 'NB bonus = $0 this month'),
        ('REN bonus line', f'Retention rate >= {REN_MIN_RETENTION*100:.0f}% of agent assigned book this month', 'Pay the REN tier bonus ($200-$800+, if REN premium >= $25k)', 'REN bonus = $0 this month'),
        ('RWR bonus line', 'NB gate passes', 'Pay the RWR tier bonus ($125-$500+, half of NB at same tier)', 'RWR bonus = $0 this month - even with 50 rewrites'),
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


def build_why_current_drops(wb):
    """Answer the question: 'why does current bonus drop when we switch RWR to REN, if REN pays more than RWR?'"""
    ws = wb.create_sheet('Why Current Drops')
    ws['A1'] = "Why does Current bonus DROP when RWR converts to REN?"
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:G1')

    ws['A2'] = ("Short answer: the CURRENT plan does NOT pay anything for renewals. It only counts NB + RWR "
                "toward the 35-policy tier. So when RWRs move to RENs, the count drops, the tier drops, the bonus drops. "
                "Renewals paying 'more' than rewrites is true in the NEW plans - in the current plan renewals pay $0.")
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:G2')

    r = 4
    ws.cell(row=r, column=1, value='THE CURRENT PLAN AND HOW IT BREAKS UNDER THE SWAP').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    r += 1

    cur_facts = [
        "Current rule: count NB + RWR. Below 30 = $0. 30-34 = $250. 35-49 = $350. 50+ = $450 + $10 per policy above 50.",
        "Renewals (REN) are NOT counted. They pay $0 regardless of how many you keep.",
        "If we shift 50% of RWR -> REN, the NB+RWR count drops. The tier drops. The bonus drops.",
        "Even though the agent has MORE renewals (a good outcome), the plan does not reward it.",
        "That is the broken incentive: today the agency pays MORE for churning the book than for retaining it.",
    ]
    for t in cur_facts:
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=12)
        c.alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        ws.row_dimensions[r].height = 28
        r += 1
    r += 1

    # Show one concrete agent example side-by-side
    ws.cell(row=r, column=1, value='CONCRETE EXAMPLE - Dialinerys Dieguez, March 2026').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    r += 1

    headers = ['Scenario', 'NB', 'RWR', 'REN', 'NB+RWR count', 'Current tier rule', 'Current PAYS']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    real_d = AGENTS['Dialinerys Dieguez']['March']
    swap_d = swap_rwr_to_ren(real_d, 0.5)
    nb_r, rwr_r, ren_r = real_d['NB'][0], real_d['RWR'][0], real_d['REN'][0]
    nb_s, rwr_s, ren_s = swap_d['NB'][0], swap_d['RWR'][0], swap_d['REN'][0]
    cur_r = current_bonus(nb_r, rwr_r)
    cur_s = current_bonus(nb_s, rwr_s)

    rows = [
        ('REAL (today)', nb_r, rwr_r, ren_r, nb_r+rwr_r, f"{nb_r+rwr_r} policies -> ${450+10*max(0,nb_r+rwr_r-50)}+ tier", f"${cur_r}"),
        ('SWAP (50% RWR -> REN)', nb_s, rwr_s, ren_s, nb_s+rwr_s, f"{nb_s+rwr_s} policies -> ${450+10*max(0,nb_s+rwr_s-50)}+ tier", f"${cur_s}"),
        ('DELTA', '', f"-{rwr_r-rwr_s}", f"+{ren_s-ren_r}", f"-{(nb_r+rwr_r)-(nb_s+rwr_s)}", f"Lower tier", f"-${cur_r-cur_s}"),
    ]
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = SUB_FILL
            if row[0] == 'DELTA':
                c.font = Font(bold=True, color='9C0006')
                c.fill = WARN_FILL
        r += 1
    r += 1

    # And under the new plan
    ws.cell(row=r, column=1, value='SAME EXAMPLE UNDER THE NEW BONUS PLAN - the renewal NOW pays').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    r += 1

    headers2 = ['Scenario', 'NB tier bonus', 'RWR tier bonus', 'REN tier bonus', 'Total target', 'WITH MIN paid', '']
    for i, h in enumerate(headers2, 1):
        if h: style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    rb = calc_proposal_a(real_d['NB'], real_d['RWR'], real_d['REN'])
    sb = calc_proposal_a(swap_d['NB'], swap_d['RWR'], swap_d['REN'])

    new_rows = [
        ('REAL (today)', f"${rb['nb_target']:.0f}", f"${rb['rwr_target']:.0f}", f"${rb['ren_target']:.0f}", f"${rb['paid']:.0f}", f"${rb['paid_after_min']:.0f}"),
        ('SWAP (50% RWR -> REN)', f"${sb['nb_target']:.0f}", f"${sb['rwr_target']:.0f}", f"${sb['ren_target']:.0f}", f"${sb['paid']:.0f}", f"${sb['paid_after_min']:.0f}"),
        ('DELTA', '', f"-${rb['rwr_target']-sb['rwr_target']:.0f}", f"+${sb['ren_target']-rb['ren_target']:.0f}", f"+${sb['paid']-rb['paid']:.0f}", f"+${sb['paid_after_min']-rb['paid_after_min']:.0f}"),
    ]
    for row in new_rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if r % 2 == 0: c.fill = SUB_FILL
            if row[0] == 'DELTA':
                c.font = Font(bold=True, color='006100')
                c.fill = PatternFill('solid', fgColor='C6EFCE')
        r += 1
    r += 1

    ws.cell(row=r, column=1, value='THE BIG PICTURE').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    r += 1
    big = [
        "Under the CURRENT plan, the agency literally pays LESS when the agent does the right thing (retain instead of rewrite).",
        "Under the new bonus plan, the agent gets MORE for the same behavior change - because REN has its own premium tier ladder ($200-$800) that pays once retention rate hits 30%, instead of paying $0 like today.",
        "Per-policy comparison: RWR pays $2 in new plans; REN pays $4-$6. Net per swapped policy: +$2 to +$4 of bonus. Across hundreds of policies, that adds up.",
        "Bottom line: the swap is not the cause of the drop. The current plan's design is. Switching plans fixes it.",
    ]
    for t in big:
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=12)
        c.alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        ws.row_dimensions[r].height = 30
        r += 1

    set_col_widths(ws, [22, 12, 12, 12, 16, 28, 16])


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
                   'B no-min', 'B WITH MIN', 'B vs Previous', 'Gates']
        for i, h in enumerate(headers, 1):
            style_header(ws.cell(row=r, column=i, value=h))
        r += 1

        prev_total = a_nomin = a_min = b_nomin = b_min = 0
        for month in MONTHS:
            d = AGENTS[agent][month]
            cur = current_bonus(d['NB'][0], d['RWR'][0])
            a = calc_proposal_a(d['NB'], d['RWR'], d['REN'])
            b = calc_proposal_b(d['NB'], d['RWR'], d['REN'])
            gates = f"NB:{'Y' if a['nb_qual'] else 'n'} REN:{'Y' if a['ren_qual'] else 'n'} RWR:{'Y' if a['rwr_qual'] else 'n'}"
            vals = [month, d['NB'][0], d['NB'][1], d['REN'][0], d['REN'][1], d['RWR'][0],
                    cur, a['paid'], a['paid_after_min'], a['paid_after_min'] - cur,
                    b['paid'], b['paid_after_min'], b['paid_after_min'] - cur, gates]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (2, 4, 6): style_int(c)
                elif i in (3, 5): style_dollar(c)
                elif i == 14: style_data(c)
                else: style_dollar(c)
                if i == 7: c.fill = CURRENT_FILL
                if i in (8, 9): c.fill = PROP_A_FILL
                if i == 9: c.font = Font(bold=True)
                if i in (11, 12): c.fill = PROP_B_FILL
                if i == 12: c.font = Font(bold=True)
            prev_total += cur
            a_nomin += a['paid']
            a_min += a['paid_after_min']
            b_nomin += b['paid']
            b_min += b['paid_after_min']
            r += 1

        ws.cell(row=r, column=1, value='4-Month Total').font = Font(bold=True)
        ws.cell(row=r, column=7, value=prev_total)
        ws.cell(row=r, column=8, value=a_nomin)
        ws.cell(row=r, column=9, value=a_min)
        ws.cell(row=r, column=10, value=a_min - prev_total)
        ws.cell(row=r, column=11, value=b_nomin)
        ws.cell(row=r, column=12, value=b_min)
        ws.cell(row=r, column=13, value=b_min - prev_total)
        for i in range(1, 15):
            c = ws.cell(row=r, column=i)
            c.fill = SUB_FILL
            if i in (7, 8, 9, 10, 11, 12, 13):
                style_dollar(c)
                c.font = Font(bold=True)
        r += 1

        # Plain-English read for this agent
        ws.cell(row=r, column=1, value=f"Read: under PREVIOUS plan {agent.split()[0]} earned ${prev_total:,.0f} in 4 months. Under NEW Proposal B WITHOUT minimums it would be ${b_nomin:,.0f}. WITH ${NB_MIN_PREMIUM:,}/${REN_MIN_PREMIUM:,} minimums it is ${b_min:,.0f}.").alignment = Alignment(wrap_text=True)
        ws.cell(row=r, column=1).font = Font(italic=True, size=10, color='666666')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
        ws.row_dimensions[r].height = 30
        r += 2

    set_col_widths(ws, [12, 8, 11, 8, 11, 8, 12, 11, 12, 13, 11, 12, 13, 20])


def build_rules_side_by_side(wb):
    ws = wb.create_sheet('Rules Side by Side')
    ws['A1'] = 'Rules: Current vs Proposal A vs Proposal B'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:D1')

    ws['A2'] = 'One-row-per-rule view of all three plans. A and B are the main proposals (do NOT mix them).'
    ws['A2'].font = Font(italic=True, size=11, color='1F4E78')
    ws.merge_cells('A2:D2')

    r = 4
    headers = ['Rule', 'Current', 'Proposal A (Premium)', 'Proposal B (Coverage)']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    real = {'cur': 0, 'a': 0, 'a_min': 0, 'b': 0, 'b_min': 0}
    swap = {'cur': 0, 'a': 0, 'a_min': 0, 'b': 0, 'b_min': 0}
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

    rows = [
        ('Bonus base', 'Count tier (NB+RWR)', 'Written premium tier', 'Coverage type'),
        ('NB - low premium (< $1,200)', 'Count toward tier', '$6', '$11 (base coverage)'),
        ('NB - medium premium', 'Count toward tier', '$8-$13 (tier)', '$11 (base) + $6 if BI+UM'),
        ('NB - high premium ($3,000+)', 'Count toward tier', '$13 + $2/$1k, cap $28', '$11 (base) + $6 if BI+UM'),
        ('REN - any', 'PAYS NOTHING', '$5 / $6 / $7 by premium tier', '$7 (base) + $3 if BI+UM'),
        ('RWR - any', 'Count toward tier (BAD)', '$2 flat', '$2 flat'),
        ('Coverage detail', 'Not tracked', 'Not paid separately', 'BI+UM bundle = +$6 NB / +$3 REN'),
        ('PIF add', 'Not paid', '+$9 NB / +$13 high NB / +$5 REN', '+$9 NB / +$13 high NB / +$5 REN'),
        ('Collected % kicker', 'NONE', '+10% / +15% / +20% / +25%', 'Same as A'),
        ('Monthly minimum (NEW)', '35 policies NB+RWR', f'NB >= ${NB_MIN_PREMIUM/1000:.0f}k AND REN >= ${REN_MIN_PREMIUM/1000:.0f}k', 'Same as A'),
        ('RWR gate', 'No - RWR counts toward 35', 'RWR paid only if BOTH NB+REN gates pass', 'Same as A'),
        ('Profit protection', 'None per-policy', '40% safe net review + 90-day chargeback + minimums', 'Same as A'),
        ('Best for', 'Status quo only', 'Simple payroll, premium-focused', 'Coverage upsell, agent motivation'),
        ('Tradeoff', 'Rewards churning', 'Less generous than B', 'More generous, requires coverage tracking'),
        ('Recommended decision', 'Phase out', 'Conservative choice', 'PILOT (recommended)'),
        ('4-month cost - REAL, no min', f"${real['cur']:,.0f}", f"${real['a']:,.0f}", f"${real['b']:,.0f}"),
        ('4-month cost - REAL, WITH min', f"${real['cur']:,.0f}", f"${real['a_min']:,.0f}", f"${real['b_min']:,.0f}"),
        ('4-month cost - SWAP, no min', f"${swap['cur']:,.0f}", f"${swap['a']:,.0f}", f"${swap['b']:,.0f}"),
        ('4-month cost - SWAP, WITH min', f"${swap['cur']:,.0f}", f"${swap['a_min']:,.0f}", f"${swap['b_min']:,.0f}"),
    ]
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if i == 2: c.fill = CURRENT_FILL
            elif i == 3: c.fill = PROP_A_FILL
            elif i == 4: c.fill = PROP_B_FILL
        ws.row_dimensions[r].height = 30
        r += 1

    set_col_widths(ws, [32, 26, 34, 38])


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
    ex_headers = ['Scenario', 'Proposal A pays', 'Proposal B pays']
    for i, h in enumerate(ex_headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    examples = [
        ('NB $1,000 PIP+PD, 25% down', '$6 x 1.15 = $6.90', '$11 x 1.15 = $12.65'),
        ('NB $2,000 with BI+UM, 25% down', '$11 x 1.15 = $12.65', '($11 + $6) x 1.15 = $19.55'),
        ('NB $2,500 PIF with BI+UM', '($13 + $9 PIF) x 1.25 = $27.50', '($11 + $6 + $9 PIF) x 1.25 = $32.50'),
        ('Renewal $1,500 with BI+UM, 25%', '$6 x 1.15 = $6.90', '($7 + $3) x 1.15 = $11.50'),
        ('Rewrite $1,200', '$2 x 1.15 = $2.30', '$2 x 1.15 = $2.30'),
    ]
    for ex in examples:
        for i, v in enumerate(ex, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if i == 1: c.fill = SUB_FILL
            elif i == 2: c.fill = PROP_A_FILL
            elif i == 3: c.fill = PROP_B_FILL
        ws.row_dimensions[r].height = 28
        r += 1

    set_col_widths(ws, [34, 36, 38])


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
        ('Royalty rate', '21%', 'Franchise royalty removed before safe net.', 'Standard Fiesta franchise royalty.'),
        ('Overhead reserve', '30%', 'Rent, tech, admin, bank fees (NOT salaries).', 'Lowered from 40% because salaries are now tracked via the minimum requirement.'),
        ('Agent salary (basis for minimum)', '$3,300/mo midpoint', '$16-$22/hour x 40 hours/week = $2,773-$3,813/mo.', 'Florida P&C agent baseline.'),
        ('Blended carrier commission', '11%', 'Used for the safe-net calculation.', 'Range 8% (Bristol West) to 15% (Geico/Kemper). 11% is the book-weighted blend.'),
        ('Commission model', 'Advance + chargeback OR as-earned', 'Most carriers advance on written, charge back if unpaid. Net = collected x rate.', 'Some carriers (e.g., some MGAs) pay as-earned on collected.'),
        ('Bonus paid', 'TARGET (no automatic cap)', 'Pay equals the calculated target. Soft review threshold only.', '-'),
        ('Review threshold (standard)', '40% of safe net', 'If monthly target > 40% safe net, ownership reviews.', 'Most months stay below.'),
        ('Review threshold (PIF)', '55% of safe net', 'Higher because PIF has no chargeback risk.', '-'),
        ('Chargeback period (PRIMARY PROTECTION)', '90 days', 'Cancel/rewrite reversal window. The actual profit shield.', 'Matches carrier commission chargeback exposure.'),
        ('NB tier ladder', '$250 / $375 / $525 / $725 / $1,000', 'Monthly NB written premium tiers at $45k / $55k / $70k / $85k / $100k.', 'Mirrors today\'s 30/38/50 count-tier idea but premium-based. Bigger jumps at the top tiers.'),
        ('NB above $100k', '+$50 per $5k extra', '1% linear above the top tier. No cap.', '$110k -> $1,100. $150k -> $1,500.'),
        ('RWR tier ladder', '$100 / $200 / $275 / $350 / $500', 'Same breakpoints as NB ($45k-$100k), pays roughly half of NB at each tier.', 'Rewrites pay about half of new business.'),
        ('RWR above $100k', '+$25 per $5k extra', '0.5% linear above the top tier.', '$110k RWR -> $550.'),
        ('REN tier ladder', '$250 / $350 / $450 / $550 / $800', 'Own breakpoints at $25k / $40k / $65k / $80k / $100k REN written premium.', 'Pays roughly the intersection between NB and RWR, with a slight boost.'),
        ('REN above $100k', '+$40 per $5k extra', '0.8% linear above the top tier.', '$150k REN -> $1,200. Rewards growing the renewal book past $100k.'),
        ('NB minimum', f'${NB_MIN_PREMIUM:,}/mo NB written premium', 'Equal to NB T1 entry. No bonus below this.', 'Same threshold gates RWR.'),
        ('REN minimum', f'{REN_MIN_RETENTION*100:.0f}% retention rate', 'Agent must retain at least 30% of their assigned renewal book.', 'Computed from agency renewal history.'),
        ('RWR gate', 'NB minimum met', 'RWR pays only when NB premium >= $45k. No separate RWR floor.', 'Encourages NB writing - RWR alone never earns bonus.'),
        ('Source data', 'agents_nbrwr_pivots_good_.xlsx', 'NB / RWR / REN counts and premiums for Jan-Apr 2026.', '6 agents: Abel, Dialinerys, Melissa, Flavia, Thalia, Monica.'),
        ('50% Swap scenario', '50% RWR to REN', 'Half of rewrites reclassified as renewals.', 'Premium and collected $ stay the same.'),
        ('100% Flip scenario', 'RWR <-> REN', 'Rewrites and renewals fully swapped - "if you had been renewing all along".', 'The strongest behavior-change test.'),
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
    build_proposal_a(wb)                  # The Bonus Plan rules
    build_bonus_examples(wb)              # 3 sheets: NB / RWR / REN examples
    build_agent_examples_flip(wb)         # Per-agent under 100% FLIP (cleanest view)
    build_chargeback(wb)
    build_manual_tracker(wb)
    build_source_data(wb)

    out = '/home/user/fiesta-bonus/output/Bonus_Plan_SIMPLE.xlsx'
    wb.save(out)
    print(f"Saved: {out}")
    return out


if __name__ == '__main__':
    build()
