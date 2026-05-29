"""Backtest the boss's proposed comp plan against the current plan
using the same 6 agents x 4 months from Bonus_Plan_SIMPLE."""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from build_bonus_workbook import (
    AGENTS, AGENT_NAMES, MONTHS, current_bonus,
    TITLE_FONT, SECTION_FONT, SECTION_FILL, SUB_FILL, CURRENT_FILL, PROP_A_FILL,
    style_header, style_data, style_dollar, style_int, set_col_widths,
)

# ============================================================================
# BOSS'S PROPOSED PLAN
# (from FIESTACOMPENSATIONPLANMASTERWITHBASELINE_with_440.xlsx)
# ============================================================================
PROPOSAL_TIERS = {
    'Trainee':    {'base': 34000, 'nb_pct': 0.05, 'rn_pct': 0.000, 'rwr_pct': 0.00, 'min_nb_count': 0},
    'T1 Builder': {'base': 36000, 'nb_pct': 0.08, 'rn_pct': 0.015, 'rwr_pct': 0.00, 'min_nb_count': 25},
    'T2 Senior':  {'base': 40000, 'nb_pct': 0.12, 'rn_pct': 0.025, 'rwr_pct': 0.02, 'min_nb_count': 50},
    'T3 Elite':   {'base': 45000, 'nb_pct': 0.16, 'rn_pct': 0.035, 'rwr_pct': 0.02, 'min_nb_count': 80},
    'Elite Top':  {'base': 50000, 'nb_pct': 0.20, 'rn_pct': 0.050, 'rwr_pct': 0.02, 'min_nb_count': 110},
}
NB_GATE = 45000
CARRIER_NB_COMM = 0.10
CARRIER_REN_COMM = 0.08
CARRIER_RWR_COMM = 0.08


def collected_kicker_per_policy(coll_pct):
    if coll_pct >= 1.00: return 8
    if coll_pct >= 0.25: return 5
    if coll_pct >= 0.20: return 2
    if coll_pct >= 0.15: return 1
    return 0


def assign_tier(avg_nb):
    if avg_nb >= 110: return 'Elite Top'
    if avg_nb >=  80: return 'T3 Elite'
    if avg_nb >=  50: return 'T2 Senior'
    if avg_nb >=  25: return 'T1 Builder'
    return 'Trainee'


def compute_proposal_bonus(agent_data, tier):
    """Compute the boss-proposed monthly bonus for one agent-month."""
    t = PROPOSAL_TIERS[tier]
    nb_c, nb_p, nb_col = agent_data['NB']
    rwr_c, rwr_p, _ = agent_data['RWR']
    ren_c, ren_p, _ = agent_data['REN']

    gate_met = nb_p >= NB_GATE
    if not gate_met:
        return {
            'gate_met': False,
            'nb_var': 0, 'rn_var': 0, 'rwr_var': 0,
            'kicker': 0, 'total': 0,
        }

    nb_var = nb_p * CARRIER_NB_COMM * t['nb_pct']
    rn_var = ren_p * CARRIER_REN_COMM * t['rn_pct']
    rwr_var = rwr_p * CARRIER_RWR_COMM * t['rwr_pct']
    nb_coll_pct = nb_col / nb_p if nb_p else 0
    kicker = nb_c * collected_kicker_per_policy(nb_coll_pct)
    return {
        'gate_met': True,
        'nb_var': nb_var, 'rn_var': rn_var, 'rwr_var': rwr_var,
        'kicker': kicker, 'total': nb_var + rn_var + rwr_var + kicker,
    }


def assign_all_tiers():
    """Assign tier per agent based on 4-month avg NB count."""
    result = {}
    for agent in AGENT_NAMES:
        avg_nb = sum(AGENTS[agent][m]['NB'][0] for m in MONTHS) / 4
        result[agent] = (assign_tier(avg_nb), avg_nb)
    return result


GREEN_FILL = PatternFill('solid', fgColor='00B050')
RED_FILL = PatternFill('solid', fgColor='C00000')
GREY_FILL = PatternFill('solid', fgColor='D9D9D9')
GATE_PASS_FILL = PatternFill('solid', fgColor='C6EFCE')
GATE_MISS_FILL = PatternFill('solid', fgColor='FFC7CE')


def build_summary_sheet(wb, agent_tiers, per_agent):
    ws = wb.create_sheet('Summary')
    ws['A1'] = "Boss Proposal vs Current Plan - 4-Month Backtest"
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:F1')
    ws['A2'] = ("Same 6 agents and 4 months (Jan-Apr 2026) we used for the new-bonus proposal. "
                "Current plan = today's actual count-tier ($250/$350/$450 + $10/policy above 50). "
                "Boss proposal = the producer ladder with $45K NB premium gate + collected kicker.")
    ws['A2'].font = Font(italic=True, size=10, color='666666')
    ws.merge_cells('A2:F2')

    r = 4
    # HEADLINE
    ws.cell(row=r, column=1, value='HEADLINE - 4-MONTH TOTAL (6 agents)').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1

    headers = ['Plan', '4-Month Total', 'vs Current', 'Annualized (x 3)', 'Notes', '']
    for i, h in enumerate(headers, 1):
        if h: style_header(ws.cell(row=r, column=i, value=h))
    r += 1

    cur_total = sum(per_agent[a]['cur'] for a in AGENT_NAMES)
    prop_total = sum(per_agent[a]['prop'] for a in AGENT_NAMES)
    rows = [
        ('CURRENT (paying today)', cur_total, 0, cur_total*3,
         'Count NB+RWR policies. <30 = $0, 30-34 = $250, 35-49 = $350, 50+ = $450 + $10/policy above 50.'),
        ('BOSS PROPOSAL', prop_total, prop_total - cur_total, prop_total*3,
         f'Tier ladder + $45K NB gate + collected kicker. Only {sum(1 for a in AGENT_NAMES for m in MONTHS if AGENTS[a][m]["NB"][1] >= NB_GATE)} of 24 agent-months pass the $45K NB gate.'),
    ]
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            if i == 1: style_data(c); c.font = Font(bold=True)
            elif i == 5: style_data(c)
            else: style_dollar(c)
            if 'CURRENT' in row[0]: c.fill = CURRENT_FILL
            else: c.fill = PROP_A_FILL
        ws.row_dimensions[r].height = 28
        r += 1
    r += 1

    # PER AGENT
    ws.cell(row=r, column=1, value='PER AGENT - 4-MONTH TOTALS').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1
    headers2 = ['Agent', 'Tier (auto-assigned)', 'Avg NB/mo', 'CURRENT', 'BOSS PROPOSAL', 'DELTA']
    for i, h in enumerate(headers2, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1
    for agent in AGENT_NAMES:
        tier, avg_nb = agent_tiers[agent]
        cur = per_agent[agent]['cur']
        prop = per_agent[agent]['prop']
        delta = prop - cur
        ws.cell(row=r, column=1, value=agent).font = Font(bold=True)
        ws.cell(row=r, column=2, value=tier)
        c = ws.cell(row=r, column=3, value=avg_nb)
        c.number_format = '0.0'
        c_cur = ws.cell(row=r, column=4, value=cur); style_dollar(c_cur); c_cur.fill = CURRENT_FILL
        c_prop = ws.cell(row=r, column=5, value=prop); style_dollar(c_prop); c_prop.fill = PROP_A_FILL
        c_delta = ws.cell(row=r, column=6, value=delta); style_dollar(c_delta)
        c_delta.font = Font(bold=True, color='FFFFFF')
        c_delta.fill = GREEN_FILL if delta > 0 else (RED_FILL if delta < 0 else GREY_FILL)
        c_delta.alignment = Alignment(horizontal='center')
        style_data(ws.cell(row=r, column=1))
        style_data(ws.cell(row=r, column=2))
        r += 1
    # Total row
    ws.cell(row=r, column=1, value='TOTAL').font = Font(bold=True)
    for col in range(1, 7): ws.cell(row=r, column=col).fill = SUB_FILL
    c = ws.cell(row=r, column=4, value=cur_total); style_dollar(c); c.font = Font(bold=True)
    c = ws.cell(row=r, column=5, value=prop_total); style_dollar(c); c.font = Font(bold=True)
    c = ws.cell(row=r, column=6, value=prop_total - cur_total); style_dollar(c)
    c.font = Font(bold=True, color='FFFFFF')
    c.fill = GREEN_FILL if prop_total - cur_total > 0 else RED_FILL
    c.alignment = Alignment(horizontal='center')
    r += 2

    # READ
    ws.cell(row=r, column=1, value='READ').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1
    pass_months = sum(1 for a in AGENT_NAMES for m in MONTHS if AGENTS[a][m]['NB'][1] >= NB_GATE)
    reads = [
        f"The boss's proposal pays ${prop_total:,.0f} over 4 months vs today's ${cur_total:,.0f} - a {(prop_total-cur_total)/cur_total*100:.0f}% drop.",
        f"The $45K NB premium gate is the dominant cause: only {pass_months} of 24 agent-months ({pass_months/24*100:.0f}%) clear it. The other {24-pass_months} months pay $0 variable - and the kicker is gated too.",
        f"None of our 6 agents reach T2 Senior (50+ NB sustained 3 months). 2 are T1 Builders (Dialinerys, Monica), 4 are Trainee tier.",
        f"Trainee tier pays 5% of NB agency commission. T1 pays 8%. With $40k NB premium x 10% comm = $4,000 agency comm, T1 variable = ~$320/mo - small dollars vs today's $350-$700 count tier.",
        f"Annualized: today pays ${cur_total*3:,.0f}/yr to these 6, boss proposal would pay ${prop_total*3:,.0f}/yr - savings of ${(cur_total-prop_total)*3:,.0f}/yr.",
        f"Comparison to our designed proposal: REAL scenario pays $1,675 - effectively the SAME outcome as the boss's $1,810. Both gate-driven by $45K NB.",
    ]
    for txt in reads:
        c = ws.cell(row=r, column=1, value=txt)
        c.font = Font(size=11)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 38
        r += 1

    set_col_widths(ws, [24, 22, 12, 14, 16, 14])


def build_month_detail(wb, agent_tiers):
    ws = wb.create_sheet('Month-by-Month Detail')
    ws['A1'] = 'Month-by-Month Bonus Comparison'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:O1')
    ws['A2'] = 'One row per agent-month. Shows the gate result and the bonus pieces that go into the boss-proposed total.'
    ws['A2'].font = Font(italic=True, size=10, color='666666')
    ws.merge_cells('A2:O2')

    r = 4
    headers = ['Agent', 'Tier', 'Month',
               'NB #', 'NB $', 'NB Coll %',
               'RWR #', 'REN $',
               '$45K Gate',
               'NB var', 'RN var', 'RWR var', 'Kicker',
               'BOSS Proposal', 'CURRENT', 'Delta']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    ws.row_dimensions[r].height = 26
    r += 1

    for agent in AGENT_NAMES:
        tier, _ = agent_tiers[agent]
        for month in MONTHS:
            d = AGENTS[agent][month]
            res = compute_proposal_bonus(d, tier)
            cur = current_bonus(d['NB'][0], d['RWR'][0])
            nb_c, nb_p, nb_col = d['NB']
            nb_coll_pct = nb_col / nb_p if nb_p else 0
            rwr_c = d['RWR'][0]
            ren_p = d['REN'][1]

            vals = [agent, tier, month, nb_c, nb_p, nb_coll_pct,
                    rwr_c, ren_p,
                    'PASS' if res['gate_met'] else 'MISS',
                    res['nb_var'], res['rn_var'], res['rwr_var'], res['kicker'],
                    res['total'], cur, res['total'] - cur]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                if i in (1, 2, 3, 9): style_data(c)
                elif i in (4, 7): style_int(c)
                elif i == 6:
                    style_data(c)
                    c.number_format = '0.0%'
                else: style_dollar(c)

            ws.cell(row=r, column=1).font = Font(bold=True)
            ws.cell(row=r, column=9).fill = GATE_PASS_FILL if res['gate_met'] else GATE_MISS_FILL
            ws.cell(row=r, column=9).font = Font(bold=True)
            ws.cell(row=r, column=9).alignment = Alignment(horizontal='center')
            ws.cell(row=r, column=15).fill = CURRENT_FILL
            ws.cell(row=r, column=14).fill = PROP_A_FILL
            delta_cell = ws.cell(row=r, column=16)
            delta = res['total'] - cur
            delta_cell.fill = GREEN_FILL if delta > 0 else (RED_FILL if delta < 0 else GREY_FILL)
            delta_cell.font = Font(bold=True, color='FFFFFF')
            delta_cell.alignment = Alignment(horizontal='center')
            r += 1

    set_col_widths(ws, [22, 12, 9, 6, 10, 10, 6, 9, 9, 9, 8, 9, 8, 13, 10, 11])
    ws.freeze_panes = 'D5'


def build_rules_reference(wb):
    ws = wb.create_sheet('Proposal Rules')
    ws['A1'] = "Boss Proposal Rules (as used in this backtest)"
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:F1')

    r = 3
    ws.cell(row=r, column=1, value='PRODUCER TIER LADDER').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1
    headers = ['Tier', 'Min NB count', 'Base salary', 'NB %', 'RN %', 'RWR %']
    for i, h in enumerate(headers, 1):
        style_header(ws.cell(row=r, column=i, value=h))
    r += 1
    for tier, t in PROPOSAL_TIERS.items():
        vals = [tier, t['min_nb_count'], t['base'], t['nb_pct'], t['rn_pct'], t['rwr_pct']]
        for i, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=i, value=v)
            if i == 1: style_data(c); c.font = Font(bold=True)
            elif i in (2,): style_int(c)
            elif i == 3: style_dollar(c)
            else:
                style_data(c)
                c.number_format = '0.0%'
        r += 1
    r += 1

    # Gates + kicker
    ws.cell(row=r, column=1, value='GATES, KICKER, ASSUMPTIONS').font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1
    rows = [
        ('$45K NB Premium Gate', 'Hard gate', 'Variable comp + kicker paid ONLY if monthly NB written premium >= $45K. Below gate -> $0 bonus.'),
        ('Carrier NB commission rate', '10%', 'Used to compute NB agency commission.'),
        ('Carrier RN commission rate', '8%', 'Used to compute RN agency commission.'),
        ('Carrier RWR commission rate', '8%', 'Used to compute RWR agency commission.'),
        ('Collected Kicker - <15%',  '$0 per policy', 'Per NB policy.'),
        ('Collected Kicker - 15-19%', '+$1 per policy', 'Per NB policy.'),
        ('Collected Kicker - 20-24%', '+$2 per policy', 'Per NB policy.'),
        ('Collected Kicker - 25-99%', '+$5 per policy', 'Per NB policy.'),
        ('Collected Kicker - 100% PIF','+$8 per policy', 'Per NB policy.'),
        ('Tier assignment', '4-month avg NB count', 'In the live plan tiers are sustained over 3-6 months. Backtest uses the 4-month average for each agent.'),
        ('What we DIDNT model', 'PIF / E-Pay / Ancillary spiffs / NSD bonuses', 'No per-agent monthly data for these in the source file. They would add a smaller layer on top.'),
    ]
    headers2 = ['Rule', 'Value', 'How it applies', '', '', '']
    for i, h in enumerate(headers2, 1):
        if h: style_header(ws.cell(row=r, column=i, value=h))
    r += 1
    for row in rows:
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            style_data(c)
            if i == 1: c.font = Font(bold=True)
        r += 1

    set_col_widths(ws, [30, 22, 60, 6, 6, 6])


def main():
    agent_tiers = assign_all_tiers()
    per_agent = {a: {'cur': 0, 'prop': 0} for a in AGENT_NAMES}
    for agent in AGENT_NAMES:
        tier, _ = agent_tiers[agent]
        for month in MONTHS:
            d = AGENTS[agent][month]
            res = compute_proposal_bonus(d, tier)
            cur = current_bonus(d['NB'][0], d['RWR'][0])
            per_agent[agent]['cur'] += cur
            per_agent[agent]['prop'] += res['total']

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_summary_sheet(wb, agent_tiers, per_agent)
    build_month_detail(wb, agent_tiers)
    build_rules_reference(wb)
    out = '/home/user/fiesta-bonus/output/Boss_Proposal_Backtest.xlsx'
    wb.save(out)
    print(f"Saved: {out}")


if __name__ == '__main__':
    main()
