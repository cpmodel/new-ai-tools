"""Word report: bugs found in the legacy CPAT engineer power model and improvements made in CPAT-AI-Mitigation-MVP
v1.06 (PowerModel_BugsAndImprovements_v1_0.docx). The effect of each fix is computed with power_model_ref_v0_1 on the
cached legacy inputs of Egypt (baseline and carbon-tax run), so every number in the report is reproducible.

    python make_bugs_doc_v1_0.py      (needs node with the docx package)
"""
import json
import os
import subprocess
import tempfile

import power_model_ref_v0_2 as PM          # v0.2 = v0.1 + planned retirement (same results here)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'PowerModel_BugsAndImprovements_v1_0.docx')
YEARS = list(range(2022, 2041))
K30, K40 = YEARS.index(2030), YEARS.index(2040)

BUGS = [  # key, title, legacy rows, what legacy does, what it should do, why it matters
    ('services', 'Services electricity demand grows with the industrial multiple after 2023', 'H3, row 3731',
     'The 2023 formula multiplies services demand by the services multiple (row 3713); from 2024 the copied formula '
     'points six rows up, to the industrial multiple (row 3707). Services have a higher GDP elasticity (1.1 against '
     '0.8) and a different price elasticity.',
     'Use the services multiple in every year, as in 2023 and as the label of row 3713 says.',
     'Services are about a quarter of Egypt\'s electricity demand, so this is the largest single effect.'),
    ('aee', 'Autonomous efficiency gains use the efficiency elasticity instead of the usage elasticity',
     'H1, rows 3691, 3697, 3703, 3709',
     'Efficiency multiple = 1 / (1 + alpha)^(1 + eps_F), with eps_F the "Price Elasticity" row 3677.',
     'Documentation equation (3.4.2.3) and the fuel equation use (1 + eps_U): 1 / (1 + alpha)^(1 + eps_U), with eps_U '
     'the usage ("Demand Elasticity") row 3676.',
     'Efficiency gains reduce demand slightly too much every year.'),
    ('income', 'Income factor on the GDP elasticity applied twice and anchored to 2023', 'H1, rows 3692-3710',
     'GDP term = (1 + g)^((eps_Y / $M3686) x M3686 x M2433). Row 3686 already holds the average elasticity x the '
     'income factor (row 2433), and $M fixes the denominator to 2023, so the effective elasticity is '
     'eps_Y x inc(t)^2 / inc(2023).',
     'eps_Y x inc(t): the income factor once, in the year of the calculation.',
     'The income factor falls from 1.03 to 0.95 by 2040, so the error grows over time.'),
    ('lt_alloc', 'Long-term storage cost allocation (33%) applied twice', 'J1, rows 3911 and 3912',
     'Row 3911 multiplies the long-term storage cost by the allocation share; row 3912 multiplies it again.',
     'Apply the 33% allocation once.',
     'No effect for Egypt (variable renewables stay below the 75% threshold); it matters for high-renewable '
     'scenarios.'),
    ('invcost', 'Investment spending uses the capital cost at time factor 1', 'L20, rows 4276-4284',
     'New investment (USD m) = nameplate MW x the capital cost column O of table A2, which is the cost before the '
     'capital-cost time factor.',
     'Use the capital cost of the year (table A13), as the levelised and amortised costs do.',
     'Investment spending on solar is overstated by up to 80% (2040, time factor 0.55), wind by up to 9%; it '
     'is reporting only (the investment decision uses the right cost).'),
    ('cmin', 'Investment logit normalised by the cheapest type, including types that may not be built',
     'L1/L4, rows 4046, 4073-4081',
     'The logit weight is EXP(-K x cost / cheapest cost), and the cheapest cost is taken over all nine types, '
     'including types whose new investment is not permitted (coal, nuclear, other renewables, biomass for Egypt).',
     'Normalise by the cheapest permitted type, so that a type that cannot be built does not change the shares of '
     'those that can.',
     'No effect in the Egypt baseline; in the carbon-tax run nuclear (not permitted) is the cheapest type from '
     '2024, which makes the investment logit steeper.'),
    ('retlag', 'Early coal retirement mixes nameplate and effective capacity', 'K4, row 3973',
     'The room for early coal retirement subtracts last year\'s non-coal retirements in nameplate MW from additions '
     'and demand growth in effective MW.',
     'Subtract effective retirements.',
     'Latent for Egypt (no coal). Too little early coal retirement where solar (low capacity factor) retires.'),
    ('coalret', 'Effective coal retirement is not capped like the nameplate retirement', 'K5/K6, rows 3994, 4009-4011',
     'Nameplate coal retirement is MIN(planned + early, capacity); the effective retirement uses planned + early '
     'without the cap.',
     'Effective retirement = nameplate retirement x capacity factor.',
     'Latent for Egypt. Where early retirement is large, capacity needed is overstated.'),
    ('rebate', 'Feebate rebate per kWh consumed is added to the cost per kWh generated', 'D6, rows 3407-3409',
     'The rebate is the average feebate per kWh generated divided by consumption/generation (x 1.22), and is then '
     'added to the cost per kWh generated of every type.',
     'Rebate per kWh generated, so the feebate is revenue-neutral at base-year shares.',
     'Latent for Egypt (no power feebate in the runs). Over-rebates by 22% when a power feebate is set.'),
    ('kcheck', 'Capacity needed is computed from last year\'s plan, not from actual capacity', 'K7, row 4024',
     'Effective capacity before retirements = last year\'s planned capacity (capacity after retirements + investment '
     'needed), whether or not the investment could be placed.',
     'Use the actual effective capacity of the year (capacity x capacity factor).',
     'Latent: it only matters when scale-up limits leave a deficit; then the model under-invests next year too.'),
]
SHORT = {'services': 'Services demand', 'aee': 'Efficiency exponent', 'income': 'Income factor',
         'lt_alloc': 'LT storage allocation', 'invcost': 'Investment spending', 'cmin': 'Investment logit',
         'retlag': 'Early-retirement units', 'coalret': 'Coal effective retirement', 'rebate': 'Feebate rebate',
         'kcheck': 'Capacity check'}
MINOR = [
    ('Stray reference in dispatch', 'I1, rows 3773 and 3779',
     'Required generation from coal and gas is MAX(required - other generation, $G4291 x total), and G4291 is an '
     'unrelated, empty cell of table L21. Harmless while the cell is empty; not built in the MVP.'),
    ('Total nameplate investment row is zero', 'L19, row 4273',
     'The total multiplies by $X3008, an empty cell, instead of summing the types. Reporting only; the MVP sums.'),
]
MVP_FIXES = [
    'VAT on the power policy term: v1.05 added the policy cost without VAT; legacy row 3614 multiplies it by (1 + VAT). '
    'v1.06 follows legacy (no numerical effect for Egypt, VAT 0).',
    'Long-term storage allocation: v1.05 left the 33% allocation out (it read the legacy note "incomplete" as "not '
    'applied"); v1.06 applies it once.',
    'Variable renewable share: v1.05 counted wind and solar; legacy row 3851 also counts other renewables. v1.06 '
    'follows legacy (other renewables are 0 for Egypt).',
]
DESIGN = [
    ('Capacity lags demand by one year', 'K7, row 4027 "Next Year Demand" uses this year\'s demand; the capacity '
     'built becomes available next year, so capacity is always one year of growth short and gas runs at a higher '
     'capacity factor. Kept as legacy (the check row says the plan works as intended). Recommended: use next year\'s '
     'demand, which is computable without a circular reference because demand uses last year\'s price.'),
    ('Levelised variable cost ends with the model horizon', 'D9 averages variable costs only to the last model year, '
     'so the window shrinks to one year at the end of the horizon. Kept (MVP horizon 2040, legacy 2041). Recommended: '
     'extend the last year\'s cost over the plant lifetime.'),
    ('2023-2024 demand shock', 'Legacy applies an exogenous shock to power demand (-20.8% in 2023, +9.2% in 2024). '
     'The MVP fuel blocks do not apply it, so it is a switch (Settings C19), off by default. The legacy test '
     'runs with it on.'),
    ('Demand responds to USD prices', 'Power demand responds to electricity prices in real USD. In the corrected '
     'Egypt data the residential price falls from 0.075 to 0.054 $/kWh (2021-2022, devaluation), so MVP demand jumps '
     'about 22% in 2023; the fuel blocks show the same jump. This is the open price-vintage decision. Options: switch '
     'the legacy demand shock on (Settings C19), or let demand respond to prices in real local currency.'),
    ('Spot fuel prices', 'Legacy uses a 5-year moving average of fuel prices for power; the MVP uses spot prices '
     '(the legacy option "use spot fuel prices"), so fuel price changes reach power costs earlier.'),
    ('Features not built', 'PPAs, manual capacity and investment, planned coal retirement data, power demand from the '
     'elasticity model, a carbon price on power use, and "apply final deficit to gas" are not built. All are off or 0 '
     'in the Egypt runs.'),
]
IMPROVEMENTS = [
    'One formula per block, readable: the second investment round (legacy tables L11-L17, about 60 rows of ranks, '
    'cumulative sums and lookups) is one row: permit x MAX(0, MIN(space, deficit - space of cheaper types)).',
    'The dispatch of coal and gas (legacy I2, 19 rows) is five rows: minimum, logit weight, generation, plus the '
    'residual and the cheapest cost.',
    'Amortised and levelised fixed costs are one row per type each (legacy B3-B8 and C1-C7), with constants derived '
    'once in the data step (sheet Inputs_power).',
    'Electricity prices are telescoped (price = last price + pass-through x change in supply cost x (1 + VAT)), so '
    'the components held at their last historical value drop out.',
    'Every block has one formula across years, types and scenarios, and the right column calls a named LAMBDA.',
    'Power CO2 is part of section 13, so a new ETS that covers power now sees power emissions and the cap can bind '
    'on them.',
    'Every number is checked: a Python reference reproduces 208 legacy rows of the baseline and the carbon-tax run to '
    '1e-13 with the bugs switched on, and the workbook equals the reference with the bugs fixed.',
]


def effects():
    def kpi(o):
        gen = lambda k: sum(o['gen'][i][k] for i in range(9)) * 8.76 / 1000
        return {'gen30': gen(K30), 'gen40': gen(K40), 'co2_40': sum(o['co2'][i][K40] for i in range(9)),
                'gas40': o['shr'][1][K40], 'vre40': o['vre'][K40],
                'usd': sum(o['usd'][i][k] for i in range(9) for k in range(len(YEARS))) / 1000}
    base = {run: kpi(PM.engine(PM.legacy_inputs(run), YEARS, fix=frozenset(), shock=True)) for run in ('base', 'ctax')}
    rows = []
    for name in list(PM.FIXES) + ['all']:
        fx = PM.ALL_FIXES if name == 'all' else frozenset([name])
        b = kpi(PM.engine(PM.legacy_inputs('base'), YEARS, fix=fx, shock=True))
        c = kpi(PM.engine(PM.legacy_inputs('ctax'), YEARS, fix=fx, shock=True))
        pct = lambda a, z: '0' if abs(a / z - 1) < 5e-5 else f'{(a / z - 1) * 100:+.1f}%'
        rows.append([name if name != 'all' else 'All fixes', pct(b['gen30'], base['base']['gen30']),
                     pct(b['gen40'], base['base']['gen40']), pct(b['co2_40'], base['base']['co2_40']),
                     pct(c['gas40'], base['ctax']['gas40']), pct(c['vre40'], base['ctax']['vre40']),
                     pct(b['usd'], base['base']['usd'])])
    return base, rows


def blocks():
    base, eff = effects()
    title = {k: t for k, t, *_ in BUGS}
    bl = [{'t': 'title', 'text': 'CPAT power model: bugs found and improvements made'},
          {'t': 'subtitle', 'text': 'Legacy CPAT engineer power model (CPAT 1.0pre_456, Mitigation section 4) and '
                                    'CPAT-AI-Mitigation-MVP v1.06 (Egypt). 10 October 2026.'},
          {'t': 'h1', 'text': '1. Summary'},
          {'t': 'p', 'text': 'The engineer power model of legacy CPAT (power demand, dispatch, retirement and '
                             'investment) was rebuilt in CPAT-AI-Mitigation-MVP v1.06. The aim was to match legacy '
                             'except where legacy is wrong. To find out where it is wrong, every formula of the legacy '
                             'power block was read and replicated in Python.'},
          {'t': 'bullets', 'items': [
              '**10 bugs** change results or would change them in other countries. **5** change the Egypt results '
              'today: services demand, the efficiency exponent, the income factor, investment spending and (in the '
              'carbon-tax run) the investment logit. **5** are latent for Egypt (no coal, no power feebate, '
              'renewables below 75%, no investment deficit).',
              '**2 minor defects** only affect reporting or a stray empty cell.',
              '**3 corrections** of the MVP\'s own v1.05 power prices, found while matching legacy.',
              'With all fixes, Egypt\'s electricity generation in 2040 is **' + eff[-1][2] + '** higher than legacy '
              'in the baseline, and power CO2 in 2040 is **' + eff[-1][3] + '** higher. Most of it comes from the '
              'services and income-factor bugs, both of which suppressed demand growth.',
              'The MVP is verified both ways. A Python reference **reproduces 208 legacy rows** (baseline and '
              'carbon-tax run, 2022-2040) **to 1e-13** with the bugs switched on. The workbook equals the same '
              'reference with the bugs fixed.']},
          {'t': 'h1', 'text': '2. How the legacy model was checked'},
          {'t': 'bullets', 'items': [
              'The legacy workbook (.xlsb) was converted to .xlsx with LibreOffice to read its formulas. For each row, '
              'the formulas of all year columns were compared (relative references), to find rows where the formula '
              'changes from one year to the next.',
              'A Python reference (power_model_ref_v0_1.py) implements the legacy method year by year. It is fed with '
              'the cached legacy inputs: fuel and policy costs per kWh, GDP growth, the income factor, the demand '
              'shock and the prices of 2021-2024.',
              'It reproduces every cached legacy row of the engineer model and the price chain: demand by subsector, '
              'required generation, capacity, generation, shares, fuel use, retirements, both investment rounds, '
              'nameplate and USD investment, investment shares, levelised costs and electricity prices. This holds '
              'for the baseline and the carbon-tax run, which gives two independent tests.',
              'Each bug fix is a switch in the reference. The effect of each fix below is the reference with that '
              'fix on against the reference with no fix (legacy). Inputs are the legacy Egypt inputs and the legacy '
              'demand shock, so the numbers isolate the fix.']},
          {'t': 'h1', 'text': '3. Bugs'},
          {'t': 'table', 'head': ['#', 'Bug', 'Legacy rows', 'Egypt effect'],
           'widths': [400, 4600, 1900, 2700],
           'rows': [[str(i + 1), t, rows, ('changes results' if k in ('services', 'aee', 'income')
                                           else ('investment spending (reporting)' if k == 'invcost'
                                                 else ('carbon-tax run' if k == 'cmin' else 'latent')))]
                    for i, (k, t, rows, *_r) in enumerate(BUGS)]}]
    for i, (k, t, rows, does, should, why) in enumerate(BUGS):
        bl += [{'t': 'h2', 'text': f'3.{i + 1} {t}'},
               {'t': 'p', 'text': f'**Where:** {rows}.'},
               {'t': 'p', 'text': f'**Legacy:** {does}'},
               {'t': 'p', 'text': f'**Correct:** {should}'},
               {'t': 'p', 'text': f'**Why it matters:** {why}'}]
    bl += [{'t': 'h1', 'text': '4. Effect of each fix (legacy Egypt inputs)'},
           {'t': 'p', 'text': 'Change against legacy. Generation and CO2 from the baseline. Gas and variable-renewable '
                              '(VRE) shares from the carbon-tax run. Investment is the 2022-2040 total of the baseline. '
                              f'Legacy levels: generation {base["base"]["gen30"]:,.0f} TWh (2030) and '
                              f'{base["base"]["gen40"]:,.0f} TWh (2040); power CO2 {base["base"]["co2_40"]:.1f} MtCO2 '
                              f'(2040); carbon-tax run gas share {base["ctax"]["gas40"]:.1%} and VRE share '
                              f'{base["ctax"]["vre40"]:.1%} (2040); investment USD {base["base"]["usd"]:,.0f} bn.'},
           {'t': 'table', 'head': ['Fix', 'Generation 2030', 'Generation 2040', 'Power CO2 2040',
                                   'Gas share 2040 (carbon tax)', 'VRE share 2040 (carbon tax)', 'Investment USD'],
            'widths': [1500, 1300, 1300, 1300, 1500, 1500, 1300],
            'rows': [[SHORT.get(r[0], r[0]) if r[0] != 'All fixes' else '**All fixes together**'] + r[1:]
                     for r in eff]},
           {'t': 'note', 'text': 'Effects do not add up exactly, because the fixes interact (for example, more demand '
                                 'also means more investment).'},
           {'t': 'h1', 'text': '5. Minor defects'},
           {'t': 'table', 'head': ['Defect', 'Legacy rows', 'Description'], 'widths': [2200, 1600, 5800],
            'rows': [list(m) for m in MINOR]},
           {'t': 'h1', 'text': '6. Corrections of the MVP\'s own power prices (v1.05)'},
           {'t': 'bullets', 'items': MVP_FIXES},
           {'t': 'h1', 'text': '7. Design choices kept, and recommended improvements'},
           {'t': 'table', 'head': ['Topic', 'What it is and what is recommended'], 'widths': [2400, 7200],
            'rows': [list(d) for d in DESIGN]},
           {'t': 'h1', 'text': '8. Improvements made in the MVP'},
           {'t': 'bullets', 'items': IMPROVEMENTS},
           {'t': 'h1', 'text': '9. Files'},
           {'t': 'bullets', 'items': [
               'CPAT-AI-Mitigation-MVP-v1.06.xlsm: the workbook (Mitigation section 3; LegacyDiff lists every '
               'difference with legacy).',
               'power_model_ref_v0_1.py: the Python reference; FIXES lists the bugs and the legacy rows.',
               'check_report_v1.06.md: the checks, including the legacy test and the effect of each fix.',
               'PowerModel_Method_v0_1.md: the method of the power model.',
               'make_bugs_doc_v1_0.py: builds this report (numbers computed when it runs).']}]
    return bl


def main():
    data = {'title': 'CPAT power model: bugs found and improvements made',
            'footer': 'CPAT power model - bugs and improvements (v1.0)', 'blocks': blocks()}
    src = os.path.join(tempfile.mkdtemp(), 'blocks.json')
    with open(src, 'w', encoding='utf-8') as f:
        json.dump(data, f)
    subprocess.run(['node', os.path.join(HERE, 'render_docx_v1_0.js'), src, OUT], check=True)


if __name__ == '__main__':
    main()
