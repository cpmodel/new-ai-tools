"""v2.3 (version number aligned with CPAT-AI-Mitigation-MVP v2.3; content as the previous version in Old/). 

v0.2: planned retirement by type (legacy K5 offers it for coal only): where a type's switch planned_retirement is
Yes (power_tech.csv, or inp['pflag']), the planned retirement data pret|<type> (power_paths.csv, nameplate MW, or
inp['pret']) replace the rate-based retirement capacity x 1/lifetime; retirement stays capped at the capacity, and
coal adds the cost-based early retirement. Memo outputs: total capacity, total nameplate investment, demand by group,
the cheapest of coal variable cost and wind and solar investment cost (ecmin) and the coal logit share (ecshr).

Python reference of the CPAT engineer power model (legacy Mitigation section 4, rows 2954-4297): demand (H),
dispatch (I), storage (J), retirement and capacity (K), investment (L), with the price side of v1.05 (B, C, D, E, G).

Year by year, as legacy: demand uses last year's electricity price; capacity of year t is last year's capacity less
retirements plus new investment; dispatch fills demand with fixed-output types first, then coal and gas by a logit on
variable cost; investment fills the gap between demand and capacity after retirements by a logit on levelised cost,
capped by scale-up rates, with a second round that fills the rest cheapest first.

fix: the bug fixes to apply (names in FIXES). The legacy test runs with fix = frozenset() and reproduces the cached
legacy rows; the MVP runs with ALL_FIXES. Used by check_v1_07.py and make_issues_doc_v2_3.py; independent of the
workbook builder (reads data/*.csv only).
"""
import csv
import math
import os

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, 'data')
T = ['coa', 'nga', 'oop', 'nuc', 'wnd', 'sol', 'hyd', 'ore', 'bio']
GROUPS = ['tra', 'res', 'ind', 'srv']
GROUP_OF = {'rod': 'tra', 'ral': 'tra', 'avi': 'tra', 'nav': 'tra', 'res': 'res', 'srv': 'srv'}   # others: ind
PJ_PER_KTOE = 0.041868
FIXES = {  # name: (legacy rows, short description)
    'services': ('3731', 'services electricity demand grows with the industrial multiple from 2024 (copy error)'),
    'aee': ('3691, 3697, 3703, 3709', 'autonomous efficiency exponent uses the efficiency elasticity instead of the '
                                      'usage elasticity'),
    'income': ('3692, 3698, 3704, 3710', 'income factor on the GDP elasticity applied twice, anchored to 2023 '
                                         '($M3686)'),
    'lt_alloc': ('3911, 3912', 'allocation of long-term storage cost to VRE (33%) applied twice'),
    'invcost': ('4276-4284', 'investment spending uses the capital cost at factor 1 instead of the capital cost of '
                             'the year'),
    'cmin': ('4046, 4073-4081', 'investment logit normalised by the cheapest type including types not permitted'),
    'retlag': ('3973', 'early coal retirement subtracts nameplate retirements from effective capacity additions'),
    'coalret': ('4009-4011', 'effective coal retirement not capped like the nameplate retirement'),
    'rebate': ('3407-3409', 'feebate rebate per kWh consumed added to the cost per kWh generated (x 1.22)'),
    'kcheck': ('4024', 'capacity needed is computed from the planned capacity of last year, not the actual capacity '
                       '(differs when investment could not be placed)'),
}
ALL_FIXES = frozenset(FIXES)


def read(name):
    with open(os.path.join(DATA, name), encoding='utf-8') as f:
        return list(csv.DictReader(f))


def base_data(years):
    tech = {r['code']: r for r in read('power_tech.csv')}
    paths = {r['key']: [float(r[str(y)]) for y in years] for r in read('power_paths.csv')}
    par = {r['key']: float(r['value']) for r in read('power_params.csv')}
    dem0 = {r['subsector']: float(r['ktoe_2022']) for r in read('power_demand.csv')}
    return tech, paths, par, dem0


def engine(inp, years, fix=ALL_FIXES, shock=False):
    """inp: elast {group: (eps_y, eps_u, eps_f, alpha)}; gdp, inc, shock, cpf (years); rp2021_res, rp2021_ind;
    rphist_res, rphist_ind (years, used up to last_hist); vat_res, vat_ind; last_hist; vbc[i], d5[i] ($/kWh, years);
    shp_gj[i] ($/GJ fuel, years) or d6[i] ($/kWh incl. rebate, years); tvc_extra[i] (optional values after the last
    year for the D9 window, legacy 2041). Returns {name: series or [series by type]}."""
    n = len(years)
    tech, paths, par, dem0 = base_data(years)
    tf = [tech[f] for f in T]
    num = lambda i, k: float(tf[i][k])
    cf = [num(i, 'capacity_factor') for i in range(9)]
    eff = [num(i, 'efficiency') for i in range(9)]
    gj_kwh = [0.0036 / eff[i] if tf[i]['price_code'] else 0.0 for i in range(9)]
    permit_yes = [tf[i]['invest_permitted'] == 'Yes' for i in range(9)]
    x, alloc_lt = par['lt_start'], par['lt_alloc']
    kd, ki, cmin_floor = par['k_dispatch'], par['k_invest'], par['min_cost']
    shares = par['sh_nxp'] + par['sh_tsd'] + par['sh_eou'] + par['sh_add']
    o = {k: [[0.0] * n for _ in T] for k in (
        'cap', 'effcap', 'gen', 'shr', 'pe', 'co2', 'ret', 'reteff', 'pni', 'first', 'second', 'nei', 'nim', 'usd',
        'phi', 'tic', 'lfx', 'lvc', 'fix', 'gnc', 'cax', 'caxav', 'stoav', 'maxinv', 'tvc', 'shk', 'permit',
        'gmin', 'lgd', 'lgi', 'space', 'dsum')}
    pflag = inp.get('pflag', [1.0 if tf[i].get('planned_retirement') == 'Yes' else 0.0 for i in range(9)])
    pret = inp.get('pret', [paths.get(f'pret|{f}', [0.0] * n) for f in T])
    o['pret'] = [list(pret[i]) for i in range(9)]
    for k in ('captot', 'nimtot', 'ecmin', 'ecshr', 'req', 'vre', 'msc', 'cph', 'gncav', 'ccpav', 'demtot', 'early', 'need', 'reb', 'sc_res', 'sc_ind',
              'rppre_res', 'rppre_ind', 'rp_res', 'rp_ind', 'cons_res', 'cons_ind', 'reqcg', 'vcmin', 'rem', 'gtot',
              'dgen', 'egrow', 'kafter', 'cmin', 'deficit', 'usdtot', 'sgap_res', 'sgap_ind', 'rvat_res', 'rvat_ind',
              'rcarb_res', 'rcarb_ind', 'rgap_res', 'rgap_ind'):
        o[k] = [0.0] * n
    o['mult'] = {g: [0.0] * n for g in GROUPS}
    o['pdem'] = {g: [0.0] * n for g in GROUPS}
    dem = {s: [0.0] * n for s in dem0}
    ef = {r['key']: float(r['ef_tco2_per_gj']) for r in read('ef_co2.csv')}
    efp = [ef.get(f'EGY|{f}|pow', 0.0) for f in T]
    shr0 = [num(i, 'gen0_mwy') / sum(num(j, 'gen0_mwy') for j in range(9)) for i in range(9)]
    for k in range(n):                           # D: variable cost incl. shadow price (D6) and rebate, all years
        for i in range(9):
            o['shk'][i][k] = inp['d6'][i][k] if 'd6' in inp else inp['shp_gj'][i][k] * gj_kwh[i]
        if 'd6' not in inp and k > 0:            # rebate from the second year, fixed base-year shares (legacy D6)
            avg = sum(o['shk'][i][k] * shr0[i] for i in range(9))
            o['reb'][k] = -avg * (1 if 'rebate' in fix else par['gen_per_cons'])
        for i in range(9):
            o['tvc'][i][k] = inp['vbc'][i][k] + inp['d5'][i][k] + o['shk'][i][k] + o['reb'][k]
    for i in range(9):                           # D9: discounted average of variable cost to the end of the horizon
        w = num(i, 'wacc')
        path = o['tvc'][i] + list(inp['tvc_extra'][i] if 'tvc_extra' in inp else [])
        dfs = [(1 + w) ** -j for j in range(len(path))]
        for k in range(n):
            o['lvc'][i][k] = sum(a * b for a, b in zip(dfs[k:], path[k:])) / sum(dfs[k:])
            o['dsum'][i][k] = sum(dfs[k:n])
    k30 = 0.0
    for k, y in enumerate(years):
        # ---------------- H demand
        if k == 0:
            for s in dem:
                dem[s][0] = dem0[s]
            w_res = {'tra': 0.5, 'res': 1.0, 'ind': 0.0, 'srv': 0.0}
            for g in GROUPS:
                o['pdem'][g][0] = w_res[g] * inp['rp2021_res'] + (1 - w_res[g]) * inp['rp2021_ind']
                o['mult'][g][0] = 1.0
        else:
            pr = lambda g, kk: (inp['rp2021_' + g] if kk < 0 else o['rp_' + g][kk])   # price of year index kk
            p = {g: (pr(g, k - 1), pr(g, k - 2)) for g in ('res', 'ind')}
            p['tra'] = ((p['res'][0] + p['ind'][0]) / 2, (p['res'][1] + p['ind'][1]) / 2)
            p['srv'] = p['ind']
            for g in GROUPS:
                o['pdem'][g][k] = p[g][0]
                ey, eu, epf, a = inp['elast'][g]
                aee = 1 / (1 + a + par['extra_eff']) ** (1 + (eu if 'aee' in fix else epf)) - 1
                inc = inp['inc'][k] if 'income' in fix else inp['inc'][k] ** 2 / inp['inc'][1]
                gd = (1 + inp['gdp'][k]) ** (ey * inc) - 1
                c0, c1 = inp['cpf'][k], inp['cpf'][k - 1]
                pe_ = (p[g][0] / p[g][1]) ** eu * ((p[g][0] + c0) / (p[g][1] + c1)) ** (epf + eu * epf) - 1
                o['mult'][g][k] = (1 + aee) * (1 + gd) * (1 + pe_) * (1 + (inp['shock'][k] if shock else 0.0))
            for s in dem:
                g = GROUP_OF.get(s, 'ind')
                if s == 'srv' and y >= 2024 and 'services' not in fix:
                    g = 'ind'
                dem[s][k] = dem[s][k - 1] * o['mult'][g][k]
        tot = sum(dem[s][k] for s in dem)
        o['demtot'][k] = tot
        o['cons_res'][k] = dem['res'][k] * par['ktoe_gwh']
        o['cons_ind'][k] = (tot - dem['res'][k]) * par['ktoe_gwh']
        o['req'][k] = sum(num(i, 'gen0_mwy') for i in range(9)) if k == 0 else \
            tot * (1 + shares) * par['ktoe_gwh'] * 1000 / 8760
        # ---------------- K1 capacity
        for i in range(9):
            o['cap'][i][k] = num(i, 'cap0_mw') if k == 0 else \
                o['cap'][i][k - 1] - o['ret'][i][k - 1] + o['nim'][i][k - 1]
            o['effcap'][i][k] = o['cap'][i][k] * cf[i]
            o['permit'][i][k] = 1.0 if permit_yes[i] and y >= num(i, 'invest_start') else 0.0
        toteff = sum(o['effcap'][i][k] for i in range(9))
        o['captot'][k] = sum(o['cap'][i][k] for i in range(9))
        # ---------------- I dispatch
        if True:                                 # dispatch (base year: memo rows only, generation is data)
            cg = [i for i in range(9) if tf[i]['dispatch'] == 'cg']
            for i in range(9):
                if i not in cg:
                    o['gen'][i][k] = o['effcap'][i][k]
            reqcg = max(o['req'][k] - sum(o['gen'][i][k] for i in range(9) if i not in cg), 0.0)
            o['reqcg'][k] = reqcg
            gmax = {i: o['cap'][i][k] * par['maxcf_cg'] for i in cg}
            gmin = {i: min(max(reqcg - sum(gmax[j] for j in cg if j != i), 0.0), gmax[i]) for i in cg}
            rem = max(reqcg - sum(gmin.values()), 0.0)
            vc = {i: max(o['tvc'][i][k], cmin_floor) for i in cg}
            m = min(vc.values())
            lg = {i: math.exp(-kd * vc[i] / m) for i in cg}
            o['rem'][k], o['vcmin'][k] = rem, m
            for i in cg:
                o['gmin'][i][k], o['lgd'][i][k] = gmin[i], lg[i]
            for i in cg:
                o['gen'][i][k] = min(gmax[i], gmin[i] + rem * lg[i] / sum(lg.values()))
            if k == 0:
                for i in range(9):
                    o['gen'][i][0] = num(i, 'gen0_mwy')
        gtot = sum(o['gen'][i][k] for i in range(9))
        o['gtot'][k] = gtot
        o['dgen'][k] = max(0.0, gtot - o['gtot'][k - 1]) if k else 0.0
        for i in range(9):
            o['shr'][i][k] = o['gen'][i][k] / gtot if gtot else 0.0
            o['pe'][i][k] = o['gen'][i][k] * 8.76 / par['ktoe_gwh'] / eff[i]
            o['co2'][i][k] = o['pe'][i][k] * PJ_PER_KTOE * efp[i]
        v = sum(o['shr'][i][k] for i in range(9) if int(tf[i]['vre_share']))
        o['vre'][k] = v
        # ---------------- J storage
        cph = (paths['cbat'][k] + paths['cint'][k] / par['st_ratio']) / (par['dlf_ren'] * 8760) + paths['obat'][k] / 8760
        m_lt = 2 * (v - x) / (1 - x) ** 2 if v > x else 0.0
        lt = m_lt * paths['cel'][k] * alloc_lt * (1 if 'lt_alloc' in fix else alloc_lt)
        o['cph'][k], o['msc'][k] = cph, par['st_alloc'] * par['st_marg_hours'] * v * cph + lt
        # ---------------- B, C, E (v1.05) and D9, E1
        for i in range(9):
            N, w, cfi = num(i, 'life_years'), num(i, 'wacc'), cf[i]
            dlf = (1 - (1 + w) ** -N) / (1 - (1 + w) ** -1)
            vflag = int(tf[i]['vre'])
            o['cax'][i][k] = num(i, 'cax0_usd_per_kw') * paths[f'tcf|{T[i]}'][k]
            if k == 0:
                o['caxav'][i][0], o['stoav'][i][0] = o['cax'][i][0], vflag * o['msc'][0]
            else:
                ph = o['phi'][i][k - 1]
                o['caxav'][i][k] = o['caxav'][i][k - 1] * (1 - ph) + o['cax'][i][k - 1] * ph
                o['stoav'][i][k] = o['stoav'][i][k - 1] * (1 - ph) + vflag * o['msc'][k] * ph
            dtc, tcx, opf = num(i, 'dtc_usd_per_kw'), num(i, 'tcx_usd_per_kw'), num(i, 'opf_usd_per_kwh')
            o['fix'][i][k] = (o['caxav'][i][k] * (1 / N + w / 2) + dtc / N) / (cfi * 8760) + o['stoav'][i][k] + opf
            o['lfx'][i][k] = (o['cax'][i][k] + dtc + tcx) / (cfi * 8760 * dlf) + opf + num(i, 'rns_usd_per_kwh') \
                + vflag * o['msc'][k]
            o['gnc'][i][k] = o['fix'][i][k] + inp['vbc'][i][k]
        o['gncav'][k] = sum(o['shr'][i][k] * o['gnc'][i][k] for i in range(9))
        o['ccpav'][k] = sum(o['shr'][i][k] * inp['d5'][i][k] for i in range(9)) * par['gen_per_cons']
        # ---------------- G prices (supply cost lagged one year; VAT on the policy term, as legacy)
        for g in ('res', 'ind'):
            sc = (o['gncav'][k - 1] if k else o['gncav'][0]) + par[f'tmc|{g}']
            vat = inp['vat_' + g]
            pre = inp['rphist_' + g][k] if y <= inp['last_hist'] else \
                o['rppre_' + g][k - 1] + par['pass'] * (sc - o['sc_' + g][k - 1]) * (1 + vat)
            o['sc_' + g][k], o['rppre_' + g][k] = sc, pre
            o['rp_' + g][k] = pre + o['ccpav'][k] * par['pass'] * (1 + vat)
            o['sgap_' + g][k] = sc - pre / (1 + vat)
            cons = o['cons_' + g][k]
            o['rvat_' + g][k] = pre * vat / (1 + vat) * cons
            o['rcarb_' + g][k] = o['ccpav'][k] * cons
            o['rgap_' + g][k] = o['sgap_' + g][k] * cons
        for i in range(9):
            o['tic'][i][k] = max(o['lfx'][i][k] + o['lvc'][i][k], cmin_floor)
        # ---------------- K4 cost-based early coal retirement, K5/K6 retirements
        maxinv = [num(i, 'scaleup_rate') * toteff * o['permit'][i][k] for i in range(9)]
        incdem = max(0.0, sum(o['gen'][i][k - 1] for i in range(9)) - sum(o['gen'][i][k - 2] for i in range(9))) \
            if k >= 2 else 0.0
        retlag = sum((o['reteff'] if 'retlag' in fix else o['ret'])[i][k - 1] for i in range(1, 9)) if k >= 1 else 0.0
        growth = max(0.0, maxinv[4] + maxinv[5] - incdem - retlag)
        o['egrow'][k] = growth
        mx = min(o['effcap'][0][k] * par['early_max'], growth)
        cc = [o['tvc'][0][k], o['tic'][4][k], o['tic'][5][k]]                # legacy K4: coal cost not floored
        e = [math.exp(-ki * c / min(cc)) for c in cc]
        o['ecmin'][k], o['ecshr'][k] = min(cc), e[0] / sum(e)
        early_np = max(0.0, mx * (1 - e[0] / sum(e))) / cf[0] * par['early_factor'] if cf[0] else 0.0
        o['early'][k] = early_np
        for i in range(9):
            planned = pret[i][k] if pflag[i] else o['cap'][i][k] * num(i, 'retire_rate')
            if i == 0:
                o['ret'][0][k] = min(planned + early_np, o['cap'][0][k])
                o['reteff'][0][k] = (o['ret'][0][k] if 'coalret' in fix else planned + early_np) * cf[0]
            else:
                o['ret'][i][k] = min(planned, o['cap'][i][k])
                o['reteff'][i][k] = o['ret'][i][k] * cf[i]
        # ---------------- K7 effective capacity needed (to meet this year's demand next year, as legacy)
        chk = toteff if (k == 0 or 'kcheck' in fix) else k30
        after = chk - sum(o['reteff'][i][k] for i in range(9))
        need = max(o['req'][k] - after, 0.0)
        o['need'][k], k30 = need, after + need
        o['kafter'][k] = after
        # ---------------- L investment: logit on levelised cost, capped by scale-up rates, then cheapest first
        tic = [o['tic'][i][k] for i in range(9)]
        cmin = min([tic[i] for i in range(9) if o['permit'][i][k]] or [min(tic)]) if 'cmin' in fix else min(tic)
        lg = [o['permit'][i][k] * math.exp(-ki * tic[i] / cmin) for i in range(9)]
        pni = [u / sum(lg) for u in lg]
        o['cmin'][k] = cmin
        raw = [need * q for q in pni]
        first = [min(maxinv[i], raw[i]) for i in range(9)]
        deficit = sum(raw) - sum(first)
        o['deficit'][k] = deficit
        space = [(maxinv[i] - first[i]) * o['permit'][i][k] for i in range(9)]
        order = sorted(range(9), key=lambda i: tic[i] + i * 1.356e-6)     # legacy tie breaker
        second, cum_prev, run_ = [0.0] * 9, 0.0, 0.0
        for i in order:
            run_ += space[i]
            cum = min(deficit, run_)
            second[i] = o['permit'][i][k] * (cum - cum_prev)
            cum_prev = cum
        for i in range(9):
            o['pni'][i][k], o['first'][i][k], o['second'][i][k], o['maxinv'][i][k] = pni[i], first[i], second[i], maxinv[i]
            o['nei'][i][k] = first[i] + second[i]
            o['lgi'][i][k], o['space'][i][k] = lg[i], space[i]
            o['nim'][i][k] = o['nei'][i][k] / cf[i]
            capex = o['cax'][i][k] if 'invcost' in fix else num(i, 'cax0_usd_per_kw')
            o['usd'][i][k] = capex * o['nim'][i][k] / 1000
            den = o['cap'][i][k] + o['nim'][i][k] - o['ret'][i][k]
            o['phi'][i][k] = 0.0 if den == 0 else o['nim'][i][k] / den
        o['usdtot'][k] = sum(o['usd'][i][k] for i in range(9))
        o['nimtot'][k] = sum(o['nim'][i][k] for i in range(9))
    o['dem'] = dem
    o['demg'] = {g: [sum(dem[s][k] for s in dem if GROUP_OF.get(s, 'ind') == g) for k in range(n)] for g in GROUPS}
    return o


def elasticities(income_group='LMIC'):
    el = {r['key']: float(r[income_group]) for r in read('elasticities.csv')}
    return {g: (el[f'inc|ecy|{g}'], el[f'usg|ecy|{g}'], el[f'eff|ecy|{g}'], el[f'aei|ecy|{g}']) for g in GROUPS}


def legacy_inputs(run):
    """Engine inputs from the cached legacy run ('base' or 'ctax'): exogenous paths, variable costs D4-D6 incl. the
    2041 value for the D9 window, historical prices before policy and VAT."""
    years = list(range(2022, 2041))
    ref = {r['key']: r for r in read('legacy_engineer_reference.csv') if r['run'] == run}
    val = lambda key, y: float(ref[key][str(y)])
    ser = lambda key: [val(key, y) for y in years]
    return {
        'elast': elasticities(), 'gdp': ser('gdp'), 'inc': ser('inc'), 'shock': ser('shock'), 'cpf': ser('cpf'),
        'rp2021_res': val('rp_res', 2021), 'rp2021_ind': val('rp_ind', 2021),
        'rphist_res': ser('rppre_res'), 'rphist_ind': ser('rppre_ind'),
        'vat_res': val('vat_res', 2024), 'vat_ind': val('vat_ind', 2024), 'last_hist': 2024,
        'vbc': [ser(f'vbc|{f}') for f in T], 'd5': [ser(f'd5|{f}') for f in T], 'd6': [ser(f'd6|{f}') for f in T],
        'tvc_extra': [[val(f'tvc|{f}', 2041)] for f in T]}


def legacy_test(run, fix=frozenset(), shock=True):
    """Engine on the legacy inputs vs the cached legacy rows: [(max relative diff, key, year, legacy, engine)]."""
    years = list(range(2022, 2041))
    o = engine(legacy_inputs(run), years, fix=fix, shock=shock)
    out = []
    for r in read('legacy_engineer_reference.csv'):
        if r['run'] != run:
            continue
        key = r['key']
        name, f = (key.split('|') + [''])[:2]
        if name in ('gdp', 'inc', 'shock', 'cpf', 'vbc', 'd5', 'd6', 'cfp', 'df', 'vat_res', 'vat_ind'):
            continue                                       # inputs
        if name == 'dem':
            got = o['dem'][f]
        elif f:
            got = o[name][T.index(f)]
        else:
            got = o[name]
        leg = [float(r[str(y)]) for y in years]
        d, yr = max((abs(a - b) / max(1.0, abs(a)), y) for a, b, y in zip(leg, got, years))
        out.append((d, key, yr))
    return sorted(out, reverse=True)
