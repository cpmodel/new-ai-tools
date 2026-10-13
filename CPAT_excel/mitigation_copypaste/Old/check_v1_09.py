"""Checks for CPAT-AI-Mitigation-MVP-v1.08.xlsm (v1.08: the workbook works without LAMBDA support: no cell calls a
LAMBDA, the shipped file recalculates in LibreOffice without a single error, every value equals the independent Python
recomputation; the optional LAMBDA library (sheet Lambdas) is tested by copying each row's LAMBDA formula over the
whole row) (v1.07: every cell formula at most 100 characters (section 0);
LAMBDAs also outside the right column (ROWLABEL in column H), so every variant expands all LAMBDA calls; planned
retirement by type vs power_model_ref_v0_2 (test with data switched on); electricity demand memo lines; power
indicators on MTOutputs and Charts; regression vs v1.06 on every shared code with no intended change)
(v1.06: power engineer model, section 3 H-L: power_model_ref_v0_1 = cached legacy
rows with the bugs switched on; workbook = Python with all bug fixes; power CO2 in section 13; regression vs v1.05 apart
from the intended power and CO2 changes) (v1.05: power generation costs and electricity prices, section 3: Python
reference = cached legacy rows with the legacy inputs, workbook = Python with the MVP inputs, uniformity, outline)
(v1.04: first-run test macro CheckBatchRun and sheet MacroCheck,
run in LibreOffice, including a negative test) (v1.03: scenario definitions, VBA batch run, StoredResults,
ScenarioCompare; the macro itself is run in LibreOffice and compared with the Python emulation)
(v1.02: new ETS with cap, benchmarks, volatility, goal seek)
(v1.01: MTOutputs and Charts sheets) (v0.12 revenues; v0.13 CO2; v0.14 corrected Egypt prices;
v0.15 layout: row group buttons below, rolled up, Mitigation first at 75%; v0.16: one level of column groups incl.
2030-2039, light beige for the LAMBDA column and the price history block, blank column between scenario groups). Needs LibreOffice (soffice) to recalculate.

LibreOffice 24.2 cannot evaluate LAMBDA, so the 2035 LAMBDA cells are checked by expansion (body with the
arguments substituted) and by dragging the plain 2034 formula forward.

1. Shipped file: errors only in the 2035 columns.
2. Uniformity (relative R1C1): one formula per variable for base year / plain years / 2035, across all
   subsectors, fuels and scenarios; one formula per D:G column per variable; Section-1 input rows; paths;
   label and code columns.
3. LAMBDA encoding.
4. Values vs an independent Python recomputation from the CSVs (both 2035 variants): international prices,
   supply cost, taxes, retail price before new policies, the subsector variables and fuel use.
5. Scenario tests (MTInputs column + Mitigation group copied as a user would): copy = source; zero carbon price =
   baseline; legacy Egypt carbon-tax settings = legacy row 8582; fuel price reform, feebates, price source
   IMF-WB* with High adjustment and a nominal carbon price vs Python; feebates lower fuel use only where covered;
   global price controls None and Manual vs Python.
6. Regression vs v0.9 on every output code the two versions share, except the intended price changes
   (other oil products like the other oil products; VAT-rate assumption).
7. Labels and codes; 8. outline; 9. MTInputs = template; 10. format; 11. sanity and results.

    python check_v1_08.py   ->  writes check_report_v1.08.md
"""
import csv
import importlib.util
import math
import os
import re
import subprocess
import tempfile
import types

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI

import build_v1_09 as B
import power_model_ref_v0_2 as PM

HERE = os.path.dirname(os.path.abspath(__file__))
REF = re.compile(r"(?<![A-Za-z_\d.\"])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])")
TOL = 1e-9
VAT_CONSUMER = {'coa.res', 'nga.res', 'gso.all', 'die.all', 'lpg.all', 'ker.all', 'oop.all'}   # independent copy
report = []


def log(line=''):
    print(line)
    report.append(line)


def recalc_path(path):
    out = tempfile.mkdtemp()
    subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', out, path],
                   check=True, capture_output=True, timeout=900)
    return os.path.join(out, os.path.splitext(os.path.basename(path))[0] + '.xlsx')


def recalc_wb(path):
    return openpyxl.load_workbook(recalc_path(path), data_only=True)


def add_mto_blocks(wb, groups):
    """Add MTOutputs blocks as a user would: copy the last block (header to blank row) directly below."""
    ws = wb['MTOutputs']
    for g in groups:
        s0, d0 = B.mto_row(g - 1, 0) - 1, B.mto_row(g, 0) - 1
        for k in range(B.MTO_BLOCK):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(s0 + k, c).value
                if isinstance(v, str) and v.startswith('='):
                    v = Translator(v, origin=f'{L(c)}{s0 + k}').translate_formula(f'{L(c)}{d0 + k}')
                ws.cell(d0 + k, c).value = v
    return wb


def check_mto(wbv, groups, label):
    """Every MTOutputs indicator cell equals the Mitigation cell of its output code (searched in Python)."""
    wm, wo = wbv['Mitigation'], wbv['MTOutputs']
    found, worst, n, missing = {}, 0.0, 0, []
    for g in groups:
        code, base, years = B.group_cols(g)
        for r in range(1, wm.max_row + 1):
            k = wm.cell(r, code).value
            if isinstance(k, str):
                found[k] = [wm.cell(r, c).value for c in [base] + years]
    for g in groups:
        num = wo.cell(B.mto_row(g, 0) - 1, 5).value
        for k, (stem, _f) in enumerate(B.MTO_INDICATORS):
            r = B.mto_row(g, k)
            key = wo.cell(r, 1).value
            if key not in found or num != g:
                missing.append(key)
                continue
            for j in range(B.NY):
                a, b = wo.cell(r, B.MTO_C0 + j).value, found[key][j]
                worst = max(worst, abs(a - b) if isinstance(a, (int, float)) and isinstance(b, (int, float))
                            else float('inf'))
                n += 1
    ok = not missing and worst < TOL
    log(f'- MTOutputs {label}: {len(groups)} scenario blocks x {len(B.MTO_INDICATORS)} indicators, {n} cells equal '
        f'the Mitigation cell of their output code (max abs diff {worst:.3g}); block numbers 1..{groups[-1]} '
        f'{"OK" if ok else "FAIL " + str(missing[:5])}')
    return ok


def check_charts():
    wb = openpyxl.load_workbook(B.OUT)
    charts = wb['Charts']._charts
    stems = [s for s, _ in B.MTO_INDICATORS]
    good = len(charts) == len(B.CHARTS)
    for ch, (ttl, stem, _y) in zip(charts, B.CHARTS):
        refs = [s.val.numRef.f.replace("'", '') for s in ch.series]
        r = lambda g: B.mto_row(g, stems.index(stem))
        exp = [f"MTOutputs!${L(B.MTO_C0)}${r(g)}:${L(B.MTO_C0 + B.NY - 1)}${r(g)}" for g, *_ in B.SCENARIOS]
        good &= refs == exp
    log(f'- Charts: {len(charts)} line charts ({", ".join(c[0] for c in B.CHARTS)}), one series per scenario block '
        f'on the right MTOutputs rows {"OK" if good else "FAIL"}')
    return good


def recalc(path):
    return openpyxl.load_workbook(recalc_path(path), data_only=True)['Mitigation']


def r1c1(formula, row, col):
    def sub(m):
        c, rw = CI(m.group(2)), int(m.group(4))
        return (f'R{rw}' if m.group(3) else f'R[{rw - row}]') + (f'C{c}' if m.group(1) else f'C[{c - col}]')
    return REF.sub(sub, str(formula))


ERR_VALUES = ('#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', 'Err:')


def is_err(v):
    return isinstance(v, str) and v.startswith(ERR_VALUES)


def split_args(text):
    """Split a function's argument text at top-level commas (INDEX(...) arguments contain commas)."""
    args, depth, cur = [], 0, ''
    for ch in text:
        if ch == ',' and depth == 0:
            args.append(cur)
            cur = ''
            continue
        depth += (ch == '(') - (ch == ')')
        cur += ch
    return args + [cur]


def expand(formula, lambdas=None):
    """Replace every call of a named LAMBDA (anywhere in the formula) by its body with the arguments substituted."""
    lambdas = lambdas or B.LAMBDAS
    pat = re.compile(r'(?<![A-Za-z_.])(' + '|'.join(lambdas) + r')\(')
    while True:
        m = pat.search(formula)
        if not m:
            return formula
        depth, k = 1, m.end()
        while depth:
            depth += (formula[k] == '(') - (formula[k] == ')')
            k += 1
        params, body = lambdas[m.group(1)]
        args = split_args(formula[m.end():k - 1])
        assert len(args) == len(params), formula
        for prm in sorted(params, key=len, reverse=True):
            body = re.sub(rf'(?<![A-Za-z_.]){prm}(?![A-Za-z_\d(])', f'({args[params.index(prm)]})', body)
        formula = formula[:m.start()] + '(' + body + ')' + formula[k:]


def rows_of(var):
    """(row, sector-or-subsector, fuel) of a variable: section 2 price rows or sector-section rows."""
    if var in B.PVARS:
        return [(B.R_PV0 + B.PVOFF[var] + i, pc.split('.')[1], pc.split('.')[0]) for i, (pc, _l) in enumerate(B.PRICES)]
    if var in B.VARS:                      # subsector variables first (power CO2 shares the name co2)
        return B.var_rows(var)
    return [(r, c_, b_) for r, v, b_, c_, _k in B.PW_ROWS if v == var]


def groups_in(ws):
    return [g for g in range(1, 40) if ws.cell(B.R_SCEN, B.group_cols(g)[0]).value not in (None, '')]


LAMBDA_CALL = re.compile(r'(?<![A-Za-z_.])(' + '|'.join(B.LAMBDAS) + r')\(')


def lambda_rows(ws, c):
    """Rows whose cell in column c calls a named LAMBDA."""
    return [r for r in range(1, ws.max_row + 1)
            if isinstance(ws.cell(r, c).value, str) and LAMBDA_CALL.search(ws.cell(r, c).value)]


def variant(wb, how):
    """how = 'expand': LAMBDA bodies substituted in the right column; 'drag': plain formula of the previous year
    dragged into the right column; 'everywhere': the right-column LAMBDA formula copied back over the whole row
    (base year included, except fuel use whose base year is data), then expanded."""
    ws = wb['Mitigation']
    if how == 'everywhere':                # v1.08: the optional LAMBDA library (sheet Lambdas) over whole rows
        lam = wb['Lambdas']
        src_c = B.group_cols(1)[2][-1]
        ener = {r for r, *_ in B.var_rows('ener')} | {r for r, v, *_ in B.PW_ROWS if v in B.PW_BASE}
        for i in range(5, lam.max_row + 1):
            r, f = lam.cell(i, 1).value, lam.cell(i, 6).value
            if not r:
                continue
            for g in groups_in(ws):
                _, base, years = B.group_cols(g)
                for k in ([base] if r not in ener else []) + years:
                    ws.cell(r, k).value = expand(Translator('=' + f, origin=f'{L(src_c)}{r}')
                                                 .translate_formula(f'{L(k)}{r}'))
    for wsx in wb.worksheets:              # v1.07: LAMBDA calls elsewhere (ROWLABEL in column H)
        for row in wsx.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith('=') and LAMBDA_CALL.search(cell.value):
                    cell.value = expand(cell.value)
    path = os.path.join(tempfile.mkdtemp(), f'{how}.xlsx')
    wb.save(path)
    return path


# ------------------------------------------------------------------ independent recomputation
def read(name):
    with open(os.path.join(HERE, 'data', name), encoding='utf-8') as f:
        return list(csv.DictReader(f))


def mt_used():
    tpl = openpyxl.load_workbook(B.MT_TEMPLATE)['MTInputs']
    return {r: tpl.cell(r, 6).value for r in range(8, 416)}


def series(row, years):
    return {y: float(row[str(y)]) for y in years if row.get(str(y)) not in ('', None)}


ETS_VOL = {'Medium': (0.42, 0.5), 'High': (0.63, 0.75), 'Low': (0.28, 1 / 3), 'Zero': (0.0, 0.0)}   # legacy
_BASE = {}


def baseline_co2():
    """Baseline CO2 by ETS sector (scenario 1, template price controls), from the Python recomputation."""
    if not _BASE:
        r = recompute(scenario_mt(1), base={})
        _BASE.update({s: [sum(r[('co2', s, f)][k] for f in [x[0] for x in B.FUELS] if ('co2', s, f) in r)
                          for k in range(len(B.YEARS) + 1)] for s in B.COV_SECTORS})
    return _BASE


ECP_SEC = lambda s: 'pow' if s == 'pow' else {'tra': 'trs', 'bld': 'res'}.get(  # independent copy (legacy ECP sectors)
    next(x[2] for x in B.SUBSECTORS if x[0] == s), 'ind')


def recompute(mt, controls=None, bench=None, ovr=None, base=None, shock_on=0, power=None, ecp=None):
    """mt: MTInputs row -> value for one scenario; controls: global GovPriceControls (default: template); bench:
    {group: (start, target)} ETS benchmarks (default B.ETS_GROUPS); ovr: {year: price} override row (default blank);
    base: baseline CO2 by ETS sector (default: scenario 1 recomputed); ecp: {code: {year: nominal value}} existing
    carbon pricing data of the country (default: data/ecp.csv rows of EGY, none).
    v1.09: existing carbon pricing, power fuel use in the fuel totals, power revenues in the revenue totals.
    Returns {(var, sector-or-subsector, fuel): [values base..last]} for gp, sp, txo, rpb and the subsector
    variables, all from the CSVs."""
    tpl = mt_used()
    ry, controls = tpl[B.MT_NAMES['ResultsYear']], controls or tpl[B.MT_NAMES['GovPriceControls']]
    units = {r['fuel']: float(r['gj_per_price_unit']) for r in read('fuel_units.csv')}
    dom = {r['country_year']: r for r in read('prices_dom.csv')}
    ef = {r['key']: float(r['ef_tco2_per_gj']) for r in read('ef_co2.csv')}
    el = {r['key']: float(r['LMIC']) for r in read('elasticities.csv')}
    use0 = {(r['subsector'], r['fuel']): float(r['ktoe']) for r in read('energy_use.csv')}
    weo = read('weo.csv')[0]
    yrs_all = list(range(2021, 2046))
    us = {r['key']: series(r, yrs_all) for r in read('weo_us.csv')}
    cpi = {y: us['USA|pcpi'][ry] / us['USA|pcpi'][y] for y in us['USA|pcpi']}
    dfl = {y: us['USA|ngdp_d'][ry] / us['USA|ngdp_d'][y] for y in us['USA|ngdp_d']}
    pint = {r['key']: series(r, yrs_all) for r in read('prices_int.csv')}
    gas = {'Global': 'glo', 'LNG': 'lng', 'Europe': 'eur', 'North Am': 'nam'}[
        next(r['gas_market'] for r in read('price_assumptions.csv') if r['country_name'] == 'Egypt')]
    sub = {s[0]: s for s in B.SUBSECTORS}
    fue = {f[0]: f for f in B.FUELS}
    years = [B.BASE_YEAR] + B.YEARS
    last_hist = B.HIST_YEARS[-1]
    n = B.MT_NAMES

    def val(col, y):
        v = dom[f'EGY{y}'].get(col, '')
        return float(v) if v not in ('', None) else 0.0

    def path(y, y0, y1, s0, t, linear=True):
        if y < y0:
            return 0.0
        k = y - y0 if (y <= y1 or linear) else y1 - y0
        return s0 + (t - s0) / max(y1 - y0, 1) * k
    nominal = str(mt[n['NomorReal']]).startswith('Nominal')
    cp = {y: path(y, mt[n['CPIntro']], mt[n['CPOutro']], mt[n['CPLevelStart']], mt[n['CPLevelTarget']],
                  mt[n['ExtendCarbonPriceBeyondOutro']] == 'Linear*') * (cpi[y] if nominal else 1) for y in years}
    fpr = {pc: {y: path(y, mt[140], mt[141], mt[r0] or 0, mt[r1] or 0) for y in years}
           for pc, _l, r0, r1 in B.PRICE_FUELS}
    fb = {s: {y: path(y, mt[n['D_FeebateIntro']], mt[n['D_FeebateOutro']], mt[n[a]], mt[n[b]]) for y in years}
          for s, a, b in B.FEEBATE_SECTORS}
    y0, y1 = mt[n['D_ETSIntro']], mt[n['D_ETSOutro']]
    ets_on = {y: str(mt[n['D_NewETS']]).startswith('Yes') and y >= y0 for y in years}
    ets_cov = {s: {y: 1.0 if ets_on[y] and bool(mt[B.MT_ROWS_FIXED['etscov0'] + k]) else 0.0 for y in years}
               for k, s in enumerate(B.COV_SECTORS)}
    bench = bench or {g: (b0, b1) for g, b0, b1 in B.ETS_GROUPS}
    bm = {g: {y: 0.0 if y < y0 else b0 + (b1 - b0) / max(y1 - y0, 1) * (min(y, y1) - y0) for y in years}
          for g, (b0, b1) in bench.items()}
    obr = {s: {y: min(1.0, max(0.0, bm[B.ETS_GROUP_OF[s]][y])) for y in years} for s in B.COV_SECTORS}
    vol = ETS_VOL[str(mt[n['D_ETSVolatility']]).replace('*', '')][0]
    imp = ETS_VOL[str(mt[n['D_ETSVolImpact']]).replace('*', '')][1]
    vadj = (1 + mt[n['D_ETSCTRisk']] * imp) / (1 + vol * imp)
    base = baseline_co2() if base is None else base
    bco2 = {s: base.get(s, [0.0] * len(years)) for s in B.COV_SECTORS}
    yi = {y: k for k, y in enumerate(years)}
    bce = {y: sum(bco2[s][yi[y]] * ets_cov[s][y] for s in B.COV_SECTORS) for y in years}
    rs, rt = mt[n['D_ETSChangeRelStart']], mt[n['D_ETSChangeRelTarget']]
    cap = {}
    for y in years:
        if not ets_on[y]:
            cap[y] = 0.0
        elif y > y1 and str(mt[n['D_ETSCapCont']]).startswith('Constant'):
            cap[y] = cap[y - 1]
        else:
            cap[y] = bce[y] * (1 + rs + (rt - rs) / max(y1 - y0, 1) * (min(y, y1) - y0))
    semi = {s: B.ETS_SEMI_OF[s] for s in B.COV_SECTORS}      # legacy sector table (also tested in Variables)
    se = {y: (sum(bco2[s][yi[y]] * ets_cov[s][y] * semi[s] * (1 - 0.5 * obr[s][y]) for s in B.COV_SECTORS) / bce[y]
              if bce[y] > 0 else 0.0) for y in years}
    est = {y: math.log(cap[y] / bce[y]) / (se[y] * vadj) if ets_on[y] and 0 < cap[y] < bce[y] and se[y] < 0 else 0.0
           for y in years}
    override = str(mt[n['D_ETSPriceOverride']]).startswith('Yes')
    ets_p = {y: ((cp[y] if (ovr or {}).get(y) in (None, '') else ovr[y]) if override else est[y]) if ets_on[y]
             else 0.0 for y in years}
    ets_pe = {y: ets_p[y] * vadj for y in years}
    # existing carbon pricing (v1.09): real prices (data x CPI index), coverage of the last data year
    if ecp is None:
        ecp = {r['code']: {int(y): r[str(y)] for y in B.ECP_YEARS} for r in read('ecp.csv') if r['country'] == 'EGY'}
    ecp_last = max(B.ECP_YEARS)
    ecpv = lambda code, y: float(ecp.get(code, {}).get(y) or 0.0)
    cov = lambda code: ecpv(code, ecp_last)
    xp = {}
    for sch, code, sw, gr in (('xct', 'ecr.all', 'ExistingCTApply', 'ExistingCTGrowth'),
                              ('xets', 'eer.all', 'ExistingETSApply', 'ExistingETSGrowth')):
        dat = {y: ecpv(code, min(y, ecp_last)) * cpi[min(y, ecp_last)] for y in years}
        on, g_ = str(mt[n[sw]]).startswith('Yes'), float(mt[n[gr]] or 0)
        xp[sch] = {y: dat[y] * (1 + g_) ** max(0, y - ecp_last) if on else 0.0 for y in years}
        xp[sch + '0'] = {y: ecpv(code, min(y, ecp_last, last_hist)) * cpi[min(y, ecp_last, last_hist)] for y in years}
    comp = str(mt[n['CTaxComplimentaryToETS']]).startswith('Yes')
    sc_used = lambda s, sw: (1 - cov(f'ees.{ECP_SEC(s)}')) if comp else float(bool(sw))

    res = {}
    res[('cptraj', '', '')] = [cp[y] for y in years]
    for k_, d_ in (('ets.p', ets_p), ('ets.pe', ets_pe), ('ets.cap', cap), ('ets.est', est), ('ets.bce', bce),
                   ('ets.se', se), ('xct.p', xp['xct']), ('xct.p0', xp['xct0']), ('xets.p', xp['xets']),
                   ('xets.p0', xp['xets0'])):
        res[(k_, '', '')] = [d_[y] for y in years]
    for s in B.COV_SECTORS:
        res[('bco2', s, '')] = bco2[s]
        res[('obr', s, '')] = [obr[s][y] for y in years]
    # international prices
    src = str(mt[n['IntEnerPricForeSource']]).replace('*', '')
    adj = str(mt[n['IntEnerPricForecastAdjustment']])
    gp = {}
    for fu, _lab, _key, hi, lo in B.GP_ROWS:
        comm = {'oil': 'oil', 'coa': 'coa', 'nga': f'nga.{gas}'}.get(fu)
        fac = hi if adj.startswith('High') else (lo if adj.startswith('Low') else 1)
        gp[fu] = [1.0 if comm is None else pint[f'{src}|{comm}'][y] * dfl[y] * (fac if y > last_hist else 1)
                  for y in years]
        res[('gp', 'int', fu)] = gp[fu]
    # domestic prices
    rpb, vrate = {}, {}
    for pc, _lab in B.PRICES:
        fuel, sec = pc.split('.')
        gj = units[fuel]
        pccr = 1.0 if fuel == 'bio' else val(f'mit.ps.{pc}', B.BASE_YEAR)
        pccb = 0.0 if pccr <= 0.25 else (0.5 if pccr <= 0.5 else 1.0)
        pcc = 0.8 if controls.startswith('Manual') else (1.0 if controls.startswith('None') else pccb)
        mar = 0.0 if fuel == 'bio' else val(f'mit.mar.{pc}', B.BASE_YEAR)
        pcost = val(f'mit.{fuel}.prod.cost', B.BASE_YEAR) if fuel in ('coa', 'nga') else 0.0
        fixsp = (mar + pccb * pcost) * cpi[B.BASE_YEAR] / gj
        def vat(y):                      # dataset rate if the cell is filled (0 included), else the assumption
            raw = dom[f'EGY{y}'].get(f'mit.vatrate.{pc}', '')
            if raw not in ('', None):
                return float(raw)
            return val('VAT_WEO', y) if pc in VAT_CONSUMER else 0.0
        vr = vat(last_hist)
        vrate[pc] = vr

        def hist_txo(y):
            vry = vat(y)
            raw = (val(f'mit.rp.{pc}', y) / (1 + vry) - val(f'mit.sp.{pc}', y) if fuel in B.RP_FROM_DATA
                   else val(f'mit.txo.{pc}', y))
            return raw * cpi[y] / gj
        sp_l, txo_l = val(f'mit.sp.{pc}', last_hist) * cpi[last_hist] / gj, hist_txo(last_hist)
        g = gp[{1: 'oil', 2: 'coa', 3: 'nga', 4: 'one'}[B.GP_POS[fuel]]]
        sp, txo = [], []
        for k, y in enumerate(years):
            if y <= last_hist:
                sp.append(val(f'mit.sp.{pc}', y) * cpi[y] / gj)
                txo.append(hist_txo(y))
            else:
                sp.append(fixsp + (sp[-1] - fixsp) * g[k] / g[k - 1])
                cs_l = txo_l * (1 - pcc)
                cs = (sp_l - sp[-1] + cs_l) * (1 - pcc)
                txo.append(txo_l * pcc + (max(cs_l, cs) if cs_l >= 0 else cs))
        rpb[pc] = [(a + b) * (1 + vr) for a, b in zip(sp, txo)]
        res[('sp', sec, fuel)], res[('txo', sec, fuel)], res[('rpb', sec, fuel)] = sp, txo, rpb[pc]
        res[('txf', sec, fuel)] = [(sp_l - s_ + txo_l * (1 - pcc)) * (1 - pcc) for s_ in sp]
        res[('vat', sec, fuel)] = [(a + b) * vr for a, b in zip(sp, txo)]
        res[('etx', sec, fuel)] = [max(t, 0.0) for t in txo]
        res[('esub', sec, fuel)] = [max(-t, 0.0) for t in txo]
        res[('esubpu', sec, fuel)] = [max(-t, 0.0) * gj for t in txo]
    for s, f in B.PAIRS:
        _, _, grp, es, efs, cgs, fbs = sub[s]
        _, _, elf, aef, rule, _ = fue[f]
        code = f'{f}.{cgs if rule == "sub" else "all"}'
        e = ef[f'EGY|{f}|{efs}']
        fcov = 1 if f == 'bio' else bool(mt[n['MCov' + f.capitalize()]])
        scov = sc_used(s, mt[n['MCov' + s.capitalize()]])
        ctc = cov(f'ecs.{ECP_SEC(s)}') * cov(f'ecf.{f}')
        etc = cov(f'ees.{ECP_SEC(s)}') * cov(f'eef.{f}')
        fbcov = bool(mt[66 + B.COV_SECTORS.index(s)])
        eY, eU, eF = el[f'inc|{elf}|{es}'], el[f'usg|{elf}|{es}'], el[f'eff|{elf}|{es}']
        a = el[f'aei|{aef}|{es}']
        v = {k: [] for k in B.VARS}
        for k, y in enumerate(years):
            v['ctxnew'].append(cp[y] * e * fcov * scov * (1 - ets_cov[s][y]))
            v['ets'].append(ets_pe[y] * e * ets_cov[s][y] * (1 - obr[s][y]))
            v['ntx'].append(fpr[code][y] / units[f])
            v['nce'].append(v['ctxnew'][-1] + v['ets'][-1] + v['ntx'][-1])
            v['xcp'].append(e * (ctc * xp['xct'][y] + etc * xp['xets'][y]))
            v['xcpd'].append(e * (ctc * (xp['xct'][y] - xp['xct0'][y]) + etc * (xp['xets'][y] - xp['xets0'][y])))
            v['atp'].append(max(rpb[code][k] + (v['nce'][-1] + v['xcpd'][-1]) * (1 + vrate[code]), 0.01))
            v['shp'].append((fb[fbs][y] * fbcov * 1.0 + ets_pe[y] * ets_cov[s][y] * obr[s][y]) * e)
            if y == B.BASE_YEAR:
                v['ener'].append(use0[(s, f)])
            else:
                ratio = v['atp'][-1] / v['atp'][-2]
                eff = (v['atp'][-1] + v['shp'][-1]) / (v['atp'][-2] + v['shp'][-2])
                v['ener'].append(v['ener'][-1] * (1 / (1 + a)) ** (1 + eU) * (1 + float(weo[str(y)])) ** eY
                                 * ratio ** eU * eff ** (eF * (1 + eU)))
        kpj = 0.041868                                   # PJ per ktoe (independent copy)
        k_code = [pc for pc, _l in B.PRICES].index(code)
        v['co2'] = [u * kpj * e for u in v['ener']]
        for k, y in enumerate(years):
            ext = res[('etx', code.split('.')[1], f)][k] + res[('vat', code.split('.')[1], f)][k]
            v['rtx'].append(v['ener'][k] * kpj * (ext + v['xcpd'][k] * (1 + vrate[code])))
            v['rxcp'].append(v['ener'][k] * kpj * v['xcp'][k])
            v['rsub'].append(v['ener'][k] * kpj * res[('esub', code.split('.')[1], f)][k])
            v['rnew'].append(v['ener'][k] * kpj * (v['ctxnew'][k] + v['ntx'][k] + v['ets'][k] / vadj
                                                    + v['nce'][k] * vrate[code]))
        for k in B.VARS:
            res[(k, s, f)] = v[k]
    for var in B.REV_VARS + ['co2']:
        res[(var, 'all', 'all')] = [sum(res[(var, s, f)][k] for s, f in B.PAIRS) for k in range(len(years))]
    res[('co2.ets', 'all', 'all')] = [sum(res[('co2', s, f)][k] * ets_cov[s][y] for s, f in B.PAIRS)
                                      for k, y in enumerate(years)]
    res[('rtot', 'all', 'all')] = [a - b + c for a, b, c in zip(res[('rtx', 'all', 'all')],
                                                                 res[('rsub', 'all', 'all')],
                                                                 res[('rnew', 'all', 'all')])]
    # 3. Power (engineer model, power_model_ref_v0_2 with all bug fixes): MVP prices and policies as inputs
    tech = {r['code']: r for r in read('power_tech.csv')}
    paths_p = {r['key']: [float(r[str(y)]) for y in years] for r in read('power_paths.csv')}
    spow = sc_used('pow', mt[n['MCovPow']])
    xck = []
    fbcov_pow = bool(mt[66])
    vbc, d5, shp = [], [], []
    for f in POWER_TECH:
        t = tech[f]
        pc, eff = t['price_code'], float(t['efficiency'])
        k_ = 0.0036 / eff if pc else 0.0
        e = ef.get(f'EGY|{f}|pow', 0.0) if pc else 0.0
        fu = pc.split('.')[0] if pc else ''
        fcov = 1 if fu == 'bio' else (bool(mt[n['MCov' + fu.capitalize()]]) if pc else 0)
        fixed = float(t['vop_usd_per_kwh']) + float(t['fuel_fixed_usd_per_kwh'])
        isfuel = f in [x[0] for x in B.FUELS]
        ctc_t = cov('ecs.pow') * cov(f'ecf.{f}') if isfuel else 0.0
        etc_t = cov('ees.pow') * cov(f'eef.{f}') if isfuel else 0.0
        xck.append([e * k_ * (ctc_t * (xp['xct'][y] - xp['xct0'][y]) + etc_t * (xp['xets'][y] - xp['xets0'][y]))
                    for y in years])
        vbc.append([fixed + (rpb[pc][k] * k_ if pc else 0.0) + xck[-1][k] for k in range(len(years))])
        d5.append([e * k_ * (cp[y] * fcov * spow * (1 - ets_cov['pow'][y]) + ets_pe[y] * ets_cov['pow'][y]
                             * (1 - obr['pow'][y])) + (fpr[pc][y] / units[fu] * k_ if pc else 0.0) for y in years])
        shp.append([(fb['pow'][y] * fbcov_pow * 1.0 + ets_pe[y] * ets_cov['pow'][y] * obr['pow'][y]) * e
                    for y in years])
    hist = {g: [val(f'mit.rp.ecy.{g}', y) * cpi[y] if y <= last_hist else 0.0 for y in years] for g in ('res', 'ind')}
    pinp = {'elast': PM.elasticities(), 'gdp': [float(weo[str(y)]) for y in years], 'inc': paths_p['incf'],
            'shock': paths_p['shock'], 'cpf': [0.0] * len(years),
            'rp2021_res': val('mit.rp.ecy.res', 2021) * cpi[2021], 'rp2021_ind': val('mit.rp.ecy.ind', 2021) * cpi[2021],
            'rphist_res': hist['res'], 'rphist_ind': hist['ind'], 'vat_res': val('mit.vatrate.ecy.res', last_hist),
            'vat_ind': val('mit.vatrate.ecy.ind', last_hist), 'last_hist': last_hist, 'vbc': vbc, 'd5': d5,
            'shp_gj': shp, **(power or {})}
    o = PM.engine(pinp, years, fix=PM.ALL_FIXES, shock=bool(shock_on))
    for v in B.PW_VARS:
        if v in ('xck', 'prtx', 'prsub', 'prxc', 'prnew', 'prxl', 'rnvat'):    # v1.09: set below
            continue
        kind = B.PW_KIND[v]
        for m in B.PW_MEMBERS.get(kind, ['']):
            if kind == 'tech':
                i = POWER_TECH.index(m)
                val_ = {'vbc': vbc[i], 'ccp': d5[i]}.get(v) or (paths_p[f'tcf|{m}'] if v == 'tcf' else o[v][i])
                res[(v, 'pow', m)] = val_
            elif kind == 'user':
                res[(v, m, 'ecy')] = o[f'{v}_{m}']
            elif kind == 'dgroup':
                res[(v, m, 'ecy')] = o[v][m]
            elif kind == 'dsub':
                res[(v, m, 'ecy')] = o['dem'][m]
            else:
                res[(v, '', '')] = paths_p[v] if v in B.PW_DATA else o[v]
    # v1.09: power revenues by type (M), VAT on the passed-through policy cost (G), totals with power
    par = PM.base_data(years)[2]
    ny = len(years)
    fflag = [1.0 if f in [x[0] for x in B.FUELS] else 0.0 for f in POWER_TECH]
    etr_pow = [ets_pe[y] * ets_cov['pow'][y] * (1 - obr['pow'][y]) for y in years]
    rf = 1 / vadj if vadj > 0 else 0.0
    pw = {v: [[0.0] * ny for _ in POWER_TECH] for v in ('prtx', 'prsub', 'prxc', 'prnew', 'prxl', 'xck')}
    for i, f in enumerate(POWER_TECH):
        pc = tech[f]['price_code']
        sec_, fu_ = (pc.split('.')[1], pc.split('.')[0]) if pc else ('pow', 'coa')   # no fuel: position 1 (x flag 0)
        isfuel = fflag[i] == 1.0
        ctc_t = cov('ecs.pow') * cov(f'ecf.{f}') if isfuel else 0.0
        etc_t = cov('ees.pow') * cov(f'eef.{f}') if isfuel else 0.0
        for k, y in enumerate(years):
            q = fflag[i] * o['pe'][i][k] * 0.041868
            pw['prtx'][i][k] = q * (res[('etx', sec_, fu_)][k] + res[('vat', sec_, fu_)][k])
            pw['prsub'][i][k] = q * res[('esub', sec_, fu_)][k]
            pw['xck'][i][k] = xck[i][k]
            pw['prxc'][i][k] = xck[i][k] * 8.76 * o['gen'][i][k]
            pw['prnew'][i][k] = o['gen'][i][k] * 8.76 * d5[i][k] - o['co2'][i][k] * etr_pow[k] * (1 - rf)
            pw['prxl'][i][k] = o['co2'][i][k] * (ctc_t * xp['xct'][y] + etc_t * xp['xets'][y])
    vat_g = {'res': pinp['vat_res'], 'ind': pinp['vat_ind']}
    rnvat = {g: [o['ccpav'][k] * par['pass'] * vat_g[g] * o[f'cons_{g}'][k] for k in range(ny)] for g in ('res', 'ind')}
    for v in pw:
        for i, m in enumerate(POWER_TECH):
            res[(v, 'pow', m)] = pw[v][i]
    for g in ('res', 'ind'):
        res[('rnvat', g, 'ecy')] = rnvat[g]
    tsum = lambda v: [sum(pw[v][i][k] for i in range(9)) for k in range(ny)]
    gsum = lambda v: [o[f'{v}_res'][k] + o[f'{v}_ind'][k] for k in range(ny)]
    rpow = {'rtx': [a_ + b_ + c_ for a_, b_, c_ in zip(tsum('prtx'), tsum('prxc'), gsum('rvat'))],
            'rsub': [a_ + b_ for a_, b_ in zip(tsum('prsub'), gsum('rgap'))],
            'rnew': [a_ + b_ for a_, b_ in zip(tsum('prnew'), rnvat['res'])],
            'rxcp': tsum('prxl')}
    rpow['rnew'] = [a_ + b_ for a_, b_ in zip(rpow['rnew'], rnvat['ind'])]
    for v in B.REV_VARS:
        res[(v, 'pow', 'all')] = rpow[v]
        res[(v, 'all', 'all')] = [a_ + b_ for a_, b_ in zip(res[(v, 'all', 'all')], rpow[v])]
    res[('rtot', 'all', 'all')] = [a - b + c for a, b, c in zip(res[('rtx', 'all', 'all')],
                                                                 res[('rsub', 'all', 'all')],
                                                                 res[('rnew', 'all', 'all')])]
    epow = [sum(fflag[i] * o['pe'][i][k] for i in range(9)) for k in range(ny)]
    res[('ener', 'pow', 'all')] = epow
    res[('ener', 'all', 'all')] = [sum(res[('ener', s, f)][k] for s, f in B.PAIRS) + epow[k] for k in range(ny)]
    pco2 = [sum(o['co2'][i][k] for i in range(9)) for k in range(len(years))]
    res[('co2', 'all', 'all')] = [a_ + b_ for a_, b_ in zip(res[('co2', 'all', 'all')], pco2)]
    res[('co2.ets', 'all', 'all')] = [a_ + b_ * ets_cov['pow'][y] for a_, b_, y in
                                      zip(res[('co2.ets', 'all', 'all')], pco2, years)]
    res[('co2.pow', '', '')] = pco2
    return res


POWER_TECH = ['coa', 'nga', 'oop', 'nuc', 'wnd', 'sol', 'hyd', 'ore', 'bio']    # independent copy (legacy order)


def check_power_method():
    """Method = legacy: power_model_ref_v0_2 with no bug fixes and the legacy demand shock, fed with the cached legacy
    inputs (fuel and policy costs D4-D6, GDP, prices of 2021 and the historical years), reproduces every cached legacy
    row of the engineer model and the price chain, for the baseline and the carbon-tax run."""
    ok = True
    for run, label in (('base', 'baseline'), ('ctax', 'carbon-tax run')):
        res = PM.legacy_test(run)
        worst = res[0]
        good = worst[0] < 1e-9 and len(res) >= 200
        ok &= good
        log(f'- {label}: {len(res)} legacy rows (demand by subsector, required generation, capacity, generation, shares, '
            f'fuel use, retirements, investment shares and both rounds, nameplate and USD investment, B1 phi, E1 '
            f'investment cost, C7, B2, B6, E3, J, G1 prices, early retirement), 2022-2040: max relative diff '
            f'{worst[0]:.1e} (at {worst[1]} {worst[2]}) {"OK" if good else "FAIL"}')
    years = list(range(2022, 2041))
    k40 = years.index(2040)
    log('')
    log('Effect of each bug fix on the legacy Egypt inputs (2040; % change against legacy):')
    log('')
    log('| Fix | legacy rows | generation (base) | power CO2 (base) | generation (carbon tax) | VRE share (carbon tax) | '
        'investment USD 2022-2040 (base) |')
    log('|---|---|---|---|---|---|---|')

    def kpi(o):
        return (sum(o['gen'][i][k40] for i in range(9)), sum(o['co2'][i][k40] for i in range(9)),
                o['vre'][k40], sum(o['usd'][i][k] for i in range(9) for k in range(len(years))))
    leg = {run: kpi(PM.engine(PM.legacy_inputs(run), years, fix=frozenset(), shock=True)) for run in ('base', 'ctax')}
    for name, (rows, _d) in list(PM.FIXES.items()) + [('all', ('', ''))]:
        fx = PM.ALL_FIXES if name == 'all' else frozenset([name])
        b = kpi(PM.engine(PM.legacy_inputs('base'), years, fix=fx, shock=True))
        c = kpi(PM.engine(PM.legacy_inputs('ctax'), years, fix=fx, shock=True))
        pct = lambda a, z: f'{(a / z - 1) * 100:+.2f}%' if z else '0'
        log(f'| {name} | {rows} | {pct(b[0], leg["base"][0])} | {pct(b[1], leg["base"][1])} | '
            f'{pct(c[0], leg["ctax"][0])} | {pct(c[2], leg["ctax"][2])} | {pct(b[3], leg["base"][3])} |')
    return ok


def scenario_mt(g):
    mt = mt_used()
    for k, v in B.MT_SCENARIO_INPUTS[g].items():
        mt[B.MT_NAMES[k]] = v
    return mt


def gp_rows():
    return [(B.R_GP0 + k, 'int', fu) for k, (fu, *_r) in enumerate(B.GP_ROWS)]


def compare(ws, g, exp, label):
    _, base, years = B.group_cols(g)
    worst, n = (0.0, None), 0
    pk = ('cptraj', 'ets.p', 'ets.pe', 'ets.cap', 'ets.est', 'ets.bce', 'ets.se', 'xct.p', 'xct.p0', 'xets.p',
          'xets.p0')
    rev = [(B.R_REV[f'{v}.all'], 'all', 'all') for v in B.REV_VARS + ['rtot']]
    for var, rows in ([(k, [(B.R_POL[k], '', '')]) for k in pk]
                      + [(v, [(B.R_POL[f'{v}.{s}'], s, '') for s in B.COV_SECTORS]) for v in ('bco2', 'obr')]
                      + [(v, [row]) for v, row in zip(B.REV_VARS + ['rtot'], rev)]
                      + [(v, [(B.R_REV[f'{v}.pow'], 'pow', 'all')]) for v in B.REV_VARS]
                      + [('ener', [(B.R_TOTAL, 'all', 'all'), (B.R_EPOW, 'pow', 'all')])]
                      + [('co2', [(B.R_EMI['co2.all'], 'all', 'all')])]
                      + [('co2.ets', [(B.R_EMI['co2.ets'], 'all', 'all')])] + [('gp', gp_rows())]
                      + [('co2.pow', [(B.R_EMI['co2.pow'], '', '')])]
                      + [(v, rows_of(v)) for v in B.PVARS + B.VARS]
                      + [(v, [(r, c_, b_) for r, v_, b_, c_, _k in B.PW_ROWS if v_ == v]) for v in B.PW_VARS]):
        for r, s, f in rows:
            for k, c in enumerate([base] + years):
                v, e = ws.cell(r, c).value, exp[(var, s, f)][k]
                d = float('inf') if not isinstance(v, (int, float)) else abs(v - e) / max(1.0, abs(e))
                n += 1
                if d > worst[0]:
                    worst = (d, f'{L(c)}{r}')
    ok = worst[0] < TOL
    log(f'- {label}: carbon price, ETS rows (price, tax-equivalent, cap, estimate, baseline, semi-elasticity, '
        f'baseline CO2 and OBR share by sector), revenue, CO2 and covered CO2 totals, gp, {len(B.PVARS)} price-fuel and {len(B.VARS)} '
        f'subsector variables, {len(B.PW_VARS)} power variables, {n} cells, max relative '
        f'diff {worst[0]:.2e} (at {worst[1]}) {"OK" if ok else "FAIL"}')
    return ok


# ------------------------------------------------------------------ checks
def check_shipped(ws):
    """LibreOffice cannot evaluate LAMBDA, so the right column shows errors. The levelised variable cost (rows dsum,
    lvc) looks ahead to the right column, so in LibreOffice its errors also reach the earlier years of the power block
    and the CO2 rows of section 13 that include power (Excel evaluates them; the expanded variant checks the values).
    Anywhere else an error is a failure."""
    wbv = openpyxl.load_workbook(recalc_path(B.OUT), data_only=True)
    ok, parts = True, []
    for wsx in wbv.worksheets:
        errs = [f'{c.coordinate}' for row in wsx.iter_rows() for c in row if is_err(c.value)]
        ok &= not errs
        if errs:
            parts.append(f'{wsx.title}: {len(errs)} {errs[:6]}')
    wbf = openpyxl.load_workbook(B.OUT)
    calls = [f'{w.title}!{c.coordinate}' for w in wbf.worksheets for row in w.iter_rows() for c in row
             if isinstance(c.value, str) and c.value.startswith('=') and LAMBDA_CALL.search(c.value)]
    ok &= not calls
    log(f'- no cell calls a LAMBDA: {"OK" if not calls else "FAIL " + str(calls[:6])}')
    log(f'- shipped file recalculated by LibreOffice (no LAMBDA support): error values in all '
        f'{len(wbv.worksheets)} sheets: {"none OK" if not parts else "FAIL " + "; ".join(parts)}')
    return ok


def uniform(ws, cells, label):
    forms = {r1c1(ws.cell(r, c).value, r, c) for r, c in cells}
    ok = len(forms) == 1
    log(f'- {label}: {len(cells)} cells, {len(forms)} formula {"OK" if ok else "FAIL"}')
    return ok, next(iter(forms))


def check_uniformity(wb):
    """One relative-R1C1 formula per variable and column block: section-2 prices have a history block (base year to
    the last historical year, data) and a projection block; every other variable a base-year column and a plain
    block that runs to the right column (v1.08: no LAMBDA in cells); sheet Lambdas lists the optional LAMBDA formula
    of every calculated row."""
    ws = wb['Mitigation']
    ok = True
    grp = [B.group_cols(g) for g, *_ in B.SCENARIOS]
    nh = B.HIST_YEARS[-1] - B.BASE_YEAR           # plain years in the history block after the base year
    for var in B.PVARS + B.VARS:
        rows = [r for r, _s, _f in rows_of(var)]
        if var in ('sp', 'txo'):
            blocks = [(f'history {B.BASE_YEAR}-{B.HIST_YEARS[-1]}', lambda b, ys: [b] + ys[:nh]),
                      (f'projection {B.HIST_YEARS[-1] + 1}-{B.YEARS[-1]}', lambda b, ys: ys[nh:])]
        elif var in B.PVARS:
            blocks = [(f'{B.BASE_YEAR}-{B.YEARS[-1]}', lambda b, ys: [b] + ys)]
        else:
            blocks = [('base year', lambda b, ys: [b]), (f'{B.YEARS[0]}-{B.YEARS[-1]}', lambda b, ys: ys)]
        good, forms = True, []
        for label, cols in blocks:
            g, f = uniform(ws, [(r, c) for r in rows for _, b, ys in grp for c in cols(b, ys)], f'`{var}` {label}')
            good &= g
            forms.append(f)
        g4 = all(uniform(ws, [(r, CI(pc)) for r in rows], f'`{var}` column {pc}')[0] for pc in B.PARAM_COLS)
        for f in forms:
            log(f'  - `{f[:150]}`')
        if var in ('sp', 'txo'):
            plain_if = any('Settings!R10C3' in f for f in forms)      # the last-historical-year switch
            log(f'  - history and projection blocks without the history/projection IF (no reference to the last '
                f'historical year): {"OK" if not plain_if else "FAIL"}')
            good &= not plain_if
        ok &= good and g4
    mt_rows = [B.R_POL[k] for k, _v, _f, _s, kind, _m in B.POL if kind == 'mt'] + [B.R_PRI[p] for p in B.PRI_PARAMS]
    ok &= uniform(ws, [(r, c) for r in mt_rows for _, b, ys in grp for c in [b] + ys],
                  'Section 1 and 2 MTInputs rows')[0]
    ok &= uniform(ws, [(r, c) for r, *_ in gp_rows() for _, b, ys in grp for c in [b] + ys],
                  'section 2 international prices `gp` (data rows, plain)')[0]
    ok &= uniform(ws, [(B.R_GPN0 + k, c) for k in range(len(B.GP_ROWS)) for _, b, ys in grp for c in [b] + ys],
                  'section 2 international prices before the adjustment `gpn` (data rows, plain)')[0]
    for k_ in ('gp.src', 'gp.adj'):
        ok &= uniform(ws, [(B.R_PRI[k_], c) for _, b, ys in grp for c in [b] + ys], f'section 2 `{k_}` (plain)')[0]
    eld = [h + 1 for h in B.SUB_HEAD.values()]
    ok &= uniform(ws, [(r, c) for r in eld for _, b, ys in grp for c in [b] + ys],
                  'electricity demand memo lines in the sector blocks')[0]
    for r, lab in ((B.R_CPI, 'infl'), (B.R_DEFL, 'defl')):
        ok &= uniform(ws, [(r, c) for _, b, ys in grp for c in [b] + ys], f'index row `{lab}`')[0]
    calc = [k for k, _v, _f, _s, kind, _m in B.POL if kind == 'calc']
    exact = ('cptraj', 'cp.t', 'cp.nom', 'ets.on', 'ets.vol', 'ets.imp', 'ets.vadj', 'ets.bce', 'ets.chg', 'ets.cap',
             'ets.se', 'ets.est', 'ets.p', 'ets.pe', 'ets.rf')
    for prefix in exact + ('fpr.', 'fb.', 'etsc.', 'etsb.', 'obr.', 'shps.', 'ssc.', 'ctr.', 'etr.', 'shpx.'):
        keys = [k for k in calc if (k == prefix if prefix in exact else k.startswith(prefix))]
        assert keys, prefix
        ok &= uniform(ws, [(B.R_POL[k], c) for k in keys for _, b, ys in grp for c in [b] + ys],
                      f'Section 1 `{prefix}*` plain')[0]
    heads = list(B.SUB_HEAD.values())
    ok &= uniform(ws, [(h, c) for h in heads for _, b, ys in grp for c in [b] + ys], 'subsector heading totals')[0]
    for var in B.PW_VARS:                       # section 3 (power)
        rows = [r for r, v, *_ in B.PW_ROWS if v == var]
        if var in B.PW_DATA or var in B.PW_SUMS or var in B.PW_NOLAM:
            blocks = [(f'{B.BASE_YEAR}-{B.LAST_YEAR} (data, sum or look-ahead: one formula)', lambda b, ys: [b] + ys)]
        elif var == 'rppre':
            blocks = [(f'history {B.BASE_YEAR}-{B.HIST_YEARS[-1]}', lambda b, ys: [b] + ys[:nh]),
                      (f'projection {B.HIST_YEARS[-1] + 1}-{B.YEARS[-1]}', lambda b, ys: ys[nh:])]
        elif var in B.PW_BASE:
            blocks = [('base year', lambda b, ys: [b]), (f'{B.YEARS[0]}-{B.YEARS[-1]}', lambda b, ys: ys)]
        else:
            blocks = [(f'{B.BASE_YEAR}-{B.YEARS[-1]}', lambda b, ys: [b] + ys)]
        good = True
        for label, cols in blocks:
            g, f = uniform(ws, [(r, c) for r in rows for _, b, ys in grp for c in cols(b, ys)], f'power `{var}` {label}')
            good &= g
            log(f'  - `{f[:150]}`')
        good &= all(uniform(ws, [(r, CI(pc)) for r in rows], f'power `{var}` column {pc}')[0] for pc in B.PARAM_COLS)
        ok &= good
    lam = wb['Lambdas']
    lam_rows = {lam.cell(i, 1).value: lam.cell(i, 6).value for i in range(5, lam.max_row + 1) if lam.cell(i, 1).value}
    calc_rows = ({r for v in B.PVARS + B.VARS for r, *_ in rows_of(v)} | {B.R_POL[k] for k in calc}
                 | {r for r, v, *_ in B.PW_ROWS if v not in B.PW_DATA | B.PW_SUMS | B.PW_NOLAM})
    good = set(lam_rows) == calc_rows and all(LAMBDA_CALL.search('=' + f) for f in lam_rows.values())
    ok &= good
    log(f'- sheet Lambdas: an optional LAMBDA formula for each of the {len(calc_rows)} calculated rows '
        f'({len(lam_rows)} listed, each calls a named LAMBDA) '
        f'{"OK" if good else "FAIL " + str(sorted(calc_rows ^ set(lam_rows))[:10])}')
    cd = ws.column_dimensions
    ycols = [L(c) for _, b, ys in grp for c in ys if 2030 <= B.year_of(c) <= 2039]
    grouped = list('BCDEFGIJK') + ycols
    visible = ['A', 'H'] + [L(c) for _, b, ys in grp for c in [b] + ys if not 2030 <= B.year_of(c) <= 2039]
    hidden = (all(cd[c].hidden and cd[c].outline_level == 1 for c in grouped)
              and not any(cd[c].hidden or (cd[c].outline_level or 0) > 0 for c in visible)
              and max((cd[c].outline_level or 0) for c in cd) == 1)
    log(f'- column groups, one level: B:G, I:K and 2030-2039 of each scenario ({len(ycols)} columns), rolled up; '
        f'other columns visible: {"OK" if hidden else "FAIL"}')
    spacer = all(ws.cell(r, ys[-1] + 1).value is None for _, b, ys in grp for r in range(1, ws.max_row + 1)) \
        and all(code == B.COL_G0 + (g - 1) * B.GW for g, (code, b, ys) in enumerate(grp, 1)) and B.GW == B.NY + 2
    log(f'- blank spacer column after each scenario group (group width {B.GW}): {"OK" if spacer else "FAIL"}')
    light = 'EEECE1'
    lam_ok = not any(ws.cell(r, ys[-1]).fill.start_color.rgb.endswith(light) for _, b, ys in grp
                     for r in calc_rows)          # v1.08: the right column is an ordinary column
    nh = B.HIST_YEARS[-1] - B.BASE_YEAR
    hist_ok = all(ws.cell(r, c).fill.start_color.rgb.endswith(light) for v in ('sp', 'txo', 'rppre') for r, *_ in rows_of(v)
                  for _, b, ys in grp for c in ys[:nh]) \
        and all(ws.cell(r, b).fill.start_color.rgb.endswith('DDD9C4') for v in ('sp', 'txo', 'rppre') for r, *_ in rows_of(v)
                for _, b, ys in grp) \
        and not any(ws.cell(r, ys[nh]).fill.start_color.rgb.endswith(light) for v in ('sp', 'txo')
                    for r, *_ in rows_of(v) for _, b, ys in grp)
    log(f'- light beige (EEECE1): right column not marked (plain formulas) {"OK" if lam_ok else "FAIL"}; 2023-2024 history block of sp, txo, rppre '
        f'(base year darker beige, projection white) {"OK" if hist_ok else "FAIL"}')
    hidden = hidden and spacer and lam_ok and hist_ok
    return ok and hidden


def check_encoding():
    ok = True
    for name, (params, body) in B.LAMBDAS.items():
        xml = B.lambda_xml(name)
        bare = [p for p in params if re.search(rf'(?<![A-Za-z_.]){p}(?![A-Za-z_\d(])', xml)]
        good = not bare and xml.startswith('_xlfn.LAMBDA(')
        ok &= good
        log(f'- {name}: {len(params)} parameters, {xml.count("_xlpm.")} _xlpm. tokens, bare names {bare} '
            f'{"OK" if good else "FAIL"}')
    return ok


def mt_set(wsm, col, row_or_name, value):
    r = row_or_name if isinstance(row_or_name, int) else B.MT_NAMES[row_or_name]
    wsm.cell(r, col).value = value


def add_scenario(wb, src_g, dst_g, name, edits, mit_edits=None):
    """Copy MTInputs column of src_g to dst_g's column (keeping the auto number) and the Mitigation group."""
    ws, wsm = wb['Mitigation'], wb['MTInputs']
    ms, md = B.MT_COL0 + src_g - 1, B.MT_COL0 + dst_g - 1
    for r in range(1, 416):
        v = wsm.cell(r, ms).value
        if isinstance(v, str) and v.startswith('='):
            v = Translator(v, origin=f'{L(ms)}{r}').translate_formula(f'{L(md)}{r}')
        wsm.cell(r, md).value = v
    wsm.cell(B.MT_ROW_SCEN, md).value = f'={L(md - 1)}{B.MT_ROW_SCEN}+1'
    wsm.cell(B.MT_ROW_NAME, md).value = name
    for k, v in edits.items():
        mt_set(wsm, md, k, v)
    s0, d0 = B.group_cols(src_g)[0], B.group_cols(dst_g)[0]
    for k in range(B.GW):
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, s0 + k).value
            if isinstance(v, str) and v.startswith('='):
                v = Translator(v, origin=f'{L(s0 + k)}{r}').translate_formula(f'{L(d0 + k)}{r}')
            ws.cell(r, d0 + k).value = v
    ws.cell(B.R_SCEN, d0).value = f'={L(d0 - B.GW)}{B.R_SCEN}+1'
    for key, v in (mit_edits or {}).items():          # typed Section-1 rows (benchmarks), every year column
        for c in [d0 + 1] + B.group_cols(dst_g)[2]:
            ws.cell(B.R_POL[key], c).value = v


LEGACY_CP = {2022: 0, 2023: 0, 2024: 0, 2025: 0, 2026: 0, 2027: 12.5, 2028: 25, 2029: 37.5, 2030: 50,
             2031: 62.5, 2032: 75, 2033: 87.5, 2034: 100, 2035: 112.5}   # legacy Mitigation row 8582
FPR_TEST = {140: 2027, 141: 2030, 150: 0.05, 164: 0.20, 146: 0.5, 160: 2.0}   # gasoline $/l, gas res $/GJ
ETS_TEST = {'D_NewETS': 'Yes', 'D_ETSIntro': 2027, 'D_ETSOutro': 2035, 'D_ETSChangeRelStart': -0.02,
            'D_ETSChangeRelTarget': -0.15, 'D_ETSCapCont': 'Constant'}   # template coverage: power, mch, irn, nfm, cem
ETS_EQUIV = {**ETS_TEST, 'D_ETSPriceOverride': 'Yes', 'D_ETSVolatility': 'Zero',   # blank override: carbon price;
             'D_ETSCTRisk': 0}                                                    # adjustment 1
BENCH0 = {f'{k}.{g}': 0 for k in ('etsbs', 'etsbt') for g, *_ in B.ETS_GROUPS}      # full auctioning
PRICE_TEST = {'IntEnerPricForeSource': 'IMF-WB*', 'IntEnerPricForecastAdjustment': 'High',
              'NomorReal': 'Nominal'}
FB_TEST = {'D_FeebateIntro': 2027, 'D_FeebateOutro': 2030, 'D_Feb_Level_Start_Trans': 10,
           'D_Feb_Level_Target_Trans': 50, 'D_Feb_Level_Start_Ind': 5, 'D_Feb_Level_Target_Ind': 25,
           67: True, 74: True, 75: True}                                  # feebate coverage: rod, mch, irn


def check_scenarios():
    wb = openpyxl.load_workbook(B.OUT)
    zero = {'CPLevelStart': 0, 'CPLevelTarget': 0}
    add_scenario(wb, 2, 3, 'Test 3: copy of 2', {})
    add_scenario(wb, 3, 4, 'Test 4: copy of 2, carbon price 0', zero)
    add_scenario(wb, 4, 5, 'Test 5: legacy Egypt carbon tax', {'CPIntro': 2026, 'CPLevelStart': 0,
                                                                'CPLevelTarget': 50, 'CPOutro': 2030})
    add_scenario(wb, 4, 6, 'Test 6: fuel price reform', FPR_TEST)
    add_scenario(wb, 4, 7, 'Test 7: feebates', FB_TEST)
    add_scenario(wb, 2, 8, 'Test 8: IMF-WB*, High, nominal carbon price', PRICE_TEST)
    add_scenario(wb, 2, 9, 'Test 9: new ETS with cap (fast estimate) + carbon tax elsewhere', ETS_TEST)
    add_scenario(wb, 2, 10, 'Test 10: ETS override blank, no volatility, full auctioning', ETS_EQUIV, BENCH0)
    wb_t = recalc_wb(variant(add_mto_blocks(wb, range(3, 11)), 'expand'))
    ws = wb_t['Mitigation']
    ok = True
    for a, b, label in [(2, 3, 'scenario 3 (copy of 2) = scenario 2'),
                        (1, 4, 'scenario 4 (copy of 2, carbon price 0 in MTInputs) = scenario 1')]:
        ca, cb = B.group_cols(a), B.group_cols(b)
        diff, n = 0.0, 0
        for r in range(B.R_YEAR, ws.max_row + 1):
            if r in (B.R_SCEN, B.R_NAME):
                continue
            for x, y in zip([ca[1]] + ca[2], [cb[1]] + cb[2]):
                va, vb = ws.cell(r, x).value, ws.cell(r, y).value
                if isinstance(va, (int, float)) and isinstance(vb, (int, float)):
                    if b == 4 and r in (B.R_POL['CPLevelStart'], B.R_POL['CPLevelTarget']):
                        continue
                    diff, n = max(diff, abs(va - vb)), n + 1
                elif va != vb:
                    diff = float('inf')
        good = diff < TOL
        ok &= good
        log(f'- {label}: {n} numeric cells, max abs diff {diff:.3g} {"OK" if good else "FAIL"}')
    for g in (3, 4, 5, 6, 7, 8, 9, 10):
        code, base, _ = B.group_cols(g)
        num_, name = ws.cell(B.R_SCEN, code).value, ws.cell(B.R_NAME, base).value
        c = ws.cell(B.SUB_HEAD['rod'] + B.VOFF['ener'], code).value
        good = num_ == g and str(c).endswith(f'.e.{g}') and str(name).startswith(f'Test {g}')
        ok &= good
        log(f'- pasted group {g}: number {num_}, name "{name}", code `{c}` {"OK" if good else "FAIL"}')
    _, base, years = B.group_cols(5)
    got = {ws.cell(B.R_YEAR, c).value: ws.cell(B.R_CP, c).value for c in [base] + years}
    d = max(abs(got[y] - v) for y, v in LEGACY_CP.items())
    ok &= d < TOL
    log(f'- scenario 5 (legacy Egypt carbon tax): carbon price {[got[y] for y in (2026, 2027, 2030, 2035)]} vs '
        f'legacy row 8582, max abs diff {d:.3g} {"OK" if d < TOL else "FAIL"}')
    base_mt = scenario_mt(2)
    for g, edits in ((5, {'CPIntro': 2026, 'CPLevelStart': 0, 'CPLevelTarget': 50, 'CPOutro': 2030}),
                     (6, {**zero, **FPR_TEST}), (7, {**zero, **FB_TEST}), (8, PRICE_TEST), (9, ETS_TEST),
                     (10, ETS_EQUIV)):
        mt = dict(base_mt)
        for k, v in edits.items():
            mt[k if isinstance(k, int) else B.MT_NAMES[k]] = v
        bench = {g_: (0, 0) for g_, *_ in B.ETS_GROUPS} if g == 10 else None
        ok &= compare(ws, g, recompute(mt, bench=bench), f'scenario {g} vs Python')
    r_ntx = B.SUB_HEAD['rod'] + B.VOFF['ntx'] + 2
    r_shp = B.SUB_HEAD['rod'] + B.VOFF['shp'] + 2
    _, b6, y6 = B.group_cols(6)
    _, b7, y7 = B.group_cols(7)
    log(f'  - scenario 6, road gasoline new excise ($/GJ) 2026-2031: '
        f'{[round(ws.cell(r_ntx, c).value, 3) for c in y6[3:9]]}')
    log(f'  - scenario 7, road gasoline feebate shadow price ($/GJ) 2026-2031: '
        f'{[round(ws.cell(r_shp, c).value, 3) for c in y7[3:9]]}')
    (_, b1, y1), (_, b4, y4) = B.group_cols(1), B.group_cols(7)
    mt7 = dict(scenario_mt(2))
    for k, v in {**{'CPLevelStart': 0, 'CPLevelTarget': 0}, **FB_TEST}.items():
        mt7[k if isinstance(k, int) else B.MT_NAMES[k]] = v
    rate = {s: mt7[B.MT_NAMES[b]] for s, a, b in B.FEEBATE_SECTORS}
    fbsec = {s[0]: s[6] for s in B.SUBSECTORS}
    covered = {s for s, *_ in B.SUBSECTORS if mt7[66 + B.COV_SECTORS.index(s)] and rate[fbsec[s]]}
    d_unc = max(abs(ws.cell(r, x).value - ws.cell(r, y).value) for r, s, _f in B.var_rows('ener')
                if s not in covered for x, y in zip([b1] + y1, [b4] + y4))
    lower = [(s, ws.cell(B.SUB_HEAD[s], y4[-2]).value / ws.cell(B.SUB_HEAD[s], y1[-2]).value - 1)
             for s in covered if ws.cell(B.SUB_HEAD[s], y1[-2]).value > 0]   # mch: no fuel use in the base data
    good = d_unc < TOL and len(lower) >= 2 and all(x < 0 for _s, x in lower)
    ok &= good
    log(f'- scenario 7 vs scenario 1 (covered with a non-zero rate: {", ".join(sorted(covered))}): '
        f'other subsectors unchanged (max abs diff {d_unc:.3g}); covered subsectors '
        f'in {B.YEARS[-2]}: {", ".join(f"{s} {x:+.2%}" for s, x in sorted(lower))} {"OK" if good else "FAIL"}')
    _, b8, y8 = B.group_cols(8)
    r_oil = B.R_GP0
    log(f'  - scenario 8: crude oil gp 2024-2026 (real) {[round(ws.cell(r_oil, c).value, 2) for c in y8[1:4]]}, '
        f'carbon price 2027-2030 (nominal 20 x infl) {[round(ws.cell(B.R_CP, c).value, 3) for c in y8[4:8]]}')
    # ETS with blank override, no volatility and full auctioning = carbon tax: same nce and fuel use as scenario 2
    (_, b2, y2), (_, b9, y9) = B.group_cols(2), B.group_cols(10)
    d_nce = max(abs(ws.cell(r, x).value - ws.cell(r, y).value) for v in ('nce', 'atp', 'ener')
                for r, *_ in B.var_rows(v) for x, y in zip([b2] + y2, [b9] + y9))
    ets_sum = sum(ws.cell(r, y9[-2]).value for r, *_ in B.var_rows('ets'))
    ctx_cem = sum(ws.cell(r, y9[-2]).value for r, s_, _f in B.var_rows('ctxnew') if s_ == 'cem')
    good = d_nce < TOL and ets_sum > 0 and abs(ctx_cem) < TOL
    ok &= good
    log(f'- scenario 10 (new ETS on power and industry from 2027, override row blank = carbon price, volatility '
        f'Zero, policy risk 0, benchmarks 0): nce, atp and fuel use equal scenario 2 (max abs diff {d_nce:.3g}); ETS cost > 0 '
        f'({ets_sum:.2f}, sum over rows {B.YEARS[-2]}); carbon tax in cement 0 ({ctx_cem:.3g}) '
        f'{"OK" if good else "FAIL"}')
    _, b9, y9 = B.group_cols(9)
    row = lambda k, sec=B.R_POL: [round(ws.cell(sec[k], c).value, 3) for c in y9[3:14]]
    log(f'  - scenario 9 (cap -2% in 2027 to -15% in 2035 vs baseline covered emissions; benchmarks 1.0 -> 0.8; '
        f'volatility Medium), 2026-2036:')
    for k, lab in (('ets.bce', 'baseline covered emissions (MtCO2)'), ('ets.cap', 'cap (MtCO2)'),
                   ('ets.vadj', 'volatility adjustment'), ('ets.se', 'effective semi-elasticity'),
                   ('ets.est', 'fast estimate ($/tCO2)'), ('ets.p', 'permit price used'),
                   ('obr.cem', 'OBR share, cement')):
        log(f'    - {lab}: {row(k)}')
    for k, lab in (('co2.ets', 'covered emissions (MtCO2)'), ('co2.gap', 'covered / cap - 1'),
                   ('ets.next', 'next price proposal')):
        log(f'    - {lab}: {row(k, B.R_EMI)}')
    va = ws.cell(B.R_POL['ets.vadj'], y9[-2]).value
    good9 = abs(va - 1 / 1.1) < 1e-12
    ok &= good9
    log(f'  - volatility adjustment at Medium = 1/1.1 as legacy row 2191 ({va:.6f}) {"OK" if good9 else "FAIL"}')
    chk = max(abs(ws.cell(B.R_EMI['bco2.chk'], c).value) for g in range(1, 11)
              for c in [B.group_cols(g)[1]] + B.group_cols(g)[2])
    ok &= chk < 1e-9
    log(f'- bco2.chk (baseline data vs scenario 1 CO2) in all 10 groups: max abs {chk:.3g} '
        f'{"OK" if chk < 1e-9 else "FAIL"}')
    for controls in ('None', 'Manual'):          # global setting: MTInputs scenario-1 column
        wbc = openpyxl.load_workbook(B.OUT)
        wbc['MTInputs'].cell(B.MT_NAMES['GovPriceControls'], B.MT_COL0).value = controls
        wsc = recalc(variant(wbc, 'expand'))
        for g in (1, 2):
            ok &= compare(wsc, g, recompute(scenario_mt(g), controls),
                          f'price controls {controls} (global), scenario {g} vs Python')
    ok &= check_mto(wb_t, list(range(1, 11)), 'with blocks 3-10 added by copying (scenario tests)')
    return ok


PRICE_VARS = {'sp', 'txo', 'rpb', 'vat', 'etx', 'esub', 'esubpu', 'atp', 'ener', 'co2', 'rtx', 'rsub', 'rnew',
              'rnet', 'rtot'}


POWER_OLD = {'tcf', 'gns', 'phi', 'cbat', 'cint', 'obat', 'cel', 'cons', 'cax', 'caxav', 'stoav', 'fix', 'lfx', 'vbc',
             'ccp', 'gnc', 'gncav', 'ccpav', 'sc', 'rppre', 'rp', 'sgap', 'rvat', 'rcarb', 'rgap', 'vre', 'cph', 'msc'}


def foo_dependent(k):
    """v1.09 intended changes: power generation fuel use enters the fuel totals (total, by fuel, reference and
    change rows of section 11) and power revenues enter the revenue totals (rtx, rsub, rnew, rnet, rtot and the
    reference and change rows). Egypt has no existing carbon pricing, so nothing else changes."""
    p = k.split('.')
    if p[2] == 'ener':
        return p[3] in ('all', 'ref', 'pct', 'chk')
    return p[2] in ('rtx', 'rsub', 'rnew', 'rnet', 'rtot') and p[3] in ('all', 'ref', 'chg')


def _unused(k):
    """Fuel-use codes that change with the food & forestry elasticities (v0.11): foo itself and its aggregates."""
    p = k.split('.')
    return p[2] == 'ener' and (p[3] in ('foo', 'all', 'bld', 'ref', 'pct', 'chk'))


def check_regression(ws_new):
    src_old = os.path.join(HERE, 'Old', 'build_v1_08.py')
    bo = types.ModuleType('bold')                  # run as if in this folder (it reads data/ at import)
    bo.__file__ = os.path.join(HERE, 'build_v1_08.py')
    exec(compile(open(src_old, encoding='utf-8').read(), src_old, 'exec'), bo.__dict__)
    wbo = openpyxl.load_workbook(os.path.join(HERE, 'Old', 'CPAT-AI-Mitigation-MVP-v1.08.xlsm'))
    wso = wbo['Mitigation']
    for g in (1, 2):
        c = bo.group_cols(g)[2][-1]
        for r in range(1, wso.max_row + 1):
            v = wso.cell(r, c).value
            if isinstance(v, str) and v.startswith('='):
                wso.cell(r, c).value = expand(v, bo.LAMBDAS)
    path = os.path.join(tempfile.mkdtemp(), 'old.xlsx')
    wbo.save(path)
    old = recalc(path)
    vo, vn = {}, {}
    for g in (1, 2):
        for src, mod, dst in ((old, bo, vo), (ws_new, B, vn)):
            code, base, years = mod.group_cols(g)
            for r in range(1, src.max_row + 1):
                k = src.cell(r, code).value
                if isinstance(k, str) and k.startswith('egy.'):
                    dst[k] = [src.cell(r, c).value for c in [base] + years]
    shared = set(vo) & set(vn)
    intended = {k for k in shared if foo_dependent(k)}
    common = sorted(shared - intended)
    dk = {k: max([abs(a - b) for a, b in zip(vo[k], vn[k])          # zip: years both versions have (to 2035)
                  if isinstance(a, (int, float)) and isinstance(b, (int, float))] or [0]) for k in common}
    diff = max(dk.values())
    ok = diff < TOL and len(common) > 1000
    log(f'- {len(common)} output codes in both versions ({len(intended)} intended changes), '
        f'{B.BASE_YEAR}-{B.LAST_YEAR}: max abs diff {diff:.3g} {"OK" if ok else "FAIL " + str([k for k in dk if dk[k] > TOL][:8])}')
    new = sorted({k.split('.')[2] for k in set(vn) - set(vo)})
    log(f'- variables added: {", ".join(new)}')
    fam = sorted({'.'.join(k.split('.')[2:4]) for k in intended if max(
        [abs(a - b) for a, b in zip(vo[k], vn[k]) if isinstance(a, (int, float)) and isinstance(b, (int, float))]
        or [0]) > TOL})
    log(f'- intended changes that differ (power in the fuel and revenue totals): {", ".join(fam)}')
    return ok


ECP_TEST = {  # existing carbon pricing test data for EGY (nominal $/tCO2 by year; coverage shares, all years)
    'ecr.all': {2021: 5.0, 2022: 5.0, 2023: 6.0, 2024: 8.0, 2025: 10.0}, 'ecs.trs': 1.0, 'ecs.res': 0.5,
    'ecf.gso': 1.0, 'ecf.die': 1.0, 'ecf.nga': 1.0, 'ecf.lpg': 0.5,
    'eer.all': {2021: 0.0, 2022: 0.0, 2023: 0.0, 2024: 15.0, 2025: 20.0}, 'ees.pow': 1.0, 'ees.ind': 0.6,
    'eef.coa': 1.0, 'eef.nga': 1.0, 'eef.oop': 1.0}
ECP_SCEN = {3: ('Test 3: existing pricing grows 3% / 2% a year', {'ExistingCTGrowth': 0.03, 'ExistingETSGrowth': 0.02}),
            4: ('Test 4: existing pricing switched off', {'ExistingCTApply': 'No', 'ExistingETSApply': 'No'}),
            5: ('Test 5: new carbon tax complementary to the existing ETS', {'CTaxComplimentaryToETS': 'Yes'})}


STALE = ['not yet in the section-12', 'Power and electricity are not yet included', 'power revenues stay',
         'power 0 until', 'no power sector', '; no power,', 'Power ETS revenue about 10% high', 'interim data until',
         'not in the fuel totals above)']
DOC_REF = re.compile(r'[A-Za-z][\w-]*_v\d+(?:[._]\d+)*\.(?:md|py|docx|bas)')


def check_consistency(ws):
    """v1.09 full consistency check of the shipped workbook: check rows and accounting identities in every scenario
    group, file references in the workbook texts exist (current versions, not Old/), version strings, no stale
    statements about power or existing carbon pricing."""
    ok = True
    groups = [g for g, *_ in B.SCENARIOS]
    cols = [c for g in groups for c in [B.group_cols(g)[1]] + B.group_cols(g)[2]]
    v = lambda r, c: ws.cell(r, c).value
    chk = {k: max(abs(v(r, c)) for c in cols) for k, r in (('ener.chk', B.R_CHK), ('rev.chk', B.R_REV['rev.chk']),
                                                          ('co2.chk', B.R_EMI['co2.chk']),
                                                          ('bco2.chk', B.R_EMI['bco2.chk']))}
    good = all(x < 1e-6 for x in chk.values())
    ok &= good
    log(f'- check rows in groups {groups}: ' + ', '.join(f'{k} {x:.2g}' for k, x in chk.items()) +
        f' {"OK" if good else "FAIL"}')
    R_ = B.R_REV
    iden = []
    for c in cols:
        iden.append(abs(v(R_['rnet.all'], c) - (v(R_['rtx.all'], c) - v(R_['rsub.all'], c))))
        iden.append(abs(v(R_['rtot.all'], c) - (v(R_['rnet.all'], c) + v(R_['rnew.all'], c))))
        iden.append(abs(v(B.R_TOTAL, c) - sum(v(B.SEC_SUM[g], c) for _n, g, _t in B.SECTIONS) - v(B.R_EPOW, c)))
        pe = sum(v(B.R_PW[('pe', m)], c) * (m in [f for f, *_ in B.FUELS]) for m, _n in B.PW_TECH)
        iden.append(abs(v(B.R_EPOW, c) - pe))
        co2p = sum(v(B.R_PW[('co2', m)], c) for m, _n in B.PW_TECH)
        iden.append(abs(v(B.R_EMI['co2.pow'], c) - co2p))
        iden.append(abs(v(B.R_EMI['co2.sec.pow'], c) - co2p))
    d = max(iden)
    good = d < 1e-6
    ok &= good
    log(f'- identities in every column: rnet = rtx - rsub, rtot = rnet + rnew, fuel use = sector totals + power, '
        f'power fuel use = fuel types of pe, power CO2 = CO2 by type = ETS sector power: max abs {d:.2g} '
        f'{"OK" if good else "FAIL"}')
    wbf = openpyxl.load_workbook(B.OUT)
    texts = [(w.title, c.coordinate, c.value) for w in wbf.worksheets for row in w.iter_rows() for c in row
             if isinstance(c.value, str) and not c.value.startswith('=')]
    refs = sorted({m for _t, _c, s in texts for m in DOC_REF.findall(s)})
    hist = {m for _t, _c, s in texts if s.startswith(('v0.', 'v1.')) for m in DOC_REF.findall(s)}
    vlog = {m for t, _c, s in texts if t == 'Settings' for m in DOC_REF.findall(s)}
    missing = [m for m in refs if m not in hist | vlog and not os.path.exists(os.path.join(HERE, m))
               and not os.path.exists(os.path.join(HERE, 'presentation', m))]
    ok &= not missing
    log(f'- file references in the workbook texts ({len(refs)}): current ones exist in the folder: '
        f'{"OK" if not missing else "FAIL " + str(missing)}')
    stale = [(t, c, p) for t, c, s in texts for p in STALE if p in s]
    ok &= not stale
    log(f'- stale statements about power, revenues or interim data in the workbook texts: '
        f'{"none OK" if not stale else "FAIL " + str(stale[:6])}')
    st = wbf['Settings']
    last_log = max((r for r in range(B.SET_LOG + 2, st.max_row + 1) if st.cell(r, 2).value), default=None)
    title_ok = f'v{B.VERSION}' in str(wbf['Mitigation']['A2'].value or '') or any(
        f'v{B.VERSION}' in str(c.value) for row in wbf['Mitigation'].iter_rows(max_row=2) for c in row)
    readme = open(os.path.join(HERE, 'README.md'), encoding='utf-8').read()
    good = (st.cell(last_log, 2).value == B.VERSION and title_ok and f'{B.NAME}-v{B.VERSION}.xlsm' in readme
            and f'build_v{B.VERSION.replace(".", "_")}.py' in readme)
    ok &= good
    log(f'- versions: Settings log ends with {st.cell(last_log, 2).value}, Mitigation title v{B.VERSION}, folder README '
        f'names the workbook and builder of v{B.VERSION} {"OK" if good else "FAIL"}')
    notes = [f for f in os.listdir(HERE) if f.endswith('.md') and DOC_REF.fullmatch(f)]
    stems = {}
    for f in notes:
        stems.setdefault(re.sub(r'_v\d.*', '', f), []).append(f)
    dup = {k: x for k, x in stems.items() if len(x) > 1}
    ok &= not dup
    log(f'- one current version of each method note in the folder ({len(notes)} notes): '
        f'{"OK" if not dup else "FAIL " + str(dup)}')
    return ok


def check_ecp():
    """v1.09: existing carbon pricing with non-zero test data for Egypt (rows added to the all-country table of
    sheet ECP), in the shipped scenarios and three test scenarios, against the Python recomputation; properties."""
    wb = openpyxl.load_workbook(B.OUT)
    we = wb['ECP']
    r = B.ECP_A0
    while we.cell(r, 1).value not in (None, ''):
        r += 1
    data = {}
    for code, v in ECP_TEST.items():
        vals = v if isinstance(v, dict) else {y: v for y in B.ECP_YEARS}
        for j, x in enumerate([f'EGY|{code}', 'EGY', code] + [vals[y] for y in B.ECP_YEARS], 1):
            we.cell(r, j).value = x
        data[code] = vals
        r += 1
    for g, (name, edits) in ECP_SCEN.items():
        add_scenario(wb, 2, g, name, edits)
    ws = recalc(variant(wb, 'expand'))
    ok = True
    for g in (1, 2, 3, 4, 5):
        mt = dict(scenario_mt(1 if g == 1 else 2))
        for k, v in ECP_SCEN.get(g, ('', {}))[1].items():
            mt[B.MT_NAMES[k]] = v
        ok &= compare(ws, g, recompute(mt, ecp=data), f'ECP test data, scenario {g} vs Python')
    yr = lambda g: dict(zip([ws.cell(B.R_YEAR, c).value for c in [B.group_cols(g)[1]] + B.group_cols(g)[2]],
                            [B.group_cols(g)[1]] + B.group_cols(g)[2]))
    c2, c4 = yr(2), yr(4)
    lh = B.HIST_YEARS[-1]
    d_hist = max(abs(ws.cell(r_, c2[y]).value or 0) for r_, *_ in B.var_rows('xcpd') for y in c2 if y <= lh)
    d_after = max(ws.cell(r_, c2[2030]).value for r_, *_ in B.var_rows('xcpd'))
    lvl = sum(ws.cell(r_, c2[2030]).value for r_, *_ in B.var_rows('xcp'))
    off = min(ws.cell(r_, c4[2030]).value for r_, *_ in B.var_rows('xcpd'))
    rx = [round(ws.cell(B.R_REV[k], c2[2030]).value, 1) for k in ('rxcp.all', 'rxcp.pow')]
    xk = max(ws.cell(B.R_PW[('xck', m)], c2[2030]).value for m, _n in B.PW_TECH)
    good = d_hist < TOL and d_after > 0 and lvl > 0 and off < 0 and rx[0] > 0 and rx[1] > 0 and xk > 0
    ok &= good
    log(f'- scenario 2 with the test data: change against the price data xcpd = 0 up to {lh} (max {d_hist:.2g}), '
        f'> 0 in 2030 (data 2025 above 2024: max {d_after:.3f} $/GJ); level xcp > 0; power xck 2030 max {xk:.5f} $/kWh; '
        f'existing carbon pricing revenue 2030 (all, power) {rx} USD m; scenario 4 (switched off): xcpd < 0 '
        f'(min {off:.3f}) {"OK" if good else "FAIL"}')
    c5 = yr(5)
    ctr = {s: ws.cell(B.R_POL[f'ctr.{s}'], c5[2030]).value for s in ('pow', 'cem', 'rod')}
    cp = ws.cell(B.R_CP, c5[2030]).value
    good = abs(ctr['pow']) < TOL and abs(ctr['cem'] - cp * 0.4) < 1e-9 and abs(ctr['rod'] - cp) < 1e-9
    ok &= good
    log(f'- scenario 5 (complementary): new carbon tax rate 2030 = carbon price {cp:.2f} x (1 - existing ETS sector '
        f'coverage): power {ctr["pow"]:.2f}, cement {ctr["cem"]:.2f}, road {ctr["rod"]:.2f} {"OK" if good else "FAIL"}')
    chk = max(abs(ws.cell(rr, c).value) for g in (1, 2, 3, 4, 5) for c in [B.group_cols(g)[1]] + B.group_cols(g)[2]
              for rr in (B.R_CHK, B.R_REV['rev.chk'], B.R_EMI['co2.chk']))
    ok &= chk < 1e-6
    log(f'- check rows ener.chk, rev.chk, co2.chk in groups 1-5: max abs {chk:.2g} {"OK" if chk < 1e-6 else "FAIL"}')
    return ok


def check_codes(wsf, wsv):
    ok = True
    described = [r for r in range(1, wsf.max_row + 1)
                 if isinstance(wsf[f'H{r}'].value, str) and wsf[f'H{r}'].value.startswith('=INDEX(VarLab')]
    for col in 'HIJ':
        ok &= uniform(wsf, [(r, CI(col)) for r in described], f'column {col}')[0]
    for g, *_ in B.SCENARIOS:
        code = B.group_cols(g)[0]
        ok &= uniform(wsf, [(r, code) for r in described], f'code column {L(code)} (scenario {g})')[0]
    k1, k2 = B.group_cols(1)[0], B.group_cols(2)[0]
    h = B.SUB_HEAD['rod']
    expected = [(h + B.VOFF['ener'] + 2, k1, 'egy.mit.ener.rod.gso.e.1', 'Fuel use | Road | Gasoline'),
                (B.R_PV0 + B.PVOFF['sp'], k1, 'egy.mit.sp.pow.coa.a.1', 'Supply cost (pre-tax price) | Power | Coal'),
                (B.R_PV0 + B.PVOFF['rpb'] + 6, k2, 'egy.mit.rpb.all.gso.a.2',
                 'Retail price before new policies | All subsectors | Gasoline'),
                (B.R_GP0, k1, 'egy.mit.gp.int.oil.1', 'International energy price (real; source and adjustment from '
                                                     'MTInputs) | International | Crude oil'),
                (B.R_CPI, k1, 'egy.mit.infl.1', None),
                (h + B.VOFF['ets'] + 1, k2, 'egy.mit.ets.rod.nga.a.2', 'New ETS cost in the fuel price (tax-equivalent, auctioned part) | Road | Natural gas'),
                (B.R_PV0 + B.PVOFF['esub'] + 7, k1, 'egy.mit.esub.all.die.a.1',
                 'Existing consumer subsidy (negative part of txo, positive number) | All subsectors | Diesel'),
                (B.R_POL['ets.p'], k2, 'egy.mit.ets.p.2', None),
                (h + B.VOFF['rnew'] + 2, k2, 'egy.mit.rnew.rod.gso.2', None),
                (B.R_REV['rsub.tra'], k1, 'egy.mit.rsub.tra.all.1', None),
                (B.R_REV['rtot.chg'], k2, 'egy.mit.rtot.chg.all.all.2', None),
                (h + B.VOFF['co2'] + 3, k1, 'egy.mit.co2.rod.die.e.1', 'CO2 emissions from fuel combustion | Road | Diesel'),
                (B.R_EMI['co2.f.nga'], k2, 'egy.mit.co2.all.nga.e.2', None),
                (B.R_EMI['co2.pct'], k2, 'egy.mit.co2.pct.all.all.2', None),
                (h + B.VOFF['atp'] + 2, k2, 'egy.mit.atp.rod.gso.e.2', 'After-tax price | Road | Gasoline'),
                (h + B.VOFF['ctxnew'] + 3, k2, 'egy.mit.ctxnew.rod.die.a.2', 'New carbon tax | Road | Diesel'),
                (h + B.VOFF['shp'] + 2, k1, 'egy.mit.shp.rod.gso.1', None),
                (h, k1, 'egy.mit.ener.rod.all.e.1', 'Fuel use | Road | All fuels'),
                (B.SEC_SUM['tra'], k1, 'egy.mit.ener.tra.all.e.1', 'Fuel use | Transport | All fuels'),
                (B.R_CP, k2, 'egy.mit.cptraj.2', 'Carbon price trajectory used'),
                (B.R_POL['sc.oen'], k1, 'egy.mit.ctcov.oen.all.1', None),
                (B.R_TOTAL, k2, 'egy.mit.ener.all.all.e.2', 'Fuel use | All subsectors | All fuels'),
                (B.R_PW[('gnc', 'nga')], k1, 'egy.mit.gnc.pow.nga.1',
                 'Generation cost before new policies (cost recovery) | Power | Natural gas'),
                (B.R_PW[('caxav', 'sol')], k2, 'egy.mit.caxav.pow.sol.2', None),
                (B.R_PW[('gncav', '')], k1, 'egy.mit.gncav.1', 'Average generation cost (weighted by generation shares)'),
                (B.R_PW[('rp', 'res')], k2, 'egy.mit.rp.res.ecy.2', 'End-user electricity price | Residential | Electricity'),
                (B.R_PW[('msc', '')], k1, 'egy.mit.msc.1', None),
                (B.R_PW[('ccp', 'oop')], k2, 'egy.mit.ccp.pow.oop.2', None),
                (B.R_PW[('gen', 'sol')], k1, 'egy.mit.gen.pow.sol.1', 'Generation | Power | Solar'),
                (B.R_PW[('dem', 'srv')], k2, 'egy.mit.dem.srv.ecy.2', None),
                (B.R_PW[('mult', 'tra')], k1, 'egy.mit.mult.tra.ecy.1', None),
                (B.R_EMI['co2.pow'], k1, 'egy.mit.co2.pow.all.e.1', 'CO2 emissions from fuel combustion | Power | All fuels')]
    for r, c, code, label in expected:
        got, lab = wsv.cell(r, c).value, wsv[f'H{r}'].value
        good = got == code and (label is None or lab == label)
        ok &= good
        log(f'- {L(c)}{r}: `{got}` | {lab} {"OK" if good else "FAIL (expected " + code + ")"}')
    return ok


def check_outline(wsf):
    ok = wsf.sheet_properties.outlinePr.summaryBelow is True
    log(f'- row group buttons below the groups (summaryBelow): {"OK" if ok else "FAIL"}')
    lv = lambda r: (wsf.row_dimensions[r].outlineLevel or 0, bool(wsf.row_dimensions[r].hidden))
    checks = [('carbon price (summary of section 1)', [B.R_CP], (0, False)),
              ('section 1 inputs and paths', [B.R_POL[k] for k, *_ in B.POL if k != 'cptraj'], (1, True)),
              ('section 2 retail prices before new policies', [r for r, *_ in rows_of('rpb')], (0, False)),
              ('section 2 selectors, gp, sp, txo, vat, etx, esub, esubpu', [B.R_PRI[p] for p in B.PRI_PARAMS]
               + [r for r, *_ in gp_rows()] + [B.R_PRI['gp.src'], B.R_PRI['gp.adj']]
               + [B.R_GPN0 + k for k in range(len(B.GP_ROWS))]
               + [r for v in B.PVARS if v != 'rpb' for r, *_ in rows_of(v)], (1, True)),
              ('sector totals', list(B.SEC_SUM.values()), (0, False)),
              ('subsector headings and electricity memo lines', list(B.SUB_HEAD.values())
               + [h + 1 for h in B.SUB_HEAD.values()], (1, True)),
              ('subsector variables', [r for v in B.VARS for r, *_ in B.var_rows(v)], (2, True)),
              ('power: average generation cost and end-user prices', [r for r, v, *_ in B.PW_ROWS
                                                                      if v in B.PW_SUMMARY], (0, False)),
              ('power: sub-headings and other rows', list(B.PW_HEAD.values()) + [r for r, v, *_ in B.PW_ROWS
                                                                                 if v not in B.PW_SUMMARY], (1, True)),
              ('results and electricity memo by sector', list(range(B.R_SUB0, B.R_PCT + 1))
               + list(range(B.R_ELD['all'] - 1, B.R_ELD[B.ELD_GROUPS[-1]] + 1)), (1, True)),
              ('revenue totals', [B.R_REV[k] for k, *_r, sec in B.REV_ROWS if sec == 'all'], (0, False)),
              ('revenue by sector', [B.R_REV[k] for k, *_r, sec in B.REV_ROWS if sec != 'all'], (1, True)),
              ('CO2 total and change', [B.R_EMI[k] for k in ('co2.all', 'co2.chg', 'co2.pct')], (0, False)),
              ('CO2 by sector, fuel, scenario 1', [B.R_EMI[k] for k, *_r in B.EMI_ROWS
                                                   if k not in ('co2.all', 'co2.chg', 'co2.pct')], (1, True))]
    lvl = {r: wsf.row_dimensions[r].outlineLevel or 0 for r in range(1, wsf.max_row + 2)}
    rolled = all(wsf.row_dimensions[r].hidden for r in lvl if lvl[r])
    flags = all(bool(wsf.row_dimensions[r + 1].collapsed) == (lvl[r + 1] < lvl[r])
                for r in range(1, wsf.max_row) if lvl[r])
    ok &= rolled and flags
    log(f'- opens rolled up: every grouped row hidden; collapsed flag on the row after each group '
        f'{"OK" if rolled and flags else "FAIL"}')
    wbf = openpyxl.load_workbook(B.OUT)
    first = wbf.sheetnames[0] == 'Mitigation' and wbf.active.title == 'Mitigation'
    zoom = wbf['Mitigation'].sheet_view.zoomScale == 75
    ok &= first and zoom
    log(f'- Mitigation is the first and active tab: {"OK" if first else "FAIL"}; zoom 75%: {"OK" if zoom else "FAIL"}')
    for label, rows, exp in checks:
        good = all(lv(r) == exp for r in rows)
        ok &= good
        log(f'- {label}: {len(rows)} rows at level {exp[0]}, {"hidden" if exp[1] else "visible"} '
            f'{"OK" if good else "FAIL"}')
    return ok


def check_mtinputs():
    tpl = openpyxl.load_workbook(B.MT_TEMPLATE)['MTInputs']
    ws = openpyxl.load_workbook(B.OUT)['MTInputs']
    diffs = [f'{L(c)}{r}' for r in range(1, tpl.max_row + 1) for c in range(1, 9)
             if tpl.cell(r, c).value != ws.cell(r, c).value]
    ok = not diffs and ws.column_dimensions['D'].hidden and ws.column_dimensions['E'].hidden
    log(f'- columns A:H, rows 1-{tpl.max_row}: {len(diffs)} cells differ from the template; D:E hidden '
        f'{"OK" if ok else "FAIL " + str(diffs[:10])}')
    return ok


def check_format(wsf):
    bands = [3, B.B_POL, B.B_PRI, B.B_POW, B.B_RES, B.B_REV, B.B_EMI] + list(B.SEC_BAND.values())
    ok = all(str(wsf[f'A{r}'].font.color.rgb).endswith('FFFFFF') and wsf[f'A{r}'].font.b for r in bands)
    texts = [wsf[f'A{r}'].value[:22] for r in sorted(bands)]
    log(f'- section bands {texts}: white bold text {"OK" if ok else "FAIL"}')
    ip = openpyxl.load_workbook(B.OUT)['Inputs_prices']
    marked = sorted({ip.cell(B.IP_HDR, c.column).value[:40] for row in ip.iter_rows(min_row=B.IP_R0, max_row=B.IP_R1)
                     for c in row if c.fill.start_color.rgb.endswith('FFFF00')})
    oop = B.IP_R0 + [pc for pc, _l in B.PRICES].index('oop.all')
    good = (all(ip[f'{B.IPC[k]}{r}'].fill.start_color.rgb.endswith('FFFF00') for k in ('vatc', 'vrx')
                for r in range(B.IP_R0, B.IP_R1 + 1))
            and all(ip[f'{B.IPC[k]}{oop}'].fill.start_color.rgb.endswith('FFFF00') for k in ('pccr', 'mar')))
    ok &= good
    pdm = openpyxl.load_workbook(B.OUT)['Prices_dom']
    with open(os.path.join(HERE, 'data', 'prices_dom_changes.csv'), encoding='utf-8') as f:
        ch = list(csv.DictReader(f))
    hdr = {pdm.cell(5, j).value: j for j in range(1, pdm.max_column + 1)}
    rowi = {pdm.cell(i, 1).value: i for i in range(6, pdm.max_row + 1)}
    n_marked = sum(1 for c in ch if pdm.cell(rowi[c['country_year']], hdr[c['column']]).fill.start_color.rgb.endswith('FFFF00'))
    vals = all(abs(float(pdm.cell(rowi[c['country_year']], hdr[c['column']]).value or 0) - float(c['new'] or 0)) < 1e-9
               for c in ch)
    unmarked_other = sum(1 for i in range(6, pdm.max_row + 1) for j in range(6, pdm.max_column + 1)
                         if not str(pdm.cell(i, 1).value).startswith('EGY')
                         and pdm.cell(i, j).fill.start_color.rgb.endswith('FFFF00'))
    good = n_marked == len(ch) and vals and unmarked_other == 0
    ok &= good
    log(f'- Prices_dom: {n_marked} of {len(ch)} changed Egypt cells bright yellow with the new values; no other '
        f'country marked {"OK" if good else "FAIL"}')
    ld = openpyxl.load_workbook(B.OUT)['LegacyDiff']
    n_ld = sum(1 for r in range(5, ld.max_row + 1) if ld.cell(r, 2).value)
    names = openpyxl.load_workbook(B.OUT).sheetnames
    good_ld = n_ld == len(B.LEGACY_DIFFS) >= 20 and names[:8] == ['Mitigation', 'MTOutputs', 'StoredResults',
                                                                   'ScenarioCompare', 'MacroCheck', 'Charts', 'ReadMe',
                                                                   'LegacyDiff']
    ok &= good_ld
    log(f'- tabs {names[:8]}; sheet LegacyDiff: {n_ld} differences listed {"OK" if good_ld else "FAIL"}')
    log(f'- Inputs_prices: data changed by assumption marked bright yellow (columns {marked}; other oil products '
        f'pass-through and margin) {"OK" if good else "FAIL"}')
    return ok


def check_sanity(ws):
    ok = True
    (_, b1, y1), (_, b2, y2) = B.group_cols(1), B.group_cols(2)
    cols1, cols2, years = [b1] + y1, [b2] + y2, [B.BASE_YEAR] + B.YEARS
    pre = max(abs(ws.cell(r, x).value - ws.cell(r, y).value) for r, *_ in B.var_rows('ener')
              for x, y, yr in zip(cols1, cols2, years) if yr < 2027)
    ok &= pre < TOL
    log(f'- scenario 1 vs 2 before 2027 (fuel use): max abs diff {pre:.3g} {"OK" if pre < TOL else "FAIL"}')
    oen = max(abs(ws.cell(r, x).value - ws.cell(r, y).value) for r, s, f in B.var_rows('ener') if s == 'oen'
              for x, y in zip(cols1, cols2))
    ok &= oen < TOL
    log(f'- other energy use not taxed (MCovOen FALSE): fuel use scenario 1 = 2, max abs diff {oen:.3g} '
        f'{"OK" if oen < TOL else "FAIL"}')
    chk = max(abs(ws.cell(B.R_CHK, c).value) for c in cols1 + cols2)
    secs = max(abs(sum(ws.cell(B.SEC_SUM[g], c).value for g in B.SEC_SUM) + ws.cell(B.R_EPOW, c).value
                   - ws.cell(B.R_TOTAL, c).value) for c in cols1 + cols2)
    ok &= chk < 1e-6 and secs < 1e-6
    log(f'- check row {B.R_CHK}: max abs {chk:.3g}; sector totals and power sum to total: max abs {secs:.3g} '
        f'{"OK" if chk < 1e-6 and secs < 1e-6 else "FAIL"}')
    ip = openpyxl.load_workbook(recalc_path(B.OUT), data_only=True)['Inputs_prices']
    pcc = {ip.cell(r, CI(B.IPC['key'])).value: ip.cell(r, CI(B.IPC['pcc'])).value for r in range(B.IP_R0, B.IP_R1 + 1)}
    flat = [pc for pc, _l in B.PRICES if pcc[pc] == 0]
    i_l = years.index(B.HIST_YEARS[-1])
    rows = {pc: r for (r, *_), (pc, _l) in zip(rows_of('rpb'), B.PRICES)}
    dflat = max(abs(ws.cell(rows[pc], c).value - ws.cell(rows[pc], cols1[i_l]).value) for pc in flat
                for c in cols1[i_l:-1])
    ok &= dflat < 1e-9
    log(f'- pass-through 0 ({", ".join(flat)}): retail price before new policies stays at its {B.HIST_YEARS[-1]} '
        f'value, max abs diff {dflat:.3g} {"OK" if dflat < 1e-9 else "FAIL"}')
    log('')
    log('Retail price before new policies, scenario 1 (real $/GJ of ResultsYear), and chosen pass-through:')
    log('')
    log('| Price fuel | pass-through | 2022 | 2024 | 2027 | 2030 | 2040 | VAT 2024 | existing tax 2024 | '
        'existing subsidy 2024 | subsidy per price unit 2024 |')
    log('|---|---|---|---|---|---|---|---|---|---|---|')
    i24 = cols1[years.index(2024)]
    for pc, _l in B.PRICES:
        vals = [ws.cell(rows[pc], cols1[years.index(y)]).value for y in (2022, 2024, 2027, 2030, 2040)]
        k = [pc2 for pc2, _l2 in B.PRICES].index(pc)
        ext = [ws.cell(B.R_PV0 + B.PVOFF[v] + k, i24).value for v in ('vat', 'etx', 'esub', 'esubpu')]
        log(f'| {pc} | {pcc[pc]:g} | ' + ' | '.join(f'{v:.3f}' for v in vals + ext) + ' |')
    rsum = max(abs(sum(ws.cell(B.R_REV[f'{v}.{g}'], c).value for g in [g_ for _n, g_, _t in B.SECTIONS] + ['pow'])
                   - ws.cell(B.R_REV[f'{v}.all'], c).value) for v in B.REV_VARS for c in cols1 + cols2)
    zero_new = max(abs(ws.cell(B.R_REV['rnew.all'], c).value) for c in cols1)
    good = rsum < 1e-6 and zero_new < TOL
    ok &= good
    log(f'- revenues: sectors and power sum to totals (max abs {rsum:.3g}); no new-policy revenue in scenario 1 '
        f'(max abs {zero_new:.3g}) {"OK" if good else "FAIL"}')
    esum = max(max(abs(sum(ws.cell(B.R_EMI[f'co2.{g}'], c).value for _n, g, _t in B.SECTIONS)
                       + ws.cell(B.R_EMI['co2.pow'], c).value - ws.cell(B.R_EMI['co2.all'], c).value),
                   abs(sum(ws.cell(B.R_EMI[f'co2.f.{f}'], c).value for f, *_x in B.FUELS)
                       - ws.cell(B.R_EMI['co2.all'], c).value)) for c in cols1 + cols2)
    bio = max(abs(ws.cell(B.R_EMI['co2.f.bio'], c).value) for c in cols1 + cols2)
    pre = max(abs(ws.cell(B.R_EMI['co2.chg'], c).value - ws.cell(B.R_EMI['co2.pow'], c).value
                  + ws.cell(B.R_EMI['co2.pow'], c1).value) for c, c1, yr in zip(cols2, cols1, years) if yr < 2027)
    pre_pow = max(abs(ws.cell(B.R_EMI['co2.pow'], c).value - ws.cell(B.R_EMI['co2.pow'], c1).value)
                  for c, c1, yr in zip(cols2, cols1, years) if yr < 2027)
    good = esum < 1e-9 and bio < TOL and pre < TOL
    ok &= good
    log(f'- CO2: sectors incl. power and fuels sum to the total (max abs {esum:.3g}); biomass 0; no change vs scenario 1 '
        f'before 2027 outside power (max abs {pre:.3g}) {"OK" if good else "FAIL"}')
    log(f'  - power CO2 before 2027 differs by up to {pre_pow:.3f} MtCO2: investment looks ahead (levelised variable cost '
        f'includes the future carbon price), as legacy')
    log('')
    log('CO2 emissions from fuel combustion incl. power generation (MtCO2; no process emissions):')
    log('')
    log('| Year | Scenario 1 | Scenario 2 | Change | Transport | Buildings | Industry | Other | Power (scen. 2) |')
    log('|---|---|---|---|---|---|---|---|---|')
    for x, y, yr in zip(cols1, cols2, years):
        if yr in (2022, 2024, 2027, 2030, 2035, 2040):
            a, b = ws.cell(B.R_EMI['co2.all'], x).value, ws.cell(B.R_EMI['co2.all'], y).value
            secs = [ws.cell(B.R_EMI[f'co2.{g}'], y).value for _n, g, _t in B.SECTIONS] + [ws.cell(B.R_EMI['co2.pow'], y).value]
            log(f'| {yr} | {a:,.1f} | {b:,.1f} | {b / a - 1:+.1%} | ' + ' | '.join(f'{v:,.1f}' for v in secs) + ' |')
    log('')
    log('Revenues (USD million, real 2026):')
    log('')
    log('| Year | existing taxes | existing subsidies | new policies | total | change vs scenario 1 |')
    log('|---|---|---|---|---|---|')
    for g, cols in ((1, cols1), (2, cols2)):
        for x, yr in zip(cols, years):
            if yr in (2022, 2027, 2030, 2040):
                vals = [ws.cell(B.R_REV[k], x).value for k in ('rtx.all', 'rsub.all', 'rnew.all', 'rtot.all',
                                                               'rtot.chg')]
                log(f'| {yr} (scenario {g}) | ' + ' | '.join(f'{v:,.0f}' for v in vals) + ' |')
    log('')
    log('| Year | Scenario 1 (ktoe) | Scenario 2 (ktoe) | Change |')
    log('|---|---|---|---|')
    for x, y, yr in zip(cols1, cols2, years):
        if yr in (2022, 2023, 2024, 2026, 2027, 2030, 2035, 2040):
            a, b = ws.cell(B.R_TOTAL, x).value, ws.cell(B.R_TOTAL, y).value
            log(f'| {yr} | {a:,.0f} | {b:,.0f} | {b / a - 1:+.1%} |')
    return ok


def check_power_results(ws):
    """Power block: properties that must hold, and the key numbers of scenarios 1 and 2."""
    (_, b1, y1), (_, b2, y2) = B.group_cols(1), B.group_cols(2)
    cols1, cols2, years = [b1] + y1, [b2] + y2, [B.BASE_YEAR] + B.YEARS
    row = lambda v, m='': B.R_PW[(v, m)]
    val = lambda v, m, c: ws.cell(row(v, m), c).value
    tech = [f for f, _n in B.PW_TECH]
    ok = True
    ccp1 = max(abs(ws.cell(r, c).value) for r, v, *_ in B.PW_ROWS if v in ('ccp', 'ccpav', 'rcarb') for c in cols1)
    vbc = max(abs(val('vbc', f, x) - val('vbc', f, y)) for f in tech for x, y in zip(cols1, cols2))
    shares = max(abs(sum(val('shr', f, c) for f in tech) - 1) for c in cols1 + cols2)
    good = ccp1 < TOL and vbc < TOL and shares < 1e-9
    ok &= good
    log(f'- scenario 1: no new policy cost in power (max abs {ccp1:.3g}); variable costs before policy equal in '
        f'scenarios 1 and 2 (max abs {vbc:.3g}); generation shares sum to 1 (max abs {shares:.3g}) '
        f'{"OK" if good else "FAIL"}')
    meet = max(abs(val('gtot', '', c) - val('req', '', c)) for c in cols1[1:] + cols2[1:])
    cap = max(abs(val('cap', f, c) - (val('cap', f, c - 1) - val('ret', f, c - 1) + val('nim', f, c - 1)))
              for f in tech for c in cols1[1:] + cols2[1:])
    over = max(sum(val('nei', f, c) for f in tech) - val('need', '', c) for c in cols1 + cols2)
    resid = [(B.year_of(c), val('need', '', c) - sum(val('nei', f, c) for f in tech)) for c in cols1]
    neg = min(val('nei', f, c) for f in tech for c in cols1 + cols2)
    good = meet < 1e-6 and cap < 1e-6 and over < 1e-6 and neg > -1e-9
    ok &= good
    log(f'- generation = required generation from {B.YEARS[0]} (max abs {meet:.3g} MWy); capacity = last year less '
        f'retirements plus new nameplate investment (max abs {cap:.3g} MW); new effective investment >= 0 and at most the '
        f'capacity needed (max excess {over:.3g} MW) {"OK" if good else "FAIL"}')
    log('  - scale-up limits leave investment unplaced (MW, scenario 1; gas then runs at a higher capacity factor, as '
        'legacy): ' + (', '.join(f'{y} {x:,.0f}' for y, x in resid if x > 1e-6) or 'none'))
    lag = max(abs(val('sc', g, c) - val('gncav', '', c - 1) - (val('sc', g, cols1[1]) - val('gncav', '', cols1[0])))
              for g, _l in B.PW_GROUPS for c in cols1[1:])
    ok &= lag < TOL
    log(f'- supply cost = average generation cost of the previous year + constant T&D add-on (max abs {lag:.3g}) '
        f'{"OK" if lag < TOL else "FAIL"}')
    c40_1, c40_2 = cols1[-1], cols2[-1]
    good = val('vre', '', c40_2) >= val('vre', '', c40_1) and \
        ws.cell(B.R_EMI['co2.pow'], c40_2).value <= ws.cell(B.R_EMI['co2.pow'], c40_1).value
    ok &= good
    log(f'- carbon price (scenario 2) shifts power towards renewables by 2040: VRE share {val("vre", "", c40_1):.3f} -> '
        f'{val("vre", "", c40_2):.3f}, power CO2 {ws.cell(B.R_EMI["co2.pow"], c40_1).value:.2f} -> '
        f'{ws.cell(B.R_EMI["co2.pow"], c40_2).value:.2f} MtCO2 {"OK" if good else "FAIL"}')
    log('')
    log('Power, scenario 1 (baseline) and 2 (carbon price $20 from 2027):')
    log('')
    log('| Year | generation TWh (1) | (2) | gas share (1) | (2) | VRE share (1) | (2) | power CO2 Mt (1) | (2) | '
        'price res $/kWh (1) | (2) | capacity GW (1) | investment USD m (1) |')
    log('|---|---|---|---|---|---|---|---|---|---|---|---|---|')
    for k, yr in enumerate(years):
        if yr in (2022, 2023, 2025, 2027, 2030, 2035, 2040):
            c1, c2 = cols1[k], cols2[k]
            v = [val('gtot', '', c1) * 8.76 / 1000, val('gtot', '', c2) * 8.76 / 1000, val('shr', 'nga', c1),
                 val('shr', 'nga', c2), val('vre', '', c1), val('vre', '', c2), ws.cell(B.R_EMI['co2.pow'], c1).value,
                 ws.cell(B.R_EMI['co2.pow'], c2).value, val('rp', 'res', c1), val('rp', 'res', c2),
                 sum(val('cap', f, c1) for f in tech) / 1000, val('usdtot', '', c1)]
            fm = ['{:.1f}', '{:.1f}', '{:.3f}', '{:.3f}', '{:.3f}', '{:.3f}', '{:.1f}', '{:.1f}', '{:.4f}', '{:.4f}',
                  '{:.1f}', '{:,.0f}']
            log(f'| {yr} | ' + ' | '.join(f_.format(x) for f_, x in zip(fm, v)) + ' |')
    log('')
    log('Generation cost by type, scenario 1, 2022 and 2040 ($/kWh): ' + ', '.join(
        f'{n} {val("gnc", f, cols1[0]):.3f} / {val("gnc", f, cols1[-1]):.3f}' for f, n in B.PW_TECH))
    return ok


def check_formula_length():
    """v1.07: every cell formula of the shipped workbook has at most 100 characters (as Excel shows it, without the
    file prefixes _xlfn. / _xlpm.); the named LAMBDAs may be longer."""
    wb = openpyxl.load_workbook(B.OUT)
    n, worst, over, lam_calls = 0, (0, ''), [], 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith('='):
                    k = len(v.replace('_xlfn.', '').replace('_xlpm.', ''))
                    n += 1
                    lam_calls += bool(LAMBDA_CALL.search(v))
                    if k > worst[0]:
                        worst = (k, f'{ws.title}!{c.coordinate}')
                    if k > 100:
                        over.append(f'{ws.title}!{c.coordinate}')
    ok = not over
    names = [d for d in wb.defined_names if d not in B.LAMBDAS]
    longest = max((len(B.lambda_xml(k).replace('_xlfn.', '').replace('_xlpm.', '')), k) for k in B.LAMBDAS)
    log(f'- {n} cell formulas in {len(wb.worksheets)} sheets ({lam_calls} call a named LAMBDA): longest {worst[0]} '
        f'characters ({worst[1]}); over 100: {len(over)} {"OK" if ok else "FAIL " + str(over[:8])}')
    log(f'- {len(B.LAMBDAS)} named LAMBDAs (longest: {longest[1]}, {longest[0]} characters); {len(names)} defined '
        f'names for data tables and fixed ranges: {", ".join(names)}')
    return ok


PRET_TEST = {'nga': (2025, 1500.0), 'oop': (2024, 400.0)}   # planned retirement test: type -> (from year, MW/yr)


def check_planned_retirement():
    """Planned retirement switched on for gas and oil with data (PowerTech, PowerPaths edited as a user would):
    workbook = power_model_ref_v0_2 in scenarios 1 and 2; retirement = the data; capacity balance holds."""
    wb = openpyxl.load_workbook(B.OUT)
    pt, pp = wb['PowerTech'], wb['PowerPaths']
    hdr = {pt.cell(3, j).value: j for j in range(1, pt.max_column + 1)}
    years = [B.BASE_YEAR] + B.YEARS
    ycol = {pp.cell(3, j).value: j for j in range(1, pp.max_column + 1)}
    pflag, pret = [], []
    for f in POWER_TECH:
        on = f in PRET_TEST
        r = next(i for i in range(4, pt.max_row + 1) if pt.cell(i, 1).value == f)
        pt.cell(r, hdr['planned_retirement']).value = 'Yes' if on else 'No'
        rp = next(i for i in range(4, pp.max_row + 1) if pp.cell(i, 1).value == f'pret|{f}')
        series_ = [PRET_TEST[f][1] if on and y >= PRET_TEST[f][0] else 0.0 for y in years]
        for y, v in zip(years, series_):
            pp.cell(rp, ycol[y]).value = v
        pflag.append(1.0 if on else 0.0)
        pret.append(series_)
    ws = recalc(variant(wb, 'expand'))
    ok = True
    for g in (1, 2):
        ok &= compare(ws, g, recompute(scenario_mt(g), power={'pflag': pflag, 'pret': pret}),
                      f'planned retirement on (gas 1,500 MW/yr from 2025, oil 400 MW/yr from 2024), scenario {g} vs '
                      f'Python')
    _, b1, y1 = B.group_cols(1)
    cols = [b1] + y1
    ret = max(abs(ws.cell(B.R_PW[('ret', f)], c).value - min(PRET_TEST[f][1] if y >= PRET_TEST[f][0] else 0.0,
                                                               ws.cell(B.R_PW[('cap', f)], c).value))
              for f in PRET_TEST for c, y in zip(cols, years))
    good = ret < 1e-6
    ok &= good
    k30 = years.index(2030)
    log(f'- retirement of gas and oil = the planned data (capped at capacity), max abs diff {ret:.3g} MW; gas capacity '
        f'2030 {ws.cell(B.R_PW[("cap", "nga")], cols[k30]).value:,.0f} MW, new investment 2030 '
        f'{ws.cell(B.R_PW[("nimtot", "")], cols[k30]).value:,.0f} MW {"OK" if good else "FAIL"}')
    return ok


def check_memo(ws):
    """Electricity demand memo lines: each subsector line = power demand of the subsector (section 3 H); section 11
    by sector sums the subsector lines; the total = total power demand; fuel totals exclude them."""
    ok = True
    for g in (1, 2):
        _, b, ys = B.group_cols(g)
        cols = [b] + ys
        d_sub = max(abs(ws.cell(B.SUB_HEAD[s] + 1, c).value - (ws.cell(B.R_PW[('dem', s)], c).value
                                                                if s in B.DEM_SUBS else 0.0))
                    for s in B.SUB_HEAD for c in cols)
        d_tot = max(abs(ws.cell(B.R_ELD['all'], c).value - ws.cell(B.R_PW[('demtot', '')], c).value) for c in cols)
        d_grp = max(abs(sum(ws.cell(B.SUB_HEAD[s[0]] + 1, c).value for s in B.SUBSECTORS if s[2] == grp)
                        - ws.cell(B.R_ELD[grp], c).value) for grp in B.ELD_GROUPS[1:] for c in cols)
        d_dg = max(abs(sum(ws.cell(B.R_PW[('demg', m)], c).value for m, _l in B.DEM_GROUPS)
                       - ws.cell(B.R_PW[('demtot', '')], c).value) for c in cols)
        good = max(d_sub, d_tot, d_grp, d_dg) < 1e-6
        ok &= good
        log(f'- scenario {g}: memo by subsector = power demand of the subsector (max abs {d_sub:.3g}); by sector = sum of '
            f'its subsectors ({d_grp:.3g}); total = total power demand ({d_tot:.3g}); demand by group sums to the '
            f'total ({d_dg:.3g}) {"OK" if good else "FAIL"}')
    _, b1, y1 = B.group_cols(1)
    log('  - electricity demand by sector, scenario 1 (ktoe), 2022 / 2030 / 2040: ' + '; '.join(
        f'{grp} ' + ' / '.join(f'{ws.cell(B.R_ELD[grp], c).value:,.0f}' for c in (b1, y1[7], y1[-1]))
        for grp in B.ELD_GROUPS))
    return ok


def check_goalseek():
    """ets_goalseek_v0_4.py on a copy with the ETS test scenario as group 3: the cap is met within tolerance."""
    import ets_goalseek_v0_4 as G
    wb = openpyxl.load_workbook(B.OUT)
    add_scenario(wb, 2, 3, 'Test 3: new ETS with cap, goal seek', ETS_TEST)
    path = os.path.join(tempfile.mkdtemp(), 'goalseek.xlsx')
    wb.save(path)
    lines = []
    p, worst, it = G.solve(path, [3], tol=0.005, max_iter=20, log=lines.append)    # as the macro (ETS_MAXITER)
    for ln in lines:
        log(ln)
    ok = worst[3] < 0.005
    years = [B.BASE_YEAR] + B.YEARS
    est = recompute({**scenario_mt(2), **{B.MT_NAMES[k]: v for k, v in ETS_TEST.items()}})[('ets.est', '', '')]
    log(f'- goal seek, scenario 3 (ETS test settings): worst |covered/cap - 1| {worst[3]:.4f} after {it} iterations '
        f'{"OK" if ok else "FAIL"}')
    log('  - price $/tCO2, fast estimate -> goal seek: ' + ', '.join(
        f'{y} {est[k]:.1f} -> {p[3][k]:.1f}' for k, y in enumerate(years) if y in (2027, 2030, 2035, 2040)))
    return ok


# ------------------------------------------------------------------ v1.03: scenarios, stored results, macro
STEM_KEYS = {'mit.co2.all.all.e': ('co2', 'all', 'all'), 'mit.rtx.all.all': ('rtx', 'all', 'all'),
             'mit.rsub.all.all': ('rsub', 'all', 'all'), 'mit.rnew.all.all': ('rnew', 'all', 'all'),
             'mit.rtot.all.all': ('rtot', 'all', 'all'), 'mit.cptraj': ('cptraj', '', ''),
             'mit.ets.p': ('ets.p', '', ''), 'mit.co2.ets.all.all': ('co2.ets', 'all', 'all'),
             'mit.co2.cap.all.all': ('ets.cap', '', ''), 'mit.rp.res.ecy': ('rp', 'res', 'ecy'),
             'mit.rp.ind.ecy': ('rp', 'ind', 'ecy'), 'mit.gncav': ('gncav', '', ''),
             'mit.co2.pow.all.e': ('co2.pow', '', ''), 'mit.gtot': ('gtot', '', ''), 'mit.vre': ('vre', '', ''),
             'mit.shr.pow.nga': ('shr', 'pow', 'nga'), 'mit.usdtot': ('usdtot', '', '')}


def stored_blocks(ws):
    """{id: {stem: [values base..last]}} and {id: name} of StoredResults (values workbook or shipped file)."""
    out, names = {}, {}
    r = B.SR_FIRST
    while ws.cell(r, 3).value not in (None, ''):
        sid = ws.cell(r, 1).value
        out.setdefault(sid, {})[ws.cell(r, 3).value] = [ws.cell(r, B.SR_Y0 + j).value for j in range(B.NY)]
        names[sid] = ws.cell(r, 2).value
        r += 1
    return out, names


def definition_mt(num):
    mt = mt_used()
    edits = next(e for n, _nm, _r, e in B.DEFINITIONS if n == num)
    for k, v in edits.items():
        mt[k if isinstance(k, int) else B.MT_NAMES[k]] = v
    return mt


def rel(a, b):
    return abs(a - b) / max(1.0, abs(b)) if isinstance(a, (int, float)) and isinstance(b, (int, float)) else float('inf')


def check_definitions():
    ws = openpyxl.load_workbook(B.OUT)['MTInputs']
    ok = True
    for n, nm, run, edits in B.DEFINITIONS:
        c = B.MT_COL0 + n - 1
        good = (ws.cell(B.MT_ROW_SCEN, c).value == f'={L(c - 1)}{B.MT_ROW_SCEN}+1' and ws.cell(B.MT_ROW_NAME, c).value == nm
                and ws.cell(B.MT_ROW_RUN, c).value == run
                and all(ws.cell(k if isinstance(k, int) else B.MT_NAMES[k], c).value == v for k, v in edits.items()))
        ok &= good
        log(f'- definition {n} (column {L(c)}): "{nm}", Run? {run}, {len(edits)} inputs changed from the template '
            f'{"OK" if good else "FAIL"}')
    live = B.MT_COL0 + 1
    good = ws.cell(B.MT_ROW_RUN, B.MT_COL0).value.startswith('Baseline') and ws.cell(B.MT_ROW_RUN, live).value.startswith('Live')
    ok &= good
    log(f'- J4 / K4: "{ws.cell(B.MT_ROW_RUN, B.MT_COL0).value}" / "{ws.cell(B.MT_ROW_RUN, live).value}" {"OK" if good else "FAIL"}')
    return ok


def check_vba():
    from oletools.olevba import VBA_Parser
    vp = VBA_Parser(B.OUT)
    mods = {n.rsplit('.', 1)[0]: c for (_f, _s, n, c) in vp.extract_macros()}
    src = open(B.BAS, encoding='cp1252').read().replace('\r\n', '\n')
    got = mods.get('CPATScenarios', '').replace('\r\n', '\n')
    wb = openpyxl.load_workbook(B.OUT)
    codes = [w.sheet_properties.codeName for w in wb.worksheets]
    good = (got.strip() == src.strip() and set(codes) | {'ThisWorkbook'} == set(mods) - {'CPATScenarios'}
            and wb.code_name == 'ThisWorkbook')
    log(f'- VBA project: module CPATScenarios = {os.path.basename(B.BAS)} ({len(src.splitlines())} lines); document '
        f'modules ThisWorkbook + {len(codes)} sheets matching the sheet code names {"OK" if good else "FAIL"}')
    procs = re.findall(r'^Public (?:Sub|Function) (\w+)', src, flags=re.M)
    log(f'  - public macros: {", ".join(procs)}')
    return good


def check_stored(ws_vals):
    """Shipped StoredResults vs the shipped workbook recalculated (baseline, definition 3 = live scenario 2) and vs the
    independent Python recomputation (definitions 4 and 5; 5 with its stored ETS prices)."""
    wb = openpyxl.load_workbook(B.OUT, data_only=True)
    st, names = stored_blocks(wb['StoredResults'])
    ok = sorted(st) == [1] + [n for n, _nm, run, _e in B.DEFINITIONS if run == 'Yes'] and \
        all(len(v) == len(B.MTO_INDICATORS) for v in st.values())
    log(f'- StoredResults: scenario IDs {sorted(st)}, {len(B.MTO_INDICATORS)} rows each; values only: '
        f'{"OK" if ok else "FAIL"}')
    mo = ws_vals['MTOutputs']
    for sid, src in ((1, 1), (3, 2)):
        h = B.MTO_R0 + (src - 1) * B.MTO_BLOCK
        d = max(rel(st[sid][stem][j], mo.cell(h + 1 + k, B.MTO_C0 + j).value)
                for k, (stem, _f) in enumerate(B.MTO_INDICATORS) for j in range(B.NY)
                if isinstance(mo.cell(h + 1 + k, B.MTO_C0 + j).value, (int, float)))
        ok &= d < 1e-9
        log(f'- stored ID {sid} = MTOutputs block {src} of the shipped workbook recalculated (same inputs): max rel diff '
            f'{d:.2e} {"OK" if d < 1e-9 else "FAIL"}')
    years = [B.BASE_YEAR] + B.YEARS
    for sid in (4, 5):
        mt = definition_mt(sid)
        ovr = None
        if sid == 5:          # the stored prices as override (the goal seek switches the override on)
            ovr = {y: st[sid]['mit.ets.p'][k] for k, y in enumerate(years)}
            mt[B.MT_NAMES['D_ETSPriceOverride']] = 'Yes'
        exp = recompute(mt, ovr=ovr)
        d = max(rel(st[sid][stem][k], exp[key][k]) for stem, key in STEM_KEYS.items() for k in range(len(years)))
        ok &= d < 1e-9
        log(f'- stored ID {sid} ({names[sid][:40]}...) vs independent Python recomputation ({len(STEM_KEYS)} '
            f'indicators: CO2, revenues, carbon and ETS price, covered CO2, cap, electricity prices, generation cost): max rel diff {d:.2e} '
            f'{"OK" if d < 1e-9 else "FAIL"}')
    cov, cap = st[5]['mit.co2.ets.all.all'], st[5]['mit.co2.cap.all.all']
    gap = max(abs(a / b - 1) for a, b in zip(cov, cap) if b)
    ok &= gap < 0.005
    log(f'- stored ID 5 (ETS package): covered emissions meet the cap, worst |covered/cap - 1| {gap:.4f} '
        f'{"OK" if gap < 0.005 else "FAIL"}; ETS price 2027/2030/2035/2040: '
        f'{[round(st[5]["mit.ets.p"][y - B.BASE_YEAR], 1) for y in (2027, 2030, 2035, 2040)]}')
    return ok


def check_compare():
    wbf = openpyxl.load_workbook(B.OUT)
    ws = wbf['ScenarioCompare']
    vals = recalc_wb(B.OUT)
    wv = vals['ScenarioCompare']
    st, names = stored_blocks(vals['StoredResults'])
    y = wv['C4'].value
    n = len(B.CMP_STEMS)
    r_l, r_d, r_p = 9, 9 + n + 2, 9 + 2 * (n + 2)
    cols = range(4, 5 + B.CMP_SLOTS)
    ok = True
    worst = 0.0
    for c in cols:
        sid = wv.cell(6, c).value
        for k, (stem, _f) in enumerate(B.CMP_STEMS):
            v = wv.cell(r_l + k, c).value
            if sid in (None, ''):
                ok &= v in (None, '')
                continue
            e = st[sid][stem][y - B.BASE_YEAR]
            worst = max(worst, rel(v, e))
            if c > 4:
                worst = max(worst, rel(wv.cell(r_d + k, c).value, e - st[1][stem][y - B.BASE_YEAR]))
                b = st[1][stem][y - B.BASE_YEAR]
                if b:
                    worst = max(worst, rel(wv.cell(r_p + k, c).value, e / b - 1))
    ok &= worst < 1e-9
    log(f'- ScenarioCompare {y}: IDs {[wv.cell(6, c).value for c in cols if wv.cell(6, c).value]}, names '
        f'{[str(wv.cell(7, c).value)[:18] for c in cols if wv.cell(6, c).value]}; levels, differences and % differences '
        f'equal StoredResults (max rel diff {worst:.2e}) {"OK" if worst < 1e-9 else "FAIL"}')
    for label, r0, cs in (('levels', r_l, list(cols)), ('differences', r_d, list(cols)[1:]),
                          ('% differences', r_p, list(cols)[1:])):
        ok &= uniform(ws, [(r0 + k, c) for k in range(n) for c in cs], f'ScenarioCompare {label}')[0]
    ok &= uniform(ws, [(7, c) for c in cols], 'ScenarioCompare names')[0]
    i = list(B.CMP_STEMS).index(('mit.co2.all.all.e', '#,##0.0'))
    log(f'  - CO2 {y} (MtCO2): ' + ', '.join(f'{wv.cell(6, c).value}: {wv.cell(r_l + i, c).value:,.1f}'
                                            for c in cols if wv.cell(6, c).value))
    return ok


def check_macro():
    """Run CPATScenarios.RunAll in LibreOffice on the shipped workbook (LAMBDAs expanded, StoredResults emptied) and
    compare what the macro stores with the shipped StoredResults (Python emulation); the live column and the ETS
    override row must be restored."""
    from vba_project_v0_1 import embed_vba
    wb = openpyxl.load_workbook(B.OUT)
    sr = wb['StoredResults']
    for row in sr.iter_rows(min_row=B.SR_FIRST, max_row=sr.max_row):
        for c in row:
            c.value = None
    xlsx = variant(wb, 'expand')
    xlsm = xlsx[:-5] + '.xlsm'
    embed_vba(xlsx, xlsm, {'CPATScenarios': B.vba_module_source()})
    out = xlsx[:-5] + '_ran.xlsx'
    res = subprocess.run(['python3', os.path.join(HERE, 'lo_run_macro_v0_1.py'), xlsm, 'CPATScenarios.RunAll', out],
                         capture_output=True, text=True, timeout=3000)
    log(f'- LibreOffice ran the embedded macro: {res.stdout.strip().splitlines()[-1] if res.stdout.strip() else res.stderr[-300:]}')
    if not os.path.exists(out):
        log('  FAIL: no output')
        return False
    ran = openpyxl.load_workbook(out, data_only=True)
    got, gnames = stored_blocks(ran['StoredResults'])
    ship, snames = stored_blocks(openpyxl.load_workbook(B.OUT, data_only=True)['StoredResults'])
    worst, where = 0.0, None
    for sid in ship:
        for stem in ship[sid]:
            for j in range(B.NY):
                d = rel(got.get(sid, {}).get(stem, [None] * B.NY)[j], ship[sid][stem][j])
                if d > worst:
                    worst, where = d, (sid, stem, B.BASE_YEAR + j)
    good = sorted(got) == sorted(ship) and gnames == snames and worst < 1e-6
    log(f'  - macro stored IDs {sorted(got)} with the same names; values vs the Python emulation: max rel diff '
        f'{worst:.2e}{" at " + str(where) if where else ""} {"OK" if good else "FAIL"}')
    orig = openpyxl.load_workbook(B.OUT)
    live = B.MT_COL0 + 1
    mt_o, mt_r = orig['MTInputs'], openpyxl.load_workbook(out)['MTInputs']
    lo_bool = {'=TRUE()': True, '=FALSE()': False}         # LibreOffice saves boolean cells as =TRUE() / =FALSE()
    norm = lambda v: lo_bool.get(v, v) if isinstance(v, str) else v
    diffs = [r for r in range(1, 416) if r != B.MT_ROW_SCEN and mt_o.cell(r, live).value != norm(mt_r.cell(r, live).value)
             and not (isinstance(mt_o.cell(r, live).value, str) and mt_o.cell(r, live).value.startswith('='))]
    ovr = [ran['Mitigation'].cell(B.R_POL['ets.ovr'], c).value for c in [B.group_cols(2)[1]] + B.group_cols(2)[2]]
    good2 = not diffs and all(v in (None, '') for v in ovr)
    log(f'  - live column K restored ({len(diffs)} cells differ) and ETS override row empty again '
        f'{"OK" if good2 else "FAIL " + str(diffs[:5])}')
    return good and good2


def run_macro_lo(wb, proc, tag):
    """Expanded variant of wb with the VBA module embedded, run proc in LibreOffice; (stdout line, values workbook)."""
    from vba_project_v0_1 import embed_vba
    xlsx = variant(wb, 'expand')
    xlsm = xlsx[:-5] + f'_{tag}.xlsm'
    embed_vba(xlsx, xlsm, {'CPATScenarios': B.vba_module_source()})
    out = xlsx[:-5] + f'_{tag}_ran.xlsx'
    res = subprocess.run(['python3', os.path.join(HERE, 'lo_run_macro_v0_1.py'), xlsm, f'CPATScenarios.{proc}', out],
                         capture_output=True, text=True, timeout=3000)
    line = res.stdout.strip().splitlines()[-1] if res.stdout.strip() else res.stderr[-300:]
    return line, (openpyxl.load_workbook(out, data_only=True) if os.path.exists(out) else None)


def check_first_run():
    """v1.04: the first-run test macro CheckBatch, run in LibreOffice on the shipped workbook (stored results kept):
    Result PASS and the largest relative difference below 1e-6; and on a copy where a row code is broken: FAIL
    before any run."""
    ok = True
    ws = openpyxl.load_workbook(B.OUT)['MacroCheck']
    good = ws[f'A{B.CHK_FIRST}'].value == 'Result' and ws[f'B{B.CHK_FIRST}'].value == 'not run yet'
    ok &= good
    log(f'- sheet MacroCheck as shipped: "{ws[f"B{B.CHK_FIRST}"].value}" {"OK" if good else "FAIL"}')
    line, ran = run_macro_lo(openpyxl.load_workbook(B.OUT), 'CheckBatch', 'check')
    log(f'- LibreOffice ran CheckBatch: {line}')
    if ran is None:
        return False
    mc = ran['MacroCheck']
    rows = [(mc.cell(r, 1).value, mc.cell(r, 2).value) for r in range(B.CHK_FIRST, B.CHK_FIRST + 40)
            if mc.cell(r, 1).value]
    for a, b in rows:
        log(f'  - {a}: {b}')
    result = dict(rows).get('Result')
    worst = float(dict(rows).get('Largest relative difference (baseline and batch)', 'inf'))
    lookups = sum(1 for a, b in rows if str(b).startswith('OK ('))
    good = result == 'PASS' and worst < 1e-6 and lookups == 12
    ok &= good
    log(f'- first-run test: Result {result}, {lookups} lookups found, largest relative difference {worst:.1e} '
        f'{"OK" if good else "FAIL"}')
    wb = openpyxl.load_workbook(B.OUT)
    wb['Mitigation'].cell(B.R_POL['ets.ovr'], 1).value = 'ets.ovr.broken'
    line, ran = run_macro_lo(wb, 'CheckBatch', 'neg')
    mc = ran['MacroCheck'] if ran else None
    res = mc.cell(B.CHK_FIRST, 2).value if mc else None
    fails = [mc.cell(r, 1).value for r in range(B.CHK_FIRST + 1, B.CHK_FIRST + 40)
             if mc and str(mc.cell(r, 2).value or '').startswith('FAIL')] if mc else []
    good = str(res).startswith('FAIL') and fails == ['Mitigation: row ets.ovr (ETS override prices)'] and \
        line.split(' returned ')[-1].startswith('-1')
    ok &= good
    log(f'- negative test (row code ets.ovr broken): CheckBatch {line.split(" returned ")[-1]}; Result "{res}"; '
        f'failed lookup: {fails} {"OK" if good else "FAIL"}')
    return ok


def main():
    log(f'# Check report - {B.NAME}-v{B.VERSION}')
    results = []
    log('\n## 0. Formula length (at most 100 characters per cell formula)')
    results.append(check_formula_length())
    log('\n## 1. Shipped file recalculated in LibreOffice (no LAMBDA support)')
    results.append(check_shipped(None))
    log('\n## 2. Formula uniformity (relative R1C1)')
    results.append(check_uniformity(openpyxl.load_workbook(B.OUT)))
    log('\n## 3. LAMBDA encoding')
    results.append(check_encoding())
    log('\n## 4. Values vs independent Python recomputation')
    ws_a = recalc(B.OUT)
    ws_c = recalc(variant(openpyxl.load_workbook(B.OUT), 'everywhere'))
    for g, *_ in B.SCENARIOS:
        exp = recompute(scenario_mt(g))
        results.append(compare(ws_a, g, exp, f'A. as shipped (plain formulas, no LAMBDA), scenario {g}'))
        results.append(compare(ws_c, g, exp, f'B. optional LAMBDA library (sheet Lambdas) copied over every '
                                             f'calculated row, LAMBDAs expanded, scenario {g}'))
    log('\n## 4b. Power method = legacy (Python reference with the legacy inputs vs cached legacy rows)')
    results.append(check_power_method())
    log('\n## 5. Scenario tests (copied MTInputs column + Mitigation group)')
    results.append(check_scenarios())
    log('\n## 5b. ETS goal seek (ets_goalseek_v0_4.py)')
    results.append(check_goalseek())
    log('\n## 5c. Existing carbon pricing with non-zero test data (v1.09)')
    results.append(check_ecp())
    log('\n## 6. Regression vs v1.08 (shared output codes)')
    results.append(check_regression(ws_a))
    log('\n## 6b. MTOutputs and Charts')
    results.append(check_mto(recalc_wb(variant(openpyxl.load_workbook(B.OUT), 'expand')), [1, 2], 'as shipped'))
    results.append(check_charts())
    log('\n## 7. Labels and codes')
    results.append(check_codes(openpyxl.load_workbook(B.OUT)['Mitigation'], ws_a))
    log('\n## 8. Row outline')
    results.append(check_outline(openpyxl.load_workbook(B.OUT)['Mitigation']))
    log('\n## 9. MTInputs')
    results.append(check_mtinputs())
    log('\n## 10. Format')
    results.append(check_format(openpyxl.load_workbook(B.OUT)['Mitigation']))
    log('\n## 11. Sanity and results')
    results.append(check_sanity(ws_a))
    log('\n## 11b. Power sector: properties and results')
    results.append(check_power_results(ws_a))
    log('\n## 11c. Planned retirement switched on (v1.07)')
    results.append(check_planned_retirement())
    log('\n## 11d. Electricity demand memo lines (v1.07)')
    results.append(check_memo(ws_a))
    log('\n## 12. Scenario definitions, VBA module, stored results, comparison, macro run')
    results.append(check_definitions())
    results.append(check_vba())
    results.append(check_stored(recalc_wb(variant(openpyxl.load_workbook(B.OUT), 'expand'))))
    results.append(check_compare())
    results.append(check_macro())
    log('\n## 13. First-run test macro CheckBatchRun (v1.04)')
    results.append(check_first_run())
    log('\n## 14. Consistency (v1.09): check rows, identities, references, versions, stale texts')
    results.append(check_consistency(ws_a))
    log(f'\n**Overall: {"PASS" if all(results) else "FAIL"}**')
    with open(os.path.join(HERE, f'check_report_v{B.VERSION}.md'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(report) + '\n')


if __name__ == '__main__':
    main()
