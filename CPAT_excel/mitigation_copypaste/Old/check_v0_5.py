"""Checks for CPAT_Mitigation_CopyPaste_v0.5.xlsx. Needs LibreOffice (soffice) to recalculate.

LibreOffice 24.2 cannot evaluate LAMBDA, so the 2035 column (named LAMBDAs) is checked two ways:
  A. expansion: each LAMBDA call is replaced by its body with the arguments substituted (what Excel
     computes), then recalculated;
  B. drag-forward: the plain 2034 formula is dragged into 2035 (the user's alternative).
Both must match an independent Python recomputation.

1. Shipped file: errors only in the 2035 column (LAMBDA, LibreOffice limitation) and nowhere else.
2. Formula uniformity per block (relative R1C1): one base-year, one plain, one LAMBDA formula.
3. LAMBDA encoding: every parameter carries _xlpm., no bare parameter names.
4. Values (variants A and B) vs independent Python recomputation from data/*.csv.
5. Copy test: group 2 copied to group 3 = group 2; copied to group 4 with carbon price 0 = group 1.
6. Regression: all four blocks equal v0.4 (Old/).
9. Row outline: blocks grouped under their summary lines, collapsed; totals grouped, open.
10. Format and summary lines: white band text; summary lines = total fuel use and policy carbon price.
8. Labels and codes: one formula per column; expected CPAT codes; copied groups number themselves.
7. Sanity: scenarios equal before 2027; oen and bio unaffected; totals check row 0.

    python check_v0_5.py   ->  writes check_report_v0.5.md
"""
import csv
import importlib.util
import os
import re
import subprocess
import tempfile

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI

import build_v0_5 as B

HERE = os.path.dirname(os.path.abspath(__file__))
REF = re.compile(r"(?<![A-Za-z_\d.\"])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])")
CALL = re.compile(r'^=(' + '|'.join(B.LAMBDAS) + r')\((.*)\)$')
TOL = 1e-9
report = []


def log(line=''):
    print(line)
    report.append(line)


def recalc(path):
    out = tempfile.mkdtemp()
    subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', out, path],
                   check=True, capture_output=True, timeout=600)
    return openpyxl.load_workbook(os.path.join(out, os.path.basename(path)), data_only=True)['Mitigation']


def r1c1(formula, row, col):
    def sub(m):
        c, rw = CI(m.group(2)), int(m.group(4))
        return (f'R{rw}' if m.group(3) else f'R[{rw - row}]') + (f'C{c}' if m.group(1) else f'C[{c - col}]')
    return REF.sub(sub, formula)


def block_rows(b0):
    return B.data_rows(b0)


def is_err(v):
    return isinstance(v, str) and (v.startswith('#') or v.startswith('Err:'))


def expand(formula):
    """Replace a named-LAMBDA call by its body with the arguments substituted."""
    m = CALL.match(formula)
    if not m:
        return formula
    params, body = B.LAMBDAS[m.group(1)]
    args = m.group(2).split(',')
    assert len(args) == len(params), formula
    for prm in sorted(params, key=len, reverse=True):
        body = re.sub(rf'(?<![A-Za-z_.]){prm}(?![A-Za-z_\d(])', f'({args[params.index(prm)]})', body)
    return '=' + body


def lambda_cells(ws, groups):
    return [(r, years[-1]) for name, (b0, *_r) in B.BLOCKS.items() for r in block_rows(b0)
            for _, years in groups]


def save_variant(wb, how):
    """how = 'expand' (LAMBDA bodies substituted) or 'drag' (plain 2034 formula dragged into 2035)."""
    ws = wb['Mitigation']
    if how == 'expand_old':  # v0.2 layout: group = base, years, gap
        groups = [(12 + (g - 1) * 15, list(range(13 + (g - 1) * 15, 26 + (g - 1) * 15))) for g in (1, 2)]
    else:
        groups = [B.group_cols(g)[1:] for g in range(1, 50) if ws.cell(B.R_YEAR, B.group_cols(g)[0]).value]
    for r, c in lambda_cells(ws, groups):
        if how.startswith('expand'):
            ws.cell(r, c).value = expand(ws.cell(r, c).value)
        else:
            ws.cell(r, c).value = Translator(ws.cell(r, c - 1).value, origin=f'{L(c - 1)}{r}').translate_formula(
                f'{L(c)}{r}')
    path = os.path.join(tempfile.mkdtemp(), f'{how}.xlsx')
    wb.save(path)
    return path


# ------------------------------------------------------------------ independent recomputation
def read(name):
    with open(os.path.join(HERE, 'data', name), encoding='utf-8') as f:
        return list(csv.DictReader(f))


def recompute():
    units = {r['fuel']: float(r['gj_per_price_unit']) for r in read('fuel_units.csv')}
    prices = {r['country_year']: r for r in read('prices_dom.csv')}['EGY2022']
    ef = {r['key']: float(r['ef_tco2_per_gj']) for r in read('ef_co2.csv')}
    el = {r['key']: float(r['LMIC']) for r in read('elasticities.csv')}
    use0 = {(r['subsector'], r['fuel']): float(r['ktoe']) for r in read('energy_use.csv')}
    weo = read('weo.csv')[0]
    sub = {s[0]: s for s in B.SUBSECTORS}
    fue = {f[0]: f for f in B.FUELS}
    res = {}
    for g, _, cp in B.SCENARIOS:
        for s, f in B.PAIRS:
            _, _, grp, es, efs, cgs, cov = sub[s]
            _, _, elf, aef, rule, _ = fue[f]
            code = f'{f}.{cgs if rule == "sub" else "all"}'
            sp = float(prices[f'mit.sp.{code}']) / units[f]
            btax = (float(prices[f'mit.rp.{code}']) - float(prices[f'mit.sp.{code}'])) / units[f]
            e = ef[f'EGY|{f}|{efs}']
            eY, eU, eF = el[f'inc|{elf}|{es}'], el[f'usg|{elf}|{es}'], el[f'eff|{elf}|{es}']
            a = el[f'aei|{aef}|{es}']
            pre, tax, post, use = [sp], [btax], [sp + btax], [use0[(s, f)]]
            for y in B.YEARS:
                pre.append(pre[-1])
                tax.append(btax + cp.get(y, 0) * e * cov)
                post.append(pre[-1] + tax[-1])
                ratio = post[-1] / post[-2]
                use.append(use[-1] * (1 / (1 + a)) ** (1 + eU) * (1 + float(weo[str(y)])) ** eY
                           * ratio ** eU * ratio ** (eF * (1 + eU)))
            for blk, vals in (('pre', pre), ('tax', tax), ('post', post), ('use', use)):
                res[(g, blk, s, f)] = vals
    return res


def compare_python(ws, label):
    exp = recompute()
    worst = (0.0, None)
    for name, (b0, *_r) in B.BLOCKS.items():
        for i, r in enumerate(block_rows(b0)):
            s, f = B.PAIRS[i]
            for g, *_ in B.SCENARIOS:
                _, base, years = B.group_cols(g)
                for k, c in enumerate([base] + years):
                    v, e = ws.cell(r, c).value, exp[(g, name, s, f)][k]
                    d = float('inf') if not isinstance(v, (int, float)) else abs(v - e) / max(1.0, abs(e))
                    if d > worst[0]:
                        worst = (d, f'{L(c)}{r}')
    ok = worst[0] < TOL
    log(f'- {label}: {len(B.BLOCKS)} blocks x {B.NROW} rows x {len(B.SCENARIOS)} scenarios x '
        f'{B.NY} years, max relative diff {worst[0]:.2e} (at {worst[1]}) {"OK" if ok else "FAIL"}')
    return ok


# ------------------------------------------------------------------ checks
def check_shipped(ws):
    groups = [B.group_cols(g)[1:] for g, *_ in B.SCENARIOS]
    lam_cols = {years[-1] for _, years in groups}
    errs = [(c.row, c.column) for row in ws.iter_rows() for c in row if is_err(c.value)]
    outside = [f'{L(c)}{r}' for r, c in errs if c not in lam_cols]
    ok = not outside
    log(f'- error values: {len(errs)}, all in the {B.YEARS[-1]} (LAMBDA) columns '
        f'{sorted(L(c) for c in lam_cols)}: {"OK" if ok else "FAIL, outside: " + str(outside[:10])}')
    return ok


def check_uniformity(wb):
    ws = wb['Mitigation']
    ok = True
    groups = [B.group_cols(g)[1:] for g, *_ in B.SCENARIOS]
    for name, (b0, *_r) in B.BLOCKS.items():
        sets = {'base': set(), 'plain': set(), 'lambda': set()}
        for base, years in groups:
            for r in block_rows(b0):
                sets['base'].add(r1c1(ws.cell(r, base).value, r, base))
                for c in years[:-1]:
                    sets['plain'].add(r1c1(ws.cell(r, c).value, r, c))
                sets['lambda'].add(r1c1(ws.cell(r, years[-1]).value, r, years[-1]))
        params = {pc: {r1c1(str(ws[f'{pc}{r}'].value), r, CI(pc)) for r in block_rows(b0)} for pc in B.PARAM_COLS}
        good = all(len(v) == 1 for v in sets.values()) and all(len(v) == 1 for v in params.values())
        ok &= good
        log(f'- `{name}`: base {len(sets["base"])}, plain {B.YEARS[0]}-{B.YEARS[-2]} {len(sets["plain"])}, '
            f'LAMBDA {B.YEARS[-1]} {len(sets["lambda"])}, parameter columns D:G '
            f'{[len(v) for v in params.values()]} formula(s) {"OK" if good else "FAIL"}')
        log(f'  - plain: `{next(iter(sets["plain"]))}`')
        log(f'  - LAMBDA: `{next(iter(sets["lambda"]))}`')
    hidden = all(ws.column_dimensions[c].hidden for c in ['D']) and ws.column_dimensions['D'].max == 7
    log(f'- parameter columns D:G grouped and hidden: {"OK" if hidden else "FAIL"}')
    return ok and hidden


def check_encoding():
    ok = True
    for name, (params, _) in B.LAMBDAS.items():
        xml = B.lambda_xml(name)
        bare = [p for p in params if re.search(rf'(?<![A-Za-z_.]){p}(?![A-Za-z_\d(])', xml)]
        expected = xml.count('_xlpm.') == len(params) + sum(
            len(re.findall(rf'(?<![A-Za-z_.]){p}(?![A-Za-z_\d(])', B.LAMBDAS[name][1])) for p in params)
        good = not bare and expected and xml.startswith('_xlfn.LAMBDA(')
        ok &= good
        log(f'- {name}: {len(params)} parameters, {xml.count("_xlpm.")} _xlpm. tokens, bare names {bare} '
            f'{"OK" if good else "FAIL"}')
    return ok


def check_copy():
    wb = openpyxl.load_workbook(B.OUT)
    ws = wb['Mitigation']
    last_row = ws.max_row
    for src, dst, zero in [(2, 3, False), (2, 4, True)]:
        s0, d0 = B.group_cols(src)[0], B.group_cols(dst)[0]
        for k in range(B.GW):
            for r in range(1, last_row + 1):
                v = ws.cell(r, s0 + k).value
                if isinstance(v, str) and v.startswith('='):
                    v = Translator(v, origin=f'{L(s0 + k)}{r}').translate_formula(f'{L(d0 + k)}{r}')
                ws.cell(r, d0 + k).value = v
        if zero:
            for c in range(d0 + 1, d0 + 1 + B.NY):
                ws.cell(B.R_CP, c).value = 0
    ws = recalc(save_variant(wb, 'expand'))
    ok = True
    for a, b, label in [(2, 3, 'scenario 3 (copy of 2) = scenario 2'),
                        (1, 4, 'scenario 4 (copy of 2, carbon price 0) = scenario 1')]:
        ca, cb = B.group_cols(a), B.group_cols(b)
        diff, n = 0.0, 0
        for r in range(B.R_YEAR, last_row + 1):
            if r in (B.R_CP, B.R_NAME):
                continue
            for x, y in zip([ca[1]] + ca[2], [cb[1]] + cb[2]):
                va, vb = ws.cell(r, x).value, ws.cell(r, y).value
                if isinstance(va, (int, float)) and isinstance(vb, (int, float)):
                    diff, n = max(diff, abs(va - vb)), n + 1
                elif va != vb:
                    diff = float('inf')
        good = diff < TOL
        ok &= good
        log(f'- {label}: {n} numeric cells, max abs diff {diff:.3g} {"OK" if good else "FAIL"}')
    for g in (3, 4):
        code = B.group_cols(g)[0]
        num_ = ws.cell(B.R_YEAR, code).value
        r = B.B_USE + 2
        good = num_ == g and str(ws.cell(r, code).value).endswith(f'.e.{g}')
        ok &= good
        log(f'- pasted group {g}: scenario number {num_}, code `{ws.cell(r, code).value}` {"OK" if good else "FAIL"}')
    return ok


def check_regression(ws):
    spec = importlib.util.spec_from_file_location('b01', os.path.join(HERE, 'Old', 'build_v0_4.py'))
    b01 = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(b01)
    wbo = openpyxl.load_workbook(os.path.join(HERE, 'Old', 'CPAT_Mitigation_CopyPaste_v0.4.xlsx'))
    wso = wbo['Mitigation']
    for _, (b0, *_r) in b01.BLOCKS.items():          # expand the LAMBDA year with the v0.4 layout
        for r in range(b0 + 1, b0 + 1 + b01.NROW):
            for g in (1, 2):
                c = b01.group_cols(g)[2][-1]
                wso.cell(r, c).value = expand(wso.cell(r, c).value)
    path = os.path.join(tempfile.mkdtemp(), 'old.xlsx')
    wbo.save(path)
    old = recalc(path)
    old_b = {'pre': b01.B_PRE, 'tax': b01.B_TAX, 'post': b01.B_POST, 'use': b01.B_USE}
    diff, n = 0.0, 0
    for name, (b0, *_r) in B.BLOCKS.items():
        for i in range(B.NROW):
            for g, *_ in B.SCENARIOS:
                _, ob, oy = b01.group_cols(g)
                _, nb, ny = B.group_cols(g)
                for oc, nc in zip([ob] + oy, [nb] + ny):
                    diff = max(diff, abs(old.cell(old_b[name] + 1 + i, oc).value - ws.cell(b0 + 2 + i, nc).value))
                    n += 1
    ok = diff < TOL
    log(f'- {n} block cells vs v0.4: max abs diff {diff:.3g} {"OK" if ok else "FAIL"}')
    return ok


def check_sanity(ws):
    ok = True
    (_, b1, y1), (_, b2, y2) = B.group_cols(1), B.group_cols(2)
    cols1, cols2, years = [b1] + y1, [b2] + y2, [B.BASE_YEAR] + B.YEARS
    pre = max(abs(ws.cell(r, c1).value - ws.cell(r, c2).value)
              for r in block_rows(B.B_USE) for c1, c2, y in zip(cols1, cols2, years) if y < 2027)
    log(f'- scenario 1 vs 2 before 2027 (fuel use): max abs diff {pre:.3g} {"OK" if pre < TOL else "FAIL"}')
    ok &= pre < TOL
    for label, pick in [('oen rows (no coverage)', lambda s, f: s == 'oen'), ('bio rows (EF 0)', lambda s, f: f == 'bio')]:
        d = max(abs(ws.cell(r, c1).value - ws.cell(r, c2).value)
                for i, r in enumerate(block_rows(B.B_USE)) if pick(*B.PAIRS[i]) for c1, c2 in zip(cols1, cols2))
        log(f'- {label}: scenario 1 = scenario 2, max abs diff {d:.3g} {"OK" if d < TOL else "FAIL"}')
        ok &= d < TOL
    r_tot = B.R_TOTAL
    chk = max(abs(ws.cell(B.R_CHK, c).value) for c in cols1 + cols2)
    log(f'- check row {B.R_CHK} (totals by subsector and by fuel vs total): max abs {chk:.3g} {"OK" if chk < 1e-6 else "FAIL"}')
    ok &= chk < 1e-6
    log('')
    log('| Year | Scenario 1 (ktoe) | Scenario 2 (ktoe) | Change |')
    log('|---|---|---|---|')
    for c1, c2, y in zip(cols1, cols2, years):
        if y in (2022, 2026, 2027, 2030, 2035):
            a, b = ws.cell(r_tot, c1).value, ws.cell(r_tot, c2).value
            log(f'| {y} | {a:,.0f} | {b:,.0f} | {b / a - 1:+.1%} |')
    return ok


def check_codes(wsf, wsv):
    ok = True
    described = [r for r in range(1, wsf.max_row + 1)
                 if isinstance(wsf[f'H{r}'].value, str) and wsf[f'H{r}'].value.startswith('=INDEX(Variables')]
    for col in 'HIJ':
        forms = {r1c1(wsf[f'{col}{r}'].value, r, CI(col)) for r in described}
        good = len(forms) == 1
        ok &= good
        log(f'- column {col}: {len(described)} rows, {len(forms)} formula {"OK" if good else "FAIL"}')
    for g, *_ in B.SCENARIOS:
        code = B.group_cols(g)[0]
        rows = [r for r in described if wsf.cell(r, code).value]
        forms = {r1c1(wsf.cell(r, code).value, r, code) for r in rows}
        good = len(forms) == 1 and len(rows) == len(described)
        ok &= good
        log(f'- code column {L(code)} (scenario {g}): {len(rows)} rows, {len(forms)} formula {"OK" if good else "FAIL"}')
    k1, k2 = B.group_cols(1)[0], B.group_cols(2)[0]
    u = B.B_USE + 2 + B.PAIRS.index(('rod', 'gso'))
    r_tot = B.R_TOTAL
    expected = [(u, k1, 'egy.mit.ener.rod.gso.e.1', 'Fuel use | Road | Gasoline'),
                (u, k2, 'egy.mit.ener.rod.gso.e.2', None),
                (u - 3 * B.BLOCK_H, k1, 'egy.mit.sp.rod.gso.a.1', 'Pre-tax price (supply cost) | Road | Gasoline'),
                (u - B.BLOCK_H, k2, 'egy.mit.atp.rod.gso.e.2', 'After-tax price | Road | Gasoline'),
                (B.R_GDP, k1, 'egy.mit.gdp.pos.pct.1', 'Real GDP growth'),
                (B.R_CP, k2, 'egy.mit.cptraj.2', 'Carbon price'),
                (B.B_TOT + 1, k1, 'egy.mit.ener.rod.all.e.1', 'Fuel use | Road | All fuels'),
                (r_tot, k2, 'egy.mit.ener.all.all.e.2', 'Fuel use | All subsectors | All fuels'),
                (B.B_TAX + 1, k1, 'egy.mit.cptraj.ref.1', 'Policy carbon price (as in the scenario assumptions)')]
    for r, c, code, label in expected:
        got, lab = wsv.cell(r, c).value, wsv[f'H{r}'].value
        good = got == code and (label is None or lab == label)
        ok &= good
        log(f'- {L(c)}{r}: `{got}` | {lab} {"OK" if good else "FAIL (expected " + code + ")"}')
    nums = [wsv.cell(B.R_YEAR, B.group_cols(g)[0]).value for g, *_ in B.SCENARIOS]
    good = nums == [g for g, *_ in B.SCENARIOS]
    ok &= good
    log(f'- scenario numbers at the top of the code columns: {nums} {"OK" if good else "FAIL"}')
    return ok


def check_outline(wsf):
    ok = wsf.sheet_properties.outlinePr.summaryBelow is False
    for name, (b0, *_r) in B.BLOCKS.items():
        rows = range(b0 + 2, b0 + B.NROW + 3)
        good = (all(wsf.row_dimensions[r].outlineLevel == 1 and wsf.row_dimensions[r].hidden for r in rows)
                and all(wsf.row_dimensions[x].outlineLevel == 0 and not wsf.row_dimensions[x].hidden
                        for x in (b0, b0 + 1)))
        ok &= good
        log(f'- block `{name}`: rows {rows[0]}-{rows[-1]} grouped under summary line {b0 + 1} (band {b0} and summary '
            f'visible), collapsed {"OK" if good else "FAIL"}')
    rows = range(B.R_SUB0, B.R_PCT + 1)
    good = all(wsf.row_dimensions[r].outlineLevel == 1 and not wsf.row_dimensions[r].hidden for r in rows)
    ok &= good
    log(f'- totals: rows {rows[0]}-{rows[-1]} grouped under band row {B.B_TOT}, open {"OK" if good else "FAIL"}')
    return ok


def check_format(wsf, wsv):
    ok = True
    bands = [3, B.B_PRE, B.B_TAX, B.B_POST, B.B_USE, B.B_TOT]
    white = all(str(wsf[f'A{r}'].font.color.rgb).endswith('FFFFFF') and wsf[f'A{r}'].font.b for r in bands)
    ok &= white
    log(f'- band rows {bands}: white bold text {"OK" if white else "FAIL"}')
    for g, *_ in B.SCENARIOS:
        _, base, years = B.group_cols(g)
        d_cp = max(abs(wsv.cell(B.B_TAX + 1, c).value - wsv.cell(B.R_CP, c).value) for c in [base] + years)
        d_tot = max(abs(wsv.cell(B.R_TOTAL, c).value - sum(wsv.cell(r, c).value for r in B.data_rows(B.B_USE)))
                    for c in [base] + years)
        good = d_cp < TOL and d_tot < 1e-6
        ok &= good
        log(f'- scenario {g}: tax summary = carbon price (diff {d_cp:.3g}); fuel-use summary = sum of 128 rows '
            f'(diff {d_tot:.3g}) {"OK" if good else "FAIL"}')
    return ok


def main():
    log(f'# Check report - CPAT_Mitigation_CopyPaste_v{B.VERSION}')
    log('')
    log('## 1. Shipped file recalculated in LibreOffice (no LAMBDA support)')
    r1 = check_shipped(recalc(B.OUT))
    log('')
    log('## 2. Formula uniformity (relative R1C1)')
    r2 = check_uniformity(openpyxl.load_workbook(B.OUT))
    log('')
    log('## 3. LAMBDA encoding')
    r3 = check_encoding()
    log('')
    log('## 4. Values vs independent Python recomputation')
    ws_a = recalc(save_variant(openpyxl.load_workbook(B.OUT), 'expand'))
    r4a = compare_python(ws_a, f'A. LAMBDA bodies expanded in {B.YEARS[-1]}')
    ws_b = recalc(save_variant(openpyxl.load_workbook(B.OUT), 'drag'))
    r4b = compare_python(ws_b, f'B. plain formula dragged into {B.YEARS[-1]}')
    log('')
    log('## 5. Scenario copy test')
    r5 = check_copy()
    log('')
    log('## 6. Regression vs v0.4')
    r6 = check_regression(ws_a)
    log('')
    log('## 7. Labels and codes')
    r8 = check_codes(openpyxl.load_workbook(B.OUT)['Mitigation'], ws_a)
    log('')
    log('## 8. Row outline')
    r9 = check_outline(openpyxl.load_workbook(B.OUT)['Mitigation'])
    log('')
    log('## 9. Format and summary lines')
    r10 = check_format(openpyxl.load_workbook(B.OUT)['Mitigation'], ws_a)
    log('')
    log('## 10. Sanity and results')
    r7 = check_sanity(ws_a)
    log('')
    allok = all([r1, r2, r3, r4a, r4b, r5, r6, r7, r8, r9, r10])
    log(f'**Overall: {"PASS" if allok else "FAIL"}**')
    with open(os.path.join(HERE, f'check_report_v{B.VERSION}.md'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(report) + '\n')


if __name__ == '__main__':
    main()
