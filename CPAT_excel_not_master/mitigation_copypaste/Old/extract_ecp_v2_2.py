"""Existing carbon pricing (ECP) data from the legacy CPAT workbook, sheet ECP -> data/ecp.csv (v2.2; written for model v1.09, version
number aligned with CPAT-AI-Mitigation-MVP v2.2).

Legacy sheet ECP (CPAT 1.0pre_456_NoPropData.xlsb; 'as of Apr 1, 2025'; sources WB Carbon Pricing Dashboard, ICAP)
holds, by country and year 2015-2025, indicator codes <iso3>.mit.<code>:
  carbon tax: ecr.all (rate, $/tCO2 nominal), ecr.<fuel> (fuel-specific rates, a few countries), ecf.<fuel> (fuel
              coverage share), ecs.<pow|res|ind|trs> (sector coverage share), crv.all (revenue, not used);
  ETS:        eer.all (permit price, $/tCO2 nominal; national block, plus eu.mit.eer.all for the regional EU ETS),
              eef.<fuel>, ees.<sector> (coverage; national and regional blocks), erv.all, efa.all (not used).
Legacy Mitigation rows 1675-1764 read these (two ETS blocks: national and regional) and legacy rows 2270-2288 use the
existing ETS sector coverage for the switch CTaxComplimentaryToETS.

Output (one ETS per country): key = ISO3|code, with
  - the national ETS where the country has one; otherwise the regional (EU) ETS: EU price and the country's regional
    coverage rows (countries with both, e.g. DEU, AUT: national ETS kept; see CAVEATS);
  - the carbon tax rate ecr.all and coverage ecf/ecs; fuel-specific rates ecr.<fuel> are kept in the CSV but not used
    by the model (v1.09);
  - years 2021-2025 (US CPI data start in 2021; the model's base year is 2022); coverage shares carried forward
    over blank years (legacy leaves some years blank, e.g. DEU ees.pow 2024-2025).
Egypt has no entries in legacy ECP: the model then reads zeros (the ECP sheet looks up with IFERROR(..., 0)).
"""
import csv
import os
from pyxlsb import open_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
LEGACY = os.path.join(HERE, '..', '..', 'cpat_excel_original', 'CPAT 1.0pre_456_NoPropData.xlsb')
OUT = os.path.join(HERE, 'data', 'ecp.csv')
YEARS = list(range(2021, 2026))
CODES = (['ecr.all'] + [f'ecr.{f}' for f in ('coa', 'nga', 'gso', 'die', 'lpg', 'ker', 'oop', 'jfu', 'bio')]
         + [f'ecs.{s}' for s in ('pow', 'res', 'ind', 'trs')]
         + [f'ecf.{f}' for f in ('coa', 'nga', 'gso', 'die', 'lpg', 'ker', 'oop', 'jfu', 'bio')]
         + ['eer.all'] + [f'ees.{s}' for s in ('pow', 'res', 'ind', 'trs')]
         + [f'eef.{f}' for f in ('coa', 'nga', 'gso', 'die', 'lpg', 'ker', 'oop', 'jfu', 'bio')])


def read_ecp():
    """Rows of sheet ECP: (row number, {column letter: value})."""
    rows = []
    with open_workbook(LEGACY) as wb, wb.get_sheet('ECP') as sh:
        for row in sh.rows():
            vals = {}
            for c in row:
                if c.v not in (None, ''):
                    vals[chr(65 + c.c) if c.c < 26 else None] = c.v
            if vals:
                rows.append((row[0].r + 1, vals))
    return rows


def main():
    rows = read_ecp()
    head = next(v for r, v in rows if r == 4)
    ycol = {int(float(v)): k for k, v in head.items() if k and k >= 'F' and str(v).replace('.0', '').isdigit()}
    block = None
    data = {}                                       # (block, iso3, code) -> {year: value}
    for r, v in rows:
        b = str(v.get('B', ''))
        if 'E' not in v and b in ('National ETS', 'Regional ETS'):
            block = 'nat' if b.startswith('National') else 'reg'
        if r < 600:
            block = 'tax'
        code = str(v.get('E', ''))
        if '.mit.' not in code:
            continue
        iso, c = code.split('.mit.')
        if c not in CODES:
            continue
        vals = {}
        for y in YEARS:
            x = v.get(ycol[y])
            try:
                vals[y] = float(x)
            except (TypeError, ValueError):
                vals[y] = None
        data[(block, iso.upper(), c)] = vals
    out = {}
    countries = sorted({k[1] for k in data if k[1] != 'EU'})
    for iso in countries:
        for c in CODES:
            if c.startswith('ec'):
                src = data.get(('tax', iso, c))
            elif ('nat', iso, 'eer.all') in data:
                src = data.get(('nat', iso, c))
            elif any(k[0] == 'reg' and k[1] == iso for k in data):
                src = data.get(('reg', 'EU', c)) if c == 'eer.all' else data.get(('reg', iso, c))
            else:
                src = None
            if src and any(x not in (None, 0.0) for x in src.values()):
                src = dict(src)
                if c[:3] in ('ecs', 'ecf', 'ees', 'eef'):      # coverage: carry the last value forward over blanks
                    last = None
                    for y in YEARS:
                        last = src[y] if src[y] is not None else last
                        src[y] = last
                out[(iso, c)] = src
    with open(OUT, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['key', 'country', 'code'] + YEARS)
        for (iso, c), vals in sorted(out.items()):
            w.writerow([f'{iso}|{c}', iso, c] + ['' if vals[y] is None else round(vals[y], 6) for y in YEARS])
    print(f'{len(out)} rows, {len({k[0] for k in out})} countries -> {OUT}')


if __name__ == '__main__':
    main()
