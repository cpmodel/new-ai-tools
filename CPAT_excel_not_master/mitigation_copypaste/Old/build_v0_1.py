"""Build CPAT_Mitigation_CopyPaste_v0.1.xlsx (design 2: time across, scenarios as column groups).

Price -> fuel-use prototype of the CPAT mitigation module whose formulas are fully copy-pasteable:
every block has one formula text (in relative R1C1 terms) across all its projection cells, and a whole
scenario group (15 columns) can be copied to the next group position and keeps working.

Reads data/*.csv (made by extract_data_v0_1.py). Writes formulas only (no cached values); open in Excel
or recalculate with LibreOffice (see check_v0_1.py).

    python build_v0_1.py
"""
import csv
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

VERSION = '0.1'
HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, 'data')
OUT = os.path.join(HERE, f'CPAT_Mitigation_CopyPaste_v{VERSION}.xlsx')

# ---------------------------------------------------------------- dimensions
SUBSECTORS = [  # code, name, group, elasticity sector, EF sector, coal/gas price sector, carbon-price coverage
    ('rod', 'Road', 'tra', 'tra', 'rod', 'res', 1),
    ('ral', 'Rail', 'tra', 'tra', 'rod', 'ind', 1),
    ('avi', 'Domestic aviation', 'tra', 'tra', 'rod', 'ind', 1),
    ('nav', 'Domestic navigation', 'tra', 'tra', 'rod', 'ind', 1),
    ('res', 'Residential', 'bld', 'res', 'res', 'res', 1),
    ('foo', 'Food & forestry', 'bld', 'ind', 'res', 'ind', 1),
    ('srv', 'Services (public & private)', 'bld', 'srv', 'res', 'ind', 1),
    ('mch', 'Mining & chemicals', 'ind', 'ind', 'ind', 'ind', 1),
    ('irn', 'Iron & steel', 'ind', 'ind', 'ind', 'ind', 1),
    ('nfm', 'Other metals', 'ind', 'ind', 'ind', 'ind', 1),
    ('mac', 'Machinery', 'ind', 'ind', 'ind', 'ind', 1),
    ('cem', 'Cement', 'ind', 'ind', 'ind', 'ind', 1),
    ('omn', 'Other manufacturing', 'ind', 'ind', 'ind', 'ind', 1),
    ('cst', 'Construction', 'ind', 'ind', 'ind', 'ind', 1),
    ('ftr', 'Fuel transformation', 'ind', 'ind', 'ind', 'ind', 1),
    ('oen', 'Other energy use', 'oen', 'ind', 'ind', 'ind', 0),
]
FUELS = [  # code, name, elasticity fuel, AEEI fuel, price-sector rule, price unit
    ('coa', 'Coal', 'coa', 'coa', 'sub', '$/GJ'),
    ('nga', 'Natural gas', 'nga', 'nga', 'sub', '$/GJ'),
    ('gso', 'Gasoline', 'gso', 'oil', 'all', '$/liter'),
    ('die', 'Diesel', 'die', 'oil', 'all', '$/liter'),
    ('lpg', 'LPG', 'oil', 'oil', 'all', '$/liter'),
    ('ker', 'Kerosene', 'oil', 'oil', 'all', '$/liter'),
    ('oop', 'Other oil products', 'oil', 'oil', 'all', '$/bbl'),
    ('bio', 'Biomass', 'bio', 'bio', 'all', '$/GJ'),
]
NS, NF = len(SUBSECTORS), len(FUELS)
NROW = NS * NF                          # 128 rows per block (subsector outer, fuel inner)
SNAMES = {s[0]: s[1] for s in SUBSECTORS}
FNAMES = {f[0]: f[1] for f in FUELS}

BASE_YEAR, LAST_YEAR = 2022, 2035
YEARS = list(range(BASE_YEAR + 1, LAST_YEAR + 1))
N_SETS = 4                              # parameter sets (= scenarios that fit without inserting columns)
N_SLOTS = 5                             # parameter slots per set (P5 = buffer)
SCENARIOS = [(1, 'Baseline (no carbon price)', {}),
             (2, 'Carbon price $20/tCO2 from 2027', {y: 20 for y in YEARS if y >= 2027})]
SPARE_SCENARIOS = [(3, 'Spare (copy a group here)'), (4, 'Spare (copy a group here)')]

# ---------------------------------------------------------------- Mitigation layout
COL_KEY0 = 11                           # K: first key column
KEYS = ['Elasticity sector', 'Elasticity fuel', 'AEEI fuel', 'Price code', 'EF sector', 'GJ per price unit']
COL_P0 = COL_KEY0 + len(KEYS)           # Q: first parameter column
COL_P1 = COL_P0 + N_SETS * N_SLOTS - 1  # AJ: last parameter column
COL_G0 = COL_P1 + 1                     # AK: first scenario group
GW = 2 + len(YEARS)                     # group width: label column, base year, projection years
PSET = f'${L(COL_P0)}{{r}}:${L(COL_P1)}{{r}}'

R_SCEN, R_YEAR, R_NAME, R_GDP, R_CP = 4, 5, 6, 7, 8
BLOCK_H = NROW + 2                      # band row + 128 rows + blank row
B_PRE, B_TAX, B_POST, B_USE = 10, 10 + BLOCK_H, 10 + 2 * BLOCK_H, 10 + 3 * BLOCK_H
B_TOT = 10 + 4 * BLOCK_H

# ---------------------------------------------------------------- styles (NORMS section 2)
def fill(hex_):
    return PatternFill('solid', start_color=hex_, end_color=hex_)


F_TITLE, F_BAND, F_INPUT, F_CODE = fill('00B050'), fill('92D050'), fill('EBF1DE'), fill('DCE6F1')
F_BASE, F_UNUSED, F_CHECK = fill('DDD9C4'), fill('F2F2F2'), fill('FFF2CC')
FONT = Font(name='Arial', size=9)
FONT_B = Font(name='Arial', size=9, bold=True)
FONT_T = Font(name='Arial', size=14, bold=True, color='FFFFFF')
THIN = Side(style='thin', color='808080')


def put(ws, cell, value, font=FONT, fill_=None, fmt=None):
    c = ws[cell]
    c.value = value
    c.font = font
    if fill_:
        c.fill = fill_
    if fmt:
        c.number_format = fmt
    return c


def title(ws, text, width):
    for j in range(1, width + 1):
        ws.cell(2, j).fill = F_TITLE
    put(ws, 'A2', text, FONT_T, F_TITLE)


def band(ws, row, text, width, first=1):
    for j in range(first, width + 1):
        ws.cell(row, j).fill = F_BAND
    put(ws, f'{L(first)}{row}', text, FONT_B, F_BAND)


def read_csv(name):
    with open(os.path.join(DATA, name), encoding='utf-8') as f:
        return list(csv.reader(f))


def num(v):
    if v in ('', None):
        return None
    try:
        return float(v)
    except ValueError:
        return v


# ---------------------------------------------------------------- data sheets
def data_sheet(wb, name, title_text, header, rows, source):
    ws = wb.create_sheet(name)
    title(ws, title_text, len(header))
    put(ws, 'A1', f'Source: {source}', FONT)
    for j, h in enumerate(header, 1):
        put(ws, f'{L(j)}3', num(h) if isinstance(h, str) and h.isdigit() else h, FONT_B, F_INPUT)
    for i, r in enumerate(rows, 4):
        for j, v in enumerate(r, 1):
            ws.cell(i, j, num(v)).font = FONT
    ws.freeze_panes = 'B4'
    ws.column_dimensions['A'].width = 18
    return ws


def build_data_sheets(wb):
    c = read_csv('countries.csv')
    data_sheet(wb, 'Countries', 'Countries: code, region and income group', c[0], c[1:],
               'legacy Elasticities rows 138-141 (country columns)')
    e = read_csv('elasticities.csv')
    data_sheet(wb, 'Elasticities', 'Elasticities by income group (legacy CPAT values)', e[0], e[1:],
               "legacy Elasticities rows 143-282 ('Simple' option, income-group specific)")
    hdr, rows = read_csv('prices_dom_header.csv'), read_csv('prices_dom.csv')
    ws = wb.create_sheet('Prices_dom')  # legacy Prices_dom layout (NORMS 1.3): bands rows 1-4, codes row 5
    for i, r in enumerate(hdr[1:5], 1):
        for j, v in enumerate(r, 1):
            if v:
                ws.cell(i, j, v).font = FONT
    for j, h in enumerate(rows[0], 1):
        put(ws, f'{L(j)}5', h, FONT_B, F_INPUT)
    for i, r in enumerate(rows[1:], 6):
        for j, v in enumerate(r, 1):
            ws.cell(i, j, num(v) if j > 4 else (int(float(v)) if j == 5 else v)).font = FONT
    ws.freeze_panes = 'F6'
    u = read_csv('energy_use.csv')
    data_sheet(wb, 'EnergyCons', 'Base-year energy use by subsector and fuel (ktoe)', u[0], u[1:],
               'legacy Mitigation rows 583-602 via kernel v1.6 Data_Energy (Egypt only so far)')
    w = read_csv('weo.csv')
    data_sheet(wb, 'WEO', 'Macro series by country (real GDP growth)', w[0], w[1:],
               'legacy Mitigation row 2446 via kernel v1.6 Data_Macro (Egypt only so far)')
    f = read_csv('ef_co2.csv')
    data_sheet(wb, 'EF_GHG', 'Fuel CO2 emission factors, IIASA (tCO2/GJ), by country, fuel and EF sector',
               f[0], f[1:], 'legacy Mitigation rows 980-1009 (Egypt only so far; no inventory adjustment)')
    return len(rows) - 1, len(u) - 1, len(w) - 1, len(f) - 1, len(c) - 1, len(e) - 1


# ---------------------------------------------------------------- Mapping and Settings
def build_mapping(wb):
    ws = wb.create_sheet('Mapping')
    title(ws, 'Mapping: subsectors and fuels to CPAT parameter groups', 8)
    put(ws, 'A3', 'Subsectors', FONT_B)
    for j, h in enumerate(['Code', 'Name', 'Group', 'Elasticity sector', 'EF sector',
                           'Coal/gas price sector', 'Carbon-price coverage'], 1):
        put(ws, f'{L(j)}4', h, FONT_B, F_INPUT)
    for i, s in enumerate(SUBSECTORS, 5):
        for j, v in enumerate(s, 1):
            put(ws, f'{L(j)}{i}', v, FONT, F_INPUT if j > 2 else None)
    units = {r[0]: float(r[1]) for r in read_csv('fuel_units.csv')[1:]}
    put(ws, 'A22', 'Fuels', FONT_B)
    for j, h in enumerate(['Code', 'Name', 'Elasticity fuel', 'AEEI fuel', 'Price sector rule',
                           'Price unit', 'GJ per price unit'], 1):
        put(ws, f'{L(j)}23', h, FONT_B, F_INPUT)
    for i, f in enumerate(FUELS, 24):
        for j, v in enumerate(list(f) + [units[f[0]]], 1):
            put(ws, f'{L(j)}{i}', v, FONT, F_INPUT if j > 2 else None)
    put(ws, 'A33', "Price sector rule: 'sub' = use the subsector's coal/gas price sector (res or ind, "
                   "as CPAT: rod and res use residential coal/gas prices); 'all' = the fuel's single price.", FONT)
    put(ws, 'A34', 'GJ per price unit: legacy Mitigation rows 980-1009, conversion factor (volume unit to GJ).', FONT)
    put(ws, 'A35', 'Coverage: legacy docs 3.3.4.6 - no carbon taxes in other energy use (oen).', FONT)
    ws.column_dimensions['B'].width = 26
    for col in 'CDEFG':
        ws.column_dimensions[col].width = 16


def build_settings(wb):
    ws = wb.create_sheet('Settings')
    title(ws, 'Settings', 6)
    rows = [(4, 'Country code', 'EGY', True), (5, 'Country name',
            '=INDEX(Countries!$B$4:$B$1000,MATCH($C$4,Countries!$A$4:$A$1000,0))', False),
            (6, 'Income group (selects elasticities)',
             '=INDEX(Countries!$D$4:$D$1000,MATCH($C$4,Countries!$A$4:$A$1000,0))', False),
            (7, 'Base year', BASE_YEAR, True)]
    for r, lab, val, inp in rows:
        put(ws, f'B{r}', lab)
        put(ws, f'C{r}', val, FONT, F_INPUT if inp else None)
    put(ws, 'B9', 'Scenario number', FONT_B, F_INPUT)
    put(ws, 'C9', 'Scenario name', FONT_B, F_INPUT)
    for i, (n, name) in enumerate([(s[0], s[1]) for s in SCENARIOS] + SPARE_SCENARIOS, 10):
        put(ws, f'B{i}', n)
        put(ws, f'C{i}', name, FONT, F_INPUT)
    put(ws, 'B16', 'Version log', FONT_B)
    for j, h in enumerate(['Version', 'Date', 'Description', 'Max abs regression diff'], 2):
        put(ws, f'{L(j)}17', h, FONT_B, F_INPUT)
    put(ws, 'B18', VERSION)
    put(ws, 'C18', '2026-10-08')
    put(ws, 'D18', 'First build: design 2 price -> fuel-use prototype, Egypt base data, 2 scenarios.')
    put(ws, 'E18', 'n/a (first version)')
    ws.column_dimensions['B'].width = 34
    ws.column_dimensions['C'].width = 34
    ws.column_dimensions['D'].width = 70


# ---------------------------------------------------------------- Mitigation
def p(k, col, r):
    """Pick parameter slot k (1..5) for the scenario in this column's group."""
    return f'INDEX({PSET.format(r=r)},1,({col}$4-1)*{N_SLOTS}+{k})'


def lookup_price(code_prefix, r, year_ref):
    return (f'INDEX(Prices_dom!$A$6:$CU$2000,MATCH(Settings!$C$4&{year_ref},Prices_dom!$A$6:$A$2000,0),'
            f'MATCH("{code_prefix}."&$N{r},Prices_dom!$A$5:$CU$5,0))')


def lookup_elast(typ, fuel_col, r):
    return (f'=INDEX(Elasticities!$F$4:$I$200,MATCH("{typ}|"&${fuel_col}{r}&"|"&$K{r},Elasticities!$A$4:$A$200,0),'
            f'MATCH(Settings!$C$6,Elasticities!$F$3:$I$3,0))')


BLOCKS = {  # band row: (name, code, unit, slot labels, group-cell formula maker)
    'pre': (B_PRE, 'Pre-tax price (supply cost)', 'pbt', '$/GJ',
            ['Annual growth', '-', '-', '-', 'buffer']),
    'tax': (B_TAX, 'Tax: base tax + carbon price x EF x coverage', 'tax', '$/GJ',
            ['Base tax ($/GJ)', 'EF (tCO2/GJ)', 'Coverage', '-', 'buffer']),
    'post': (B_POST, 'Post-tax price = pre-tax + tax', 'ptp', '$/GJ', ['-', '-', '-', '-', 'buffer']),
    'use': (B_USE, 'Fuel use (CPAT docs 3.3.3)', 'enc', 'ktoe',
            ['eps_Y GDP', 'eps_U usage', 'eps_F efficiency', 'alpha AEEI', 'buffer']),
}
SHORT = {'pre': 'Pre-tax price', 'tax': 'Tax', 'post': 'Post-tax price', 'use': 'Fuel use'}
USED_SLOTS = {'pre': [1], 'tax': [1, 2, 3], 'post': [], 'use': [1, 2, 3, 4]}


def set1_formula(block, k, r):
    if block == 'pre' and k == 1:
        return 0, True                                   # annual real growth of the pre-tax price (input)
    if block == 'tax':
        if k == 1:
            return (f'=({lookup_price("mit.rp", r, "Settings!$C$7")}-{lookup_price("mit.sp", r, "Settings!$C$7")})'
                    f'/$P{r}'), False
        if k == 2:
            return (f'=INDEX(EF_GHG!$E$4:$E$1000,MATCH(Settings!$C$4&"|"&$B{r}&"|"&$O{r},'
                    f'EF_GHG!$A$4:$A$1000,0))'), False
        if k == 3:
            return f'=INDEX(Mapping!$G$5:$G$20,MATCH($C{r},Mapping!$A$5:$A$20,0))', False
    if block == 'use':
        return {1: lookup_elast('inc', 'L', r), 2: lookup_elast('usg', 'L', r),
                3: lookup_elast('eff', 'L', r), 4: lookup_elast('aei', 'M', r)}[k], False
    return None, False


def group_formula(block, r, col, prev, is_base):
    if block == 'pre':
        if is_base:
            return f'={lookup_price("mit.sp", r, f"{col}$5")}/$P{r}'
        return f'={prev}{r}*(1+{p(1, col, r)})'
    if block == 'tax':
        return f'={p(1, col, r)}+{col}${R_CP}*{p(2, col, r)}*{p(3, col, r)}'
    if block == 'post':
        return f'={col}{r - 2 * BLOCK_H}+{col}{r - BLOCK_H}'
    if block == 'use':
        if is_base:
            return (f'=INDEX(EnergyCons!$F$4:$F$5000,MATCH(Settings!$C$4&"|"&{col}$5&"|"&$C{r}&"|"&$B{r},'
                    f'EnergyCons!$A$4:$A$5000,0))')
        ratio = f'({col}{r - BLOCK_H}/{prev}{r - BLOCK_H})'
        eY, eU, eF, a = (p(k, col, r) for k in (1, 2, 3, 4))
        return (f'={prev}{r}*(1/(1+{a}))^(1+{eU})*(1+{col}${R_GDP})^{eY}'
                f'*{ratio}^{eU}*{ratio}^({eF}*(1+{eU}))')


def group_cols(g):
    start = COL_G0 + (g - 1) * GW
    return start, start + 1, list(range(start + 2, start + GW))


def build_mitigation(wb):
    ws = wb.create_sheet('Mitigation', 1)
    groups = [s[0] for s in SCENARIOS]
    last = COL_G0 + len(groups) * GW - 1
    # Row 1 headers (NORMS 1.1), row 2 title band.
    for j, h in enumerate(['', 'Fuel', 'Sector', 'Description', 'Unit', 'Source', 'Input Code',
                           'Output Code', 'Note', 'Helper (group)'] + KEYS, 1):
        put(ws, f'{L(j)}1', h, FONT_B)
    for s in range(N_SETS):
        for k in range(N_SLOTS):
            put(ws, f'{L(COL_P0 + s * N_SLOTS + k)}1', f'S{s + 1} P{k + 1}', FONT_B)
    title(ws, 'Mitigation module - price to fuel use (copy-pasteable, design 2)', last)
    band(ws, 3, 'Scenario and macro inputs (one column group per scenario; copy a whole group to add a scenario)', last)
    labels = {R_SCEN: ('Scenario number', 'no.'), R_YEAR: ('Year', 'year'), R_NAME: ('Scenario name', ''),
              R_GDP: ('Real GDP growth', '%'), R_CP: ('Carbon price', '$/tCO2')}
    for r, (lab, unit) in labels.items():
        put(ws, f'D{r}', lab, FONT_B)
        put(ws, f'E{r}', unit)
    put(ws, f'F{R_GDP}', 'WEO (by country)')
    put(ws, f'F{R_CP}', 'Scenario input')
    put(ws, f'H{R_GDP}', 'mit.gdp.pos.pct')
    put(ws, f'H{R_CP}', 'mit.cp')

    for g, gname, cp in SCENARIOS:
        lab, base, years = group_cols(g)
        put(ws, f'{L(lab)}1', f'Scenario group {g}', FONT_B)
        put(ws, f'{L(lab)}{R_SCEN}', g, FONT_B, F_INPUT)
        put(ws, f'{L(lab)}{R_YEAR}', 'base ->', FONT)
        put(ws, f'{L(base)}{R_NAME}', f'=INDEX(Settings!$C$10:$C$13,{L(base)}${R_SCEN})', FONT_B)
        for c in [base] + years:
            col, prev = L(c), L(c - 1)
            put(ws, f'{col}{R_SCEN}', f'={prev}{R_SCEN}', FONT_B, fmt='0')
            put(ws, f'{col}{R_YEAR}', '=Settings!$C$7' if c == base else f'={prev}{R_YEAR}+1', FONT_B, fmt='0')
            put(ws, f'{col}{R_GDP}', (f'=INDEX(WEO!$F$4:$AZ$1000,MATCH(Settings!$C$4&"|gdp_growth",WEO!$A$4:$A$1000,0),'
                                      f'MATCH({col}${R_YEAR},WEO!$F$3:$AZ$3,0))'), fmt='0.0%')
            year = BASE_YEAR if c == base else YEARS[c - years[0]]
            put(ws, f'{col}{R_CP}', cp.get(year, 0), FONT, F_INPUT, '0.0')
            if c == base:
                for r in (R_SCEN, R_YEAR, R_GDP):
                    ws[f'{col}{r}'].fill = F_BASE

    # Blocks
    for block, (b0, name, code, unit, slots) in BLOCKS.items():
        band(ws, b0, name, last)
        for s in range(N_SETS):
            for k in range(N_SLOTS):
                put(ws, f'{L(COL_P0 + s * N_SLOTS + k)}{b0}', slots[k], FONT_B, F_BAND)
        for i in range(NROW):
            r = b0 + 1 + i
            sub, fu = SUBSECTORS[i // NF][0], FUELS[i % NF][0]
            put(ws, f'B{r}', fu)
            put(ws, f'C{r}', sub)
            put(ws, f'D{r}', f'{SHORT[block]} | {SNAMES[sub]} | {FNAMES[fu]}')
            put(ws, f'E{r}', unit)
            put(ws, f'F{r}', {'pre': 'Prices_dom (base year), then growth', 'tax': 'Prices_dom, EF_GHG, Mapping',
                              'post': 'Calculation', 'use': 'EnergyCons (base year), Elasticities'}[block])
            put(ws, f'G{r}', {'pre': f'mit.sp.{fu}', 'tax': f'mit.rp.{fu}', 'post': '',
                              'use': f'mit.enc.{fu}.{sub}'}[block], fill_=F_CODE)
            put(ws, f'H{r}', f'mit.{code}.{fu}.{sub}', fill_=F_CODE)
            put(ws, f'J{r}', f'=INDEX(Mapping!$C$5:$C$20,MATCH($C{r},Mapping!$A$5:$A$20,0))')
            keys = [f'=INDEX(Mapping!$D$5:$D$20,MATCH($C{r},Mapping!$A$5:$A$20,0))',
                    f'=INDEX(Mapping!$C$24:$C$31,MATCH($B{r},Mapping!$A$24:$A$31,0))',
                    f'=INDEX(Mapping!$D$24:$D$31,MATCH($B{r},Mapping!$A$24:$A$31,0))',
                    (f'=$B{r}&"."&IF(INDEX(Mapping!$E$24:$E$31,MATCH($B{r},Mapping!$A$24:$A$31,0))="sub",'
                     f'INDEX(Mapping!$F$5:$F$20,MATCH($C{r},Mapping!$A$5:$A$20,0)),"all")'),
                    f'=INDEX(Mapping!$E$5:$E$20,MATCH($C{r},Mapping!$A$5:$A$20,0))',
                    f'=INDEX(Mapping!$G$24:$G$31,MATCH($B{r},Mapping!$A$24:$A$31,0))']
            for j, f in enumerate(keys):
                put(ws, f'{L(COL_KEY0 + j)}{r}', f)
            # Parameter sets: set 1 holds the lookups/inputs, sets 2..4 link to set 1 (overwrite to differ).
            for s in range(N_SETS):
                for k in range(1, N_SLOTS + 1):
                    cell = f'{L(COL_P0 + s * N_SLOTS + k - 1)}{r}'
                    if k not in USED_SLOTS[block]:
                        put(ws, cell, None, fill_=F_UNUSED)
                        continue
                    if s == 0:
                        val, is_input = set1_formula(block, k, r)
                        put(ws, cell, val, fill_=F_INPUT if is_input else None, fmt='0.000')
                    else:
                        put(ws, cell, f'={L(COL_P0 + k - 1)}{r}', fmt='0.000')
            # Scenario groups
            fmt = '#,##0.0' if block == 'use' else '0.000'
            for g in groups:
                _, base, years = group_cols(g)
                for c in [base] + years:
                    put(ws, f'{L(c)}{r}', group_formula(block, r, L(c), L(c - 1), c == base),
                        fill_=F_BASE if c == base else None, fmt=fmt)

    # Totals (fuel use)
    band(ws, B_TOT, 'Totals - fuel use (ktoe)', last)
    u0, u1 = B_USE + 1, B_USE + NROW
    r_sub0, r_fuel0 = B_TOT + 1, B_TOT + NS + 2
    r_tot = r_fuel0 + NF + 1
    rows = ([(r_sub0 + i, 'all', s[0], f'Total fuel use | {s[1]}', 'C') for i, s in enumerate(SUBSECTORS)]
            + [(r_fuel0 + i, f[0], 'all', f'Total fuel use | {f[1]}', 'B') for i, f in enumerate(FUELS)])
    for r, fu, sub, desc, keycol in rows:
        put(ws, f'B{r}', fu)
        put(ws, f'C{r}', sub)
        put(ws, f'D{r}', desc)
        put(ws, f'E{r}', 'ktoe')
        put(ws, f'H{r}', f'mit.enc.{fu}.{sub}', fill_=F_CODE)
        for g in groups:
            _, base, years = group_cols(g)
            for c in [base] + years:
                col = L(c)
                put(ws, f'{col}{r}', f'=SUMIF(${keycol}${u0}:${keycol}${u1},${keycol}{r},{col}${u0}:{col}${u1})',
                    fill_=F_BASE if c == base else None, fmt='#,##0.0')
    extra = [(r_tot, 'Total fuel use (all subsectors, all fuels)', 'ktoe', 'mit.enc.all.all',
              lambda col: f'=SUM({col}{r_sub0}:{col}{r_sub0 + NS - 1})', '#,##0.0', None),
             (r_tot + 1, 'Check: sum by fuel - sum by subsector (should be 0)', 'ktoe', '',
              lambda col: f'=SUM({col}{r_fuel0}:{col}{r_fuel0 + NF - 1})-{col}{r_tot}', '0.000', F_CHECK),
             (r_tot + 3, 'Total fuel use, scenario group 1 (same year)', 'ktoe', '',
              lambda col: (f'=INDEX(${L(COL_G0)}{r_tot}:${L(last + 10 * GW)}{r_tot},'
                           f'MATCH({col}${R_YEAR},${L(COL_G0)}${R_YEAR}:${L(last + 10 * GW)}${R_YEAR},0))'),
              '#,##0.0', None),
             (r_tot + 4, 'Change in total fuel use vs scenario group 1', '%', 'mit.enc.all.all.pct',
              lambda col: f'={col}{r_tot}/{col}{r_tot + 3}-1', '0.00%', None)]
    for r, desc, unit, code, f, fmt, fl in extra:
        put(ws, f'D{r}', desc, FONT_B)
        put(ws, f'E{r}', unit)
        if code:
            put(ws, f'H{r}', code, fill_=F_CODE)
        for g in groups:
            _, base, years = group_cols(g)
            for c in [base] + years:
                put(ws, f'{L(c)}{r}', f(L(c)), fill_=fl or (F_BASE if c == base else None), fmt=fmt)

    # Layout
    ws.freeze_panes = f'{L(COL_KEY0)}10'
    widths = {'A': 3, 'B': 5, 'C': 5, 'D': 40, 'E': 7, 'F': 14, 'G': 13, 'H': 15, 'I': 6, 'J': 7}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    for j in range(COL_KEY0, COL_P0):
        ws.column_dimensions[L(j)].width = 8
    for j in range(COL_P0, COL_P1 + 1):
        ws.column_dimensions[L(j)].width = 8
        if j > COL_P0 + N_SLOTS - 1:
            ws.column_dimensions[L(j)].outlineLevel = 1  # parameter sets 2..4 can be collapsed
    for g in groups:
        lab, _, _ = group_cols(g)
        ws.column_dimensions[L(lab)].width = 7
        for c in range(lab, lab + GW):
            ws.cell(R_SCEN, c).border = Border(top=THIN)
        for c in range(lab + 1, lab + GW):
            ws.column_dimensions[L(c)].width = 9
    return last, r_tot


# ---------------------------------------------------------------- ReadMe
def build_readme(wb, last, r_tot):
    ws = wb.active
    ws.title = 'ReadMe'
    title(ws, f'CPAT mitigation module - copy-pasteable prototype v{VERSION}', 3)
    g1, g2 = group_cols(1), group_cols(2)
    lines = [
        ('Purpose', 'Auditable, copy-pasteable replacement of the CPAT mitigation price -> fuel-use chain. '
                    'Every block has one formula (same R1C1 text) across all its projection cells; a whole '
                    'scenario group copies to the next group position and keeps working.'),
        ('Status', f'v{VERSION} prototype. Price -> fuel use only (no emissions, power or revenue yet).'),
        ('Layout (design 2)', 'Columns: descriptors A-J (NORMS 1.1), keys K-P, parameter sets Q-AJ, then one '
                              f'column group per scenario ({GW} columns: label, base year, '
                              f'{YEARS[0]}-{YEARS[-1]}). Rows: macro rows 4-8, then blocks.'),
        ('Nesting', 'Across: scenario (outer) > year (inner). Down: variable block (outer) > subsector > fuel '
                    f'(inner). {NS} CPAT energy-use subsectors x {NF} fuels = {NROW} rows per block.'),
        ('Blocks', f'Pre-tax price (row {B_PRE}), Tax (row {B_TAX}), Post-tax price (row {B_POST}), '
                   f'Fuel use (row {B_USE}), Totals (row {B_TOT}).'),
        ('Pre-tax price', 'base year: Prices_dom supply cost (mit.sp) / GJ per price unit; then '
                          'p(t) = p(t-1) x (1 + P1 growth).'),
        ('Tax', 'tax(t) = P1 base tax + carbon price(t) x P2 EF x P3 coverage. Base tax = (retail - supply '
                'cost) / GJ per unit, so it includes existing excise, subsidies and VAT.'),
        ('Post-tax price', 'pre-tax + tax (same row in the blocks above, fixed offsets).'),
        ('Fuel use', 'F(t) = F(t-1) x (1/(1+alpha))^(1+eps_U) x (1+g)^eps_Y x (p(t)/p(t-1))^eps_U x '
                     '(p(t)/p(t-1))^(eps_F(1+eps_U)) - CPAT documentation 3.3.3, without the Covid factor and '
                     'shadow price.'),
        ('Parameters', f'Each row has {N_SETS} parameter sets of {N_SLOTS} slots (P5 = buffer). Set 1 holds the '
                       'lookups; sets 2-4 link to set 1 - overwrite a cell to make a parameter scenario-specific. '
                       'Formulas pick the set with INDEX(row, (scenario number - 1) x 5 + slot).'),
        ('Lookups by country', 'Settings!C4 selects the country. Elasticities (via income group) and prices '
                               'cover all CPAT countries; base-year energy use, GDP growth and EFs are Egypt only '
                               'so far (other countries return #N/A until their rows are added).'),
        ('Add a scenario', f'(1) Copy columns {L(g2[0])}:{L(g2[2][-1])} (a whole group) and paste at the next '
                           f'free group position (column {L(g2[0] + GW)}). (2) Type the new scenario number '
                           f'in its label cell (row {R_SCEN}). (3) Edit the carbon prices in row {R_CP}. '
                           f'(4) Optionally overwrite parameter set N on the left. Up to {N_SETS} scenarios fit '
                           'without inserting parameter columns.'),
        ('Colours', 'Green = input; tan = base-year column; light blue = codes; grey = unused parameter slot; '
                    'yellow = check (NORMS section 2).'),
        ('Deviation from NORMS 1.1', 'Years do not start in column K: keys, parameter sets and scenario '
                                     'groups follow the descriptor columns (design 2).'),
        ('Data', 'Data sheets: Countries, Elasticities, Prices_dom, EnergyCons, WEO, EF_GHG; mappings in '
                 'Mapping. Built by build_v0_1.py from data/*.csv (extract_data_v0_1.py).'),
        ('Checks', f'Row {r_tot + 1} (sum by fuel - sum by subsector) should be 0. check_v0_1.py tests formula '
                   'uniformity, the scenario-copy behaviour and an independent recomputation.'),
    ]
    for i, (k, v) in enumerate(lines, 4):
        put(ws, f'A{i}', k, FONT_B)
        c = put(ws, f'B{i}', v)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        ws[f'A{i}'].alignment = Alignment(vertical='top')
    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['B'].width = 120


def main():
    wb = Workbook()
    build_settings(wb)
    build_mapping(wb)
    last, r_tot = build_mitigation(wb)
    div = wb.create_sheet('DATA->')
    put(div, 'A1', 'Data sheets follow', FONT_B)
    build_data_sheets(wb)
    build_readme(wb, last, r_tot)
    order = ['ReadMe', 'Settings', 'Mitigation', 'Mapping', 'DATA->', 'Countries', 'Elasticities',
             'Prices_dom', 'EnergyCons', 'WEO', 'EF_GHG']
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    wb.save(OUT)
    print('Saved', OUT)


if __name__ == '__main__':
    main()
