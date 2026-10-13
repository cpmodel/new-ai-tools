"""Build CPAT_Mitigation_CopyPaste_v0.8.xlsx (design 2: time across, scenarios as column groups).

Price -> fuel-use prototype of the CPAT mitigation module with fully copy-pasteable formulas.

v0.7 layout follows the legacy CPAT Mitigation sheet: numbered sections (1 Policies, 3 Power sector, 5 Transport,
6 Buildings, 7 Industrial, 8 Other energy use, 11 Results). Within a sector section each subsector has a heading
line (its total fuel use) and its variables, each over the 8 fuels:
    sp (pre-tax price), ctxnew (new carbon tax), ntx (new excise: fuel price reform), nce (total new policy),
    tax (base tax + nce), atp (after-tax price), shp (shadow price on the efficiency margin), ener (fuel use).
v0.8: the shadow price enters the efficiency margin of the fuel-use equation (legacy / cpat_coded ec.py):
    ((atp + shp) / (atp_prev + shp_prev)) ^ eps_F, with shp = sector shadow price x EF x share impacting
    efficiency (ssc = feebate coverage x efficiency-margin adjustment, 1.0 for feebates). All policy paths
    continue linearly after their target year (one default; the carbon price keeps its MTInputs switch).
The distance between variables is the same in every subsector, so every variable has one formula everywhere.

- Data input is a separate step: sheet Inputs (one row per fuel|subsector) holds the mappings and parameter
  lookups; Mitigation reads it in the hidden parameter columns D:G and the base-year column. Section 1 reads the
  scenario inputs from MTInputs (one Used-for-calculation column per scenario) by MTInputs row (column D) and the
  scenario number (row 5), then computes the policy paths (carbon price, fuel price reform, feebates).
- Calculation cells contain no searching lookups; a few pick a Section-1 row by position (INDEX(range, position)),
  the position being a data-step parameter in D:G.
- 2023-2034 plain formulas; 2035 calls named LAMBDAs (PRETAX, TAX, POSTTAX, FUELUSE) for sp, tax, atp, ener.
- Each scenario group starts with a code column building country.mit.<var>.<subsector>.<fuel>.<suffix>.<scen>.

Reads data/*.csv (extract_data_v0_1.py) and templates/MTInputs_template.xlsx. Writes formulas only.

    python build_v0_8.py
"""
import csv
import os
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName

VERSION = '0.8'
HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, 'data')
OUT = os.path.join(HERE, f'CPAT_Mitigation_CopyPaste_v{VERSION}.xlsx')

# ---------------------------------------------------------------- dimensions
SUBSECTORS = [  # code, name, group, elasticity sector, EF sector, coal/gas price sector, feebate sector
    ('rod', 'Road', 'tra', 'tra', 'rod', 'res', 'tra'),
    ('ral', 'Rail', 'tra', 'tra', 'rod', 'ind', 'tra'),
    ('avi', 'Domestic aviation', 'tra', 'tra', 'rod', 'ind', 'tra'),
    ('nav', 'Domestic navigation', 'tra', 'tra', 'rod', 'ind', 'tra'),
    ('res', 'Residential', 'bld', 'res', 'res', 'res', 'res'),
    ('foo', 'Food & forestry', 'bld', 'ind', 'res', 'ind', 'res'),
    ('srv', 'Services (public & private)', 'bld', 'srv', 'res', 'ind', 'res'),
    ('mch', 'Mining & chemicals', 'ind', 'ind', 'ind', 'ind', 'ind'),
    ('irn', 'Iron & steel', 'ind', 'ind', 'ind', 'ind', 'ind'),
    ('nfm', 'Other metals', 'ind', 'ind', 'ind', 'ind', 'ind'),
    ('mac', 'Machinery', 'ind', 'ind', 'ind', 'ind', 'ind'),
    ('cem', 'Cement', 'ind', 'ind', 'ind', 'ind', 'ind'),
    ('omn', 'Other manufacturing', 'ind', 'ind', 'ind', 'ind', 'ind'),
    ('cst', 'Construction', 'ind', 'ind', 'ind', 'ind', 'ind'),
    ('ftr', 'Fuel transformation', 'ind', 'ind', 'ind', 'ind', 'ind'),
    ('oen', 'Other energy use', 'oen', 'ind', 'ind', 'ind', 'ind'),
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
SNAMES = {s[0]: s[1] for s in SUBSECTORS}
FNAMES = {f[0]: f[1] for f in FUELS}
PAIRS = [(s[0], f[0]) for s in SUBSECTORS for f in FUELS]
SECTIONS = [  # legacy section number, group code, title
    (5, 'tra', 'Transport sector'), (6, 'bld', 'Buildings sector'), (7, 'ind', 'Industrial sector'),
    (8, 'oen', 'Other energy use')]
EXTRA_LABELS = [('all', 'All subsectors'), ('pow', 'Power'), ('tra', 'Transport'), ('bld', 'Buildings'),
                ('ind', 'Industry')]                      # labels only (totals, policy rows)
COV_SECTORS = ['pow'] + [s[0] for s in SUBSECTORS]        # order of the MTInputs coverage switches (17)
PRICE_FUELS = [  # price code, label, MTInputs rows: starting increase, final increase (fuel price reform)
    ('coa.pow', 'Coal, Power', 142, 156), ('coa.res', 'Coal, Residential', 143, 157),
    ('coa.ind', 'Coal, Industry', 144, 158), ('nga.pow', 'Natural gas, Power', 145, 159),
    ('nga.res', 'Natural gas, Residential', 146, 160), ('nga.ind', 'Natural gas, Industry', 147, 161),
    ('ecy.res', 'Electricity, Residential', 148, 162), ('ecy.ind', 'Electricity, Industry', 149, 163),
    ('gso.all', 'Gasoline, All', 150, 164), ('die.all', 'Diesel, All', 151, 165),
    ('lpg.all', 'LPG, All', 152, 166), ('ker.all', 'Kerosene, All', 153, 167),
    ('oop.all', 'Other oil products, All', 154, 168), ('bio.all', 'Biomass, All', 155, 169)]
FEEBATE_SECTORS = [  # code, MTInputs names of the starting and target rate
    ('pow', 'D_Feb_Level_Start_Pow', 'D_Feb_Level_Target_Pow'),
    ('tra', 'D_Feb_Level_Start_Trans', 'D_Feb_Level_Target_Trans'),
    ('res', 'D_Feb_Level_Start_Res', 'D_Feb_Level_Target_Res'),
    ('ind', 'D_Feb_Level_Start_Ind', 'D_Feb_Level_Target_Ind')]
MT_ROWS_FIXED = {'fpr.yr0': 140, 'fpr.yr1': 141, 'fbcov0': 66}   # MTInputs rows without NameOfParameter

BASE_YEAR, LAST_YEAR = 2022, 2035
YEARS = list(range(BASE_YEAR + 1, LAST_YEAR + 1))
SCENARIOS = [(1, 'Baseline (no carbon price)', {}),
             (2, 'Carbon price $20/tCO2 from 2027', {y: 20 for y in YEARS if y >= 2027})]
MT_SCENARIO_INPUTS = {
    1: {'CPIntro': 2027, 'CPLevelStart': 0, 'CPLevelTarget': 0, 'CPOutro': 2030, 'MCovOen': False},
    2: {'CPIntro': 2027, 'CPLevelStart': 20, 'CPLevelTarget': 20, 'CPOutro': 2030, 'MCovOen': False},
}
MT_TEMPLATE = next(os.path.join(d, 'templates', 'MTInputs_template.xlsx')      # repo templates folder,
                   for d in [os.path.abspath(os.path.join(HERE, *['..'] * k)) for k in range(2, 5)]   # also from Old/
                   if os.path.exists(os.path.join(d, 'templates', 'MTInputs_template.xlsx')))
MT_COL0 = 10                            # J: first scenario column on MTInputs
MT_ROW_SCEN, MT_ROW_NAME = 5, 6
MT_LAST = 'AZ'
CP_PARAMS = ['CPIntro', 'CPLevelStart', 'CPLevelTarget', 'CPOutro', 'ExtendCarbonPriceBeyondOutro']

# ---------------------------------------------------------------- Mapping layout
MAP_SUB0 = 5                            # subsector table rows 5-20, extra label rows after
MAP_SUB1 = MAP_SUB0 + NS - 1
MAP_SUBLAB1 = MAP_SUB1 + len(EXTRA_LABELS)
MAP_FUEL0 = MAP_SUBLAB1 + 4             # fuel table
MAP_FUEL1 = MAP_FUEL0 + NF - 1
MAP_FUELLAB1 = MAP_FUEL1 + 2            # + 'all' and 'ecy' label rows
MAP_PF0 = MAP_FUELLAB1 + 4              # price-fuel table (fuel price reform)
MAP_PF1 = MAP_PF0 + len(PRICE_FUELS) - 1
MAP_FB0 = MAP_PF1 + 4                   # feebate sectors
MAP_FB1 = MAP_FB0 + len(FEEBATE_SECTORS) - 1

# ---------------------------------------------------------------- Inputs sheet (one row per fuel|subsector)
IN_R0 = 5
IN_R1 = IN_R0 + len(PAIRS) - 1
INPUT_COLS = [
    ('Key (fuel|subsector)', 'key'), ('Fuel', 'fuel'), ('Subsector', 'sub'),
    ('Elasticity sector', 'esec'), ('Elasticity fuel', 'efuel'), ('AEEI fuel', 'afuel'),
    ('Price code', 'pcode'), ('EF sector', 'efsec'), ('GJ per price unit', 'gj'),
    ('Fuel position (coverage)', 'fpos'), ('Sector position (coverage)', 'spos'),
    ('Price-fuel position (fuel price reform)', 'pfpos'), ('Feebate sector position', 'fbpos'),
    ('eps_Y GDP', 'eY'), ('eps_U usage', 'eU'), ('eps_F efficiency', 'eF'), ('alpha AEEI', 'a'),
    ('Pre-tax price growth', 'grow'), ('Base pre-tax price ($/GJ)', 'p0'), ('Base tax ($/GJ)', 't0'),
    ('EF (tCO2/GJ)', 'ef'), ('Base-year fuel use (ktoe)', 'f0'),
]
IC = {code: L(j) for j, (_, code) in enumerate(INPUT_COLS, 1)}

# ---------------------------------------------------------------- Mitigation layout
PARAM_COLS = ['D', 'E', 'F', 'G']       # hidden by default
DESC = {'H': 'Description', 'I': 'Unit', 'J': 'Source'}
COL_G0 = 11                             # K: code column of scenario 1; base year 2022 in L (as legacy)
NY = 1 + len(YEARS)
GW = NY + 1
R_YEAR, R_SCEN, R_NAME, R_GDP = 4, 5, 6, 7

# Section 1 (policies): row plan built in order; each entry = (key, var, fuel, sector, kind, mt_row or None)
#   kind: 'mt' = lookup of MTInputs row (column D) for this column's scenario; 'calc' = path formula;
#         'const' = typed constant (no MTInputs switch)
POL = []


def _pol(key, var, fuel='', sector='', kind='mt', mt=None):
    POL.append((key, var, fuel, sector, kind, mt))


def _mt_rows():
    import openpyxl
    src = openpyxl.load_workbook(MT_TEMPLATE)['MTInputs']
    return {src.cell(r, 8).value: r for r in range(8, src.max_row + 1) if src.cell(r, 8).value}


MT_NAMES = _mt_rows()
_pol('cptraj', 'cptraj', kind='calc')                                     # summary line under the band
for p in CP_PARAMS:
    _pol(p, p, mt=MT_NAMES[p])
for f, *_ in FUELS:
    if f == 'bio':
        _pol('fc.bio', 'ctcov', 'bio', 'all', kind='const')                # no MTInputs switch: TRUE
    else:
        _pol(f'fc.{f}', 'ctcov', f, 'all', mt=MT_NAMES['MCov' + f.capitalize()])
for s in COV_SECTORS:
    _pol(f'sc.{s}', 'ctcov', 'all', s, mt=MT_NAMES['MCov' + s.capitalize()])
_pol('fpr.yr0', 'fpr.yr0', mt=MT_ROWS_FIXED['fpr.yr0'])
_pol('fpr.yr1', 'fpr.yr1', mt=MT_ROWS_FIXED['fpr.yr1'])
for pc, _lab, r0, r1 in PRICE_FUELS:
    _pol(f'fprs.{pc}', 'fpr.s', pc.split('.')[0], pc.split('.')[1], mt=r0)
for pc, _lab, r0, r1 in PRICE_FUELS:
    _pol(f'fprf.{pc}', 'fpr.f', pc.split('.')[0], pc.split('.')[1], mt=r1)
for pc, *_ in PRICE_FUELS:
    _pol(f'fpr.{pc}', 'fpr', pc.split('.')[0], pc.split('.')[1], kind='calc')
_pol('fb.yr0', 'fb.yr0', mt=MT_NAMES['D_FeebateIntro'])
_pol('fb.yr1', 'fb.yr1', mt=MT_NAMES['D_FeebateOutro'])
for s, n0, n1 in FEEBATE_SECTORS:
    _pol(f'fbs.{s}', 'fb.s', 'all', s, mt=MT_NAMES[n0])
for s, n0, n1 in FEEBATE_SECTORS:
    _pol(f'fbt.{s}', 'fb.t', 'all', s, mt=MT_NAMES[n1])
for s, *_ in FEEBATE_SECTORS:
    _pol(f'fb.{s}', 'fb', 'all', s, kind='calc')
for k, s in enumerate(COV_SECTORS):
    _pol(f'fbc.{s}', 'fbcov', 'all', s, mt=MT_ROWS_FIXED['fbcov0'] + k)
_pol('ssc.adj', 'ssc.adj', kind='const')                                   # feebates: 1.0 (cpat_coded default)
for s, *_ in FEEBATE_SECTORS:
    _pol(f'shps.{s}', 'shps', 'all', s, kind='calc')                        # shadow price by sector, $/tCO2
for s in COV_SECTORS:
    _pol(f'ssc.{s}', 'ssc', 'all', s, kind='calc')                          # share impacting efficiency
CONST_VALUES = {'fc.bio': True, 'ssc.adj': 1}
SUBHEADS = {'cptraj': None, CP_PARAMS[0]: 'Carbon tax (MTInputs rows 18-21, 265)',
            'fc.coa': 'Carbon tax coverage: fuels (MTInputs rows 23-29)',
            'sc.pow': 'Carbon tax coverage: sectors (MTInputs rows 31-47)',
            'fpr.yr0': 'Fuel price reform (MTInputs rows 140-169; increases in price units)',
            'fb.yr0': 'Feebates (MTInputs rows 53-64; rates in USD/tCO2)',
            'fbc.pow': 'Feebates sector coverage (MTInputs rows 66-82)',
            'ssc.adj': 'Shadow prices: by sector ($/tCO2, legacy rows 2345-2349) and share impacting efficiency '
                       '(legacy rows 2402-2419)'}

B_POL = 9                               # band: 1. Policies
R_POL = {}                              # key -> row
_r = B_POL + 1
for key, *_ in POL:
    if SUBHEADS.get(key):
        _r += 1                         # sub-heading row before this entry
    R_POL[key] = _r
    _r += 1
R_CP = R_POL['cptraj']
B_POW = _r + 1                          # 3. Power sector (placeholder)
R_POW_NOTE = B_POW + 1
NVAR = 8
VARS = ['sp', 'ctxnew', 'ntx', 'nce', 'tax', 'atp', 'shp', 'ener']   # per subsector, each over the 8 fuels
VSTEP = NF + 1                          # 8 fuel rows + 1 blank row
SB = 1 + NVAR * VSTEP                   # subsector block: heading line + variables
VOFF = {v: 1 + k * VSTEP for k, v in enumerate(VARS)}   # first row of each variable relative to the heading
SEC_BAND, SEC_SUM, SUB_HEAD = {}, {}, {}
_r = R_POW_NOTE + 2
for num_, grp, title_ in SECTIONS:
    SEC_BAND[grp] = _r
    SEC_SUM[grp] = _r + 1
    _r += 2
    for s in [x for x in SUBSECTORS if x[2] == grp]:
        SUB_HEAD[s[0]] = _r
        _r += SB
    _r += 1
S_FIRST, S_LAST = SEC_BAND['tra'], _r - 2          # rows spanned by the sector sections
B_RES = _r                              # 11. Results - energy consumption
R_TOTAL = B_RES + 1
R_SUB0 = R_TOTAL + 2
R_FUEL0 = R_SUB0 + NS + 1
R_CHK = R_FUEL0 + NF + 1
R_REF = R_CHK + 2
R_PCT = R_REF + 1


def var_rows(var):
    """All (row, subsector, fuel) of a variable across the sector sections."""
    return [(SUB_HEAD[s] + VOFF[var] + i, s, f) for s, *_ in SUBSECTORS for i, (f, *_r2) in enumerate(FUELS)]


# Hidden parameter columns D:G for each variable (Inputs columns), and their labels.
VAR_PARAMS = {'sp': ['grow'], 'ctxnew': ['ef', 'fpos', 'spos'], 'ntx': ['pfpos', 'gj'], 'nce': [],
              'tax': ['t0'], 'atp': [], 'shp': ['fbpos', 'ef', 'spos'], 'ener': ['eY', 'eU', 'eF', 'a']}
PARAM_LABELS = {'grow': 'growth', 't0': 'base tax', 'ef': 'EF', 'fpos': 'fuel position', 'spos': 'sector position',
                'pfpos': 'price-fuel position', 'gj': 'GJ per unit', 'fbpos': 'feebate sector position',
                'eY': 'eps_Y', 'eU': 'eps_U', 'eF': 'eps_F', 'a': 'alpha'}

# Variable dictionary (sheet Variables): code, label, unit, code suffix, source, legacy reference, param labels
VARIABLES = [
    ('gdp.pos.pct', 'Real GDP growth', '%', '', 'WEO (lookup by country)', 'legacy Mitigation row 2446'),
] + [(p, '=MT', '=MT', '', 'MTInputs (scenario column)', f'MTInputs row {MT_NAMES[p]}') for p in CP_PARAMS] + [
    ('cptraj', 'Carbon price trajectory used', '$/tCO2 real', '', 'Calculation (legacy rows 2248, 2251)',
     'legacy Mitigation row 2251: egy.mit.cptraj.1'),
    ('ctcov', 'Carbon tax coverage (Apply tax?)', 'switch', '', 'MTInputs (scenario column)',
     'MTInputs MCov* rows 23-47; biomass has no switch (EF 0)'),
    ('fpr.yr0', 'Fuel price reform: starting year', 'year', '', 'MTInputs (scenario column)', 'MTInputs row 140'),
    ('fpr.yr1', 'Fuel price reform: target year', 'year', '', 'MTInputs (scenario column)', 'MTInputs row 141'),
    ('fpr.s', 'Fuel price reform: starting price increase', 'price unit', '', 'MTInputs (scenario column)',
     'MTInputs rows 142-155'),
    ('fpr.f', 'Fuel price reform: final price increase', 'price unit', '', 'MTInputs (scenario column)',
     'MTInputs rows 156-169'),
    ('fpr', 'Fuel price reform: price increase path', 'price unit', '', 'Calculation (assumed linear)', 'new'),
    ('fb.yr0', 'Feebates: start date', 'year', '', 'MTInputs (scenario column)', 'MTInputs row 53'),
    ('fb.yr1', 'Feebates: target date', 'year', '', 'MTInputs (scenario column)', 'MTInputs row 54'),
    ('fb.s', 'Feebates: starting rate', 'USD/tCO2', '', 'MTInputs (scenario column)', 'MTInputs rows 56-59'),
    ('fb.t', 'Feebates: target rate', 'USD/tCO2', '', 'MTInputs (scenario column)', 'MTInputs rows 61-64'),
    ('fb', 'Feebates: rate path', 'USD/tCO2', '', 'Calculation (assumed linear)', 'legacy rows 2000-2004'),
    ('fbcov', 'Feebates sector coverage (Apply?)', 'switch', '', 'MTInputs (scenario column)',
     'MTInputs rows 66-82'),
    ('ssc.adj', 'Efficiency-margin adjustment, feebates', 'share', '', 'cpat_coded default (feebates 1.0)',
     'MTInputs rows 247-250 hold the adjustments for regulations (not yet used)'),
    ('shps', 'Shadow price by sector (feebates; later + non-auctioned ETS, regulations)', 'USD/tCO2', '',
     'Calculation', 'legacy Mitigation rows 2345-2349'),
    ('ssc', 'Share of shadow price impacting efficiency', 'share', '', 'Feebate coverage x adjustment',
     'legacy Mitigation rows 2403-2419: egy.mit.ssc.rod.1'),
    ('sp', 'Pre-tax price (supply cost)', '$/GJ', 'a', 'Inputs (base year), then growth',
     'legacy Mitigation row 2498: egy.mit.sp.ind.coa.a.1', 'growth'),
    ('ctxnew', 'New carbon tax', '$/GJ', 'a', 'Carbon price x EF x fuel and sector coverage',
     'legacy Mitigation row 2510: egy.mit.ctxnew.ind.coa.a.1', 'EF', 'fuel position', 'sector position'),
    ('ntx', 'New excise tax (fuel price reform)', '$/GJ', 'a', 'Fuel price reform path / GJ per unit',
     'legacy Mitigation row 2512: egy.mit.ntx.ind.coa.a.1', 'price-fuel position', 'GJ per unit'),
    ('nce', 'Total new policy', '$/GJ', 'a', 'New carbon tax + new excise tax',
     'legacy Mitigation row 2514: egy.mit.nce.ind.coa.a.1'),
    ('tax', 'Tax (base tax + total new policy)', '$/GJ', 'a', 'Inputs (base tax); total new policy',
     'new: legacy splits taxes into txo, vat, ctx, ets, nce', 'base tax'),
    ('atp', 'After-tax price', '$/GJ', 'e', 'Pre-tax price + tax',
     'legacy Mitigation row 5416: egy.mit.atp.rod.gso.e.1 (legacy unit $/liter for liquids)'),
    ('shp', 'Shadow price on the efficiency margin', '$/GJ', '', 'Sector shadow price x EF x share',
     'legacy Mitigation row 2377: egy.mit.shp.coa.pow.1', 'feebate sector position', 'EF', 'sector position'),
    ('ener', 'Fuel use', 'ktoe', 'e', 'Inputs (base year), CPAT eq. 3.3.3',
     'legacy Mitigation row 5430: egy.mit.ener.rod.gso.e.1', 'eps_Y', 'eps_U', 'eps_F', 'alpha'),
    ('ener.chk', 'Check: totals by subsector and by fuel minus total fuel use (should be 0)', 'ktoe', '',
     'Check', 'new'),
    ('ener.ref', 'Fuel use, scenario 1 (same year)', 'ktoe', 'e', 'Lookup', 'new'),
    ('ener.pct', 'Change in fuel use vs scenario 1', '%', 'e', 'Calculation', 'new'),
]
VR0 = 5
VR1 = VR0 + 60

LAMBDAS = {
    'PRETAX': (['prev', 'growth'], 'prev*(1+growth)'),
    'TAX': (['base_tax', 'new_policy'], 'base_tax+new_policy'),
    'POSTTAX': (['pre_tax_p', 'tax_p'], 'pre_tax_p+tax_p'),
    'FUELUSE': (['f_prev', 'p_now', 'p_prev', 'shp_now', 'shp_prev', 'gdp_g', 'eps_y', 'eps_u', 'eps_f', 'alpha'],
                'f_prev*(1/(1+alpha))^(1+eps_u)*(1+gdp_g)^eps_y*(p_now/p_prev)^eps_u'
                '*((p_now+shp_now)/(p_prev+shp_prev))^(eps_f*(1+eps_u))'),
}
LAMBDA_NOTES = {
    'PRETAX': 'Pre-tax price: previous year x (1 + growth)',
    'TAX': 'Tax ($/GJ): base tax + total new policy (new carbon tax + new excise tax)',
    'POSTTAX': 'After-tax price: pre-tax price + tax',
    'FUELUSE': 'Fuel use, CPAT documentation 3.3.3, shadow price on the efficiency margin (no Covid factor)',
}


def lambda_xml(name):
    """Excel file encoding of a LAMBDA: _xlfn.LAMBDA(_xlpm.x, ..., body with _xlpm.x)."""
    params, body = LAMBDAS[name]
    for prm in sorted(params, key=len, reverse=True):
        body = re.sub(rf'(?<![A-Za-z_.]){prm}(?![A-Za-z_\d(])', f'_xlpm.{prm}', body)
    return '_xlfn.LAMBDA(' + ','.join(f'_xlpm.{prm}' for prm in params) + ',' + body + ')'

# ---------------------------------------------------------------- styles (NORMS section 2)
def fill(hex_):
    return PatternFill('solid', start_color=hex_, end_color=hex_)


F_TITLE, F_BAND, F_INPUT, F_CODE = fill('00B050'), fill('92D050'), fill('EBF1DE'), fill('DCE6F1')
F_BASE, F_UNUSED, F_CHECK, F_LAMBDA = fill('DDD9C4'), fill('F2F2F2'), fill('FFF2CC'), fill('FCE4D6')
FONT = Font(name='Arial', size=9)
FONT_B = Font(name='Arial', size=9, bold=True)
FONT_T = Font(name='Arial', size=14, bold=True, color='FFFFFF')
FONT_BAND = Font(name='Arial', size=9, bold=True, color='FFFFFF')
FONT_SUM = Font(name='Arial', size=9, bold=True)
FONT_LAMBDA = Font(name='Arial', size=9, color='9C0006')


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

def band(ws, row, text, width):
    for j in range(1, width + 1):
        ws.cell(row, j).fill = F_BAND
    put(ws, f'A{row}', text, FONT_BAND, F_BAND)

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

def build_data_sheets(wb):
    c = read_csv('countries.csv')
    data_sheet(wb, 'Countries', 'Countries: code, region and income group', c[0], c[1:],
               'legacy Elasticities rows 138-141 (country columns)')
    e = read_csv('elasticities.csv')
    data_sheet(wb, 'Elasticities', 'Elasticities by income group (legacy CPAT values, CPAT sector and fuel groups)',
               e[0], e[1:], "legacy Elasticities rows 143-282 ('Simple' option, income-group specific)")
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


def build_mtinputs(wb):
    """MTInputs: the template copied cell by cell (values and styles, rows 1-415, columns A:H), plus one
    'Used for calculation' column per scenario from column J."""
    from copy import copy
    import openpyxl
    src = openpyxl.load_workbook(MT_TEMPLATE)['MTInputs']
    ws = wb.create_sheet('MTInputs')
    for row in src.iter_rows(min_row=1, max_row=src.max_row, max_col=8):
        for c in row:
            d = ws.cell(c.row, c.column, c.value)
            if c.has_style:
                d.font, d.fill, d.border = copy(c.font), copy(c.fill), copy(c.border)
                d.alignment, d.number_format = copy(c.alignment), c.number_format
    for col, dim in src.column_dimensions.items():
        ws.column_dimensions[col].width = dim.width
        ws.column_dimensions[col].hidden = dim.hidden
    names = {src.cell(r, 8).value: r for r in range(8, src.max_row + 1) if src.cell(r, 8).value}
    for k in MT_SCENARIO_INPUTS[1]:
        assert k in names, k
    for g, gname, _ in SCENARIOS:
        col = MT_COL0 + g - 1
        cl = L(col)
        put(ws, f'{cl}{MT_ROW_SCEN}', g if g == 1 else f'={L(col - 1)}{MT_ROW_SCEN}+1', FONT_B,
            F_INPUT if g == 1 else None, '0')
        put(ws, f'{cl}{MT_ROW_NAME}', gname, FONT_B, F_INPUT)
        c = put(ws, f'{cl}7', 'Used for calculation', FONT_B, F_INPUT)
        c.alignment = Alignment(wrap_text=True)
        for r in range(8, src.max_row + 1):
            name, label, used = src.cell(r, 8).value, src.cell(r, 2).value, src.cell(r, 6).value
            if r == 10:
                val = f'={cl}${MT_ROW_SCEN}'          # Scenario ID
            elif r == 12:
                val = f'={cl}${MT_ROW_NAME}'          # Scenario name
            elif name in MT_SCENARIO_INPUTS[g]:
                val = MT_SCENARIO_INPUTS[g][name]
            else:
                val = used
            if val is None:
                continue
            fnt = Font(name='Arial', size=9, bold=True, color='C00000') if name in MT_SCENARIO_INPUTS[g] else FONT
            put(ws, f'{cl}{r}', val, fnt, F_INPUT)
        ws.column_dimensions[cl].width = 16
    ws.column_dimensions['I'].width = 3
    put(ws, f'{L(MT_COL0)}3', 'Scenario columns (Used for calculation): number in row 5, name in row 6. Red = '
                              'changed from the template value. Copy the last column one to the right to add a '
                              'scenario.', FONT_B)
    ws.freeze_panes = 'D8'


def fix_outline_levels(path):
    """openpyxl drops sheetFormatPr/@outlineLevelCol; Excel needs it to size the outline bar for the hidden
    D:G column group. Add it to the sheet that has a row outline (Mitigation)."""
    import shutil
    import zipfile
    tmp = path + '.tmp'
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('xl/worksheets/sheet') and re.search(rb'outlineLevelRow="\d+"', data):
                assert b'outlineLevelCol' not in data
                data = re.sub(rb'(outlineLevelRow="\d+")', rb'\1 outlineLevelCol="1"', data, count=1)
            zout.writestr(item, data)
    shutil.move(tmp, path)


# ---------------------------------------------------------------- Mapping, Settings, Inputs, Variables
def build_mapping(wb):
    ws = wb.create_sheet('Mapping')
    title(ws, 'Mapping: subsectors, fuels and policy keys to CPAT parameter groups', 8)
    put(ws, f'A{MAP_SUB0 - 2}', 'Subsectors', FONT_B)
    for j, h in enumerate(['Code', 'Name', 'Group', 'Elasticity sector', 'EF sector', 'Coal/gas price sector',
                           'Feebate sector', 'Coverage position'], 1):
        put(ws, f'{L(j)}{MAP_SUB0 - 1}', h, FONT_B, F_INPUT)
    for i, s in enumerate(SUBSECTORS):
        r = MAP_SUB0 + i
        for j, v in enumerate(list(s) + [COV_SECTORS.index(s[0]) + 1], 1):
            put(ws, f'{L(j)}{r}', v, FONT, F_INPUT if j > 2 else None)
    for i, (code, name) in enumerate(EXTRA_LABELS):
        put(ws, f'A{MAP_SUB1 + 1 + i}', code)
        put(ws, f'B{MAP_SUB1 + 1 + i}', name)
    units = {r[0]: float(r[1]) for r in read_csv('fuel_units.csv')[1:]}
    put(ws, f'A{MAP_FUEL0 - 2}', 'Fuels', FONT_B)
    for j, h in enumerate(['Code', 'Name', 'Elasticity fuel', 'AEEI fuel', 'Price sector rule', 'Price unit',
                           'GJ per price unit', 'Fuel position'], 1):
        put(ws, f'{L(j)}{MAP_FUEL0 - 1}', h, FONT_B, F_INPUT)
    for i, f in enumerate(FUELS):
        for j, v in enumerate(list(f) + [units[f[0]], i + 1], 1):
            put(ws, f'{L(j)}{MAP_FUEL0 + i}', v, FONT, F_INPUT if j > 2 else None)
    for i, (code, name) in enumerate([('all', 'All fuels'), ('ecy', 'Electricity')]):
        put(ws, f'A{MAP_FUEL1 + 1 + i}', code)
        put(ws, f'B{MAP_FUEL1 + 1 + i}', name)
    put(ws, f'A{MAP_PF0 - 2}', 'Price fuels (fuel price reform, MTInputs rows 142-169)', FONT_B)
    for j, h in enumerate(['Price code', 'Label', 'MTInputs row: starting increase', 'MTInputs row: final increase',
                           'Position'], 1):
        put(ws, f'{L(j)}{MAP_PF0 - 1}', h, FONT_B, F_INPUT)
    for i, (pc, lab, r0, r1) in enumerate(PRICE_FUELS):
        for j, v in enumerate([pc, lab, r0, r1, i + 1], 1):
            put(ws, f'{L(j)}{MAP_PF0 + i}', v)
    put(ws, f'A{MAP_FB0 - 2}', 'Feebate sectors (MTInputs rows 56-64)', FONT_B)
    for j, h in enumerate(['Code', 'Starting rate (NameOfParameter)', 'Target rate (NameOfParameter)',
                           'Position'], 1):
        put(ws, f'{L(j)}{MAP_FB0 - 1}', h, FONT_B, F_INPUT)
    for i, (s, n0, n1) in enumerate(FEEBATE_SECTORS):
        for j, v in enumerate([s, n0, n1, i + 1], 1):
            put(ws, f'{L(j)}{MAP_FB0 + i}', v)
    notes = ["Price sector rule: 'sub' = the subsector's coal/gas price sector (rod and res use residential "
             "coal/gas prices, as CPAT); 'all' = the fuel's single price.",
             'GJ per price unit: legacy Mitigation rows 980-1009 (conversion factor, volume unit to GJ).',
             'Coverage position: order of the MTInputs sector switches (power first, then the 16 subsectors).',
             'Feebate sector for food & forestry and services = residential (assumption; legacy groups them as '
             'Commercial).',
             f'Rows {MAP_SUB1 + 1}-{MAP_SUBLAB1} and {MAP_FUEL1 + 1}-{MAP_FUELLAB1}: labels only (totals, policy '
             'rows).']
    for i, n in enumerate(notes):
        put(ws, f'A{MAP_FB1 + 3 + i}', n)
    ws.column_dimensions['B'].width = 26
    for col in 'CDEFGH':
        ws.column_dimensions[col].width = 16


def build_settings(wb):
    ws = wb.create_sheet('Settings')
    title(ws, 'Settings', 6)
    rows = [(4, 'Country code', 'EGY', True),
            (5, 'Country name', '=INDEX(Countries!$B$4:$B$1000,MATCH($C$4,Countries!$A$4:$A$1000,0))', False),
            (6, 'Income group (selects elasticities)',
             '=INDEX(Countries!$D$4:$D$1000,MATCH($C$4,Countries!$A$4:$A$1000,0))', False),
            (7, 'Base year', BASE_YEAR, True)]
    for r, lab, val, inp in rows:
        put(ws, f'B{r}', lab)
        put(ws, f'C{r}', val, FONT, F_INPUT if inp else None)
    put(ws, 'B10', 'Version log', FONT_B)
    for j, h in enumerate(['Version', 'Date', 'Description', 'Max abs regression diff'], 2):
        put(ws, f'{L(j)}11', h, FONT_B, F_INPUT)
    log = [
        ('0.1', '2026-10-08', 'First build: design 2 price -> fuel-use prototype, Egypt base data, 2 scenarios.',
         'n/a (first version)'),
        ('0.2', '2026-10-08', 'Inputs sheet by fuel|subsector (lookups separated from formulas); 4 hidden parameter '
                              'columns D:G, global parameters; base year 2022 in column L; named LAMBDAs in 2035.',
         '0 (vs v0.1)'),
        ('0.3', '2026-10-08', 'Variable codes in column A; labels from Variables; code column per scenario with '
                              'full CPAT codes (auto-numbered).', '0 (vs v0.2)'),
        ('0.4', '2026-10-08', 'Row outline: blocks roll up.', '0 (vs v0.3)'),
        ('0.5', '2026-10-08', 'White band text; summary lines under bands.', '0 (vs v0.4)'),
        ('0.6', '2026-10-08', 'MTInputs sheet per scenario; carbon price from MTInputs.', '0 (vs v0.5)'),
        ('0.7', '2026-10-08', 'Legacy CPAT section layout (1 Policies, 3 Power, 5-8 sectors, 11 Results; subsector > '
                              'variable > fuel); policy wedges: new carbon tax with MTInputs coverage, new excise '
                              '(fuel price reform), total new policy; feebate shadow price (not yet used).',
         '0 (common codes vs v0.6)'),
        ('0.8', '2026-10-08', 'Shadow price on the efficiency margin of fuel use; shadow prices by sector and '
                              'share impacting efficiency in section 1; all policy paths continue linearly; '
                              'biomass carbon-tax coverage TRUE.', '0 (common codes vs v0.7, no feebate)')]
    for i, row in enumerate(log, 12):
        for j, v in enumerate(row, 2):
            put(ws, f'{L(j)}{i}', v)
    ws.column_dimensions['B'].width = 34
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 100
    ws.column_dimensions['E'].width = 26


def map_sub(col, r):
    return f'INDEX(Mapping!${col}${MAP_SUB0}:${col}${MAP_SUB1},MATCH($C{r},Mapping!$A${MAP_SUB0}:$A${MAP_SUB1},0))'


def map_fuel(col, r):
    return (f'INDEX(Mapping!${col}${MAP_FUEL0}:${col}${MAP_FUEL1},'
            f'MATCH($B{r},Mapping!$A${MAP_FUEL0}:$A${MAP_FUEL1},0))')


def build_inputs(wb):
    ws = wb.create_sheet('Inputs')
    title(ws, 'Inputs by fuel | subsector: every lookup used by Mitigation (data step, no calculations)',
          len(INPUT_COLS))
    put(ws, 'A3', 'One row per fuel | subsector (same taxonomy as Mitigation columns B:C). Mitigation reads these '
                  'columns with INDEX/MATCH on the key. Green = input.', FONT)
    for j, (h, _) in enumerate(INPUT_COLS, 1):
        c = put(ws, f'{L(j)}4', h, FONT_B, F_INPUT)
        c.alignment = Alignment(wrap_text=True, vertical='top')

    def elast(typ, fcol, r):
        return (f'=INDEX(Elasticities!$F$4:$I$200,MATCH("{typ}|"&${fcol}{r}&"|"&${IC["esec"]}{r},'
                f'Elasticities!$A$4:$A$200,0),MATCH(Settings!$C$6,Elasticities!$F$3:$I$3,0))')

    def price(prefix, r):
        return (f'INDEX(Prices_dom!$A$6:$CU$2000,MATCH(Settings!$C$4&Settings!$C$7,Prices_dom!$A$6:$A$2000,0),'
                f'MATCH("{prefix}."&${IC["pcode"]}{r},Prices_dom!$A$5:$CU$5,0))')

    for i, (s, f) in enumerate(PAIRS):
        r = IN_R0 + i
        vals = {
            'key': f'=$B{r}&"|"&$C{r}', 'fuel': f, 'sub': s,
            'esec': '=' + map_sub('D', r), 'efuel': '=' + map_fuel('C', r), 'afuel': '=' + map_fuel('D', r),
            'pcode': f'=$B{r}&"."&IF({map_fuel("E", r)}="sub",{map_sub("F", r)},"all")',
            'efsec': '=' + map_sub('E', r), 'gj': '=' + map_fuel('G', r),
            'fpos': '=' + map_fuel('H', r), 'spos': '=' + map_sub('H', r),
            'pfpos': (f'=INDEX(Mapping!$E${MAP_PF0}:$E${MAP_PF1},MATCH(${IC["pcode"]}{r},'
                      f'Mapping!$A${MAP_PF0}:$A${MAP_PF1},0))'),
            'fbpos': (f'=INDEX(Mapping!$D${MAP_FB0}:$D${MAP_FB1},MATCH({map_sub("G", r)},'
                      f'Mapping!$A${MAP_FB0}:$A${MAP_FB1},0))'),
            'eY': elast('inc', IC['efuel'], r), 'eU': elast('usg', IC['efuel'], r),
            'eF': elast('eff', IC['efuel'], r), 'a': elast('aei', IC['afuel'], r),
            'grow': 0,
            'p0': f'={price("mit.sp", r)}/${IC["gj"]}{r}',
            't0': f'=({price("mit.rp", r)}-{price("mit.sp", r)})/${IC["gj"]}{r}',
            'ef': (f'=INDEX(EF_GHG!$E$4:$E$1000,MATCH(Settings!$C$4&"|"&$B{r}&"|"&${IC["efsec"]}{r},'
                   f'EF_GHG!$A$4:$A$1000,0))'),
            'f0': (f'=INDEX(EnergyCons!$F$4:$F$5000,MATCH(Settings!$C$4&"|"&Settings!$C$7&"|"&$C{r}&"|"&$B{r},'
                   f'EnergyCons!$A$4:$A$5000,0))'),
        }
        for _, code in INPUT_COLS:
            fmt = {'p0': '0.000', 't0': '0.000', 'ef': '0.0000', 'f0': '#,##0.0', 'grow': '0.0%',
                   'eY': '0.00', 'eU': '0.00', 'eF': '0.00', 'a': '0.0%'}.get(code)
            put(ws, f'{IC[code]}{r}', vals[code], FONT, F_INPUT if code in ('fuel', 'sub', 'grow') else None, fmt)
    ws.freeze_panes = 'D5'
    ws.column_dimensions['A'].width = 10
    for j in range(2, len(INPUT_COLS) + 1):
        ws.column_dimensions[L(j)].width = 10
    ws.row_dimensions[4].height = 48


def build_variables(wb):
    ws = wb.create_sheet('Variables')
    title(ws, 'Variables: label, unit, code suffix, source and parameter columns by variable code', 10)
    put(ws, 'A3', 'Mitigation looks up Description, Unit and Source here and builds the output code '
                  'country.mit.<variable>.<subsector>.<fuel>.<suffix>.<scenario>. Param 1-4 = Mitigation D:G.', FONT)
    for j, h in enumerate(['Variable code', 'Label', 'Unit', 'Code suffix', 'Source', 'Legacy reference',
                           'Param 1 (D)', 'Param 2 (E)', 'Param 3 (F)', 'Param 4 (G)'], 1):
        put(ws, f'{L(j)}4', h, FONT_B, F_INPUT)
    for i, v in enumerate(VARIABLES, VR0):
        for j, x in enumerate(v, 1):
            if x == '=MT':  # label / unit of an MTInputs parameter: looked up there by NameOfParameter
                src = 'B' if j == 2 else 'C'
                x = f'=INDEX(MTInputs!${src}$8:${src}$415,MATCH($A{i},MTInputs!$H$8:$H$415,0))&""'
            put(ws, f'{L(j)}{i}', x, FONT, F_INPUT if j < 6 and not str(x).startswith('=') else None)
    for col, w in {'A': 14, 'B': 46, 'C': 9, 'D': 9, 'E': 34, 'F': 60, 'G': 16, 'H': 16, 'I': 16, 'J': 16}.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------- formula helpers
def group_cols(g):
    """(code column, base-year column, projection columns) of scenario group g."""
    code = COL_G0 + (g - 1) * GW
    return code, code + 1, list(range(code + 2, code + 1 + NY))


def var_lookup(col, r):
    return f'INDEX(Variables!${col}${VR0}:${col}${VR1},MATCH($A{r},Variables!$A${VR0}:$A${VR1},0))'


def label_formula(r):
    return (f'={var_lookup("B", r)}&IF($C{r}="",""," | "'
            f'&INDEX(Mapping!$B${MAP_SUB0}:$B${MAP_SUBLAB1},MATCH($C{r},Mapping!$A${MAP_SUB0}:$A${MAP_SUBLAB1},0))'
            f'&" | "&INDEX(Mapping!$B${MAP_FUEL0}:$B${MAP_FUELLAB1},'
            f'MATCH($B{r},Mapping!$A${MAP_FUEL0}:$A${MAP_FUELLAB1},0)))')


def code_formula(r, code_col):
    sfx = var_lookup('D', r)
    return (f'=LOWER(Settings!$C$4)&".mit."&$A{r}&IF($C{r}="","","."&$C{r}&"."&$B{r})'
            f'&IF({sfx}="","","."&{sfx})&"."&{code_col}${R_SCEN}')


def inputs_lookup(code, r):
    return f'INDEX(Inputs!${IC[code]}${IN_R0}:${IC[code]}${IN_R1},MATCH($B{r}&"|"&$C{r},Inputs!$A${IN_R0}:$A${IN_R1},0))'


def pol_range(col, key0, n):
    """Absolute-row range of n consecutive Section-1 rows starting at key0, in column col."""
    r0 = R_POL[key0]
    return f'{col}${r0}:{col}${r0 + n - 1}'


def calc_formula(var, r, col, prev, is_base, use_lambda):
    """Sector-section cell of variable var in row r (column col; prev = previous year column)."""
    o = lambda v: f'{col}{r - VOFF[var] + VOFF[v]}'          # same fuel, other variable, this column
    op = lambda v: f'{prev}{r - VOFF[var] + VOFF[v]}'        # same fuel, other variable, previous column
    if var == 'sp':
        if is_base:
            return '=' + inputs_lookup('p0', r)
        return f'=PRETAX({prev}{r},$D{r})' if use_lambda else f'={prev}{r}*(1+$D{r})'
    if var == 'ctxnew':
        return (f'={col}${R_CP}*$D{r}*INDEX({pol_range(col, "fc.coa", NF)},$E{r})'
                f'*INDEX({pol_range(col, "sc.pow", len(COV_SECTORS))},$F{r})')
    if var == 'ntx':
        return f'=INDEX({pol_range(col, "fpr." + PRICE_FUELS[0][0], len(PRICE_FUELS))},$D{r})/$E{r}'
    if var == 'nce':
        return f'={o("ctxnew")}+{o("ntx")}'
    if var == 'tax':
        return f'=TAX($D{r},{o("nce")})' if use_lambda else f'=$D{r}+{o("nce")}'
    if var == 'atp':
        return f'=POSTTAX({o("sp")},{o("tax")})' if use_lambda else f'={o("sp")}+{o("tax")}'
    if var == 'shp':
        return (f'=INDEX({pol_range(col, "shps." + FEEBATE_SECTORS[0][0], len(FEEBATE_SECTORS))},$D{r})*$E{r}'
                f'*INDEX({pol_range(col, "ssc.pow", len(COV_SECTORS))},$F{r})')
    if var == 'ener':
        if is_base:
            return '=' + inputs_lookup('f0', r)
        if use_lambda:
            return (f'=FUELUSE({prev}{r},{o("atp")},{op("atp")},{o("shp")},{op("shp")},{col}${R_GDP},'
                    f'$D{r},$E{r},$F{r},$G{r})')
        ratio = f'({o("atp")}/{op("atp")})'
        eff = f'(({o("atp")}+{o("shp")})/({op("atp")}+{op("shp")}))'
        return (f'={prev}{r}*(1/(1+$G{r}))^(1+$E{r})*(1+{col}${R_GDP})^$D{r}'
                f'*{ratio}^$E{r}*{eff}^($F{r}*(1+$E{r}))')
    raise ValueError(var)


LAMBDA_VARS = {'sp', 'tax', 'atp', 'ener'}


def mt_lookup(r, col):
    """Section-1 input: MTInputs row (column D) for this column's scenario (row 5)."""
    return (f'=INDEX(MTInputs!${L(MT_COL0)}$1:${MT_LAST}$415,$D{r},'
            f'MATCH({col}${R_SCEN},MTInputs!${L(MT_COL0)}${MT_ROW_SCEN}:${MT_LAST}${MT_ROW_SCEN},0))')


def path_formula(col, y0, y1, start, target, linear_ext=None):
    """0 before y0; linear from start (at y0) to target (at y1). After y1: linear continuation (the one
    default for all paths); the carbon price keeps its MTInputs switch (linear_ext cell, "Linear*" = continue,
    otherwise flat at the target, as legacy)."""
    y = f'{col}${R_YEAR}'
    slope = f'({target}-{start})/MAX({y1}-{y0},1)'
    if linear_ext is None:
        return f'=IF({y}<{y0},0,{start}+{slope}*({y}-{y0}))'
    return (f'=IF({y}<{y0},0,{start}+{slope}'
            f'*IF(OR({y}<={y1},{linear_ext}="Linear*"),{y}-{y0},{y1}-{y0}))')


def pol_calc(key, col):
    c = lambda k: f'{col}${R_POL[k]}'          # fixed rows (years, carbon-tax inputs): absolute
    rel = lambda k: f'{col}{R_POL[k]}'          # start / final values: relative, so a path drags down
    if key == 'cptraj':
        return path_formula(col, c('CPIntro'), c('CPOutro'), c('CPLevelStart'), c('CPLevelTarget'),
                            c('ExtendCarbonPriceBeyondOutro'))
    if key.startswith('fpr.'):
        pc = key[4:]
        return path_formula(col, c('fpr.yr0'), c('fpr.yr1'), rel('fprs.' + pc), rel('fprf.' + pc))
    if key.startswith('fb.') and key not in ('fb.yr0', 'fb.yr1'):
        s = key[3:]
        return path_formula(col, c('fb.yr0'), c('fb.yr1'), rel('fbs.' + s), rel('fbt.' + s))
    if key.startswith('shps.'):
        return f'={rel("fb." + key[5:])}'                       # feebates only, for now
    if key.startswith('ssc.'):
        return f'={rel("fbc." + key[4:])}*{c("ssc.adj")}'
    raise ValueError(key)


# ---------------------------------------------------------------- Mitigation
def build_mitigation(wb):
    ws = wb.create_sheet('Mitigation', 1)
    groups = [s[0] for s in SCENARIOS]
    last = COL_G0 + len(groups) * GW - 1
    outline = {}                                            # row -> (level, hidden)
    for j, h in enumerate(['Variable', 'Fuel', 'Sector'], 1):
        put(ws, f'{L(j)}1', h, FONT_B)
    for col in PARAM_COLS:
        put(ws, f'{col}1', f'Param {PARAM_COLS.index(col) + 1}', FONT_B)
    for col, h in DESC.items():
        put(ws, f'{col}1', h, FONT_B)
    title(ws, 'Mitigation module - price to fuel use (copy-pasteable, design 2, legacy CPAT sections)', last)
    band(ws, 3, 'Scenario assumptions (one column group per scenario: code column + years; copy a whole group '
                'to add a scenario)', last)
    put(ws, f'H{R_YEAR}', 'Year', FONT_B)
    put(ws, f'H{R_SCEN}', 'Scenario number (MTInputs column)', FONT_B)
    put(ws, f'H{R_NAME}', 'Scenario name (from MTInputs)', FONT_B)

    def describe(r, font=FONT):
        put(ws, f'H{r}', label_formula(r), font)
        put(ws, f'I{r}', '=' + var_lookup('C', r), font)
        put(ws, f'J{r}', '=' + var_lookup('E', r), font)

    def ids(r, var, fuel, sector, font=FONT):
        put(ws, f'A{r}', var, font)
        put(ws, f'B{r}', fuel, font)
        put(ws, f'C{r}', sector, font)
        describe(r, font)

    def codes(r, font=FONT):
        for g in groups:
            cc = L(group_cols(g)[0])
            put(ws, f'{cc}{r}', code_formula(r, cc), font, F_CODE)

    def each_col(fn):
        for g in groups:
            _, base, years = group_cols(g)
            for c in [base] + years:
                fn(c, c == base, c == years[-1])

    # Scenario assumption rows
    ids(R_GDP, 'gdp.pos.pct', '', '', FONT_B)
    codes(R_GDP)
    for g, gname, _ in SCENARIOS:
        code, base, years = group_cols(g)
        cc = L(code)
        put(ws, f'{cc}1', f'Scenario {g}: output code', FONT_B)
        put(ws, f'{cc}{R_SCEN}', 1 if g == 1 else f'={L(code - GW)}{R_SCEN}+1', FONT_B,
            F_INPUT if g == 1 else None, '0')
        put(ws, f'{L(base)}{R_NAME}', (f'=INDEX(MTInputs!${L(MT_COL0)}${MT_ROW_NAME}:${MT_LAST}${MT_ROW_NAME},'
                                       f'MATCH({L(base)}${R_SCEN},MTInputs!${L(MT_COL0)}${MT_ROW_SCEN}:'
                                       f'${MT_LAST}${MT_ROW_SCEN},0))'), FONT_B)

    def top(c, is_base, is_last):
        col, prev = L(c), L(c - 1)
        bf = F_BASE if is_base else None
        put(ws, f'{col}{R_YEAR}', '=Settings!$C$7' if is_base else f'={prev}{R_YEAR}+1', FONT_B, bf, '0')
        put(ws, f'{col}{R_SCEN}', f'={prev}{R_SCEN}', FONT, bf, '0')
        put(ws, f'{col}{R_GDP}', (f'=INDEX(WEO!$F$4:$AZ$1000,MATCH(Settings!$C$4&"|gdp_growth",WEO!$A$4:$A$1000,0),'
                                  f'MATCH({col}${R_YEAR},WEO!$F$3:$AZ$3,0))'), fill_=bf, fmt='0.0%')
    each_col(top)

    # 1. Policies
    band(ws, B_POL, '1. Policies (scenario inputs from MTInputs by row and scenario number; policy paths)', last)
    put(ws, f'D{B_POL}', 'MTInputs row', FONT_BAND, F_BAND)
    for key, var, fuel, sector, kind, mt in POL:
        r = R_POL[key]
        if SUBHEADS.get(key):
            put(ws, f'H{r - 1}', SUBHEADS[key], Font(name='Arial', size=9, bold=True, italic=True))
            outline[r - 1] = (1, True)
        summary = key == 'cptraj'
        fnt = FONT_SUM if summary or kind == 'calc' else FONT
        ids(r, var, fuel, sector, fnt)
        if mt:
            put(ws, f'D{r}', mt, FONT, F_INPUT)
        codes(r, fnt)
        outline[r] = (0, False) if summary else (1, True)

        def pcell(c, is_base, is_last, r=r, key=key, kind=kind, fnt=fnt):
            col = L(c)
            bf = F_BASE if is_base else None
            if kind == 'mt':
                val = mt_lookup(r, col)
            elif kind == 'calc':
                val = pol_calc(key, col)
            else:
                val = CONST_VALUES[key]
            fmt = '0' if var in ('CPIntro', 'CPOutro', 'fpr.yr0', 'fpr.yr1', 'fb.yr0', 'fb.yr1') else '0.00'
            if var in ('ctcov', 'fbcov') and kind == 'const':
                fmt = 'General'
            put(ws, f'{col}{r}', val, fnt, F_INPUT if kind == 'const' else bf, fmt)
        each_col(pcell)
    ws.row_dimensions[R_CP].collapsed = True

    # 3. Power sector (placeholder)
    band(ws, B_POW, '3. Power sector (elasticity-based / engineer model): not yet modelled', last)
    put(ws, f'H{R_POW_NOTE}', 'Placeholder: the power sector comes in a later bucket. MTInputs power switches '
                              '(MCovPow, feebate power) are read in section 1 but not used yet.',
        Font(name='Arial', size=9, italic=True))

    # 5-8. Sector sections
    for num_, grp, title_ in SECTIONS:
        b = SEC_BAND[grp]
        band(ws, b, f'{num_}. {title_}', last)
        subs = [s[0] for s in SUBSECTORS if s[2] == grp]
        r_sum = SEC_SUM[grp]
        ids(r_sum, 'ener', 'all', grp, FONT_SUM)
        codes(r_sum, FONT_SUM)
        heads = [SUB_HEAD[s] for s in subs]
        a, z = heads[0], heads[-1] + SB - 1

        def secsum(c, is_base, is_last, r=r_sum, a=a, z=z):
            col = L(c)
            put(ws, f'{col}{r}', f'=SUMIFS({col}{a}:{col}{z},$A{a}:$A{z},"ener",$B{a}:$B{z},"<>all")', FONT_SUM,
                F_BASE if is_base else None, '#,##0.0')
        each_col(secsum)
        outline[r_sum] = (0, False)
        ws.row_dimensions[r_sum].collapsed = False
        for s in subs:
            h = SUB_HEAD[s]
            ids(h, 'ener', 'all', s, FONT_SUM)
            codes(h, FONT_SUM)
            outline[h] = (1, False)
            ws.row_dimensions[h].collapsed = True
            e0 = h + VOFF['ener']

            def subsum(c, is_base, is_last, h=h, e0=e0):
                col = L(c)
                put(ws, f'{col}{h}', f'=SUM({col}{e0}:{col}{e0 + NF - 1})', FONT_SUM, F_BASE if is_base else None,
                    '#,##0.0')
            each_col(subsum)
            for k, var in enumerate(VARS):
                params = VAR_PARAMS[var]
                for i, (f, *_r) in enumerate(FUELS):
                    r = h + VOFF[var] + i
                    ids(r, var, f, s)
                    codes(r)
                    outline[r] = (2, True)
                    for j, pc in enumerate(PARAM_COLS):
                        if j < len(params):
                            put(ws, f'{pc}{r}', '=' + inputs_lookup(params[j], r), fmt='0.000')
                        else:
                            put(ws, f'{pc}{r}', None, fill_=F_UNUSED)
                    fmt = '#,##0.0' if var == 'ener' else '0.000'

                    def cell(c, is_base, is_last, r=r, var=var, fmt=fmt):
                        lam = is_last and var in LAMBDA_VARS
                        put(ws, f'{L(c)}{r}', calc_formula(var, r, L(c), L(c - 1), is_base, lam),
                            FONT_LAMBDA if lam else FONT, F_LAMBDA if lam else (F_BASE if is_base else None), fmt)
                    each_col(cell)
                outline[h + VOFF[var] + NF] = (2, True)              # blank row after each variable

    # 11. Results - energy consumption
    band(ws, B_RES, '11. Results - energy consumption', last)
    a, z = S_FIRST, S_LAST
    ids(R_TOTAL, 'ener', 'all', 'all', FONT_SUM)
    codes(R_TOTAL, FONT_SUM)

    def tot(c, is_base, is_last):
        col = L(c)
        put(ws, f'{col}{R_TOTAL}', f'=SUMIFS({col}${a}:{col}${z},$A${a}:$A${z},"ener",$B${a}:$B${z},"<>all")',
            FONT_SUM, F_BASE if is_base else None, '#,##0.0')
    each_col(tot)
    for i, (s, *_r) in enumerate(SUBSECTORS):
        r = R_SUB0 + i
        ids(r, 'ener', 'all', s)
        codes(r)

        def bysub(c, is_base, is_last, r=r):
            col = L(c)
            put(ws, f'{col}{r}', (f'=SUMIFS({col}${a}:{col}${z},$A${a}:$A${z},"ener",$C${a}:$C${z},$C{r},'
                                  f'$B${a}:$B${z},"<>all")'), fill_=F_BASE if is_base else None, fmt='#,##0.0')
        each_col(bysub)
    for i, (f, *_r) in enumerate(FUELS):
        r = R_FUEL0 + i
        ids(r, 'ener', f, 'all')
        codes(r)

        def byfuel(c, is_base, is_last, r=r):
            col = L(c)
            put(ws, f'{col}{r}', f'=SUMIFS({col}${a}:{col}${z},$A${a}:$A${z},"ener",$B${a}:$B${z},$B{r})',
                fill_=F_BASE if is_base else None, fmt='#,##0.0')
        each_col(byfuel)
    far = L(last + 10 * GW)
    extra = [(R_CHK, 'ener.chk', '', '',
              lambda col: (f'=ABS(SUM({col}{R_SUB0}:{col}{R_SUB0 + NS - 1})-{col}{R_TOTAL})'
                           f'+ABS(SUM({col}{R_FUEL0}:{col}{R_FUEL0 + NF - 1})-{col}{R_TOTAL})'), '0.000', F_CHECK),
             (R_REF, 'ener.ref', 'all', 'all',
              lambda col: (f'=INDEX(${L(COL_G0)}{R_TOTAL}:${far}{R_TOTAL},'
                           f'MATCH({col}${R_YEAR},${L(COL_G0)}${R_YEAR}:${far}${R_YEAR},0))'), '#,##0.0', None),
             (R_PCT, 'ener.pct', 'all', 'all', lambda col: f'={col}{R_TOTAL}/{col}{R_REF}-1', '0.00%', None)]
    for r, var, fu, sub, f, fmt, fl in extra:
        ids(r, var, fu, sub, FONT_B)
        codes(r)

        def ex(c, is_base, is_last, r=r, f=f, fmt=fmt, fl=fl):
            put(ws, f'{L(c)}{r}', f(L(c)), fill_=fl or (F_BASE if is_base else None), fmt=fmt)
        each_col(ex)
    for r in range(R_SUB0, R_PCT + 1):
        outline[r] = (1, False)

    # Outline: summary rows above details; levels as set above.
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_format.outlineLevelRow = 2
    for r, (lvl, hid) in outline.items():
        if lvl:
            ws.row_dimensions[r].outlineLevel = lvl
            ws.row_dimensions[r].hidden = hid
    ws.column_dimensions.group('D', 'G', outline_level=1, hidden=True)
    ws.sheet_format.outlineLevelCol = 1
    ws.freeze_panes = f'{L(COL_G0 + 1)}{R_SCEN + 1}'
    for col, w in {'A': 11, 'B': 5, 'C': 5, 'H': 44, 'I': 9, 'J': 18}.items():
        ws.column_dimensions[col].width = w
    for g in groups:
        code, base, years = group_cols(g)
        ws.column_dimensions[L(code)].width = 24
        for c in [base] + years:
            ws.column_dimensions[L(c)].width = 9
    return last


# ---------------------------------------------------------------- ReadMe
def build_readme(wb):
    ws = wb.active
    ws.title = 'ReadMe'
    title(ws, f'CPAT mitigation module - copy-pasteable prototype v{VERSION}', 3)
    code2 = group_cols(2)[0]
    lam_lines = '; '.join(f'{n}({", ".join(p)}) = {b}' for n, (p, b) in LAMBDAS.items())
    h = SUB_HEAD['rod']
    lines = [
        ('Purpose', 'Auditable, copy-pasteable replacement of the CPAT mitigation price -> fuel-use chain. Each '
                    'variable has one formula across all subsectors, fuels and years, and a whole scenario group '
                    'copies to a new scenario and keeps working.'),
        ('Status', f'v{VERSION} prototype: policies (carbon tax, fuel price reform, feebate shadow price) and '
                   'price -> fuel use for transport, buildings, industry and other energy use. Power, emissions, '
                   'revenue and ETS not yet.'),
        ('Layout (as legacy CPAT)', f'Sections: 1. Policies (row {B_POL}), 3. Power sector (row {B_POW}, '
                                    f'placeholder), 5. Transport (row {SEC_BAND["tra"]}), 6. Buildings (row '
                                    f'{SEC_BAND["bld"]}), 7. Industrial (row {SEC_BAND["ind"]}), 8. Other energy '
                                    f'use (row {SEC_BAND["oen"]}), 11. Results - energy consumption (row {B_RES}). '
                                    'Numbers follow the legacy Mitigation sheet.'),
        ('Nesting', 'Across: scenario (outer) > year (inner). Down: sector section > subsector > variable > fuel. '
                    f'Each subsector: a heading line (its total fuel use), then {", ".join(VARS)}, each over the '
                    f'{NF} fuels with one blank row ({SB} rows per subsector).'),
        ('Variables', 'sp pre-tax price; ctxnew new carbon tax = carbon price x EF x fuel coverage x sector '
                      'coverage; ntx new excise = fuel price reform path / GJ per unit; nce total new policy = '
                      'ctxnew + ntx; tax = base tax + nce; atp after-tax price = sp + tax; shp shadow price = '
                      'sector shadow price x EF x share impacting efficiency; ener fuel use (CPAT eq. 3.3.3): the '
                      'usage term uses atp, the efficiency term (atp + shp), as legacy CPAT.'),
        ('Roll-up', 'Sector sections show their total and the subsector heading lines; click + on a heading to '
                    'see its variables (outline level 2). Section 1 shows the carbon price; its inputs and paths '
                    'are rolled up. Buttons 1 / 2 / 3 at the top left set all levels.'),
        ('Section 1 (policies)', 'Each input row reads MTInputs by its row number (column D, hidden) and the '
                                 f'scenario number (row {R_SCEN}). Paths: carbon price as legacy (0 before '
                                 'CPIntro, linear to CPLevelTarget by CPOutro, linear continuation if Linear*); '
                                 'fuel price reform and feebates: 0 before the start year, linear from the '
                                 'starting to the final/target value, continuing linearly afterwards (one default). '
                                 'Shadow prices: by sector ($/tCO2) = feebate path (later + non-auctioned ETS, '
                                 'regulations); share impacting efficiency = feebate coverage x adjustment (1.0).'),
        ('Data step vs formulas', 'Lookups happen in Inputs (one row per fuel|subsector), in Mitigation D:G '
                                  '(hidden; labels per variable in Variables G:J), the base-year column and the '
                                  'Section-1 input rows. Calculation cells contain no searching lookups; ctxnew, '
                                  'ntx and shp pick a Section-1 row by position with INDEX(range, position).'),
        ('LAMBDA column', f'{YEARS[-1]} (orange) calls named LAMBDAs for sp, tax, atp and ener; drag it back over '
                          f'the row to use them everywhere, or drag {YEARS[-2]} forward. {lam_lines}.'),
        ('Codes', 'Code column before each scenario group: country.mit.<variable>.<subsector>.<fuel>.<suffix>.'
                  f'<scenario> (e.g. egy.mit.ener.rod.gso.e.1). Scenario number at the top (row {R_SCEN}).'),
        ('Add a scenario', 'MTInputs: copy the last scenario column one column right (its number updates) and edit '
                           f'its inputs. Mitigation: copy columns {L(code2)}:{L(code2 + GW - 1)} and paste at '
                           f'{L(code2 + GW)}; numbers, names, codes and all policy paths follow.'),
        ('Example', f'Road (heading row {h}): sp rows {h + VOFF["sp"]}-{h + VOFF["sp"] + NF - 1}, ener rows '
                    f'{h + VOFF["ener"]}-{h + VOFF["ener"] + NF - 1}.'),
        ('Colours', 'Green bands with white text = sections; green = input; tan = base year; light blue = codes; '
                    'grey = unused parameter slot; orange = LAMBDA column; yellow = check.'),
        ('Data', 'Countries, Elasticities, Prices_dom, EnergyCons, WEO, EF_GHG; Mapping; Variables; MTInputs. '
                 f'Built by build_v{VERSION.replace(".", "_")}.py; checked by check_v{VERSION.replace(".", "_")}.py.'),
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
    build_inputs(wb)
    build_variables(wb)
    build_mtinputs(wb)
    build_mitigation(wb)
    div = wb.create_sheet('DATA->')
    put(div, 'A1', 'Data sheets follow', FONT_B)
    build_data_sheets(wb)
    build_readme(wb)
    for name in LAMBDAS:
        wb.defined_names[name] = DefinedName(name, attr_text=lambda_xml(name), comment=LAMBDA_NOTES[name])
    order = ['ReadMe', 'Settings', 'MTInputs', 'Mitigation', 'Inputs', 'Variables', 'Mapping', 'DATA->',
             'Countries', 'Elasticities', 'Prices_dom', 'EnergyCons', 'WEO', 'EF_GHG']
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    wb.save(OUT)
    fix_outline_levels(OUT)
    print('Saved', OUT)


if __name__ == '__main__':
    main()
