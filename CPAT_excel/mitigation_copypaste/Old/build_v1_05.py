"""Build CPAT-AI-Mitigation-MVP-v1.05.xlsm (design 2: time across, scenarios as column groups).

v1.00: the prototype is renamed CPAT-AI-Mitigation-MVP (earlier versions: CPAT_Mitigation_CopyPaste_v0.x) and
    upgraded to v1.00; content as v0.16.
v1.01: sheet MTOutputs collects key results for every scenario by output code (one row block per scenario,
    copy-pasteable: copy the last block below to add a scenario; its number goes up by one), and sheet Charts plots
    them (CO2, fuel use, fiscal effect, carbon price by scenario).
v1.02: new ETS with a cap (ETS_Legacy_Algorithm_v0_1.md). The cap is a change relative to baseline covered emissions
    (MTInputs 85-87); baseline CO2 by ETS sector is a data block in section 1 (refreshed from scenario 1 by the
    build, option A of ETS_Cap_Design_Options_v0_1.md). Permit price: fast estimate LN(cap / baseline) / (covered
    semi-elasticity x allocation effectiveness x volatility adjustment), or an override row (MTInputs
    D_ETSPriceOverride = Yes; pasted from ets_goalseek_v0_1.py, blank = carbon price path). Benchmarks by ETS sector
    group replace the auction share: output-based (OBR) share = benchmark; the permit price is split into a
    tax-equivalent part p x (1 - OBR) in the fuel price and a shadow price p x OBR on the efficiency margin (as
    feebates). Volatility adjustment (1 + carbon tax policy risk x impact) / (1 + ETS volatility x impact) from
    MTInputs 95-97 (legacy ETS+LTS table; 1/1.1 at Medium). Section 13 compares covered emissions with the cap and
    proposes the next price (ets.next).
v1.03: multiple scenarios through stored results. MTInputs: J = scenario 1 (baseline), K = scenario 2 (live policy
    scenario, calculated by Mitigation group 2), L onwards = scenario definitions (row 4 Run? = Yes/No; combined
    packages added). The VBA module CPATScenarios (CPATScenarios_v0_1.bas; v0_2 from v1.04, embedded by vba_project_v0_1.py, so the
    workbook is an .xlsm) copies each definition into K, recalculates, runs the ETS goal seek when the definition
    applies a new ETS, and stores the MTOutputs block of scenario 2 as values in sheet StoredResults (the baseline is
    stored once). ScenarioCompare looks up key results for one year from StoredResults and compares them with the
    baseline. The build fills StoredResults with a Python emulation of the macro (LibreOffice recalculation, same
    goal seek: ets_goalseek_v0_2.py).
v1.04: first-run test for Excel. Module CPATScenarios_v0_2.bas adds the macro CheckBatchRun: it checks that the
    macros find every sheet, row and column they use, compares the host's own recalculation of the baseline with the
    stored baseline (Excel evaluates the LAMBDAs natively), reruns the batch and compares every stored value with the
    values stored before; the report goes to the new sheet MacroCheck (Result PASS / CHECK / FAIL).
v1.05: power sector, price side (PowerPrices_Method_v0_1.md): Mitigation section 3 with legacy sub-tables A Inputs,
    B amortised fixed costs (vintages), C levelised fixed costs, D variable costs and new policy cost per kWh, E total
    and average generation cost, G end-user electricity prices and power revenues, J storage cost of variable
    renewables. Data step: sheet Inputs_power (one row per generation type, derived constants) and data sheets
    PowerTech, PowerPaths, PowerParams (extract_power_v0_1.py). Generation shares, investment shares and
    consumption are INTERIM data (legacy baseline) until the engineer model is built. MTOutputs adds the residential
    and industrial electricity prices and the average generation cost.

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
v0.9: domestic price projection (legacy method, PriceProjection_Method_v0.2.md) in a new section
    2. Retail energy prices, by the 12 price fuels (coal and gas x power / residential / industry, 4 oil products,
    other oil products, biomass): international prices gp (source and adjustment from MTInputs), supply cost sp
    (fixed part + part floating with gp), excise and other taxes txo (pass-through rule), retail price before
    new policies rpb. Everything in real USD of ResultsYear (CPI index for domestic data and nominal inputs,
    US GDP deflator index for international prices; rows infl and defl at the top). Per subsector the variables
    are now ctxnew, ntx, nce, atp (= rpb of its price fuel + nce x (1 + VAT rate)), shp, ener.
v0.10 (user decisions): other oil products are treated like the other oil products (pass-through and margin from
    the IMF dataset instead of the legacy hardcodes 1 and 0); VAT-rate assumption: where the dataset's VAT rate is
    blank, the country's general VAT rate (VAT_WEO, Egypt 14%) applies to residential coal and gas and to the
    all-sector oil products; power, industry and biomass 0. Cells whose data changed by assumption are bright
    yellow with red text (Inputs_prices).
v0.11 (user decisions): horizon 2040; food & forestry takes the services (buildings) elasticities; new ETS (price
    = carbon price inputs, as effective as a carbon tax, replaces the carbon tax in ETS-covered sectors; auction share
    kept for revenue); sectoral shadow prices = feebates + regulations (placeholder 0); section 2 completed with
    existing taxes and subsidies (VAT payment, existing tax, existing subsidy, subsidy per price unit) for revenue
    and for later per-fuel policy inputs; section 12 Revenues prepared. Column blocks: history and projection have
    separate plain formulas; the right column (2040) of every calculated row calls a named LAMBDA, which carries
    the IF between history and projection where needed. Sheet LegacyDiff lists the differences with legacy CPAT.
v0.12: revenues, as three separate calculations per subsector and fuel (USD million, real of ResultsYear):
    rtx existing tax revenue = fuel use x (existing tax etx + VAT vat of the price fuel); rsub existing subsidy cost
    = fuel use x existing subsidy esub; rnew new-policy revenue = fuel use x (ctxnew + ntx + ets x auctioned share +
    nce x VAT rate). Section 12 sums them by sector, nets existing taxes and subsidies, and compares with scenario 1.
v0.13: CO2 emissions from fuel combustion per subsector and fuel, co2 = fuel use x PJ/ktoe x EF (MtCO2);
    section 13 sums them by sector and fuel and compares with scenario 1. Other GHGs and local pollutants later.
v0.14: corrected Egypt price block (user, data/source/Egypt_Price_Data_2026-10-09.xlsx, applied to
    data/prices_dom.csv by update_prices_egypt_v0_1.py); changed cells bright yellow in Prices_dom. VAT rule: the
    dataset rate wherever the cell is filled (explicit 0 included); the VAT_WEO assumption only for blank cells.
v0.15 (layout only): columns B:C and I:K grouped (level 1) with D:G nested (level 2); row groups have their
    summary / button below (summaryBelow); everything opens rolled up; Mitigation is the first tab, zoom 75%.
v0.16 (layout only): one level of column groups: B:G (fuel, sector, parameters), I:K, and 2030-2039 in every
    scenario group; light beige (EEECE1) for blocks that are copy-pasteable internally but not draggable into the
    white or base-year (darker beige DDD9C4) columns: the LAMBDA column and the 2023-2024 history block of sp, txo;
    a blank spacer column closes each scenario group (copy code column to spacer to add a scenario).
The distance between variables is the same in every subsector, so every variable has one formula everywhere.

- Data input is a separate step: sheet Inputs (one row per fuel|subsector) holds the mappings and parameter
  lookups; Mitigation reads it in the hidden parameter columns D:G and the base-year column. Section 1 reads the
  scenario inputs from MTInputs (one Used-for-calculation column per scenario) by MTInputs row (column D) and the
  scenario number (row 5), then computes the policy paths (carbon price, fuel price reform, feebates).
- Calculation cells contain no searching lookups; a few pick a Section-1 row by position (INDEX(range, position)),
  the position being a data-step parameter in D:G.
- 2023-2034 plain formulas; 2035 calls named LAMBDAs (SUPPLYCOST, OTHERTAX, POSTTAX, FUELUSE) for sp, txo,
  atp, ener.
- Each scenario group starts with a code column building country.mit.<var>.<subsector>.<fuel>.<suffix>.<scen>.

Reads data/*.csv (extract_data_v0_2.py) and templates/MTInputs_template.xlsx. Writes formulas only.

    python build_v1_05.py
"""
import csv
import os
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI_
from openpyxl.workbook.defined_name import DefinedName

VERSION = '1.05'
NAME = 'CPAT-AI-Mitigation-MVP'
HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, 'data')
OUT = os.path.join(HERE, f'{NAME}-v{VERSION}.xlsm')
BAS = os.path.join(HERE, 'CPATScenarios_v0_2.bas')          # VBA module (embedded; also for import)

# ---------------------------------------------------------------- dimensions
SUBSECTORS = [  # code, name, group, elasticity sector, EF sector, coal/gas price sector, feebate sector
    ('rod', 'Road', 'tra', 'tra', 'rod', 'res', 'tra'),
    ('ral', 'Rail', 'tra', 'tra', 'rod', 'ind', 'tra'),
    ('avi', 'Domestic aviation', 'tra', 'tra', 'rod', 'ind', 'tra'),
    ('nav', 'Domestic navigation', 'tra', 'tra', 'rod', 'ind', 'tra'),
    ('res', 'Residential', 'bld', 'res', 'res', 'res', 'res'),
    ('foo', 'Food & forestry', 'bld', 'srv', 'res', 'ind', 'res'),   # v0.11: services (buildings) elasticities
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
                ('ind', 'Industry'), ('int', 'International')]                      # labels only (totals, policy rows)
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
PRICES = [(pc, lab) for pc, lab, *_ in PRICE_FUELS if not pc.startswith('ecy')]   # 12 price fuels (section 2)
NP = len(PRICES)
GP_ROWS = [  # international price rows: fuel code, label, commodity key (D), High factor (E), Low factor (F)
    ('oil', 'Crude oil', 'oil', 1.5, 0.5), ('coa', 'Coal', 'coa', 1.25, 0.75),
    ('nga', 'Natural gas (country gas market)', '="nga."&Settings!$C$12', 1.25, 0.75),
    ('one', 'Flat (no price change)', None, 1, 1)]
GP_POS = {'gso': 1, 'die': 1, 'lpg': 1, 'ker': 1, 'oop': 1, 'coa': 2, 'nga': 3, 'bio': 4}
RP_FROM_DATA = {'coa', 'nga', 'bio'}    # historical retail price from data; oil products: sp + txo (legacy)
VAT_CONSUMER = {'coa.res', 'nga.res', 'gso.all', 'die.all', 'lpg.all', 'ker.all', 'oop.all'}   # VAT assumption applies
MT_ROWS_FIXED = {'fpr.yr0': 140, 'fpr.yr1': 141, 'fbcov0': 66, 'etscov0': 99}   # MTInputs rows without NameOfParameter

BASE_YEAR, LAST_YEAR = 2022, 2040
YEARS = list(range(BASE_YEAR + 1, LAST_YEAR + 1))
SCENARIOS = [(1, 'Baseline (no carbon price)', {}),
             (2, 'Carbon price $20/tCO2 from 2027', {y: 20 for y in YEARS if y >= 2027})]
MT_SCENARIO_INPUTS = {
    1: {'CPIntro': 2027, 'CPLevelStart': 0, 'CPLevelTarget': 0, 'CPOutro': 2030, 'MCovOen': False},
    2: {'CPIntro': 2027, 'CPLevelStart': 20, 'CPLevelTarget': 20, 'CPOutro': 2030, 'MCovOen': False},
}
FB_PACKAGE = {'D_FeebateIntro': 2027, 'D_FeebateOutro': 2030, 'D_Feb_Level_Start_Trans': 10,
              'D_Feb_Level_Target_Trans': 50, 'D_Feb_Level_Start_Ind': 5, 'D_Feb_Level_Target_Ind': 25,
              67: True, 74: True, 75: True}                    # feebate coverage (MTInputs rows): road, mch, irn
DEFINITIONS = [  # scenario definitions on MTInputs (columns after the live scenario 2): number, name, Run?, inputs
    (3, 'Carbon price $20/tCO2 from 2027', 'Yes', dict(MT_SCENARIO_INPUTS[2])),
    (4, 'Package A: carbon tax to $50 by 2030 + feebates (transport, industry)', 'Yes',
     {'CPIntro': 2027, 'CPLevelStart': 10, 'CPLevelTarget': 50, 'CPOutro': 2030, 'MCovOen': False, **FB_PACKAGE}),
    (5, 'Package B: ETS on power and industry (cap -5% 2027 to -20% 2035) + carbon tax $25 elsewhere', 'Yes',
     {'CPIntro': 2027, 'CPLevelStart': 25, 'CPLevelTarget': 25, 'CPOutro': 2030, 'MCovOen': False, 'D_NewETS': 'Yes',
      'D_ETSIntro': 2027, 'D_ETSOutro': 2035, 'D_ETSChangeRelStart': -0.05, 'D_ETSChangeRelTarget': -0.2,
      'D_ETSCapCont': 'Constant'}),
]
MT_ROW_RUN = 4                          # MTInputs: Run? flag of a scenario definition
MT_TEMPLATE = next(os.path.join(d, 'templates', 'MTInputs_template.xlsx')      # repo templates folder,
                   for d in [os.path.abspath(os.path.join(HERE, *['..'] * k)) for k in range(2, 5)]   # also from Old/
                   if os.path.exists(os.path.join(d, 'templates', 'MTInputs_template.xlsx')))
MT_COL0 = 10                            # J: first scenario column on MTInputs
MT_ROW_SCEN, MT_ROW_NAME = 5, 6
MT_LAST = 'AZ'
ETS_PARAMS = ['D_NewETS', 'D_ETSChangeRelStart', 'D_ETSChangeRelTarget', 'D_ETSCapCont', 'D_ETSIntro', 'D_ETSOutro',
              'D_ETSPriceOverride', 'D_ETSVolatility', 'D_ETSVolImpact', 'D_ETSCTRisk']
ETS_GROUPS = [  # ETS sector groups for benchmarks (= feebate sectors): code, benchmark at start, at target (assumption)
    ('pow', 1.0, 0.8), ('tra', 1.0, 0.8), ('res', 1.0, 0.8), ('ind', 1.0, 0.8)]
ETS_SEMI = {'pow': -0.0027939, 'tra': -0.002833, 'res': -0.0038637, 'ind': -0.0054897}   # legacy rows 1815-1831, EGY
CP_PARAMS = ['CPIntro', 'CPLevelStart', 'CPLevelTarget', 'CPOutro', 'ExtendCarbonPriceBeyondOutro', 'NomorReal']
PRI_PARAMS = ['IntEnerPricForeSource', 'IntEnerPricForecastAdjustment']   # section 2 selectors

# ---------------------------------------------------------------- Mapping layout
MAP_SUB0 = 5                            # subsector table rows 5-20, extra label rows after
MAP_SUB1 = MAP_SUB0 + NS - 1
MAP_SUBLAB1 = MAP_SUB1 + len(EXTRA_LABELS)
MAP_FUEL0 = MAP_SUBLAB1 + 4             # fuel table
MAP_FUEL1 = MAP_FUEL0 + NF - 1
FUEL_EXTRA = [('all', 'All fuels'), ('ecy', 'Electricity'), ('oil', 'Crude oil'), ('one', 'Flat (no change)'),
              ('nuc', 'Nuclear'), ('wnd', 'Wind'), ('sol', 'Solar'), ('hyd', 'Hydro'), ('ore', 'Other renewables')]
MAP_FUELLAB1 = MAP_FUEL1 + len(FUEL_EXTRA)   # label-only fuel rows
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
    ('Price position (section 2)', 'ppos'), ('VAT rate', 'vr'),
    ('EF (tCO2/GJ)', 'ef'), ('Base-year fuel use (ktoe)', 'f0'),
]
IC = {code: L(j) for j, (_, code) in enumerate(INPUT_COLS, 1)}

# ---------------------------------------------------------------- Inputs_prices sheet (one row per price fuel)
IP_R0 = 5
IP_R1 = IP_R0 + NP - 1
HIST_YEARS = [2021, 2022, 2023, 2024]   # years in the IMF price dataset (Prices_dom)
IP_COLS = [
    ('Price code (fuel.sector)', 'key'), ('Label', 'lab'), ('Fuel', 'fuel'), ('Sector', 'sec'),
    ('Position (section 2 row)', 'pos'), ('International price position (gp row)', 'gpos'),
    ('GJ per price unit', 'gj'), ('Retail price from data? (1 = coal, gas, biomass)', 'rpd'),
    ('Raw pass-through coefficient (mit.ps, base year)', 'pccr'), ('Bucketed pass-through', 'pccb'),
    ('Chosen pass-through (Settings: price controls)', 'pcc'), ('Margin (mit.mar, base year, nominal price unit)', 'mar'),
    ('Domestic production cost (base year, nominal $/GJ)', 'pcost'),
    ('Fixed supply cost fixsp (real $/GJ)', 'fixsp'),
    ('VAT applies to final consumers? (assumption: residential and all-sector oil products)', 'vatc'),
    ('VAT rate assumption = general VAT rate (VAT_WEO) x consumer flag, used only where the data cell is blank', 'vrx'),
    ('VAT rate used (data where filled, 0 included; else assumption), last historical year', 'vr'),
    ('Supply cost, last historical year (real $/GJ)', 'spL'), ('Excise and other taxes, last historical year (real $/GJ)', 'txoL'),
] + [(f'Supply cost {y} (real $/GJ)', f'sp{y}') for y in HIST_YEARS] + \
    [(f'Excise and other taxes {y} (real $/GJ)', f'txo{y}') for y in HIST_YEARS]
IPC = {code: L(j) for j, (_, code) in enumerate(IP_COLS, 1)}

# ---------------------------------------------------------------- Mitigation layout
PARAM_COLS = ['D', 'E', 'F', 'G']       # hidden by default
COL_OUTLINE = {c: 1 for c in 'BCDEFGIJK'}   # column groups, one level (v0.16); plus 2030-2039 per scenario group
COL_YEARS_GROUP = (2030, 2039)
DESC = {'H': 'Description', 'I': 'Unit', 'J': 'Source'}
COL_G0 = 11                             # K: code column of scenario 1; base year 2022 in L (as legacy)
NY = 1 + len(YEARS)
GW = NY + 2                             # code column + years + one blank spacer column (v0.16)
R_YEAR, R_SCEN, R_NAME, R_GDP, R_CPI, R_DEFL = 4, 5, 6, 7, 8, 9

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
for p_ in ETS_PARAMS:
    _pol(p_, p_, mt=MT_NAMES[p_])
for k, s in enumerate(COV_SECTORS):
    _pol(f'etscv.{s}', 'etscov', 'all', s, mt=MT_ROWS_FIXED['etscov0'] + k)   # ETS sector coverage (MTInputs)
for s in COV_SECTORS:
    _pol(f'etsc.{s}', 'etsc', 'all', s, kind='calc')                       # effective coverage (ETS on, started)
for g, *_ in ETS_GROUPS:
    _pol(f'etsbs.{g}', 'etsb.s', 'all', g, kind='const')                    # benchmark at the start year
for g, *_ in ETS_GROUPS:
    _pol(f'etsbt.{g}', 'etsb.t', 'all', g, kind='const')                    # benchmark at the target year
for g, *_ in ETS_GROUPS:
    _pol(f'etsb.{g}', 'etsb', 'all', g, kind='calc')                        # benchmark path (flat after target)
for s in COV_SECTORS:
    _pol(f'obr.{s}', 'obr', 'all', s, kind='calc')                          # output-based (OBR) share by sector
for s in COV_SECTORS:
    _pol(f'bco2.{s}', 'bco2', 'all', s, kind='data')                        # baseline CO2 (data from scenario 1)
for k_ in ('ets.vadj', 'ets.bce', 'ets.cap', 'ets.se', 'ets.est'):
    _pol(k_, k_, kind='calc')
_pol('ets.ovr', 'ets.ovr', kind='data')                                      # override prices (goal seek), $/tCO2
for k_ in ('ets.p', 'ets.pe', 'ets.rf'):
    _pol(k_, k_, kind='calc')
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
    _pol(f'reg.{s}', 'reg', 'all', s, kind='const')                         # regulations shadow price: placeholder 0
for s, *_ in FEEBATE_SECTORS:
    _pol(f'shps.{s}', 'shps', 'all', s, kind='calc')                        # shadow price by sector, $/tCO2
for s in COV_SECTORS:
    _pol(f'ssc.{s}', 'ssc', 'all', s, kind='calc')                          # share impacting efficiency
CONST_VALUES = {'fc.bio': True, 'ssc.adj': 1, **{f'reg.{s}': 0 for s, *_ in FEEBATE_SECTORS},
                **{f'etsbs.{g}': b0 for g, b0, _b1 in ETS_GROUPS}, **{f'etsbt.{g}': b1 for g, _b0, b1 in ETS_GROUPS}}
ETS_GROUP_OF = {'pow': 'pow', **{s[0]: s[6] for s in SUBSECTORS}}           # ETS sector -> benchmark group
ETS_SEMI_OF = {'pow': ETS_SEMI['pow'], **{s[0]: ETS_SEMI['tra' if s[2] == 'tra' else ('res' if s[0] == 'res' else 'ind')]
                                          for s in SUBSECTORS}}           # legacy sector table (foo, srv: industry)
POL_PARAMS = {f'obr.{s}': ([g for g, *_ in ETS_GROUPS].index(ETS_GROUP_OF[s]) + 1, ETS_SEMI_OF[s])
              for s in COV_SECTORS}                                         # D: benchmark group, E: semi-elasticity
DATA_VALUES = {}                                                            # 'data' rows: filled by refresh_bco2
SUBHEADS = {'cptraj': None, CP_PARAMS[0]: 'Carbon tax (MTInputs rows 18-21, 265, 266)',
            'fc.coa': 'Carbon tax coverage: fuels (MTInputs rows 23-29)',
            'sc.pow': 'Carbon tax coverage: sectors (MTInputs rows 31-47)',
            'fpr.yr0': 'Fuel price reform (MTInputs rows 140-169; increases in price units)',
            'D_NewETS': 'New ETS (MTInputs rows 84-97: cap relative to baseline, years, price override, volatility; '
                        'replaces the carbon tax in covered sectors)',
            'etscv.pow': 'ETS sector coverage (MTInputs rows 99-115) and effective coverage (ETS applied and started)',
            'etsbs.pow': 'ETS benchmarks by sector group (typed inputs, share of emission intensity allocated free; '
                         'linear from the start to the target year, flat after); OBR share by sector (D: group, E: '
                         'semi-elasticity, legacy)',
            'bco2.pow': 'Baseline CO2 by ETS sector (MtCO2, data: values of scenario 1 rows co2.sec, section 13; '
                        'refreshed by the build; check row bco2.chk)',
            'ets.vadj': 'ETS price: volatility adjustment, baseline covered emissions, cap, effective semi-elasticity, '
                        'fast estimate, override row (pasted values), price used, tax-equivalent price, revenue factor',
            'fb.yr0': 'Feebates (MTInputs rows 53-64; rates in USD/tCO2)',
            'fbc.pow': 'Feebates sector coverage (MTInputs rows 66-82)',
            'ssc.adj': 'Shadow prices: regulations (placeholder), by sector ($/tCO2 = feebates + regulations; legacy '
                       'rows 2345-2349) and share impacting efficiency (legacy rows 2402-2419)'}

B_POL = 11                              # band: 1. Policies
R_POL = {}                              # key -> row
_r = B_POL + 1
for key, *_ in POL:
    if SUBHEADS.get(key):
        _r += 1                         # sub-heading row before this entry
    R_POL[key] = _r
    _r += 1
R_CP = R_POL['cptraj']
B_PRI = _r + 1                          # 2. Retail energy prices (before new policies)
R_PRI = {}
R_PRI['head.int'] = B_PRI + 1           # sub-heading: international prices
for k, p in enumerate(PRI_PARAMS):
    R_PRI[p] = B_PRI + 2 + k
R_GP0 = B_PRI + 2 + len(PRI_PARAMS)     # gp rows (oil, coa, nga, one)
R_PRI['head.dom'] = R_GP0 + len(GP_ROWS)
R_PV0 = R_PRI['head.dom'] + 1           # sp, txo, rpb blocks over the 12 price fuels
PVARS = ['sp', 'txo', 'rpb', 'vat', 'etx', 'esub', 'esubpu']
PVOFF = {v: k * (NP + 1) for k, v in enumerate(PVARS)}
B_POW = R_PV0 + len(PVARS) * (NP + 1) + 1   # 3. Power sector: generation costs and power prices (v1.05)
PW_TECH = [('coa', 'Coal'), ('nga', 'Natural gas'), ('oop', 'Oil'), ('nuc', 'Nuclear'), ('wnd', 'Wind'),
           ('sol', 'Solar'), ('hyd', 'Hydro'), ('ore', 'Other renewables'), ('bio', 'Biomass')]   # legacy order
NT = len(PW_TECH)
PW_GROUPS = [('res', 'Residential'), ('ind', 'Industry')]     # end-user groups (legacy G1, G2)
PW_LAYOUT = [  # legacy sub-table letter, heading, items (kind, variable); kind: tech (9 rows), user (2), single (1)
    ('A', 'A. Inputs: capex time factor; generation and investment shares (INTERIM data, legacy baseline, until the '
          'engineer model); storage cost paths', [('tech', 'tcf'), ('tech', 'gns'), ('tech', 'phi'), ('single', 'cbat'),
                                                  ('single', 'cint'), ('single', 'obat'), ('single', 'cel')]),
    ('B', 'B. Amortised fixed costs of the existing stock ($/kWh): capital cost, vintage averages of capital and '
          'storage cost (Phi of the previous year)', [('tech', 'cax'), ('tech', 'caxav'), ('tech', 'stoav'),
                                                       ('tech', 'fix')]),
    ('C', 'C. Levelised fixed costs of new plants ($/kWh; for investment decisions, not for prices)', [('tech', 'lfx')]),
    ('D', 'D. Variable costs before new policies, and new policy cost per kWh generated ($/kWh)',
     [('tech', 'vbc'), ('tech', 'ccp')]),
    ('E', 'E. Total costs: generation cost by type, average generation cost, average new policy cost per kWh '
          'consumed ($/kWh)', [('tech', 'gnc'), ('single', 'gncav'), ('single', 'ccpav')]),
    ('G', 'G. End-user electricity prices (real $/kWh; supply cost lagged one year) and power revenues (USD million; '
          'consumption INTERIM data)', [('user', 'cons'), ('user', 'sc'), ('user', 'rppre'), ('user', 'rp'),
                                        ('user', 'sgap'), ('user', 'rvat'), ('user', 'rcarb'), ('user', 'rgap')]),
    ('J', 'J. Systems integration cost of variable renewables: storage ($/kWh)',
     [('single', 'vre'), ('single', 'cph'), ('single', 'msc')]),
]
PW_DATA = {'tcf', 'gns', 'phi', 'cbat', 'cint', 'obat', 'cel', 'cons'}     # data lookups (PowerPaths), no LAMBDA
PW_INTERIM = {'gns', 'phi', 'cons'}                                         # interim data until the engineer model
PW_BASE = {'caxav', 'stoav', 'sc'}                                          # own base-year formula
PW_SUMMARY = {'gncav', 'rp'}                                                # visible when rolled up
R_PW, PW_HEAD, PW_ROWS = {}, {}, []      # (var, member) -> row; letter -> heading row; (row, var, B, C, kind)
_r = B_POW + 1
for _let, _head, _items in PW_LAYOUT:
    PW_HEAD[_let] = _r
    _r += 1
    _single = False
    for _kind, _v in _items:
        if _kind == 'single':
            R_PW[(_v, '')] = _r
            PW_ROWS.append((_r, _v, '', '', _kind))
            _r += 1
            _single = True
            continue
        if _single:
            _r += 1
        _mem = PW_TECH if _kind == 'tech' else PW_GROUPS
        for _k, (_m, _n) in enumerate(_mem):
            R_PW[(_v, _m)] = _r + _k
            PW_ROWS.append((_r + _k, _v, _m if _kind == 'tech' else 'ecy', 'pow' if _kind == 'tech' else _m, _kind))
        _r += len(_mem) + 1
        _single = False
    if _single:
        _r += 1
R_POW_NOTE = _r
PW_KIND = {v: k for _l, _h, items in PW_LAYOUT for k, v in items}
PW_VARS = [v for _l, _h, items in PW_LAYOUT for _k, v in items]
PW_PARAMS = {  # Mitigation D:G of each power variable: ('key',) data key | ('t', col) Inputs_power by type |
    # ('g', col) Inputs_power by user group | ('p', key) PowerParams | ('c', value) constant
    'tcf': [('key',)], 'gns': [('key',), ('t', 'vsh')], 'phi': [('key',)], 'cbat': [('key',)], 'cint': [('key',)],
    'obat': [('key',)], 'cel': [('key',)], 'cons': [('key',)],
    'cax': [('t', 'cax0')], 'caxav': [], 'stoav': [('t', 'vre')], 'fix': [('t', 'acap'), ('t', 'afix')],
    'lfx': [('t', 'lcap'), ('t', 'lfix'), ('t', 'vre')], 'vbc': [('t', 'vfix'), ('t', 'ppos'), ('t', 'k')],
    'ccp': [('t', 'efk'), ('t', 'fpos'), ('t', 'pfpos'), ('t', 'kpu')], 'gnc': [],
    'gncav': [('c', 1)], 'ccpav': [('p', 'gen_per_cons')], 'vre': [('c', 1)],
    'cph': [('p', 'st_ratio'), ('p', 'dlf_ren')], 'msc': [('p', 'st_alloc'), ('p', 'st_marg_hours'), ('p', 'lt_start')],
    'sc': [('g', 'tmc')], 'rppre': [('g', 'pass'), ('g', 'pos'), ('g', 'vr')], 'rp': [], 'sgap': [('g', 'vr')],
    'rvat': [('g', 'vshare')], 'rcarb': [('c', 1)], 'rgap': [('c', 1)]}
PW_FMT = {'tcf': '0.000', 'gns': '0.000', 'phi': '0.000', 'cbat': '#,##0.0', 'cint': '#,##0.0', 'obat': '0.00',
          'cax': '#,##0', 'caxav': '#,##0', 'cons': '#,##0', 'rvat': '#,##0.0', 'rcarb': '#,##0.0', 'rgap': '#,##0.0',
          'vre': '0.000'}
VARS = ['ctxnew', 'ets', 'ntx', 'nce', 'atp', 'shp', 'ener', 'co2', 'rtx', 'rsub', 'rnew']   # per subsector, 8 fuels
REV_VARS = ['rtx', 'rsub', 'rnew']
NVAR = len(VARS)
VSTEP = NF + 1                          # 8 fuel rows + 1 blank row
SB = 1 + NVAR * VSTEP                   # subsector block: heading line + variables
VOFF = {v: 1 + k * VSTEP for k, v in enumerate(VARS)}   # first row of each variable relative to the heading
SEC_BAND, SEC_SUM, SUB_HEAD = {}, {}, {}
_r = R_POW_NOTE + 3
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
B_REV = R_PCT + 2                       # 12. Revenues
REV_ROWS = []                           # (key, var, fuel, sector) in order; groups = sector sections
for v_ in REV_VARS:
    REV_ROWS.append((f'{v_}.all', v_, 'all', 'all'))
    REV_ROWS += [(f'{v_}.{g}', v_, 'all', g) for _n, g, _t in SECTIONS]
REV_ROWS += [('rnet.all', 'rnet', 'all', 'all'), ('rtot.all', 'rtot', 'all', 'all'),
             ('rtot.ref', 'rtot.ref', 'all', 'all'), ('rtot.chg', 'rtot.chg', 'all', 'all')]
R_REV = {k: B_REV + 1 + i for i, (k, *_r) in enumerate(REV_ROWS)}
R_REV_NOTE = B_REV + 1 + len(REV_ROWS) + 1
B_EMI = R_REV_NOTE + 4                  # 13. Emissions: CO2 from fuel combustion
EMI_ROWS = ([('co2.all', 'co2', 'all', 'all')] + [(f'co2.{g}', 'co2', 'all', g) for _n, g, _t in SECTIONS]
            + [(f'co2.f.{f}', 'co2', f, 'all') for f, *_x in FUELS]
            + [('co2.ref', 'co2.ref', 'all', 'all'), ('co2.chg', 'co2.chg', 'all', 'all'),
               ('co2.pct', 'co2.pct', 'all', 'all')]
            + [(f'co2.sec.{s}', 'co2.sec', 'all', s) for s in COV_SECTORS]
            + [(k, k, 'all', 'all') for k in ('co2.ets', 'co2.cap', 'co2.gap', 'ets.next', 'bco2.chk')])
R_EMI = {k: B_EMI + 1 + i for i, (k, *_r) in enumerate(EMI_ROWS)}
R_EMI_NOTE = B_EMI + 1 + len(EMI_ROWS) + 1


def var_rows(var):
    """All (row, subsector, fuel) of a variable across the sector sections."""
    return [(SUB_HEAD[s] + VOFF[var] + i, s, f) for s, *_ in SUBSECTORS for i, (f, *_r2) in enumerate(FUELS)]


# Hidden parameter columns D:G for each variable (Inputs columns), and their labels.
VAR_PARAMS = {'ctxnew': ['ef', 'fpos', 'spos'], 'ets': ['ef', 'spos'], 'ntx': ['pfpos', 'gj'], 'nce': [],
              'atp': ['ppos', 'vr'], 'shp': ['fbpos', 'ef', 'spos'], 'ener': ['eY', 'eU', 'eF', 'a'],
              'co2': ['ef'], 'rtx': ['ppos'], 'rsub': ['ppos'], 'rnew': ['vr']}
PVAR_PARAMS = {'sp': ['fixsp', 'gpos', 'pos'], 'txo': ['txoL', 'spL', 'pcc', 'pos'], 'rpb': ['vr'], 'vat': ['vr'],
               'etx': [], 'esub': [], 'esubpu': ['gj']}
PARAM_LABELS = {'ppos': 'price position', 'vr': 'VAT rate', 'ef': 'EF', 'fpos': 'fuel position', 'spos': 'sector position',
                'pfpos': 'price-fuel position', 'gj': 'GJ per unit', 'fbpos': 'feebate sector position',
                'eY': 'eps_Y', 'eU': 'eps_U', 'eF': 'eps_F', 'a': 'alpha'}

# Variable dictionary (sheet Variables): code, label, unit, code suffix, source, legacy reference, param labels
VARIABLES = [
    ('gdp.pos.pct', 'Real GDP growth', '%', '', 'WEO (lookup by country)', 'legacy Mitigation row 2446'),
    ('infl', 'Inflation index: US CPI, ResultsYear = 1 (nominal -> real for domestic prices and nominal policy '
             'inputs)', 'index', '', 'WEO USA|pcpi: CPI(ResultsYear) / CPI(year)',
     'legacy Mitigation rows 456 / 725 (inflation index, 2026 = 100)'),
    ('defl', 'Deflator index: US GDP deflator, ResultsYear = 1 (nominal -> real for international prices)', 'index',
     '', 'WEO USA|ngdp_d: deflator(ResultsYear) / deflator(year)', 'legacy Mitigation row 467 (row 630 = 1 / it)'),
] + [(p, '=MT', '=MT', '', 'MTInputs (scenario column)', f'MTInputs row {MT_NAMES[p]}')
     for p in CP_PARAMS + PRI_PARAMS + ETS_PARAMS] + [
    ('cptraj', 'Carbon price trajectory used', '$/tCO2 real', '', 'Calculation (legacy rows 2248, 2251); x infl if '
     'NomorReal = Nominal', 'legacy Mitigation row 2251: egy.mit.cptraj.1'),
    ('ctcov', 'Carbon tax coverage (Apply tax?)', 'switch', '', 'MTInputs (scenario column)',
     'MTInputs MCov* rows 23-47; biomass has no switch (EF 0)'),
    ('fpr.yr0', 'Fuel price reform: starting year', 'year', '', 'MTInputs (scenario column)', 'MTInputs row 140'),
    ('fpr.yr1', 'Fuel price reform: target year', 'year', '', 'MTInputs (scenario column)', 'MTInputs row 141'),
    ('fpr.s', 'Fuel price reform: starting price increase', 'price unit', '', 'MTInputs (scenario column)',
     'MTInputs rows 142-155'),
    ('fpr.f', 'Fuel price reform: final price increase', 'price unit', '', 'MTInputs (scenario column)',
     'MTInputs rows 156-169'),
    ('fpr', 'Fuel price reform: price increase path', 'price unit', '', 'Calculation (assumed linear)', 'new'),
    ('etscov', 'ETS sector coverage (Apply?)', 'switch', '', 'MTInputs (scenario column)', 'MTInputs rows 99-115'),
    ('etsc', 'Effective ETS coverage (new ETS applied, started, sector covered)', '0/1', '', 'Calculation', 'new'),
    ('etsb.s', 'ETS benchmark at the start year (share of emission intensity allocated free)', 'share', '',
     'Typed input (assumption 1.0)', 'new: replaces the auctioned proportion (MTInputs rows 90-92)'),
    ('etsb.t', 'ETS benchmark at the target year (share of emission intensity allocated free)', 'share', '',
     'Typed input (assumption 0.8)', 'new: replaces the auctioned proportion (MTInputs rows 90-92)'),
    ('etsb', 'ETS benchmark path', 'share', '', 'Linear from D_ETSIntro to D_ETSOutro, flat after', 'new'),
    ('obr', 'ETS output-based (OBR) share = benchmark (auctioned share = 1 - OBR)', 'share', '',
     'MIN(1, MAX(0, benchmark of the sector group))', 'legacy rows 1863-1864 (one auctioned proportion)'),
    ('bco2', 'Baseline CO2 by ETS sector (data from scenario 1)', 'MtCO2', '', 'Values of scenario 1 co2.sec',
     'legacy rows 1888-1919 (baseline emissions)'),
    ('ets.vadj', 'ETS volatility adjustment (tax-equivalent price per $ of permit price)', 'factor', '',
     '(1 + carbon tax policy risk x impact) / (1 + ETS volatility x impact); Settings table',
     'legacy row 2191 (1/1.1); ETS+LTS rows 5-9'),
    ('ets.bce', 'Baseline covered emissions', 'MtCO2', '', 'SUMPRODUCT(bco2, etsc)', 'legacy row 1956'),
    ('ets.cap', 'ETS cap', 'MtCO2', '', 'Baseline covered emissions x (1 + change path); constant after the target '
     'year if D_ETSCapCont = Constant', 'legacy rows 1958-1960'),
    ('ets.se', 'Effective covered semi-elasticity (allocation effectiveness: auctioned 1, OBR 0.5)', 'per $/tCO2',
     '', 'Baseline-weighted semi-elasticities x (1 - 0.5 OBR)', 'legacy rows 1845, 1866-1867'),
    ('ets.est', 'ETS price, fast estimate', '$/tCO2 real', '', 'LN(cap / baseline) / (semi-elasticity x volatility '
     'adjustment)', 'legacy row 1868 (QuickEstimateOfETSPrices; linear there)'),
    ('ets.ovr', 'ETS price override (pasted values, e.g. from ets_goalseek_v0_1.py)', '$/tCO2 real', '',
     'Typed / pasted values', 'legacy Manual inputs O15:AB15 (OverrideOfETSPrices)'),
    ('ets.p', 'ETS permit price used', '$/tCO2 real', '', 'Override row if D_ETSPriceOverride = Yes (blank: carbon '
     'price path), else the fast estimate', 'legacy row 1873'),
    ('ets.pe', 'ETS tax-equivalent price (permit price x volatility adjustment)', '$/tCO2 real', '',
     'ets.p x ets.vadj', 'legacy row 2191 applied to the ETS price'),
    ('ets.rf', 'ETS revenue factor (permit price per $ of tax-equivalent price)', 'factor', '', '1 / ets.vadj',
     'new'),
    ('reg', 'Shadow price from regulations (placeholder)', 'USD/tCO2', '', 'Constant 0 (not yet)',
     'legacy rows 2345-2349 include regulations'),
    ('fb.yr0', 'Feebates: start date', 'year', '', 'MTInputs (scenario column)', 'MTInputs row 53'),
    ('fb.yr1', 'Feebates: target date', 'year', '', 'MTInputs (scenario column)', 'MTInputs row 54'),
    ('fb.s', 'Feebates: starting rate', 'USD/tCO2', '', 'MTInputs (scenario column)', 'MTInputs rows 56-59'),
    ('fb.t', 'Feebates: target rate', 'USD/tCO2', '', 'MTInputs (scenario column)', 'MTInputs rows 61-64'),
    ('fb', 'Feebates: rate path', 'USD/tCO2', '', 'Calculation (assumed linear)', 'legacy rows 2000-2004'),
    ('fbcov', 'Feebates sector coverage (Apply?)', 'switch', '', 'MTInputs (scenario column)',
     'MTInputs rows 66-82'),
    ('ssc.adj', 'Efficiency-margin adjustment, feebates', 'share', '', 'cpat_coded default (feebates 1.0)',
     'MTInputs rows 247-250 hold the adjustments for regulations (not yet used)'),
    ('shps', 'Shadow price by sector (feebates + regulations; the ETS acts through the price)', 'USD/tCO2', '',
     'Calculation', 'legacy Mitigation rows 2345-2349'),
    ('ssc', 'Share of shadow price impacting efficiency', 'share', '', 'Feebate coverage x adjustment',
     'legacy Mitigation rows 2403-2419: egy.mit.ssc.rod.1'),
    ('gp', 'International energy price (real; source and adjustment from MTInputs)', 'source unit', '',
     'Prices_int (nominal) x defl x High/Low factor after the last historical year',
     'legacy Mitigation rows 633-695 (sources by year)', 'commodity key', 'High factor', 'Low factor'),
    ('sp', 'Supply cost (pre-tax price)', '$/GJ real', 'a', 'Historical: Inputs_prices; then fixsp + floating '
     'part x gp(t) / gp(t-1)', 'legacy Mitigation row 2498: egy.mit.sp.ind.coa.a.1; cpat_coded prices.py', 'fixed supply cost',
     'gp position', 'price position'),
    ('txo', 'Excise and other taxes (before new policies)', '$/GJ real', 'a',
     'Historical: Inputs_prices; then fixed part txo_L x pcc + floating part (pass-through rule)',
     'legacy txo = fadtx + fixs + cs (+ nce); cpat_coded prices.py', 'txo last hist. year', 'sp last hist. year',
     'pass-through', 'price position'),
    ('rpb', 'Retail price before new policies', '$/GJ real', 'a', '(sp + txo) x (1 + VAT rate)',
     'legacy retail price rp without nce; cpat_coded prices.py', 'VAT rate'),
    ('ets', 'New ETS cost in the fuel price (tax-equivalent, auctioned part)', '$/GJ', 'a',
     'ETS tax-equivalent price x EF x effective ETS coverage x (1 - OBR share)',
     'legacy ets inside txo; cpat_coded prices/ets.py', 'EF', 'sector position'),
    ('vat', 'VAT payment before new policies', '$/GJ real', 'a', '(sp + txo) x VAT rate', 'legacy vat; cpat_coded '
     'prices.py', 'VAT rate'),
    ('etx', 'Existing excise and other taxes (positive part of txo)', '$/GJ real', 'a', 'max(txo, 0)',
     'legacy fadtx + positive cs (split here by sign)'),
    ('esub', 'Existing consumer subsidy (negative part of txo, positive number)', '$/GJ real', 'a', 'max(-txo, 0)',
     'legacy fixs + negative cs (split here by sign)'),
    ('esubpu', 'Existing consumer subsidy per price unit (input for a later per-fuel policy)', 'price unit real',
     'a', 'esub x GJ per price unit', 'new (units as MTInputs rows 142-169)', 'GJ per unit'),
    ('ctxnew', 'New carbon tax', '$/GJ', 'a', 'Carbon price x EF x fuel and sector coverage x (1 - ETS coverage)',
     'legacy Mitigation row 2510: egy.mit.ctxnew.ind.coa.a.1', 'EF', 'fuel position', 'sector position'),
    ('ntx', 'New excise tax (fuel price reform)', '$/GJ', 'a', 'Fuel price reform path / GJ per unit',
     'legacy Mitigation row 2512: egy.mit.ntx.ind.coa.a.1', 'price-fuel position', 'GJ per unit'),
    ('nce', 'Total new policy', '$/GJ', 'a', 'New carbon tax + new excise tax',
     'legacy Mitigation row 2514: egy.mit.nce.ind.coa.a.1'),
    ('atp', 'After-tax price', '$/GJ real', 'e', 'max(rpb of the price fuel + nce x (1 + VAT rate), 0.01)',
     'legacy Mitigation row 5416: egy.mit.atp.rod.gso.e.1 (legacy unit $/liter for liquids)', 'price position',
     'VAT rate'),
    ('shp', 'Shadow price on the efficiency margin', '$/GJ', '', '(Sector shadow price x share + ETS tax-equivalent '
     'price x ETS coverage x OBR share) x EF',
     'legacy Mitigation row 2377: egy.mit.shp.coa.pow.1', 'feebate sector position', 'EF', 'sector position'),
    ('ener', 'Fuel use', 'ktoe', 'e', 'Inputs (base year), CPAT eq. 3.3.3',
     'legacy Mitigation row 5430: egy.mit.ener.rod.gso.e.1', 'eps_Y', 'eps_U', 'eps_F', 'alpha'),
    ('co2', 'CO2 emissions from fuel combustion', 'MtCO2', 'e', 'Fuel use x PJ/ktoe x EF',
     'legacy CO2 emissions by sector and fuel (inventory-adjusted EFs)', 'EF'),
    ('co2.ref', 'CO2 emissions, scenario 1 (same year)', 'MtCO2', '', 'Lookup', 'new'),
    ('co2.chg', 'Change in CO2 emissions vs scenario 1', 'MtCO2', '', 'co2 - co2.ref', 'new'),
    ('co2.pct', 'Change in CO2 emissions vs scenario 1', '%', '', 'co2 / co2.ref - 1', 'new'),
    ('co2.sec', 'CO2 emissions by ETS sector', 'MtCO2', '', 'Sum over fuels (power: not yet modelled, 0)', 'new'),
    ('co2.ets', 'CO2 emissions covered by the new ETS', 'MtCO2', '', 'SUMPRODUCT(co2.sec, etsc)', 'legacy row 1959'),
    ('co2.cap', 'ETS cap', 'MtCO2', '', 'ets.cap (section 1)', 'legacy row 1960'),
    ('co2.gap', 'Covered emissions vs cap (0 = cap met)', '%', '', 'co2.ets / co2.cap - 1', 'legacy rows 1961-1963'),
    ('ets.next', 'ETS price proposal for the next iteration', '$/tCO2 real', '', 'p x (LN(cap/baseline) / '
     'LN(covered/baseline)) ^ convergence exponent (Settings); fast estimate if no price yet',
     'legacy PricesAdjustedByErrorFactor (row 1966)'),
    ('bco2.chk', 'Check: baseline CO2 data minus scenario 1 CO2 (0 = data up to date)', 'MtCO2', '',
     'SUM(bco2) - co2.ref', 'new'),
    ('rtx', 'Revenue from existing taxes (excise and other taxes + VAT)', 'USD m real', '',
     'Fuel use x PJ/ktoe x (etx + vat)', 'legacy revenue block (existing taxes)', 'price position'),
    ('rsub', 'Cost of existing consumer subsidies (positive = cost)', 'USD m real', '', 'Fuel use x PJ/ktoe x esub',
     'legacy subsidy cost', 'price position'),
    ('rnew', 'Revenue from new policies (carbon tax, excise, auctioned ETS, VAT on them)', 'USD m real', '',
     'Fuel use x PJ/ktoe x (ctxnew + ntx + ets x revenue factor + nce x VAT rate)', 'legacy new revenue',
     'VAT rate'),
    ('rnet', 'Net revenue from existing taxes and subsidies', 'USD m real', '', 'rtx - rsub', 'new'),
    ('rtot', 'Total revenue (net existing + new policies)', 'USD m real', '', 'rnet + rnew', 'new'),
    ('rtot.ref', 'Total revenue, scenario 1 (same year)', 'USD m real', '', 'Lookup', 'new'),
    ('rtot.chg', 'Change in total revenue vs scenario 1 (fiscal effect)', 'USD m real', '', 'rtot - rtot.ref',
     'new'),
    ('ener.chk', 'Check: totals by subsector and by fuel minus total fuel use (should be 0)', 'ktoe', '',
     'Check', 'new'),
    ('ener.ref', 'Fuel use, scenario 1 (same year)', 'ktoe', 'e', 'Lookup', 'new'),
    ('ener.pct', 'Change in fuel use vs scenario 1', '%', 'e', 'Calculation', 'new'),
    # 3. Power sector (v1.05)
    ('tcf', 'Capex time factor (capital cost relative to its level at factor 1)', 'factor', '', 'PowerPaths',
     'legacy Mitigation rows 3038-3046 (A7)', 'data key'),
    ('gns', 'Generation share (INTERIM: legacy baseline run until the engineer model)', 'share', '', 'PowerPaths',
     'legacy rows 3452-3460 (E0)', 'data key', 'counts in VRE share'),
    ('phi', 'New investment as a share of remaining capacity (INTERIM: legacy baseline run until the engineer model)',
     'share', '', 'PowerPaths', 'legacy rows 3169-3177 (B1)', 'data key'),
    ('cbat', 'Battery storage capital cost', '$/kWh storage', '', 'PowerPaths', 'legacy row 3887 (J2)', 'data key'),
    ('cint', 'Battery interface capital cost', '$/kW', '', 'PowerPaths', 'legacy row 3888 (J2)', 'data key'),
    ('obat', 'Battery storage operating cost', '$/kWh storage/yr', '', 'PowerPaths', 'legacy row 3891 (J2)',
     'data key'),
    ('cel', 'Long-term storage cost per marginal unit (electrolyser and storage capital, fixed and variable opex)',
     '$/kWh', '', 'PowerPaths', 'legacy rows 3900-3902 (J2)', 'data key'),
    ('cons', 'Electricity consumption (INTERIM: legacy baseline run until electricity demand is modelled)', 'GWh', '',
     'PowerPaths', 'legacy rows 3621, 3650 (G1, G2)', 'data key'),
    ('cax', 'Capital cost of new plants', '$/kW real', '', 'cax0 x capex time factor', 'legacy rows 3099-3107 (A13)',
     'cax0 (factor 1)'),
    ('caxav', 'Weighted average capital cost of the stock (vintages)', '$/kW real', '',
     'Base year cax; then caxav(t-1) x (1 - Phi(t-1)) + cax(t-1) x Phi(t-1)', 'legacy rows 3181-3189 (B2)'),
    ('stoav', 'Weighted average storage cost of the stock (vintages)', '$/kWh', '',
     'Base year msc x VRE flag; then stoav(t-1) x (1 - Phi(t-1)) + msc(t) x VRE flag x Phi(t-1)',
     'legacy rows 3225-3233 (B6)', 'VRE flag'),
    ('fix', 'Amortised fixed cost of the existing stock (capital, interest, decommissioning, storage, fixed O&M)',
     '$/kWh', '', 'caxav x (1/N + wacc/2)/(cf x 8,760) + decommissioning/N/(cf x 8,760) + fixed O&M + stoav',
     'legacy rows 3247-3255 (B8 = B3 + B4 + B5 + B6 + B7)', 'capex factor', 'other fixed cost'),
    ('lfx', 'Levelised fixed cost of new plants (for investment decisions)', '$/kWh', '',
     '(cax + decommissioning + transmission capex)/(cf x 8,760 x dlf) + fixed O&M + subsidy + msc x VRE flag',
     'legacy rows 3326-3334 (C7 = C1 + ... + C6)', 'capex factor', 'other fixed cost', 'VRE flag'),
    ('vbc', 'Variable cost before new policies (fuel and variable O&M)', '$/kWh', '',
     'Variable O&M + fixed fuel cost + rpb of the price fuel (section 2, $/GJ) x 0.0036/efficiency',
     'legacy rows 3374-3382 (D4; legacy uses a 5-year moving average of fuel prices)', 'var. O&M + fixed fuel',
     'price position', 'GJ per kWh'),
    ('ccp', 'New policy cost per kWh generated (carbon tax, ETS, fuel price reform)', '$/kWh', '',
     '(carbon price x fuel and power coverage x (1 - ETS coverage) + ETS tax-equivalent price x ETS coverage x (1 - '
     'OBR share)) x tCO2/kWh + fuel price reform x 0.0036/efficiency/GJ per unit',
     'legacy D carbon cost by generation type; documentation 3.4.1.7', 'tCO2/kWh', 'fuel position',
     'price-fuel position', 'kWh factor of reform'),
    ('gnc', 'Generation cost before new policies (cost recovery)', '$/kWh', '', 'fix + vbc',
     'legacy rows 3486-3494 (E3)'),
    ('gncav', 'Average generation cost (weighted by generation shares)', '$/kWh', '', 'SUMPRODUCT(gns, gnc)',
     'legacy row 3495 (E3 weighted average)', 'factor'),
    ('ccpav', 'Average new policy cost per kWh consumed', '$/kWh', '',
     'SUMPRODUCT(gns, ccp) x generation / consumption', 'legacy G1/G2 policy row (3614, 3643)',
     'generation / consumption'),
    ('vre', 'Variable renewable share of generation (wind, solar)', 'share', '', 'SUMPRODUCT(gns, VRE-share flag)',
     'legacy row 3907', 'factor'),
    ('cph', 'Short-term storage cost per marginal hour of storage', '$/kWh', '',
     '(battery capex + interface capex / kWh per kW)/(dlf of renewables x 8,760) + battery opex / 8,760',
     'legacy row 3905 (J2)', 'kWh storage per kW', 'dlf of renewables'),
    ('msc', 'Marginal storage cost of variable renewables (wind, solar, other renewables)', '$/kWh', '',
     'allocation x 18 x vre x cph + 2 x MAX(0, vre - 0.75)/(1 - 0.75)^2 x cel', 'legacy rows 3910-3912 (J1)',
     'short-term allocation', 'short-term hours factor', 'long-term start share'),
    ('sc', 'Supply cost of electricity (average generation cost of the previous year + transmission and '
           'distribution)', '$/kWh', '', 'gncav(t-1) + T&D add-on; base year gncav(t) + T&D',
     'legacy rows 3600, 3629', 'T&D add-on'),
    ('rppre', 'Electricity price before new policies (incl. VAT)', '$/kWh real', '',
     'Data up to the last historical year (x CPI index), then rppre(t-1) + pass-through x (sc - sc(t-1)) x (1 + VAT)',
     'legacy rows 3613, 3642 (telescoped: held components cancel)', 'pass-through', 'user-group position',
     'VAT rate'),
    ('rp', 'End-user electricity price', '$/kWh real', '', 'rppre + ccpav (no VAT on new policies, as legacy; '
     'rebate and power excise 0)', 'legacy rows 3617, 3646'),
    ('sgap', 'Subsidy gap: supply cost minus price before VAT (positive = price below cost)', '$/kWh', '',
     'sc - rppre / (1 + VAT rate)', 'new (legacy producer-subsidy and over/under-estimate rows)', 'VAT rate'),
    ('rvat', 'VAT revenue from electricity', 'USD m real', '', 'rppre x VAT/(1 + VAT) x consumption', 'new',
     'VAT share of price'),
    ('rcarb', 'New-policy revenue from power (carbon tax, ETS at the tax-equivalent price, fuel price reform)',
     'USD m real', '', 'ccpav x consumption', 'legacy power carbon revenue', 'factor'),
    ('rgap', 'Fiscal cost of the electricity subsidy gap', 'USD m real', '', 'sgap x consumption', 'new', 'factor'),
]
VR0 = 5
VR1 = VR0 + 200

LAMBDAS = {   # every calculated row calls one of these in the right column (LAST_YEAR)
    'PATH': (['year', 'y0', 'y1', 'start', 'target', 'cont'],
             'IF(year<y0,0,start+(target-start)/MAX(y1-y0,1)*IF(OR(year<=y1,cont),year-y0,y1-y0))'),
    'ETSPRICE': (['apply', 'year', 'start_year', 'override', 'override_price', 'estimate', 'carbon_price'],
                 'IF(AND(LEFT(apply,3)="Yes",year>=start_year),IF(LEFT(override,3)="Yes",IF(override_price="",'
                 'carbon_price,override_price),estimate),0)'),
    'OBRSHARE': (['benchmark'], 'MIN(1,MAX(0,benchmark))'),
    'VOLADJ': (['volatility', 'impact', 'ct_risk'],
               '(1+ct_risk*INDEX(Settings!$C$18:$F$18,MATCH(SUBSTITUTE(impact,"*",""),Settings!$C$16:$F$16,0)))'
               '/(1+INDEX(Settings!$C$17:$F$17,MATCH(SUBSTITUTE(volatility,"*",""),Settings!$C$16:$F$16,0))'
               '*INDEX(Settings!$C$18:$F$18,MATCH(SUBSTITUTE(impact,"*",""),Settings!$C$16:$F$16,0)))'),
    'COVERED': (['emissions', 'coverage'], 'SUMPRODUCT(emissions,coverage)'),
    'ETSCAP': (['apply', 'year', 'y0', 'y1', 'change_start', 'change_target', 'cont', 'base', 'cap_prev'],
               'IF(AND(LEFT(apply,3)="Yes",year>=y0),IF(AND(year>y1,LEFT(cont,8)="Constant"),cap_prev,'
               'base*(1+change_start+(change_target-change_start)/MAX(y1-y0,1)*(MIN(year,y1)-y0))),0)'),
    'ETSSEMI': (['emissions', 'coverage', 'semi', 'obr', 'obr_eff'],
                'IF(SUMPRODUCT(emissions,coverage)>0,SUMPRODUCT(emissions,coverage,semi,1-(1-obr_eff)*obr)'
                '/SUMPRODUCT(emissions,coverage),0)'),
    'ETSESTIMATE': (['apply', 'year', 'start_year', 'cap', 'base', 'semi', 'vadj'],
                    'IF(AND(LEFT(apply,3)="Yes",year>=start_year,cap>0,cap<base,semi<0,vadj>0),'
                    'LN(cap/base)/(semi*vadj),0)'),
    'TAXEQUIV': (['price', 'vadj'], 'price*vadj'),
    'REVFACTOR': (['vadj'], 'IF(vadj>0,1/vadj,0)'),
    'ETSCOVER': (['apply', 'year', 'start_year', 'covered'],
                 'IF(AND(LEFT(apply,3)="Yes",year>=start_year),IF(covered,1,0),0)'),
    'SHADOWSECTOR': (['feebate', 'regulation'], 'feebate+regulation'),
    'SHADOWSHARE': (['coverage', 'adjustment'], 'IF(coverage,1,0)*adjustment'),
    'SUPPLYCOST': (['year', 'last_hist', 'hist_sp', 'fix_sp', 'sp_prev', 'gp_now', 'gp_prev'],
                   'IF(year<=last_hist,hist_sp,fix_sp+(sp_prev-fix_sp)*gp_now/gp_prev)'),
    'OTHERTAX': (['year', 'last_hist', 'hist_txo', 'txo_last', 'sp_last', 'sp_now', 'pcc'],
                 'IF(year<=last_hist,hist_txo,txo_last*pcc+IF(txo_last*(1-pcc)>=0,MAX(txo_last*(1-pcc),'
                 '(sp_last-sp_now+txo_last*(1-pcc))*(1-pcc)),(sp_last-sp_now+txo_last*(1-pcc))*(1-pcc)))'),
    'RETAILPRICE': (['sp', 'txo', 'vat_rate'], '(sp+txo)*(1+vat_rate)'),
    'VATPAY': (['sp', 'txo', 'vat_rate'], '(sp+txo)*vat_rate'),
    'TAXPART': (['txo'], 'MAX(txo,0)'),
    'SUBSIDYPART': (['txo'], 'MAX(-txo,0)'),
    'PERUNIT': (['value_gj', 'gj_per_unit'], 'value_gj*gj_per_unit'),
    'CARBONTAX': (['price', 'ef', 'fuel_cov', 'sector_cov', 'ets_cov'], 'price*ef*fuel_cov*sector_cov*(1-ets_cov)'),
    'ETSCOST': (['price', 'ef', 'ets_cov', 'obr'], 'price*ef*ets_cov*(1-obr)'),
    'NEWEXCISE': (['increase', 'gj_per_unit'], 'increase/gj_per_unit'),
    'NEWPOLICY': (['carbon_tax', 'ets', 'excise'], 'carbon_tax+ets+excise'),
    'POSTTAX': (['pre_policy_p', 'new_policy', 'vat_rate'], 'MAX(pre_policy_p+new_policy*(1+vat_rate),0.01)'),
    'SHADOWEFF': (['sector_price', 'share', 'ets_price', 'ets_cov', 'obr', 'ef'],
                  '(sector_price*share+ets_price*ets_cov*obr)*ef'),
    'REVENUE': (['fuel_use', 'pj_per_ktoe', 'rate'], 'fuel_use*pj_per_ktoe*rate'),
    'EMISSIONS': (['fuel_use', 'pj_per_ktoe', 'ef'], 'fuel_use*pj_per_ktoe*ef'),
    'NEWREVRATE': (['carbon_tax', 'excise', 'ets', 'rev_factor', 'new_policy', 'vat_rate'],
                   'carbon_tax+excise+ets*rev_factor+new_policy*vat_rate'),
    'FUELUSE': (['f_prev', 'p_now', 'p_prev', 'shp_now', 'shp_prev', 'gdp_g', 'eps_y', 'eps_u', 'eps_f', 'alpha'],
                'f_prev*(1/(1+alpha))^(1+eps_u)*(1+gdp_g)^eps_y*(p_now/p_prev)^eps_u'
                '*((p_now+shp_now)/(p_prev+shp_prev))^(eps_f*(1+eps_u))'),
    # 3. Power sector (v1.05)
    'CAPEX': (['capex_base', 'time_factor'], 'capex_base*time_factor'),
    'VINTAGE': (['avg_prev', 'new_value', 'share_prev'], 'avg_prev*(1-share_prev)+new_value*share_prev'),
    'AMORTISED': (['capex_avg', 'capex_factor', 'other_fixed', 'storage_avg'],
                  'capex_avg*capex_factor+other_fixed+storage_avg'),
    'LEVELISED': (['capex', 'capex_factor', 'other_fixed', 'vre_flag', 'storage'],
                  'capex*capex_factor+other_fixed+vre_flag*storage'),
    'VARCOST': (['fixed_part', 'fuel_price', 'gj_per_kwh'], 'fixed_part+fuel_price*gj_per_kwh'),
    'GENCOST': (['fixed_cost', 'variable_cost'], 'fixed_cost+variable_cost'),
    'WEIGHTED': (['shares', 'values', 'factor'], 'SUMPRODUCT(shares,values)*factor'),
    'STORAGEHOUR': (['battery', 'interface', 'opex', 'kwh_per_kw', 'dlf'],
                    '(battery+interface/kwh_per_kw)/(dlf*8760)+opex/8760'),
    'STORAGECOST': (['alloc', 'hours', 'vre', 'cph', 'lt_start', 'lt_cost'],
                    'alloc*hours*vre*cph+2*MAX(0,vre-lt_start)/(1-lt_start)^2*lt_cost'),
    'POWERSUPPLY': (['gen_cost', 'tmc'], 'gen_cost+tmc'),
    'POWERPRICE': (['year', 'last_hist', 'hist_price', 'price_prev', 'pass', 'sc_now', 'sc_prev', 'vat_rate'],
                   'IF(year<=last_hist,hist_price,price_prev+pass*(sc_now-sc_prev)*(1+vat_rate))'),
    'ENDUSERPOWER': (['pre_policy_p', 'policy_cost'], 'pre_policy_p+policy_cost'),
    'SUBSIDYGAP': (['supply_cost', 'pre_policy_p', 'vat_rate'], 'supply_cost-pre_policy_p/(1+vat_rate)'),
    'POWERREV': (['rate', 'factor', 'consumption'], 'rate*factor*consumption'),
}
LAMBDA_NOTES = {
    'PATH': 'Policy path: 0 before y0, linear from start (y0) to target (y1); after y1 linear if cont, else flat',
    'ETSPRICE': 'ETS permit price ($/tCO2) once a new ETS applies: override row if the override switch is Yes (blank '
                'cell: carbon price path), else the fast estimate',
    'OBRSHARE': 'Output-based (OBR) share of an ETS sector: its benchmark, between 0 and 1',
    'VOLADJ': 'ETS volatility adjustment: (1 + carbon tax policy risk x impact) / (1 + ETS volatility x impact), '
              'labels looked up in the Settings table (legacy ETS+LTS)',
    'COVERED': 'Covered emissions: SUMPRODUCT(emissions by ETS sector, effective coverage)',
    'ETSCAP': 'ETS cap: baseline covered emissions x (1 + change path from start to target year); after the target '
              'year constant if the continuation is Constant, else the relative change is kept',
    'ETSSEMI': 'Effective covered semi-elasticity: baseline-weighted sector semi-elasticities x allocation '
               'effectiveness (auctioned 1, OBR obr_eff)',
    'ETSESTIMATE': 'Fast ETS price estimate: LN(cap / baseline covered emissions) / (semi-elasticity x volatility '
                   'adjustment), 0 if the cap does not bind',
    'TAXEQUIV': 'ETS tax-equivalent price: permit price x volatility adjustment',
    'REVFACTOR': 'ETS revenue factor: permit price per $ of tax-equivalent price (1 / volatility adjustment)',
    'ETSCOVER': 'Effective ETS coverage of a sector: 1 if the new ETS applies, has started and covers it',
    'SHADOWSECTOR': 'Shadow price by sector ($/tCO2): feebate rate + regulations (placeholder)',
    'SHADOWSHARE': 'Share of the shadow price impacting efficiency: coverage x efficiency-margin adjustment',
    'SUPPLYCOST': 'Supply cost: data up to the last historical year, then fixed part + floating part x gp(t)/gp(t-1)',
    'OTHERTAX': 'Excise and other taxes: data up to the last historical year, then fixed part txo_L x pcc + floating '
                'part (sp_L - sp + txo_L(1 - pcc)) x (1 - pcc), not below its last value if that is >= 0 (legacy)',
    'RETAILPRICE': 'Retail price before new policies: (supply cost + excise and other taxes) x (1 + VAT rate)',
    'VATPAY': 'VAT payment before new policies: (supply cost + excise and other taxes) x VAT rate',
    'TAXPART': 'Existing tax: positive part of excise and other taxes',
    'SUBSIDYPART': 'Existing consumer subsidy: negative part of excise and other taxes, as a positive number',
    'PERUNIT': 'Value per GJ converted to the price unit (MTInputs units: $/liter, $/bbl, $/GJ)',
    'CARBONTAX': 'New carbon tax: price x EF x fuel coverage x sector coverage, not where the new ETS covers',
    'ETSCOST': 'New ETS in the fuel price: tax-equivalent price x EF x effective ETS coverage x auctioned share '
               '(1 - OBR)',
    'NEWEXCISE': 'New excise (fuel price reform): increase per price unit / GJ per price unit',
    'NEWPOLICY': 'Total new policy: new carbon tax + ETS + new excise',
    'POSTTAX': 'After-tax price: retail price before new policies + new policies x (1 + VAT rate), at least 0.01',
    'SHADOWEFF': 'Shadow price on the efficiency margin: (sector shadow price x share impacting efficiency + ETS '
                 'tax-equivalent price x ETS coverage x OBR share) x EF',
    'REVENUE': 'Revenue (USD million real): fuel use (ktoe) x PJ per ktoe x rate ($/GJ)',
    'EMISSIONS': 'CO2 emissions (MtCO2): fuel use (ktoe) x PJ per ktoe x emission factor (tCO2/GJ)',
    'NEWREVRATE': 'New-policy revenue rate ($/GJ): carbon tax + new excise + ETS x revenue factor + VAT on new '
                  'policies (feebates are revenue-neutral)',
    'FUELUSE': 'Fuel use, CPAT documentation 3.3.3, shadow price on the efficiency margin (no Covid factor)',
    'CAPEX': 'Capital cost of new plants ($/kW): capital cost at factor 1 x capex time factor',
    'VINTAGE': 'Vintage average: last year\'s average x (1 - last year\'s investment share) + new value x last year\'s '
               'investment share (legacy B2, B6)',
    'AMORTISED': 'Amortised fixed cost ($/kWh): average capex x capex factor + other fixed costs + storage average',
    'LEVELISED': 'Levelised fixed cost of new plants ($/kWh): capex x capex factor + other fixed costs + storage if VRE',
    'VARCOST': 'Variable cost ($/kWh): variable O&M and fixed fuel + fuel price ($/GJ) x GJ per kWh',
    'GENCOST': 'Generation cost before new policies ($/kWh): fixed + variable',
    'WEIGHTED': 'Generation-weighted sum: SUMPRODUCT(shares, values) x factor',
    'STORAGEHOUR': 'Short-term storage cost per marginal hour ($/kWh): annualised battery and interface capex + opex',
    'STORAGECOST': 'Marginal storage cost of variable renewables ($/kWh): short-term (alloc x hours x VRE share x cph) '
                   '+ long-term (2 x MAX(0, VRE share - start)/(1 - start)^2 x cost per unit)',
    'POWERSUPPLY': 'Supply cost of electricity: generation cost (previous year) + transmission and distribution',
    'POWERPRICE': 'Electricity price before new policies: data up to the last historical year, then the previous price '
                  '+ pass-through x change in supply cost x (1 + VAT)',
    'ENDUSERPOWER': 'End-user electricity price: price before new policies + average new policy cost (no VAT on it)',
    'SUBSIDYGAP': 'Subsidy gap ($/kWh): supply cost - price before VAT',
    'POWERREV': 'Power revenue (USD million): rate ($/kWh) x factor x consumption (GWh)',
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
F_BASE, F_UNUSED, F_CHECK = fill('DDD9C4'), fill('F2F2F2'), fill('FFF2CC')
F_LIGHT = fill('EEECE1')                  # light beige: internally copy-pasteable block, not draggable to white/base
F_LAMBDA = F_LIGHT
F_CHANGED = fill('FFFF00')                # bright yellow: data changed by assumption
FONT = Font(name='Arial', size=9)
FONT_B = Font(name='Arial', size=9, bold=True)
FONT_T = Font(name='Arial', size=14, bold=True, color='FFFFFF')
FONT_BAND = Font(name='Arial', size=9, bold=True, color='FFFFFF')
FONT_SUM = Font(name='Arial', size=9, bold=True)
FONT_LAMBDA = Font(name='Arial', size=9, color='9C0006')
FONT_CHANGED = Font(name='Arial', size=9, bold=True, color='FF0000')


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
    changed = {(r[0], r[1]) for r in read_csv('prices_dom_changes.csv')[1:]}
    col_of = {h: j for j, h in enumerate(rows[0], 1)}
    row_of = {r[0]: i for i, r in enumerate(rows[1:], 6)}
    for cy, col in changed:
        c = ws.cell(row_of[cy], col_of[col])
        c.fill, c.font = F_CHANGED, FONT_CHANGED
    for k in row_of:
        if any(cy == k for cy, _c in changed):
            c = ws.cell(row_of[k], 1)
            c.fill, c.font = F_CHANGED, FONT_CHANGED
    put(ws, 'B1', f'Bright yellow / red = {len(changed)} cells changed or added from the corrected Egypt price block '
                  '(user, 2026-10-09; data/prices_dom_changes.csv)', FONT_CHANGED, F_CHANGED)
    ws.freeze_panes = 'F6'
    u = read_csv('energy_use.csv')
    data_sheet(wb, 'EnergyCons', 'Base-year energy use by subsector and fuel (ktoe)', u[0], u[1:],
               'legacy Mitigation rows 583-602 via kernel v1.6 Data_Energy (Egypt only so far)')
    w = read_csv('weo.csv')
    us = read_csv('weo_us.csv')                       # key, country, variable, unit, years..., source
    assert w[0][5:] == us[0][4:-1], 'WEO year columns differ'
    us_rows = [r[:4] + [r[-1]] + r[4:-1] for r in us[1:]]
    data_sheet(wb, 'WEO', 'Macro series by country (real GDP growth; US CPI and GDP deflator for real terms)',
               w[0], w[1:] + us_rows,
               'legacy Mitigation row 2446 via kernel v1.6 Data_Macro (Egypt only so far); US PCPI and NGDP_D: '
               'legacy Mitigation rows 455 and 452')
    pi = read_csv('prices_int.csv')
    data_sheet(wb, 'Prices_int', 'International energy prices by source, nominal USD (key = source|commodity)',
               pi[0], pi[1:], 'legacy Mitigation rows 633-695 (built from legacy Prices_int; IMF-WB* = average of '
               'IMF and WB; global gas = average of LNG, North America, Europe)')
    pa = read_csv('price_assumptions.csv')
    data_sheet(wb, 'PriceAssump', 'Regional price assumptions by country: baseline taxes, gas market', pa[0],
               pa[1:], 'legacy Prices_int regional assumptions (rows 62+)')
    t = read_csv('power_tech.csv')
    data_sheet(wb, 'PowerTech', 'Power generation types: costs, lifetime, efficiency, capacity factor, WACC (legacy '
               'Egypt values)', t[0], t[1:], 'extract_power_v0_1.py: legacy Mitigation rows 2977-3350 (A2, A7, A8, '
               'A9, A14-A16, B2, B4, D1, D2)')
    pp = read_csv('power_paths.csv')
    data_sheet(wb, 'PowerPaths', 'Power year paths: capex time factors, storage costs; INTERIM generation shares, '
               'investment shares and consumption (legacy baseline run)', pp[0], pp[1:],
               'extract_power_v0_1.py: legacy Mitigation rows 3038-3902')
    pm = read_csv('power_params.csv')
    data_sheet(wb, 'PowerParams', 'Power parameters: transmission and distribution, pass-through, storage',
               pm[0], pm[1:], 'extract_power_v0_1.py: legacy Mitigation rows 3596-3883')
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
    flags = {1: 'Baseline (stored once)', 2: 'Live (batch target)'}
    columns = [(g, gname, MT_SCENARIO_INPUTS[g]) for g, gname, _ in SCENARIOS] + \
        [(n, nm, edits) for n, nm, _run, edits in DEFINITIONS]
    runs = {n: run for n, _nm, run, _e in DEFINITIONS}
    for g, gname, edits in columns:
        col = MT_COL0 + g - 1
        cl = L(col)
        put(ws, f'{cl}{MT_ROW_RUN}', runs.get(g, flags.get(g)), FONT_B, F_INPUT if g in runs else None)
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
            elif name in edits or r in edits:
                val = edits[name] if name in edits else edits[r]
            else:
                val = used
            if val is None:
                continue
            fnt = Font(name='Arial', size=9, bold=True, color='C00000') if (name in edits or r in edits) else FONT
            put(ws, f'{cl}{r}', val, fnt, F_INPUT)
        ws.column_dimensions[cl].width = 16
    ws.column_dimensions['I'].width = 3
    put(ws, f'{L(MT_COL0)}3', 'Scenario columns: number in row 5, name in row 6; red = changed from the template '
                              'value. J = baseline, K = live policy scenario (Mitigation group 2).', FONT_B)
    put(ws, f'{L(MT_COL0 + 2)}3', 'Scenario definitions: row 4 Run? = Yes for the batch run (macro RunAllScenarios '
                                  'copies each into column K, runs it and stores the results in StoredResults). Copy '
                                  'the last definition one column to the right to add one.', FONT_B)
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
            if item.filename.startswith('xl/worksheets/sheet') and re.search(rb'outlineLevelRow="\d+"', data) \
                    and b'outlineLevelCol' not in data:          # present when the workbook was saved before
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
    for i, (code, name) in enumerate(FUEL_EXTRA):
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


SET_LOG = 21                            # Settings: version log title row


def mt_setting(name):
    """Global setting read from MTInputs scenario 1 (column J) by NameOfParameter."""
    return f'=INDEX(MTInputs!${L(MT_COL0)}$8:${L(MT_COL0)}$415,MATCH("{name}",MTInputs!$H$8:$H$415,0))'


def build_settings(wb):
    ws = wb.create_sheet('Settings')
    title(ws, 'Settings', 6)
    rows = [(4, 'Country code', 'EGY', True),
            (5, 'Country name', '=INDEX(Countries!$B$4:$B$1000,MATCH($C$4,Countries!$A$4:$A$1000,0))', False),
            (6, 'Income group (selects elasticities)',
             '=INDEX(Countries!$D$4:$D$1000,MATCH($C$4,Countries!$A$4:$A$1000,0))', False),
            (7, 'Base year', BASE_YEAR, True),
            (8, 'Results in real USD of year (MTInputs ResultsYear, scenario 1)', mt_setting('ResultsYear'), False),
            (9, 'Government energy price controls (MTInputs GovPriceControls, scenario 1)',
             mt_setting('GovPriceControls'), False),
            (10, 'Last historical price year (IMF price dataset)', HIST_YEARS[-1], True),
            (11, 'Gas market (PriceAssump, by country name)',
             '=INDEX(PriceAssump!$C$4:$C$400,MATCH($C$5,PriceAssump!$A$4:$A$400,0))', False),
            (12, 'Gas price commodity key (Prices_int)',
             '=IF($C$11="LNG","lng",IF($C$11="Europe","eur",IF($C$11="North Am","nam","glo")))', False),
            (13, 'PJ per ktoe (ktoe x PJ/ktoe x $/GJ = USD million; x tCO2/GJ = MtCO2)', 0.041868, True),
            (14, 'ETS: relative effectiveness of output-based allocation (fast estimate; auctioned = 1; legacy '
                 'hardcoded)', 0.5, True),
            (15, 'ETS: convergence exponent of the next-price proposal (legacy convergence factor, row 1966)', 0.5,
             True)]
    for r, lab, val, inp in rows:
        put(ws, f'B{r}', lab)
        put(ws, f'C{r}', val, FONT, F_INPUT if inp else None)
    put(ws, 'B16', 'ETS volatility table (legacy ETS+LTS rows 5-8; MTInputs labels D_ETSVolatility, D_ETSVolImpact)',
        FONT_B)
    for j, (lab, vol, imp) in enumerate([('Medium', 0.42, 0.5), ('High', 0.63, 0.75), ('Low', 0.28, 1 / 3),
                                         ('Zero', 0, 0)]):
        put(ws, f'{L(3 + j)}16', lab, FONT_B, F_INPUT)
        put(ws, f'{L(3 + j)}17', vol, FONT, F_INPUT, '0.00')
        put(ws, f'{L(3 + j)}18', imp, FONT, F_INPUT, '0.00')
    put(ws, 'B17', 'Assumed annual price volatility of the ETS')
    put(ws, 'B18', 'Abatement cost increase per unit of volatility (impact)')
    put(ws, f'B{SET_LOG}', 'Version log', FONT_B)
    for j, h in enumerate(['Version', 'Date', 'Description', 'Max abs regression diff'], 2):
        put(ws, f'{L(j)}{SET_LOG + 1}', h, FONT_B, F_INPUT)
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
                              'biomass carbon-tax coverage TRUE.', '0 (common codes vs v0.7, no feebate)'),
        ('0.9', '2026-10-08', 'Domestic price projection (legacy method) in section 2: international prices by '
                              'source, supply cost, excise and other taxes with pass-through, retail price before new '
                              'policies; real USD of ResultsYear (CPI and deflator indices); historical prices '
                              '2022-2024 from data; per subsector atp = rpb + nce x (1 + VAT).',
         'intended change: all prices and fuel use (see check report)'),
        ('0.10', '2026-10-08', 'Other oil products treated like the other oil products (pass-through and margin '
                               'from data); VAT-rate assumption where data blank (VAT_WEO for residential and oil '
                               'products); changed data marked bright yellow.',
         'intended change: prices with VAT, other oil products, fuel use'),
        ('0.11', '2026-10-09', 'Horizon 2040; food & forestry with services elasticities; new ETS (price = carbon '
                               'price inputs, replaces the carbon tax where it covers); sectoral shadow prices '
                               '(feebates + regulations placeholder); section 2: VAT payment, existing taxes and '
                               'subsidies, subsidy per price unit; LAMBDA in the right column of every calculated row; '
                               'separate history / projection column blocks; LegacyDiff sheet; 12. Revenues prepared.',
         '0 to 2035 on unaffected codes (see check report); foo and horizon changes intended'),
        ('0.12', '2026-10-09', 'Revenues: existing tax revenue, existing subsidy cost and new-policy revenue as '
                               'separate calculations per subsector and fuel; section 12 totals by sector, net '
                               'existing, total and change vs scenario 1.', '0 on all v0.11 codes'),
        ('0.13', '2026-10-09', 'CO2 emissions from fuel combustion per subsector and fuel (fuel use x EF); '
                               'section 13 totals by sector and fuel, change vs scenario 1.', '0 on all v0.12 codes'),
        ('0.14', '2026-10-09', 'Corrected Egypt price block (user); changed cells bright yellow in Prices_dom; VAT '
                               'assumption only where the data cell is blank.', 'intended: prices, taxes, '
                               'subsidies, fuel use, revenue, CO2 (see check report)'),
        ('0.15', '2026-10-09', 'Layout: columns B:C and I:K grouped (D:G nested), row group buttons below, opens rolled '
                               'up; Mitigation first tab at 75% zoom.', '0 (all codes vs v0.14)'),
        ('0.16', '2026-10-09', 'Layout: one level of column groups (B:G, I:K, 2030-2039 per scenario); light beige '
                               'for the LAMBDA column and the 2023-2024 price history block; blank column between scenario groups.',
         '0 (all codes vs v0.15)'),
        ('1.00', '2026-10-09', 'Renamed CPAT-AI-Mitigation-MVP and upgraded to v1.00 (content as v0.16).',
         '0 (all codes vs v0.16)'),
        ('1.01', '2026-10-09', 'MTOutputs sheet: key results per scenario by output code (copy-pasteable row blocks); '
                               'Charts sheet.', '0 (all Mitigation codes vs v1.00)'),
        ('1.02', '2026-10-09', 'New ETS with a cap: baseline covered emissions (data from scenario 1), fast price '
                               'estimate, override row and goal-seek script; benchmarks replace the auction share '
                               '(OBR part as a shadow price on the efficiency margin); volatility adjustment; cap '
                               'check in section 13.', '0 (all codes shared with v1.01; ETS scenarios change)'),
        ('1.03', '2026-10-09', 'Multiple scenarios: scenario definitions on MTInputs (combined packages), VBA module '
                               'CPATScenarios (batch run with ETS goal seek, results stored as values), sheets '
                               'StoredResults and ScenarioCompare; workbook is .xlsm.', '0 (all codes vs v1.02)'),
        ('1.04', '2026-10-10', 'First-run test for Excel: macro CheckBatchRun (lookups, baseline recalculated by the '
                               'host vs stored, batch rerun vs stored values) with its report on the new sheet '
                               'MacroCheck.', '0 (all codes vs v1.03)'),
        ('1.05', '2026-10-10', 'Power sector, price side: section 3 with legacy sub-tables A, B, C, D, E, G, J '
                               '(generation costs by type, average generation cost, storage cost of variable '
                               'renewables, residential and industrial electricity prices, power revenues); sheets '
                               'Inputs_power, PowerTech, PowerPaths, PowerParams; generation and investment shares and '
                               'consumption interim data; 3 MTOutputs indicators added.',
         '0 (all codes vs v1.04)')]
    for i, row in enumerate(log, SET_LOG + 2):
        for j, v in enumerate(row, 2):
            put(ws, f'{L(j)}{i}', v)
    ws.column_dimensions['B'].width = 60
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
            'ppos': ip_lookup('pos', f'${IC["pcode"]}{r}'),
            'vr': ip_lookup('vr', f'${IC["pcode"]}{r}'),
            'ef': (f'=INDEX(EF_GHG!$E$4:$E$1000,MATCH(Settings!$C$4&"|"&$B{r}&"|"&${IC["efsec"]}{r},'
                   f'EF_GHG!$A$4:$A$1000,0))'),
            'f0': (f'=INDEX(EnergyCons!$F$4:$F$5000,MATCH(Settings!$C$4&"|"&Settings!$C$7&"|"&$C{r}&"|"&$B{r},'
                   f'EnergyCons!$A$4:$A$5000,0))'),
        }
        for _, code in INPUT_COLS:
            fmt = {'vr': '0.0%', 'ef': '0.0000', 'f0': '#,##0.0',
                   'eY': '0.00', 'eU': '0.00', 'eF': '0.00', 'a': '0.0%'}.get(code)
            put(ws, f'{IC[code]}{r}', vals[code], FONT, F_INPUT if code in ('fuel', 'sub') else None, fmt)
    ws.freeze_panes = 'D5'
    ws.column_dimensions['A'].width = 10
    for j in range(2, len(INPUT_COLS) + 1):
        ws.column_dimensions[L(j)].width = 10
    ws.row_dimensions[4].height = 48


def ip_lookup(code, key):
    """Inputs_prices column `code` for price code `key` (a cell reference or formula)."""
    c = IPC[code]
    return f'=INDEX(Inputs_prices!${c}${IP_R0}:${c}${IP_R1},MATCH({key},Inputs_prices!$A${IP_R0}:$A${IP_R1},0))'


def build_inputs_prices(wb):
    """Data step for section 2: one row per price fuel. Historical prices converted to real USD of ResultsYear
    with the CPI index of their year (row 3) and to $/GJ; forecasting coefficients (margin, production cost,
    pass-through) of the base year; values of the last historical year."""
    ws = wb.create_sheet('Inputs_prices')
    title(ws, 'Inputs by price fuel: historical prices (real $/GJ) and forecasting coefficients (data step)',
          len(IP_COLS))
    put(ws, 'A3', 'CPI index of the column year (ResultsYear = 1) ->', FONT_B)
    for j, (h, _) in enumerate(IP_COLS, 1):
        c = put(ws, f'{L(j)}4', h, FONT_B, F_INPUT)
        c.alignment = Alignment(wrap_text=True, vertical='top')
    pcpi = 'MATCH("USA|pcpi",WEO!$A$4:$A$1000,0)'
    for y in HIST_YEARS:
        for v in ('sp', 'txo'):
            col = IPC[f'{v}{y}']
            put(ws, f'{col}4', y, FONT_B, F_INPUT, '0')
            put(ws, f'{col}3', (f'=INDEX(WEO!$F$4:$AZ$1000,{pcpi},MATCH(Settings!$C$8,WEO!$F$3:$AZ$3,0))'
                                f'/INDEX(WEO!$F$4:$AZ$1000,{pcpi},MATCH({col}$4,WEO!$F$3:$AZ$3,0))'), FONT, None,
                '0.0000')
    sp0, sp1 = IPC[f'sp{HIST_YEARS[0]}'], IPC[f'sp{HIST_YEARS[-1]}']
    tx0, tx1 = IPC[f'txo{HIST_YEARS[0]}'], IPC[f'txo{HIST_YEARS[-1]}']

    def dom(var, year, r):
        """IMF price dataset value (blank = 0) of variable var for this row's price code, nominal."""
        return (f'IFERROR(INDEX(Prices_dom!$A$6:$CU$2000,MATCH(Settings!$C$4&{year},Prices_dom!$A$6:$A$2000,0),'
                f'MATCH("mit.{var}."&$A{r},Prices_dom!$A$5:$CU$5,0))+0,0)')

    def blank(var, year, r):
        """TRUE if the dataset cell is empty (no value), as opposed to an explicit 0."""
        return (f'IFERROR(ISBLANK(INDEX(Prices_dom!$A$6:$CU$2000,MATCH(Settings!$C$4&{year},Prices_dom!$A$6:$A$2000,0),'
                f'MATCH("mit.{var}."&$A{r},Prices_dom!$A$5:$CU$5,0))),TRUE)')

    def vat_weo(year):
        """General VAT rate of the country (Prices_dom VAT_WEO), blank = 0."""
        return (f'IFERROR(INDEX(Prices_dom!$A$6:$CU$2000,MATCH(Settings!$C$4&{year},Prices_dom!$A$6:$A$2000,0),'
                f'MATCH("VAT_WEO",Prices_dom!$A$5:$CU$5,0))+0,0)')

    for i, (pc, lab) in enumerate(PRICES):
        r = IP_R0 + i
        fuel, sec = pc.split('.')
        c = lambda k: f'${IPC[k]}{r}'
        base_cpi = f'INDEX(${sp0}$3:${sp1}$3,MATCH(Settings!$C$7,${sp0}$4:${sp1}$4,0))'
        vals = {
            'key': pc, 'lab': lab, 'fuel': fuel, 'sec': sec, 'pos': i + 1, 'gpos': GP_POS[fuel],
            'gj': '=' + map_fuel('G', r).replace(f'$B{r}', c('fuel')),
            'rpd': 1 if fuel in RP_FROM_DATA else 0,
            'pccr': 1 if fuel == 'bio' else f'={dom("ps", "Settings!$C$7", r)}',
            'pccb': f'=IF({c("pccr")}<=0.25,0,IF({c("pccr")}<=0.5,0.5,1))',
            'pcc': f'=IF(LEFT(Settings!$C$9,6)="Manual",0.8,IF(LEFT(Settings!$C$9,4)="None",1,{c("pccb")}))',
            'mar': 0 if fuel == 'bio' else f'={dom("mar", "Settings!$C$7", r)}',
            'pcost': (f'=IFERROR(INDEX(Prices_dom!$A$6:$CU$2000,MATCH(Settings!$C$4&Settings!$C$7,'
                      f'Prices_dom!$A$6:$A$2000,0),MATCH("mit.{fuel}.prod.cost",Prices_dom!$A$5:$CU$5,0))+0,0)'
                      if fuel in ('coa', 'nga') else 0),
            'fixsp': f'=({c("mar")}+{c("pccb")}*{c("pcost")})*{base_cpi}/{c("gj")}',
            'vatc': 1 if pc in VAT_CONSUMER else 0,
            'vrx': f'={c("vatc")}*{vat_weo("Settings!$C$10")}',
            'vr': f'=IF({blank("vatrate", "Settings!$C$10", r)},{c("vrx")},{dom("vatrate", "Settings!$C$10", r)})',
            'spL': f'=INDEX(${sp0}{r}:${sp1}{r},MATCH(Settings!$C$10,${sp0}$4:${sp1}$4,0))',
            'txoL': f'=INDEX(${tx0}{r}:${tx1}{r},MATCH(Settings!$C$10,${tx0}$4:${tx1}$4,0))',
        }
        for y in HIST_YEARS:
            sc, tc = IPC[f'sp{y}'], IPC[f'txo{y}']
            vr_y = (f'IF({blank("vatrate", f"{tc}$4", r)},{c("vatc")}*{vat_weo(f"{tc}$4")},'
                    f'{dom("vatrate", f"{tc}$4", r)})')
            vals[f'sp{y}'] = f'={dom("sp", f"{sc}$4", r)}*{sc}$3/{c("gj")}'
            vals[f'txo{y}'] = (f'=IF({c("rpd")}=1,{dom("rp", f"{tc}$4", r)}/(1+{vr_y})-{dom("sp", f"{tc}$4", r)},'
                               f'{dom("txo", f"{tc}$4", r)})*{tc}$3/{c("gj")}')
        for _, code in IP_COLS:
            v = vals[code]
            fmt = '0.0%' if code == 'vr' else ('0.00' if code.startswith('pcc') else
                                                ('0' if code in ('pos', 'gpos', 'rpd') else '0.000'))
            changed = code in ('vatc', 'vrx') or (fuel == 'oop' and code in ('pccr', 'mar'))
            fmt = '0.0%' if code == 'vrx' else ('0' if code == 'vatc' else fmt)
            put(ws, f'{IPC[code]}{r}', v, FONT_CHANGED if changed else FONT,
                F_CHANGED if changed else (F_INPUT if not str(v).startswith('=') else None), fmt)
    notes = [
        'Units: $/GJ in real USD of ResultsYear (Settings C8). Data (Prices_dom) are nominal USD per price unit '
        '($/GJ coal, gas, biomass; $/liter gasoline, diesel, LPG, kerosene; $/bbl other oil products): x CPI index '
        'of the data year (row 3) / GJ per price unit.',
        'Historical excise and other taxes (legacy): coal, gas, biomass = retail price / (1 + VAT rate) - supply '
        'cost; oil products = txo of the dataset (legacy builds their retail price as sp + txo, x (1 + VAT rate)).',
        'Pass-through: raw = mit.ps of the base year (blank = 0; biomass = 1); bucketed: '
        '<= 0.25 -> 0, <= 0.5 -> 0.5, else 1; chosen: Bucketed (default), Manual = 0.8, None = 1.',
        'fixsp = (margin + bucketed pass-through x domestic production cost [coal, gas]) x CPI index of the base '
        'year / GJ per unit; margin of biomass = 0 (legacy hardcode).',
        'BRIGHT YELLOW / RED = DATA CHANGED BY ASSUMPTION (v0.10, user decisions): (1) other oil products take '
        'pass-through and margin from the dataset like the other oil products (legacy hardcodes 1 and 0; Egypt: '
        'pass-through 0, margin 7.95 $/bbl); (2) VAT rate: the dataset rate where it is filled, otherwise the '
        'general VAT rate VAT_WEO for residential coal and gas and all-sector oil products, 0 for power, industry '
        'and biomass. v0.14: the corrected Egypt block fills every VAT cell (explicit 0 for gas and oil products, '
        '14% residential coal; any VAT sits inside txo), so the assumption no longer applies to Egypt.',
        'VAT rate used = dataset mit.vatrate wherever the cell is filled, explicit 0 included; the assumption only '
        'for blank cells (same rule for each historical year).']
    for k, n in enumerate(notes):
        put(ws, f'A{IP_R1 + 2 + k}', n)
    ws.freeze_panes = 'C5'
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 22
    for j in range(3, len(IP_COLS) + 1):
        ws.column_dimensions[L(j)].width = 11
    ws.row_dimensions[4].height = 72


# ---------------------------------------------------------------- Inputs_power sheet (v1.05; one row per generation type)
IPW_R0 = 5
IPW_R1 = IPW_R0 + NT - 1
IPW_COLS = [  # header, code, PowerTech column (None = formula)
    ('Generation type (code)', 'key', None), ('Name', 'name', 'name'),
    ('Capital cost at capex time factor 1 ($/kW, real)', 'cax0', 'cax0_usd_per_kw'),
    ('Lifetime N (years)', 'N', 'life_years'), ('Efficiency (NCV)', 'nu', 'efficiency'),
    ('Capacity factor', 'cf', 'capacity_factor'), ('WACC (country + technology premium)', 'wacc', 'wacc'),
    ('Decommissioning cost ($/kW)', 'dtc', 'dtc_usd_per_kw'),
    ('Transmission capex, levelised ($/kW)', 'tcx', 'tcx_usd_per_kw'), ('Fixed O&M ($/kWh)', 'opf', 'opf_usd_per_kwh'),
    ('Variable O&M ($/kWh)', 'vop', 'vop_usd_per_kwh'), ('Fixed fuel cost ($/kWh; nuclear)', 'ffix',
                                                            'fuel_fixed_usd_per_kwh'),
    ('Renewable subsidy ($/kWh; negative = subsidy)', 'rns', 'rns_usd_per_kwh'),
    ('Price fuel (section 2 price code; blank = no fuel)', 'pcode', 'price_code'),
    ('Carries the storage cost (VRE flag: wind, solar, other renewables)', 'vre', 'vre'),
    ('Counts in the VRE share (wind, solar; legacy row 3907)', 'vsh', 'vre_share'),
    ('Discounted lifetime dlf = (1 - (1+wacc)^-N) / (1 - (1+wacc)^-1)', 'dlf', None),
    ('Amortised capex factor (1/N + wacc/2) / (cf x 8,760)', 'acap', None),
    ('Other amortised fixed cost: dtc / N / (cf x 8,760) + fixed O&M ($/kWh)', 'afix', None),
    ('Levelised capex factor 1 / (cf x 8,760 x dlf)', 'lcap', None),
    ('Other levelised fixed cost: (dtc + tcx) x levelised factor + fixed O&M + subsidy ($/kWh)', 'lfix', None),
    ('Variable O&M + fixed fuel cost ($/kWh)', 'vfix', None),
    ('Fuel per kWh: 0.0036 / efficiency (GJ/kWh; 0 without a price fuel)', 'k', None),
    ('Price position (section 2 row; 1 without a price fuel)', 'ppos', None),
    ('EF (tCO2/GJ; power sector)', 'ef', None), ('Carbon intensity EF x GJ per kWh (tCO2/kWh)', 'efk', None),
    ('Fuel position (carbon-tax coverage)', 'fpos', None), ('Price-fuel position (fuel price reform)', 'pfpos', None),
    ('GJ per price unit', 'gj', None),
    ('Fuel price reform per kWh: GJ per kWh / GJ per price unit (0 without a price fuel)', 'kpu', None),
]
IPWC = {code: L(j) for j, (_h, code, _s) in enumerate(IPW_COLS, 1)}
IPW_G0 = IPW_R1 + 5                     # user-group table: CPI row G0 - 2, header G0 - 1, rows G0..
IPW_G1 = IPW_G0 + len(PW_GROUPS) - 1
IPWG_COLS = [('User group', 'grp'), ('Label', 'lab'), ('Position', 'pos'),
             ('Transmission and distribution ($/kWh)', 'tmc'), ('Pass-through of supply-cost changes', 'pass'),
             ('VAT rate, last historical year (data; blank = 0)', 'vr'), ('VAT share of the price: VAT / (1 + VAT)',
                                                                         'vshare')] + \
    [(f'Retail price {y} (real $/kWh, incl. VAT)', f'rp{y}') for y in HIST_YEARS]
IPWG = {code: L(j) for j, (_h, code) in enumerate(IPWG_COLS, 1)}


def pw_param(name):
    """PowerParams value by key."""
    return f'INDEX(PowerParams!$B$4:$B$50,MATCH("{name}",PowerParams!$A$4:$A$50,0))'


def build_inputs_power(wb):
    """Data step for section 3: one row per generation type (PowerTech data and the constants derived from them,
    so that each Mitigation formula needs at most four parameters) and one row per end-user group (T&D add-on,
    pass-through, VAT rate, historical electricity prices in real $/kWh)."""
    ws = wb.create_sheet('Inputs_power')
    title(ws, 'Inputs by generation type and end-user group (data step for section 3; no calculations of results)',
          len(IPW_COLS))
    put(ws, 'A3', 'One row per generation type (PowerTech data, legacy order) and derived constants; below, one row '
                  'per end-user group. Mitigation section 3 reads these columns in its hidden parameter columns D:G.',
        FONT)
    for j, (h, _c, _s) in enumerate(IPW_COLS, 1):
        c = put(ws, f'{L(j)}4', h, FONT_B, F_INPUT)
        c.alignment = Alignment(wrap_text=True, vertical='top')
    tk = lambda hdr, r: (f'INDEX(PowerTech!$A$4:$Z$50,MATCH($A{r},PowerTech!$A$4:$A$50,0),'
                         f'MATCH("{hdr}",PowerTech!$A$3:$Z$3,0))')
    for i, (f, _n) in enumerate(PW_TECH):
        r = IPW_R0 + i
        c = lambda k: f'${IPWC[k]}{r}'
        nofuel = f'{c("pcode")}=""'
        vals = {
            'key': f,
            'dlf': f'=(1-(1+{c("wacc")})^-{c("N")})/(1-(1+{c("wacc")})^-1)',
            'acap': f'=(1/{c("N")}+{c("wacc")}/2)/({c("cf")}*8760)',
            'afix': f'={c("dtc")}/{c("N")}/({c("cf")}*8760)+{c("opf")}',
            'lcap': f'=1/({c("cf")}*8760*{c("dlf")})',
            'lfix': f'=({c("dtc")}+{c("tcx")})*{c("lcap")}+{c("opf")}+{c("rns")}',
            'vfix': f'={c("vop")}+{c("ffix")}',
            'k': f'=IF({nofuel},0,0.0036/{c("nu")})',
            'ppos': f'=IF({nofuel},1,{ip_lookup("pos", c("pcode"))[1:]})',
            'ef': (f'=IFERROR(INDEX(EF_GHG!$E$4:$E$1000,MATCH(Settings!$C$4&"|"&$A{r}&"|pow",'
                   f'EF_GHG!$A$4:$A$1000,0)),0)'),
            'efk': f'={c("ef")}*{c("k")}',
            'fpos': (f'=IF({nofuel},1,INDEX(Mapping!$H${MAP_FUEL0}:$H${MAP_FUEL1},MATCH($A{r},'
                     f'Mapping!$A${MAP_FUEL0}:$A${MAP_FUEL1},0)))'),
            'pfpos': (f'=IF({nofuel},1,INDEX(Mapping!$E${MAP_PF0}:$E${MAP_PF1},MATCH({c("pcode")},'
                      f'Mapping!$A${MAP_PF0}:$A${MAP_PF1},0)))'),
            'gj': (f'=IF({nofuel},1,INDEX(Mapping!$G${MAP_FUEL0}:$G${MAP_FUEL1},MATCH($A{r},'
                   f'Mapping!$A${MAP_FUEL0}:$A${MAP_FUEL1},0)))'),
            'kpu': f'={c("k")}/{c("gj")}',
        }
        for _h, code, src in IPW_COLS:
            v = vals.get(code) or ('=' + tk(src, r) + ('&""' if code in ('name', 'pcode') else ''))
            fmt = {'cax0': '#,##0.0', 'N': '0', 'vre': '0', 'vsh': '0', 'ppos': '0', 'fpos': '0', 'pfpos': '0',
                   'dlf': '0.000', 'gj': '0.00', 'wacc': '0.0%', 'cf': '0.000', 'nu': '0.000'}.get(code, '0.000000')
            put(ws, f'{IPWC[code]}{r}', v, FONT, F_INPUT if code == 'key' else None, fmt)
    # end-user groups
    put(ws, f'A{IPW_G0 - 3}', 'End-user groups (legacy G1 residential, G2 industrial)', FONT_B)
    put(ws, f'A{IPW_G0 - 2}', 'CPI index of the column year (ResultsYear = 1) ->', FONT_B)
    for j, (h, code) in enumerate(IPWG_COLS, 1):
        c = put(ws, f'{L(j)}{IPW_G0 - 1}', h, FONT_B, F_INPUT)
        c.alignment = Alignment(wrap_text=True, vertical='top')
    pcpi = 'MATCH("USA|pcpi",WEO!$A$4:$A$1000,0)'
    for y in HIST_YEARS:
        col = IPWG[f'rp{y}']
        put(ws, f'{col}{IPW_G0 - 1}', y, FONT_B, F_INPUT, '0')
        put(ws, f'{col}{IPW_G0 - 2}', (f'=INDEX(WEO!$F$4:$AZ$1000,{pcpi},MATCH(Settings!$C$8,WEO!$F$3:$AZ$3,0))'
                                       f'/INDEX(WEO!$F$4:$AZ$1000,{pcpi},MATCH({col}${IPW_G0 - 1},WEO!$F$3:$AZ$3,0))'),
            FONT, None, '0.0000')

    def dom(var, year, r):
        return (f'IFERROR(INDEX(Prices_dom!$A$6:$CU$2000,MATCH(Settings!$C$4&{year},Prices_dom!$A$6:$A$2000,0),'
                f'MATCH("mit.{var}.ecy."&$A{r},Prices_dom!$A$5:$CU$5,0))+0,0)')
    for i, (g, lab) in enumerate(PW_GROUPS):
        r = IPW_G0 + i
        vals = {'grp': g, 'lab': lab, 'pos': i + 1, 'tmc': f'={pw_param("tmc|" + g)}', 'pass': f'={pw_param("pass")}',
                'vr': f'={dom("vatrate", "Settings!$C$10", r)}', 'vshare': f'=${IPWG["vr"]}{r}/(1+${IPWG["vr"]}{r})'}
        for y in HIST_YEARS:
            col = IPWG[f'rp{y}']
            vals[f'rp{y}'] = f'={dom("rp", f"{col}${IPW_G0 - 1}", r)}*{col}${IPW_G0 - 2}'
        for _h, code in IPWG_COLS:
            v = vals[code]
            put(ws, f'{IPWG[code]}{r}', v, FONT, F_INPUT if not str(v).startswith('=') else None,
                '0' if code == 'pos' else ('0.0%' if code in ('vr', 'vshare') else '0.0000'))
    notes = [
        'Derived constants follow PowerPrices_Method_v0_1.md: amortised fixed cost = caxav x acap + afix + stoav '
        '(legacy B3-B7 in one line); levelised fixed cost = cax x lcap + lfix + VRE x msc (legacy C1-C6); variable '
        'cost = vfix + rpb x k (legacy D4, with k = 0.0036 / efficiency).',
        'Types without a price fuel (nuclear, wind, solar, hydro, other renewables) take position 1 with factor 0 '
        '(k, efk, kpu = 0), so the price and policy terms vanish.',
        'Historical electricity prices: Prices_dom mit.rp.ecy.res / .ind (corrected Egypt block, nominal $/kWh incl. '
        'VAT) x CPI index of the year. VAT rate: Prices_dom mit.vatrate.ecy.* of the last historical year (blank = 0).']
    for k, n in enumerate(notes):
        put(ws, f'A{IPW_G1 + 2 + k}', n)
    ws.freeze_panes = 'C5'
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 18
    for j in range(3, len(IPW_COLS) + 1):
        ws.column_dimensions[L(j)].width = 11
    ws.row_dimensions[4].height = 96
    ws.row_dimensions[IPW_G0 - 1].height = 48


LEGACY_DIFFS = [  # area, legacy CPAT, this model, reason, effect, type
    ('Elasticities: food & forestry', 'Industry elasticities (cpat_coded EC_SECTORS)', 'Services (buildings) '
     'elasticities', 'User decision v0.11; legacy groups food & forestry with services as Commercial',
     'Food & forestry fuel response', 'Decision'),
    ('Other oil products: pass-through and margin', 'Hardcoded 1 and 0', 'From the IMF dataset like the other '
     'oil products (Egypt 0 and 7.95 $/bbl)', 'User decision v0.10: the legacy rule drove the retail price to the '
     '0.01 floor and fuel use 10x by 2030', 'Large: other oil products, total fuel use', 'Decision'),
    ('VAT rate where the dataset is blank', '0', 'General VAT rate VAT_WEO for residential coal and gas and '
     'all-sector oil products; 0 for power, industry, biomass. Only for blank cells (v0.14)', 'User decision v0.10',
     'None for Egypt since v0.14 (corrected data fill every VAT cell)', 'Assumption'),
    ('Egypt price data', 'Price block in the legacy release', 'Corrected Egypt block supplied by the user '
     '(2026-10-09; 291 cells, years 2019-2020 added); changed cells bright yellow in Prices_dom', 'User data '
     'correction', 'Taxes and subsidies of gas and oil products; small supply-cost changes', 'Data'),
    ('Oil-product share in the VAT of transport fuels', 'VAT rate x oil-product share', 'Share = 1',
     'Not yet extracted', 'None for Egypt today', 'Simplification'),
    ('New ETS price', 'Cap relative to baseline; quick estimate (linear in the semi-elasticity) and a VBA goal seek '
     'on an override row', 'Same cap; fast estimate in log form LN(cap/baseline)/(semi x effectiveness x volatility '
     'adjustment); override row pasted from ets_goalseek_v0_1.py (Python, same damped log-space step, no VBA); '
     'baseline covered emissions are data from scenario 1 (bco2)', 'No circular reference; legacy note "arguably '
     'exp/ln"', 'ETS scenarios', 'Method'),
    ('ETS allocation', 'One auctioned proportion for all sectors (MTInputs 90-92)', 'Benchmarks by sector group '
     '(pow, tra, res, ind; typed rows); OBR share = benchmark; MTInputs 90-92 not used', 'User decision v1.02 '
     '(benchmarks); the split afterwards is as legacy', 'ETS scenarios', 'Decision'),
    ('ETS volatility', 'Relative price ETS vs tax hardcoded 1.1 (row 1843); volatility inputs feed a factor not '
     'used', 'Volatility adjustment (1 + policy risk x impact)/(1 + volatility x impact) from MTInputs 95-97 '
     '(1/1.1 at Medium) on the tax-equivalent price and in the fast estimate', 'User suggestion v1.02 '
     '(volatility-dependent effectiveness)', 'ETS scenarios', 'Method'),
    ('Existing ETS and existing carbon taxes', 'Included where they exist', 'Not modelled', 'None for Egypt',
     'Other countries', 'Not yet'),
    ('Sectoral shadow prices', 'Feebates + non-auctioned ETS + regulations', 'Feebates + regulations '
     '(placeholder 0) + OBR part of the ETS', 'Regulations not yet', 'Regulation scenarios', 'Not yet'),
    ('Policy paths after the target year', 'Per instrument (carbon price: MTInputs switch)', 'Linear continuation '
     'for fuel price reform and feebates; carbon price keeps the switch; ETS benchmarks constant; ETS cap per '
     'MTInputs switch',
     'User decision v0.8 (one default)', 'Post-target years', 'Decision'),
    ('Scenario batch runs', 'Macros copy each MTInputs scenario column into the column used for calculation and '
     'store the outputs', 'Same idea: RunAllScenarios copies each definition (MTInputs L onwards, Run? = Yes) into '
     'the live column K, runs the ETS goal seek when needed, stores MTOutputs values in StoredResults; '
     'ScenarioCompare compares one year', 'User request v1.03', 'Multi-scenario analysis', 'Same'),
    ('Feebate grouping', 'Main grouping (buildings = residential rate)', 'Same: food & forestry and services take '
     'the residential rate; other energy use the industry rate', 'User decision v0.8', 'None', 'Same'),
    ('Biomass carbon-tax coverage', 'No switch', 'TRUE (EF 0)', 'User decision v0.8', 'None', 'Decision'),
    ('Margins and production costs', 'cpat_coded applies no inflation conversion', 'x CPI index of the base year',
     'Consistent real terms (assumption A4)', 'Coal and gas supply cost level', 'To check (bucket 5)'),
    ('Existing taxes and subsidies', 'txo = fixed tax (fadtx) + fixed subsidy (fixs) + floating part (cs)',
     'Split by the sign of txo: etx = max(txo, 0), esub = max(-txo, 0)', 'Simple and revenue-ready; same totals',
     'Presentation only', 'Simplification'),
    ('Subsidy and price-control phase-outs', 'MTInputs 123-138, 191-195', 'Not yet (phase-out factors 1)',
     'Later per-fuel policy input; esubpu prepared', 'Reform scenarios', 'Not yet'),
    ('Producer subsidies', 'In the supply cost', 'Left out', 'Zero for Egypt', 'Other countries', 'Not yet'),
    ('Ad-valorem baseline taxes; Europe LNG shift', 'Per Prices_int assumptions', 'Fixed taxes only; no shift',
     'Egypt: Fixed taxes, Global gas market', 'Other countries', 'Not yet'),
    ('Fuel-use equation', 'Includes Covid factor and additional policy efficiency gains (MTInputs 242-245)',
     'CPAT eq. 3.3.3 without them', 'Not yet', 'Small', 'Not yet'),
    ('Power quantities (engineer model)', 'Dispatch, investment logit, capacity, power emissions (tables H, I, K, L)',
     'Not yet: generation shares, investment shares and consumption are INTERIM data (legacy baseline run)',
     'Next bucket (v1.05 builds the price side first)', 'Power prices respond to fuel and carbon costs, not to the '
     'scenario\'s generation mix; no power emissions, power not in the totals of sections 11-13', 'Not yet'),
    ('Power: generation costs (v1.05)', 'Tables B3-B8 (amortised), C1-C7 (levelised), A5-A6 discount factors',
     'One row per type for amortised and for levelised fixed cost; discounted lifetime in closed form; constants '
     'derived in Inputs_power', 'Fewer, readable rows (PowerPrices_Method_v0_1.md S1-S3)', 'None (method test: '
     'legacy rows reproduced to 1e-15)', 'Simplification'),
    ('Power: fuel prices', '5-year moving average of fuel prices (A18; spot prices a dashboard option)',
     'Spot prices: section 2 retail price before new policies of the price fuel (coal and gas for power, oil, '
     'biomass)', 'Simplification S5; the legacy option "use spot fuel prices"', 'Fuel-price changes reach power '
     'costs one to four years earlier', 'Simplification'),
    ('Power: end-user prices', 'Supply cost + producer subsidy + other tax + over/under-estimate residual (held at '
     'the last historical year), then VAT', 'Telescoped: price(t) = price(t-1) + pass-through x change in supply '
     'cost x (1 + VAT) after the last historical year', 'Simplification S4: the held components cancel',
     'None (same numbers)', 'Simplification'),
    ('Power: historical electricity prices and VAT', 'Legacy vintage (residential 0.112 $/kWh incl. 15.6% VAT in '
     '2022)', 'Corrected Egypt block (residential 0.054, 0.048, 0.043 $/kWh nominal 2022-2024; VAT 0)', 'User data '
     '(2026-10-09)', 'Electricity price levels', 'Data'),
    ('Power: coal implicit cost and extra shadow price', 'D3, D6', 'Not built', 'Both 0 for Egypt (no coal in the mix)',
     'Other countries with coal', 'Not yet'),
    ('Power: new-policy revenue', 'Carbon revenue from power', 'ccpav x consumption; the ETS part valued at the '
     'tax-equivalent price (subsector revenues use the permit price on the auctioned part)', 'Simplification',
     'Power ETS revenue about 10% high at Medium volatility', 'Simplification'),
    ('Emissions', 'All GHGs and local pollutants; inventory-adjusted EFs; power and process emissions',
     'CO2 from fuel combustion only, IIASA EFs without inventory adjustment; no power, no process emissions',
     'User decision v0.13: CO2 first', 'Emission levels', 'Not yet'),
    ('Revenue structure', 'Revenue by instrument incl. electricity, producer subsidies, ETS allowances, '
     'GDP feedback', 'Three separate calculations per subsector and fuel: existing taxes (txo positive part + '
     'VAT), existing consumer subsidies, new policies (carbon tax, excise, auctioned ETS, VAT on them); no power, '
     'no producer subsidies, no GDP feedback', 'User request v0.12; scope of the prototype', 'Revenue levels',
     'Simplification'),
    ('Horizon and indices', 'To 2050; CPI held flat after 2031', 'To 2040; same indices; values after 2030 not a '
     'legacy target', 'User decision v0.9 / v0.11', 'Post-2030', 'Decision'),
    ('Data vintage', 'Proprietary price forecasts; IEA balances', 'NoPropData extracts; energy use via kernel v1.6; '
     'IIASA EFs without inventory adjustment', 'No proprietary data', 'Levels; legacy price forecasts #N/A',
     'Data'),
    ('Formula structure', 'Different formulas by block and year', 'One formula per variable block; LAMBDA in the '
     'right column (2040) with the history/projection IF inside', 'Auditability and copy-paste', 'None '
     '(numerically identical)', 'Design'),
]


def build_legacy_diff(wb):
    ws = wb.create_sheet('LegacyDiff')
    title(ws, 'Differences with legacy CPAT (decisions, assumptions, simplifications, not yet modelled)', 7)
    put(ws, 'A3', 'Keep this list current with each version. Type: Decision = user decision; Assumption = data '
                  'assumption (bright yellow in Inputs_prices); Simplification; Not yet; Data; Design; Same.', FONT)
    for j, h in enumerate(['#', 'Area', 'Legacy CPAT', 'This model', 'Why', 'Effect', 'Type'], 1):
        c = put(ws, f'{L(j)}4', h, FONT_B, F_INPUT)
        c.alignment = Alignment(wrap_text=True, vertical='top')
    for i, row in enumerate(LEGACY_DIFFS, 5):
        for j, v in enumerate((i - 4,) + row, 1):
            c = put(ws, f'{L(j)}{i}', v, FONT_CHANGED if row[-1] == 'Assumption' and j == 7 else FONT,
                    F_CHANGED if row[-1] == 'Assumption' and j == 7 else None)
            c.alignment = Alignment(wrap_text=True, vertical='top')
    for col, w in {'A': 4, 'B': 26, 'C': 36, 'D': 44, 'E': 40, 'F': 26, 'G': 14}.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'C5'


# ---------------------------------------------------------------- MTOutputs and Charts (v1.01)
MTO_INDICATORS = [  # code stem after the country code (scenario number appended), number format
    ('mit.ener.all.all.e', '#,##0'), ('mit.ener.tra.all.e', '#,##0'), ('mit.ener.bld.all.e', '#,##0'),
    ('mit.ener.ind.all.e', '#,##0'), ('mit.ener.oen.all.e', '#,##0'), ('mit.ener.pct.all.all.e', '0.0%'),
    ('mit.co2.all.all.e', '#,##0.0'), ('mit.co2.tra.all.e', '#,##0.0'), ('mit.co2.bld.all.e', '#,##0.0'),
    ('mit.co2.ind.all.e', '#,##0.0'), ('mit.co2.chg.all.all', '#,##0.0'), ('mit.co2.pct.all.all', '0.0%'),
    ('mit.rtx.all.all', '#,##0'), ('mit.rsub.all.all', '#,##0'), ('mit.rnew.all.all', '#,##0'),
    ('mit.rtot.all.all', '#,##0'), ('mit.rtot.chg.all.all', '#,##0'),
    ('mit.cptraj', '0.00'), ('mit.ets.p', '0.00'), ('mit.co2.ets.all.all', '#,##0.0'),
    ('mit.co2.cap.all.all', '#,##0.0'),
    ('mit.rpb.all.gso.a', '0.00'), ('mit.rpb.all.die.a', '0.00'), ('mit.rpb.res.nga.a', '0.00'),
    ('mit.atp.rod.gso.e', '0.00'),
    ('mit.rp.res.ecy', '0.0000'), ('mit.rp.ind.ecy', '0.0000'), ('mit.gncav', '0.0000'),
]
MTO_R0 = 6                              # first block header row
MTO_BLOCK = 1 + len(MTO_INDICATORS) + 1  # header + indicators + blank row
MTO_C0 = 6                              # F: first year column
MIT_RANGE = '$A$1:$ZZ$5000'             # search area on Mitigation (room for copied scenario groups)
CHARTS = [  # chart title, indicator code stem, y-axis title
    ('CO2 emissions from fuel combustion', 'mit.co2.all.all.e', 'MtCO2'),
    ('Fuel use, all subsectors', 'mit.ener.all.all.e', 'ktoe'),
    ('Fiscal effect vs scenario 1 (total revenue change)', 'mit.rtot.chg.all.all', 'USD million, real 2026'),
    ('Carbon price', 'mit.cptraj', 'USD/tCO2, real 2026'),
]


def mto_row(g, k):
    """Row of indicator k (0-based) in the block of scenario g."""
    return MTO_R0 + (g - 1) * MTO_BLOCK + 1 + k


def build_mtoutputs(wb):
    """MTOutputs: one row block per scenario; each row finds its output code on Mitigation (code column of the
    scenario = first column of row 5 holding the scenario number) and reads the year columns by offset."""
    ws = wb.create_sheet('MTOutputs')
    ny = len([BASE_YEAR] + YEARS)
    last = MTO_C0 + ny - 1
    title(ws, f'MTOutputs - key results by scenario (read from Mitigation by output code)', last)
    put(ws, 'A3', 'One block per scenario: the number in column E of the block header drives every row. To add a '
                  'scenario, copy the last block (header to blank row) and paste it directly below: the number goes '
                  'up by one. Rows: add a code stem in column B and copy a row\'s formulas. Values follow Mitigation.',
        FONT)
    for j, h in enumerate(['Output code', 'Code stem', 'Description (from Mitigation)', 'Unit', 'Scenario'], 1):
        put(ws, f'{L(j)}4', h, FONT_B, F_INPUT)
    for k in range(ny):
        col = L(MTO_C0 + k)
        put(ws, f'{col}4', '=Settings!$C$7' if k == 0 else f'={L(MTO_C0 + k - 1)}4+1', FONT_B,
            F_BASE if k == 0 else (F_LIGHT if k == ny - 1 else F_INPUT), '0')
    code_col = f'MATCH($E{{r}},Mitigation!$A$5:$ZZ$5,0)'
    for g, gname, _ in SCENARIOS:
        h = MTO_R0 + (g - 1) * MTO_BLOCK
        band(ws, h, '', last)
        put(ws, f'A{h}', 'Scenario block', FONT_BAND, F_BAND)
        put(ws, f'E{h}', 1 if g == 1 else f'=E{h - 1}+1', FONT_BAND, F_BAND, '0')
        put(ws, f'C{h}', (f'=INDEX(MTInputs!${L(MT_COL0)}${MT_ROW_NAME}:${MT_LAST}${MT_ROW_NAME},'
                          f'MATCH($E{h},MTInputs!${L(MT_COL0)}${MT_ROW_SCEN}:${MT_LAST}${MT_ROW_SCEN},0))'),
            FONT_BAND, F_BAND)
        for k, (stem, fmt) in enumerate(MTO_INDICATORS):
            r = h + 1 + k
            cc = code_col.format(r=r)
            find = f'MATCH($A{r},INDEX(Mitigation!{MIT_RANGE},0,{cc}),0)'
            put(ws, f'A{r}', f'=LOWER(Settings!$C$4)&"."&$B{r}&"."&$E{r}', FONT, F_CODE)
            put(ws, f'B{r}', stem, FONT, F_INPUT)
            put(ws, f'C{r}', f'=IFERROR(INDEX(Mitigation!$H$1:$H$5000,{find}),"code not found")')
            put(ws, f'D{r}', f'=IFERROR(INDEX(Mitigation!$I$1:$I$5000,{find}),"")')
            put(ws, f'E{r}', f'=E{r - 1}', FONT, None, '0')
            for kk in range(ny):
                col = L(MTO_C0 + kk)
                put(ws, f'{col}{r}', (f'=IFERROR(INDEX(Mitigation!{MIT_RANGE},{find},{cc}+1+{col}$4-Settings!$C$7),'
                                      f'"")'), FONT, F_BASE if kk == 0 else None, fmt)
        r = h + MTO_BLOCK - 1
        put(ws, f'E{r}', f'=E{r - 1}', Font(name='Arial', size=9, color='BFBFBF'), None, '0')
    ws.freeze_panes = f'{L(MTO_C0)}5'
    for col, w in {'A': 30, 'B': 22, 'C': 52, 'D': 12, 'E': 8}.items():
        ws.column_dimensions[col].width = w
    for k in range(ny):
        ws.column_dimensions[L(MTO_C0 + k)].width = 9
    ws.sheet_view.zoomScale = 85


def build_charts(wb):
    """Line charts of MTOutputs rows, one series per scenario block (add a series after adding a scenario)."""
    from openpyxl.chart import LineChart, Reference, Series
    ws = wb.create_sheet('Charts')
    title(ws, 'Charts - key results by scenario (from MTOutputs)', 20)
    put(ws, 'A3', 'One series per scenario block on MTOutputs. After adding a scenario block, add its row as a new '
                  'series (Select Data). Values in real 2026 USD where monetary.', FONT)
    src = wb['MTOutputs']
    ny = len([BASE_YEAR] + YEARS)
    cats = Reference(src, min_col=MTO_C0, max_col=MTO_C0 + ny - 1, min_row=4, max_row=4)
    stems = [s for s, _ in MTO_INDICATORS]
    for i, (ttl, stem, ytitle) in enumerate(CHARTS):
        ch = LineChart()
        ch.title, ch.y_axis.title, ch.x_axis.title = ttl, ytitle, 'Year'
        ch.height, ch.width = 7.5, 15
        k = stems.index(stem)
        for g, gname, _ in SCENARIOS:
            r = mto_row(g, k)
            ser = Series(Reference(src, min_col=MTO_C0, max_col=MTO_C0 + ny - 1, min_row=r, max_row=r),
                         title=f'Scenario {g}: {gname}')
            ch.series.append(ser)
        ch.set_categories(cats)
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.legend.position = 'b'
        ws.add_chart(ch, f'{"A" if i % 2 == 0 else "K"}{5 + (i // 2) * 16}')


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


def call(name, *args):
    assert len(args) == len(LAMBDAS[name][0]), name
    return f'{name}(' + ','.join(args) + ')'


def year_of(c):
    return BASE_YEAR + (c - COL_G0) % GW - 1


def calc_formula(var, r, col, prev, is_base, use_lambda):
    """Sector-section cell of variable var in row r (column col; prev = previous year column)."""
    o = lambda v: f'{col}{r - VOFF[var] + VOFF[v]}'          # same fuel, other variable, this column
    op = lambda v: f'{prev}{r - VOFF[var] + VOFF[v]}'        # same fuel, other variable, previous column
    etsc = f'INDEX({pol_range(col, "etsc.pow", len(COV_SECTORS))},'
    if var == 'ctxnew':
        a = (f'{col}${R_CP}', f'$D{r}', f'INDEX({pol_range(col, "fc.coa", NF)},$E{r})',
             f'INDEX({pol_range(col, "sc.pow", len(COV_SECTORS))},$F{r})', f'{etsc}$F{r})')
        return '=' + (call('CARBONTAX', *a) if use_lambda else f'{a[0]}*{a[1]}*{a[2]}*{a[3]}*(1-{a[4]})')
    obr = f'INDEX({pol_range(col, "obr.pow", len(COV_SECTORS))},'
    if var == 'ets':
        a = (f'{col}${R_POL["ets.pe"]}', f'$D{r}', f'{etsc}$E{r})', f'{obr}$E{r})')
        return '=' + (call('ETSCOST', *a) if use_lambda else f'{a[0]}*{a[1]}*{a[2]}*(1-{a[3]})')
    if var == 'ntx':
        a = (f'INDEX({pol_range(col, "fpr." + PRICE_FUELS[0][0], len(PRICE_FUELS))},$D{r})', f'$E{r}')
        return '=' + (call('NEWEXCISE', *a) if use_lambda else '/'.join(a))
    if var == 'nce':
        a = (o('ctxnew'), o('ets'), o('ntx'))
        return '=' + (call('NEWPOLICY', *a) if use_lambda else '+'.join(a))
    if var == 'atp':
        rpb = f'INDEX({col}${R_PV0 + PVOFF["rpb"]}:{col}${R_PV0 + PVOFF["rpb"] + NP - 1},$D{r})'
        return (f'={call("POSTTAX", rpb, o("nce"), f"$E{r}")}' if use_lambda
                else f'=MAX({rpb}+{o("nce")}*(1+$E{r}),0.01)')
    if var == 'shp':
        a = (f'INDEX({pol_range(col, "shps." + FEEBATE_SECTORS[0][0], len(FEEBATE_SECTORS))},$D{r})',
             f'INDEX({pol_range(col, "ssc.pow", len(COV_SECTORS))},$F{r})', f'{col}${R_POL["ets.pe"]}',
             f'{etsc}$F{r})', f'{obr}$F{r})', f'$E{r}')
        return '=' + (call('SHADOWEFF', *a) if use_lambda else f'({a[0]}*{a[1]}+{a[2]}*{a[3]}*{a[4]})*{a[5]}')
    if var == 'co2':
        a = (o('ener'), 'Settings!$C$13', f'$D{r}')
        return '=' + (call('EMISSIONS', *a) if use_lambda else '*'.join(a))
    if var in REV_VARS:
        pr = lambda v: f'INDEX({col}${R_PV0 + PVOFF[v]}:{col}${R_PV0 + PVOFF[v] + NP - 1},$D{r})'
        k = 'Settings!$C$13'
        if var == 'rtx':
            rate = f'{pr("etx")}+{pr("vat")}'
        elif var == 'rsub':
            rate = pr('esub')
        else:
            a = (o('ctxnew'), o('ntx'), o('ets'), f'{col}${R_POL["ets.rf"]}', o('nce'), f'$D{r}')
            rate = call('NEWREVRATE', *a) if use_lambda else f'{a[0]}+{a[1]}+{a[2]}*{a[3]}+{a[4]}*{a[5]}'
        if use_lambda:
            return '=' + call('REVENUE', o('ener'), k, rate)
        return f'={o("ener")}*{k}*({rate})'
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


LH = 'Settings!$C$10'                   # last historical price year


def gp_range(col):
    return f'{col}${R_GP0}:{col}${R_GP0 + len(GP_ROWS) - 1}'


def gp_formula(r, col):
    """International price, real: Prices_int (nominal, source from MTInputs) x deflator index x High/Low factor
    after the last historical year. Row without commodity key (D empty) = 1 (flat)."""
    src, adj = f'{col}${R_PRI[PRI_PARAMS[0]]}', f'{col}${R_PRI[PRI_PARAMS[1]]}'
    return (f'=IF($D{r}="",1,INDEX(Prices_int!$F$4:$AD$100,MATCH(SUBSTITUTE({src},"*","")&"|"&$D{r},'
            f'Prices_int!$A$4:$A$100,0),MATCH({col}${R_YEAR},Prices_int!$F$3:$AD$3,0))*{col}${R_DEFL}'
            f'*IF({col}${R_YEAR}>{LH},IF(LEFT({adj},4)="High",$E{r},IF(LEFT({adj},3)="Low",$F{r},1)),1))')


def price_formula(var, r, col, prev, block):
    """Section-2 cell of price variable var in row r. block: 'hist' (years up to the last historical price year,
    base year included: data from Inputs_prices by price position), 'proj' (projected years) or 'lam' (right
    column: the named LAMBDA, which holds the IF between history and projection)."""
    o = lambda v: f'{col}{r - PVOFF[var] + PVOFF[v]}'
    hist = lambda v, pos: (f'INDEX(Inputs_prices!${IPC[v + str(HIST_YEARS[0])]}${IP_R0}:'
                           f'${IPC[v + str(HIST_YEARS[-1])]}${IP_R1},{pos},MATCH({col}${R_YEAR},'
                           f'Inputs_prices!${IPC[v + str(HIST_YEARS[0])]}$4:${IPC[v + str(HIST_YEARS[-1])]}$4,0))')
    lam = block == 'lam'
    if var == 'sp':
        gp_now, gp_prev = f'INDEX({gp_range(col)},$E{r})', f'INDEX({gp_range(prev)},$E{r})'
        if block == 'hist':
            return '=' + hist('sp', f'$F{r}')
        if lam:
            return '=' + call('SUPPLYCOST', f'{col}${R_YEAR}', LH, f'IFERROR({hist("sp", f"$F{r}")},0)', f'$D{r}',
                              f'{prev}{r}', gp_now, gp_prev)
        return f'=$D{r}+({prev}{r}-$D{r})*{gp_now}/{gp_prev}'
    if var == 'txo':
        if block == 'hist':
            return '=' + hist('txo', f'$G{r}')
        if lam:
            return '=' + call('OTHERTAX', f'{col}${R_YEAR}', LH, f'IFERROR({hist("txo", f"$G{r}")},0)', f'$D{r}',
                              f'$E{r}', o('sp'), f'$F{r}')
        cs = f'($E{r}-{o("sp")}+$D{r}*(1-$F{r}))*(1-$F{r})'
        return f'=$D{r}*$F{r}+IF($D{r}*(1-$F{r})>=0,MAX($D{r}*(1-$F{r}),{cs}),{cs})'
    if var == 'rpb':
        return '=' + (call('RETAILPRICE', o('sp'), o('txo'), f'$D{r}') if lam
                      else f'({o("sp")}+{o("txo")})*(1+$D{r})')
    if var == 'vat':
        return '=' + (call('VATPAY', o('sp'), o('txo'), f'$D{r}') if lam else f'({o("sp")}+{o("txo")})*$D{r}')
    if var == 'etx':
        return '=' + (call('TAXPART', o('txo')) if lam else f'MAX({o("txo")},0)')
    if var == 'esub':
        return '=' + (call('SUBSIDYPART', o('txo')) if lam else f'MAX(-{o("txo")},0)')
    if var == 'esubpu':
        return '=' + (call('PERUNIT', o('esub'), f'$D{r}') if lam else f'{o("esub")}*$D{r}')
    raise ValueError(var)


PP_COLS = (L(5), L(4 + NY))             # PowerPaths: year columns (E = base year)


def pw_block(var, col):
    """Absolute-row range of a power variable over the generation types (or user groups) in column col."""
    mem = PW_TECH if PW_KIND[var] == 'tech' else PW_GROUPS
    r0 = R_PW[(var, mem[0][0])]
    return f'{col}${r0}:{col}${r0 + len(mem) - 1}'


def pw_param_formula(spec, var, r):
    """Mitigation D:G cell of a power row: data key, Inputs_power lookup, PowerParams value or constant."""
    if spec[0] == 'key':
        kind = PW_KIND[var]
        return f'=$A{r}&"|"&$B{r}' if kind == 'tech' else (f'=$A{r}&"|"&$C{r}' if kind == 'user' else f'=$A{r}')
    if spec[0] == 't':
        c = IPWC[spec[1]]
        return f'=INDEX(Inputs_power!${c}${IPW_R0}:${c}${IPW_R1},MATCH($B{r},Inputs_power!$A${IPW_R0}:$A${IPW_R1},0))'
    if spec[0] == 'g':
        c = IPWG[spec[1]]
        return f'=INDEX(Inputs_power!${c}${IPW_G0}:${c}${IPW_G1},MATCH($C{r},Inputs_power!$A${IPW_G0}:$A${IPW_G1},0))'
    if spec[0] == 'p':
        return '=' + pw_param(spec[1])
    return spec[1]


def pw_formula(var, r, col, prev, block):
    """Section-3 cell of power variable var in row r. block: 'all' (data rows, every column), 'base', 'plain',
    'hist' / 'proj' (rppre) or 'lam' (right column: the named LAMBDA)."""
    lam = block == 'lam'
    first = {'tech': PW_TECH[0][0], 'user': PW_GROUPS[0][0]}.get(PW_KIND[var])
    o = lambda v, c=col: f'{c}{r - R_PW[(var, first)] + R_PW[(v, first)]}'     # same member, other variable
    one = lambda v, c=col: f'{c}${R_PW[(v, "")]}'                                 # single row
    if var in PW_DATA:
        return (f'=INDEX(PowerPaths!${PP_COLS[0]}$4:${PP_COLS[1]}$200,MATCH($D{r},PowerPaths!$A$4:$A$200,0),'
                f'MATCH({col}${R_YEAR},PowerPaths!${PP_COLS[0]}$3:${PP_COLS[1]}$3,0))')
    if var == 'cax':
        return '=' + (call('CAPEX', f'$D{r}', o('tcf')) if lam else f'$D{r}*{o("tcf")}')
    if var in ('caxav', 'stoav'):
        new = (lambda c: o('cax', c)) if var == 'caxav' else (lambda c: f'$D{r}*{one("msc", c)}')
        if block == 'base':
            return '=' + (o('cax') if var == 'caxav' else new(col))
        new_v = new(prev) if var == 'caxav' else new(col)          # capex of last year; storage cost of this year
        a = (f'{prev}{r}', new_v, o('phi', prev))
        return '=' + (call('VINTAGE', *a) if lam else f'{a[0]}*(1-{a[2]})+{a[1]}*{a[2]}')
    if var == 'fix':
        a = (o('caxav'), f'$D{r}', f'$E{r}', o('stoav'))
        return '=' + (call('AMORTISED', *a) if lam else f'{a[0]}*{a[1]}+{a[2]}+{a[3]}')
    if var == 'lfx':
        a = (o('cax'), f'$D{r}', f'$E{r}', f'$F{r}', one('msc'))
        return '=' + (call('LEVELISED', *a) if lam else f'{a[0]}*{a[1]}+{a[2]}+{a[3]}*{a[4]}')
    if var == 'vbc':
        rpb = f'INDEX({col}${R_PV0 + PVOFF["rpb"]}:{col}${R_PV0 + PVOFF["rpb"] + NP - 1},$E{r})'
        return '=' + (call('VARCOST', f'$D{r}', rpb, f'$F{r}') if lam else f'$D{r}+{rpb}*$F{r}')
    if var == 'ccp':
        cp, fc = f'{col}${R_CP}', f'INDEX({pol_range(col, "fc.coa", NF)},$E{r})'
        sc, etsc, obr = (f'{col}${R_POL["sc.pow"]}', f'{col}${R_POL["etsc.pow"]}', f'{col}${R_POL["obr.pow"]}')
        pe = f'{col}${R_POL["ets.pe"]}'
        fpr = f'INDEX({pol_range(col, "fpr." + PRICE_FUELS[0][0], len(PRICE_FUELS))},$F{r})*$G{r}'
        if lam:
            return '=' + call('NEWPOLICY', call('CARBONTAX', cp, f'$D{r}', fc, sc, etsc),
                              call('ETSCOST', pe, f'$D{r}', etsc, obr), fpr)
        return f'={cp}*$D{r}*{fc}*{sc}*(1-{etsc})+{pe}*$D{r}*{etsc}*(1-{obr})+{fpr}'
    if var == 'gnc':
        return '=' + (call('GENCOST', o('fix'), o('vbc')) if lam else f'{o("fix")}+{o("vbc")}')
    if var in ('gncav', 'ccpav', 'vre'):
        vals = {'gncav': pw_block('gnc', col), 'ccpav': pw_block('ccp', col),
                'vre': f'$E${R_PW[("gns", PW_TECH[0][0])]}:$E${R_PW[("gns", PW_TECH[0][0])] + NT - 1}'}[var]
        a = (pw_block('gns', col), vals, f'$D{r}')
        return '=' + (call('WEIGHTED', *a) if lam else f'SUMPRODUCT({a[0]},{a[1]})*{a[2]}')
    if var == 'cph':
        a = (one('cbat'), one('cint'), one('obat'), f'$D{r}', f'$E{r}')
        return '=' + (call('STORAGEHOUR', *a) if lam else f'({a[0]}+{a[1]}/{a[3]})/({a[4]}*8760)+{a[2]}/8760')
    if var == 'msc':
        a = (f'$D{r}', f'$E{r}', one('vre'), one('cph'), f'$F{r}', one('cel'))
        return '=' + (call('STORAGECOST', *a) if lam else
                      f'{a[0]}*{a[1]}*{a[2]}*{a[3]}+2*MAX(0,{a[2]}-{a[4]})/(1-{a[4]})^2*{a[5]}')
    if var == 'sc':
        gen = one('gncav') if block == 'base' else one('gncav', prev)
        return '=' + (call('POWERSUPPLY', gen, f'$D{r}') if lam else f'{gen}+$D{r}')
    if var == 'rppre':
        hist = (f'INDEX(Inputs_power!${IPWG["rp" + str(HIST_YEARS[0])]}${IPW_G0}:${IPWG["rp" + str(HIST_YEARS[-1])]}'
                f'${IPW_G1},$E{r},MATCH({col}${R_YEAR},Inputs_power!${IPWG["rp" + str(HIST_YEARS[0])]}${IPW_G0 - 1}:'
                f'${IPWG["rp" + str(HIST_YEARS[-1])]}${IPW_G0 - 1},0))')
        if block == 'hist':
            return '=' + hist
        if lam:
            return '=' + call('POWERPRICE', f'{col}${R_YEAR}', LH, f'IFERROR({hist},0)', f'{prev}{r}', f'$D{r}',
                              o('sc'), o('sc', prev), f'$F{r}')
        return f'={prev}{r}+$D{r}*({o("sc")}-{o("sc", prev)})*(1+$F{r})'
    if var == 'rp':
        return '=' + (call('ENDUSERPOWER', o('rppre'), one('ccpav')) if lam else f'{o("rppre")}+{one("ccpav")}')
    if var == 'sgap':
        return '=' + (call('SUBSIDYGAP', o('sc'), o('rppre'), f'$D{r}') if lam else f'{o("sc")}-{o("rppre")}/(1+$D{r})')
    if var in ('rvat', 'rcarb', 'rgap'):
        rate = {'rvat': o('rppre'), 'rcarb': one('ccpav'), 'rgap': o('sgap')}[var]
        return '=' + (call('POWERREV', rate, f'$D{r}', o('cons')) if lam else f'{rate}*$D{r}*{o("cons")}')
    raise ValueError(var)


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


def pol_calc(key, col, use_lambda=False):
    c = lambda k: f'{col}${R_POL[k]}'          # fixed rows (years, carbon-tax inputs): absolute
    rel = lambda k: f'{col}{R_POL[k]}'          # start / final values: relative, so a path drags down
    y = f'{col}${R_YEAR}'
    if key == 'cptraj':                      # nominal inputs (NomorReal = Nominal) x CPI index (legacy row 2249)
        nom = f'IF(LEFT({c("NomorReal")},7)="Nominal",{col}${R_CPI},1)'
        if use_lambda:
            return '=' + call('PATH', y, c('CPIntro'), c('CPOutro'), c('CPLevelStart'), c('CPLevelTarget'),
                              f'{c("ExtendCarbonPriceBeyondOutro")}="Linear*"') + '*' + nom
        path = path_formula(col, c('CPIntro'), c('CPOutro'), c('CPLevelStart'), c('CPLevelTarget'),
                            c('ExtendCarbonPriceBeyondOutro'))
        return f'=({path[1:]})*{nom}'
    if key.startswith('fpr.') or (key.startswith('fb.') and key not in ('fb.yr0', 'fb.yr1')):
        pre, k = ('fpr', key[4:]) if key.startswith('fpr.') else ('fb', key[3:])
        y0, y1 = c(f'{pre}.yr0'), c(f'{pre}.yr1')
        s0, s1 = (rel('fprs.' + k), rel('fprf.' + k)) if pre == 'fpr' else (rel('fbs.' + k), rel('fbt.' + k))
        if use_lambda:
            return '=' + call('PATH', y, y0, y1, s0, s1, 'TRUE')
        return path_formula(col, y0, y1, s0, s1)
    on = f'AND(LEFT({c("D_NewETS")},3)="Yes",{y}>={c("D_ETSIntro")})'
    bco2 = pol_range(col, 'bco2.pow', len(COV_SECTORS))
    etsc = pol_range(col, 'etsc.pow', len(COV_SECTORS))
    if key == 'ets.p':
        a = (c('D_NewETS'), y, c('D_ETSIntro'), c('D_ETSPriceOverride'), c('ets.ovr'), c('ets.est'), c('cptraj'))
        return '=' + (call('ETSPRICE', *a) if use_lambda else
                      f'IF({on},IF(LEFT({a[3]},3)="Yes",IF({a[4]}="",{a[6]},{a[4]}),{a[5]}),0)')
    if key.startswith('etsb.'):              # benchmark: linear from start to target year, flat after
        y0, y1, s0, s1 = c('D_ETSIntro'), c('D_ETSOutro'), rel('etsbs.' + key[5:]), rel('etsbt.' + key[5:])
        if use_lambda:
            return '=' + call('PATH', y, y0, y1, s0, s1, 'FALSE')
        return f'=IF({y}<{y0},0,{s0}+({s1}-{s0})/MAX({y1}-{y0},1)*(MIN({y},{y1})-{y0}))'
    if key.startswith('obr.'):
        b = f'INDEX({pol_range(col, "etsb.pow", len(ETS_GROUPS))},$D{R_POL[key]})'
        return '=' + (call('OBRSHARE', b) if use_lambda else f'MIN(1,MAX(0,{b}))')
    if key == 'ets.vadj':
        a = (c('D_ETSVolatility'), c('D_ETSVolImpact'), c('D_ETSCTRisk'))
        if use_lambda:
            return '=' + call('VOLADJ', *a)
        imp = f'INDEX(Settings!$C$18:$F$18,MATCH(SUBSTITUTE({a[1]},"*",""),Settings!$C$16:$F$16,0))'
        vol = f'INDEX(Settings!$C$17:$F$17,MATCH(SUBSTITUTE({a[0]},"*",""),Settings!$C$16:$F$16,0))'
        return f'=(1+{a[2]}*{imp})/(1+{vol}*{imp})'
    if key == 'ets.bce':
        return '=' + (call('COVERED', bco2, etsc) if use_lambda else f'SUMPRODUCT({bco2},{etsc})')
    if key == 'ets.cap':
        a = (c('D_NewETS'), y, c('D_ETSIntro'), c('D_ETSOutro'), c('D_ETSChangeRelStart'), c('D_ETSChangeRelTarget'),
             c('D_ETSCapCont'), c('ets.bce'), f'{L(CI_(col) - 1)}${R_POL["ets.cap"]}')
        if use_lambda:
            return '=' + call('ETSCAP', *a)
        return (f'=IF({on},IF(AND({y}>{a[3]},LEFT({a[6]},8)="Constant"),{a[8]},{a[7]}*(1+{a[4]}+({a[5]}-{a[4]})'
                f'/MAX({a[3]}-{a[2]},1)*(MIN({y},{a[3]})-{a[2]}))),0)')
    if key == 'ets.se':
        semi, obr = f'$E${R_POL["obr.pow"]}:$E${R_POL["obr.pow"] + len(COV_SECTORS) - 1}', \
            pol_range(col, 'obr.pow', len(COV_SECTORS))
        if use_lambda:
            return '=' + call('ETSSEMI', bco2, etsc, semi, obr, 'Settings!$C$14')
        return (f'=IF(SUMPRODUCT({bco2},{etsc})>0,SUMPRODUCT({bco2},{etsc},{semi},1-(1-Settings!$C$14)*{obr})'
                f'/SUMPRODUCT({bco2},{etsc}),0)')
    if key == 'ets.est':
        a = (c('D_NewETS'), y, c('D_ETSIntro'), c('ets.cap'), c('ets.bce'), c('ets.se'), c('ets.vadj'))
        if use_lambda:
            return '=' + call('ETSESTIMATE', *a)
        return f'=IF(AND({on},{a[3]}>0,{a[3]}<{a[4]},{a[5]}<0,{a[6]}>0),LN({a[3]}/{a[4]})/({a[5]}*{a[6]}),0)'
    if key == 'ets.pe':
        return '=' + (call('TAXEQUIV', c('ets.p'), c('ets.vadj')) if use_lambda else f'{c("ets.p")}*{c("ets.vadj")}')
    if key == 'ets.rf':
        v = c('ets.vadj')
        return '=' + (call('REVFACTOR', v) if use_lambda else f'IF({v}>0,1/{v},0)')
    if key.startswith('etsc.'):
        a = (c('D_NewETS'), y, c('D_ETSIntro'), rel('etscv.' + key[5:]))
        return '=' + (call('ETSCOVER', *a) if use_lambda
                      else f'IF(AND(LEFT({a[0]},3)="Yes",{y}>={a[2]}),IF({a[3]},1,0),0)')
    if key.startswith('shps.'):
        a = (rel('fb.' + key[5:]), rel('reg.' + key[5:]))
        return '=' + (call('SHADOWSECTOR', *a) if use_lambda else '+'.join(a))
    if key.startswith('ssc.'):
        a = (rel('fbc.' + key[4:]), c('ssc.adj'))
        return '=' + (call('SHADOWSHARE', *a) if use_lambda else f'IF({a[0]},1,0)*{a[1]}')
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
    title(ws, f'{NAME} v{VERSION} - mitigation module: policies, prices, fuel use, revenues, CO2 (copy-pasteable)', last)
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
    for r_, v_ in ((R_CPI, 'infl'), (R_DEFL, 'defl')):
        ids(r_, v_, '', '')
        codes(r_)
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
        for r_, key in ((R_CPI, 'USA|pcpi'), (R_DEFL, 'USA|ngdp_d')):
            m = f'MATCH("{key}",WEO!$A$4:$A$1000,0)'
            put(ws, f'{col}{r_}', (f'=INDEX(WEO!$F$4:$AZ$1000,{m},MATCH(Settings!$C$8,WEO!$F$3:$AZ$3,0))'
                                   f'/INDEX(WEO!$F$4:$AZ$1000,{m},MATCH({col}${R_YEAR},WEO!$F$3:$AZ$3,0))'),
                fill_=bf, fmt='0.0000')
    each_col(top)

    sub_font = Font(name='Arial', size=9, bold=True, italic=True)

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
        for pc_, v_ in zip(PARAM_COLS, POL_PARAMS.get(key, ())):
            put(ws, f'{pc_}{r}', v_, FONT, F_INPUT, '0.0000000' if pc_ == 'E' else '0')
        codes(r, fnt)
        outline[r] = (0, False) if summary else (1, True)

        def pcell(c, is_base, is_last, r=r, key=key, kind=kind, fnt=fnt):
            col = L(c)
            bf = F_BASE if is_base else None
            if kind == 'mt':
                val = mt_lookup(r, col)
            elif kind == 'calc':
                val = pol_calc(key, col, is_last)
            elif kind == 'data':
                val = DATA_VALUES.get((key, year_of(c)))
            else:
                val = CONST_VALUES[key]
            fmt = '0' if var in ('CPIntro', 'CPOutro', 'fpr.yr0', 'fpr.yr1', 'fb.yr0', 'fb.yr1') else '0.00'
            if var in ('ctcov', 'fbcov') and kind == 'const':
                fmt = 'General'
            lam = kind == 'calc' and is_last
            if var in ('bco2', 'ets.bce', 'ets.cap'):
                fmt = '#,##0.000'
            elif var == 'ets.se':
                fmt = '0.0000000'
            put(ws, f'{col}{r}', val, FONT_LAMBDA if lam else fnt,
                F_LAMBDA if lam else (F_INPUT if kind in ('const', 'data') else bf), fmt)
        each_col(pcell)

    # 2. Retail energy prices (before new policies)
    band(ws, B_PRI, '2. Retail energy prices before new policies (real $/GJ of ResultsYear; by price fuel; '
                    'legacy method, PriceProjection_Method_v0.3.md)', last)
    put(ws, f'H{R_PRI["head.int"]}', 'International energy prices (MTInputs rows 220, 258; real via the deflator '
                                     'index; only ratios are used)', sub_font)
    put(ws, f'H{R_PRI["head.dom"]}', 'Domestic prices by price fuel: supply cost sp, excise and other taxes txo, '
                                     'retail price before new policies rpb', sub_font)
    outline[R_PRI['head.int']] = (1, True)
    outline[R_PRI['head.dom']] = (1, True)
    for k, p in enumerate(PRI_PARAMS):
        r = R_PRI[p]
        ids(r, p, '', '')
        put(ws, f'D{r}', MT_NAMES[p], FONT, F_INPUT)
        codes(r)
        outline[r] = (1, True)

        def scell(c, is_base, is_last, r=r):
            put(ws, f'{L(c)}{r}', mt_lookup(r, L(c)), fill_=F_BASE if is_base else None)
        each_col(scell)
    for k, (fu, lab, key, hi, lo) in enumerate(GP_ROWS):
        r = R_GP0 + k
        ids(r, 'gp', fu, 'int')
        codes(r)
        outline[r] = (1, True)
        for pc, v in zip(PARAM_COLS, [key, hi, lo, None]):
            put(ws, f'{pc}{r}', v, FONT, F_UNUSED if v is None else F_INPUT)

        def gcell(c, is_base, is_last, r=r):
            put(ws, f'{L(c)}{r}', gp_formula(r, L(c)), fill_=F_BASE if is_base else None, fmt='0.00')
        each_col(gcell)
    for var in PVARS:
        params = PVAR_PARAMS[var]
        for i, (pc, lab) in enumerate(PRICES):
            r = R_PV0 + PVOFF[var] + i
            fuel, sec = pc.split('.')
            ids(r, var, fuel, sec)
            codes(r)
            outline[r] = (0, False) if var == 'rpb' else (1, True)
            for j, pcol in enumerate(PARAM_COLS):
                if j < len(params):
                    put(ws, f'{pcol}{r}', ip_lookup(params[j], f'$B{r}&"."&$C{r}'), fmt='0.000')
                else:
                    put(ws, f'{pcol}{r}', None, fill_=F_UNUSED)

            def pcell2(c, is_base, is_last, r=r, var=var):
                lam = is_last
                block = 'lam' if lam else ('hist' if year_of(c) <= HIST_YEARS[-1] else 'proj')
                fl = F_LAMBDA if lam else (F_BASE if is_base else (F_LIGHT if block == 'hist' and var in ('sp', 'txo')
                                                                    else None))
                put(ws, f'{L(c)}{r}', price_formula(var, r, L(c), L(c - 1), block),
                    FONT_LAMBDA if lam else FONT, fl, '0.000')
            each_col(pcell2)
        outline[R_PV0 + PVOFF[var] + NP] = (1, True)            # blank row after each variable

    # 3. Power sector: generation costs and power prices (v1.05, PowerPrices_Method_v0_1.md)
    band(ws, B_POW, '3. Power sector: generation costs and electricity prices (legacy section 4 sub-tables A-G, J; '
                    'quantities from the engineer model later)', last)
    for let, head, _items in PW_LAYOUT:
        put(ws, f'H{PW_HEAD[let]}', head, sub_font)
        outline[PW_HEAD[let]] = (1, True)
    f_interim = Font(name='Arial', size=9, italic=True, color='C65911')
    for r, var, fu, sec, kind in PW_ROWS:
        summary = var in PW_SUMMARY
        fnt = FONT_SUM if summary else (f_interim if var in PW_INTERIM else FONT)
        ids(r, var, fu, sec, fnt)
        codes(r, fnt)
        outline[r] = (0, False) if summary else (1, True)
        specs = PW_PARAMS[var]
        for j, pcol in enumerate(PARAM_COLS):
            if j < len(specs):
                put(ws, f'{pcol}{r}', pw_param_formula(specs[j], var, r), FONT,
                    F_INPUT if specs[j][0] == 'c' else None, 'General' if specs[j][0] == 'key' else '0.000000')
            else:
                put(ws, f'{pcol}{r}', None, fill_=F_UNUSED)
        fmt = PW_FMT.get(var, '0.0000')

        def pwcell(c, is_base, is_last, r=r, var=var, fnt=fnt, fmt=fmt):
            data = var in PW_DATA
            lam = is_last and not data
            if data:
                block = 'all'
            elif lam:
                block = 'lam'
            elif var == 'rppre':
                block = 'hist' if year_of(c) <= HIST_YEARS[-1] else 'proj'
            else:
                block = 'base' if is_base and var in PW_BASE else 'plain'
            fl = F_LAMBDA if lam else (F_BASE if is_base else (F_LIGHT if block == 'hist' else None))
            put(ws, f'{L(c)}{r}', pw_formula(var, r, L(c), L(c - 1), block), FONT_LAMBDA if lam else fnt, fl, fmt)
        each_col(pwcell)
    for r in range(B_POW + 1, R_POW_NOTE + 2):
        outline.setdefault(r, (1, True))                    # blank rows between blocks and the notes roll up too
    notes = [
        'Interim data (orange italic): generation shares gns, investment shares phi and consumption cons are the legacy '
        'baseline run until the engineer model and electricity demand are built, so a scenario changes power costs '
        'and prices through fuel and carbon costs but not yet through its own generation mix.',
        'Interactions: section 2 fuel prices -> D; section 1 policies -> D (ccp) -> E (ccpav) -> G (rp); generation '
        'mix -> E and J (VRE share -> storage cost -> B and C); E -> G with a one-year lag (sc uses gncav of the '
        'previous year). C (levelised) feeds investment only. Power results are not yet in the totals of sections '
        '11-13.']
    for k, n in enumerate(notes):
        put(ws, f'H{R_POW_NOTE + k}', n, Font(name='Arial', size=9, italic=True))

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
        for s in subs:
            h = SUB_HEAD[s]
            ids(h, 'ener', 'all', s, FONT_SUM)
            codes(h, FONT_SUM)
            outline[h] = (1, False)
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
                    fmt = '#,##0.0' if var in ['ener'] + REV_VARS else ('0.0000' if var == 'co2' else '0.000')

                    def cell(c, is_base, is_last, r=r, var=var, fmt=fmt):
                        lam = is_last
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

    # 12. Revenues: existing taxes, existing subsidies and new policies as separate calculations
    band(ws, B_REV, '12. Revenues (USD million, real of ResultsYear): existing taxes, existing subsidies and new '
                    'policies, calculated separately', last)
    spans = {g: (SUB_HEAD[[x[0] for x in SUBSECTORS if x[2] == g][0]],
                 SUB_HEAD[[x[0] for x in SUBSECTORS if x[2] == g][-1]] + SB - 1) for _n, g, _t in SECTIONS}
    far = L(last + 10 * GW)
    for key, var, fu, sec in REV_ROWS:
        r = R_REV[key]
        summary = sec == 'all'
        fnt = FONT_SUM if summary else FONT
        ids(r, var, fu, sec, fnt)
        codes(r, fnt)
        outline[r] = (0, False) if summary else (1, True)

        def rcell(c, is_base, is_last, r=r, key=key, var=var, sec=sec, fnt=fnt):
            col = L(c)
            if var in REV_VARS:
                a, z = (S_FIRST, S_LAST) if sec == 'all' else spans[sec]
                f = f'=SUMIFS({col}${a}:{col}${z},$A${a}:$A${z},"{var}",$B${a}:$B${z},"<>all")'
            elif var == 'rnet':
                f = f'={col}{R_REV["rtx.all"]}-{col}{R_REV["rsub.all"]}'
            elif var == 'rtot':
                f = f'={col}{R_REV["rnet.all"]}+{col}{R_REV["rnew.all"]}'
            elif var == 'rtot.ref':
                rr = R_REV['rtot.all']
                f = (f'=INDEX(${L(COL_G0)}{rr}:${far}{rr},'
                     f'MATCH({col}${R_YEAR},${L(COL_G0)}${R_YEAR}:${far}${R_YEAR},0))')
            else:
                f = f'={col}{R_REV["rtot.all"]}-{col}{R_REV["rtot.ref"]}'
            put(ws, f'{col}{r}', f, fnt, F_BASE if is_base else None, '#,##0.0')
        each_col(rcell)
    notes = [
        'Per subsector and fuel (rows rtx, rsub, rnew in each subsector block): rtx = fuel use x PJ/ktoe x (existing '
        'tax etx + VAT vat of the price fuel); rsub = fuel use x PJ/ktoe x existing subsidy esub (fiscal cost, '
        'positive); rnew = fuel use x PJ/ktoe x (new carbon tax + new excise + ETS x revenue factor + VAT on new '
        'policies nce x VAT rate). Feebates are revenue-neutral.',
        'Existing taxes and subsidies use the scenario\'s own fuel use, so a policy also changes them (the fiscal '
        'effect vs scenario 1 is rtot.chg). Power and electricity are not yet included.',
        'Per-fuel policy input (later): esubpu (existing subsidy per MTInputs price unit, section 2) can feed a '
        'subsidy-reform path in the fuel price reform rows (MTInputs 142-169 use the same units).']
    for k, n in enumerate(notes):
        put(ws, f'H{R_REV_NOTE + k}', n, Font(name='Arial', size=9, italic=True))

    # 13. Emissions: CO2 from fuel combustion
    band(ws, B_EMI, '13. Emissions - CO2 from fuel combustion (MtCO2; power, process emissions and other gases '
                    'not yet)', last)
    for key, var, fu, sec in EMI_ROWS:
        r = R_EMI[key]
        summary = key in ('co2.all', 'co2.chg', 'co2.pct')
        fnt = FONT_SUM if summary else FONT
        ids(r, var, fu, sec, fnt)
        codes(r, fnt)
        outline[r] = (0, False) if summary else (1, True)

        def ecell(c, is_base, is_last, r=r, key=key, var=var, fu=fu, sec=sec, fnt=fnt):
            col = L(c)
            if var == 'co2' and fu == 'all':
                a, z = (S_FIRST, S_LAST) if sec == 'all' else spans[sec]
                f = f'=SUMIFS({col}${a}:{col}${z},$A${a}:$A${z},"co2",$B${a}:$B${z},"<>all")'
                fmt = '#,##0.00'
            elif var == 'co2':
                a, z = S_FIRST, S_LAST
                f = f'=SUMIFS({col}${a}:{col}${z},$A${a}:$A${z},"co2",$B${a}:$B${z},$B{r})'
                fmt = '#,##0.00'
            elif var == 'co2.ref':
                rr = R_EMI['co2.all']
                f = (f'=INDEX(${L(COL_G0)}{rr}:${far}{rr},'
                     f'MATCH({col}${R_YEAR},${L(COL_G0)}${R_YEAR}:${far}${R_YEAR},0))')
                fmt = '#,##0.00'
            elif var == 'co2.chg':
                f = f'={col}{R_EMI["co2.all"]}-{col}{R_EMI["co2.ref"]}'
                fmt = '#,##0.00'
            elif var == 'co2.pct':
                f = f'={col}{R_EMI["co2.all"]}/{col}{R_EMI["co2.ref"]}-1'
                fmt = '0.0%'
            elif var == 'co2.sec':
                a, z = S_FIRST, S_LAST
                f = f'=SUMIFS({col}${a}:{col}${z},$A${a}:$A${z},"co2",$C${a}:$C${z},$C{r},$B${a}:$B${z},"<>all")'
                fmt = '#,##0.000'
            elif var == 'co2.ets':
                r0 = R_EMI['co2.sec.pow']
                f = (f'=SUMPRODUCT({col}${r0}:{col}${r0 + len(COV_SECTORS) - 1},'
                     f'{pol_range(col, "etsc.pow", len(COV_SECTORS))})')
                fmt = '#,##0.000'
            elif var == 'co2.cap':
                f = f'={col}${R_POL["ets.cap"]}'
                fmt = '#,##0.000'
            elif var == 'co2.gap':
                f = f'=IF({col}{R_EMI["co2.cap"]}>0,{col}{R_EMI["co2.ets"]}/{col}{R_EMI["co2.cap"]}-1,0)'
                fmt = '0.0%'
            elif var == 'ets.next':
                p_, cap, bce, cov = (f'{col}${R_POL["ets.p"]}', f'{col}${R_POL["ets.cap"]}',
                                     f'{col}${R_POL["ets.bce"]}', f'{col}{R_EMI["co2.ets"]}')
                f = (f'=IF(OR({p_}<=0,{cap}<=0,{cap}>={bce},{cov}>={bce}),{col}${R_POL["ets.est"]},'
                     f'{p_}*(LN({cap}/{bce})/LN({cov}/{bce}))^Settings!$C$15)')
                fmt = '0.00'
            else:                                                       # bco2.chk
                f = f'=SUM({pol_range(col, "bco2.pow", len(COV_SECTORS))})-{col}{R_EMI["co2.ref"]}'
                fmt = '0.000'
            put(ws, f'{col}{r}', f, fnt, F_CHECK if var == 'bco2.chk' else (F_BASE if is_base else None), fmt)
        each_col(ecell)
    put(ws, f'H{R_EMI_NOTE}', 'Per subsector and fuel: co2 = fuel use (ktoe) x PJ/ktoe x EF (tCO2/GJ, IIASA, EF_GHG '
                              'sheet; biomass 0). No inventory adjustment, no power sector, no process emissions; '
                              'CH4, N2O and local pollutants later. ETS rows: co2.sec by ETS sector (power 0 until '
                              'modelled), co2.ets covered emissions, co2.cap the cap, co2.gap the distance to the cap, '
                              'ets.next the next price (paste its values into ets.ovr with D_ETSPriceOverride = Yes '
                              'and repeat, or run ets_goalseek_v0_1.py); bco2.chk = 0 when the baseline data in '
                              'section 1 match scenario 1.', Font(name='Arial', size=9, italic=True))

    # Outline (v0.15): summary / group button below each row group; everything opens rolled up.
    ws.sheet_properties.outlinePr.summaryBelow = True
    ws.sheet_format.outlineLevelRow = 2
    level = {r: lvl for r, (lvl, _hid) in outline.items() if lvl}
    for r, lvl in level.items():
        ws.row_dimensions[r].outlineLevel = lvl
        ws.row_dimensions[r].hidden = True
    for r in sorted(level):                  # the row after a group's last row carries its collapsed flag
        if level.get(r + 1, 0) < level[r]:
            ws.row_dimensions[r + 1].collapsed = True
    ws.freeze_panes = f'{L(COL_G0 + 1)}{R_SCEN + 1}'
    for col, w in {'A': 11, 'B': 5, 'C': 5, 'H': 44, 'I': 9, 'J': 18}.items():
        ws.column_dimensions[col].width = w
    for g in groups:
        code, base, years = group_cols(g)
        ws.column_dimensions[L(code)].width = 24
        for c in [base] + years:
            ws.column_dimensions[L(c)].width = 9
        ws.column_dimensions[L(years[-1] + 1)].width = 3          # blank spacer column between scenario groups
    ycols = [c for g in groups for c in group_cols(g)[2] if COL_YEARS_GROUP[0] <= year_of(c) <= COL_YEARS_GROUP[1]]
    for col in list(COL_OUTLINE) + [L(c) for c in ycols]:   # one level: B:G, I:K, 2030-2039 per scenario; rolled up
        ws.column_dimensions[col].outline_level = 1
        ws.column_dimensions[col].hidden = True
    after = ['H', L(COL_G0 + 1)] + [L(c + 1) for c in ycols if year_of(c) == COL_YEARS_GROUP[1]]
    for col in after:                        # columns right of each group carry the collapsed flag
        ws.column_dimensions[col].collapsed = True
    ws.sheet_format.outlineLevelCol = 1
    ws.sheet_view.zoomScale = 75
    return last


# ---------------------------------------------------------------- ReadMe
def build_readme(wb):
    ws = wb.active
    ws.title = 'ReadMe'
    title(ws, f'{NAME} v{VERSION} - AI-generated, copy-pasteable CPAT mitigation module', 3)
    code2 = group_cols(2)[0]
    lam_lines = '; '.join(f'{n}({", ".join(p)}) = {b}' for n, (p, b) in LAMBDAS.items())
    h = SUB_HEAD['rod']
    lines = [
        ('Purpose', 'Auditable, copy-pasteable replacement of the CPAT mitigation price -> fuel-use chain. Each '
                    'variable has one formula across all subsectors, fuels and years, and a whole scenario group '
                    'copies to a new scenario and keeps working.'),
        ('Status', f'v{VERSION} prototype to {LAST_YEAR}: policies (carbon tax, new ETS, fuel price reform, feebates '
                   'with sectoral shadow prices), domestic prices with existing taxes and subsidies, and price -> '
                   'fuel use for transport, buildings, industry and other energy use, revenues (section 12), CO2 '
                   '(section 13), and power generation costs and electricity prices (section 3; generation mix '
                   'interim data until the engineer model). Differences with legacy: sheet LegacyDiff.'),
        ('Layout (as legacy CPAT)', f'Sections: 1. Policies (row {B_POL}), 2. Retail energy prices (row {B_PRI}), '
                                    f'3. Power sector (row {B_POW}), 5. Transport (row '
                                    f'{SEC_BAND["tra"]}), 6. Buildings (row {SEC_BAND["bld"]}), 7. Industrial (row '
                                    f'{SEC_BAND["ind"]}), 8. Other energy use (row {SEC_BAND["oen"]}), 11. Results - '
                                    f'energy consumption (row {B_RES}). Numbers follow the legacy Mitigation sheet.'),
        ('Nesting', 'Across: scenario (outer) > year (inner). Down: sector section > subsector > variable > fuel. '
                    f'Each subsector: a heading line (its total fuel use), then {", ".join(VARS)}, each over the '
                    f'{NF} fuels with one blank row ({SB} rows per subsector). Section 2: sp, txo, rpb over the '
                    f'{NP} price fuels (coal and gas x power / residential / industry; gasoline, diesel, LPG, '
                    'kerosene, other oil products, biomass for all sectors).'),
        ('ASSUMPTIONS: units', 'All prices, taxes and policy values are in REAL US DOLLARS OF THE RESULTS YEAR '
                               '(MTInputs ResultsYear, 2026; Settings C8), per GJ. There is no inflation after the '
                               'conversion: a constant real price is a constant number. Fuel use is in ktoe.'),
        ('ASSUMPTIONS: two indices', f'Row {R_CPI} infl = US CPI(ResultsYear) / US CPI(year) converts nominal '
                                     'DOMESTIC data (IMF price dataset: supply cost, taxes, retail price, margin, '
                                     'production cost) and nominal policy inputs. Row '
                                     f'{R_DEFL} defl = US GDP deflator(ResultsYear) / deflator(year) converts '
                                     'nominal INTERNATIONAL prices (legacy uses the deflator there). The two differ '
                                     'by up to 1.2% before 2030 (2022: 1.137 vs 1.124), so both are kept. The CPI '
                                     'is held flat after 2031 in the legacy data (WEO horizon); results after 2030 '
                                     'are not compared with legacy.'),
        ('ASSUMPTIONS: inputs', 'Domestic price data are nominal USD of their year -> x infl of that year. '
                                'International prices are nominal -> x defl; only their year-on-year ratios move '
                                'the supply cost, so their units and level drop out. MTInputs policy values are '
                                'REAL (ResultsYear USD): carbon price, feebate rates and fuel price reform '
                                'increases. Only the carbon price has a switch: NomorReal = Nominal makes it '
                                f'nominal, then x infl (legacy row 2249). Margins and production costs are base-year '
                                '(2022) nominal values -> x infl(2022), held constant in real terms.'),
        ('ASSUMPTIONS: history', f'Years up to the last historical price year (Settings C10, 2024) take prices from '
                                 'data (Inputs_prices); later years are projected. Base-year (2022) fuel use is '
                                 'data; fuel use from 2023 on is modelled with these prices (also in 2023-2024).'),
        ('ASSUMPTIONS: global vs scenario', 'Global (Settings, from MTInputs scenario 1): ResultsYear, government '
                                            'price controls (pass-through), last historical year, gas market of the '
                                            'country. Per scenario (MTInputs column): international price source '
                                            'and adjustment, carbon price and its nominal/real switch, fuel price '
                                            'reform, feebates, coverage.'),
        ('ASSUMPTIONS: legacy quirk', 'Legacy rule kept: with pass-through 0.5 or 0.8 txo jumps in the first '
                                      'projected year. See PriceProjection_Method_v0.3.md.'),
        ('DATA CHANGED (v0.14)', 'Egypt price block replaced by the corrected block supplied by the user '
                                 '(data/source/Egypt_Price_Data_2026-10-09.xlsx): changed cells are bright yellow '
                                 'in Prices_dom (list: data/prices_dom_changes.csv). Its VAT rates are explicit (0 for '
                                 'gas and oil products), so the v0.10 VAT assumption no longer applies to Egypt.'),
        ('ASSUMPTIONS: DATA CHANGED (v0.10)', 'Bright yellow cells with red text in Inputs_prices: (1) other oil '
                                              'products take pass-through and margin from the dataset like the other '
                                              'oil products (legacy hardcodes 1 and 0, which drove their retail price '
                                              'to the 0.01 floor); (2) VAT rate: dataset rate where filled, else the '
                                              'general VAT rate VAT_WEO (Egypt 14%) for residential coal and gas and '
                                              'all-sector oil products; power, industry and biomass 0. New policies '
                                              'pay this VAT.'),
        ('Price projection', 'gp = international price (source, gas market, High/Low factor after the last '
                             'historical year). sp = fixsp + (sp(t-1) - fixsp) x gp(t) / gp(t-1), fixsp = margin + '
                             'pass-through x production cost (coal, gas). txo = txo_L x pcc + cs, cs = (sp_L - sp + '
                             'txo_L(1 - pcc)) x (1 - pcc), not below txo_L(1 - pcc) if that is >= 0. rpb = (sp + '
                             'txo) x (1 + VAT rate). atp = max(rpb + nce x (1 + VAT rate), 0.01): new policies are '
                             'fully passed through and pay VAT. Pass-through 0 keeps rpb at its last historical value '
                             '(Egypt gas and oil products); 1 passes supply-cost changes on in full.'),
        ('Variables', 'ctxnew new carbon tax = carbon price x EF x fuel coverage x sector coverage x (1 - ETS '
                      'coverage); ets new ETS = tax-equivalent permit price x EF x effective ETS coverage x (1 - '
                      'OBR share); ntx new excise = fuel '
                      'price reform path / GJ per unit; nce = ctxnew + ets + ntx; atp after-tax price; shp shadow '
                      'price = (sector shadow price x share impacting efficiency + ETS tax-equivalent price x ETS '
                      'coverage x OBR share) x EF; ener fuel use (CPAT eq. 3.3.3): '
                      'the usage term uses atp, the efficiency term (atp + shp), as legacy.'),
        ('ETS (v1.02)', 'MTInputs rows 84-115 per scenario; in covered sectors the ETS replaces the carbon tax. Cap = '
                        'baseline covered emissions (bco2 data from scenario 1 x coverage) x (1 + change path). Price: '
                        'fast estimate LN(cap / baseline) / (effective semi-elasticity x volatility adjustment), or '
                        'the override row ets.ovr when D_ETSPriceOverride = Yes (blank cell: carbon price path). '
                        'Benchmarks by sector group (typed rows) give the OBR share; the permit price splits into a '
                        'tax-equivalent part in the fuel price (x (1 - OBR)) and a shadow price on the efficiency '
                        'margin (x OBR). Tax-equivalent price = permit price x volatility adjustment; revenue uses the '
                        'permit price on the auctioned part. Section 13: covered emissions vs cap and the next price '
                        '(ets.next); ets_goalseek_v0_1.py iterates until the cap is met. See '
                        'ETS_Legacy_Algorithm_v0_1.md.'),
        ('Power prices (v1.05)', 'Section 3, legacy sub-tables: A inputs (capex factor; generation shares, '
                                 'investment shares and consumption are INTERIM legacy-baseline data, orange italic); '
                                 'B amortised fixed cost of the stock = caxav x acap + afix + stoav (vintage averages '
                                 'with last year\'s investment share); C levelised fixed cost of new plants (for '
                                 'investment only); D variable cost = O&M + section 2 fuel price x 0.0036/efficiency, '
                                 'and the new policy cost per kWh ccp (carbon tax, ETS, fuel price reform on power '
                                 'fuels); E generation cost gnc = fix + vbc and the average gncav; G supply cost sc = '
                                 'gncav of the previous year + T&D, price before new policies = data to 2024, then + '
                                 'pass-through x change in sc x (1 + VAT), end-user price rp = rppre + ccpav, '
                                 'subsidy gap and revenues; J storage cost of variable renewables from the VRE share. '
                                 'Constants per type: Inputs_power. Method: PowerPrices_Method_v0_1.md.'),
        ('Emissions (v0.13)', 'co2 per subsector and fuel = fuel use x PJ/ktoe x EF (tCO2/GJ), MtCO2. Section 13: '
                              'total, by sector, by fuel, change vs scenario 1. Fuel combustion only (no power, '
                              'process emissions, CH4, N2O or local pollutants yet).'),
        ('First run in Excel (v1.04)', 'Enable macros (if the file came from the internet: file Properties > '
                                       'Unblock first), then run the macro CheckBatchRun (Alt+F8). Sheet MacroCheck '
                                       'shows Result PASS when the macros find everything they need, Excel\'s own '
                                       'recalculation of the baseline matches the stored baseline, and the batch '
                                       'rerun reproduces every stored value to 1e-6 (relative). If Excel reports '
                                       'that it cannot load the VBA project, import CPATScenarios_v0_2.bas '
                                       '(Alt+F11 > File > Import) and save as .xlsm.'),
        ('Scenarios and batch runs (v1.03)', 'MTInputs: J = scenario 1 (baseline), K = scenario 2 (live policy '
                                             'scenario, Mitigation group 2), L onwards = scenario definitions (row 4 '
                                             'Run? = Yes/No). Macro RunAllScenarios (Alt+F8; module CPATScenarios, '
                                             'enable macros) stores the baseline once, then copies each definition '
                                             'into K, recalculates, runs the ETS goal seek if the definition applies '
                                             'a new ETS, and stores the MTOutputs block as values in StoredResults '
                                             '(K is restored at the end). Other macros: StoreBaseline, '
                                             'StoreLiveScenario, SolveETSLive, ClearStoredScenarios. ScenarioCompare: '
                                             'choose a year and up to 8 stored scenario IDs; levels, differences and '
                                             '% differences vs the baseline. StoredResults ships filled by a Python '
                                             'emulation of the macro.'),
        ('MTOutputs and Charts (v1.01)', 'MTOutputs collects key results (fuel use, CO2, revenues, prices) for every '
                                         'scenario by output code: one row block per scenario, the scenario number in '
                                         'column E of the block header. Copy the last block below to add a scenario. '
                                         'Charts plots them, one series per scenario.'),
        ('Revenues (v0.12)', 'Three separate calculations per subsector and fuel, USD million real: rtx existing '
                             'tax revenue = fuel use x PJ/ktoe x (etx + vat); rsub existing subsidy cost = fuel use x '
                             'PJ/ktoe x esub; rnew new-policy revenue = fuel use x PJ/ktoe x (ctxnew + ntx + ets x '
                             'revenue factor + nce x VAT rate). Section 12: totals and sectors, net existing (rtx - '
                             'rsub), total, and change vs scenario 1. Feebates are revenue-neutral.'),
        ('Existing taxes and subsidies', 'Section 2 by price fuel: vat = (sp + txo) x VAT rate; etx = max(txo, 0) '
                                         'existing excise and other taxes; esub = max(-txo, 0) existing consumer '
                                         'subsidy; esubpu = esub in MTInputs price units ($/liter, $/bbl, $/GJ), '
                                         'ready as a per-fuel policy input (subsidy reform) later. (a) For revenue: '
                                         'fuel use x (etx + vat - esub) plus new policies; see section 12.'),
        ('Roll-up', 'Mitigation opens fully rolled up (zoom 75%): section bands, sector totals, the carbon price, '
                    'retail prices, revenue and CO2 totals. Row groups have their + button BELOW the group: level 1 '
                    'shows subsector headings and inputs, level 2 the variables of each subsector (buttons 1 / 2 / 3 '
                    'at the top left set all levels). Columns (one level): B:G (fuel, sector, parameters), I:K (unit, '
                    'source, scenario-1 code) and 2030-2039 in each scenario group; button 2 above the columns opens '
                    'them.'),
        ('Section 1 (policies)', 'Each input row reads MTInputs by its row number (column D, hidden) and the '
                                 f'scenario number (row {R_SCEN}). Paths: carbon price as legacy (0 before '
                                 'CPIntro, linear to CPLevelTarget by CPOutro, linear continuation if Linear*); '
                                 'fuel price reform and feebates: 0 before the start year, linear from the '
                                 'starting to the final/target value, continuing linearly afterwards (one default). '
                                 'New ETS: cap-based permit price (fast estimate or override), benchmarks give the '
                                 'OBR share, replaces the carbon tax in covered sectors. Shadow prices by sector ($/tCO2) = '
                                 'feebates + regulations (placeholder 0); share impacting efficiency = feebate '
                                 'coverage x adjustment (1.0).'),
        ('Data step vs formulas', 'Lookups happen in Inputs (one row per fuel|subsector), Inputs_prices (one row per '
                                  'price fuel), in Mitigation D:G (hidden; labels per variable in Variables G:J), the '
                                  'base-year column, the top rows, the Section-1 input rows and the gp rows (source '
                                  'is a scenario input). Historical sp and txo read Inputs_prices by position and '
                                  'year. ctxnew, ntx, shp, atp and sp pick a row by position with INDEX(range, '
                                  'position).'),
        ('LAMBDA column', f'The right column ({YEARS[-1]}, light beige) of every calculated row calls a named LAMBDA '
                          '(plain formulas elsewhere); data lookups and sums stay plain. A LAMBDA may hold the IF '
                          'between history and projection (SUPPLYCOST, OTHERTAX), so it can be dragged back over '
                          f'the whole row; the plain columns use separate blocks: data up to {HIST_YEARS[-1]}, '
                          'projection after. Changing the last historical year (Settings C10) therefore also needs '
                          f'the blocks re-dragged. Or drag {YEARS[-2]} forward. Definitions: Formulas > Name '
                          f'Manager. {lam_lines}.'),
        ('Codes', 'Code column before each scenario group: country.mit.<variable>.<sector>.<fuel>.<suffix>.'
                  f'<scenario> (e.g. egy.mit.ener.rod.gso.e.1). Scenario number at the top (row {R_SCEN}).'),
        ('Add a scenario', 'MTInputs: copy the last scenario column one column right (its number updates) and edit '
                           f'its inputs. Mitigation: copy columns {L(code2)}:{L(code2 + GW - 1)} and paste at '
                           f'{L(code2 + GW)}; numbers, names, codes and all policy paths follow.'),
        ('Example', f'Road (heading row {h}): atp rows {h + VOFF["atp"]}-{h + VOFF["atp"] + NF - 1}, ener rows '
                    f'{h + VOFF["ener"]}-{h + VOFF["ener"] + NF - 1}.'),
        ('Colours', 'Green bands with white text = sections; green = input; beige (darker) = base year; LIGHT BEIGE = '
                    'a block that is copy-pasteable within itself but not draggable into white or base-year cells: '
                    'the LAMBDA column (dark red text) and the 2023-2024 history block of supply cost and taxes; '
                    'light blue = codes; grey = unused parameter slot; pale yellow = check; BRIGHT YELLOW with red '
                    'text = data changed by assumption or correction.'),
        ('Data', 'Countries, Elasticities, Prices_dom, Prices_int, PriceAssump, EnergyCons, WEO, EF_GHG, PowerTech, '
                 'PowerPaths, PowerParams; Mapping; '
                 f'Variables; MTInputs. Built by build_v{VERSION.replace(".", "_")}.py; checked by '
                 f'check_v{VERSION.replace(".", "_")}.py.'),
    ]
    for i, (k, v) in enumerate(lines, 4):
        put(ws, f'A{i}', k, FONT_B)
        c = put(ws, f'B{i}', v)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        ws[f'A{i}'].alignment = Alignment(vertical='top')
    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['B'].width = 120


def refresh_bco2(wb):
    """Option A (ETS_Cap_Design_Options_v0_1.md): the baseline CO2 by ETS sector is data. Recalculate the workbook
    (LAMBDAs expanded, LibreOffice), take scenario 1's co2.sec rows (section 13) and write their values into the
    bco2 rows of every scenario group (what a user does with copy / paste values). Scenario 1 has no ETS, so its
    emissions do not depend on these rows."""
    import tempfile
    import openpyxl
    import check_v1_05 as C
    tmp = os.path.join(tempfile.mkdtemp(), 'pre_bco2.xlsx')
    wb.save(tmp)
    vals = C.recalc(C.variant(openpyxl.load_workbook(tmp), 'expand'))
    _, base, years = group_cols(1)
    for s in COV_SECTORS:
        for c in [base] + years:
            v = vals.cell(R_EMI[f'co2.sec.{s}'], c).value
            assert isinstance(v, (int, float)), (s, c, v)
            DATA_VALUES[(f'bco2.{s}', year_of(c))] = v
    ws = wb['Mitigation']
    for g, *_ in SCENARIOS:
        _, base, years = group_cols(g)
        for s in COV_SECTORS:
            for c in [base] + years:
                ws.cell(R_POL[f'bco2.{s}'], c).value = DATA_VALUES[(f'bco2.{s}', year_of(c))]


SR_HEAD, SR_FIRST, SR_Y0 = 5, 6, 8      # StoredResults: header row, first data row, first year column (H)
SR_LAST = 5000
SR_COLS = ['Scenario ID', 'Scenario name', 'Code stem', 'Output code', 'Description', 'Unit', 'Stored']
CMP_STEMS = [(st, fm) for st, fm in MTO_INDICATORS if st not in ('mit.ener.pct.all.all.e', 'mit.co2.chg.all.all',
                                                                 'mit.co2.pct.all.all', 'mit.rtot.chg.all.all')]
CMP_SLOTS = 8                           # ScenarioCompare: scenario columns besides the baseline
CMP_YEAR = 2030


def stored_rows(vals, sid, src, name, stamp):
    """StoredResults rows of MTOutputs block src (values workbook), under scenario ID sid (as the macro)."""
    mo = vals['MTOutputs']
    h = MTO_R0 + (src - 1) * MTO_BLOCK
    assert mo.cell(h, 5).value == src
    name = name or mo.cell(h, 3).value
    out = []
    for k, (stem, _f) in enumerate(MTO_INDICATORS):
        r = h + 1 + k
        country = str(mo.cell(r, 1).value).split('.')[0]
        out.append([sid, name, stem, f'{country}.{stem}.{sid}', mo.cell(r, 3).value, mo.cell(r, 4).value, stamp]
                   + [mo.cell(r, MTO_C0 + j).value for j in range(NY)])
    return out


def emulate_batch(wb, log=print):
    """Python emulation of the macro CPATScenarios.RunAll on the workbook as built: the baseline, then every
    definition with Run? = Yes copied into the live column K (rows 8-415 except the formula rows 10 and 12, and its
    name), recalculated in LibreOffice (LAMBDAs expanded), with the ETS goal seek (ets_goalseek_v0_2.solve, the
    macro's algorithm) when the definition applies a new ETS without the price override. Returns StoredResults
    rows."""
    import tempfile
    import openpyxl
    import check_v1_05 as C
    import ets_goalseek_v0_2 as G
    stamp = 'build (Python emulation of RunAllScenarios)'
    tmp = os.path.join(tempfile.mkdtemp(), 'pre_batch.xlsx')
    wb.save(tmp)
    rows = stored_rows(C.recalc_wb(C.variant(openpyxl.load_workbook(tmp), 'expand')), 1, 1, None, stamp)
    live = MT_COL0 + 1
    for num, name, run, _e in DEFINITIONS:
        if run != 'Yes':
            continue
        w = openpyxl.load_workbook(tmp)
        mt = w['MTInputs']
        d = MT_COL0 + num - 1
        for r in range(8, 416):
            if r not in (10, 12):
                mt.cell(r, live).value = mt.cell(r, d).value
        mt.cell(MT_ROW_NAME, live).value = mt.cell(MT_ROW_NAME, d).value
        path = os.path.join(tempfile.mkdtemp(), f'def{num}.xlsx')
        w.save(path)
        ets = (str(mt.cell(MT_NAMES['D_NewETS'], live).value).upper().startswith('YES')
               and not str(mt.cell(MT_NAMES['D_ETSPriceOverride'], live).value).upper().startswith('YES'))
        if ets:
            p, worst, it = G.solve(path, [2], tol=0.005, max_iter=20, log=lambda _x: None)
            vals = G.recalc_with(path, [2], p)
            log(f'  definition {num}: ETS goal seek, worst |covered/cap - 1| {worst[2]:.4f} after {it} iterations')
        else:
            vals = C.recalc_wb(C.variant(openpyxl.load_workbook(path), 'expand'))
            log(f'  definition {num}: recalculated')
        rows += stored_rows(vals, num, 2, name, stamp)
    return rows


def build_stored_results(wb, rows):
    ws = wb.create_sheet('StoredResults')
    last = SR_Y0 + NY - 1
    title(ws, 'StoredResults - scenario results stored as values (macro RunAllScenarios)', last)
    put(ws, 'A3', 'Values only (no formulas): one block of MTOutputs rows per scenario, under its ID (1 = baseline, '
                  'stored once; definitions keep their MTInputs number). Macros (Alt+F8): RunAllScenarios, '
                  'StoreBaseline, StoreLiveScenario, ClearStoredScenarios. Copy / paste values here to keep results '
                  'from other files. Shipped filled by a Python emulation of RunAllScenarios.', FONT)
    for j, h in enumerate(SR_COLS, 1):
        put(ws, f'{L(j)}{SR_HEAD}', h, FONT_B, F_INPUT)
    for k in range(NY):
        put(ws, f'{L(SR_Y0 + k)}{SR_HEAD}', BASE_YEAR + k, FONT_B, F_INPUT, '0')
    fmt = dict(MTO_INDICATORS)
    for i, row in enumerate(rows, SR_FIRST):
        for j, v in enumerate(row, 1):
            put(ws, f'{L(j)}{i}', v, FONT, None, fmt[row[2]] if j >= SR_Y0 else ('0' if j == 1 else None))
    ws.freeze_panes = f'{L(SR_Y0)}{SR_FIRST}'
    for col, w in {'A': 8, 'B': 40, 'C': 22, 'D': 28, 'E': 48, 'F': 12, 'G': 20}.items():
        ws.column_dimensions[col].width = w
    for k in range(NY):
        ws.column_dimensions[L(SR_Y0 + k)].width = 9
    ws.sheet_view.zoomScale = 85


CHK_FIRST = 6                           # MacroCheck: first report row (as the macro's CHK_FIRST)


def build_macro_check(wb):
    """MacroCheck: report sheet of the macro CheckBatchRun (filled when it runs)."""
    ws = wb.create_sheet('MacroCheck')
    title(ws, 'MacroCheck - first-run test of the scenario macros (macro CheckBatchRun)', 2)
    put(ws, 'A3', 'Run the macro CheckBatchRun (Alt+F8) once after opening the file in Excel with macros enabled. It '
                  'checks the lookups the macros use, compares Excel\'s recalculation of the baseline with the stored '
                  'baseline, reruns RunAllScenarios and compares every stored value with the values stored before. '
                  'Expected: Result PASS (no value differs by more than 1e-6, relative).', FONT)
    ws['A3'].alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells('A3:B3')
    ws.row_dimensions[3].height = 54
    put(ws, 'A5', 'Check', FONT_B, F_INPUT)
    put(ws, 'B5', 'Result', FONT_B, F_INPUT)
    put(ws, f'A{CHK_FIRST}', 'Result', FONT_B)
    put(ws, f'B{CHK_FIRST}', 'not run yet', FONT_B)
    ws.column_dimensions['A'].width = 62
    ws.column_dimensions['B'].width = 70


def build_scenario_compare(wb):
    """Key results of stored scenarios for one year: levels, difference and % difference vs the baseline."""
    from openpyxl.worksheet.datavalidation import DataValidation
    ws = wb.create_sheet('ScenarioCompare')
    first, lastc = 4, 4 + CMP_SLOTS                      # D = baseline, E.. = scenarios
    title(ws, 'ScenarioCompare - stored scenarios vs the baseline in one year (from StoredResults)', lastc)
    put(ws, 'A3', 'Choose the year (C4) and the stored scenario IDs (row 6, E onwards); the baseline is ID 1. Values '
                  'come from StoredResults (run the macro RunAllScenarios to refresh them).', FONT)
    put(ws, 'B4', 'Year', FONT_B)
    put(ws, 'C4', CMP_YEAR, FONT_B, F_INPUT, '0')
    dv = DataValidation(type='list', formula1=f'"{",".join(str(BASE_YEAR + k) for k in range(NY))}"',
                        allow_blank=False)
    ws.add_data_validation(dv)
    dv.add('C4')
    put(ws, 'C6', 'Scenario ID', FONT_B)
    put(ws, 'C7', 'Scenario name', FONT_B)
    ids = [n for n, _nm, run, _e in DEFINITIONS if run == 'Yes']
    sr = lambda col: f'StoredResults!${col}${SR_FIRST}:${col}${SR_LAST}'
    ycol = (f'INDEX(StoredResults!${L(SR_Y0)}${SR_FIRST}:${L(SR_Y0 + NY - 1)}${SR_LAST},0,'
            f'MATCH($C$4,StoredResults!${L(SR_Y0)}${SR_HEAD}:${L(SR_Y0 + NY - 1)}${SR_HEAD},0))')
    for c in range(first, lastc + 1):
        cl = L(c)
        sid = 1 if c == first else (ids[c - first - 1] if c - first - 1 < len(ids) else None)
        put(ws, f'{cl}6', sid, FONT_B, F_BASE if c == first else F_INPUT, '0')
        put(ws, f'{cl}7', f'=IF({cl}$6="","",IFERROR(INDEX({sr("B")},MATCH({cl}$6,{sr("A")},0)),"not stored"))',
            FONT_B)
        ws[f'{cl}7'].alignment = Alignment(wrap_text=True, vertical='top')
    ws.row_dimensions[7].height = 48
    n = len(CMP_STEMS)
    blocks = [('Level in the year', 9), ('Difference vs baseline (scenario - baseline)', 9 + n + 2),
              ('% difference vs baseline', 9 + 2 * (n + 2))]
    for (label, r0), kind in zip(blocks, ('lvl', 'diff', 'pct')):
        band(ws, r0 - 1, label, lastc)
        for k, (stem, fm) in enumerate(CMP_STEMS):
            r, rl = r0 + k, blocks[0][1] + k
            put(ws, f'A{r}', stem, FONT, F_INPUT)
            put(ws, f'B{r}', f'=IFERROR(INDEX({sr("E")},MATCH($A{r},{sr("C")},0)),"")')
            put(ws, f'C{r}', f'=IFERROR(INDEX({sr("F")},MATCH($A{r},{sr("C")},0)),"")')
            for c in range(first, lastc + 1):
                cl = L(c)
                if kind == 'lvl':
                    f = (f'=IF({cl}$6="","",IF(COUNTIFS({sr("A")},{cl}$6,{sr("C")},$A{r})=0,"",'
                         f'SUMIFS({ycol},{sr("A")},{cl}$6,{sr("C")},$A{r})))')
                    put(ws, f'{cl}{r}', f, FONT, F_BASE if c == first else None, fm)
                elif c > first:
                    f = (f'=IF(OR({cl}{rl}="",$D{rl}=""),"",{cl}{rl}-$D{rl})' if kind == 'diff' else
                         f'=IF(OR({cl}{rl}="",$D{rl}="",$D{rl}=0),"",{cl}{rl}/$D{rl}-1)')
                    put(ws, f'{cl}{r}', f, FONT, None, fm if kind == 'diff' else '0.0%')
    ws.freeze_panes = 'D8'
    for col, w in {'A': 22, 'B': 52, 'C': 12}.items():
        ws.column_dimensions[col].width = w
    for c in range(first, lastc + 1):
        ws.column_dimensions[L(c)].width = 16
    ws.sheet_view.zoomScale = 85


def main():
    wb = Workbook()
    build_settings(wb)
    build_mapping(wb)
    build_inputs(wb)
    build_inputs_prices(wb)
    build_inputs_power(wb)
    build_legacy_diff(wb)
    build_variables(wb)
    build_mtinputs(wb)
    build_mitigation(wb)
    build_mtoutputs(wb)
    build_charts(wb)
    build_stored_results(wb, [])
    build_scenario_compare(wb)
    build_macro_check(wb)
    div = wb.create_sheet('DATA->')
    put(div, 'A1', 'Data sheets follow', FONT_B)
    build_data_sheets(wb)
    build_readme(wb)
    for name in LAMBDAS:
        wb.defined_names[name] = DefinedName(name, attr_text=lambda_xml(name), comment=LAMBDA_NOTES[name])
    order = ['Mitigation', 'MTOutputs', 'StoredResults', 'ScenarioCompare', 'MacroCheck', 'Charts', 'ReadMe', 'LegacyDiff',
             'Settings',
             'MTInputs', 'Inputs', 'Inputs_prices', 'Inputs_power', 'Variables', 'Mapping', 'DATA->',
             'Countries', 'Elasticities', 'Prices_dom', 'Prices_int', 'PriceAssump', 'EnergyCons', 'WEO', 'EF_GHG',
             'PowerTech', 'PowerPaths', 'PowerParams']
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    for w in wb.worksheets:
        w.sheet_view.tabSelected = w.title == 'Mitigation'
    refresh_bco2(wb)
    rows = emulate_batch(wb)
    idx = wb.sheetnames.index('StoredResults')
    del wb['StoredResults']
    build_stored_results(wb, rows)
    wb._sheets.insert(idx, wb._sheets.pop(wb.sheetnames.index('StoredResults')))
    wb.code_name = 'ThisWorkbook'
    for k, w in enumerate(wb.worksheets, 1):           # code names = document modules of the VBA project
        w.sheet_properties.codeName = f'Sheet{k}'
    import tempfile
    from vba_project_v0_1 import embed_vba
    tmp = os.path.join(tempfile.mkdtemp(), f'{NAME}-v{VERSION}.xlsx')
    wb.save(tmp)
    fix_outline_levels(tmp)
    embed_vba(tmp, OUT, {'CPATScenarios': vba_module_source()})
    print('Saved', OUT)


def vba_module_source():
    """The module text of CPATScenarios_v0_2.bas without its Attribute VB_Name line (added by the VBA writer)."""
    text = open(BAS, encoding='cp1252').read()
    assert text.startswith('Attribute VB_Name = "CPATScenarios"')
    return text.split('\n', 1)[1]


if __name__ == '__main__':
    main()
