"""Build CPAT-AI-Mitigation-MVP-v1.07.xlsm (design 2: time across, scenarios as column groups).

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
    goal seek: ets_goalseek_v0_2.py; v1.07: ets_goalseek_v0_3.py).
v1.04: first-run test for Excel. Module CPATScenarios_v0_2.bas adds the macro CheckBatchRun: it checks that the
    macros find every sheet, row and column they use, compares the host's own recalculation of the baseline with the
    stored baseline (Excel evaluates the LAMBDAs natively), reruns the batch and compares every stored value with the
    values stored before; the report goes to the new sheet MacroCheck (Result PASS / CHECK / FAIL).
v1.05: power sector, price side (PowerPrices_Method_v0_2.md): Mitigation section 3 with legacy sub-tables A Inputs,
    B amortised fixed costs (vintages), C levelised fixed costs, D variable costs and new policy cost per kWh, E total
    and average generation cost, G end-user electricity prices and power revenues, J storage cost of variable
    renewables. Data step: sheet Inputs_power (one row per generation type, derived constants) and data sheets
    PowerTech, PowerPaths, PowerParams (extract_power_v0_1.py). Generation shares, investment shares and
    consumption are INTERIM data (legacy baseline) until the engineer model is built. MTOutputs adds the residential
    and industrial electricity prices and the average generation cost.
v1.06: power sector, engineer model (PowerModel_Method_v0_1.md): sub-tables H demand (by subsector, elasticities of
    electricity, GDP, last year's price), I dispatch (fixed-output types at their capacity factor, coal and gas by
    minima and a logit on variable cost), K capacity and retirements (incl. cost-based early coal retirement), L
    investment (logit on levelised cost within scale-up limits, then cheapest remaining space), D6 shadow price and
    rebate, D9 levelised variable cost, E1 investment cost; generation shares, investment shares and consumption are
    now modelled (no interim data). Aligned with legacy except for ten bugs (power_model_ref_v0_1.FIXES;
    PowerModel_BugsAndImprovements_v1_0.docx). Power CO2 enters section 13 (total, by fuel, ETS sector power).
v1.07: short formulas: every cell formula that does not call a named LAMBDA has at most 100 characters (the LAMBDAs
    and the right-column calls may be longer). Tools: defined names for the data tables (NAMES; e.g. WEOVal, WEOKey,
    WEOYear), helper rows and columns that show intermediate steps (carbon-price path years and nominal factor; ETS
    on/off, volatility and impact, cap change path, log ratio of the next price; international price before the
    adjustment; early-retirement cost floor and coal logit share; Inputs_prices data by year; MTOutputs and
    ScenarioCompare lookup positions), fuel-use parameters G = AEEI factor and F = efficiency exponent (Inputs), and
    the LAMBDA ROWLABEL in column H. Planned retirement by generation type (PowerTech planned_retirement, PowerPaths
    pret|<type>; Egypt: no data, switches No). Memo lines of electricity demand by subsector (sector blocks) and by
    sector (section 11; not in the fuel totals: power demand keeps its own elasticities, sector policies do not act
    on it). Power indicators on MTOutputs and Charts. User decisions: the 2023 demand jump (USD electricity prices
    after the devaluation) is accepted; sector policies stay separate from power demand.
v1.08: the workbook works without LAMBDA support (user: main functionality without LAMBDAs). Every cell holds a plain
    formula of at most 100 characters, the right column (2040) included; column H reads the member labels from a
    table on Mapping (no ROWLABEL). The named LAMBDAs stay in the Name Manager as an optional library: sheet Lambdas
    lists, for every calculated row, the LAMBDA formula that can replace the plain formulas of the row (tested by
    the checks: copied over whole rows, same results).
v1.09: existing carbon pricing and power in the main thread. Sheet ECP holds the legacy existing carbon pricing data
    (all countries, extract_ecp_v0_1.py) and the selected country's values (real prices, coverage shares); section 1
    reads the MTInputs switches (rows 208-214), the existing carbon tax and ETS prices used (data, growth after the
    last data year, switch) and the level embedded in the price data. Per subsector and fuel: xcp (level, $/GJ) and
    xcpd (change against the price data, in the after-tax price and in existing-tax revenue); rxcp revenue memo.
    Power: xck (change per kWh) in the variable cost; revenues by generation type (sub-table M) and electricity VAT
    and subsidy gap enter section 12 (rows .pow); generation fuel use enters section 11 (row ener.pow, totals by
    fuel). CTaxComplimentaryToETS = Yes: the new carbon tax covers 1 - existing ETS sector coverage. Check rows
    rev.chk and co2.chk.

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

VERSION = '1.09'
LAMBDA_ROWS = []                        # v1.08: (row, variable, fuel, sector, LAMBDA formula of the scenario-1 right
                                        # column) - the optional LAMBDA library, written to sheet Lambdas, not to cells
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
ECP_PARAMS = ['ExistingCTApply', 'ExistingCTGrowth', 'ExistingETSApply', 'ExistingETSGrowth',
              'CTaxComplimentaryToETS']   # v1.09: MTInputs rows 208-214
ECP_YEARS = list(range(2021, 2026))     # legacy ECP data years used (US CPI from 2021; legacy sheet 2015-2025)
ECP_SECTOR = {'pow': 'pow', **{s[0]: ('trs' if s[2] == 'tra' else ('res' if s[2] == 'bld' else 'ind'))
                               for s in SUBSECTORS}}   # subsector -> legacy ECP sector (pow, trs, res, ind)
ECP_CODES = (['ecr.all', 'eer.all'] + [f'ecs.{s}' for s in ('pow', 'trs', 'res', 'ind')]
             + [f'ecf.{f[0]}' for f in FUELS] + [f'ees.{s}' for s in ('pow', 'trs', 'res', 'ind')]
             + [f'eef.{f[0]}' for f in FUELS])
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
    ('Price sector rule of the fuel (sub / all)', 'prule'), ('Coal/gas price sector of the subsector', 'psec'),
    ('Price code', 'pcode'), ('EF sector', 'efsec'), ('GJ per price unit', 'gj'),
    ('Fuel position (coverage)', 'fpos'), ('Sector position (coverage)', 'spos'),
    ('Price-fuel position (fuel price reform)', 'pfpos'),
    ('eps_Y GDP', 'eY'), ('eps_U usage', 'eU'), ('eps_F efficiency', 'eF'), ('alpha AEEI', 'a'),
    ('AEEI factor (1/(1+alpha))^(1+eps_U)', 'aeef'), ('Efficiency exponent eps_F x (1+eps_U)', 'effx'),
    ('Price position (section 2)', 'ppos'), ('VAT rate', 'vr'),
    ('EF (tCO2/GJ)', 'ef'), ('Base-year fuel use (ktoe)', 'f0'),
    ('ECP sector of the subsector (legacy: pow, trs, res, ind)', 'ecps'),
    ('Existing carbon tax coverage: sector x fuel share (sheet ECP)', 'ctc'),
    ('Existing ETS coverage: sector x fuel share (sheet ECP)', 'etc'),
]
IC = {code: L(j) for j, (_, code) in enumerate(INPUT_COLS, 1)}

# ---------------------------------------------------------------- Inputs_prices sheet (one row per price fuel)
IP_PROW, IP_CPI, IP_VATW, IP_HDR = 3, 4, 5, 6   # v1.07: Prices_dom row, CPI index, VAT_WEO of the year; header
IP_R0 = 7
IP_R1 = IP_R0 + NP - 1
HIST_YEARS = [2021, 2022, 2023, 2024]   # years in the IMF price dataset (Prices_dom)
IP_COLS = [
    ('Price code (fuel.sector)', 'key'), ('Label', 'lab'), ('Fuel', 'fuel'), ('Sector', 'sec'),
    ('Position (section 2 row)', 'pos'), ('International price position (gp row)', 'gpos'),
    ('GJ per price unit', 'gj'), ('Retail price from data? (1 = coal, gas, biomass)', 'rpd'),
    ('Raw pass-through coefficient (mit.ps, base year)', 'pccr'), ('Bucketed pass-through', 'pccb'),
    ('Chosen pass-through (Settings: price controls)', 'pcc'), ('Margin (mit.mar, base year, nominal price unit)', 'mar'),
    ('Domestic production cost (base year, nominal $/GJ)', 'pcost'),
    ('CPI index of the base year', 'bcpi'), ('Fixed supply cost fixsp (real $/GJ)', 'fixsp'),
    ('VAT applies to final consumers? (assumption: residential and all-sector oil products)', 'vatc'),
    ('VAT rate assumption = general VAT rate (VAT_WEO) x consumer flag, used only where the data cell is blank', 'vrx'),
    ('VAT rate used (data where filled, 0 included; else assumption), last historical year', 'vr'),
    ('Supply cost, last historical year (real $/GJ)', 'spL'), ('Excise and other taxes, last historical year (real $/GJ)', 'txoL'),
] + [(f'Supply cost {y} (real $/GJ)', f'sp{y}') for y in HIST_YEARS] + \
    [(f'Excise and other taxes {y} (real $/GJ)', f'txo{y}') for y in HIST_YEARS] + \
    [(f'{lab} {y}', f'{code}{y}') for code, lab in (   # v1.07: data of each year (helper columns)
        ('rpn', 'Retail price, nominal per price unit (Prices_dom mit.rp; blank = 0)'),
        ('spn', 'Supply cost, nominal per price unit (mit.sp; blank = 0)'),
        ('txn', 'Excise and other taxes, nominal per price unit (mit.txo; blank = 0)'),
        ('vrb', 'VAT rate cell blank in the dataset (1 / 0)'), ('vrd', 'VAT rate in the dataset (blank = 0)'),
        ('vry', 'VAT rate used (data where filled; blank: VAT_WEO x consumer flag)')) for y in HIST_YEARS]
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
_pol('cp.t', 'cp.t', kind='calc')                                         # v1.07: years along the price path
_pol('cp.nom', 'cp.nom', kind='calc')                                     # v1.07: nominal-to-real factor
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
_pol('ets.on', 'ets.on', kind='calc')                                     # v1.07: new ETS applied and started
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
for k_ in ('ets.vol', 'ets.imp', 'ets.vadj', 'ets.bce', 'ets.chg', 'ets.cap', 'ets.se', 'ets.est'):
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
for s in COV_SECTORS:                                                      # v1.07: policy rates by sector ($/tCO2)
    _pol(f'ctr.{s}', 'ctr', 'all', s, kind='calc')                          # carbon tax (not where the ETS covers)
for s in COV_SECTORS:
    _pol(f'etr.{s}', 'etr', 'all', s, kind='calc')                          # ETS cost in the fuel price
for s in COV_SECTORS:
    _pol(f'shpx.{s}', 'shpx', 'all', s, kind='calc')                        # shadow price on the efficiency margin
for p_ in ECP_PARAMS:                                                      # v1.09: existing carbon pricing
    _pol(p_, p_, mt=MT_NAMES[p_])
for sch in ('xct', 'xets'):
    for k_ in ('d', 'p', 'p0'):
        _pol(f'{sch}.{k_}', f'{sch}.{k_}', kind='ecp' if k_ != 'p' else 'calc')
CONST_VALUES = {'fc.bio': True, 'ssc.adj': 1, **{f'reg.{s}': 0 for s, *_ in FEEBATE_SECTORS},
                **{f'etsbs.{g}': b0 for g, b0, _b1 in ETS_GROUPS}, **{f'etsbt.{g}': b1 for g, _b0, b1 in ETS_GROUPS}}
ETS_GROUP_OF = {'pow': 'pow', **{s[0]: s[6] for s in SUBSECTORS}}           # ETS sector -> benchmark group
ETS_SEMI_OF = {'pow': ETS_SEMI['pow'], **{s[0]: ETS_SEMI['tra' if s[2] == 'tra' else ('res' if s[0] == 'res' else 'ind')]
                                          for s in SUBSECTORS}}           # legacy sector table (foo, srv: industry)
POL_PARAMS = {f'obr.{s}': ([g for g, *_ in ETS_GROUPS].index(ETS_GROUP_OF[s]) + 1, ETS_SEMI_OF[s])
              for s in COV_SECTORS}                                         # D: benchmark group, E: semi-elasticity
POL_PARAMS.update({f'shpx.{s}': ([f for f, *_ in FEEBATE_SECTORS].index(ETS_GROUP_OF[s]) + 1,)
                   for s in COV_SECTORS})                                   # D: feebate sector position
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
            'ets.vol': 'ETS price: volatility and impact (Settings table), volatility adjustment, baseline covered '
                       'emissions, cap change path, cap, effective semi-elasticity, fast estimate, override row (pasted '
                       'values), price used, tax-equivalent price, revenue factor',
            'fb.yr0': 'Feebates (MTInputs rows 53-64; rates in USD/tCO2)',
            'fbc.pow': 'Feebates sector coverage (MTInputs rows 66-82)',
            'ssc.adj': 'Shadow prices: regulations (placeholder), by sector ($/tCO2 = feebates + regulations; legacy '
                       'rows 2345-2349) and share impacting efficiency (legacy rows 2402-2419)',
            'ExistingCTApply': 'Existing carbon pricing (v1.09; MTInputs rows 208-214, sheet ECP): switches and growth, '
                               'carbon tax and ETS price (data of the year, real; last data year x (1 + growth) after; '
                               'used = data x switch; embedded = level in the price data, held after the last '
                               'historical price year)',
            'ctr.pow': 'Policy rates by sector ($/tCO2; v1.07): carbon tax (carbon price x sector coverage x (1 - ETS '
                       'coverage)), ETS cost in the fuel price (tax-equivalent price x ETS coverage x (1 - OBR)), '
                       'shadow price on the efficiency margin (feebates and regulations x share + ETS tax-equivalent '
                       'price x ETS coverage x OBR; D: feebate sector position)'}

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
R_PRI['gp.src'] = B_PRI + 2 + len(PRI_PARAMS)   # v1.07: price source key (selector without the *)
R_PRI['gp.adj'] = R_PRI['gp.src'] + 1   # v1.07: adjustment choice (1 none, 2 High, 3 Low)
R_GPN0 = R_PRI['gp.adj'] + 1            # v1.07: gpn rows, international price as in Prices_int (nominal)
R_GP0 = R_GPN0 + len(GP_ROWS)           # gp rows (oil, coa, nga, one)
R_PRI['head.dom'] = R_GP0 + len(GP_ROWS)
R_PV0 = R_PRI['head.dom'] + 1           # sp, txo, rpb blocks over the 12 price fuels
PVARS = ['sp', 'txf', 'txo', 'rpb', 'vat', 'etx', 'esub', 'esubpu']     # v1.07: txf (floating part of txo)
PVOFF = {v: k * (NP + 1) for k, v in enumerate(PVARS)}
B_POW = R_PV0 + len(PVARS) * (NP + 1) + 1   # 3. Power sector: engineer model, generation costs, prices (v1.05-v1.06)
PW_TECH = [('coa', 'Coal'), ('nga', 'Natural gas'), ('oop', 'Oil'), ('nuc', 'Nuclear'), ('wnd', 'Wind'),
           ('sol', 'Solar'), ('hyd', 'Hydro'), ('ore', 'Other renewables'), ('bio', 'Biomass')]   # legacy order
NT = len(PW_TECH)
PW_GROUPS = [('res', 'Residential'), ('ind', 'Industry')]     # end-user price groups (legacy G1, G2)
DEM_GROUPS = [('tra', 'Transport'), ('res', 'Residential'), ('ind', 'Industry'), ('srv', 'Services')]   # legacy H1
DEM_SUBS = [r[0] for r in csv.reader(open(os.path.join(DATA, 'power_demand.csv'), encoding='utf-8'))][1:]
DEM_GROUP_OF = {'rod': 'tra', 'ral': 'tra', 'avi': 'tra', 'nav': 'tra', 'res': 'res', 'srv': 'srv'}   # others: ind
PW_LAYOUT = [  # legacy sub-table letter, heading, items (kind, variable)
    # kind: tech (9 rows, one per generation type), user (2: residential, industry), dgroup (4 demand groups),
    #       dsub (15 subsectors), single (1 row)
    ('A', 'A. Inputs: capex time factor, storage cost paths, income factor on GDP elasticities, exogenous demand shock',
     [('tech', 'tcf'), ('single', 'cbat'), ('single', 'cint'), ('single', 'obat'), ('single', 'cel'),
      ('single', 'incf'), ('single', 'shock')]),
    ('B', 'B. Amortised fixed costs of the existing stock ($/kWh): capital cost, vintage averages of capital and '
          'storage cost (investment share of the previous year, from L)', [('tech', 'cax'), ('tech', 'caxav'),
                                                                           ('tech', 'stoav'), ('tech', 'fix')]),
    ('C', 'C. Levelised fixed costs of new plants ($/kWh; for investment decisions, not for prices)', [('tech', 'lfx')]),
    ('D', 'D. Variable costs ($/kWh): change in existing carbon pricing against the price data (in the cost before '
          'new policies), before new policies, new policy cost, shadow price and rebate (feebates, ETS '
          'output-based part), total; levelised variable cost to the end of the horizon (D9)',
     [('tech', 'xck'), ('tech', 'vbc'), ('tech', 'ccp'), ('tech', 'shk'), ('single', 'reb'), ('tech', 'tvc'),
      ('tech', 'dsum'), ('tech', 'lvc')]),
    ('E', 'E. Total costs: generation cost by type and average; average new policy cost per kWh consumed; levelised '
          'investment cost (E1, for investment)', [('tech', 'gnc'), ('single', 'gncav'), ('single', 'ccpav'),
                                                    ('tech', 'tic')]),
    ('G', 'G. End-user electricity prices (real $/kWh; supply cost lagged one year) and power revenues (USD million)',
     [('user', 'cons'), ('user', 'sc'), ('user', 'rppre'), ('user', 'rp'), ('user', 'sgap'), ('user', 'rvat'),
      ('user', 'rcarb'), ('user', 'rnvat'), ('user', 'rgap')]),
    ('H', 'H. Power demand (ktoe): price of the previous year by demand group, growth multiple (efficiency, GDP, price), '
          'demand by subsector and by demand group, total, required generation (MWy)',
     [('dgroup', 'pdem'), ('dgroup', 'mult'), ('dsub', 'dem'), ('dgroup', 'demg'), ('single', 'demtot'),
      ('single', 'req')]),
    ('I', 'I. Dispatch (MWy): fixed-output types at their capacity factor; coal and gas fill the rest (minimum each, '
          'then a logit on variable cost); generation shares, fuel use (ktoe), CO2 (MtCO2)',
     [('single', 'reqcg'), ('single', 'vcmin'), ('single', 'rem'), ('tech', 'gmin'), ('tech', 'lgd'), ('tech', 'gen'),
      ('single', 'gtot'), ('single', 'dgen'), ('tech', 'shr'), ('tech', 'pe'), ('tech', 'co2')]),
    ('J', 'J. Systems integration cost of variable renewables: storage ($/kWh)',
     [('single', 'vre'), ('single', 'cph'), ('single', 'msc')]),
    ('K', 'K. Capacity (MW) and retirements: nameplate and effective capacity, planned retirement (data, used where '
          'switched on), retirements, cost-based early coal retirement, effective capacity needed',
     [('tech', 'cap'), ('single', 'captot'), ('tech', 'effcap'), ('tech', 'pret'), ('tech', 'ret'), ('tech', 'reteff'),
      ('single', 'egrow'), ('single', 'ecmin'), ('single', 'ecshr'), ('single', 'early'), ('single', 'kafter'),
      ('single', 'need')]),
    ('L', 'L. Investment (MW effective, MW nameplate, USD million): logit on levelised cost within scale-up limits, '
          'then the cheapest remaining space; share of new investment in the stock (for B)',
     [('single', 'cmin'), ('single', 'deficit'), ('tech', 'permit'), ('tech', 'maxinv'), ('tech', 'lgi'),
      ('tech', 'pni'), ('tech', 'first'), ('tech', 'space'), ('tech', 'second'), ('tech', 'nei'), ('tech', 'nim'),
      ('single', 'nimtot'), ('tech', 'usd'), ('single', 'usdtot'), ('tech', 'phi')]),
    ('M', 'M. Power-sector revenues (USD million, real; v1.09): existing taxes and subsidies on generation fuels, change '
          'in existing carbon pricing, new policies (ETS at the permit price), existing carbon pricing (memo); with '
          'electricity VAT and subsidy gap (G) they form the power rows of section 12',
     [('tech', 'prtx'), ('tech', 'prsub'), ('tech', 'prxc'), ('tech', 'prnew'), ('tech', 'prxl')]),
]
PW_DATA = {'tcf', 'cbat', 'cint', 'obat', 'cel', 'incf', 'shock', 'pret'}   # data lookups (PowerPaths), no LAMBDA
PW_SUMS = {'demtot', 'gtot', 'usdtot', 'captot', 'nimtot', 'demg'}       # plain sums, no LAMBDA
PW_NOLAM = {'dsum', 'lvc'}            # look one column ahead: plain formula in every column, no LAMBDA
PW_BASE = {'caxav', 'stoav', 'sc', 'pdem', 'mult', 'dem', 'req', 'gen', 'cap', 'egrow', 'reb', 'dgen'}  # own base year
PW_SUMMARY = {'gncav', 'rp', 'gtot', 'vre', 'captot'}                  # visible when rolled up
PW_MEMBERS = {'tech': [m for m, _n in PW_TECH], 'user': [m for m, _n in PW_GROUPS],
              'dgroup': [m for m, _n in DEM_GROUPS], 'dsub': DEM_SUBS}
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
        _mem = PW_MEMBERS[_kind]
        for _k, _m in enumerate(_mem):
            R_PW[(_v, _m)] = _r + _k
            _bc = {'tech': (_m, 'pow'), 'user': ('ecy', _m), 'dgroup': ('ecy', _m), 'dsub': ('ecy', _m)}[_kind]
            PW_ROWS.append((_r + _k, _v, _bc[0], _bc[1], _kind))
        _r += len(_mem) + 1
        _single = False
    if _single:
        _r += 1
R_POW_NOTE = _r
PW_KIND = {v: k for _l, _h, items in PW_LAYOUT for k, v in items}
PW_VARS = [v for _l, _h, items in PW_LAYOUT for _k, v in items]
PW_PARAMS = {  # Mitigation D:G of each power variable: ('key',) data key | ('t', col) Inputs_power by type |
    # ('g', col) by price group | ('h', col) by demand group | ('s', col) by subsector | ('p', key) PowerParams |
    # ('q', key) Inputs_power derived parameter | ('c', value) constant
    'tcf': [('key',)], 'cbat': [('key',)], 'cint': [('key',)], 'obat': [('key',)], 'cel': [('key',)],
    'incf': [('key',)], 'shock': [('key',)],
    'cax': [('t', 'cax0')], 'caxav': [], 'stoav': [('t', 'vre')], 'fix': [('t', 'acap'), ('t', 'afix')],
    'lfx': [('t', 'lcap'), ('t', 'lfix'), ('t', 'vre')], 'vbc': [('t', 'vfix'), ('t', 'ppos'), ('t', 'k')],
    'xck': [('t', 'efk'), ('t', 'ctc'), ('t', 'etc')],
    'prtx': [('t', 'ppos'), ('t', 'fflag')], 'prsub': [('t', 'ppos'), ('t', 'fflag')], 'prxc': [], 'prnew': [],
    'prxl': [('t', 'ctc'), ('t', 'etc')], 'rnvat': [('g', 'pass'), ('g', 'vr')],
    'ccp': [('t', 'efk'), ('t', 'fpos'), ('t', 'pfpos'), ('t', 'kpu')], 'shk': [('t', 'efk'), ('t', 'shr0')],
    'reb': [], 'tvc': [], 'dsum': [('t', 'wacc')], 'lvc': [('t', 'wacc')], 'gnc': [],
    'gncav': [('c', 1)], 'ccpav': [('p', 'gen_per_cons')], 'tic': [('p', 'min_cost')],
    'cons': [('g', 'wtot'), ('g', 'wres'), ('p', 'ktoe_gwh')], 'sc': [('g', 'tmc')],
    'rppre': [('g', 'pass'), ('g', 'pos'), ('g', 'vr')], 'rp': [('g', 'pass'), ('g', 'vr')], 'sgap': [('g', 'vr')],
    'rvat': [('g', 'vshare')], 'rcarb': [('c', 1)], 'rgap': [('c', 1)],
    'pdem': [('h', 'wres'), ('h', 'p2021')], 'mult': [('h', 'aeef'), ('h', 'ey'), ('h', 'eps')],
    'dem': [('s', 'gpos'), ('s', 'd0')], 'demg': [('h', 'pos')], 'demtot': [], 'req': [('q', 'req_factor')],
    'reqcg': [], 'vcmin': [('p', 'min_cost')], 'rem': [], 'gmin': [('t', 'cg'), ('p', 'maxcf_cg')],
    'lgd': [('t', 'cg'), ('p', 'k_dispatch'), ('p', 'min_cost')], 'gen': [('t', 'cg'), ('p', 'maxcf_cg'), ('t', 'gen0')],
    'gtot': [], 'dgen': [], 'shr': [('t', 'vsh')], 'pe': [('t', 'pef'), ('t', 'fflag')], 'co2': [('t', 'co2f')],
    'vre': [('c', 1)], 'cph': [('p', 'st_ratio'), ('p', 'dlf_ren')],
    'msc': [('p', 'st_alloc'), ('p', 'st_marg_hours'), ('p', 'lt_start'), ('p', 'lt_alloc')],
    'cap': [('t', 'cap0')], 'captot': [], 'effcap': [('t', 'cf')], 'pret': [('key',)],
    'ret': [('t', 'rr'), ('t', 'eflag'), ('t', 'puse')], 'reteff': [('t', 'cf')],
    'egrow': [], 'ecmin': [], 'ecshr': [('p', 'k_invest')],
    'early': [('p', 'early_max'), ('p', 'early_factor'), ('q', 'cf_coal')],
    'kafter': [], 'need': [],
    'cmin': [('p', 'min_cost')], 'deficit': [], 'permit': [('t', 'perm'), ('t', 'start')], 'maxinv': [('t', 'rate')],
    'lgi': [('p', 'k_invest')], 'pni': [], 'first': [], 'space': [], 'second': [('t', 'tb')], 'nei': [],
    'nim': [('t', 'ninv')], 'nimtot': [], 'usd': [], 'usdtot': [], 'phi': []}
PW_FMT = {'tcf': '0.000', 'cbat': '#,##0.0', 'cint': '#,##0.0', 'obat': '0.00', 'cax': '#,##0', 'caxav': '#,##0',
          'cons': '#,##0', 'rvat': '#,##0.0', 'rcarb': '#,##0.0', 'rgap': '#,##0.0', 'rnvat': '#,##0.0',
          'prtx': '#,##0.0', 'prsub': '#,##0.0', 'prxc': '#,##0.0', 'prnew': '#,##0.0', 'prxl': '#,##0.0',
          'xck': '0.00000', 'vre': '0.000', 'incf': '0.0000',
          'shock': '0.0%', 'pdem': '0.0000', 'mult': '0.0000', 'dem': '#,##0.0', 'demtot': '#,##0.0', 'req': '#,##0',
          'reqcg': '#,##0', 'rem': '#,##0', 'gmin': '#,##0', 'lgd': '0.0000', 'gen': '#,##0', 'gtot': '#,##0',
          'dgen': '#,##0', 'shr': '0.0%', 'pe': '#,##0', 'co2': '#,##0.00', 'cap': '#,##0', 'effcap': '#,##0',
          'ret': '#,##0', 'reteff': '#,##0', 'egrow': '#,##0', 'early': '#,##0', 'kafter': '#,##0', 'need': '#,##0',
          'deficit': '#,##0', 'permit': '0', 'maxinv': '#,##0', 'lgi': '0.0000', 'pni': '0.0%', 'first': '#,##0',
          'space': '#,##0', 'second': '#,##0', 'nei': '#,##0', 'nim': '#,##0', 'usd': '#,##0', 'usdtot': '#,##0',
          'phi': '0.0000', 'dsum': '0.000', 'pret': '#,##0', 'captot': '#,##0', 'nimtot': '#,##0', 'demg': '#,##0.0',
          'ecmin': '0.0000', 'ecshr': '0.0000'}
VARS = ['ctxnew', 'ets', 'ntx', 'nce', 'xcp', 'xcpd', 'atp', 'shp', 'ener', 'co2', 'rtx', 'rsub', 'rnew',
        'rxcp']                         # per subsector, 8 fuels (v1.09: xcp, xcpd, rxcp)
REV_VARS = ['rtx', 'rsub', 'rnew', 'rxcp']
NVAR = len(VARS)
VSTEP = NF + 1                          # 8 fuel rows + 1 blank row
SB = 2 + NVAR * VSTEP                   # subsector block: heading line, electricity memo line (v1.07), variables
VOFF = {v: 2 + k * VSTEP for k, v in enumerate(VARS)}   # first row of each variable relative to the heading
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
R_EPOW = R_SUB0 + NS                    # v1.09: power generation fuel use (after the subsectors)
R_FUEL0 = R_EPOW + 2
R_CHK = R_FUEL0 + NF + 1
R_REF = R_CHK + 2
R_PCT = R_REF + 1
ELD_GROUPS = ['all'] + [g for _n, g, _t in SECTIONS]
R_ELD = {g: R_PCT + 2 + k for k, g in enumerate(ELD_GROUPS)}   # v1.07: electricity demand memo by sector
B_REV = R_ELD[ELD_GROUPS[-1]] + 2       # 12. Revenues
REV_ROWS = []                           # (key, var, fuel, sector) in order; groups = sector sections
for v_ in REV_VARS:
    REV_ROWS.append((f'{v_}.all', v_, 'all', 'all'))
    REV_ROWS += [(f'{v_}.{g}', v_, 'all', g) for _n, g, _t in SECTIONS]
    REV_ROWS.append((f'{v_}.pow', v_, 'all', 'pow'))                 # v1.09: power sector (section 3 G and M)
REV_ROWS += [('rnet.all', 'rnet', 'all', 'all'), ('rtot.all', 'rtot', 'all', 'all'),
             ('rtot.ref', 'rtot.ref', 'all', 'all'), ('rtot.chg', 'rtot.chg', 'all', 'all'),
             ('rev.chk', 'rev.chk', 'all', 'all')]
R_REV = {k: B_REV + 1 + i for i, (k, *_r) in enumerate(REV_ROWS)}
R_REV_NOTE = B_REV + 1 + len(REV_ROWS) + 1
B_EMI = R_REV_NOTE + 4                  # 13. Emissions: CO2 from fuel combustion
EMI_ROWS = ([('co2.all', 'co2', 'all', 'all')] + [(f'co2.{g}', 'co2', 'all', g) for _n, g, _t in SECTIONS]
            + [('co2.pow', 'co2', 'all', 'pow')]                                     # v1.06: power generation
            + [(f'co2.f.{f}', 'co2', f, 'all') for f, *_x in FUELS]
            + [('co2.ref', 'co2.ref', 'all', 'all'), ('co2.chg', 'co2.chg', 'all', 'all'),
               ('co2.pct', 'co2.pct', 'all', 'all')]
            + [(f'co2.sec.{s}', 'co2.sec', 'all', s) for s in COV_SECTORS]
            + [(k, k, 'all', 'all') for k in ('co2.ets', 'co2.cap', 'co2.gap', 'ets.lr', 'ets.next', 'bco2.chk',
                                              'co2.chk')])
R_EMI = {k: B_EMI + 1 + i for i, (k, *_r) in enumerate(EMI_ROWS)}
R_EMI_NOTE = B_EMI + 1 + len(EMI_ROWS) + 1


def var_rows(var):
    """All (row, subsector, fuel) of a variable across the sector sections."""
    return [(SUB_HEAD[s] + VOFF[var] + i, s, f) for s, *_ in SUBSECTORS for i, (f, *_r2) in enumerate(FUELS)]


# Hidden parameter columns D:G for each variable (Inputs columns), and their labels.
VAR_PARAMS = {'ctxnew': ['ef', 'fpos', 'spos'], 'ets': ['ef', 'spos'], 'ntx': ['pfpos', 'gj'], 'nce': [],
              'xcp': ['ef', 'ctc', 'etc'], 'xcpd': ['ef', 'ctc', 'etc'],
              'atp': ['ppos', 'vr'], 'shp': ['ef', 'spos'], 'ener': ['eY', 'eU', 'effx', 'aeef'],
              'co2': ['ef'], 'rtx': ['ppos', 'vr'], 'rsub': ['ppos'], 'rnew': ['vr'], 'rxcp': []}
PVAR_PARAMS = {'sp': ['fixsp', 'gpos', 'pos'], 'txf': ['txoL', 'spL', 'pcc'], 'txo': ['txoL', 'spL', 'pcc', 'pos'], 'rpb': ['vr'], 'vat': ['vr'],
               'etx': [], 'esub': [], 'esubpu': ['gj']}
PARAM_LABELS = {'ppos': 'price position', 'vr': 'VAT rate', 'ef': 'EF', 'fpos': 'fuel position', 'spos': 'sector position',
                'pfpos': 'price-fuel position', 'gj': 'GJ per unit', 'fbpos': 'feebate sector position',
                'eY': 'eps_Y', 'eU': 'eps_U', 'eF': 'eps_F', 'a': 'alpha', 'effx': 'efficiency exponent',
                'aeef': 'AEEI factor', 'ctc': 'existing carbon tax coverage', 'etc': 'existing ETS coverage'}

# Variable dictionary (sheet Variables): code, label, unit, code suffix, source, legacy reference, param labels
VARIABLES = [
    ('gdp.pos.pct', 'Real GDP growth', '%', '', 'WEO (lookup by country)', 'legacy Mitigation row 2446'),
    ('infl', 'Inflation index: US CPI, ResultsYear = 1 (nominal -> real for domestic prices and nominal policy '
             'inputs)', 'index', '', 'WEO USA|pcpi: CPI(ResultsYear) / CPI(year)',
     'legacy Mitigation rows 456 / 725 (inflation index, 2026 = 100)'),
    ('defl', 'Deflator index: US GDP deflator, ResultsYear = 1 (nominal -> real for international prices)', 'index',
     '', 'WEO USA|ngdp_d: deflator(ResultsYear) / deflator(year)', 'legacy Mitigation row 467 (row 630 = 1 / it)'),
] + [(p, '=MT', '=MT', '', 'MTInputs (scenario column)', f'MTInputs row {MT_NAMES[p]}')
     for p in CP_PARAMS + PRI_PARAMS + ETS_PARAMS + ECP_PARAMS] + [
    ('cptraj', 'Carbon price trajectory used', '$/tCO2 real', '', 'Calculation (legacy rows 2248, 2251); x infl if '
     'NomorReal = Nominal', 'legacy Mitigation row 2251: egy.mit.cptraj.1'),
    ('cp.t', 'Carbon price path: years since the start year (held at the target year unless the path continues '
             'linearly)', 'years', '', 'IF(Extend = Linear*, year, MIN(year, CPOutro)) - CPIntro', 'legacy row 2248'),
    ('cp.nom', 'Carbon price inputs: nominal-to-real factor', 'factor', '', 'CPI index if NomorReal = Nominal, else 1',
     'legacy row 2249'),
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
    ('ets.on', 'New ETS applied and started (1/0)', '0/1', '', 'D_NewETS = Yes and year >= D_ETSIntro', 'new'),
    ('ets.vol', 'ETS: assumed annual price volatility (Settings table, label D_ETSVolatility)', 'share', '',
     'Settings row 17', 'legacy ETS+LTS rows 5-8'),
    ('ets.imp', 'ETS: abatement cost increase per unit of volatility (Settings table, label D_ETSVolImpact)', 'factor',
     '', 'Settings row 18', 'legacy ETS+LTS rows 5-8'),
    ('ets.vadj', 'ETS volatility adjustment (tax-equivalent price per $ of permit price)', 'factor', '',
     '(1 + carbon tax policy risk x impact) / (1 + ETS volatility x impact)',
     'legacy row 2191 (1/1.1); ETS+LTS rows 5-9'),
    ('ets.chg', 'ETS cap: change relative to baseline covered emissions', 'share', '',
     'Linear from D_ETSChangeRelStart (start year) to D_ETSChangeRelTarget (target year), flat after',
     'legacy rows 1958-1960'),
    ('ets.bce', 'Baseline covered emissions', 'MtCO2', '', 'SUMPRODUCT(bco2, etsc)', 'legacy row 1956'),
    ('ets.cap', 'ETS cap', 'MtCO2', '', 'Baseline covered emissions x (1 + change path); constant after the target '
     'year if D_ETSCapCont = Constant', 'legacy rows 1958-1960'),
    ('ets.se', 'Effective covered semi-elasticity (allocation effectiveness: auctioned 1, OBR 0.5)', 'per $/tCO2',
     '', 'Baseline-weighted semi-elasticities x (1 - 0.5 OBR)', 'legacy rows 1845, 1866-1867'),
    ('ets.est', 'ETS price, fast estimate', '$/tCO2 real', '', 'LN(cap / baseline) / (semi-elasticity x volatility '
     'adjustment)', 'legacy row 1868 (QuickEstimateOfETSPrices; linear there)'),
    ('ets.ovr', 'ETS price override (pasted values, e.g. from ets_goalseek_v0_4.py)', '$/tCO2 real', '',
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
    ('ctr', 'New carbon tax rate by sector (not where the new ETS covers the sector)', 'USD/tCO2', '',
     'carbon price x sector coverage x (1 - ETS coverage); complementary to an existing ETS: sector coverage = 1 - '
     'existing ETS sector coverage', 'legacy Mitigation row 2510 (before EF and fuel coverage); rows 2270-2288',
     'existing ETS sector coverage'),
    ('xct.d', 'Existing carbon tax rate: data of the year (last data year after; real)', 'USD/tCO2 real', '',
     'Sheet ECP (ecr.all, nominal x CPI index)', 'legacy Mitigation row 1727 (ECP ecr.all)'),
    ('xct.p', 'Existing carbon tax rate used', 'USD/tCO2 real', '', 'ExistingCTApply x data x (1 + ExistingCTGrowth)^'
     '(years after the last data year)', 'legacy rows 1727, 2215-2216 (MTInputs 213-214)'),
    ('xct.p0', 'Existing carbon tax rate embedded in the price data (held after the last historical price year)',
     'USD/tCO2 real', '', 'Sheet ECP, year MIN(year, last data year, last historical price year)', 'new'),
    ('xets.d', 'Existing ETS permit price: data of the year (last data year after; real)', 'USD/tCO2 real', '',
     'Sheet ECP (eer.all, nominal x CPI index)', 'legacy Mitigation rows 1680, 1695 (ECP eer.all)'),
    ('xets.p', 'Existing ETS permit price used', 'USD/tCO2 real', '', 'ExistingETSApply x data x (1 + '
     'ExistingETSGrowth)^(years after the last data year)', 'legacy rows 1680, 1695 (MTInputs 208-209)'),
    ('xets.p0', 'Existing ETS permit price embedded in the price data (held after the last historical price year)',
     'USD/tCO2 real', '', 'Sheet ECP, year MIN(year, last data year, last historical price year)', 'new'),
    ('etr', 'New ETS cost in the fuel price by sector (tax-equivalent price, auctioned part)', 'USD/tCO2', '',
     'ETS tax-equivalent price x ETS coverage x (1 - OBR share)', 'legacy ets inside txo (before the EF)'),
    ('shpx', 'Shadow price on the efficiency margin by sector (feebates and regulations x share impacting efficiency + '
             'ETS tax-equivalent price x ETS coverage x OBR share)', 'USD/tCO2', '', 'Calculation',
     'legacy rows 2377 (per tCO2 before the EF)', 'feebate sector position'),
    ('gp.src', 'International prices: source key (MTInputs selector without the *)', 'key', '',
     'SUBSTITUTE(IntEnerPricForeSource, "*", "")', 'MTInputs row 220'),
    ('gp.adj', 'International prices: adjustment after the last historical year (1 none, 2 High, 3 Low)', 'choice', '',
     'From IntEnerPricForecastAdjustment', 'MTInputs row 258'),
    ('gpn', 'International energy price, nominal, of the chosen source', 'source unit', '',
     'Prices_int (1 without commodity key)', 'legacy Mitigation rows 633-695 (sources by year)', 'commodity key'),
    ('gp', 'International energy price (real; source and adjustment from MTInputs)', 'source unit', '',
     'gpn x defl x High/Low factor after the last historical year',
     'legacy Mitigation rows 633-695 (sources by year)', 'commodity key', 'High factor', 'Low factor'),
    ('sp', 'Supply cost (pre-tax price)', '$/GJ real', 'a', 'Historical: Inputs_prices; then fixsp + floating '
     'part x gp(t) / gp(t-1)', 'legacy Mitigation row 2498: egy.mit.sp.ind.coa.a.1; cpat_coded prices.py', 'fixed supply cost',
     'gp position', 'price position'),
    ('txo', 'Excise and other taxes (before new policies)', '$/GJ real', 'a',
     'Historical: Inputs_prices; then fixed part txo_L x pcc + floating part (pass-through rule)',
     'legacy txo = fadtx + fixs + cs (+ nce); cpat_coded prices.py', 'txo last hist. year', 'sp last hist. year',
     'pass-through', 'price position'),
    ('txf', 'Floating part of excise and other taxes (pass-through rule; used after the last historical year)',
     '$/GJ real', 'a', '(sp_L - sp + txo_L x (1 - pcc)) x (1 - pcc)', 'legacy cs; cpat_coded prices.py',
     'txo last hist. year', 'sp last hist. year', 'pass-through'),
    ('rpb', 'Retail price before new policies', '$/GJ real', 'a', '(sp + txo) x (1 + VAT rate)',
     'legacy retail price rp without nce; cpat_coded prices.py', 'VAT rate'),
    ('ets', 'New ETS cost in the fuel price (tax-equivalent, auctioned part)', '$/GJ', 'a',
     'ETS cost rate of the sector (etr, section 1) x EF',
     'legacy ets inside txo; cpat_coded prices/ets.py', 'EF', 'sector position'),
    ('vat', 'VAT payment before new policies', '$/GJ real', 'a', '(sp + txo) x VAT rate', 'legacy vat; cpat_coded '
     'prices.py', 'VAT rate'),
    ('etx', 'Existing excise and other taxes (positive part of txo)', '$/GJ real', 'a', 'max(txo, 0)',
     'legacy fadtx + positive cs (split here by sign)'),
    ('esub', 'Existing consumer subsidy (negative part of txo, positive number)', '$/GJ real', 'a', 'max(-txo, 0)',
     'legacy fixs + negative cs (split here by sign)'),
    ('esubpu', 'Existing consumer subsidy per price unit (input for a later per-fuel policy)', 'price unit real',
     'a', 'esub x GJ per price unit', 'new (units as MTInputs rows 142-169)', 'GJ per unit'),
    ('ctxnew', 'New carbon tax', '$/GJ', 'a', 'Carbon tax rate of the sector (ctr, section 1) x EF x fuel coverage',
     'legacy Mitigation row 2510: egy.mit.ctxnew.ind.coa.a.1', 'EF', 'fuel position', 'sector position'),
    ('ntx', 'New excise tax (fuel price reform)', '$/GJ', 'a', 'Fuel price reform path / GJ per unit',
     'legacy Mitigation row 2512: egy.mit.ntx.ind.coa.a.1', 'price-fuel position', 'GJ per unit'),
    ('nce', 'Total new policy', '$/GJ', 'a', 'New carbon tax + new excise tax',
     'legacy Mitigation row 2514: egy.mit.nce.ind.coa.a.1'),
    ('atp', 'After-tax price', '$/GJ real', 'e', 'max(rpb of the price fuel + nce x (1 + VAT rate), 0.01)',
     'legacy Mitigation row 5416: egy.mit.atp.rod.gso.e.1 (legacy unit $/liter for liquids)', 'price position',
     'VAT rate'),
    ('shp', 'Shadow price on the efficiency margin', '$/GJ', '', 'Shadow price of the sector (shpx, section 1) x EF',
     'legacy Mitigation row 2377: egy.mit.shp.coa.pow.1', 'EF', 'sector position'),
    ('xcp', 'Existing carbon pricing (carbon tax + ETS) per GJ', '$/GJ', 'a', 'EF x (carbon tax coverage x rate + '
     'ETS coverage x price); coverage = sector share x fuel share (sheet ECP)',
     'legacy rows 1711-1723 (ets), 1752-1764 (ctx)', 'EF', 'carbon tax coverage', 'ETS coverage'),
    ('xcpd', 'Change in existing carbon pricing against the price data (enters the after-tax price)', '$/GJ', 'a',
     'EF x (carbon tax coverage x (rate - embedded rate) + ETS coverage x (price - embedded price))',
     'new (legacy: ctx and ets are parts of txo)', 'EF', 'carbon tax coverage', 'ETS coverage'),
    ('ener', 'Fuel use', 'ktoe', 'e', 'Inputs (base year), CPAT eq. 3.3.3',
     'legacy Mitigation row 5430: egy.mit.ener.rod.gso.e.1', 'eps_Y', 'eps_U', 'efficiency exponent eps_F x (1 + '
     'eps_U)', 'AEEI factor (1/(1+alpha))^(1+eps_U)'),
    ('co2', 'CO2 emissions from fuel combustion', 'MtCO2', 'e', 'Fuel use x PJ/ktoe x EF (power generation, section '
     '3: generation fuel use (ktoe) x MtCO2 per ktoe)', 'legacy CO2 emissions by sector and fuel (inventory-adjusted '
     'EFs)', 'EF (power: MtCO2 per ktoe)'),
    ('co2.ref', 'CO2 emissions, scenario 1 (same year)', 'MtCO2', '', 'Lookup', 'new'),
    ('co2.chg', 'Change in CO2 emissions vs scenario 1', 'MtCO2', '', 'co2 - co2.ref', 'new'),
    ('co2.pct', 'Change in CO2 emissions vs scenario 1', '%', '', 'co2 / co2.ref - 1', 'new'),
    ('co2.sec', 'CO2 emissions by ETS sector', 'MtCO2', '', 'Sum over fuels (power: section 3, v1.06)', 'new'),
    ('co2.ets', 'CO2 emissions covered by the new ETS', 'MtCO2', '', 'SUMPRODUCT(co2.sec, etsc)', 'legacy row 1959'),
    ('co2.cap', 'ETS cap', 'MtCO2', '', 'ets.cap (section 1)', 'legacy row 1960'),
    ('co2.gap', 'Covered emissions vs cap (0 = cap met)', '%', '', 'co2.ets / co2.cap - 1', 'legacy rows 1961-1963'),
    ('ets.lr', 'ETS: log ratio LN(cap / baseline) / LN(covered / baseline) (0 if the cap does not bind)', 'ratio',
     '', 'Calculation', 'legacy PricesAdjustedByErrorFactor (row 1966)'),
    ('ets.next', 'ETS price proposal for the next iteration', '$/tCO2 real', '', 'p x log ratio ^ convergence '
     'exponent (Settings); fast estimate if no price or no log ratio',
     'legacy PricesAdjustedByErrorFactor (row 1966)'),
    ('eldem', 'Electricity demand (memo: power model H; not in the fuel totals; sector policies do not act on it)',
     'ktoe', 'e', 'Section 3 H, demand by subsector', 'legacy H3 (rows 3725-3739)'),
    ('bco2.chk', 'Check: baseline CO2 data minus scenario 1 CO2 (0 = data up to date)', 'MtCO2', '',
     'SUM(bco2) - co2.ref', 'new'),
    ('rtx', 'Revenue from existing taxes (excise and other taxes + VAT; incl. the change in existing carbon pricing)',
     'USD m real', '', 'Fuel use x PJ/ktoe x (etx + vat + xcpd x (1 + VAT rate)); power: fuel taxes, change in '
     'existing carbon pricing, electricity VAT', 'legacy revenue block (existing taxes)', 'price position',
     'VAT rate'),
    ('rxcp', 'Revenue from existing carbon pricing (memo: level; the part embedded in the price data is in rtx)',
     'USD m real', '', 'Fuel use x PJ/ktoe x xcp (power: CO2 x coverage x price)', 'legacy ECP revenue (crv, erv; '
     'not used)'),
    ('rsub', 'Cost of existing consumer subsidies (positive = cost)', 'USD m real', '', 'Fuel use x PJ/ktoe x esub',
     'legacy subsidy cost', 'price position'),
    ('rnew', 'Revenue from new policies (carbon tax, excise, auctioned ETS, VAT on them)', 'USD m real', '',
     'Fuel use x PJ/ktoe x (ctxnew + ntx + ets x revenue factor + nce x VAT rate)', 'legacy new revenue',
     'VAT rate'),
    ('rnet', 'Net revenue from existing taxes and subsidies', 'USD m real', '', 'rtx - rsub', 'new'),
    ('rtot', 'Total revenue (net existing + new policies)', 'USD m real', '', 'rnet + rnew', 'new'),
    ('rtot.ref', 'Total revenue, scenario 1 (same year)', 'USD m real', '', 'Lookup', 'new'),
    ('rev.chk', 'Check: revenue totals minus their sector and power rows (rtx, rsub, rnew; should be 0)',
     'USD m real', '', 'Check', 'new'),
    ('co2.chk', 'Check: CO2 total minus sectors and power, and minus the fuels (should be 0)', 'MtCO2', '', 'Check',
     'new'),
    ('rtot.chg', 'Change in total revenue vs scenario 1 (fiscal effect)', 'USD m real', '', 'rtot - rtot.ref',
     'new'),
    ('ener.chk', 'Check: totals by subsector (and power) and by fuel minus total fuel use (should be 0)', 'ktoe', '',
     'Check', 'new'),
    ('ener.ref', 'Fuel use, scenario 1 (same year)', 'ktoe', 'e', 'Lookup', 'new'),
    ('ener.pct', 'Change in fuel use vs scenario 1', '%', 'e', 'Calculation', 'new'),
    # 3. Power sector (v1.05 prices, v1.06 engineer model)
    ('tcf', 'Capex time factor (capital cost relative to its level at factor 1)', 'factor', '', 'PowerPaths',
     'legacy Mitigation rows 3038-3046 (A7)', 'data key'),
    ('cbat', 'Battery storage capital cost', '$/kWh storage', '', 'PowerPaths', 'legacy row 3887 (J2)', 'data key'),
    ('cint', 'Battery interface capital cost', '$/kW', '', 'PowerPaths', 'legacy row 3888 (J2)', 'data key'),
    ('obat', 'Battery storage operating cost', '$/kWh storage/yr', '', 'PowerPaths', 'legacy row 3891 (J2)',
     'data key'),
    ('cel', 'Long-term storage cost per marginal unit (electrolyser and storage capital, fixed and variable opex)',
     '$/kWh', '', 'PowerPaths', 'legacy rows 3900-3902 (J2)', 'data key'),
    ('incf', 'Income-dependent factor on the GDP elasticities of power demand', 'factor', '', 'PowerPaths',
     'legacy row 2433 (macro block)', 'data key'),
    ('shock', 'Exogenous shock on energy demand (legacy calibration 2023-2024; used if Settings C19 = 1)', 'share', '',
     'PowerPaths', 'legacy row 2439', 'data key'),
    ('cax', 'Capital cost of new plants', '$/kW real', '', 'cax0 x capex time factor', 'legacy rows 3099-3107 (A13)',
     'cax0 (factor 1)'),
    ('caxav', 'Weighted average capital cost of the stock (vintages)', '$/kW real', '',
     'Base year cax; then caxav(t-1) x (1 - phi(t-1)) + cax(t-1) x phi(t-1)', 'legacy rows 3181-3189 (B2)'),
    ('stoav', 'Weighted average storage cost of the stock (vintages)', '$/kWh', '',
     'Base year msc x VRE flag; then stoav(t-1) x (1 - phi(t-1)) + msc(t) x VRE flag x phi(t-1)',
     'legacy rows 3225-3233 (B6)', 'VRE flag'),
    ('fix', 'Amortised fixed cost of the existing stock (capital, interest, decommissioning, storage, fixed O&M)',
     '$/kWh', '', 'caxav x (1/N + wacc/2)/(cf x 8,760) + decommissioning/N/(cf x 8,760) + fixed O&M + stoav',
     'legacy rows 3247-3255 (B8 = B3 + B4 + B5 + B6 + B7)', 'capex factor', 'other fixed cost'),
    ('lfx', 'Levelised fixed cost of new plants', '$/kWh', '',
     '(cax + decommissioning + transmission capex)/(cf x 8,760 x dlf) + fixed O&M + subsidy + msc x VRE flag',
     'legacy rows 3326-3334 (C7 = C1 + ... + C6)', 'capex factor', 'other fixed cost', 'VRE flag'),
    ('xck', 'Change in existing carbon pricing against the price data, per kWh generated', '$/kWh', '',
     'tCO2/kWh x (carbon tax coverage x (rate - embedded) + ETS coverage x (price - embedded))', 'new (legacy: '
     'existing pricing is part of txo of power fuels)', 'tCO2/kWh', 'carbon tax coverage', 'ETS coverage'),
    ('vbc', 'Variable cost before new policies (fuel and variable O&M)', '$/kWh', '',
     'Variable O&M + fixed fuel cost + rpb of the price fuel (section 2, $/GJ) x 0.0036/efficiency + xck',
     'legacy rows 3374-3382 (D4; legacy uses a 5-year moving average of fuel prices)', 'var. O&M + fixed fuel',
     'price position', 'GJ per kWh'),
    ('ccp', 'New policy cost per kWh generated (carbon tax, ETS, fuel price reform)', '$/kWh', '',
     'tCO2/kWh x (carbon tax rate of power x fuel coverage + ETS cost rate of power; section 1 ctr, etr) + fuel price '
     'reform x 0.0036/efficiency/GJ per unit',
     'legacy rows 3385-3393 (D5)', 'tCO2/kWh', 'fuel position', 'price-fuel position', 'kWh factor of reform'),
    ('shk', 'Shadow price per kWh generated (power feebates, ETS output-based part)', '$/kWh', '',
     'shadow price of power (shpx, section 1: feebate rate x share impacting efficiency + ETS tax-equivalent price x '
     'ETS coverage x OBR share) x tCO2/kWh',
     'legacy rows 3398-3406 (D6)', 'tCO2/kWh', 'base-year generation share (rebate)'),
    ('reb', 'Feebate rebate per kWh generated (fixed base-year shares; from the second year)', '$/kWh', '',
     '-SUMPRODUCT(shk, base-year generation shares)', 'legacy rows 3407-3409 (fixed: per kWh generated, legacy per '
     'kWh consumed)'),
    ('tvc', 'Total variable cost (dispatch; levelised for investment)', '$/kWh', '', 'vbc + ccp + shk + reb',
     'legacy rows 3423-3431 (D7)'),
    ('dsum', 'Sum of discount factors from the year to the end of the horizon', 'years', '',
     '(1 + wacc)^(base year - year) + next year (blank after the last year)', 'legacy A5 discount factors (row 3012)',
     'WACC'),
    ('lvc', 'Levelised variable cost: discounted average of tvc from the year to the end of the horizon', '$/kWh', '',
     '(df x tvc + (dsum - df) x lvc of next year) / dsum', 'legacy rows 3439-3447 (D9; capacity factors constant)',
     'WACC'),
    ('gnc', 'Generation cost before new policies (cost recovery)', '$/kWh', '', 'fix + vbc',
     'legacy rows 3486-3494 (E3)'),
    ('gncav', 'Average generation cost (weighted by generation shares)', '$/kWh', '', 'SUMPRODUCT(shr, gnc)',
     'legacy row 3495 (E3 weighted average)', 'factor'),
    ('ccpav', 'Average new policy cost per kWh consumed', '$/kWh', '',
     'SUMPRODUCT(shr, ccp) x generation / consumption', 'legacy row 3395 (D5 per kWh consumed)',
     'generation / consumption'),
    ('tic', 'Levelised investment cost (levelised fixed + levelised variable)', '$/kWh', '', 'MAX(lfx + lvc, minimum)',
     'legacy rows 3463-3471 (E1) = L1', 'minimum cost'),
    ('cons', 'Electricity consumption by price group', 'GWh', '', 'residential: residential demand; industry: total '
     'minus residential; x GWh per ktoe', 'legacy rows 3621, 3650 (G1, G2)', 'weight on total', 'weight on residential',
     'GWh per ktoe'),
    ('sc', 'Supply cost of electricity (average generation cost of the previous year + transmission and '
           'distribution)', '$/kWh', '', 'gncav(t-1) + T&D add-on; base year gncav(t) + T&D',
     'legacy rows 3600, 3629', 'T&D add-on'),
    ('rppre', 'Electricity price before new policies (incl. VAT)', '$/kWh real', '',
     'Data up to the last historical year (x CPI index), then rppre(t-1) + pass-through x (sc - sc(t-1)) x (1 + VAT)',
     'legacy rows 3613, 3642 (telescoped: held components cancel)', 'pass-through', 'user-group position',
     'VAT rate'),
    ('rp', 'End-user electricity price', '$/kWh real', '', 'rppre + ccpav x pass-through x (1 + VAT) (rebate and '
     'power excise 0)', 'legacy rows 3614, 3617, 3646', 'pass-through', 'VAT rate'),
    ('sgap', 'Subsidy gap: supply cost minus price before VAT (positive = price below cost)', '$/kWh', '',
     'sc - rppre / (1 + VAT rate)', 'new (legacy producer-subsidy and over/under-estimate rows)', 'VAT rate'),
    ('rvat', 'VAT revenue from electricity', 'USD m real', '', 'rppre x VAT/(1 + VAT) x consumption', 'new',
     'VAT share of price'),
    ('rcarb', 'New-policy cost per user group (carbon tax, ETS at the tax-equivalent price, fuel price reform; '
              'memo, revenue: prnew in M)', 'USD m real', '', 'ccpav x consumption', 'legacy power carbon revenue',
     'factor'),
    ('rnvat', 'VAT on the new-policy cost passed through to electricity users', 'USD m real', '',
     'ccpav x pass-through x VAT rate x consumption', 'new', 'pass-through', 'VAT rate'),
    ('rgap', 'Fiscal cost of the electricity subsidy gap', 'USD m real', '', 'sgap x consumption', 'new', 'factor'),
    ('pdem', 'Electricity price of the previous year used by demand (by demand group)', '$/kWh real', '',
     'weight x residential price(t-1) + (1 - weight) x industrial price(t-1); base year: 2021 prices',
     'legacy rows 3687-3689 (deliberately lagged)', 'weight on the residential price', '2021 price'),
    ('mult', 'Growth multiple of electricity demand (efficiency, GDP, price; demand shock if switched on)', 'factor',
     '', 'efficiency multiple x (1 + GDP growth)^(eps_Y x income factor) x (p/p(t-1))^(total price elasticity) x '
     '(1 + shock x switch)', 'legacy rows 3691-3713 (fixed: efficiency exponent, income factor)', 'efficiency multiple',
     'eps_Y', 'total price elasticity'),
    ('dem', 'Electricity demand', 'ktoe', '', 'demand(t-1) x multiple of the demand group',
     'legacy rows 3725-3739 (H3; fixed: services keep the services multiple)', 'demand-group position',
     'base-year demand'),
    ('demg', 'Electricity demand by demand group (memo)', 'ktoe', '', 'Sum over the subsectors of the group',
     'legacy H3 (groups of H1)', 'demand-group position'),
    ('demtot', 'Electricity demand, total', 'ktoe', '', 'Sum over subsectors', 'legacy row 3741'),
    ('req', 'Required generation (incl. net exports, transmission and statistical differences, own use)', 'MWy', '',
     'Base year: actual generation; then total demand x required-generation factor', 'legacy rows 3746-3752',
     'MWy per ktoe of demand'),
    ('reqcg', 'Generation required from coal and gas', 'MWy', '', 'MAX(required generation - fixed-output generation, '
     '0)', 'legacy row 3779'),
    ('vcmin', 'Lowest variable cost of coal and gas', '$/kWh', '', 'MAX(MIN over coal and gas of tvc, minimum)',
     'legacy rows 3790-3793', 'minimum cost'),
    ('rem', 'Coal and gas generation left after the minima', 'MWy', '', 'MAX(reqcg - sum of minima, 0)',
     'legacy row 3789'),
    ('gmin', 'Minimum generation (coal if gas runs at its maximum, and vice versa)', 'MWy', '',
     'coal/gas flag x MIN(MAX(reqcg - maximum of the other, 0), own maximum)', 'legacy rows 3787-3788',
     'coal/gas flag', 'maximum capacity factor'),
    ('lgd', 'Dispatch logit weight (coal and gas)', 'weight', '', 'flag x EXP(-K x MAX(tvc, minimum) / lowest)',
     'legacy rows 3795-3796', 'coal/gas flag', 'K dispatch', 'minimum cost'),
    ('gen', 'Generation', 'MWy', '', 'Base year: data; coal and gas: MIN(maximum, minimum + rem x logit share); '
     'others: effective capacity', 'legacy rows 3804-3812 (I3; PPAs and scale-down not used)', 'coal/gas flag',
     'maximum capacity factor', 'base-year generation'),
    ('gtot', 'Generation, total', 'MWy', '', 'Sum over types', 'legacy row 3813'),
    ('dgen', 'Increase in total generation (0 in the base year)', 'MWy', '', 'MAX(0, gtot - gtot(t-1))',
     'legacy row 3972 (used lagged)'),
    ('shr', 'Generation share', 'share', '', 'gen / gtot', 'legacy rows 3840-3848 (I6) = E0', 'counts in VRE share'),
    ('pe', 'Fuel use (primary energy) of power generation', 'ktoe', '', 'gen x 8.76 / GWh per ktoe / efficiency',
     'legacy rows 3855-3863 (I7)', 'ktoe per MWy'),
    ('vre', 'Variable renewable share of generation (wind, solar, other renewables)', 'share', '',
     'SUMPRODUCT(shr, VRE-share flag)', 'legacy rows 3851, 3907', 'factor'),
    ('cph', 'Short-term storage cost per marginal hour of storage', '$/kWh', '',
     '(battery capex + interface capex / kWh per kW)/(dlf of renewables x 8,760) + battery opex / 8,760',
     'legacy row 3905 (J2)', 'kWh storage per kW', 'dlf of renewables'),
    ('msc', 'Marginal storage cost of variable renewables (wind, solar, other renewables)', '$/kWh', '',
     'alloc x 18 x vre x cph + 2 x MAX(0, vre - 0.75)/(1 - 0.75)^2 x cel x LT allocation',
     'legacy rows 3908-3912 (J1; fixed: LT allocation once)', 'short-term allocation', 'short-term hours factor',
     'long-term start share', 'long-term allocation'),
    ('cap', 'Nameplate capacity', 'MW', '', 'Base year: data; then cap(t-1) - retirement(t-1) + new nameplate '
     'investment(t-1)', 'legacy rows 3932-3940 (K1) = L21 of the previous year', 'base-year capacity'),
    ('captot', 'Nameplate capacity, total', 'MW', '', 'Sum over types', 'legacy row 3941'),
    ('pret', 'Planned retirement (data; used where the type\'s switch is Yes)', 'MW', '', 'PowerPaths pret|<type> '
     '(Egypt: no data, 0)', 'legacy row 3992 (K5, coal only; data from sheet Power)', 'data key'),
    ('effcap', 'Effective capacity (expected generation)', 'MW', '', 'cap x capacity factor',
     'legacy rows 3956-3964 (K3)', 'capacity factor'),
    ('ret', 'Retirement of nameplate capacity', 'MW', '', 'MIN(planned retirement if switched on, else cap x '
     'retirement rate; + early coal retirement; cap)', 'legacy rows 3992-4002 (K5; planned retirement for every type, '
     'legacy coal only)', 'retirement rate', 'early-retirement flag', 'planned retirement used (1/0)'),
    ('reteff', 'Retirement of effective capacity', 'MW', '', 'ret x capacity factor',
     'legacy rows 4009-4019 (K6; fixed: coal as nameplate)', 'capacity factor'),
    ('egrow', 'Room for early coal retirement: maximum wind and solar additions less last year\'s demand increase and '
              'non-coal retirements', 'MW', '', 'MAX(0, ...)', 'legacy rows 3970-3974 (fixed: effective retirements)'),
    ('ecmin', 'Early coal retirement: lowest of coal variable cost and wind and solar investment cost', '$/kWh', '',
     'MIN(tvc coal, tic wind, tic solar)', 'legacy row 3982'),
    ('ecshr', 'Early coal retirement: coal share of the logit (coal variable cost vs wind and solar investment '
              'cost)', 'share', '', '1 / (1 + EXP(-K (tic wind - tvc coal) / lowest) + EXP(-K (tic solar - tvc coal) '
     '/ lowest))', 'legacy row 3983', 'K investment'),
    ('early', 'Cost-based early coal retirement (nameplate)', 'MW', '',
     'MAX(0, MIN(coal effective capacity x 0.8, room) x (1 - coal share of the logit)) / cf x 0.8',
     'legacy rows 3975-3993 (K4)', 'maximum share', 'factor', 'coal capacity factor'),
    ('kafter', 'Effective capacity after this year\'s retirements', 'MW', '', 'SUM(effcap) - SUM(reteff)',
     'legacy row 4026 (fixed: actual capacity, legacy last year\'s plan)'),
    ('need', 'Effective capacity investment needed', 'MW', '', 'MAX(required generation - kafter, 0)',
     'legacy row 4028 (capacity for this year\'s demand, available next year)'),
    ('cmin', 'Lowest levelised investment cost of the permitted types', '$/kWh', '', 'MAX(MIN over permitted, minimum)',
     'legacy row 4046 (fixed: permitted types only)', 'minimum cost'),
    ('deficit', 'Investment left for the second round', 'MW', '', 'MAX(need - SUM(first round), 0)',
     'legacy row 4153'),
    ('permit', 'New investment permitted (0/1)', '0/1', '', 'Yes from the start year', 'legacy rows 4061-4069 (L3)',
     'permitted (Yes/No)', 'from year'),
    ('maxinv', 'Maximum effective investment (scale-up limit)', 'MW', '', 'rate x total effective capacity x permit',
     'legacy rows 4130-4138 (L8)', 'scale-up rate'),
    ('lgi', 'Investment logit weight', 'weight', '', 'permit x EXP(-K x tic / cmin)', 'legacy rows 4073-4081 (L4)',
     'K investment'),
    ('pni', 'Share of new investment (logit)', 'share', '', 'lgi / SUM(lgi)', 'legacy rows 4085-4093 (L5)'),
    ('first', 'New effective investment, first round (logit within scale-up limits)', 'MW', '',
     'MIN(maxinv, need x pni)', 'legacy rows 4143-4151 (L9)'),
    ('space', 'Remaining space under the scale-up limit', 'MW', '', '(maxinv - first) x permit',
     'legacy rows 4158-4166 (L10)'),
    ('second', 'New effective investment, second round (cheapest remaining space first)', 'MW', '',
     'permit x MAX(0, MIN(space, deficit - space of cheaper types))', 'legacy rows 4169-4246 (L11-L17 in one row)',
     'tie breaker'),
    ('nei', 'New effective investment', 'MW', '', 'first + second', 'legacy rows 4251-4259 (L18)'),
    ('nim', 'New nameplate investment', 'MW', '', 'nei / capacity factor', 'legacy rows 4264-4272 (L19)',
     'nameplate per effective MW'),
    ('nimtot', 'New nameplate investment, total', 'MW', '', 'Sum over types', 'legacy row 4273'),
    ('usd', 'New investment', 'USD m real', '', 'nim x cax / 1000', 'legacy rows 4276-4284 (L20; fixed: capital cost '
     'of the year)'),
    ('usdtot', 'New investment, total', 'USD m real', '', 'Sum over types', 'legacy row 4285'),
    ('phi', 'New investment as a share of the stock after retirements', 'share', '', 'nim / (cap + nim - ret)',
     'legacy rows 3169-3177 (B1)'),
    # M. power-sector revenues (v1.09)
    ('prtx', 'Existing taxes on generation fuels (excise and other taxes + VAT of the price fuel)', 'USD m real', '',
     'fuel flag x fuel use (ktoe) x PJ/ktoe x (etx + vat)', 'new', 'price position', 'fuel flag'),
    ('prsub', 'Existing subsidies on generation fuels (positive = cost)', 'USD m real', '',
     'fuel flag x fuel use x PJ/ktoe x esub', 'new', 'price position', 'fuel flag'),
    ('prxc', 'Change in existing carbon pricing revenue against the price data', 'USD m real', '',
     'xck x generation (MWy x 8.76 = GWh)', 'new'),
    ('prnew', 'New-policy revenue (carbon tax, ETS at the permit price, fuel price reform)', 'USD m real', '',
     'ccp x GWh - CO2 x ETS cost rate of power x (1 - revenue factor)', 'legacy power carbon revenue (fixed: ETS at '
     'the permit price)'),
    ('prxl', 'Existing carbon pricing revenue (memo: level)', 'USD m real', '', 'CO2 x (carbon tax coverage x rate + '
     'ETS coverage x price)', 'legacy ECP revenue', 'carbon tax coverage', 'ETS coverage'),
]
VR0 = 5
VR1 = VR0 + 260

def IP_HIST(v):
    """History lookup inside the SUPPLYCOST / OTHERTAX LAMBDAs: Inputs_prices value of year, price row pos."""
    c0, c1 = IPC[v + str(HIST_YEARS[0])], IPC[v + str(HIST_YEARS[-1])]
    return (f'IFERROR(INDEX(Inputs_prices!${c0}${IP_R0}:${c1}${IP_R1},pos,MATCH(year,Inputs_prices!${c0}${IP_HDR}:'
            f'${c1}${IP_HDR},0)),0)')


LAMBDAS = {   # every calculated row calls one of these in the right column (LAST_YEAR)
    'PATH': (['year', 'y0', 'y1', 'start', 'target', 'cont'],
             'IF(year<y0,0,start+(target-start)/MAX(y1-y0,1)*IF(OR(year<=y1,cont),year-y0,y1-y0))'),
    'PATHYEARS': (['year', 'y0', 'y1', 'cont'], 'IF(cont,year,MIN(year,y1))-y0'),
    'NOMFACTOR': (['basis', 'cpi_index'], 'IF(LEFT(basis,7)="Nominal",cpi_index,1)'),
    'ETSON': (['apply', 'year', 'start_year'], 'IF(AND(LEFT(apply,3)="Yes",year>=start_year),1,0)'),
    'ETSPRICE': (['ets_on', 'override', 'override_price', 'estimate', 'carbon_price'],
                 'IF(ets_on=1,IF(LEFT(override,3)="Yes",IF(override_price="",carbon_price,override_price),estimate),0)'),
    'OBRSHARE': (['benchmark'], 'MIN(1,MAX(0,benchmark))'),
    'VOLPICK': (['label', 'values'], 'INDEX(values,MATCH(SUBSTITUTE(label,"*",""),Settings!$C$16:$F$16,0))'),
    'VOLADJ': (['volatility', 'impact', 'ct_risk'], '(1+ct_risk*impact)/(1+volatility*impact)'),
    'COVERED': (['emissions', 'coverage'], 'SUMPRODUCT(emissions,coverage)'),
    'CAPCHANGE': (['year', 'y0', 'y1', 'change_start', 'change_target'],
                  'change_start+(change_target-change_start)/MAX(y1-y0,1)*(MIN(year,y1)-y0)'),
    'ETSCAP': (['ets_on', 'year', 'y1', 'cont', 'base', 'change', 'cap_prev'],
               'IF(ets_on=1,IF(AND(year>y1,LEFT(cont,8)="Constant"),cap_prev,base*(1+change)),0)'),
    'ETSSEMI': (['emissions', 'coverage', 'semi', 'obr', 'obr_eff', 'base'],
                'IF(base>0,SUMPRODUCT(emissions,coverage,semi,1-(1-obr_eff)*obr)/base,0)'),
    'ETSESTIMATE': (['ets_on', 'cap', 'base', 'semi', 'vadj'],
                    'IF(AND(ets_on=1,cap>0,cap<base,semi<0,vadj>0),LN(cap/base)/(semi*vadj),0)'),
    'TAXEQUIV': (['price', 'vadj'], 'price*vadj'),
    'REVFACTOR': (['vadj'], 'IF(vadj>0,1/vadj,0)'),
    'ETSCOVER': (['ets_on', 'covered'], 'IF(ets_on=1,IF(covered,1,0),0)'),
    'SHADOWSECTOR': (['feebate', 'regulation'], 'feebate+regulation'),
    'SHADOWSHARE': (['coverage', 'adjustment'], 'IF(coverage,1,0)*adjustment'),
    'SUPPLYCOST': (['year', 'pos', 'fix_sp', 'sp_prev', 'gp_now', 'gp_prev'],
                   f'IF(year<=Settings!$C$10,{IP_HIST("sp")},fix_sp+(sp_prev-fix_sp)*gp_now/gp_prev)'),
    'OTHERTAX': (['year', 'pos', 'txo_last', 'sp_last', 'sp_now', 'pcc'],
                 f'IF(year<=Settings!$C$10,{IP_HIST("txo")},txo_last*pcc+IF(txo_last*(1-pcc)>=0,MAX(txo_last*(1-pcc),'
                 '(sp_last-sp_now+txo_last*(1-pcc))*(1-pcc)),(sp_last-sp_now+txo_last*(1-pcc))*(1-pcc)))'),
    'RETAILPRICE': (['sp', 'txo', 'vat_rate'], '(sp+txo)*(1+vat_rate)'),
    'VATPAY': (['sp', 'txo', 'vat_rate'], '(sp+txo)*vat_rate'),
    'TAXPART': (['txo'], 'MAX(txo,0)'),
    'SUBSIDYPART': (['txo'], 'MAX(-txo,0)'),
    'PERUNIT': (['value_gj', 'gj_per_unit'], 'value_gj*gj_per_unit'),
    'SECTORTAX': (['price', 'sector_cov', 'ets_cov', 'complement', 'xets_cov'],
                  'price*IF(LEFT(complement,3)="Yes",1-xets_cov,sector_cov)*(1-ets_cov)'),
    'ECPPRICE': (['apply', 'data', 'growth', 'year'], 'IF(LEFT(apply,3)="Yes",data*(1+growth)^MAX(0,year-ECPLast),0)'),
    'XCARBON': (['ef', 'ct_cov', 'ct_price', 'ets_cov', 'ets_price'], 'ef*(ct_cov*ct_price+ets_cov*ets_price)'),
    'SECTORETS': (['price', 'ets_cov', 'obr'], 'price*ets_cov*(1-obr)'),
    'CARBONTAX': (['rate', 'ef', 'fuel_cov'], 'rate*ef*fuel_cov'),
    'ETSCOST': (['rate', 'ef'], 'rate*ef'),
    'NEWEXCISE': (['increase', 'gj_per_unit'], 'increase/gj_per_unit'),
    'NEWPOLICY': (['carbon_tax', 'ets', 'excise'], 'carbon_tax+ets+excise'),
    'POSTTAX': (['pre_policy_p', 'new_policy', 'vat_rate'], 'MAX(pre_policy_p+new_policy*(1+vat_rate),0.01)'),
    'SHADOWTOTAL': (['sector_price', 'share', 'ets_price', 'ets_cov', 'obr'], 'sector_price*share+ets_price*ets_cov*obr'),
    'SHADOWGJ': (['shadow_price', 'ef'], 'shadow_price*ef'),
    'KWHPOLICY': (['tco2_kwh', 'ct_rate', 'fuel_cov', 'ets_rate', 'reform', 'reform_factor'],
                  'tco2_kwh*(ct_rate*fuel_cov+ets_rate)+reform*reform_factor'),
    'FLOATTAX': (['txo_last', 'sp_last', 'sp_now', 'pcc'], '(sp_last-sp_now+txo_last*(1-pcc))*(1-pcc)'),
    'REVENUE': (['fuel_use', 'pj_per_ktoe', 'rate'], 'fuel_use*pj_per_ktoe*rate'),
    'EMISSIONS': (['fuel_use', 'pj_per_ktoe', 'ef'], 'fuel_use*pj_per_ktoe*ef'),
    'NEWREVRATE': (['carbon_tax', 'excise', 'ets', 'rev_factor', 'new_policy', 'vat_rate'],
                   'carbon_tax+excise+ets*rev_factor+new_policy*vat_rate'),
    'FUELUSE': (['f_prev', 'p_now', 'p_prev', 'shp_now', 'shp_prev', 'gdp_g', 'eps_y', 'eps_u', 'eff_exp',
                 'aeei_factor'],
                'f_prev*aeei_factor*(1+gdp_g)^eps_y*(p_now/p_prev)^eps_u*((p_now+shp_now)/(p_prev+shp_prev))^eff_exp'),
    # 3. Power sector (v1.05)
    'CAPEX': (['capex_base', 'time_factor'], 'capex_base*time_factor'),
    'VINTAGE': (['avg_prev', 'new_value', 'share_prev'], 'avg_prev*(1-share_prev)+new_value*share_prev'),
    'AMORTISED': (['capex_avg', 'capex_factor', 'other_fixed', 'storage_avg'],
                  'capex_avg*capex_factor+other_fixed+storage_avg'),
    'LEVELISED': (['capex', 'capex_factor', 'other_fixed', 'vre_flag', 'storage'],
                  'capex*capex_factor+other_fixed+vre_flag*storage'),
    'VARCOST': (['fixed_part', 'fuel_price', 'gj_per_kwh', 'xcarbon'], 'fixed_part+fuel_price*gj_per_kwh+xcarbon'),
    'PNEWREV': (['gen', 'policy_cost', 'co2', 'ets_rate', 'rev_factor'],
                'gen*8.76*policy_cost-co2*ets_rate*(1-rev_factor)'),
    'GENCOST': (['fixed_cost', 'variable_cost'], 'fixed_cost+variable_cost'),
    'WEIGHTED': (['shares', 'values', 'factor'], 'SUMPRODUCT(shares,values)*factor'),
    'STORAGEHOUR': (['battery', 'interface', 'opex', 'kwh_per_kw', 'dlf'],
                    '(battery+interface/kwh_per_kw)/(dlf*8760)+opex/8760'),
    'STORAGECOST': (['alloc', 'hours', 'vre', 'cph', 'lt_start', 'lt_cost', 'lt_alloc'],
                    'alloc*hours*vre*cph+2*MAX(0,vre-lt_start)/(1-lt_start)^2*lt_cost*lt_alloc'),
    'POWERSUPPLY': (['gen_cost', 'tmc'], 'gen_cost+tmc'),
    'POWERPRICE': (['year', 'pos', 'price_prev', 'pass', 'sc_now', 'sc_prev', 'vat_rate'],
                   'IF(year<=Settings!$C$10,HIST,price_prev+pass*(sc_now-sc_prev)*(1+vat_rate))'),   # HIST: below
    'ENDUSERPOWER': (['pre_policy_p', 'policy_cost', 'pass', 'vat_rate'], 'pre_policy_p+policy_cost*pass*(1+vat_rate)'),
    'SUBSIDYGAP': (['supply_cost', 'pre_policy_p', 'vat_rate'], 'supply_cost-pre_policy_p/(1+vat_rate)'),
    'POWERREV': (['unit_value', 'factor', 'consumption'], 'unit_value*factor*consumption'),
    # 3. Power sector, engineer model (v1.06)
    'REBATE': (['fees', 'shares'], '-SUMPRODUCT(fees,shares)'),
    'VARTOTAL': (['base_cost', 'policy_cost', 'shadow', 'rebate'], 'base_cost+policy_cost+shadow+rebate'),
    'INVCOST': (['levelised_fixed', 'levelised_var', 'minimum'], 'MAX(levelised_fixed+levelised_var,minimum)'),
    'CONSUMPTION': (['total', 'res', 'w_total', 'w_res', 'gwh_per_ktoe'], '(w_total*total+w_res*res)*gwh_per_ktoe'),
    'PRICEMIX': (['w_res', 'p_res', 'p_ind'], 'w_res*p_res+(1-w_res)*p_ind'),
    'DEMANDMULT': (['eff_mult', 'gdp_g', 'eps_y', 'income', 'p_now', 'p_prev', 'eps_p', 'shock', 'shock_on'],
                   'eff_mult*(1+gdp_g)^(eps_y*income)*(p_now/p_prev)^eps_p*(1+shock*shock_on)'),
    'GROW': (['prev', 'multiple'], 'prev*multiple'),
    'SCALE': (['value', 'factor'], 'value*factor'),
    'RESIDUAL': (['total', 'part'], 'MAX(total-part,0)'),
    'FLOORMIN': (['values', 'flags', 'min_value'], 'MAX(_xlfn.MINIFS(values,flags,1),min_value)'),
    'MINDISPATCH': (['flag', 'required', 'max_all', 'max_own'], 'flag*MIN(MAX(required-(max_all-max_own),0),max_own)'),
    'LOGITW': (['flag', 'k', 'cost', 'lowest'], 'flag*EXP(-k*cost/lowest)'),
    'DISPATCH': (['flag', 'fixed_output', 'max_gen', 'min_gen', 'remainder', 'weight', 'weight_sum'],
                 'IF(flag=1,MIN(max_gen,min_gen+remainder*weight/weight_sum),fixed_output)'),
    'INCREASE': (['value_now', 'value_prev'], 'MAX(0,value_now-value_prev)'),
    'SHARE': (['part', 'total'], 'IF(total=0,0,part/total)'),
    'CARRY': (['cap_prev', 'ret_prev', 'inv_prev'], 'cap_prev-ret_prev+inv_prev'),
    'RETIRE': (['cap', 'ret_rate', 'early_flag', 'early', 'planned_on', 'planned'],
               'MIN(IF(planned_on=1,planned,cap*ret_rate)+early_flag*early,cap)'),
    'EARLYROOM': (['add_wind', 'add_solar', 'demand_increase', 'retirements'],
                  'MAX(0,add_wind+add_solar-demand_increase-(retirements))'),
    'LOWEST3': (['a', 'b', 'c'], 'MIN(a,b,c)'),
    'COALSHARE': (['k', 'c_coal', 'c_wind', 'c_solar', 'lowest'],
                  '1/(1+EXP(-k*(c_wind-c_coal)/lowest)+EXP(-k*(c_solar-c_coal)/lowest))'),
    'EARLYRET': (['coal_eff', 'max_share', 'room', 'coal_share', 'factor', 'cf_coal'],
                 'MAX(0,MIN(coal_eff*max_share,room)*(1-coal_share))/cf_coal*factor'),
    'DIFF': (['minuend', 'subtrahend'], 'minuend-subtrahend'),
    'PERMITTED': (['outcome', 'start', 'year'], 'IF(AND(outcome="Yes",year>=start),1,0)'),
    'SCALEUP': (['up_rate', 'capacity', 'permit'], 'up_rate*capacity*permit'),
    'FIRSTROUND': (['max_inv', 'need', 'share'], 'MIN(max_inv,need*share)'),
    'SPACE': (['max_inv', 'first', 'permit'], '(max_inv-first)*permit'),
    'CHEAPFILL': (['permit', 'space', 'deficit', 'cheaper_space'], 'permit*MAX(0,MIN(space,deficit-cheaper_space))'),
    'ADDED': (['round_1', 'round_2'], 'round_1+round_2'),
    'INVSPEND': (['nameplate', 'capex'], 'nameplate*capex/1000'),
    'VINTSHARE': (['new', 'cap', 'ret'], 'IF(cap+new-ret=0,0,new/(cap+new-ret))'),
}
LAMBDA_NOTES = {
    'PATH': 'Policy path: 0 before y0, linear from start (y0) to target (y1); after y1 linear if cont, else flat',
    'PATHYEARS': 'Years along a policy path: year - y0, held at y1 - y0 after the target year unless cont',
    'NOMFACTOR': 'Nominal-to-real factor of the carbon price inputs: the CPI index if they are nominal, else 1',
    'ETSON': '1 if the new ETS applies (Yes) and has started, else 0',
    'ETSPRICE': 'ETS permit price ($/tCO2) once a new ETS applies: override row if the override switch is Yes (blank '
                'cell: carbon price path), else the fast estimate',
    'OBRSHARE': 'Output-based (OBR) share of an ETS sector: its benchmark, between 0 and 1',
    'VOLPICK': 'Value of the Settings volatility table (legacy ETS+LTS) for an MTInputs label (Medium, High, Low, '
               'Zero; a trailing * ignored)',
    'VOLADJ': 'ETS volatility adjustment: (1 + carbon tax policy risk x impact) / (1 + ETS volatility x impact)',
    'CAPCHANGE': 'Relative change of the ETS cap: linear from the start to the target year, flat after',
    'COVERED': 'Covered emissions: SUMPRODUCT(emissions by ETS sector, effective coverage)',
    'ETSCAP': 'ETS cap: baseline covered emissions x (1 + relative change); after the target year constant if the '
              'continuation is Constant, else the relative change is kept',
    'ETSSEMI': 'Effective covered semi-elasticity: baseline-weighted sector semi-elasticities x allocation '
               'effectiveness (auctioned 1, OBR obr_eff)',
    'ETSESTIMATE': 'Fast ETS price estimate: LN(cap / baseline covered emissions) / (semi-elasticity x volatility '
                   'adjustment), 0 if the cap does not bind',
    'TAXEQUIV': 'ETS tax-equivalent price: permit price x volatility adjustment',
    'REVFACTOR': 'ETS revenue factor: permit price per $ of tax-equivalent price (1 / volatility adjustment)',
    'ETSCOVER': 'Effective ETS coverage of a sector: 1 if the new ETS applies, has started and covers it',
    'SHADOWSECTOR': 'Shadow price by sector ($/tCO2): feebate rate + regulations (placeholder)',
    'SHADOWSHARE': 'Share of the shadow price impacting efficiency: coverage x efficiency-margin adjustment',
    'SUPPLYCOST': 'Supply cost: data up to the last historical year (Inputs_prices, row pos), then fixed part + '
                  'floating part x gp(t)/gp(t-1)',
    'OTHERTAX': 'Excise and other taxes: data up to the last historical year, then fixed part txo_L x pcc + floating '
                'part (sp_L - sp + txo_L(1 - pcc)) x (1 - pcc), not below its last value if that is >= 0 (legacy)',
    'RETAILPRICE': 'Retail price before new policies: (supply cost + excise and other taxes) x (1 + VAT rate)',
    'VATPAY': 'VAT payment before new policies: (supply cost + excise and other taxes) x VAT rate',
    'TAXPART': 'Existing tax: positive part of excise and other taxes',
    'SUBSIDYPART': 'Existing consumer subsidy: negative part of excise and other taxes, as a positive number',
    'PERUNIT': 'Value per GJ converted to the price unit (MTInputs units: $/liter, $/bbl, $/GJ)',
    'SECTORTAX': 'Carbon tax rate of a sector ($/tCO2): carbon price x sector coverage (complementary to an existing '
                 'ETS: 1 - its sector coverage), not where the new ETS covers',
    'ECPPRICE': 'Existing carbon price used ($/tCO2 real): data of the year x (1 + growth) after the last data year, '
                'if the existing scheme applies (Yes), else 0',
    'XCARBON': 'Existing carbon pricing per unit: EF (or CO2) x (carbon tax coverage x price + ETS coverage x '
               'price); with price changes: the change against the price data',
    'SECTORETS': 'ETS cost rate of a sector in the fuel price ($/tCO2): tax-equivalent price x effective ETS coverage '
                 'x auctioned share (1 - OBR)',
    'CARBONTAX': 'New carbon tax ($/GJ): carbon tax rate of the sector x EF x fuel coverage',
    'ETSCOST': 'New ETS in the fuel price ($/GJ): ETS cost rate of the sector x EF; ETS rate x effective coverage '
               'and auctioned share are in the rate. Legacy: tax-equivalent price x EF x effective ETS coverage x '
               'auctioned share '
               '(1 - OBR)',
    'NEWEXCISE': 'New excise (fuel price reform): increase per price unit / GJ per price unit',
    'NEWPOLICY': 'Total new policy: new carbon tax + ETS + new excise',
    'POSTTAX': 'After-tax price: retail price before new policies + new policies x (1 + VAT rate), at least 0.01',
    'SHADOWTOTAL': 'Shadow price on the efficiency margin of a sector ($/tCO2): sector shadow price x share impacting '
                   'efficiency + ETS tax-equivalent price x ETS coverage x OBR share',
    'SHADOWGJ': 'Shadow price per unit of energy: shadow price ($/tCO2) x emission factor (tCO2 per GJ or per kWh)',
    'KWHPOLICY': 'New policy cost per kWh generated: tCO2/kWh x (carbon tax x fuel coverage + ETS cost) + fuel price '
                 'reform x its kWh factor',
    'FLOATTAX': 'Floating part of excise and other taxes: (sp_L - sp + txo_L (1 - pcc)) x (1 - pcc) (legacy)',
    'REVENUE': 'Revenue (USD million real): fuel use (ktoe) x PJ per ktoe x rate ($/GJ)',
    'EMISSIONS': 'CO2 emissions (MtCO2): fuel use (ktoe) x PJ per ktoe x emission factor (tCO2/GJ)',
    'NEWREVRATE': 'New-policy revenue rate ($/GJ): carbon tax + new excise + ETS x revenue factor + VAT on new '
                  'policies (feebates are revenue-neutral)',
    'FUELUSE': 'Fuel use, CPAT documentation 3.3.3, shadow price on the efficiency margin (no Covid factor); '
               'AEEI factor = (1/(1+alpha))^(1+eps_U), efficiency exponent = eps_F x (1+eps_U) (Inputs)',
    'CAPEX': 'Capital cost of new plants ($/kW): capital cost at factor 1 x capex time factor',
    'VINTAGE': 'Vintage average: last year\'s average x (1 - last year\'s investment share) + new value x last year\'s '
               'investment share (legacy B2, B6)',
    'AMORTISED': 'Amortised fixed cost ($/kWh): average capex x capex factor + other fixed costs + storage average',
    'LEVELISED': 'Levelised fixed cost of new plants ($/kWh): capex x capex factor + other fixed costs + storage if VRE',
    'VARCOST': 'Variable cost ($/kWh): variable O&M and fixed fuel + fuel price ($/GJ) x GJ per kWh + change in '
               'existing carbon pricing',
    'PNEWREV': 'Power new-policy revenue (USD million): generation (MWy) x 8.76 x policy cost ($/kWh) - CO2 x ETS '
               'cost rate x (1 - revenue factor)',
    'GENCOST': 'Generation cost before new policies ($/kWh): fixed + variable',
    'WEIGHTED': 'Generation-weighted sum: SUMPRODUCT(shares, values) x factor',
    'STORAGEHOUR': 'Short-term storage cost per marginal hour ($/kWh): annualised battery and interface capex + opex',
    'STORAGECOST': 'Marginal storage cost of variable renewables ($/kWh): short-term (alloc x hours x VRE share x cph) '
                   '+ long-term (2 x MAX(0, VRE share - start)/(1 - start)^2 x cost per unit)',
    'POWERSUPPLY': 'Supply cost of electricity: generation cost (previous year) + transmission and distribution',
    'POWERPRICE': 'Electricity price before new policies: data up to the last historical year (Inputs_power, row '
                  'pos), then the previous price '
                  '+ pass-through x change in supply cost x (1 + VAT)',
    'ENDUSERPOWER': 'End-user electricity price: price before new policies + average new policy cost (no VAT on it)',
    'SUBSIDYGAP': 'Subsidy gap ($/kWh): supply cost - price before VAT',
    'POWERREV': 'Power revenue (USD million): rate ($/kWh) x factor x consumption (GWh)',
    'REBATE': 'Feebate rebate per kWh generated: -SUMPRODUCT(shadow price per kWh, base-year generation shares)',
    'VARTOTAL': 'Total variable cost: before new policies + new policy cost + shadow price + rebate',
    'INVCOST': 'Levelised investment cost: MAX(levelised fixed + levelised variable cost, minimum)',
    'CONSUMPTION': 'Electricity consumption (GWh) of a price group: (w_total x total + w_res x residential demand) x '
                   'GWh per ktoe',
    'PRICEMIX': 'Price used by a demand group: weighted residential and industrial price',
    'DEMANDMULT': 'Growth multiple of electricity demand: efficiency multiple x GDP term x price term x (1 + shock x '
                  'switch)',
    'GROW': 'Previous value x multiple',
    'SCALE': 'Value x factor',
    'RESIDUAL': 'MAX(total - part, 0)',
    'FLOORMIN': 'Lowest value among flagged members (flag = 1), at least the floor',
    'MINDISPATCH': 'Minimum generation of coal or gas: what is still needed if the other runs at its maximum',
    'LOGITW': 'Logit weight: flag x EXP(-K x cost / lowest cost)',
    'DISPATCH': 'Generation: coal and gas MIN(maximum, minimum + remainder x logit share); other types their '
                'effective capacity',
    'INCREASE': 'MAX(0, value - previous value)',
    'SHARE': 'Part / total (0 if the total is 0)',
    'CARRY': 'Capacity: last year less last year\'s retirements plus last year\'s new investment',
    'RETIRE': 'Retirement: MIN(planned retirement (data, if switched on) or capacity x rate, + early retirement '
              '(coal), capacity)',
    'LOWEST3': 'Lowest of three values',
    'COALSHARE': 'Coal share of the early-retirement logit: EXP(-K c_coal / lowest) / (sum over coal, wind, solar), '
                 'written as 1 / (1 + EXP(-K (c_wind - c_coal) / lowest) + EXP(-K (c_solar - c_coal) / lowest))',
    'EARLYROOM': 'Room for early coal retirement: maximum wind and solar additions less demand increase and '
                 'retirements',
    'EARLYRET': 'Cost-based early coal retirement (nameplate): MIN(coal effective capacity x maximum share, room) x '
                '(1 - logit share of coal) / capacity factor x factor',
    'DIFF': 'Difference',
    'PERMITTED': '1 if new investment is permitted (Yes) and the year is at least the start year',
    'SCALEUP': 'Maximum effective investment: scale-up rate x total effective capacity x permit',
    'FIRSTROUND': 'First-round investment: MIN(maximum, need x logit share)',
    'SPACE': 'Remaining space: (maximum - first round) x permit',
    'CHEAPFILL': 'Second round, cheapest first: permit x MAX(0, MIN(space, deficit - space of cheaper types))',
    'ADDED': 'First + second round',
    'INVSPEND': 'Investment (USD million): nameplate MW x capital cost ($/kW) / 1000',
    'VINTSHARE': 'Share of new investment in the stock after retirements: new / (cap + new - ret)',
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
               'Egypt values)', t[0], t[1:], 'extract_power_v0_3.py: legacy Mitigation rows 2977-3350 (A2, A7, A8, '
               'A9, A14-A16, B2, B4, D1, D2)')
    pp = read_csv('power_paths.csv')
    data_sheet(wb, 'PowerPaths', 'Power year paths: capex time factors, storage costs, income factor, demand shock, '
               'planned retirement by type', pp[0], pp[1:],
               'extract_power_v0_3.py: legacy Mitigation rows 3038-3902')
    pd_ = read_csv('power_demand.csv')
    data_sheet(wb, 'PowerDemand', 'Electricity demand by subsector, base year (ktoe)', pd_[0], pd_[1:],
               'extract_power_v0_3.py: legacy Mitigation rows 3725-3739 (H3, energy balance 2022)')
    pm = read_csv('power_params.csv')
    data_sheet(wb, 'PowerParams', 'Power parameters: transmission and distribution, pass-through, storage',
               pm[0], pm[1:], 'extract_power_v0_3.py: legacy Mitigation rows 3596-3883')
    f = read_csv('ef_co2.csv')
    data_sheet(wb, 'EF_GHG', 'Fuel CO2 emission factors, IIASA (tCO2/GJ), by country, fuel and EF sector',
               f[0], f[1:], 'legacy Mitigation rows 980-1009 (Egypt only so far; no inventory adjustment)')


ECP_HDR = 5                             # sheet ECP: header row of the selected-country table (years)
ECP_CPI = ECP_HDR + 1                   # nominal-to-real factor by year
ECP_R0 = ECP_CPI + 1                    # first code row of the selected country
ECP_R1 = ECP_R0 + len(ECP_CODES) - 1
ECP_YL = L(4 + len(ECP_YEARS))          # last year column of the selected-country table
ECP_A0 = ECP_R1 + 5                     # first data row of the all-country table (header above)
ECP_AL = L(3 + len(ECP_YEARS))


def build_ecp(wb):
    """Sheet ECP (v1.09): existing carbon pricing. Top: the selected country's values used by the model (lookups by
    Settings C4 in the all-country table, 0 where the country has no entry; prices x CPI index = real USD of
    ResultsYear; coverage used = last data year). Below: all countries (data/ecp.csv, extract_ecp_v0_1.py)."""
    ws = wb.create_sheet('ECP')
    title(ws, 'Existing carbon pricing (ECP): carbon tax and ETS prices and coverage, legacy ECP data', 4 + len(ECP_YEARS))
    put(ws, 'A1', 'Source: legacy CPAT sheet ECP (as of Apr 1, 2025; WB Carbon Pricing Dashboard, ICAP) via '
                  'extract_ecp_v0_1.py (data/ecp.csv)', FONT)
    put(ws, 'A3', 'Selected country (Settings C4): values used by Mitigation section 1 and Inputs (0 where the country has '
                  'no entry). Prices: nominal data x CPI index (real USD/tCO2 of ResultsYear). Coverage used: last data '
                  'year (shares). ecr/eer = carbon tax rate / ETS price; ecs/ees = sector coverage (pow, trs, res, '
                  'ind); ecf/eef = fuel coverage.', FONT)
    for j, h in enumerate(['Code', 'Label', 'Unit', 'Coverage used (last data year)'], 1):
        put(ws, f'{L(j)}{ECP_HDR}', h, FONT_B, F_INPUT)
    for k, y in enumerate(ECP_YEARS):
        put(ws, f'{L(5 + k)}{ECP_HDR}', y, FONT_B, F_INPUT, '0')
        put(ws, f'{L(5 + k)}{ECP_CPI}', f'=CPIRes/INDEX(WEOVal,MATCH("USA|pcpi",WEOKey,0),MATCH({L(5 + k)}${ECP_HDR},'
                                        f'WEOYear,0))', FONT, None, '0.0000')
    put(ws, f'B{ECP_CPI}', 'Nominal-to-real factor (CPI index of ResultsYear / year)', FONT)
    put(ws, f'C{ECP_CPI}', 'factor', FONT)
    sec_lab = {'pow': 'power', 'trs': 'transport', 'res': 'buildings', 'ind': 'industry'}
    for i, code in enumerate(ECP_CODES):
        r = ECP_R0 + i
        kind, m = code.split('.')
        price = kind in ('ecr', 'eer')
        lab = {'ecr': 'Carbon tax rate', 'eer': 'ETS permit price',
               'ecs': f'Carbon tax coverage, sector {sec_lab.get(m, m)}', 'ees': f'ETS coverage, sector {sec_lab.get(m, m)}',
               'ecf': f'Carbon tax coverage, fuel {FNAMES.get(m, m)}', 'eef': f'ETS coverage, fuel {FNAMES.get(m, m)}'}[kind]
        put(ws, f'A{r}', code, FONT)
        put(ws, f'B{r}', lab, FONT)
        put(ws, f'C{r}', 'USD/tCO2 real' if price else 'share', FONT)
        if not price:
            put(ws, f'D{r}', f'={ECP_YL}{r}', FONT, None, '0.000')
        for k, y in enumerate(ECP_YEARS):
            cl = L(5 + k)
            f = f'=IFERROR(INDEX(ECPAllVal,MATCH(CtryCode&"|"&$A{r},ECPAllKey,0),MATCH({cl}${ECP_HDR},ECPAllYear,0)),0)'
            put(ws, f'{cl}{r}', f + (f'*{cl}${ECP_CPI}' if price else ''), FONT, None, '0.00' if price else '0.000')
    put(ws, f'A{ECP_A0 - 2}', 'All countries (legacy ECP; nominal USD/tCO2 for prices; coverage shares carried forward '
                              'over blank years; one ETS per country: national, else regional EU ETS)', FONT_B)
    rows = read_csv('ecp.csv')
    for j, h in enumerate(rows[0], 1):
        put(ws, f'{L(j)}{ECP_A0 - 1}', num(h) if h.isdigit() else h, FONT_B, F_INPUT)
    for i, row in enumerate(rows[1:]):
        for j, v in enumerate(row, 1):
            ws.cell(ECP_A0 + i, j, num(v) if j > 3 else v).font = FONT
    assert len(rows) - 1 <= 2000, 'ECP all-country table larger than the named range'
    assert rows[0][3:] == [str(y) for y in ECP_YEARS], 'ecp.csv years differ from ECP_YEARS'
    ws.freeze_panes = f'B{ECP_HDR + 1}'
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 36


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
                           'Feebate sector', 'Coverage position', 'ECP sector'], 1):
        put(ws, f'{L(j)}{MAP_SUB0 - 1}', h, FONT_B, F_INPUT)
    for i, s in enumerate(SUBSECTORS):
        r = MAP_SUB0 + i
        for j, v in enumerate(list(s) + [COV_SECTORS.index(s[0]) + 1, ECP_SECTOR[s[0]]], 1):
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
             'ECP sector: sector of the legacy existing carbon pricing data (sheet ECP: pow, trs, res, ind); '
             'buildings (residential, food & forestry, services) = res, industry and other energy use = ind.',
             'Feebate sector for food & forestry and services = residential (assumption; legacy groups them as '
             'Commercial).',
             f'Rows {MAP_SUB1 + 1}-{MAP_SUBLAB1} and {MAP_FUEL1 + 1}-{MAP_FUELLAB1}: labels only (totals, policy '
             'rows).']
    for i, n in enumerate(notes):
        put(ws, f'A{MAP_FB1 + 3 + i}', n)
    ws.column_dimensions['B'].width = 26
    for col in 'CDEFGHI':
        ws.column_dimensions[col].width = 16


SET_LOG = 27                            # Settings: version log title row (v1.07: rows 20-25 hold derived settings)


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
    put(ws, 'B19', 'Power demand: apply the legacy 2023-2024 energy demand shock (1 = yes; legacy row 2439; the fuel '
                   'blocks do not apply it)')
    put(ws, 'C19', 0, FONT, F_INPUT, '0')
    derived = [  # v1.07: derived settings with defined names (short formulas elsewhere)
        (20, 'Output code prefix (country.mit.; defined name CodePrefix)', '=LOWER($C$4)&".mit."', '@'),
        (22, 'US CPI of ResultsYear (WEO USA|pcpi; CPIRes, base of the real-price index)',
         '=INDEX(WEOVal,MATCH("USA|pcpi",WEOKey,0),MATCH($C$8,WEOYear,0))', '0.000'),
        (23, 'US GDP deflator of ResultsYear (WEO USA|ngdp_d; DeflRes)',
         '=INDEX(WEOVal,MATCH("USA|ngdp_d",WEOKey,0),MATCH($C$8,WEOYear,0))', '0.000'),
        (24, 'Prices_dom row of the country and base year (PDomBase)', '=MATCH($C$4&$C$7,PDomKey,0)', '0'),
        (25, 'Prices_dom row of the country and last historical price year (PDomLast)',
         '=MATCH($C$4&$C$10,PDomKey,0)', '0')]
    for r, lab, val, fmt in derived:
        put(ws, f'B{r}', lab)
        put(ws, f'C{r}', val, FONT, None, fmt)
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
         '0 (all codes vs v1.04)'),
        ('1.06', '2026-10-10', 'Power engineer model: demand, dispatch, retirement and capacity, investment (legacy H, '
                               'I, K, L) with D6, D9, E1; generation and investment shares and consumption modelled; '
                               'power CO2 in section 13; ten legacy bugs fixed (separate Word report); demand-shock '
                               'switch (Settings C19); 5 MTOutputs indicators added.',
         'intended: power block, CO2 totals, ETS with power (see check report)'),
        ('1.07', '2026-10-11', 'Short formulas: every cell formula has at most 100 characters (defined names for the '
                               'data tables, helper rows and columns, policy rates by sector in section 1, LAMBDA '
                               'ROWLABEL for descriptions); planned retirement by generation type (data rows, switch '
                               'per type, off for Egypt); memo lines of electricity demand by subsector and sector; '
                               '18 power indicators on MTOutputs, 6 power charts. Decisions: 2023 demand jump '
                               'accepted; sector policies kept separate from power demand.',
         '0 on all v1.06 codes (see check report)'),
        ('1.08', '2026-10-11', 'Works without LAMBDA support: plain formulas (at most 100 characters) in every cell, the '
                               'right column (2040) included; descriptions from a member-label table on Mapping; the '
                               'named LAMBDAs kept as an optional library (sheet Lambdas: the LAMBDA formula of every '
                               'calculated row).', '0 on all v1.07 codes (see check report)'),
        ('1.09', '2026-10-11', 'Existing carbon pricing (sheet ECP: legacy data for 52 countries, none for Egypt; '
                               'MTInputs 208-214; change against the price data in prices and revenue; complementary '
                               'carbon tax); power in the main thread: generation fuel use in the fuel totals, power '
                               'revenues (fuel taxes and subsidies, electricity VAT and subsidy gap, new policies with '
                               'the ETS at the permit price) in the revenue totals; check rows rev.chk, co2.chk.',
         'Fuel and revenue totals now include power (intended); all other v1.08 codes 0 (see check report)')]
    for i, row in enumerate(log, SET_LOG + 2):
        for j, v in enumerate(row, 2):
            put(ws, f'{L(j)}{i}', v)
    ws.column_dimensions['B'].width = 60
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 100
    ws.column_dimensions['E'].width = 26


def map_sub(col, r):
    return f'INDEX(Mapping!${col}${MAP_SUB0}:${col}${MAP_SUB1},MATCH($C{r},SubKey,0))'


def map_fuel(col, r):
    return f'INDEX(Mapping!${col}${MAP_FUEL0}:${col}${MAP_FUEL1},MATCH($B{r},FuelKey,0))'


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
        return (f'=INDEX(ElaVal,MATCH("{typ}|"&${fcol}{r}&"|"&${IC["esec"]}{r},ElaKey,0),'
                f'MATCH(Settings!$C$6,ElaGroup,0))')

    for i, (s, f) in enumerate(PAIRS):
        r = IN_R0 + i
        vals = {
            'key': f'=$B{r}&"|"&$C{r}', 'fuel': f, 'sub': s,
            'esec': '=' + map_sub('D', r), 'efuel': '=' + map_fuel('C', r), 'afuel': '=' + map_fuel('D', r),
            'prule': '=' + map_fuel('E', r), 'psec': '=' + map_sub('F', r),
            'pcode': f'=$B{r}&"."&IF(${IC["prule"]}{r}="sub",${IC["psec"]}{r},"all")',
            'efsec': '=' + map_sub('E', r), 'gj': '=' + map_fuel('G', r),
            'fpos': '=' + map_fuel('H', r), 'spos': '=' + map_sub('H', r),
            'pfpos': (f'=INDEX(Mapping!$E${MAP_PF0}:$E${MAP_PF1},MATCH(${IC["pcode"]}{r},'
                      f'Mapping!$A${MAP_PF0}:$A${MAP_PF1},0))'),
            'eY': elast('inc', IC['efuel'], r), 'eU': elast('usg', IC['efuel'], r),
            'eF': elast('eff', IC['efuel'], r), 'a': elast('aei', IC['afuel'], r),
            'aeef': f'=(1/(1+${IC["a"]}{r}))^(1+${IC["eU"]}{r})', 'effx': f'=${IC["eF"]}{r}*(1+${IC["eU"]}{r})',
            'ppos': ip_lookup('pos', f'${IC["pcode"]}{r}'),
            'vr': ip_lookup('vr', f'${IC["pcode"]}{r}'),
            'ef': f'=INDEX(EFVal,MATCH(Settings!$C$4&"|"&$B{r}&"|"&${IC["efsec"]}{r},EFKey,0))',
            'f0': f'=INDEX(ConsVal,MATCH(Settings!$C$4&"|"&Settings!$C$7&"|"&$C{r}&"|"&$B{r},ConsKey,0))',
            'ecps': '=' + map_sub('I', r),
            'ctc': ecp_cov('ecs', 'ecf', f'${IC["ecps"]}{r}', f'$B{r}'),
            'etc': ecp_cov('ees', 'eef', f'${IC["ecps"]}{r}', f'$B{r}'),
        }
        for _, code in INPUT_COLS:
            fmt = {'ctc': '0.000', 'etc': '0.000', 'vr': '0.0%', 'ef': '0.0000', 'f0': '#,##0.0', 'aeef': '0.0000', 'effx': '0.00',
                   'eY': '0.00', 'eU': '0.00', 'eF': '0.00', 'a': '0.0%'}.get(code)
            put(ws, f'{IC[code]}{r}', vals[code], FONT, F_INPUT if code in ('fuel', 'sub') else None, fmt)
    ws.freeze_panes = 'D5'
    ws.column_dimensions['A'].width = 10
    for j in range(2, len(INPUT_COLS) + 1):
        ws.column_dimensions[L(j)].width = 10
    ws.row_dimensions[4].height = 48


def ecp_cov(sk, fk, sec, fuel, guard=False):
    """Existing carbon pricing coverage (v1.09): sector share x fuel share of sheet ECP (IFERROR for power types
    without a fuel)."""
    f = (f'INDEX(ECPCov,MATCH("{sk}."&{sec},ECPKey,0))*INDEX(ECPCov,MATCH("{fk}."&{fuel},ECPKey,0))')
    return f'=IFERROR({f},0)' if guard else '=' + f


def ip_lookup(code, key):
    """Inputs_prices column `code` for price code `key` (a cell reference or formula)."""
    c = IPC[code]
    return f'=INDEX(Inputs_prices!${c}${IP_R0}:${c}${IP_R1},MATCH({key},Inputs_prices!$A${IP_R0}:$A${IP_R1},0))'


def build_inputs_prices(wb):
    """Data step for section 2: one row per price fuel. Historical prices converted to real USD of ResultsYear
    with the CPI index of their year and to $/GJ; forecasting coefficients (margin, production cost, pass-through) of
    the base year; values of the last historical year. v1.07: the data of each historical year are read in their own
    helper columns (nominal retail price, supply cost, excise and other taxes, VAT blank flag, VAT data, VAT used), with
    the Prices_dom row, CPI index and general VAT rate of the year in rows 3-5, so every formula stays short."""
    ws = wb.create_sheet('Inputs_prices')
    title(ws, 'Inputs by price fuel: historical prices (real $/GJ) and forecasting coefficients (data step)',
          len(IP_COLS))
    put(ws, f'A{IP_PROW}', 'Prices_dom row of the column year ->', FONT_B)
    put(ws, f'A{IP_CPI}', 'CPI index of the column year (ResultsYear = 1) ->', FONT_B)
    put(ws, f'A{IP_VATW}', 'General VAT rate VAT_WEO of the column year (blank = 0) ->', FONT_B)
    for j, (h, _) in enumerate(IP_COLS, 1):
        c = put(ws, f'{L(j)}{IP_HDR}', h, FONT_B, F_INPUT)
        c.alignment = Alignment(wrap_text=True, vertical='top')
    for code in IPC:
        y = re.match(r'[a-z]+(\d{4})$', code)
        if not y:
            continue
        col = IPC[code]
        put(ws, f'{col}{IP_HDR}', int(y.group(1)), FONT_B, F_INPUT, '0')
        put(ws, f'{col}{IP_PROW}', f'=MATCH(Settings!$C$4&{col}${IP_HDR},PDomKey,0)', FONT, None, '0')
        put(ws, f'{col}{IP_CPI}', f'=CPIRes/INDEX(WEOVal,MATCH("USA|pcpi",WEOKey,0),MATCH({col}${IP_HDR},WEOYear,0))',
            FONT, None, '0.0000')
        put(ws, f'{col}{IP_VATW}', f'=IFERROR(INDEX(PDomVal,{col}${IP_PROW},MATCH("VAT_WEO",PDomHdr,0))+0,0)', FONT,
            None, '0.0%')
    blk = lambda v: (IPC[f'{v}{HIST_YEARS[0]}'], IPC[f'{v}{HIST_YEARS[-1]}'])

    def dom(var, prow, r):
        """IMF price dataset value (blank = 0) of mit.<var>.<price code> in the Prices_dom row prow, nominal."""
        return f'IFERROR(INDEX(PDomVal,{prow},MATCH("mit.{var}."&$A{r},PDomHdr,0))+0,0)'

    for i, (pc, lab) in enumerate(PRICES):
        r = IP_R0 + i
        fuel, sec = pc.split('.')
        c = lambda k: f'${IPC[k]}{r}'
        at = lambda v, y: f'INDEX(${blk(v)[0]}{r}:${blk(v)[1]}{r},MATCH({y},${blk(v)[0]}${IP_HDR}:${blk(v)[1]}${IP_HDR},0))'
        base_cpi = f'INDEX(${blk("sp")[0]}${IP_CPI}:${blk("sp")[1]}${IP_CPI},MATCH(Settings!$C$7,' \
                   f'${blk("sp")[0]}${IP_HDR}:${blk("sp")[1]}${IP_HDR},0))'
        vals = {
            'key': pc, 'lab': lab, 'fuel': fuel, 'sec': sec, 'pos': i + 1, 'gpos': GP_POS[fuel],
            'gj': '=' + map_fuel('G', r).replace(f'$B{r}', c('fuel')),
            'rpd': 1 if fuel in RP_FROM_DATA else 0,
            'pccr': 1 if fuel == 'bio' else f'={dom("ps", "PDomBase", r)}',
            'pccb': f'=IF({c("pccr")}<=0.25,0,IF({c("pccr")}<=0.5,0.5,1))',
            'pcc': f'=IF(LEFT(Settings!$C$9,6)="Manual",0.8,IF(LEFT(Settings!$C$9,4)="None",1,{c("pccb")}))',
            'mar': 0 if fuel == 'bio' else f'={dom("mar", "PDomBase", r)}',
            'pcost': (f'=IFERROR(INDEX(PDomVal,PDomBase,MATCH("mit.{fuel}.prod.cost",PDomHdr,0))+0,0)'
                      if fuel in ('coa', 'nga') else 0),
            'bcpi': f'={base_cpi}',
            'fixsp': f'=({c("mar")}+{c("pccb")}*{c("pcost")})*{c("bcpi")}/{c("gj")}',
            'vatc': 1 if pc in VAT_CONSUMER else 0,
            'vr': f'={at("vry", "Settings!$C$10")}',
            'spL': f'={at("sp", "Settings!$C$10")}',
            'txoL': f'={at("txo", "Settings!$C$10")}',
        }
        last_vatw = (f'INDEX(${blk("vry")[0]}${IP_VATW}:${blk("vry")[1]}${IP_VATW},MATCH(Settings!$C$10,'
                     f'${blk("vry")[0]}${IP_HDR}:${blk("vry")[1]}${IP_HDR},0))')
        vals['vrx'] = f'={c("vatc")}*{last_vatw}'          # display: the assumption (vr reads the year columns)
        for y in HIST_YEARS:
            h = lambda v: f'{IPC[v + str(y)]}'
            vals[f'rpn{y}'] = f'={dom("rp", f"{h("rpn")}${IP_PROW}", r)}'
            vals[f'spn{y}'] = f'={dom("sp", f"{h("spn")}${IP_PROW}", r)}'
            vals[f'txn{y}'] = f'={dom("txo", f"{h("txn")}${IP_PROW}", r)}'
            vrd = f'INDEX(PDomVal,{h("vrb")}${IP_PROW},MATCH("mit.vatrate."&$A{r},PDomHdr,0))'
            vals[f'vrb{y}'] = f'=IF(IFERROR(ISBLANK({vrd}),TRUE),1,0)'
            vals[f'vrd{y}'] = f'={dom("vatrate", f"{h("vrd")}${IP_PROW}", r)}'
            vals[f'vry{y}'] = f'=IF(${h("vrb")}{r}=1,{c("vatc")}*{h("vry")}${IP_VATW},${h("vrd")}{r})'
            vals[f'sp{y}'] = f'=${h("spn")}{r}*{h("sp")}${IP_CPI}/{c("gj")}'
            vals[f'txo{y}'] = (f'=IF({c("rpd")}=1,${h("rpn")}{r}/(1+${h("vry")}{r})-${h("spn")}{r},${h("txn")}{r})'
                               f'*{h("txo")}${IP_CPI}/{c("gj")}')
        for _, code in IP_COLS:
            v = vals[code]
            fmt = '0.0%' if code == 'vr' or code.startswith('vr') and code[2:3] in 'dy' else (
                '0.00' if code.startswith('pcc') else ('0' if code in ('pos', 'gpos', 'rpd') or code.startswith('vrb')
                                                       else '0.000'))
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
    ws.freeze_panes = f'C{IP_R0}'
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 22
    for j in range(3, len(IP_COLS) + 1):
        ws.column_dimensions[L(j)].width = 11
    ws.row_dimensions[IP_HDR].height = 72


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
    ('Base-year capacity, nameplate (MW)', 'cap0', 'cap0_mw'), ('Base-year generation (MWy)', 'gen0', 'gen0_mwy'),
    ('Retirement rate (1 / lifetime; hydro 0)', 'rr', 'retire_rate'),
    ('New investment permitted (Yes / No)', 'perm', 'invest_permitted'), ('Permitted from year', 'start', 'invest_start'),
    ('Scale-up limit: maximum effective investment as a share of effective capacity', 'rate', 'scaleup_rate'),
    ('Dispatch type (cg = coal or gas, filled by the logit; fixed = capacity x capacity factor)', 'disp', 'dispatch'),
    ('Planned retirement data used? (Yes / No; data rows PowerPaths pret|<type>)', 'puse0', 'planned_retirement'),
    ('Planned retirement used (1 / 0)', 'puse', None),
    ('Coal or gas (1 / 0)', 'cg', None), ('Early-retirement type: coal (1 / 0)', 'eflag', None),
    ('Nameplate MW per effective MW (1 / capacity factor)', 'ninv', None),
    ('Base-year generation share', 'shr0', None),
    ('Fuel use per generation: 8.76 / GWh per ktoe / efficiency (ktoe per MWy)', 'pef', None),
    ('CO2 per fuel use: EF x PJ per ktoe (MtCO2 per ktoe)', 'co2f', None),
    ('Tie breaker for the cost ranking (legacy L11)', 'tb', None),
    ('Existing carbon tax coverage: power x fuel share (sheet ECP; 0 without a fuel)', 'ctc', None),
    ('Existing ETS coverage: power x fuel share (sheet ECP; 0 without a fuel)', 'etc', None),
    ('Fuel of the sector model (1 / 0): its fuel use counts in the fuel totals (section 11)', 'fflag', None),
]
IPWC = {code: L(j) for j, (_h, code, _s) in enumerate(IPW_COLS, 1)}
IPW_G0 = IPW_R1 + 6                     # user-group table: Prices_dom row G0 - 3, CPI G0 - 2, header G0 - 1, rows
IPW_G1 = IPW_G0 + len(PW_GROUPS) - 1
IPWG_COLS = [('User group', 'grp'), ('Label', 'lab'), ('Position', 'pos'),
             ('Consumption: weight on total demand', 'wtot'), ('Consumption: weight on residential demand', 'wres'),
             ('Transmission and distribution ($/kWh)', 'tmc'), ('Pass-through of supply-cost changes', 'pass'),
             ('VAT rate, last historical year (data; blank = 0)', 'vr'), ('VAT share of the price: VAT / (1 + VAT)',
                                                                         'vshare')] + \
    [(f'Retail price {y} (real $/kWh, incl. VAT)', f'rp{y}') for y in HIST_YEARS]
IPWG = {code: L(j) for j, (_h, code) in enumerate(IPWG_COLS, 1)}
IPWH0 = IPW_G1 + 9                      # demand-group table (legacy H1): header IPWH0 - 1
IPWH1 = IPWH0 + len(DEM_GROUPS) - 1
IPWH_COLS = [('Demand group', 'grp'), ('Label', 'lab'), ('Position', 'pos'),
             ('Price weight on the residential price (transport 0.5; industry, services 0)', 'wres'),
             ('GDP elasticity eps_Y', 'ey'), ('Usage elasticity eps_U', 'eu'), ('Efficiency elasticity eps_F', 'ef'),
             ('Autonomous efficiency gain alpha', 'alpha'),
             ('Efficiency multiple (1 / (1 + alpha + additional efficiency))^(1 + eps_U)', 'aeef'),
             ('Total price elasticity eps_U + eps_F + eps_U x eps_F (legacy row 3679)', 'eps'),
             ('Price of 2021 (real $/kWh; weighted)', 'p2021')]
IPWH = {code: L(j) for j, (_h, code) in enumerate(IPWH_COLS, 1)}
IPWS0 = IPWH1 + 4                       # demand by subsector (legacy H3): header IPWS0 - 1
IPWS1 = IPWS0 + len(DEM_SUBS) - 1
IPWS_COLS = [('Subsector', 'sub'), ('Label', 'lab'), ('Demand group', 'grp'), ('Demand-group position', 'gpos'),
             ('Electricity demand, base year (ktoe)', 'd0')]
IPWS = {code: L(j) for j, (_h, code) in enumerate(IPWS_COLS, 1)}
IPWQ0 = IPWS1 + 4                       # derived system parameters: header IPWQ0 - 1
IPWQ = [('sh_nxp', 'Net exports / final demand, base year (PowerParams)'),
        ('sh_tsd', 'Transmission and statistical differences / final demand (PowerParams)'),
        ('sh_eou', 'Energy industry own use / final demand (PowerParams)'),
        ('sh_add', 'Additional statistical difference / final demand (PowerParams)'),
        ('ktoe_gwh', 'GWh per ktoe (PowerParams)'),
        ('req_factor', 'Required generation per final demand: (1 + net exports + transmission + own use + stat. diff. '
                       'shares) x GWh per ktoe x 1000 / 8760 (MWy per ktoe)'),
        ('cf_coal', 'Capacity factor of coal (early retirement, effective to nameplate)')]
IPWQ_ROW = {k: IPWQ0 + i for i, (k, _d) in enumerate(IPWQ)}
_rp0, _rp1 = IPWG['rp' + str(HIST_YEARS[0])], IPWG['rp' + str(HIST_YEARS[-1])]
LAMBDAS['POWERPRICE'] = (LAMBDAS['POWERPRICE'][0], LAMBDAS['POWERPRICE'][1].replace(
    'HIST', f'IFERROR(INDEX(Inputs_power!${_rp0}${IPW_G0}:${_rp1}${IPW_G1},pos,MATCH(year,'
            f'Inputs_power!${_rp0}${IPW_G0 - 1}:${_rp1}${IPW_G0 - 1},0)),0)'))


def pw_param(name):
    """PowerParams value by key."""
    return f'INDEX(PParVal,MATCH("{name}",PParKey,0))'


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
    tk = lambda hdr, r: f'INDEX(PTechVal,MATCH($A{r},PTechKey,0),MATCH("{hdr}",PTechHdr,0))'
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
            'ef': f'=IFERROR(INDEX(EFVal,MATCH(Settings!$C$4&"|"&$A{r}&"|pow",EFKey,0)),0)',
            'efk': f'={c("ef")}*{c("k")}',
            'fpos': f'=IF({nofuel},1,INDEX(Mapping!$H${MAP_FUEL0}:$H${MAP_FUEL1},MATCH($A{r},FuelKey,0)))',
            'pfpos': (f'=IF({nofuel},1,INDEX(Mapping!$E${MAP_PF0}:$E${MAP_PF1},MATCH({c("pcode")},'
                      f'Mapping!$A${MAP_PF0}:$A${MAP_PF1},0)))'),
            'gj': f'=IF({nofuel},1,INDEX(Mapping!$G${MAP_FUEL0}:$G${MAP_FUEL1},MATCH($A{r},FuelKey,0)))',
            'kpu': f'={c("k")}/{c("gj")}',
            'cg': f'=IF({c("disp")}="cg",1,0)', 'eflag': f'=IF($A{r}="coa",1,0)', 'ninv': f'=1/{c("cf")}',
            'puse': f'=IF({c("puse0")}="Yes",1,0)',
            'shr0': f'={c("gen0")}/SUM(${IPWC["gen0"]}${IPW_R0}:${IPWC["gen0"]}${IPW_R1})',
            'pef': f'=8.76/{pw_param("ktoe_gwh")}/{c("nu")}', 'co2f': f'={c("ef")}*Settings!$C$13',
            'tb': f'={i}*1.356E-06',
            'ctc': ecp_cov('ecs', 'ecf', '"pow"', f'$A{r}', guard=True),
            'etc': ecp_cov('ees', 'eef', '"pow"', f'$A{r}', guard=True),
            'fflag': f'=IF(ISNUMBER(MATCH($A{r},FuelKey,0)),1,0)',
        }
        for _h, code, src in IPW_COLS:
            v = vals.get(code) or ('=' + tk(src, r) + ('&""' if code in ('name', 'pcode', 'perm', 'disp', 'puse0')
                                                       else ''))
            fmt = {'cax0': '#,##0.0', 'N': '0', 'vre': '0', 'vsh': '0', 'ppos': '0', 'fpos': '0', 'pfpos': '0',
                   'dlf': '0.000', 'gj': '0.00', 'wacc': '0.0%', 'cf': '0.000', 'nu': '0.000', 'cap0': '#,##0',
                   'gen0': '#,##0', 'perm': '@', 'start': '0', 'disp': '@', 'cg': '0', 'eflag': '0', 'rate': '0.0%',
                   'puse0': '@', 'puse': '0', 'fflag': '0', 'ctc': '0.000', 'etc': '0.000',
                   'shr0': '0.0%', 'tb': '0.0E+00'}.get(code, '0.000000')
            put(ws, f'{IPWC[code]}{r}', v, FONT, F_INPUT if code == 'key' else None, fmt)
    # end-user groups
    put(ws, f'A{IPW_G0 - 4}', 'End-user groups (legacy G1 residential, G2 industrial)', FONT_B)
    put(ws, f'A{IPW_G0 - 3}', 'Prices_dom row of the column year ->', FONT_B)
    put(ws, f'A{IPW_G0 - 2}', 'CPI index of the column year (ResultsYear = 1) ->', FONT_B)
    for j, (h, code) in enumerate(IPWG_COLS, 1):
        c = put(ws, f'{L(j)}{IPW_G0 - 1}', h, FONT_B, F_INPUT)
        c.alignment = Alignment(wrap_text=True, vertical='top')
    for y in HIST_YEARS:
        col = IPWG[f'rp{y}']
        put(ws, f'{col}{IPW_G0 - 1}', y, FONT_B, F_INPUT, '0')
        put(ws, f'{col}{IPW_G0 - 3}', f'=MATCH(Settings!$C$4&{col}${IPW_G0 - 1},PDomKey,0)', FONT, None, '0')
        put(ws, f'{col}{IPW_G0 - 2}', f'=CPIRes/INDEX(WEOVal,MATCH("USA|pcpi",WEOKey,0),MATCH({col}${IPW_G0 - 1},WEOYear,0))',
            FONT, None, '0.0000')

    def dom(var, prow, r):
        """Prices_dom value (blank = 0) of mit.<var>.ecy.<group> in the Prices_dom row prow."""
        return f'IFERROR(INDEX(PDomVal,{prow},MATCH("mit.{var}.ecy."&$A{r},PDomHdr,0))+0,0)'
    for i, (g, lab) in enumerate(PW_GROUPS):
        r = IPW_G0 + i
        vals = {'grp': g, 'lab': lab, 'pos': i + 1, 'tmc': f'={pw_param("tmc|" + g)}', 'pass': f'={pw_param("pass")}',
                'wtot': 0 if g == 'res' else 1, 'wres': 1 if g == 'res' else -1,
                'vr': f'={dom("vatrate", "PDomLast", r)}', 'vshare': f'=${IPWG["vr"]}{r}/(1+${IPWG["vr"]}{r})'}
        for y in HIST_YEARS:
            col = IPWG[f'rp{y}']
            vals[f'rp{y}'] = f'={dom("rp", f"{col}${IPW_G0 - 3}", r)}*{col}${IPW_G0 - 2}'
        for _h, code in IPWG_COLS:
            v = vals[code]
            put(ws, f'{IPWG[code]}{r}', v, FONT, F_INPUT if not str(v).startswith('=') else None,
                '0' if code in ('pos', 'wtot', 'wres') else ('0.0%' if code in ('vr', 'vshare') else '0.0000'))
    # demand groups (legacy H1): elasticities of electricity by income group, price weights, 2021 price
    put(ws, f'A{IPWH0 - 2}', 'Power demand groups (legacy H1; elasticities from sheet Elasticities, fuel ecy, by income '
                             'group)', FONT_B)
    for j, (h, code) in enumerate(IPWH_COLS, 1):
        c = put(ws, f'{L(j)}{IPWH0 - 1}', h, FONT_B, F_INPUT)
        c.alignment = Alignment(wrap_text=True, vertical='top')
    el = lambda typ, r: f'=INDEX(ElaVal,MATCH("{typ}|ecy|"&$A{r},ElaKey,0),MATCH(Settings!$C$6,ElaGroup,0))'
    rp21 = lambda g: f'${IPWG["rp" + str(HIST_YEARS[0])]}${IPW_G0 + [x for x, _l in PW_GROUPS].index(g)}'
    for i, (g, lab) in enumerate(DEM_GROUPS):
        r = IPWH0 + i
        c = lambda k: f'${IPWH[k]}{r}'
        vals = {'grp': g, 'lab': lab, 'pos': i + 1, 'wres': {'tra': 0.5, 'res': 1}.get(g, 0),
                'ey': el('inc', r), 'eu': el('usg', r), 'ef': el('eff', r), 'alpha': el('aei', r),
                'aeef': f'=(1/(1+{c("alpha")}+{pw_param("extra_eff")}))^(1+{c("eu")})',
                'eps': f'={c("eu")}+{c("ef")}+{c("eu")}*{c("ef")}',
                'p2021': f'={c("wres")}*{rp21("res")}+(1-{c("wres")})*{rp21("ind")}'}
        for _h, code in IPWH_COLS:
            v = vals[code]
            put(ws, f'{IPWH[code]}{r}', v, FONT, F_INPUT if code in ('grp', 'lab', 'wres') else None,
                '0' if code == 'pos' else '0.0000')
    # demand by subsector (legacy H3)
    put(ws, f'A{IPWS0 - 2}', 'Power demand by subsector, base year (legacy H3; data sheet PowerDemand)', FONT_B)
    for j, (h, code) in enumerate(IPWS_COLS, 1):
        c = put(ws, f'{L(j)}{IPWS0 - 1}', h, FONT_B, F_INPUT)
        c.alignment = Alignment(wrap_text=True, vertical='top')
    for i, s_ in enumerate(DEM_SUBS):
        r = IPWS0 + i
        g = DEM_GROUP_OF.get(s_, 'ind')
        vals = {'sub': s_, 'lab': SNAMES.get(s_, s_), 'grp': g,
                'gpos': f'=MATCH(${IPWS["grp"]}{r},${IPWH["grp"]}${IPWH0}:${IPWH["grp"]}${IPWH1},0)',
                'd0': f'=INDEX(PowerDemand!$C$4:$C$100,MATCH($A{r},PowerDemand!$A$4:$A$100,0))'}
        for _h, code in IPWS_COLS:
            put(ws, f'{IPWS[code]}{r}', vals[code], FONT, F_INPUT if code in ('sub', 'grp') else None,
                '0' if code == 'gpos' else '#,##0.0')
    # derived system parameters
    put(ws, f'A{IPWQ0 - 2}', 'Derived system parameters', FONT_B)
    for j, h in enumerate(['Key', 'Value', 'Description'], 1):
        put(ws, f'{L(j)}{IPWQ0 - 1}', h, FONT_B, F_INPUT)
    q = lambda k: f'$B${IPWQ_ROW[k]}'
    qv = {**{k: f'={pw_param(k)}' for k in ('sh_nxp', 'sh_tsd', 'sh_eou', 'sh_add', 'ktoe_gwh')},
          'req_factor': f'=(1+{q("sh_nxp")}+{q("sh_tsd")}+{q("sh_eou")}+{q("sh_add")})*{q("ktoe_gwh")}*1000/8760',
          'cf_coal': f'=${IPWC["cf"]}${IPW_R0 + [m for m, _n in PW_TECH].index("coa")}'}
    for key, desc in IPWQ:
        r = IPWQ_ROW[key]
        put(ws, f'A{r}', key, FONT, F_INPUT)
        put(ws, f'B{r}', qv[key], FONT, None, '0.000000')
        put(ws, f'C{r}', desc)
    notes = [
        'Derived constants follow PowerPrices_Method_v0_3.md: amortised fixed cost = caxav x acap + afix + stoav '
        '(legacy B3-B7 in one line); levelised fixed cost = cax x lcap + lfix + VRE x msc (legacy C1-C6); variable '
        'cost = vfix + rpb x k (legacy D4, with k = 0.0036 / efficiency).',
        'Types without a price fuel (nuclear, wind, solar, hydro, other renewables) take position 1 with factor 0 '
        '(k, efk, kpu = 0), so the price and policy terms vanish.',
        'Historical electricity prices: Prices_dom mit.rp.ecy.res / .ind (corrected Egypt block, nominal $/kWh incl. '
        'VAT) x CPI index of the year. VAT rate: Prices_dom mit.vatrate.ecy.* of the last historical year (blank = 0).']
    for k, n in enumerate(notes):
        put(ws, f'A{IPWQ0 + len(IPWQ) + 2 + k}', n)
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
     'adjustment); override row pasted from ets_goalseek_v0_4.py (Python, same damped log-space step, no VBA); '
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
    ('Power engineer model (v1.06)', 'Tables H (demand), I (dispatch), K (retirement, capacity), L (investment)',
     'Same method, year by year (method test: 208 legacy rows of the baseline and the carbon-tax run reproduced to '
     '1e-13 with the bugs switched on)', 'User request v1.06: align except for bugs', 'Generation mix, capacity, '
     'investment and power CO2 now respond to each scenario', 'Same'),
    ('Power: bug fixes (v1.06)', 'Ten defects (PowerModel_BugsAndImprovements_v1_0.docx): services demand with the '
     'industrial multiple from 2024; efficiency exponent with eps_F; income factor twice; LT storage allocation '
     'twice; investment spending at base capital cost; investment logit normalised by non-permitted types; '
     'early-retirement units; coal effective retirement uncapped; feebate rebate x 1.22; capacity check on the plan',
     'Fixed', 'User request v1.06', 'Legacy Egypt inputs: generation 2040 +12.6% (baseline), +9.2% (carbon tax); '
     'power CO2 2040 +10.8%; the last six are latent for Egypt', 'Bug fix'),
    ('Power demand: 2023-2024 demand shock', 'Exogenous shock -20.8% (2023), +9.2% (2024) on power demand',
     'Switch Settings C19, off by default (the fuel blocks do not apply it either); the resulting 2023 demand jump '
     '(USD electricity prices after the devaluation) is accepted', 'User decision v1.07 (consistency with the fuel '
     'blocks)', 'Power demand 2023-2040 higher than legacy', 'Decision'),
    ('Power demand and sector policies', 'Power demand from its own elasticities (H); sector policies act on fuels '
     'only', 'Same; electricity demand shown as memo lines in the sector blocks and section 11 (not in the fuel '
     'totals)', 'User decision v1.07: keep sector policies separate', 'None', 'Same'),
    ('Power: planned retirement', 'Coal only (K5), data from sheet Power when Use Data = Yes; else the rate',
     'Every type: switch planned_retirement (PowerTech) and data rows pret|<type> (PowerPaths); retirement = MIN('
     'planned or rate-based + early coal, capacity)', 'User request v1.07', 'None for Egypt (no data, switches No)',
     'Design'),
    ('Power demand: road electricity', 'From the road module (EV electricity)', 'Grows with the transport multiple '
     '(legacy fallback)', 'No road module in the MVP', 'None for Egypt (0)', 'Not yet'),
    ('Power: features not built', 'PPAs, manual capacity and investment, power demand '
     'from the elasticity model, carbon price on power use (feebates), apply final deficit to gas', 'Not built '
     '(all off or 0 in the Egypt runs)', 'Scope', 'None for Egypt', 'Not yet'),
    ('Power: levelised variable cost window', 'To the last model year 2041', 'To the last model year 2040',
     'Horizon of the MVP', 'Small, at the end of the horizon', 'Decision'),
    ('Power: generation costs (v1.05)', 'Tables B3-B8 (amortised), C1-C7 (levelised), A5-A6 discount factors',
     'One row per type for amortised and for levelised fixed cost; discounted lifetime in closed form; constants '
     'derived in Inputs_power', 'Fewer, readable rows (PowerPrices_Method_v0_3.md S1-S3)', 'None (method test: '
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
    ('Power: coal implicit cost', 'D3 (calibrated on the coal share)', 'Not built (D6 shadow price and rebate built '
     'in v1.06)', '0 for Egypt (no coal in the mix)', 'Other countries with coal', 'Not yet'),
    ('Power: new-policy revenue', 'Carbon revenue from power', 'v1.09: by generation type (prnew): policy cost x '
     'generation, the ETS part at the permit price (as the subsectors); rcarb (ccpav x consumption) kept as a memo',
     'Consistency with the subsectors', 'None (fixed in v1.09)', 'Same'),
    ('Emissions', 'All GHGs and local pollutants; inventory-adjusted EFs; power and process emissions',
     'CO2 from fuel combustion only, IIASA EFs without inventory adjustment; power from v1.06; no process emissions',
     'User decision v0.13: CO2 first', 'Emission levels', 'Not yet'),
    ('Revenue structure', 'Revenue by instrument incl. electricity, producer subsidies, ETS allowances, '
     'GDP feedback', 'Three separate calculations per subsector and fuel: existing taxes (txo positive part + '
     'VAT), existing consumer subsidies, new policies (carbon tax, excise, auctioned ETS, VAT on them); power '
     '(v1.09): fuel taxes and subsidies on generation, electricity VAT, electricity subsidy gap (signed: negative '
     'when the price is above cost), new policies; no producer subsidies, no GDP feedback',
     'User request v0.12; scope of the prototype', 'Revenue levels', 'Simplification'),
    ('Existing carbon pricing', 'ECP data by country (national and regional ETS blocks, carbon tax with fuel-specific '
     'rates), coverage by year; ctx and ets inside txo; free allocation and revenue data',
     'One ETS per country (national, else regional EU ETS); carbon tax single rate; coverage of the last data year; '
     'only the change against the level in the price data enters prices and revenue (the IMF data contain the '
     'level); existing ETS revenue at full price; existing and new ETS add up where both cover a sector',
     'Simplicity; IMF retail prices include existing pricing', 'None for Egypt (no existing pricing)',
     'Simplification'),
    ('Power in the totals', 'Total energy and revenue include power', 'v1.09: generation fuel use (coal, gas, oil, '
     'biomass) in the fuel totals of section 11; power revenues in section 12; electricity final demand a memo',
     'Main thread', 'Totals now include power (36.7 Mtoe of 86.4 in 2022)', 'Same'),
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
    # v1.09: power in the main thread, existing carbon pricing
    ('mit.ener.pow.all.e', '#,##0'), ('mit.rtx.pow.all', '#,##0'), ('mit.rsub.pow.all', '#,##0'),
    ('mit.rnew.pow.all', '#,##0'), ('mit.rxcp.all.all', '#,##0'), ('mit.xct.p', '0.00'), ('mit.xets.p', '0.00'),
    ('mit.cptraj', '0.00'), ('mit.ets.p', '0.00'), ('mit.co2.ets.all.all', '#,##0.0'),
    ('mit.co2.cap.all.all', '#,##0.0'),
    ('mit.rpb.all.gso.a', '0.00'), ('mit.rpb.all.die.a', '0.00'), ('mit.rpb.res.nga.a', '0.00'),
    ('mit.atp.rod.gso.e', '0.00'),
    ('mit.rp.res.ecy', '0.0000'), ('mit.rp.ind.ecy', '0.0000'), ('mit.gncav', '0.0000'),
    ('mit.co2.pow.all.e', '#,##0.0'), ('mit.gtot', '#,##0'), ('mit.vre', '0.0%'), ('mit.shr.pow.nga', '0.0%'),
    ('mit.usdtot', '#,##0'),
    # v1.07: power indicators
    ('mit.demtot', '#,##0'), ('mit.eldem.tra.ecy.e', '#,##0'), ('mit.eldem.bld.ecy.e', '#,##0'),
    ('mit.eldem.ind.ecy.e', '#,##0'), ('mit.eldem.oen.ecy.e', '#,##0'), ('mit.req', '#,##0'),
    ('mit.captot', '#,##0'), ('mit.cap.pow.wnd', '#,##0'), ('mit.cap.pow.sol', '#,##0'), ('mit.nimtot', '#,##0'),
    ('mit.shr.pow.coa', '0.0%'), ('mit.shr.pow.oop', '0.0%'), ('mit.shr.pow.wnd', '0.0%'), ('mit.shr.pow.sol', '0.0%'),
    ('mit.shr.pow.hyd', '0.0%'), ('mit.ccpav', '0.0000'), ('mit.rcarb.res.ecy', '#,##0'), ('mit.rcarb.ind.ecy', '#,##0'),
]
MTO_R0 = 6                              # first block header row
MTO_BLOCK = 1 + len(MTO_INDICATORS) + 1  # header + indicators + blank row
MTO_C0 = 6                              # F: first year column
MIT_RANGE = '$A$1:$ZZ$5000'             # search area on Mitigation (room for copied scenario groups)
CHARTS = [  # chart title, indicator code stem, y-axis title
    ('CO2 emissions from fuel combustion', 'mit.co2.all.all.e', 'MtCO2'),
    ('Fuel use, all subsectors and power generation', 'mit.ener.all.all.e', 'ktoe'),
    ('Fiscal effect vs scenario 1 (total revenue change)', 'mit.rtot.chg.all.all', 'USD million, real 2026'),
    ('Carbon price', 'mit.cptraj', 'USD/tCO2, real 2026'),
    ('Power generation, total', 'mit.gtot', 'MWy (average MW)'),
    ('Power CO2 emissions', 'mit.co2.pow.all.e', 'MtCO2'),
    ('Variable renewable share of generation', 'mit.vre', 'share'),
    ('Residential electricity price', 'mit.rp.res.ecy', 'USD/kWh, real 2026'),
    ('Nameplate capacity, total', 'mit.captot', 'MW'),
    ('Electricity demand, total', 'mit.demtot', 'ktoe'),
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
    hc, hr = L(last + 1), L(last + 2)                       # v1.07: lookup positions (helper columns)
    put(ws, f'{hc}4', 'Mitigation code column', FONT_B, F_INPUT)
    put(ws, f'{hr}4', 'Mitigation row', FONT_B, F_INPUT)
    for k in range(ny):
        col = L(MTO_C0 + k)
        put(ws, f'{col}4', '=Settings!$C$7' if k == 0 else f'={L(MTO_C0 + k - 1)}4+1', FONT_B,
            F_BASE if k == 0 else (F_LIGHT if k == ny - 1 else F_INPUT), '0')
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
            put(ws, f'A{r}', f'=LOWER(Settings!$C$4)&"."&$B{r}&"."&$E{r}', FONT, F_CODE)
            put(ws, f'B{r}', stem, FONT, F_INPUT)
            put(ws, f'C{r}', f'=IFERROR(INDEX(MitDesc,${hr}{r}),"code not found")')
            put(ws, f'D{r}', f'=IFERROR(INDEX(MitUnit,${hr}{r}),"")')
            put(ws, f'E{r}', f'=E{r - 1}', FONT, None, '0')
            put(ws, f'{hc}{r}', f'=MATCH($E{r},MitScen,0)', FONT, None, '0')
            put(ws, f'{hr}{r}', f'=MATCH($A{r},INDEX(MitAll,0,${hc}{r}),0)', FONT, None, '0')
            for kk in range(ny):
                col = L(MTO_C0 + kk)
                put(ws, f'{col}{r}', f'=IFERROR(INDEX(MitAll,${hr}{r},${hc}{r}+1+{col}$4-Settings!$C$7),"")', FONT,
                    F_BASE if kk == 0 else None, fmt)
        r = h + MTO_BLOCK - 1
        put(ws, f'E{r}', f'=E{r - 1}', Font(name='Arial', size=9, color='BFBFBF'), None, '0')
    ws.freeze_panes = f'{L(MTO_C0)}5'
    for col, w in {'A': 30, 'B': 22, 'C': 52, 'D': 12, 'E': 8}.items():
        ws.column_dimensions[col].width = w
    for k in range(ny):
        ws.column_dimensions[L(MTO_C0 + k)].width = 9
    ws.column_dimensions[hc].width = 11
    ws.column_dimensions[hr].width = 11
    ws.sheet_view.zoomScale = 85


def build_lambdas_sheet(wb):
    """Sheet Lambdas (v1.08): the optional LAMBDA library. The workbook calculates with plain formulas only; for
    every calculated row this sheet gives the LAMBDA formula of the right column of scenario 1, which can replace the
    plain formulas of the row (paste it as a formula in the right column and copy it left over the row; the checks
    verify that the results are the same)."""
    ws = wb.create_sheet('Lambdas')
    title(ws, 'Lambdas - optional LAMBDA formulas by row (the workbook itself uses plain formulas only)', 6)
    put(ws, 'A3', f'The named LAMBDAs (Name Manager, {len(LAMBDAS)}) are an optional library: Excel 365 can use them, '
                  'other spreadsheets ignore them. Column F is the formula for the right column '
                  f'({L(group_cols(1)[2][-1])}) of scenario 1; type = before it to use it, and copy it over the row. '
                  'The base year keeps its own formula where it differs.', FONT)
    for j, h in enumerate(['Mitigation row', 'Variable', 'Fuel / member', 'Sector', 'LAMBDA', 'Formula (without =)'],
                          1):
        put(ws, f'{L(j)}4', h, FONT_B, F_INPUT)
    for i, (r, var, fu, sec, f) in enumerate(LAMBDA_ROWS, 5):
        m = re.match(r'=([A-Z0-9]+)\(', f)
        for j, v in enumerate([r, var, fu, sec, m.group(1) if m else '', f[1:]], 1):
            put(ws, f'{L(j)}{i}', v, FONT_LAMBDA if j == 6 else FONT)
    for col, w in {'A': 10, 'B': 12, 'C': 8, 'D': 8, 'E': 14, 'F': 120}.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A5'


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
                           'Param 1 (D)', 'Param 2 (E)', 'Param 3 (F)', 'Param 4 (G)',
                           'Code end (.suffix. or .)'], 1):
        put(ws, f'{L(j)}4', h, FONT_B, F_INPUT)
    for i, v in enumerate(VARIABLES, VR0):
        put(ws, f'K{i}', f'=IF($D{i}="",".","."&$D{i}&".")')
        for j, x in enumerate(v, 1):
            if x == '=MT':  # label / unit of an MTInputs parameter: looked up there by NameOfParameter
                src = 'B' if j == 2 else 'C'
                x = f'=INDEX(MTInputs!${src}$8:${src}$415,MATCH($A{i},MTInputs!$H$8:$H$415,0))&""'
            put(ws, f'{L(j)}{i}', x, FONT, F_INPUT if j < 6 and not str(x).startswith('=') else None)
    for col, w in {'A': 14, 'B': 46, 'C': 9, 'D': 9, 'E': 34, 'F': 60, 'G': 16, 'H': 16, 'I': 16, 'J': 16,
                   'K': 12}.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------- formula helpers
def group_cols(g):
    """(code column, base-year column, projection columns) of scenario group g."""
    code = COL_G0 + (g - 1) * GW
    return code, code + 1, list(range(code + 2, code + 1 + NY))


VAR_NAMES = {'B': 'VarLab', 'C': 'VarUnit', 'D': 'VarSfx', 'E': 'VarSrc', 'K': 'VarEnd'}   # Variables columns


def var_lookup(col, r):
    return f'INDEX({VAR_NAMES[col]},MATCH($A{r},VarKey,0))'


def label_formula(r):
    """Description: variable label (Variables) & member label ' | sector | fuel' (table on Mapping, key sector|fuel)."""
    return f'=INDEX(VarLab,MATCH($A{r},VarKey,0))&INDEX(MemLab,MATCH($C{r}&"|"&$B{r},MemKey,0))'


MEM = {}                                # v1.08: member-label table on Mapping: first and last row


def build_member_labels(wb):
    """Mapping: one row per sector|fuel pair used on Mitigation, with its label ' | sector | fuel' ('' without a
    sector), so that the descriptions in column H need no LAMBDA and stay short."""
    wm, ws = wb['Mitigation'], wb['Mapping']
    pairs = []
    for r in range(1, wm.max_row + 1):
        if isinstance(wm.cell(r, 8).value, str) and wm.cell(r, 8).value.startswith('=INDEX(VarLab'):
            p = (wm.cell(r, 3).value or '', wm.cell(r, 2).value or '')
            if p not in pairs:
                pairs.append(p)
    r0 = ws.max_row + 3
    put(ws, f'A{r0 - 2}', 'Member labels for Mitigation column H (key sector|fuel; add a row for a new pair)', FONT_B)
    for j, h in enumerate(['Key', 'Sector', 'Fuel', 'Label', 'Sector label', 'Fuel label'], 1):
        put(ws, f'{L(j)}{r0 - 1}', h, FONT_B, F_INPUT)
    for i, (sec, fu) in enumerate(pairs):
        r = r0 + i
        put(ws, f'A{r}', f'=$B{r}&"|"&$C{r}')
        put(ws, f'B{r}', sec, FONT, F_INPUT)
        put(ws, f'C{r}', fu, FONT, F_INPUT)
        put(ws, f'D{r}', f'=IF($B{r}="",""," | "&$E{r}&" | "&$F{r})')
        put(ws, f'E{r}', f'=IFERROR(INDEX(SecLab,MATCH($B{r},SecKey,0)),"")')
        put(ws, f'F{r}', f'=IFERROR(INDEX(FuelLab,MATCH($C{r},FuelKeyL,0)),"")')
    MEM['r0'], MEM['r1'] = r0, r0 + len(pairs) + 20         # room for 20 more pairs


def code_formula(r, code_col):
    """country.mit.<variable>.<sector>.<fuel>.<suffix>.<scenario>; CodePrefix = country.mit., VarEnd = .suffix."""
    return f'=CodePrefix&$A{r}&IF($C{r}="","","."&$C{r}&"."&$B{r})&{var_lookup("K", r)}&{code_col}${R_SCEN}'


def defined_names():
    """Defined names (v1.07) for the data tables and a few fixed ranges, so that cell formulas stay short."""
    v = lambda c: f'Variables!${c}${VR0}:${c}${VR1}'
    mit = lambda c: f'Mitigation!${c}${S_FIRST}:${c}${S_LAST}'
    obr0 = R_POL['obr.pow']
    sr = lambda c: f'StoredResults!${c}${SR_FIRST}:${c}${SR_LAST}'
    yl = L(SR_Y0 + NY - 1)
    return {
        'WEOVal': 'WEO!$F$4:$AZ$1000', 'WEOKey': 'WEO!$A$4:$A$1000', 'WEOYear': 'WEO!$F$3:$AZ$3',
        'ElaVal': 'Elasticities!$F$4:$I$200', 'ElaKey': 'Elasticities!$A$4:$A$200', 'ElaGroup': 'Elasticities!$F$3:$I$3',
        'EFVal': 'EF_GHG!$E$4:$E$1000', 'EFKey': 'EF_GHG!$A$4:$A$1000',
        'ConsVal': 'EnergyCons!$F$4:$F$5000', 'ConsKey': 'EnergyCons!$A$4:$A$5000',
        'PIntVal': 'Prices_int!$F$4:$AD$100', 'PIntKey': 'Prices_int!$A$4:$A$100', 'PIntYear': 'Prices_int!$F$3:$AD$3',
        'PDomVal': 'Prices_dom!$A$6:$CU$2000', 'PDomKey': 'Prices_dom!$A$6:$A$2000', 'PDomHdr': 'Prices_dom!$A$5:$CU$5',
        'PPathVal': f'PowerPaths!${PP_COLS[0]}$4:${PP_COLS[1]}$200', 'PPathKey': 'PowerPaths!$A$4:$A$200',
        'PPathYear': f'PowerPaths!${PP_COLS[0]}$3:${PP_COLS[1]}$3',
        'PTechVal': 'PowerTech!$A$4:$Z$50', 'PTechKey': 'PowerTech!$A$4:$A$50', 'PTechHdr': 'PowerTech!$A$3:$Z$3',
        'PParVal': 'PowerParams!$B$4:$B$50', 'PParKey': 'PowerParams!$A$4:$A$50',
        'VarKey': v('A'), 'VarLab': v('B'), 'VarUnit': v('C'), 'VarSfx': v('D'), 'VarSrc': v('E'), 'VarEnd': v('K'),
        'SubKey': f'Mapping!$A${MAP_SUB0}:$A${MAP_SUB1}', 'FuelKey': f'Mapping!$A${MAP_FUEL0}:$A${MAP_FUEL1}',
        'SecKey': f'Mapping!$A${MAP_SUB0}:$A${MAP_SUBLAB1}', 'SecLab': f'Mapping!$B${MAP_SUB0}:$B${MAP_SUBLAB1}',
        'FuelKeyL': f'Mapping!$A${MAP_FUEL0}:$A${MAP_FUELLAB1}', 'FuelLab': f'Mapping!$B${MAP_FUEL0}:$B${MAP_FUELLAB1}',
        'MemKey': f'Mapping!$A${MEM["r0"]}:$A${MEM["r1"]}', 'MemLab': f'Mapping!$D${MEM["r0"]}:$D${MEM["r1"]}',
        'SecVar': mit('A'), 'SecFuel': mit('B'), 'SecSub': mit('C'),
        'SemiEl': f'Mitigation!$E${obr0}:$E${obr0 + len(COV_SECTORS) - 1}',
        'MitAll': f'Mitigation!{MIT_RANGE}', 'MitScen': 'Mitigation!$A$5:$ZZ$5',
        'MitDesc': 'Mitigation!$H$1:$H$5000', 'MitUnit': 'Mitigation!$I$1:$I$5000',
        'SRScen': sr('A'), 'SRName': sr('B'), 'SRCode': sr('C'), 'SRDesc': sr('E'), 'SRUnit': sr('F'),
        'SRVal': f'StoredResults!${L(SR_Y0)}${SR_FIRST}:${yl}${SR_LAST}',
        'SRYear': f'StoredResults!${L(SR_Y0)}${SR_HEAD}:${yl}${SR_HEAD}',
        'CodePrefix': 'Settings!$C$20', 'CPIRes': 'Settings!$C$22', 'DeflRes': 'Settings!$C$23',
        'PDomBase': 'Settings!$C$24', 'PDomLast': 'Settings!$C$25', 'OBREff': 'Settings!$C$14',
        # v1.09: existing carbon pricing (sheet ECP): selected country, all countries
        'ECPKey': f'ECP!$A${ECP_R0}:$A${ECP_R1}', 'ECPCov': f'ECP!$D${ECP_R0}:$D${ECP_R1}',
        'ECPVal': f'ECP!$E${ECP_R0}:${ECP_YL}${ECP_R1}', 'ECPYear': f'ECP!$E${ECP_HDR}:${ECP_YL}${ECP_HDR}',
        'ECPLast': f'ECP!${ECP_YL}${ECP_HDR}', 'ECPAllKey': f'ECP!$A${ECP_A0}:$A${ECP_A0 + 2000}',
        'ECPAllVal': f'ECP!$D${ECP_A0}:${ECP_AL}${ECP_A0 + 2000}',
        'ECPAllYear': f'ECP!$D${ECP_A0 - 1}:${ECP_AL}${ECP_A0 - 1}',
        'CtryCode': 'Settings!$C$4'}


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
    if var == 'ctxnew':     # carbon tax rate of the sector x EF x fuel coverage
        a = (f'INDEX({pol_range(col, "ctr.pow", len(COV_SECTORS))},$F{r})', f'$D{r}',
             f'INDEX({pol_range(col, "fc.coa", NF)},$E{r})')
        return '=' + (call('CARBONTAX', *a) if use_lambda else f'{a[0]}*{a[1]}*{a[2]}')
    if var == 'ets':        # ETS cost rate of the sector x EF
        a = (f'INDEX({pol_range(col, "etr.pow", len(COV_SECTORS))},$E{r})', f'$D{r}')
        return '=' + (call('ETSCOST', *a) if use_lambda else f'{a[0]}*{a[1]}')
    if var == 'ntx':
        a = (f'INDEX({pol_range(col, "fpr." + PRICE_FUELS[0][0], len(PRICE_FUELS))},$D{r})', f'$E{r}')
        return '=' + (call('NEWEXCISE', *a) if use_lambda else '/'.join(a))
    if var == 'nce':
        a = (o('ctxnew'), o('ets'), o('ntx'))
        return '=' + (call('NEWPOLICY', *a) if use_lambda else '+'.join(a))
    if var in ('xcp', 'xcpd'):  # v1.09: existing carbon pricing: EF x (coverage x price), level or change
        p = lambda k: f'{col}${R_POL[k]}'
        ct, et = ((p('xct.p'), p('xets.p')) if var == 'xcp' else
                  (f'({p("xct.p")}-{p("xct.p0")})', f'({p("xets.p")}-{p("xets.p0")})'))
        a = (f'$D{r}', f'$E{r}', ct, f'$F{r}', et)
        return '=' + (call('XCARBON', *a) if use_lambda else f'{a[0]}*({a[1]}*{a[2]}+{a[3]}*{a[4]})')
    if var == 'atp':
        rpb = f'INDEX({col}${R_PV0 + PVOFF["rpb"]}:{col}${R_PV0 + PVOFF["rpb"] + NP - 1},$D{r})'
        pol = f'{o("nce")}+{o("xcpd")}'
        return (f'={call("POSTTAX", rpb, pol, f"$E{r}")}' if use_lambda
                else f'=MAX({rpb}+({pol})*(1+$E{r}),0.01)')
    if var == 'shp':        # shadow price of the sector ($/tCO2, section 1) x EF
        a = (f'INDEX({pol_range(col, "shpx.pow", len(COV_SECTORS))},$E{r})', f'$D{r}')
        return '=' + (call('SHADOWGJ', *a) if use_lambda else f'{a[0]}*{a[1]}')
    if var == 'co2':
        a = (o('ener'), 'Settings!$C$13', f'$D{r}')
        return '=' + (call('EMISSIONS', *a) if use_lambda else '*'.join(a))
    if var in REV_VARS:
        pr = lambda v: f'INDEX({col}${R_PV0 + PVOFF[v]}:{col}${R_PV0 + PVOFF[v] + NP - 1},$D{r})'
        k = 'Settings!$C$13'
        if var == 'rtx':                     # v1.09: + change in existing carbon pricing (incl. VAT)
            rate = f'{pr("etx")}+{pr("vat")}+{o("xcpd")}*(1+$E{r})'
        elif var == 'rxcp':                  # v1.09: existing carbon pricing revenue (memo, level)
            rate = o('xcp')
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
        return f'={prev}{r}*$G{r}*(1+{col}${R_GDP})^$D{r}*{ratio}^$E{r}*{eff}^$F{r}'
    raise ValueError(var)


LH = 'Settings!$C$10'                   # last historical price year


def gp_range(col):
    return f'{col}${R_GP0}:{col}${R_GP0 + len(GP_ROWS) - 1}'


def gp_formula(r, col, var='gp'):
    """International prices (v1.07 in steps): gp.src = source key (MTInputs selector without the *); gp.adj = 1 (no
    adjustment), 2 (High) or 3 (Low); gpn = Prices_int value of the source (nominal; 1 without commodity key);
    gp = gpn x deflator index x High/Low factor after the last historical year (1 without commodity key)."""
    src, adj = f'{col}${R_PRI[PRI_PARAMS[0]]}', f'{col}${R_PRI[PRI_PARAMS[1]]}'
    if var == 'gp.src':
        return f'=SUBSTITUTE({src},"*","")'
    if var == 'gp.adj':
        return f'=IF(LEFT({adj},4)="High",2,IF(LEFT({adj},3)="Low",3,1))'
    if var == 'gpn':
        return (f'=IF($D{r}="",1,INDEX(PIntVal,MATCH({col}${R_PRI["gp.src"]}&"|"&$D{r},PIntKey,0),'
                f'MATCH({col}${R_YEAR},PIntYear,0)))')
    gpn = f'{col}{r - R_GP0 + R_GPN0}'
    return (f'=IF($D{r}="",1,{gpn}*{col}${R_DEFL}*IF({col}${R_YEAR}>{LH},CHOOSE({col}${R_PRI["gp.adj"]},1,$E{r},'
            f'$F{r}),1))')


def price_formula(var, r, col, prev, block):
    """Section-2 cell of price variable var in row r. block: 'hist' (years up to the last historical price year,
    base year included: data from Inputs_prices by price position), 'proj' (projected years) or 'lam' (right
    column: the named LAMBDA, which holds the IF between history and projection)."""
    o = lambda v: f'{col}{r - PVOFF[var] + PVOFF[v]}'
    hist = lambda v, pos: (f'INDEX(Inputs_prices!${IPC[v + str(HIST_YEARS[0])]}${IP_R0}:'
                           f'${IPC[v + str(HIST_YEARS[-1])]}${IP_R1},{pos},MATCH({col}${R_YEAR},'
                           f'Inputs_prices!${IPC[v + str(HIST_YEARS[0])]}${IP_HDR}:'
                           f'${IPC[v + str(HIST_YEARS[-1])]}${IP_HDR},0))')
    lam = block == 'lam'
    if var == 'sp':
        gp_now, gp_prev = f'INDEX({gp_range(col)},$E{r})', f'INDEX({gp_range(prev)},$E{r})'
        if block == 'hist':
            return '=' + hist('sp', f'$F{r}')
        if lam:
            return '=' + call('SUPPLYCOST', f'{col}${R_YEAR}', f'$F{r}', f'$D{r}', f'{prev}{r}', gp_now, gp_prev)
        return f'=$D{r}+({prev}{r}-$D{r})*{gp_now}/{gp_prev}'
    if var == 'txo':
        if block == 'hist':
            return '=' + hist('txo', f'$G{r}')
        if lam:
            return '=' + call('OTHERTAX', f'{col}${R_YEAR}', f'$G{r}', f'$D{r}', f'$E{r}', o('sp'), f'$F{r}')
        return f'=$D{r}*$F{r}+IF($D{r}*(1-$F{r})>=0,MAX($D{r}*(1-$F{r}),{o("txf")}),{o("txf")})'
    if var == 'txf':
        a = (f'$D{r}', f'$E{r}', o('sp'), f'$F{r}')
        return '=' + (call('FLOATTAX', *a) if lam else f'({a[1]}-{a[2]}+{a[0]}*(1-{a[3]}))*(1-{a[3]})')
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


def pw_block(var, col, cols_abs=False):
    """Absolute-row range of a power variable over its members (types, groups, subsectors) in column col."""
    mem = PW_MEMBERS[PW_KIND[var]]
    r0 = R_PW[(var, mem[0])]
    return f'{col}${r0}:{col}${r0 + len(mem) - 1}'


def pw_pcol(var, pc):
    """Absolute range of a power variable's parameter column pc (e.g. $D$100:$D$108)."""
    mem = PW_MEMBERS[PW_KIND[var]]
    r0 = R_PW[(var, mem[0])]
    return f'${pc}${r0}:${pc}${r0 + len(mem) - 1}'


def ipw_param(key):
    return f'Inputs_power!$B${IPWQ_ROW[key]}'


def pw_param_formula(spec, var, r):
    """Mitigation D:G cell of a power row: data key, Inputs_power lookup, PowerParams value or constant."""
    if spec[0] == 'key':
        kind = PW_KIND[var]
        return f'=$A{r}&"|"&$B{r}' if kind == 'tech' else (f'=$A{r}&"|"&$C{r}' if kind == 'user' else f'=$A{r}')
    if spec[0] == 't':
        c = IPWC[spec[1]]
        return f'=INDEX(Inputs_power!${c}${IPW_R0}:${c}${IPW_R1},MATCH($B{r},Inputs_power!$A${IPW_R0}:$A${IPW_R1},0))'
    if spec[0] in ('g', 'h', 's'):
        tab = {'g': (IPWG, IPW_G0, IPW_G1), 'h': (IPWH, IPWH0, IPWH1), 's': (IPWS, IPWS0, IPWS1)}[spec[0]]
        c = tab[0][spec[1]]
        return (f'=INDEX(Inputs_power!${c}${tab[1]}:${c}${tab[2]},MATCH($C{r},Inputs_power!$A${tab[1]}:$A${tab[2]},'
                f'0))')
    if spec[0] == 'p':
        return '=' + pw_param(spec[1])
    if spec[0] == 'q':
        return '=' + ipw_param(spec[1])
    return spec[1]


def pw_formula(var, r, col, prev, block):
    """Section-3 cell of power variable var in row r. block: 'all' (data rows, every column), 'base', 'plain',
    'hist' / 'proj' (rppre) or 'lam' (right column: the named LAMBDA)."""
    lam = block == 'lam'
    nxt = L(CI_(col) + 1)
    kind = PW_KIND[var]
    first = PW_MEMBERS[kind][0] if kind != 'single' else None
    o = lambda v, c=col: f'{c}{r - R_PW[(var, first)] + R_PW[(v, first)]}'      # same member, other variable
    one = lambda v, c=col: f'{c}${R_PW[(v, "")]}'                                 # single row
    mem = lambda v, m, c=col: f'{c}${R_PW[(v, m)]}'                              # a given member (absolute row)
    blk = lambda v, c=col: pw_block(v, c)
    y = f'{col}${R_YEAR}'
    cg = pw_pcol('gen', 'D')                                                      # coal / gas flags (gen rows)

    def fn(name, *a, plain):
        return '=' + (call(name, *a) if lam else plain)
    if var in PW_DATA:
        return f'=INDEX(PPathVal,MATCH($D{r},PPathKey,0),MATCH({col}${R_YEAR},PPathYear,0))'
    if var == 'demg':       # memo: demand of the subsectors of this demand group (dem rows: D = group position)
        return f'=SUMIF({pw_pcol("dem", "D")},$D{r},{blk("dem")})'
    if var in PW_SUMS:
        src = {'demtot': 'dem', 'gtot': 'gen', 'usdtot': 'usd', 'captot': 'cap', 'nimtot': 'nim'}[var]
        return f'=SUM({blk(src)})'
    if var == 'cax':
        return fn('CAPEX', f'$D{r}', o('tcf'), plain=f'$D{r}*{o("tcf")}')
    if var in ('caxav', 'stoav'):
        if block == 'base':
            return '=' + (o('cax') if var == 'caxav' else f'$D{r}*{one("msc")}')
        new_v = o('cax', prev) if var == 'caxav' else f'$D{r}*{one("msc")}'      # capex of last year; storage now
        a = (f'{prev}{r}', new_v, o('phi', prev))
        return fn('VINTAGE', *a, plain=f'{a[0]}*(1-{a[2]})+{a[1]}*{a[2]}')
    if var == 'fix':
        a = (o('caxav'), f'$D{r}', f'$E{r}', o('stoav'))
        return fn('AMORTISED', *a, plain=f'{a[0]}*{a[1]}+{a[2]}+{a[3]}')
    if var == 'lfx':
        a = (o('cax'), f'$D{r}', f'$E{r}', f'$F{r}', one('msc'))
        return fn('LEVELISED', *a, plain=f'{a[0]}*{a[1]}+{a[2]}+{a[3]}*{a[4]}')
    if var == 'vbc':        # v1.09: + change in existing carbon pricing (xck)
        rpb = f'INDEX({col}${R_PV0 + PVOFF["rpb"]}:{col}${R_PV0 + PVOFF["rpb"] + NP - 1},$E{r})'
        return fn('VARCOST', f'$D{r}', rpb, f'$F{r}', o('xck'), plain=f'$D{r}+{rpb}*$F{r}+{o("xck")}')
    if var in ('xck', 'prxl'):   # v1.09: existing carbon pricing: change per kWh (xck), revenue level (prxl)
        p = lambda k: f'{col}${R_POL[k]}'
        if var == 'xck':
            a = (f'$D{r}', f'$E{r}', f'({p("xct.p")}-{p("xct.p0")})', f'$F{r}', f'({p("xets.p")}-{p("xets.p0")})')
        else:
            a = (o('co2'), f'$D{r}', p('xct.p'), f'$E{r}', p('xets.p'))
        return fn('XCARBON', *a, plain=f'{a[0]}*({a[1]}*{a[2]}+{a[3]}*{a[4]})')
    if var in ('prtx', 'prsub'):   # v1.09: existing taxes / subsidies on generation fuels (price fuel of section 2)
        pr = lambda v: f'INDEX({col}${R_PV0 + PVOFF[v]}:{col}${R_PV0 + PVOFF[v] + NP - 1},$D{r})'
        rate = f'{pr("etx")}+{pr("vat")}' if var == 'prtx' else pr('esub')
        a = (f'$E{r}*{o("pe")}', 'Settings!$C$13', rate)
        return fn('REVENUE', *a, plain=f'{a[0]}*{a[1]}*({a[2]})')
    if var == 'prxc':       # change in existing carbon pricing revenue: generation (GWh) x xck
        return fn('POWERREV', o('xck'), '8.76', o('gen'), plain=f'{o("xck")}*8.76*{o("gen")}')
    if var == 'prnew':      # new policies: cost per kWh x GWh, ETS part at the permit price (x revenue factor)
        a = (o('gen'), o('ccp'), o('co2'), f'{col}${R_POL["etr.pow"]}', f'{col}${R_POL["ets.rf"]}')
        return fn('PNEWREV', *a, plain=f'{a[0]}*8.76*{a[1]}-{a[2]}*{a[3]}*(1-{a[4]})')
    if var == 'ccp':        # tCO2/kWh x (carbon tax x fuel coverage + ETS cost) + fuel price reform per kWh
        a = (f'$D{r}', f'{col}${R_POL["ctr.pow"]}', f'INDEX({pol_range(col, "fc.coa", NF)},$E{r})',
             f'{col}${R_POL["etr.pow"]}',
             f'INDEX({pol_range(col, "fpr." + PRICE_FUELS[0][0], len(PRICE_FUELS))},$F{r})', f'$G{r}')
        return fn('KWHPOLICY', *a, plain=f'{a[0]}*({a[1]}*{a[2]}+{a[3]})+{a[4]}*{a[5]}')
    if var == 'shk':        # shadow price per kWh: shadow price of power ($/tCO2, section 1) x tCO2/kWh
        a = (f'{col}${R_POL["shpx.pow"]}', f'$D{r}')
        return fn('SHADOWGJ', *a, plain=f'{a[0]}*{a[1]}')
    if var == 'reb':        # feebate rebate per kWh generated (base-year shares), from the second year
        if block == 'base':
            return '=0'
        a = (blk('shk'), pw_pcol('shk', 'E'))
        return fn('REBATE', *a, plain=f'-SUMPRODUCT({a[0]},{a[1]})')
    if var == 'tvc':
        a = (o('vbc'), o('ccp'), o('shk'), one('reb'))
        return fn('VARTOTAL', *a, plain='+'.join(a))
    if var == 'dsum':       # sum of discount factors to the end of the horizon (next column: blank after the last)
        a = (f'$D{r}', 'Settings!$C$7', y, f'{nxt}{r}')
        return fn('DISCSUM', *a, plain=f'(1+{a[0]})^({a[1]}-{a[2]})+{a[3]}')
    if var == 'lvc':        # levelised variable cost: discounted average of tvc to the end of the horizon (D9)
        a = (f'$D{r}', 'Settings!$C$7', y, o('tvc'), o('dsum'), f'{nxt}{r}')
        df = f'(1+{a[0]})^({a[1]}-{a[2]})'
        return fn('LEVVAR', *a, plain=f'({df}*{a[3]}+({a[4]}-{df})*{a[5]})/{a[4]}')
    if var == 'gnc':
        return fn('GENCOST', o('fix'), o('vbc'), plain=f'{o("fix")}+{o("vbc")}')
    if var in ('gncav', 'ccpav', 'vre'):
        vals = {'gncav': blk('gnc'), 'ccpav': blk('ccp'), 'vre': pw_pcol('shr', 'D')}[var]
        a = (blk('shr'), vals, f'$D{r}')
        return fn('WEIGHTED', *a, plain=f'SUMPRODUCT({a[0]},{a[1]})*{a[2]}')
    if var == 'tic':
        a = (o('lfx'), o('lvc'), f'$D{r}')
        return fn('INVCOST', *a, plain=f'MAX({a[0]}+{a[1]},{a[2]})')
    if var == 'cph':
        a = (one('cbat'), one('cint'), one('obat'), f'$D{r}', f'$E{r}')
        return fn('STORAGEHOUR', *a, plain=f'({a[0]}+{a[1]}/{a[3]})/({a[4]}*8760)+{a[2]}/8760')
    if var == 'msc':
        a = (f'$D{r}', f'$E{r}', one('vre'), one('cph'), f'$F{r}', one('cel'), f'$G{r}')
        return fn('STORAGECOST', *a, plain=(f'{a[0]}*{a[1]}*{a[2]}*{a[3]}+2*MAX(0,{a[2]}-{a[4]})/(1-{a[4]})^2*{a[5]}'
                                            f'*{a[6]}'))
    if var == 'cons':
        a = (one('demtot'), mem('dem', 'res'), f'$D{r}', f'$E{r}', f'$F{r}')
        return fn('CONSUMPTION', *a, plain=f'({a[2]}*{a[0]}+{a[3]}*{a[1]})*{a[4]}')
    if var == 'sc':
        gen = one('gncav') if block == 'base' else one('gncav', prev)
        return fn('POWERSUPPLY', gen, f'$D{r}', plain=f'{gen}+$D{r}')
    if var == 'rppre':
        hist = (f'INDEX(Inputs_power!${IPWG["rp" + str(HIST_YEARS[0])]}${IPW_G0}:${IPWG["rp" + str(HIST_YEARS[-1])]}'
                f'${IPW_G1},$E{r},MATCH({col}${R_YEAR},Inputs_power!${IPWG["rp" + str(HIST_YEARS[0])]}${IPW_G0 - 1}:'
                f'${IPWG["rp" + str(HIST_YEARS[-1])]}${IPW_G0 - 1},0))')
        if block == 'hist':
            return '=' + hist
        if lam:
            return '=' + call('POWERPRICE', f'{col}${R_YEAR}', f'$E{r}', f'{prev}{r}', f'$D{r}', o('sc'),
                              o('sc', prev), f'$F{r}')
        return f'={prev}{r}+$D{r}*({o("sc")}-{o("sc", prev)})*(1+$F{r})'
    if var == 'rp':         # policy term x pass-through x (1 + VAT), as legacy row 3614
        a = (o('rppre'), one('ccpav'), f'$D{r}', f'$E{r}')
        return fn('ENDUSERPOWER', *a, plain=f'{a[0]}+{a[1]}*{a[2]}*(1+{a[3]})')
    if var == 'sgap':
        return fn('SUBSIDYGAP', o('sc'), o('rppre'), f'$D{r}', plain=f'{o("sc")}-{o("rppre")}/(1+$D{r})')
    if var in ('rvat', 'rcarb', 'rgap'):
        rate = {'rvat': o('rppre'), 'rcarb': one('ccpav'), 'rgap': o('sgap')}[var]
        return fn('POWERREV', rate, f'$D{r}', o('cons'), plain=f'{rate}*$D{r}*{o("cons")}')
    if var == 'rnvat':      # v1.09: VAT on the new-policy cost passed through (pass-through x VAT rate)
        rate = f'{one("ccpav")}*$D{r}'
        return fn('POWERREV', rate, f'$E{r}', o('cons'), plain=f'{rate}*$E{r}*{o("cons")}')
    # ---------------- H demand
    if var == 'pdem':       # price of the previous year, weighted by demand group (base year: 2021 data)
        if block == 'base':
            return f'=$E{r}'
        a = (f'$D{r}', mem('rp', 'res', prev), mem('rp', 'ind', prev))
        return fn('PRICEMIX', *a, plain=f'{a[0]}*{a[1]}+(1-{a[0]})*{a[2]}')
    if var == 'mult':
        if block == 'base':
            return '=1'
        a = (f'$D{r}', f'{col}${R_GDP}', f'$E{r}', one('incf'), o('pdem'), o('pdem', prev), f'$F{r}', one('shock'),
             'Settings!$C$19')
        return fn('DEMANDMULT', *a, plain=f'{a[0]}*(1+{a[1]})^({a[2]}*{a[3]})*({a[4]}/{a[5]})^{a[6]}*(1+{a[7]}*{a[8]})')
    if var == 'dem':
        if block == 'base':
            return f'=$E{r}'
        a = (f'{prev}{r}', f'INDEX({blk("mult")},$D{r})')
        return fn('GROW', *a, plain=f'{a[0]}*{a[1]}')
    if var == 'req':
        if block == 'base':
            return f'=SUM({blk("gen")})'
        return fn('SCALE', one('demtot'), f'$D{r}', plain=f'{one("demtot")}*$D{r}')
    # ---------------- I dispatch
    if var == 'reqcg':      # generation required from coal and gas
        a = (one('req'), f'SUMPRODUCT({blk("effcap")},1-{cg})')
        return fn('RESIDUAL', *a, plain=f'MAX({a[0]}-{a[1]},0)')
    if var == 'vcmin':
        a = (blk('tvc'), cg, f'$D{r}')
        return fn('FLOORMIN', *a, plain=f'MAX(_xlfn.MINIFS({a[0]},{a[1]},1),{a[2]})')
    if var == 'rem':
        a = (one('reqcg'), f'SUM({blk("gmin")})')
        return fn('RESIDUAL', *a, plain=f'MAX({a[0]}-{a[1]},0)')
    if var == 'gmin':       # minimum of coal (gas) if the other runs at its maximum
        a = (f'$D{r}', one('reqcg'), f'SUMPRODUCT({blk("cap")},{cg})*$E{r}', f'{o("cap")}*$E{r}')
        return fn('MINDISPATCH', *a, plain=f'{a[0]}*MIN(MAX({a[1]}-({a[2]}-{a[3]}),0),{a[3]})')
    if var == 'lgd':
        a = (f'$D{r}', f'$E{r}', f'MAX({o("tvc")},$F{r})', one('vcmin'))
        return fn('LOGITW', *a, plain=f'{a[0]}*EXP(-{a[1]}*{a[2]}/{a[3]})')
    if var == 'gen':
        if block == 'base':
            return f'=$F{r}'
        a = (f'$D{r}', o('effcap'), f'{o("cap")}*$E{r}', o('gmin'), one('rem'), o('lgd'), f'SUM({blk("lgd")})')
        return fn('DISPATCH', *a, plain=f'IF({a[0]}=1,MIN({a[2]},{a[3]}+{a[4]}*{a[5]}/{a[6]}),{a[1]})')
    if var == 'dgen':
        if block == 'base':
            return '=0'
        return fn('INCREASE', one('gtot'), one('gtot', prev), plain=f'MAX(0,{one("gtot")}-{one("gtot", prev)})')
    if var == 'shr':
        return fn('SHARE', o('gen'), one('gtot'), plain=f'IF({one("gtot")}=0,0,{o("gen")}/{one("gtot")})')
    if var in ('pe', 'co2', 'effcap', 'reteff', 'nim'):
        src = {'pe': 'gen', 'co2': 'pe', 'effcap': 'cap', 'reteff': 'ret', 'nim': 'nei'}[var]
        return fn('SCALE', o(src), f'$D{r}', plain=f'{o(src)}*$D{r}')
    # ---------------- K capacity and retirements
    if var == 'cap':
        if block == 'base':
            return f'=$D{r}'
        a = (o('cap', prev), o('ret', prev), o('nim', prev))
        return fn('CARRY', *a, plain=f'{a[0]}-{a[1]}+{a[2]}')
    if var == 'ret':        # planned retirement (data, where switched on) or rate-based; + early coal retirement
        a = (o('cap'), f'$D{r}', f'$E{r}', one('early'), f'$F{r}', o('pret'))
        return fn('RETIRE', *a, plain=f'MIN(IF({a[4]}=1,{a[5]},{a[0]}*{a[1]})+{a[2]}*{a[3]},{a[0]})')
    if var == 'egrow':      # room for early coal retirement: max wind and solar additions less demand growth and
        add = (mem('maxinv', 'wnd'), mem('maxinv', 'sol'))     # (effective) non-coal retirements of last year
        if block == 'base':
            return '=' + (call('EARLYROOM', *add, '0', '0') if lam else f'MAX(0,{add[0]}+{add[1]})')
        a = add + (one('dgen', prev), f'SUM({blk("reteff", prev)})-{mem("reteff", "coa", prev)}')
        return fn('EARLYROOM', *a, plain=f'MAX(0,{a[0]}+{a[1]}-{a[2]}-({a[3]}))')
    cc, cw, cs = mem('tvc', 'coa'), mem('tic', 'wnd'), mem('tic', 'sol')     # early retirement: coal vs wind, solar
    if var == 'ecmin':      # lowest of coal variable cost and wind and solar investment cost
        return fn('LOWEST3', cc, cw, cs, plain=f'MIN({cc},{cw},{cs})')
    if var == 'ecshr':      # coal share of the logit (= e_coal / (e_coal + e_wind + e_solar), divided through)
        a = (f'$D{r}', cc, cw, cs, one('ecmin'))
        return fn('COALSHARE', *a, plain=(f'1/(1+EXP(-{a[0]}*({cw}-{cc})/{a[4]})'
                                          f'+EXP(-{a[0]}*({cs}-{cc})/{a[4]}))'))
    if var == 'early':      # cost-based early coal retirement (nameplate MW)
        a = (mem('effcap', 'coa'), f'$D{r}', one('egrow'), one('ecshr'), f'$E{r}', f'$F{r}')
        return fn('EARLYRET', *a, plain=f'MAX(0,MIN({a[0]}*{a[1]},{a[2]})*(1-{a[3]}))/{a[5]}*{a[4]}')
    if var == 'kafter':     # effective capacity after this year's retirements
        a = (f'SUM({blk("effcap")})', f'SUM({blk("reteff")})')
        return fn('DIFF', *a, plain=f'{a[0]}-{a[1]}')
    if var == 'need':       # effective capacity needed: this year's required generation, from next year
        return fn('RESIDUAL', one('req'), one('kafter'), plain=f'MAX({one("req")}-{one("kafter")},0)')
    # ---------------- L investment
    if var == 'cmin':
        a = (blk('tic'), blk('permit'), f'$D{r}')
        return fn('FLOORMIN', *a, plain=f'MAX(_xlfn.MINIFS({a[0]},{a[1]},1),{a[2]})')
    if var == 'deficit':
        a = (one('need'), f'SUM({blk("first")})')
        return fn('RESIDUAL', *a, plain=f'MAX({a[0]}-{a[1]},0)')
    if var == 'permit':
        return fn('PERMITTED', f'$D{r}', f'$E{r}', y, plain=f'IF(AND({f"$D{r}"}="Yes",{y}>=$E{r}),1,0)')
    if var == 'maxinv':
        a = (f'$D{r}', f'SUM({blk("effcap")})', o('permit'))
        return fn('SCALEUP', *a, plain=f'{a[0]}*{a[1]}*{a[2]}')
    if var == 'lgi':
        a = (o('permit'), f'$D{r}', o('tic'), one('cmin'))
        return fn('LOGITW', *a, plain=f'{a[0]}*EXP(-{a[1]}*{a[2]}/{a[3]})')
    if var == 'pni':
        a = (o('lgi'), f'SUM({blk("lgi")})')
        return fn('SHARE', *a, plain=f'IF({a[1]}=0,0,{a[0]}/{a[1]})')
    if var == 'first':
        a = (o('maxinv'), one('need'), o('pni'))
        return fn('FIRSTROUND', *a, plain=f'MIN({a[0]},{a[1]}*{a[2]})')
    if var == 'space':
        a = (o('maxinv'), o('first'), o('permit'))
        return fn('SPACE', *a, plain=f'({a[0]}-{a[1]})*{a[2]}')
    if var == 'second':     # cheapest first: the deficit less the space of cheaper types (legacy L11-L17)
        cheaper = (f'SUMPRODUCT(({blk("tic")}+{pw_pcol("second", "D")}<{o("tic")}+$D{r})*{blk("space")})')
        a = (o('permit'), o('space'), one('deficit'), cheaper)
        return fn('CHEAPFILL', *a, plain=f'{a[0]}*MAX(0,MIN({a[1]},{a[2]}-{a[3]}))')
    if var == 'nei':
        return fn('ADDED', o('first'), o('second'), plain=f'{o("first")}+{o("second")}')
    if var == 'usd':
        return fn('INVSPEND', o('nim'), o('cax'), plain=f'{o("nim")}*{o("cax")}/1000')
    if var == 'phi':
        a = (o('nim'), o('cap'), o('ret'))
        return fn('VINTSHARE', *a, plain=f'IF({a[1]}+{a[0]}-{a[2]}=0,0,{a[0]}/({a[1]}+{a[0]}-{a[2]}))')
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
        if use_lambda:
            return '=' + call('PATH', y, c('CPIntro'), c('CPOutro'), c('CPLevelStart'), c('CPLevelTarget'),
                              f'{c("ExtendCarbonPriceBeyondOutro")}="Linear*"') + '*' + c('cp.nom')
        y0, s0, s1 = c('CPIntro'), c('CPLevelStart'), c('CPLevelTarget')
        return f'=IF({y}<{y0},0,{s0}+({s1}-{s0})/MAX({c("CPOutro")}-{y0},1)*{c("cp.t")})*{c("cp.nom")}'
    if key == 'cp.t':                        # years along the path (linear after the target year if Linear*)
        a = (y, c('CPIntro'), c('CPOutro'), c('ExtendCarbonPriceBeyondOutro'))
        if use_lambda:
            return '=' + call('PATHYEARS', a[0], a[1], a[2], f'{a[3]}="Linear*"')
        return f'=IF({a[3]}="Linear*",{y},MIN({y},{a[2]}))-{a[1]}'
    if key == 'cp.nom':                      # nominal inputs: x CPI index; real: 1
        a = (c('NomorReal'), f'{col}${R_CPI}')
        return '=' + (call('NOMFACTOR', *a) if use_lambda else f'IF(LEFT({a[0]},7)="Nominal",{a[1]},1)')
    if key.startswith('fpr.') or (key.startswith('fb.') and key not in ('fb.yr0', 'fb.yr1')):
        pre, k = ('fpr', key[4:]) if key.startswith('fpr.') else ('fb', key[3:])
        y0, y1 = c(f'{pre}.yr0'), c(f'{pre}.yr1')
        s0, s1 = (rel('fprs.' + k), rel('fprf.' + k)) if pre == 'fpr' else (rel('fbs.' + k), rel('fbt.' + k))
        if use_lambda:
            return '=' + call('PATH', y, y0, y1, s0, s1, 'TRUE')
        return path_formula(col, y0, y1, s0, s1)
    on = f'{c("ets.on")}=1'
    bco2 = pol_range(col, 'bco2.pow', len(COV_SECTORS))
    etsc = pol_range(col, 'etsc.pow', len(COV_SECTORS))
    if key == 'ets.on':                      # the new ETS applies and has started
        a = (c('D_NewETS'), y, c('D_ETSIntro'))
        return '=' + (call('ETSON', *a) if use_lambda else f'IF(AND(LEFT({a[0]},3)="Yes",{y}>={a[2]}),1,0)')
    if key == 'ets.p':
        a = (c('ets.on'), c('D_ETSPriceOverride'), c('ets.ovr'), c('ets.est'), c('cptraj'))
        return '=' + (call('ETSPRICE', *a) if use_lambda else
                      f'IF({on},IF(LEFT({a[1]},3)="Yes",IF({a[2]}="",{a[4]},{a[2]}),{a[3]}),0)')
    if key.startswith('etsb.'):              # benchmark: linear from start to target year, flat after
        y0, y1, s0, s1 = c('D_ETSIntro'), c('D_ETSOutro'), rel('etsbs.' + key[5:]), rel('etsbt.' + key[5:])
        if use_lambda:
            return '=' + call('PATH', y, y0, y1, s0, s1, 'FALSE')
        return f'=IF({y}<{y0},0,{s0}+({s1}-{s0})/MAX({y1}-{y0},1)*(MIN({y},{y1})-{y0}))'
    if key.startswith('obr.'):
        b = f'INDEX({pol_range(col, "etsb.pow", len(ETS_GROUPS))},$D{R_POL[key]})'
        return '=' + (call('OBRSHARE', b) if use_lambda else f'MIN(1,MAX(0,{b}))')
    if key in ('ets.vol', 'ets.imp'):         # Settings volatility table: value for the MTInputs label
        lab, vals = ((c('D_ETSVolatility'), 'Settings!$C$17:$F$17') if key == 'ets.vol'
                     else (c('D_ETSVolImpact'), 'Settings!$C$18:$F$18'))
        if use_lambda:
            return '=' + call('VOLPICK', lab, vals)
        return f'=INDEX({vals},MATCH(SUBSTITUTE({lab},"*",""),Settings!$C$16:$F$16,0))'
    if key == 'ets.vadj':
        a = (c('ets.vol'), c('ets.imp'), c('D_ETSCTRisk'))
        return '=' + (call('VOLADJ', *a) if use_lambda else f'(1+{a[2]}*{a[1]})/(1+{a[0]}*{a[1]})')
    if key == 'ets.bce':
        return '=' + (call('COVERED', bco2, etsc) if use_lambda else f'SUMPRODUCT({bco2},{etsc})')
    if key == 'ets.chg':                     # relative cap change: linear from start to target year, flat after
        a = (y, c('D_ETSIntro'), c('D_ETSOutro'), c('D_ETSChangeRelStart'), c('D_ETSChangeRelTarget'))
        if use_lambda:
            return '=' + call('CAPCHANGE', *a)
        return f'={a[3]}+({a[4]}-{a[3]})/MAX({a[2]}-{a[1]},1)*(MIN({y},{a[2]})-{a[1]})'
    if key == 'ets.cap':
        a = (c('ets.on'), y, c('D_ETSOutro'), c('D_ETSCapCont'), c('ets.bce'), c('ets.chg'),
             f'{L(CI_(col) - 1)}${R_POL["ets.cap"]}')
        if use_lambda:
            return '=' + call('ETSCAP', *a)
        return f'=IF({on},IF(AND({y}>{a[2]},LEFT({a[3]},8)="Constant"),{a[6]},{a[4]}*(1+{a[5]})),0)'
    if key == 'ets.se':                      # baseline-weighted semi-elasticity x allocation effectiveness
        obr = pol_range(col, 'obr.pow', len(COV_SECTORS))
        if use_lambda:
            return '=' + call('ETSSEMI', bco2, etsc, 'SemiEl', obr, 'OBREff', c('ets.bce'))
        return f'=IF({c("ets.bce")}>0,SUMPRODUCT({bco2},{etsc},SemiEl,1-(1-OBREff)*{obr})/{c("ets.bce")},0)'
    if key == 'ets.est':
        a = (c('ets.on'), c('ets.cap'), c('ets.bce'), c('ets.se'), c('ets.vadj'))
        if use_lambda:
            return '=' + call('ETSESTIMATE', *a)
        return f'=IF(AND({on},{a[1]}>0,{a[1]}<{a[2]},{a[3]}<0,{a[4]}>0),LN({a[1]}/{a[2]})/({a[3]}*{a[4]}),0)'
    if key == 'ets.pe':
        return '=' + (call('TAXEQUIV', c('ets.p'), c('ets.vadj')) if use_lambda else f'{c("ets.p")}*{c("ets.vadj")}')
    if key == 'ets.rf':
        v = c('ets.vadj')
        return '=' + (call('REVFACTOR', v) if use_lambda else f'IF({v}>0,1/{v},0)')
    if key.startswith('etsc.'):
        a = (c('ets.on'), rel('etscv.' + key[5:]))
        return '=' + (call('ETSCOVER', *a) if use_lambda else f'IF({on},IF({a[1]},1,0),0)')
    if key.startswith('shps.'):
        a = (rel('fb.' + key[5:]), rel('reg.' + key[5:]))
        return '=' + (call('SHADOWSECTOR', *a) if use_lambda else '+'.join(a))
    if key.startswith('ctr.'):               # carbon tax rate of this sector ($/tCO2), not where the ETS covers it
        # v1.09: CTaxComplimentaryToETS = Yes -> coverage 1 - existing ETS sector coverage (D, sheet ECP)
        a = (f'{col}${R_CP}', rel('sc.' + key[4:]), rel('etsc.' + key[4:]), c('CTaxComplimentaryToETS'),
             f'$D{R_POL[key]}')
        return '=' + (call('SECTORTAX', *a) if use_lambda else
                      f'{a[0]}*IF(LEFT({a[3]},3)="Yes",1-{a[4]},{a[1]})*(1-{a[2]})')
    if key in ('xct.d', 'xets.d', 'xct.p0', 'xets.p0'):   # v1.09: ECP data of the year (embedded: <= last hist.)
        code = 'ecr.all' if key.startswith('xct') else 'eer.all'
        yy = f'MIN({y},ECPLast)' if key.endswith('.d') else f'MIN({y},ECPLast,{LH})'
        return f'=INDEX(ECPVal,MATCH("{code}",ECPKey,0),MATCH({yy},ECPYear,0))'
    if key in ('xct.p', 'xets.p'):           # v1.09: price used = switch x data (x growth after the last data year)
        sch = key.split('.')[0]
        sw, g = (('ExistingCTApply', 'ExistingCTGrowth') if sch == 'xct'
                 else ('ExistingETSApply', 'ExistingETSGrowth'))
        a = (c(sw), c(f'{sch}.d'), c(g), y)
        return '=' + (call('ECPPRICE', *a) if use_lambda else
                      f'IF(LEFT({a[0]},3)="Yes",{a[1]}*(1+{a[2]})^MAX(0,{y}-ECPLast),0)')
    if key.startswith('etr.'):               # ETS cost in the fuel price of this sector ($/tCO2), auctioned part
        a = (c('ets.pe'), rel('etsc.' + key[4:]), rel('obr.' + key[4:]))
        return '=' + (call('SECTORETS', *a) if use_lambda else f'{a[0]}*{a[1]}*(1-{a[2]})')
    if key.startswith('shpx.'):              # shadow price on the efficiency margin ($/tCO2) of this sector
        s = key[5:]
        a = (f'INDEX({pol_range(col, "shps." + FEEBATE_SECTORS[0][0], len(FEEBATE_SECTORS))},$D{R_POL[key]})',
             rel('ssc.' + s), c('ets.pe'), rel('etsc.' + s), rel('obr.' + s))
        return '=' + (call('SHADOWTOTAL', *a) if use_lambda else f'{a[0]}*{a[1]}+{a[2]}*{a[3]}*{a[4]}')
    if key.startswith('ssc.'):
        a = (rel('fbc.' + key[4:]), c('ssc.adj'))
        return '=' + (call('SHADOWSHARE', *a) if use_lambda else f'IF({a[0]},1,0)*{a[1]}')
    raise ValueError(key)


# ---------------------------------------------------------------- Mitigation
def lam_row(c, r, var, fuel, sector, formula):
    """Record the LAMBDA alternative of a calculated row (right column of scenario 1) for sheet Lambdas."""
    if c == group_cols(1)[2][-1]:
        LAMBDA_ROWS.append((r, var, fuel, sector, formula))


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
        put(ws, f'{col}{R_GDP}', (f'=INDEX(WEOVal,MATCH(Settings!$C$4&"|gdp_growth",WEOKey,0),'
                                  f'MATCH({col}${R_YEAR},WEOYear,0))'), fill_=bf, fmt='0.0%')
        for r_, key, base_ in ((R_CPI, 'USA|pcpi', 'CPIRes'), (R_DEFL, 'USA|ngdp_d', 'DeflRes')):
            put(ws, f'{col}{r_}', f'={base_}/INDEX(WEOVal,MATCH("{key}",WEOKey,0),MATCH({col}${R_YEAR},WEOYear,0))',
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
        if key.startswith('ctr.'):            # v1.09: existing ETS sector coverage (complementary carbon tax)
            put(ws, f'D{r}', f'=INDEX(ECPCov,MATCH("ees.{ECP_SECTOR[sector]}",ECPKey,0))', FONT, None, '0.000')
        codes(r, fnt)
        outline[r] = (0, False) if summary else (1, True)

        def pcell(c, is_base, is_last, r=r, key=key, kind=kind, fnt=fnt, var=var, fuel=fuel, sector=sector):
            col = L(c)
            bf = F_BASE if is_base else None
            if kind == 'mt':
                val = mt_lookup(r, col)
            elif kind == 'calc':
                val = pol_calc(key, col, False)
                lam_row(c, r, var, fuel, sector, pol_calc(key, col, True))
            elif kind == 'ecp':
                val = pol_calc(key, col, False)
            elif kind == 'data':
                val = DATA_VALUES.get((key, year_of(c)))
            else:
                val = CONST_VALUES[key]
            fmt = '0' if var in ('CPIntro', 'CPOutro', 'fpr.yr0', 'fpr.yr1', 'fb.yr0', 'fb.yr1') else '0.00'
            if var in ('ctcov', 'fbcov') and kind == 'const':
                fmt = 'General'
            lam = False
            if var in ECP_PARAMS[:4] and var.endswith('Growth'):
                fmt = '0.0%'
            elif var in ECP_PARAMS:
                fmt = 'General'
            if var in ('bco2', 'ets.bce', 'ets.cap'):
                fmt = '#,##0.000'
            elif var == 'ets.se':
                fmt = '0.0000000'
            put(ws, f'{col}{r}', val, FONT_LAMBDA if lam else fnt,
                F_LAMBDA if lam else (F_INPUT if kind in ('const', 'data') else bf), fmt)
        each_col(pcell)

    # 2. Retail energy prices (before new policies)
    band(ws, B_PRI, '2. Retail energy prices before new policies (real $/GJ of ResultsYear; by price fuel; '
                    'legacy method, PriceProjection_Method_v0.6.md)', last)
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
    for key_ in ('gp.src', 'gp.adj'):
        r = R_PRI[key_]
        ids(r, key_, '', '')
        codes(r)
        outline[r] = (1, True)

        def gscell(c, is_base, is_last, r=r, key_=key_):
            put(ws, f'{L(c)}{r}', gp_formula(r, L(c), key_), fill_=F_BASE if is_base else None,
                fmt='@' if key_ == 'gp.src' else '0')
        each_col(gscell)
    for r0, var_ in ((R_GPN0, 'gpn'), (R_GP0, 'gp')):
        for k, (fu, lab, key, hi, lo) in enumerate(GP_ROWS):
            r = r0 + k
            ids(r, var_, fu, 'int')
            codes(r)
            outline[r] = (1, True)
            pv = [key, hi, lo, None] if var_ == 'gp' else [key, None, None, None]
            for pc, v in zip(PARAM_COLS, pv):
                put(ws, f'{pc}{r}', v, FONT, F_UNUSED if v is None else F_INPUT)

            def gcell(c, is_base, is_last, r=r, var_=var_):
                put(ws, f'{L(c)}{r}', gp_formula(r, L(c), var_), fill_=F_BASE if is_base else None, fmt='0.00')
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

            def pcell2(c, is_base, is_last, r=r, var=var, fuel=fuel, sec=sec):
                lam = False
                block = 'hist' if year_of(c) <= HIST_YEARS[-1] else 'proj'
                lam_row(c, r, var, fuel, sec, price_formula(var, r, L(c), L(c - 1), 'lam'))
                fl = F_LAMBDA if lam else (F_BASE if is_base else (F_LIGHT if block == 'hist' and var in ('sp', 'txo')
                                                                    else None))
                put(ws, f'{L(c)}{r}', price_formula(var, r, L(c), L(c - 1), block),
                    FONT_LAMBDA if lam else FONT, fl, '0.000')
            each_col(pcell2)
        outline[R_PV0 + PVOFF[var] + NP] = (1, True)            # blank row after each variable

    # 3. Power sector: engineer model with generation costs and electricity prices (v1.05-v1.06)
    band(ws, B_POW, '3. Power sector: engineer model (demand, dispatch, retirement, investment) with generation costs '
                    'and electricity prices (legacy section 4, sub-tables A-M; PowerModel_Method_v0_4.md)', last)
    for let, head, _items in PW_LAYOUT:
        put(ws, f'H{PW_HEAD[let]}', head, sub_font)
        outline[PW_HEAD[let]] = (1, True)
    for r, var, fu, sec, kind in PW_ROWS:
        summary = var in PW_SUMMARY
        fnt = FONT_SUM if summary else FONT
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

        def pwcell(c, is_base, is_last, r=r, var=var, fnt=fnt, fmt=fmt, fu=fu, sec=sec):
            plain_all = var in PW_DATA or var in PW_SUMS or var in PW_NOLAM
            lam = False
            if not plain_all:
                lam_row(c, r, var, fu, sec, pw_formula(var, r, L(c), L(c - 1), 'lam'))
            if var in PW_DATA:
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
        'Year by year as legacy: demand (H) uses last year\'s electricity price (G); dispatch (I) fills demand from this '
        'year\'s capacity (K); the generation shares set the VRE share and storage cost (J), the average generation '
        'cost (E) and next year\'s supply cost and price (G); investment (L) adds the capacity needed for this year\'s '
        'demand to next year\'s capacity, using levelised costs (C + D9 = E1). dsum and lvc look one column ahead '
        '(levelised variable cost to the end of the horizon), so they have plain formulas in every column.',
        'Bug fixes against legacy (PowerModel_BugsAndImprovements_v1_0.docx): services demand, efficiency exponent, '
        'income factor, long-term storage allocation, investment spending, investment logit, early-retirement units, '
        'coal effective retirement, feebate rebate, capacity check. Power enters the main thread (v1.09): CO2 in '
        'section 13 (total, by fuel, ETS sector power), generation fuel use in section 11 (row ener.pow, totals by '
        'fuel), revenues (sub-table M and G: rvat, rnvat, rgap) in section 12 (rows .pow).']
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
        d0, d1 = R_PW[('dem', DEM_SUBS[0])], R_PW[('dem', DEM_SUBS[-1])]
        memo_font = Font(name='Arial', size=9, italic=True)
        for s in subs:
            h = SUB_HEAD[s]
            ids(h, 'ener', 'all', s, FONT_SUM)
            codes(h, FONT_SUM)
            outline[h] = (1, False)
            ids(h + 1, 'eldem', 'ecy', s, memo_font)          # v1.07: electricity demand of the subsector (memo)
            codes(h + 1, memo_font)
            outline[h + 1] = (1, False)

            def eldcell(c, is_base, is_last, r=h + 1):
                col = L(c)
                put(ws, f'{col}{r}', f'=SUMIF($C${d0}:$C${d1},$C{r},{col}${d0}:{col}${d1})', memo_font,
                    F_BASE if is_base else None, '#,##0.0')
            each_col(eldcell)
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

                    def cell(c, is_base, is_last, r=r, var=var, fmt=fmt, f=f, s=s):
                        lam = False
                        lam_row(c, r, var, f, s, calc_formula(var, r, L(c), L(c - 1), is_base, True))
                        put(ws, f'{L(c)}{r}', calc_formula(var, r, L(c), L(c - 1), is_base, False),
                            FONT_LAMBDA if lam else FONT, F_LAMBDA if lam else (F_BASE if is_base else None), fmt)
                    each_col(cell)
                outline[h + VOFF[var] + NF] = (2, True)              # blank row after each variable

    # 11. Results - energy consumption
    spans = {g: (SUB_HEAD[[x[0] for x in SUBSECTORS if x[2] == g][0]],
                 SUB_HEAD[[x[0] for x in SUBSECTORS if x[2] == g][-1]] + SB - 1) for _n, g, _t in SECTIONS}
    band(ws, B_RES, '11. Results - energy consumption (fuel use of the subsectors and of power generation, ktoe)',
         last)
    a, z = S_FIRST, S_LAST
    ids(R_TOTAL, 'ener', 'all', 'all', FONT_SUM)
    codes(R_TOTAL, FONT_SUM)

    def tot(c, is_base, is_last):            # v1.09: + power generation fuel use (row ener.pow)
        col = L(c)
        put(ws, f'{col}{R_TOTAL}', f'=SUMIFS({col}${a}:{col}${z},SecVar,"ener",SecFuel,"<>all")+{col}{R_EPOW}',
            FONT_SUM, F_BASE if is_base else None, '#,##0.0')
    each_col(tot)
    ids(R_EPOW, 'ener', 'all', 'pow')
    codes(R_EPOW)
    q0, q1 = R_PW[('pe', PW_TECH[0][0])], R_PW[('pe', PW_TECH[-1][0])]     # power fuel use by type (section 3 I)

    def epow(c, is_base, is_last):
        col = L(c)
        put(ws, f'{col}{R_EPOW}', f'=SUMPRODUCT({pw_block("pe", col)},{pw_pcol("pe", "E")})',
            fill_=F_BASE if is_base else None, fmt='#,##0.0')
    each_col(epow)
    for i, (s, *_r) in enumerate(SUBSECTORS):
        r = R_SUB0 + i
        ids(r, 'ener', 'all', s)
        codes(r)

        def bysub(c, is_base, is_last, r=r):
            col = L(c)
            put(ws, f'{col}{r}', f'=SUMIFS({col}${a}:{col}${z},SecVar,"ener",SecSub,$C{r},SecFuel,"<>all")',
                fill_=F_BASE if is_base else None, fmt='#,##0.0')
        each_col(bysub)
    for i, (f, *_r) in enumerate(FUELS):
        r = R_FUEL0 + i
        ids(r, 'ener', f, 'all')
        codes(r)

        def byfuel(c, is_base, is_last, r=r):
            col = L(c)
            pw = f'SUMIF($B${q0}:$B${q1},$B{r},{col}${q0}:{col}${q1})'     # v1.09: + power of this fuel
            put(ws, f'{col}{r}', f'=SUMIFS({col}${a}:{col}${z},SecVar,"ener",SecFuel,$B{r})+{pw}',
                fill_=F_BASE if is_base else None, fmt='#,##0.0')
        each_col(byfuel)
    far = L(last + 10 * GW)
    extra = [(R_CHK, 'ener.chk', '', '',
              lambda col: (f'=ABS(SUM({col}{R_SUB0}:{col}{R_EPOW})-{col}{R_TOTAL})'
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
    put(ws, f'H{R_ELD["all"] - 1}', 'Memo: electricity demand by sector (power model, section 3 H; final use of '
                                    'electricity, not in the fuel totals above, which count the fuels used to '
                                    'generate it)', Font(name='Arial', size=9, bold=True, italic=True))
    outline[R_ELD['all'] - 1] = (1, False)
    for g in ELD_GROUPS:
        r = R_ELD[g]
        ids(r, 'eldem', 'ecy', g, FONT_SUM if g == 'all' else FONT)
        codes(r)
        outline[r] = (1, False)

        def eld(c, is_base, is_last, r=r, g=g):
            col = L(c)
            if g == 'all':
                f = f'=SUM({col}{R_ELD[ELD_GROUPS[1]]}:{col}{R_ELD[ELD_GROUPS[-1]]})'
            else:
                a_, z_ = spans[g]
                f = f'=SUMIF($A${a_}:$A${z_},"eldem",{col}${a_}:{col}${z_})'
            put(ws, f'{col}{r}', f, FONT_SUM if g == 'all' else FONT, F_BASE if is_base else None, '#,##0.0')
        each_col(eld)

    # 12. Revenues: existing taxes, existing subsidies and new policies as separate calculations
    band(ws, B_REV, '12. Revenues (USD million, real of ResultsYear): existing taxes, existing subsidies and new '
                    'policies, calculated separately', last)
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
            if var in REV_VARS and sec == 'pow':     # v1.09: power (section 3: M by type, G by user group)
                parts = {'rtx': ['prtx', 'prxc', 'rvat'], 'rsub': ['prsub', 'rgap'], 'rnew': ['prnew', 'rnvat'],
                         'rxcp': ['prxl']}[var]
                f = '=' + '+'.join(f'SUM({pw_block(p_, col)})' for p_ in parts)
            elif var in REV_VARS:
                a, z = (S_FIRST, S_LAST) if sec == 'all' else spans[sec]
                crit = 'SecVar' if sec == 'all' else f'$A${a}:$A${z}'
                f = f'=SUMIFS({col}${a}:{col}${z},{crit},"{var}")'
                if sec == 'all':
                    f += f'+{col}{R_REV[var + ".pow"]}'
            elif var == 'rev.chk':
                f = '=' + '+'.join(f'ABS({col}{R_REV[v + ".all"]}-SUM({col}{R_REV[v + ".tra"]}:{col}{R_REV[v + ".pow"]}))'
                                   for v in ('rtx', 'rsub', 'rnew'))
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
            put(ws, f'{col}{r}', f, fnt, F_CHECK if var == 'rev.chk' else (F_BASE if is_base else None),
                '0.000' if var == 'rev.chk' else '#,##0.0')
        each_col(rcell)
    notes = [
        'Per subsector and fuel (rows rtx, rsub, rnew, rxcp in each subsector block): rtx = fuel use x PJ/ktoe x '
        '(existing tax etx + VAT vat of the price fuel + change in existing carbon pricing xcpd x (1 + VAT rate)); '
        'rsub = fuel use x PJ/ktoe x existing subsidy esub (fiscal cost, positive); rnew = fuel use x PJ/ktoe x (new '
        'carbon tax + new excise + ETS x revenue factor + VAT on new policies nce x VAT rate); rxcp = existing carbon '
        'pricing revenue (memo, level; the part embedded in the price data is inside rtx). Feebates are '
        'revenue-neutral.',
        'Power (rows .pow, v1.09; section 3 M and G): rtx = taxes on generation fuels + change in existing carbon '
        'pricing + electricity VAT; rsub = subsidies on generation fuels + electricity subsidy gap (supply cost minus '
        'price before VAT; negative if the price is above cost); rnew = carbon tax, ETS at the permit price and fuel '
        'price reform on generation + VAT on the policy cost passed through. Existing taxes and subsidies use the '
        'scenario\'s own quantities, so a policy also changes them (the fiscal effect vs scenario 1 is rtot.chg); '
        'rev.chk = 0 when the totals equal their rows.',
        'Per-fuel policy input (later): esubpu (existing subsidy per MTInputs price unit, section 2) can feed a '
        'subsidy-reform path in the fuel price reform rows (MTInputs 142-169 use the same units).']
    for k, n in enumerate(notes):
        put(ws, f'H{R_REV_NOTE + k}', n, Font(name='Arial', size=9, italic=True))

    # 13. Emissions: CO2 from fuel combustion
    band(ws, B_EMI, '13. Emissions - CO2 from fuel combustion (MtCO2; subsectors and power generation; process '
                    'emissions and other gases not yet)', last)
    for key, var, fu, sec in EMI_ROWS:
        r = R_EMI[key]
        summary = key in ('co2.all', 'co2.chg', 'co2.pct')
        fnt = FONT_SUM if summary else FONT
        ids(r, var, fu, sec, fnt)
        codes(r, fnt)
        outline[r] = (0, False) if summary else (1, True)

        def ecell(c, is_base, is_last, r=r, key=key, var=var, fu=fu, sec=sec, fnt=fnt):
            col = L(c)
            q0, q1 = R_PW[('co2', PW_TECH[0][0])], R_PW[('co2', PW_TECH[-1][0])]   # power co2 by type
            pw_sum = lambda crit: f'SUMIF(${crit[0]}${q0}:${crit[0]}${q1},{crit[1]},{col}${q0}:{col}${q1})'
            if var == 'co2' and fu == 'all' and sec == 'pow':
                f = f'=SUM({pw_block("co2", col)})'
                fmt = '#,##0.00'
            elif var == 'co2' and fu == 'all':
                a, z = (S_FIRST, S_LAST) if sec == 'all' else spans[sec]
                crit = 'SecVar' if sec == 'all' else f'$A${a}:$A${z}'
                f = f'=SUMIF({crit},"co2",{col}${a}:{col}${z})'
                if sec == 'all':
                    f += f'+{col}{R_EMI["co2.pow"]}'
                fmt = '#,##0.00'
            elif var == 'co2':
                a, z = S_FIRST, S_LAST
                f = f'=SUMIFS({col}${a}:{col}${z},SecVar,"co2",SecFuel,$B{r})+{pw_sum(("B", f"$B{r}"))}'
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
                f = f'=SUMIFS({col}${a}:{col}${z},SecVar,"co2",SecSub,$C{r})+{pw_sum(("C", f"$C{r}"))}'
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
            elif var == 'ets.lr':
                cap, bce, cov = (f'{col}${R_POL["ets.cap"]}', f'{col}${R_POL["ets.bce"]}', f'{col}{R_EMI["co2.ets"]}')
                f = f'=IF(OR({cap}<=0,{cap}>={bce},{cov}>={bce}),0,LN({cap}/{bce})/LN({cov}/{bce}))'
                fmt = '0.0000'
            elif var == 'ets.next':
                p_, lr = f'{col}${R_POL["ets.p"]}', f'{col}{R_EMI["ets.lr"]}'
                f = f'=IF(OR({p_}<=0,{lr}=0),{col}${R_POL["ets.est"]},{p_}*{lr}^Settings!$C$15)'
                fmt = '0.00'
            elif var == 'co2.chk':                                      # v1.09: sectors + power, fuels
                g0, g1 = R_EMI[f'co2.{SECTIONS[0][1]}'], R_EMI['co2.pow']
                f0, f1 = R_EMI[f'co2.f.{FUELS[0][0]}'], R_EMI[f'co2.f.{FUELS[-1][0]}']
                t = f'{col}{R_EMI["co2.all"]}'
                f = f'=ABS({t}-SUM({col}{g0}:{col}{g1}))+ABS({t}-SUM({col}{f0}:{col}{f1}))'
                fmt = '0.000'
            else:                                                       # bco2.chk
                f = f'=SUM({pol_range(col, "bco2.pow", len(COV_SECTORS))})-{col}{R_EMI["co2.ref"]}'
                fmt = '0.000'
            put(ws, f'{col}{r}', f, fnt, F_CHECK if var in ('bco2.chk', 'co2.chk') else (F_BASE if is_base else None),
                fmt)
        each_col(ecell)
    put(ws, f'H{R_EMI_NOTE}', 'Per subsector and fuel: co2 = fuel use (ktoe) x PJ/ktoe x EF (tCO2/GJ, IIASA, EF_GHG '
                              'sheet; biomass 0); power generation: section 3 I (fuel use x MtCO2 per ktoe), in the '
                              'total, by fuel and as ETS sector pow. No inventory adjustment, no process emissions; '
                              'CH4, N2O and local pollutants later. co2.chk = 0 when the total equals sectors + power '
                              'and the fuels. ETS rows: co2.sec by ETS sector, co2.ets covered emissions, co2.cap the '
                              'cap, co2.gap the distance to the cap, '
                              'ets.next the next price (paste its values into ets.ovr with D_ETSPriceOverride = Yes '
                              'and repeat, or run ets_goalseek_v0_4.py); bco2.chk = 0 when the baseline data in '
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
                   '(section 13), and the power sector (section 3: engineer model with demand, dispatch, retirement '
                   'and investment, generation costs and electricity prices, power CO2). Differences with legacy: '
                   'sheet LegacyDiff.'),
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
                                      'projected year. See PriceProjection_Method_v0.6.md.'),
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
                        '(ets.next); ets_goalseek_v0_4.py iterates until the cap is met. See '
                        'ETS_Legacy_Algorithm_v0_1.md.'),
        ('Version 1.09', 'Existing carbon pricing and power in the main thread. Sheet ECP: the legacy existing carbon '
                         'tax and ETS data (prices by year, sector and fuel coverage; 52 countries, none for Egypt) and '
                         'the values of the country in Settings C4. Section 1 (MTInputs 208-214): price used = data x '
                         'switch, growing after the last data year; the level embedded in the price data. Per '
                         'subsector and fuel: xcp (level), xcpd (change against the price data: in the after-tax price '
                         'and existing-tax revenue), rxcp (revenue memo); power: xck in the variable cost. Section 11 '
                         'counts power generation fuel use (row ener.pow; totals by fuel include it); section 12 has '
                         'power rows (fuel taxes and subsidies, electricity VAT and subsidy gap, new policies with the '
                         'ETS at the permit price; sub-table M of section 3) in every total; checks ener.chk, rev.chk, '
                         'co2.chk.'),
        ('Version 1.08', 'The workbook calculates without LAMBDA support (any Excel version from 2019 with MINIFS, '
                         'LibreOffice): every cell holds a plain formula of at most 100 characters, the right column '
                         '(2040) included, so every column of a row has the same formula and can be dragged. Column H '
                         'reads the member labels from a table on Mapping. The named LAMBDAs remain in the Name '
                         'Manager as an optional library: sheet Lambdas gives, for every calculated row, the LAMBDA '
                         'formula that can replace its plain formulas (same results, tested).'),
        ('Version 1.07', 'Short formulas: every cell formula has at most 100 characters; longer logic sits in the '
                         'named LAMBDAs (Name Manager), lookups use defined names for the data tables (WEOVal, '
                         'PDomVal, PPathVal, ...), and multi-step formulas show their steps in helper rows (section 1: '
                         'carbon-price path years and nominal factor, ETS on/off, volatility, impact, cap change, '
                         'policy rates by sector ctr, etr, shpx; section 2: gp.src, gp.adj, gpn, txf; section 3: '
                         'ecmin, ecshr; section 13: ets.lr) or helper columns (Inputs_prices by year, MTOutputs, '
                         'ScenarioCompare). Planned retirement: PowerTech planned_retirement = Yes uses the data rows '
                         'pret|<type> of PowerPaths instead of capacity x 1/lifetime (Egypt: no data). Memo lines '
                         'eldem: electricity demand by subsector in each sector block and by sector in section 11 '
                         '(not in the fuel totals; sector policies do not act on power demand). MTOutputs: 18 more '
                         'power indicators; Charts: 6 power charts.'),
        ('Power engineer model (v1.06)', 'Section 3 continues with legacy H-L. H demand by subsector: last year\'s '
                                         'demand x group multiple (efficiency, GDP with income factor, price of the '
                                         'previous year); required generation = total demand x balance factor. I '
                                         'dispatch: wind, solar, hydro, other renewables, oil, nuclear, biomass at '
                                         'their capacity factor; coal and gas fill the rest (minimum of each if the '
                                         'other runs at 90%, the remainder by a logit on variable cost). K: capacity = '
                                         'last year less retirements plus new investment; cost-based early coal '
                                         'retirement. L: investment needed = this year\'s required generation less '
                                         'capacity after retirements (available next year), allocated by a logit on '
                                         'levelised investment cost within scale-up limits, the rest to the cheapest '
                                         'remaining space. Power CO2 = fuel use x EF, in section 13. Ten legacy bugs '
                                         'fixed: PowerModel_BugsAndImprovements_v1_0.docx. Method: '
                                         'PowerModel_Method_v0_4.md.'),
        ('Power prices (v1.05)', 'Section 3, legacy sub-tables: A inputs (capex factor, storage cost paths); '
                                 'B amortised fixed cost of the stock = caxav x acap + afix + stoav (vintage averages '
                                 'with last year\'s investment share); C levelised fixed cost of new plants (for '
                                 'investment only); D variable cost = O&M + section 2 fuel price x 0.0036/efficiency, '
                                 'and the new policy cost per kWh ccp (carbon tax, ETS, fuel price reform on power '
                                 'fuels); E generation cost gnc = fix + vbc and the average gncav; G supply cost sc = '
                                 'gncav of the previous year + T&D, price before new policies = data to 2024, then + '
                                 'pass-through x change in sc x (1 + VAT), end-user price rp = rppre + ccpav x '
                                 'pass-through x (1 + VAT), '
                                 'subsidy gap and revenues; J storage cost of variable renewables from the VRE share. '
                                 'Constants per type: Inputs_power. Method: PowerPrices_Method_v0_3.md.'),
        ('Emissions (v0.13)', 'co2 per subsector and fuel = fuel use x PJ/ktoe x EF (tCO2/GJ), MtCO2. Section 13: '
                              'total, by sector, by fuel, change vs scenario 1. Fuel combustion incl. power '
                              'generation (v1.06; no process emissions, CH4, N2O or local pollutants yet).'),
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
        ('Revenues (v0.12, v1.09)', 'Three separate calculations per subsector and fuel, USD million real: rtx '
                                    'existing tax revenue = fuel use x PJ/ktoe x (etx + vat + xcpd x (1 + VAT rate)); '
                                    'rsub existing subsidy cost = fuel use x PJ/ktoe x esub; rnew new-policy revenue = '
                                    'fuel use x PJ/ktoe x (ctxnew + ntx + ets x revenue factor + nce x VAT rate); rxcp '
                                    'existing carbon pricing (memo). Power (section 3 M and G): fuel taxes and '
                                    'subsidies on generation, electricity VAT and subsidy gap, new policies with the '
                                    'ETS at the permit price. Section 12: totals = sectors + power, net existing (rtx - '
                                    'rsub), total, change vs scenario 1, check rev.chk. Feebates are revenue-neutral.'),
        ('Existing carbon pricing (v1.09)', 'Sheet ECP holds the legacy data (all countries) and the values of the '
                                            'country in Settings C4 (real prices; coverage of the last data year). '
                                            'MTInputs 208-214: apply the existing carbon tax / ETS (Yes/No), real '
                                            'growth after the last data year, new carbon tax complementary to the '
                                            'existing ETS. The IMF retail prices already contain existing carbon '
                                            'pricing, so only its change against the level in the price data (xcpd, '
                                            'power xck) enters prices and existing-tax revenue; rxcp shows the level.'),
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
        ('LAMBDAs (optional)', 'Every cell holds a plain formula, the right column included (v1.08), so the '
                               'workbook calculates without LAMBDA support. Prices with history have separate '
                               f'blocks: data up to {HIST_YEARS[-1]}, projection after; changing the last '
                               'historical year (Settings C10) needs the blocks re-dragged. The named LAMBDAs '
                               '(Formulas > Name Manager) are an optional library: sheet Lambdas gives the LAMBDA '
                               'formula of each calculated row; pasted into the right column and copied over the '
                               'row it gives the same results, and a LAMBDA such as SUPPLYCOST holds the IF between '
                               f'history and projection. {lam_lines}.'),
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
                 'PowerPaths, PowerParams, PowerDemand; Mapping; '
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
    import check_v1_09 as C
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
    name), recalculated in LibreOffice (LAMBDAs expanded), with the ETS goal seek (ets_goalseek_v0_4.solve, the
    macro's algorithm) when the definition applies a new ETS without the price override. Returns StoredResults
    rows."""
    import tempfile
    import openpyxl
    import check_v1_09 as C
    import ets_goalseek_v0_4 as G
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
    put(ws, 'B5', 'Year column in StoredResults', FONT_B)
    put(ws, 'C5', '=MATCH($C$4,SRYear,0)', FONT, None, '0')
    put(ws, 'C6', 'Scenario ID', FONT_B)
    put(ws, 'C7', 'Scenario name', FONT_B)
    ids = [n for n, _nm, run, _e in DEFINITIONS if run == 'Yes']
    for c in range(first, lastc + 1):
        cl = L(c)
        sid = 1 if c == first else (ids[c - first - 1] if c - first - 1 < len(ids) else None)
        put(ws, f'{cl}6', sid, FONT_B, F_BASE if c == first else F_INPUT, '0')
        put(ws, f'{cl}7', f'=IF({cl}$6="","",IFERROR(INDEX(SRName,MATCH({cl}$6,SRScen,0)),"not stored"))', FONT_B)
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
            put(ws, f'B{r}', f'=IFERROR(INDEX(SRDesc,MATCH($A{r},SRCode,0)),"")')
            put(ws, f'C{r}', f'=IFERROR(INDEX(SRUnit,MATCH($A{r},SRCode,0)),"")')
            for c in range(first, lastc + 1):
                cl = L(c)
                if kind == 'lvl':       # blank ID: no stored row has an empty ID, so the count is 0
                    f = (f'=IF(COUNTIFS(SRScen,{cl}$6,SRCode,$A{r})=0,"",'
                         f'SUMIFS(INDEX(SRVal,0,$C$5),SRScen,{cl}$6,SRCode,$A{r}))')
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
    build_member_labels(wb)
    build_lambdas_sheet(wb)
    build_mtoutputs(wb)
    build_charts(wb)
    build_stored_results(wb, [])
    build_scenario_compare(wb)
    build_macro_check(wb)
    div = wb.create_sheet('DATA->')
    put(div, 'A1', 'Data sheets follow', FONT_B)
    build_data_sheets(wb)
    build_ecp(wb)
    build_readme(wb)
    for name in LAMBDAS:
        wb.defined_names[name] = DefinedName(name, attr_text=lambda_xml(name), comment=LAMBDA_NOTES[name])
    clash = {n.upper() for n in defined_names()} & {n.upper() for n in LAMBDAS}
    assert not clash, f'defined names clash with LAMBDA names (case-insensitive): {clash}'
    for name, ref in defined_names().items():
        wb.defined_names[name] = DefinedName(name, attr_text=ref)
    order = ['Mitigation', 'MTOutputs', 'StoredResults', 'ScenarioCompare', 'MacroCheck', 'Charts', 'ReadMe', 'LegacyDiff',
             'Settings', 'Lambdas',
             'MTInputs', 'Inputs', 'Inputs_prices', 'Inputs_power', 'Variables', 'Mapping', 'DATA->',
             'Countries', 'Elasticities', 'Prices_dom', 'Prices_int', 'PriceAssump', 'EnergyCons', 'WEO', 'EF_GHG',
             'PowerTech', 'PowerPaths', 'PowerParams', 'PowerDemand', 'ECP']
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    for w in wb.worksheets:
        w.sheet_view.tabSelected = w.title == 'Mitigation'
    if os.environ.get('QUICK'):                    # debugging: no baseline-CO2 refresh, no batch emulation
        wb.save(OUT.replace('.xlsm', '_quick.xlsx'))
        print('Saved quick', OUT.replace('.xlsm', '_quick.xlsx'))
        return
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
