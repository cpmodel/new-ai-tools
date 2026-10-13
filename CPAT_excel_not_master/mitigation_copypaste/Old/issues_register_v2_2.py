"""Register of bugs, known issues and potential upgrades of CPAT-AI-Mitigation-MVP v2.2: the single source for the
Word register (make_issues_doc_v2_2.py -> CPAT-AI-Mitigation-MVP_BugsIssuesUpgrades_v2.2.docx) and the summary in the
Word documentation (make_docs_v2_2.py). check_v2_2.py tests that the IDs are unique and that every LegacyDiff row of
the types in LD_TYPES is included.

IDs: L = bug in legacy CPAT (fixed in the MVP unless stated), M = bug found and fixed in the MVP itself during
development, I = known issue or limitation of the MVP today, LD = simplification or feature not yet modelled (from the
workbook sheet LegacyDiff, row number), U = potential upgrade (priority High / Medium / Low).
"""
import build_v2_2 as B

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
MINOR = [  # L11, L12 (kept for the effect report)
    ('Stray reference in dispatch', 'I1, rows 3773 and 3779',
     'Required generation from coal and gas is MAX(required - other generation, $G4291 x total), and G4291 is an '
     'unrelated, empty cell of table L21. Harmless while the cell is empty; not built in the MVP.'),
    ('Total nameplate investment row is zero', 'L19, row 4273',
     'The total multiplies by $X3008, an empty cell, instead of summing the types. Reporting only; the MVP sums.'),
]

LEGACY = [  # L11 onwards: id, title, where, what, status
    ('L11', 'Stray reference in dispatch', 'Mitigation rows 3773, 3779 (I1)',
     'Required generation from coal and gas is MAX(required - other generation, $G4291 x total); G4291 is an '
     'unrelated, empty cell of table L21.', 'Harmless while the cell is empty; not built in the MVP.'),
    ('L12', 'Total nameplate investment row is zero', 'Mitigation row 4273 (L19)',
     'The total multiplies by $X3008, an empty cell, instead of summing the types.', 'Reporting only; the MVP sums.'),
    ('L13', 'Other oil products: hardcoded pass-through 1 and margin 0', 'Mitigation price block (other oil products)',
     'With the hardcodes the retail price of other oil products falls with the oil price to the 0.01 floor and their '
     'fuel use rises tenfold by 2030 (Egypt).', 'Changed in the MVP (v0.10, user decision): pass-through and margin '
     'from the IMF dataset like the other oil products (LegacyDiff).'),
    ('L14', 'Pass-through applied twice in the first projected year', 'Mitigation txo rows (cs)',
     'With pass-through 0.5 or 0.8 the floating tax multiplies its last value by (1 - pass-through) a second time, so '
     'taxes jump in the first projected year.', 'Kept as legacy (PriceProjection_Method_v2_2.md A10); no effect for '
     'Egypt (pass-through 0 or 1).'),
    ('L15', 'ETS: hardcoded relative price 1.1', 'Mitigation row 1843',
     'The ETS price effectiveness uses a fixed 1.1 instead of the volatility inputs of the ETS+LTS table.',
     'Replaced in the MVP by the volatility adjustment (v1.02); suggested to the legacy team (TODO).'),
    ('L16', 'ETS quick estimate linear in the semi-elasticity', 'Mitigation row 1868',
     'The fast estimate divides the reduction by the semi-elasticity (linear); the legacy note says "arguably exp/ln".',
     'MVP uses the log form LN(cap / baseline) / (semi-elasticity x effectiveness x volatility adjustment) (v1.02).'),
    ('L17', 'Version-log rows rewritten with the current version', 'MVP builders v0.3-v0.4 (not legacy)',
     'Earlier version-log rows showed the current version number.', 'Fixed in v0.5: each row has its own version.'),
]
LEGACY = [x for x in LEGACY if x[0] != 'L17']           # L17 is an MVP bug: listed in MVP_BUGS

MVP_BUGS = [  # id, title, found / fixed in, what happened, fix
    ('M1', 'Version-log rows rewritten with the current version', 'v0.3-v0.4 / v0.5',
     'The builders wrote earlier version-log rows with the current version number.',
     'Each row has its own version; old workbooks left as shipped.'),
    ('M2', 'Power: VAT left out on the policy term', 'v1.05 / v1.06',
     'The end-user price added the policy cost without VAT; legacy row 3614 charges VAT.',
     'VAT on the policy term as legacy (no effect for Egypt, VAT 0).'),
    ('M3', 'Power: long-term storage allocation left out', 'v1.05 / v1.06',
     'The legacy note "incomplete" was read as "not applied".', 'The 33% allocation is applied once.'),
    ('M4', 'Power: VRE share without other renewables', 'v1.05 / v1.06',
     'Legacy row 3851 also counts other renewables.', 'Follows legacy (other renewables 0 for Egypt).'),
    ('M5', 'Defined name clashed with a LAMBDA name', 'v1.07 build / v1.07',
     'The defined name ETSSemi equals the LAMBDA ETSSEMI ignoring case and gave #NAME?; the ETS goal seek stopped '
     'converging.', 'Renamed SemiEl; the builder asserts that no defined name equals a LAMBDA name.'),
    ('M6', 'Power ETS revenue valued at the tax-equivalent price', 'v1.05 / v1.09',
     'Power new-policy revenue (ccpav x consumption) valued the ETS at the tax-equivalent price, about 10% high at '
     'Medium volatility.', 'Revenue by generation type (prnew) with the ETS at the permit price; rcarb kept as a memo.'),
    ('M7', 'Stale texts in the workbook and notes', 'v1.08 / v1.09',
     'Section 12 and 13 notes said power was not included; PowerPaths said interim data; data sheets named old '
     'extractors; method notes described earlier versions.', 'Texts updated; the check now scans the workbook texts '
     'for stale statements and missing file references.'),
    ('M8', 'LegacyDiff rows out of date', 'v1.09 / v2.2',
     'Existing carbon pricing was still listed as "Not modelled"; the formula-structure row described the LAMBDA '
     'right column of v1.07.', 'Rows updated; the check scans for these phrases.'),
    ('M9', 'Legacy comparison note from v1.01', 'v1.02-v1.09 / v2.2',
     'legacy_comparison described the v1.01 baseline and its script wrote a different file name.',
     'Re-run on v2.2 (legacy_comparison_v2_2.md).'),
]

ISSUES = [  # id, title, description, where it is documented
    ('I1', 'Not yet run in Excel', 'The workbook and its VBA project were generated and tested without Excel '
     '(LibreOffice). Open it, enable macros and run CheckBatchRun once; if Excel rejects the VBA project, import '
     'CPATScenarios_v2_2.bas.', 'TODO; Scenarios_Method_v2_2.md'),
    ('I2', 'Excel 2019 or later', 'Power dispatch and investment use MINIFS. LAMBDA support is not needed.', 'README'),
    ('I3', 'New sector | fuel pairs need a label', 'A new pair on Mitigation needs a row in the member-label table on '
     'Mapping, or its description shows #N/A (20 spare rows in the defined name).', 'CAVEATS v1.08'),
    ('I4', 'Fuel and power demand respond to real USD prices', 'The corrected Egypt prices fall sharply in USD after '
     'the devaluation, so fuel use and electricity demand jump in 2023 (power +22%). Open decision: price vintage and '
     'last historical year, or demand on real local-currency prices.', 'TODO; legacy_comparison_v2_2.md; LegacyDiff'),
    ('I5', 'Totals include power since v1.09', 'Fuel use and revenue totals of v1.09 and later include power and are '
     'not comparable with v1.08 and earlier outputs.', 'CAVEATS v1.09'),
    ('I6', 'Electricity subsidy gap is signed', 'Negative when the tariff is above the modelled supply cost; fuel '
     'subsidies count positive parts only.', 'PowerPrices_Method_v2_2.md 3.9'),
    ('I7', 'Existing carbon pricing simplified', 'One ETS per country, single carbon tax rate, coverage of the last '
     'data year, existing ETS revenue at the full permit price, existing and new ETS add up. None of it matters for '
     'Egypt (no existing pricing).', 'PriceProjection_Method_v2_2.md 5b; LegacyDiff'),
    ('I8', 'Capacity is built for this year\'s demand', 'Legacy design kept: capacity built in a year serves this '
     'year\'s demand and is available next year, so gas runs harder. Recommended: next year\'s demand (U9).',
     'PowerModel_Method_v2_2.md'),
    ('I9', 'Levelised variable cost window ends in 2040', 'The D9 average shrinks to one year at the end of the '
     'horizon (legacy: 2041).', 'LegacyDiff'),
    ('I10', 'Spot fuel prices for power', 'Legacy uses a 5-year moving average; fuel price changes reach power one to '
     'four years earlier.', 'LegacyDiff'),
    ('I11', 'Changing the last historical price year', 'Prices have separate history and projection blocks; changing '
     'Settings C10 needs the blocks re-dragged (or the optional LAMBDA rows from sheet Lambdas).', 'README'),
    ('I12', 'Values after 2030', 'Not a legacy target; the legacy CPI is held flat after 2031.', 'LegacyDiff'),
    ('I13', 'Egypt data only for energy use, GDP and EFs', 'Elasticities, prices and existing carbon pricing cover all '
     'CPAT countries; energy use, GDP growth and emission factors are Egypt only.', 'README'),
    ('I14', 'Distance to the legacy cached baseline', 'The MVP baseline differs from the cached legacy run, mainly '
     'through the price vintage (I4) and the fixed legacy bugs.', 'legacy_comparison_v2_2.md'),
]

UPGRADES = [  # id, title, description, priority
    ('U1', 'Review in Excel', 'Run CheckBatchRun, the batch run and ScenarioCompare in Excel; confirm the ten power '
     'bug fixes (L1-L10).', 'High'),
    ('U2', 'Numerical check against legacy CPAT for Egypt', 'Bucket 5: explain the distance to the legacy run (I14) '
     'sector by sector.', 'High'),
    ('U3', 'Decide the price vintage and the demand price basis', 'Last historical price year, or demand on real '
     'local-currency prices (I4).', 'High'),
    ('U4', 'Report the legacy bugs to the CPAT team', 'L1-L16 with their effects; L15 suggested by the user.', 'High'),
    ('U5', 'Copy-paste stress test', 'Add a fuel, a subsector and a scenario by dragging; fix what breaks.', 'Medium'),
    ('U6', 'Other gases, local pollutants and process emissions', 'CH4, N2O, local pollutants, inventory-adjusted EFs.',
     'Medium'),
    ('U7', 'Regulations as shadow prices', 'Fill the regulation rows (placeholder 0) and their efficiency-margin '
     'adjustments (MTInputs 247-250).', 'Medium'),
    ('U8', 'Subsidy and price-control phase-outs', 'Use esubpu (existing subsidy per price unit) as a per-fuel policy '
     'input (MTInputs 123-138, 191-195).', 'Medium'),
    ('U9', 'Capacity for next year\'s demand', 'Use next year\'s demand in the capacity plan (I8); computable without '
     'a circular reference.', 'Medium'),
    ('U10', 'Data for more countries', 'Energy balances, GDP and EFs beyond Egypt (I13).', 'Medium'),
    ('U11', 'Existing carbon pricing in full', 'Two ETS blocks (national and regional), fuel-specific carbon tax rates, '
     'free allocation, coverage by year (I7).', 'Low'),
    ('U12', 'Longer levelised-cost window', 'Extend the last year\'s variable cost over the plant lifetime (I9).',
     'Low'),
    ('U13', 'Moving-average fuel prices for power', 'Option as legacy (I10).', 'Low'),
    ('U14', 'Partial-adjustment fuel-use model', 'Parked by the user (2026-10-09).', 'Low'),
    ('U15', 'Planned retirement data for Egypt', 'Enter PowerPaths pret|<type> where plans exist.', 'Low'),
    ('U16', 'Python port', 'The check script already recomputes every value in Python.', 'Low'),
]

IMPROVEMENTS = [
    'One plain formula per block, at most 100 characters, the right column included; no LAMBDA support needed; the '
    'optional LAMBDA library is on sheet Lambdas.',
    'The second investment round (legacy tables L11-L17, about 60 rows) is one row: permit x MAX(0, MIN(space, '
    'deficit - space of cheaper types)); coal and gas dispatch (legacy I2, 19 rows) is five rows.',
    'Amortised and levelised fixed costs are one row per type each, with constants derived once in Inputs_power.',
    'Electricity prices are telescoped, so components held at their last historical value drop out.',
    'Power is part of the main thread: fuel use, revenues and CO2 enter the totals, and a new ETS that covers power '
    'sees power emissions.',
    'Every number is checked: a Python reference reproduces 208 legacy rows of the baseline and the carbon-tax run to '
    '1e-13 with the bugs switched on, and the workbook equals the reference with the bugs fixed.',
]

LD_TYPES = ('Not yet', 'Simplification', 'To check (bucket 5)')


def legacy_diff_items():
    """LegacyDiff rows of the types in LD_TYPES: (id, title, legacy, this model, effect, type)."""
    return [(f'LD{k}', d[0], d[1], d[2], d[4], d[5]) for k, d in enumerate(B.LEGACY_DIFFS, 1) if d[5] in LD_TYPES]


def items():
    """Every register item as (id, title, category)."""
    out = [(f'L{i + 1}', t, 'Legacy bug') for i, (_k, t, *_r) in enumerate(BUGS)]
    out += [(i, t, 'Legacy bug') for i, t, *_r in LEGACY]
    out += [(i, t, 'MVP bug (fixed)') for i, t, *_r in MVP_BUGS]
    out += [(i, t, 'Known issue') for i, t, *_r in ISSUES]
    out += [(i, t, 'Simplification / not yet') for i, t, *_r in legacy_diff_items()]
    out += [(i, t, 'Upgrade') for i, t, *_r in UPGRADES]
    return out
