"""
Step 1 (pseudocode §5) -- deriving household-relevant price changes.

CPAT Excel: 'Distribution' sheet §B.II-B.III, §C.I.

Two things happen here:
1. price_change_direct[fuel]: the price change households face buying a
   fuel directly. In the full model this comes from the Mitigation module
   (see derive_price_change_from_scenarios below); it is treated as an
   *input* to this module, not derived from a raw carbon price here.
2. price_change_indirect[category]: the price change embedded in non-fuel
   goods and services, derived from price_change_direct via each GTAP
   sector's fuel-energy-intensity ('leontiefs': the domestic energy
   coefficient assigning fuel cost to sector output, i.e. an already-solved
   (I-A)^-1 row, not a matrix this module has to invert itself -- see
   IO_GTAP in data.py), then aggregated up to CPAT consumption categories
   using the GTAP<->CPAT crosswalk, weighted by each sector's share of
   total household demand.

Known gaps (data these formulas would need but isn't in
cpat_excel/Distribution/data_standardized): sectoral pass-through
coefficients (Excel §B.IV, rows 464-486) and the IEA/GAINS emissions-based
recalibration (§C.I, rows 632-664) both live on the Mitigation side of the
workbook, not in the DATA_DISTN block -- Steps 1.4/1.5 in the pseudocode are
no-ops here (assume_imperfect_passthrough is honoured only if a caller
supplies coefficients). The Pm/Pd revenue reconciliation (Step 1.3) also
isn't implemented: IO_GTAP's domestic-consumption figures are in GTAP-year,
GTAP-base-currency terms (real 2014 US$bn) with no deflator/FX series in
this dataset to convert them onto the same footing as an LCU-denominated
CP revenue figure, so 'Pd' can't be computed without pulling in Mitigation-
module GDP/deflator data too.
"""
from typing import TYPE_CHECKING

import pandas as pd

import cpat_model.constants as c

if TYPE_CHECKING:
    from cpat_model.scenario_results import ScenarioResults


# The 5 fuel groupings IO_GTAP's Leontief coefficients are given for. 'p_c'
# (petroleum & coal products) is GTAP's refined-fuels sector, which the CPAT
# model further splits into gso/die/ker/lpg at the household level -- there
# is no separate Leontief row per refined product, only for the sector as a
# whole, so those four share one embedded-energy coefficient (weighted down
# to a single 'p_c' price change first, see price_change_for_pc_sector).
LEONTIEF_FUEL_GROUPS = ['coa', 'ely', 'nga', 'oil', 'p_c']
PC_GROUP_FUELS = [c.GSO, c.DIE, c.KER, c.LPG]


def price_change_for_pc_sector(
        price_change_direct: pd.Series,
        pc_group_weights: pd.Series
        ) -> float:
    """
    GTAP's 'p_c' (petroleum & coal products) sector covers gasoline, diesel,
    kerosene and LPG at once. Collapse their individual price changes to one
    p_c price change, weighted by each fuel's household budget share (the
    closest available proxy for their relative weight within p_c; IO_GTAP
    does not split 'p_c' by refined product so a true output-weighted
    average isn't available here).

    price_change_direct: indexed by fuel code, at least covering PC_GROUP_FUELS
    pc_group_weights: indexed by fuel code, same fuels, e.g. budget shares
    """
    weights = pc_group_weights.reindex(PC_GROUP_FUELS).fillna(0.0)
    prices = price_change_direct.reindex(PC_GROUP_FUELS).fillna(0.0)
    total_weight = weights.sum()
    if total_weight == 0:
        return float(prices.mean())
    return float((prices * weights).sum() / total_weight)


def indirect_price_increase_by_gtap_sector(
        price_change_direct: pd.Series,
        pc_group_weights: pd.Series,
        io_gtap: pd.DataFrame
        ) -> pd.Series:
    """
    Step 1.2 (§7.2): for each GTAP sector, sum the embedded-energy cost
    shock from every fuel group: sum_f( price_change[f] * leontiefs[f, sector] ).

    io_gtap: this country's slice of data.load_io_gtap(), long format
    (columns dis, gtap_sector, value); dis is '<iso3>.<fuel>.leontiefs' for
    the 5 LEONTIEF_FUEL_GROUPS rows.

    return: Series indexed by GtapSectorCode, the % output-price increase
    implied by embedded fuel costs (does not include the sector's own
    direct fuel price, if any -- see aggregate_price_change_to_categories
    for how direct fuel sectors are excluded from the indirect side).
    """
    pc_price = price_change_for_pc_sector(price_change_direct, pc_group_weights)
    fuel_group_price = {
        'coa': price_change_direct.get(c.COA, 0.0),
        'ely': price_change_direct.get(c.ELE, 0.0),
        'nga': price_change_direct.get(c.NGA, 0.0),
        'oil': price_change_direct.get(c.OOP, 0.0),
        'p_c': pc_price,
    }

    leontiefs = io_gtap[io_gtap['dis'].str.endswith('.leontiefs')].copy()
    leontiefs['fuel_group'] = leontiefs['dis'].str.split('.').str[1]
    leontiefs = leontiefs[leontiefs['fuel_group'].isin(LEONTIEF_FUEL_GROUPS)]

    leontiefs['price_change'] = leontiefs['fuel_group'].map(fuel_group_price)
    leontiefs['contribution'] = leontiefs['price_change'] * leontiefs['value']

    return leontiefs.groupby('gtap_sector')['contribution'].sum()


def aggregate_price_change_to_categories(
        price_change_direct: pd.Series,
        pc_group_weights: pd.Series,
        io_gtap: pd.DataFrame,
        gtap_cpat_crosswalk: pd.DataFrame
        ) -> pd.Series:
    """
    Step 1.6: aggregate the ~59 GTAP sectors' embedded-energy price shocks
    up to the CPAT indirect consumption categories, weighted by each
    sector's share of total household demand ('hhd' in IO_GTAP) within its
    category.

    gtap_cpat_crosswalk: data.load_gtap_cpat_sector_crosswalk(), indexed by
    GtapSectorCode, with a 'CPATCategory' column ('NA' for sectors with no
    household-consumption mapping; comma-separated for 'p_c', which the
    direct-fuel side already handles, so it's excluded here).

    return: Series indexed by CPAT category code (one of
    c.DISTN_INDIRECT_CATEGORIES), the household-demand-weighted average %
    price increase for that category.
    """
    indirect_by_sector = indirect_price_increase_by_gtap_sector(
        price_change_direct, pc_group_weights, io_gtap
    )

    hhd = io_gtap[io_gtap['dis'].str.endswith('.hhd')].set_index('gtap_sector')['value']

    sectors = gtap_cpat_crosswalk.join(indirect_by_sector.rename('price_change'), how='inner')
    sectors = sectors.join(hhd.rename('hhd'), how='left')
    sectors['hhd'] = sectors['hhd'].fillna(0.0)

    # keep only sectors mapped to one of the 14 CPAT *indirect* consumption
    # categories -- this excludes 'NA' (no household mapping), the refined-
    # petroleum sector ('gso, die, ker, lpg', priced directly instead), and
    # the fuel-producing sectors themselves ('coa'/'ely'/'nga'/'oil', which
    # the crosswalk maps to their own fuel code, not a consumption category
    # -- those are the direct-effect channel, not this indirect one).
    sectors = sectors[sectors['CPATCategory'].isin(c.DISTN_INDIRECT_CATEGORIES)]

    def weighted_avg(group):
        total_weight = group['hhd'].sum()
        if total_weight <= 0:
            return group['price_change'].mean()
        return (group['price_change'] * group['hhd']).sum() / total_weight

    return sectors.groupby('CPATCategory').apply(weighted_avg, include_groups=False)


def derive_price_change_direct_from_scenarios(
        country: str,
        baseline: 'ScenarioResults',
        policy: 'ScenarioResults',
        analysis_year: int
        ) -> pd.Series:
    """
    Alternate way to get price_change_direct: from a real Mitigation-module
    run, instead of a supplied/known value. Diffs the policy scenario's
    retail prices (EnergyPrices.rp, ele_prices['rp']) against the baseline
    (no-policy) scenario's, for the analysis year -- see
    cpat_model.scenario_results.ScenarioResults for why two scenarios are
    needed.

    Moved here from distribution.py's old Step-1 implementation, unchanged
    in behaviour: only the fossil fuels' Residential-sector retail price is
    used (households' own price), electricity comes from ele_prices
    instead (EnergyPrices excludes it from `rp`).

    return dims (FuelCode), covering all of c.DISTN_DIRECT_FUELS (biomass
    fuels fixed at 0, see full_direct_price_change).
    """
    y = str(analysis_year)
    idx = pd.IndexSlice

    fossil_fuels = [
        f for f in c.DISTN_DIRECT_FUELS
        if f not in [c.ELE] + c.COOKING_BIOMASS_FUELS
    ]
    baseline_rp = baseline.energy_prices.rp.loc[idx[[country], c.RES, fossil_fuels], y]
    policy_rp = policy.energy_prices.rp.loc[idx[[country], c.RES, fossil_fuels], y]
    price_change_fossil = (policy_rp / baseline_rp) - 1.0
    price_change_fossil.index = price_change_fossil.index.droplevel(c.SECTOR_CODE)

    baseline_ele = baseline.energy_prices.ele_prices['rp'].loc[idx[[country], c.RES, c.ELE], y]
    policy_ele = policy.energy_prices.ele_prices['rp'].loc[idx[[country], c.RES, c.ELE], y]
    price_change_ele = (policy_ele / baseline_ele) - 1.0
    price_change_ele.index = price_change_ele.index.droplevel(c.SECTOR_CODE)

    priced = pd.concat([price_change_fossil, price_change_ele])
    priced.index = priced.index.droplevel(c.COUNTRY_CODE)
    return full_direct_price_change(priced)


def derive_cp_revenue_from_scenarios(country: str, policy: 'ScenarioResults', analysis_year: int) -> float:
    """
    Alternate way to get cp_revenue: carbon price (LCU/tCO2e,
    Policies.cp_trajectory) x taxed emissions (tCO2e, CO2Emissions.total_em),
    both from the policy scenario, for the analysis year. See distribution.py
    module docstring for the same caveat noted when this lived there: a
    first approximation of 'Pm', not validated unit-for-unit against the
    Excel model's own revenue figure.
    """
    y = str(analysis_year)
    return float(policy.policies.cp_trajectory.loc[country, y] * policy.em.total_em.loc[country, y])


def full_direct_price_change(price_change_direct_priced: pd.Series) -> pd.Series:
    """
    Extends a priced-fuel-only price_change_direct Series (covering
    c.COA/ELE/NGA/OOP/GSO/DIE/KER/LPG) with the traditional cooking-biomass
    fuels (c.COOKING_BIOMASS_FUELS), which the model has no price series for
    and so are assumed to see 0 direct price change -- see §5 Step 1 for
    the same assumption made when deriving this from Mitigation prices.

    return dims (FuelCode), covering all of c.DISTN_DIRECT_FUELS
    """
    biomass = pd.Series(0.0, index=c.COOKING_BIOMASS_FUELS)
    return pd.concat([price_change_direct_priced, biomass]).reindex(c.DISTN_DIRECT_FUELS).fillna(0.0)
