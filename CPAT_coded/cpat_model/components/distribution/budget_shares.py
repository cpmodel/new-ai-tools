"""
Steps 2-3 (pseudocode §5) -- elasticities/DWL and household budget shares.

CPAT Excel: 'Distribution' sheet §C.II (elasticities/DWL), §C.III-C.IV
(budget shares).

Both pull from HHSurvey/HH_Elast, which share one long/tidy shape: one row
per (sample, type, stat_type, quant_cons, variable) cell. select_hh_cells
below is the one place that shape gets pivoted into a (item -> value)
Series, reused by every function in this module.
"""
import pandas as pd

import cpat_model.constants as c


# HHSurvey/HH_Elast 'variable' suffix for each direct fuel / indirect
# category code. Sourced from the actual column values in
# cpat_excel/Distribution/data_standardized (HHSurvey/HH_Elast sheets),
# not invented -- e.g. 'ccl_share' for charcoal, 'health_srv_elasticity'
# for health services, don't follow the fuel/category code exactly.
BUDGET_SHARE_VARIABLE = {
    c.COA: 'coa_share', c.ELE: 'ely_share', c.NGA: 'nga_share', c.OOP: 'oil_share',
    c.GSO: 'gso_share', c.DIE: 'die_share', c.KER: 'ker_share', c.LPG: 'lpg_share',
    c.CHA: 'ccl_share', c.ETH: 'ethanol_share', c.FWD: 'fwd_share',
    c.APP: 'appliances_share', c.CHE: 'chemicals_share', c.CLO: 'clothing_share',
    c.COM: 'communications_share', c.EDU: 'education_share', c.FOOD_CONS: 'food_share',
    c.HEALTH_SRV: 'health_srv_share', c.HOU: 'housing_share', c.OTH: 'other_share',
    c.PAP: 'paper_share', c.PHA: 'pharma_share', c.RET: 'rectourism_share',
    c.TEQ: 'transp_eqt_share', c.TPU: 'transp_pub_share',
}

ELASTICITY_VARIABLE = {
    c.COA: 'coa_elasticity', c.ELE: 'ely_elasticity', c.NGA: 'nga_elasticity', c.OOP: 'oil_elasticity',
    c.GSO: 'gso_elasticity', c.DIE: 'die_elasticity', c.KER: 'ker_elasticity', c.LPG: 'lpg_elasticity',
    c.CHA: 'ccl_elasticity', c.ETH: None, c.FWD: 'fwd_elasticity', # no ethanol elasticity in HH_Elast
    c.APP: 'appliances_elasticity', c.CHE: 'chemicals_elasticity', c.CLO: 'clothing_elasticity',
    c.COM: 'communications_elasticity', c.EDU: 'education_elasticity', c.FOOD_CONS: 'food_elasticity',
    c.HEALTH_SRV: 'health_srv_elasticity', c.HOU: 'housing_elasticity', c.OTH: 'other_elasticity',
    c.PAP: 'paper_elasticity', c.PHA: 'pharma_elasticity', c.RET: 'rectourism_elasticity',
    c.TEQ: 'transp_eqt_elasticity', c.TPU: 'transp_pub_elasticity',
}

BASKET_QUANT_CONS = 9999


def select_hh_cells(
        hh_long: pd.DataFrame,
        sample: str,
        stat_type: str,
        variable_by_item: dict
        ) -> pd.Series:
    """
    Pivots one (sample, stat_type) slice of a HHSurvey/HH_Elast-shaped long
    DataFrame into a decile-indexed Series per item.

    hh_long: data.load_hh_survey() or data.load_hh_elast() output (or
    already country-filtered), columns include sample, type, stat_type,
    quant_cons, variable, value.
    variable_by_item: e.g. BUDGET_SHARE_VARIABLE or ELASTICITY_VARIABLE --
    maps a fuel/category code to the source 'variable' name; a None value
    means "not available", filled with 0.0.

    return: DataFrame indexed by decile (1-10, 'Basket' as index label 0),
    columns = item codes (whichever keys of variable_by_item resolve to a
    real column).
    """
    slice_df = hh_long[(hh_long['sample'] == sample) & (hh_long['stat_type'] == stat_type)]

    out = {}
    for item, var in variable_by_item.items():
        if var is None:
            continue
        item_rows = slice_df[slice_df['variable'] == var]
        if item_rows.empty:
            continue
        series = item_rows.set_index('quant_cons')['value']
        out[item] = series

    df = pd.DataFrame(out)
    # quant_cons 9999 == 'Basket' (national aggregate); relabel to 0 so it
    # sorts before decile 1 rather than after decile 10.
    df = df.rename(index={BASKET_QUANT_CONS: 0}).sort_index()
    return df


def get_budget_shares(hh_survey: pd.DataFrame, sample: str, stat_type: str = 'mean') -> pd.DataFrame:
    """
    Step 3: household budget shares (§C.III-C.IV), % of total consumption,
    for all direct fuels and indirect categories at once.

    return: DataFrame indexed by decile (0='Basket', 1-10), columns = every
    code in c.DISTN_DIRECT_FUELS + c.DISTN_INDIRECT_CATEGORIES for which
    HHSurvey has a *_share variable.
    """
    return select_hh_cells(hh_survey, sample, stat_type, BUDGET_SHARE_VARIABLE)


def get_elasticities(hh_elast: pd.DataFrame) -> pd.DataFrame:
    """
    Step 2 input: own-price elasticities of demand, decile-specific.
    HH_Elast only has an 'Overall' sample and 'mean' statistic (see
    data.load_hh_elast docstring) -- callers needing Urban/Rural or other
    statistics fall back to this Overall/mean series (documented
    approximation, no source data exists for the alternative).

    return: DataFrame indexed by decile (0='Basket', 1-10), columns = every
    code with an elasticity variable in HH_Elast.
    """
    return select_hh_cells(hh_elast, 'Overall', 'mean', ELASTICITY_VARIABLE)


def compute_deadweight_loss(
        budget_shares: pd.DataFrame,
        elasticities: pd.DataFrame,
        price_change: pd.Series
        ) -> pd.DataFrame:
    """
    Step 2.3 / §7.7 Harberger triangle:
    DWL[item, decile] = 0.5 * elasticity[item, decile] * price_change[item]^2 * budget_share[item, decile]

    budget_shares, elasticities: decile-indexed DataFrames (get_budget_shares
    / get_elasticities output), same item columns as price_change's index.
    price_change: Series indexed by item code (fraction, e.g. 0.5 = 50%).

    return: DataFrame, same shape as budget_shares, in the same % units
    (percentage points of consumption).
    """
    items = [i for i in budget_shares.columns if i in elasticities.columns and i in price_change.index]
    dwl = 0.5 * elasticities[items] * (price_change[items] ** 2) * budget_shares[items]
    return dwl
