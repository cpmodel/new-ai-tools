"""
Steps 4-6 (pseudocode §5) -- direct, indirect and total consumption-loss
effects.

CPAT Excel: 'Distribution' sheet §C.V (direct), §C.VI (indirect), §C.VII
(total).
"""
import pandas as pd

import cpat_model.constants as c


def direct_effect(
        budget_shares: pd.DataFrame,
        price_change_direct: pd.Series,
        dwl: pd.DataFrame | None = None,
        behavior_adjustment: float = 1.0
        ) -> pd.DataFrame:
    """
    Step 4: direct_effect[fuel, decile] =
        budget_share[fuel, decile] * price_change_direct[fuel] * behavior_adjustment
        - dwl[fuel, decile]

    budget_shares: decile-indexed DataFrame (budget_shares.get_budget_shares
    output), columns limited to fuel codes here by the caller (or containing
    extra indirect-category columns, which are ignored -- only columns also
    present in price_change_direct's index are used).
    price_change_direct: Series indexed by fuel code, fraction (0.5 = 50%).
    dwl: same shape as budget_shares (budget_shares.compute_deadweight_loss
    output), or None to skip the deadweight-loss adjustment.
    behavior_adjustment: scalar multiplier (DistributionInputsDict
    'behavioral_response_adj_factor' when 'adjust_for_behavioral_change' is
    True, else 1.0).

    return: DataFrame, decile-indexed, one column per fuel, % of consumption
    (negative = a cost).
    """
    fuels = [f for f in c.DISTN_DIRECT_FUELS if f in budget_shares.columns and f in price_change_direct.index]
    effect = -1 * budget_shares[fuels] * price_change_direct[fuels] * behavior_adjustment
    if dwl is not None:
        dwl_fuels = [f for f in fuels if f in dwl.columns]
        effect[dwl_fuels] = effect[dwl_fuels] - dwl[dwl_fuels]
    return effect


def apply_cooking_fuel_exemption(
        direct_effect_df: pd.DataFrame,
        exempt_fuel: str,
        exempt_share: float,
        exempt_bottom_deciles: int
        ) -> pd.DataFrame:
    """
    Step 4.2 (§4.5 exemption inputs): zero out (a share of) the direct
    effect of the most-used cooking fuel for the bottom N deciles.

    exempt_fuel: fuel code (from data.load_who_cooking's 'main_fuel' value,
    or a user override).
    exempt_share: 0-1, portion of that fuel's effect treated as exempt.
    exempt_bottom_deciles: deciles 1..N get the exemption; 0 = no exemption.
    """
    out = direct_effect_df.copy()
    if exempt_fuel not in out.columns or exempt_bottom_deciles <= 0:
        return out
    deciles = [d for d in out.index if isinstance(d, (int, float)) and 1 <= d <= exempt_bottom_deciles]
    out.loc[deciles, exempt_fuel] = out.loc[deciles, exempt_fuel] * (1 - exempt_share)
    return out


def indirect_effect(
        budget_shares: pd.DataFrame,
        price_change_indirect: pd.Series,
        dwl: pd.DataFrame | None = None,
        behavior_adjustment: float = 1.0
        ) -> pd.DataFrame:
    """
    Step 5: same formula as direct_effect, over the indirect consumption
    categories instead of direct fuels.

    return: DataFrame, decile-indexed, one column per category, % of
    consumption.
    """
    categories = [
        cat for cat in c.DISTN_INDIRECT_CATEGORIES
        if cat in budget_shares.columns and cat in price_change_indirect.index
    ]
    effect = -1 * budget_shares[categories] * price_change_indirect[categories] * behavior_adjustment
    if dwl is not None:
        dwl_cats = [cat for cat in categories if cat in dwl.columns]
        effect[dwl_cats] = effect[dwl_cats] - dwl[dwl_cats]
    return effect


def total_effect(direct_effect_df: pd.DataFrame, indirect_effect_df: pd.DataFrame) -> pd.DataFrame:
    """
    Step 6: total_effect[decile] = sum_fuel(direct_effect) + sum_category(indirect_effect)

    return: DataFrame, decile-indexed, columns ['direct', 'indirect', 'total'],
    % of consumption.
    """
    direct_total = direct_effect_df.sum(axis=1)
    indirect_total = indirect_effect_df.sum(axis=1)
    return pd.DataFrame({
        'direct': direct_total,
        'indirect': indirect_total,
        'total': direct_total + indirect_total,
    })
