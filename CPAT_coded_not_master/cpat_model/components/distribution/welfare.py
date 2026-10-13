"""
Steps 9-12 (pseudocode §5) -- post-policy consumption, Gini/Lorenz,
horizontal equity, and the compensation requirement.

CPAT Excel: 'Distribution' sheet §C.VIII-C.IX (post-policy consumption,
Gini), §D.IV (horizontal equity), §D.V (compensation).
"""
import pandas as pd


def post_policy_consumption(
        baseline_consumption: pd.Series,
        total_effect_pct: pd.Series,
        amount_recycled: pd.Series | None = None
        ) -> pd.DataFrame:
    """
    Step 9: consumption before/after the policy, with and without revenue
    recycling, plus each series' (cumulative) share of the total -- the
    Lorenz-curve input for gini_from_shares below.

    baseline_consumption: decile-indexed, adjusted (national-accounts-
    rebased) total consumption -- rebasing.rebase_to_national_accounts's
    'total_consumption' column.
    total_effect_pct: decile-indexed, % of consumption (negative = a cost) --
    effects.total_effect's 'total' column.
    amount_recycled: decile-indexed, LCU recycled to that decile (see
    distribution.py's revenue-recycling wiring), or None to skip the
    'incl. recycling' series.

    return: DataFrame indexed by decile, columns baseline/post_excl_recycling/
    post_incl_recycling (LCU) and their _share / _cumulative_share
    counterparts (deciles ordered as given -- callers should pass deciles
    sorted poorest-to-wealthiest for the cumulative columns to mean anything).
    """
    post_excl = baseline_consumption * (1 + total_effect_pct / 100.0)
    if amount_recycled is not None:
        post_incl = post_excl + amount_recycled.reindex(post_excl.index).fillna(0.0)
    else:
        post_incl = post_excl.copy()

    out = pd.DataFrame({
        'baseline': baseline_consumption,
        'post_excl_recycling': post_excl,
        'post_incl_recycling': post_incl,
    })
    for col in ['baseline', 'post_excl_recycling', 'post_incl_recycling']:
        out[f'{col}_share'] = out[col] / out[col].sum()
        out[f'{col}_cumulative_share'] = out[f'{col}_share'].cumsum()
    return out


def gini_from_shares(population_share: pd.Series, consumption_share: pd.Series) -> float:
    """
    Step 10 / §7.6: Gini coefficient from grouped (decile, or any-size-
    group) data via the trapezoidal-rule area under the Lorenz curve.

    population_share, consumption_share: decile-indexed (or any group
    size), each summing to 1, both sorted poorest -> wealthiest by the
    caller (a plain .sort_index() is enough when the index is 1..10).

    return: Gini coefficient, 0 (perfect equality) to ~1 (perfect inequality).
    """
    cum_pop = population_share.cumsum()
    cum_cons = consumption_share.cumsum()
    prev_pop = cum_pop.shift(1, fill_value=0.0)
    prev_cons = cum_cons.shift(1, fill_value=0.0)

    width = cum_pop - prev_pop
    avg_height = (cum_cons + prev_cons) / 2
    area_under_lorenz = (width * avg_height).sum()
    return 1 - 2 * area_under_lorenz


def horizontal_equity_spread(effect_by_statistic: dict) -> pd.DataFrame:
    """
    Step 11: within-decile spread of the total effect across the p25/mean/
    p75 statistics HHSurvey provides, i.e. how much households *within* the
    same decile differ from one another, not just decile-to-decile.

    effect_by_statistic: {'p25': Series, 'mean': Series, 'p75': Series},
    each decile-indexed (effects.total_effect's 'total' column, computed
    once per statistic).

    return: DataFrame indexed by decile, columns p25/mean/p75.
    """
    return pd.DataFrame(effect_by_statistic)


def compensation_shares(total_effect_pct: pd.Series, consumption_share: pd.Series) -> pd.DataFrame:
    """
    Step 12: the share of total carbon-pricing revenue that would be
    required to exactly offset each decile's loss, and the cumulative
    version (poorest -> wealthiest): 'compensate the bottom N deciles for
    X% of total revenue'.

    total_effect_pct: decile-indexed, % of consumption (negative = a cost).
    consumption_share: decile-indexed, this decile's share of total
    (baseline) consumption, summing to 1.

    return: DataFrame indexed by decile, columns share_of_revenue (sums to
    1 across deciles with a loss) and cumulative_share_of_revenue (assumes
    the index is already sorted poorest -> wealthiest).
    """
    loss = (-total_effect_pct).clip(lower=0) * consumption_share
    total_loss = loss.sum()
    share = loss / total_loss if total_loss > 0 else pd.Series(0.0, index=loss.index)
    return pd.DataFrame({
        'share_of_revenue': share,
        'cumulative_share_of_revenue': share.cumsum(),
    })
