"""
Step 8 (pseudocode §5) -- revenue recycling: labor tax/PIT reductions,
targeted transfers, public investment, current spending.

CPAT Excel: 'Distribution' sheet §C.X (transfers/investment/spending),
§C.XI (PIT).

Every function here returns a decile-indexed Series of the *share* of its
channel's revenue pool received by each decile (summing to 1 across the
targeted population, 0 elsewhere) -- distribution.py multiplies by the
pool's actual LCU amount (channel share of cp_revenue). Working in shares
throughout, rather than absolute LCU liabilities, is a deliberate
simplification: the Excel model's Personal Allowance / Targeted Exemption
methods are ultimately capped against each decile's *absolute* baseline PIT
liability (§4.9), which needs a national PIT-revenue-to-GDP figure this
dataset doesn't carry (GDP is Mitigation-side data). The approximations
below reproduce the same redistributive shape -- proportional, capped, or
targeted -- using each decile's *share* of baseline PIT paid instead; see
each function's docstring for exactly how.
"""
import pandas as pd


def normalize_shares(raw_shares: pd.Series) -> pd.Series:
    """HHSurvey's pit_share_income/pit_share_region don't reliably sum to
    100 (per the Excel model's own 'share ... (rescaled)' row, §4.9) --
    rescale to sum to 1."""
    total = raw_shares.sum()
    if total <= 0:
        return pd.Series(1.0 / len(raw_shares), index=raw_shares.index)
    return raw_shares / total


def pit_proportional_compensation(pit_share: pd.Series) -> pd.Series:
    """
    'Proportional Compensation': each decile receives the same share of the
    labor-tax revenue pool as their share of baseline PIT paid.
    """
    return normalize_shares(pit_share)


def pit_personal_allowance(pit_share: pd.Series) -> pd.Series:
    """
    'Personal Allowance': every decile gets, at most, an equal (per-decile)
    share of the pool; any decile whose PIT-paid share would exceed that
    cap has the excess redistributed among deciles still below it. This is
    the share-based analogue of the Excel model's per-capita 'maximum
    transfer' cap (§5 Step 8.1) -- capping shares at 1/n plays the same
    role as capping currency amounts at max_transfer_pc, without requiring
    an absolute PIT baseline.
    """
    shares = normalize_shares(pit_share).copy()
    n = len(shares)
    cap = 1.0 / n
    for _ in range(n): # a few passes converge; a decile freed from the cap can still exceed it once
        over_cap = shares > cap
        if not over_cap.any():
            break
        excess = (shares[over_cap] - cap).sum()
        shares[over_cap] = cap
        under_cap = ~over_cap
        if under_cap.sum() == 0:
            break
        headroom = (cap - shares[under_cap])
        if headroom.sum() <= 0:
            break
        shares[under_cap] = shares[under_cap] + excess * (headroom / headroom.sum())
    return shares


def pit_targeted_exemption(pit_share: pd.Series, exempt_bottom_deciles: int) -> pd.Series:
    """
    'Targeted Exemption': only the bottom N deciles receive anything, in
    proportion to their (relative) baseline PIT-paid share.
    """
    shares = pit_share.copy()
    exempt = [d for d in shares.index if d <= exempt_bottom_deciles]
    out = pd.Series(0.0, index=shares.index)
    if not exempt:
        return out
    out[exempt] = normalize_shares(shares[exempt])
    return out


def pit_reduction_shares(pit_share: pd.Series, method: str, exempt_bottom_deciles: int = 0) -> pd.Series:
    """Dispatches to the method named in DistributionInputsDict['labor_tax_reduction_method']."""
    if method == 'Personal Allowance':
        return pit_personal_allowance(pit_share)
    if method == 'Targeted Exemption':
        return pit_targeted_exemption(pit_share, exempt_bottom_deciles)
    return pit_proportional_compensation(pit_share)


def targeted_transfer_shares(
        population: pd.Series,
        targeted_percentile: float,
        coverage_rate: float,
        leakage_rate: float
        ) -> pd.Series:
    """
    Rules-based synthetic cash transfer (§4.4: 'Revenue Recycling Options ->
    Transfers'), as configured on the CPAT Dashboard -- not an ASPIRE
    existing-program incidence (see data.load_aspire for that alternative
    path, not wired in here).

    targeted_percentile: bottom X% of the population (by decile) is
    'targeted', e.g. 0.4 for the bottom 40%.
    coverage_rate: fraction of the targeted population that actually
    receives a transfer (reduces the pool reaching them, rather than being
    redistributed -- matches the Excel field description: leftover from
    imperfect coverage is not mentioned as being reallocated, unlike PIT's
    personal allowance).
    leakage_rate: fraction of the *untargeted* population that also
    receives a transfer.

    return: Series indexed by decile 1-10, each decile's share of the
    transfer pool (sums to 1, assuming the pool exactly funds
    coverage_rate x targeted population + leakage_rate x untargeted
    population at one equal per-capita rate).
    """
    population = population.sort_index()
    cum_share = population.cumsum() / population.sum()
    prev_cum = cum_share.shift(1, fill_value=0.0)

    # this decile's share of population that falls within the targeted
    # (bottom targeted_percentile) band -- 1.0 if wholly inside, 0.0 if
    # wholly outside, fractional for the decile straddling the cutoff
    within_target = ((targeted_percentile - prev_cum) / (cum_share - prev_cum)).clip(0, 1)

    recipients = population * (within_target * coverage_rate + (1 - within_target) * leakage_rate)
    if recipients.sum() <= 0:
        return pd.Series(0.0, index=population.index)
    return recipients / recipients.sum()


def public_investment_shares(access_index: pd.Series) -> pd.Series:
    """
    Public investment (§4.4): distributed using an infrastructure-access
    incidence weight per decile (HHSurvey's '<type>_acs_share', an index
    relative to the national mean -- see data.load_hh_survey; higher means
    that decile benefits more from access-type investment).
    """
    weights = access_index.clip(lower=0)
    if weights.sum() <= 0:
        return pd.Series(1.0 / len(weights), index=weights.index)
    return weights / weights.sum()


def current_spending_shares(population: pd.Series) -> pd.Series:
    """
    Current spending (§4.4: 'e.g. health, education, social security'):
    no incidence data for this channel exists in
    cpat_excel/Distribution/data_standardized (GDPRatios has health/
    education spend-to-GDP *ratios*, but not their distribution across
    deciles). Falls back to an equal per-capita distribution -- the
    least-informative reasonable default, not a deliberate progressivity
    assumption. Flagged here rather than silently guessed at: refine this
    if/when decile-level current-spending incidence data is found.
    """
    return population / population.sum()
