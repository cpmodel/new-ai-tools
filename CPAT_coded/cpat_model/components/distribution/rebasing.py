"""
Step 7 (pseudocode §5) -- household-survey-to-national-accounts rebasing.

CPAT Excel: 'Distribution' sheet §C.VIII.

Household budget surveys under-report total consumption relative to
national accounts (GDP-based) estimates, so results are rescaled onto a
national total before reporting currency-denominated (not just %) burdens.

This module takes the *target* national totals (population, consumption) as
plain parameters rather than deriving them from GDP here, because the GDP
time series and deflator chain those targets are normally built from is
Mitigation-side data (see cpat_model.components.gdp.gdp.GDP), not part of
cpat_excel/Distribution/data_standardized. Deriving them is thus the
caller's job: from a real GDP run when one is available (GDP.population,
GDP.ngdp_d for the deflator chain, and GDPRatios' consumption-to-GDP ratio
loaded here), or from a known/assumed total otherwise.
"""
import pandas as pd


def rebase_to_national_accounts(
        survey_population: pd.Series,
        survey_per_capita_consumption: pd.Series,
        national_total_population: float,
        national_total_consumption: float
        ) -> pd.DataFrame:
    """
    Scales the household survey's decile-level population and per-capita
    consumption so their totals match given national-accounts figures,
    preserving the survey's relative distribution across deciles.

    survey_population: HHSurvey 'popw', decile-indexed (1-10; a 'Basket'/0
    row, if present, is ignored -- only the 10 deciles are rescaled).
    survey_per_capita_consumption: HHSurvey 'cons_pc_acrent', same index.
    national_total_population, national_total_consumption: the year's
    actual totals to rebase onto (see module docstring for where these
    come from).

    return: DataFrame indexed by decile 1-10, columns:
    - population: adjusted headcount
    - per_capita_consumption: adjusted per-capita consumption
    - total_consumption: adjusted total consumption (population x per-capita)
    - consumption_share: this decile's share of total_consumption
    """
    deciles = [d for d in survey_population.index if d != 0]
    pop = survey_population.reindex(deciles)
    pc_cons = survey_per_capita_consumption.reindex(deciles)
    survey_total_consumption = pop * pc_cons

    population_adj_factor = national_total_population / pop.sum()
    consumption_adj_factor = national_total_consumption / survey_total_consumption.sum()

    adjusted_population = pop * population_adj_factor
    adjusted_total_consumption = survey_total_consumption * consumption_adj_factor
    adjusted_per_capita = adjusted_total_consumption / adjusted_population

    return pd.DataFrame({
        'population': adjusted_population,
        'per_capita_consumption': adjusted_per_capita,
        'total_consumption': adjusted_total_consumption,
        'consumption_share': adjusted_total_consumption / adjusted_total_consumption.sum(),
    })


def household_consumption_to_gdp_ratio(gdp_ratios: pd.DataFrame, latest_year: int | None = None) -> float:
    """
    National Accounts Household Consumption-to-GDP ratio (§4.6), the
    'assumed constant' proportion of GDP that is household consumption --
    used, together with a real GDP figure, to derive
    national_total_consumption for rebase_to_national_accounts above.

    gdp_ratios: data.load_gdp_ratios() output, this country's slice.
    latest_year: use this year's ratio if available, else the most recent
    year with data.

    return: ratio (0-1, e.g. 0.83 for 83% of GDP).
    """
    series = gdp_ratios[
        gdp_ratios['Indicator Name'] == 'Households and NPISHs final consumption expenditure (% of GDP)'
    ].dropna(subset=['value'])
    if latest_year is not None and (series['year'] == latest_year).any():
        value = series.loc[series['year'] == latest_year, 'value'].iloc[0]
    else:
        value = series.sort_values('year')['value'].iloc[-1]
    return value / 100.0
