from types import SimpleNamespace

import pytest
import pandas as pd

from cpat_model.components.distribution import price_changes
from cpat_model.components.distribution.data import load_io_gtap, load_gtap_cpat_sector_crosswalk
import cpat_model.constants as c


YEAR = 2030
Y = str(YEAR)
COUNTRY = 'EGY'


def _rp_df(values: dict) -> pd.DataFrame:
    """
    Builds a (CountryCode, SectorCode, FuelCode)-indexed, single-year-column
    DataFrame like EnergyPrices.rp / ele_prices['rp'], from
    {(country, sector, fuel): price}.
    """
    rows = [(country, sector, fuel, price) for (country, sector, fuel), price in values.items()]
    df = pd.DataFrame(rows, columns=[c.COUNTRY_CODE, c.SECTOR_CODE, c.FUEL_CODE, Y])
    return df.set_index(c.ID_COL_NAMES)


def _scenario_results(rp: pd.DataFrame, ele_rp: pd.DataFrame, cp: float | None = None, em: float | None = None):
    """
    Minimal stand-in for ScenarioResults carrying just what
    derive_price_change_direct_from_scenarios/derive_cp_revenue_from_scenarios
    read: energy_prices.rp, energy_prices.ele_prices['rp'],
    policies.cp_trajectory, em.total_em. Avoids constructing real
    EnergyPrices/Policies/CO2Emissions objects, which need full InputData.
    """
    energy_prices = SimpleNamespace(rp=rp, ele_prices={'rp': ele_rp})
    policies = SimpleNamespace(
        cp_trajectory=pd.DataFrame({Y: [cp]}, index=pd.Index([COUNTRY], name=c.COUNTRY_CODE))
        if cp is not None else None
    )
    em_obj = SimpleNamespace(
        total_em=pd.DataFrame({Y: [em]}, index=pd.Index([COUNTRY], name=c.COUNTRY_CODE))
        if em is not None else None
    )
    return SimpleNamespace(energy_prices=energy_prices, policies=policies, em=em_obj)


def test_derive_price_change_direct_from_scenarios():
    fossil_fuels = [f for f in c.DISTN_DIRECT_FUELS if f not in [c.ELE] + c.COOKING_BIOMASS_FUELS]

    baseline_rp = _rp_df({(COUNTRY, c.RES, f): 10.0 for f in fossil_fuels})
    policy_rp = _rp_df({(COUNTRY, c.RES, f): 11.0 for f in fossil_fuels}) # +10% for every fossil fuel

    baseline_ele = _rp_df({(COUNTRY, c.RES, c.ELE): 5.0})
    policy_ele = _rp_df({(COUNTRY, c.RES, c.ELE): 6.0}) # +20%

    baseline = _scenario_results(baseline_rp, baseline_ele)
    policy = _scenario_results(policy_rp, policy_ele)

    price_change_direct = price_changes.derive_price_change_direct_from_scenarios(
        COUNTRY, baseline, policy, YEAR
    )

    for f in fossil_fuels:
        assert price_change_direct[f] == pytest.approx(0.1)
    assert price_change_direct[c.ELE] == pytest.approx(0.2)
    for f in c.COOKING_BIOMASS_FUELS:
        assert price_change_direct[f] == 0.0

    assert sorted(price_change_direct.index) == sorted(c.DISTN_DIRECT_FUELS)


def test_derive_cp_revenue_from_scenarios():
    baseline = _scenario_results(_rp_df({}), _rp_df({}))
    policy = _scenario_results(_rp_df({}), _rp_df({}), cp=50.0, em=1_000.0)

    cp_revenue = price_changes.derive_cp_revenue_from_scenarios(COUNTRY, policy, YEAR)

    assert cp_revenue == pytest.approx(50_000.0)


def test_price_change_for_pc_sector_weighted_average():
    price_change_direct = pd.Series({c.GSO: 0.2, c.DIE: 0.4, c.KER: 0.0, c.LPG: 0.6})
    weights = pd.Series({c.GSO: 1.0, c.DIE: 1.0, c.KER: 0.0, c.LPG: 2.0})

    pc_price = price_changes.price_change_for_pc_sector(price_change_direct, weights)

    # (0.2*1 + 0.4*1 + 0*0 + 0.6*2) / (1+1+0+2) = 1.8/4 = 0.45
    assert pc_price == pytest.approx(0.45)


def test_aggregate_price_change_to_categories_uses_real_egypt_data():
    """
    Structural check against the real Egypt IO_GTAP + GTAP-CPAT crosswalk
    (cpat_excel/Distribution/data_standardized): every indirect category
    should come back, all non-negative for a price *increase* scenario, and
    a fuel-heavy category (public transportation) should show a bigger
    embedded-energy shock than a fuel-light one (communications) -- this is
    the one relative-ranking check that doesn't depend on the pass-through
    coefficients this dataset is missing (see price_changes.py module
    docstring), only on the Leontief coefficients' relative sizes.
    """
    io_gtap = load_io_gtap(['EGY'])
    crosswalk = load_gtap_cpat_sector_crosswalk()

    price_change_direct = pd.Series({
        c.COA: 0.38, c.ELE: 0.25, c.NGA: 0.59, c.OOP: 1.04,
        c.GSO: 0.19, c.DIE: 0.28, c.KER: 0.30, c.LPG: 0.59,
    })
    pc_weights = pd.Series({c.GSO: 0.53, c.DIE: 0.0, c.KER: 0.0001, c.LPG: 1.76})

    result = price_changes.aggregate_price_change_to_categories(
        price_change_direct, pc_weights, io_gtap, crosswalk
    )

    assert set(result.index) == set(c.DISTN_INDIRECT_CATEGORIES)
    assert (result >= 0).all()
    assert result[c.TPU] > result[c.COM]
