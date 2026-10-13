import pytest
import pandas as pd
import cpat_model.constants as c

from cpat_model.components.prices.domestic_prices import Headers, DomPrices, MANUAL_PRICES

def test_get_keep_col_condiotion():
    dom_prices = DomPrices.__new__(DomPrices)
    get_keep_col_condiotion = dom_prices._DomPrices__get_keep_col_condiotion # pylint: disable=protected-access

    assert (
        get_keep_col_condiotion(['mit.coa.prod.cost', 'mit.nga.prod.cost'], 'mit.nga.prod.cost')
    ) is True

    assert (
        get_keep_col_condiotion(['mit.coa.prod.cost', 'mit.nga.prod.cost'], 'mit.nga.prod.cosa')
    ) is False

    assert (
        get_keep_col_condiotion('mit.pros.', 'mit.nga.prod.cost')
    ) is False

    assert (
        get_keep_col_condiotion('mit.pros.', 'mit.pros.nga.all')
    ) is True

def test_get_sector_fuel_cols_df():
    dom_prices = DomPrices.__new__(DomPrices)
    code = 'mit.rp.'
    get_sector_fuel_cols_df = dom_prices._DomPrices__get_sector_fuel_cols_df # pylint: disable=protected-access

    df = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 3 + ['MWI'] * 3 ,
            c.FUEL_CODE: ['mit.rp.bio.all', 'mit.rp.coa.res', 'mit.rp.coa.pow'] * 2,
            2021: [float(year) for year in range(1, 7)],
            2022: [float(year) for year in range(7, 13)],
            2023: [float(year) for year in range(13, 19)],
        }
    )
    expected_df = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 3 + ['MWI'] * 3 ,
            c.SECTOR_CODE: [c.ALL, c.RES, c.POW] * 2,
            c.FUEL_CODE: [c.BIO, c.COA, c.COA] * 2,
            '2021': [float(year) for year in range(1, 7)],
            '2022': [float(year) for year in range(7, 13)],
            '2023': [float(year) for year in range(13, 19)],
        }
    ).set_index(c.ID_COL_NAMES)
    pd.testing.assert_frame_equal(
        get_sector_fuel_cols_df(df, code),
        expected_df
    )

def test_cut_years_prior_to_start_year():
    dom_prices = DomPrices.__new__(DomPrices)
    cut_years_prior_to_start_year = dom_prices._DomPrices__cut_years_prior_to_start_year # pylint: disable=protected-access

    simulation_years = (2022, 2023)
    df = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 3 + ['MWI'] * 3 ,
            c.SECTOR_CODE: [c.ALL, c.RES, c.POW] * 2,
            c.FUEL_CODE: [c.BIO, c.COA, c.COA] * 2,
            '2019': [float(year) for year in range(1, 7)],
            '2020': [float(year) for year in range(7, 13)],
            '2021': [float(year) for year in range(13, 19)],
            '2022': [float(year) for year in range(19, 25)],
            '2023': [float(year) for year in range(25, 31)],
        }
    ).set_index(c.ID_COL_NAMES)

    expected_df = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 3 + ['MWI'] * 3 ,
            c.SECTOR_CODE: [c.ALL, c.RES, c.POW] * 2,
            c.FUEL_CODE: [c.BIO, c.COA, c.COA] * 2,
            '2021': [float(year) for year in range(13, 19)],
            '2022': [float(year) for year in range(19, 25)],
            '2023': [float(year) for year in range(25, 31)],
        }
    ).set_index(c.ID_COL_NAMES)
    pd.testing.assert_frame_equal(
        cut_years_prior_to_start_year(simulation_years, df),
        expected_df
    )


def test_set_manual_price_control_coefficients():
    # TODO: update after implementing Manual inputs?
    selected_countries = ['USA', 'ALB']
    dom_prices = DomPrices.__new__(DomPrices)
    dom_prices.forecasting_coefficients = {}
    set_manual_price_control_coefficients = (
        dom_prices._DomPrices__set_manual_price_control_coefficients # pylint: disable=protected-access
    )

    expected_df = pd.DataFrame(
        0.8,
        index=pd.MultiIndex.from_tuples(
            [
                (country, sector, fuel)
                for country in sorted(selected_countries)
                for sector, fuel in [
                    *[(c.ALL, fuel) for fuel in [c.BIO, c.DIE, c.GSO, c.KER, c.LPG, c.OOP]],
                    *[(c.IND, fuel) for fuel in [c.COA, c.ELE, c.NGA]],
                    *[(c.POW, fuel) for fuel in [c.COA, c.NGA]],
                    *[(c.RES, fuel) for fuel in [c.COA, c.ELE, c.NGA]]
                ]
            ],
            names=c.ID_COL_NAMES
        ),
        columns=[Headers.MANUAL_PRICE_CONTROL_COEFFICIENTS]
    )
    set_manual_price_control_coefficients(selected_countries, MANUAL_PRICES)
    pd.testing.assert_frame_equal(
        dom_prices.forecasting_coefficients[Headers.MANUAL_PRICE_CONTROL_COEFFICIENTS],
        expected_df
    )

def test_set_chosen_price_control_coefficient():
    dom_prices = DomPrices.__new__(DomPrices)
    dom_prices.forecasting_coefficients = {}
    set_chosen_price_control_coefficient = (
        dom_prices._DomPrices__set_chosen_price_control_coefficient # pylint: disable=protected-access
    )
    selected_countries = ['USA', 'ALB']

    index = pd.MultiIndex.from_tuples(
        [
            (country, sector, fuel)
            for country in sorted(selected_countries)
            for sector, fuel in [
                *[(c.ALL, fuel) for fuel in [c.BIO, c.DIE, c.GSO, c.KER, c.LPG, c.OOP]],
                *[(c.IND, fuel) for fuel in [c.COA, c.ELE, c.NGA]],
                *[(c.POW, fuel) for fuel in [c.COA, c.NGA]],
                *[(c.RES, fuel) for fuel in [c.COA, c.ELE, c.NGA]]
            ]
        ],
        names=c.ID_COL_NAMES
    )

    dom_prices.forecasting_coefficients[Headers.BUCKETED_PRICE_CONTROL_COEFFICIENT] = pd.DataFrame(
        {
            Headers.BUCKETED_PRICE_CONTROL_COEFFICIENT: (
                [1.0, 0.5] + [1.0] * 18 + [0.0] + [1.0] * 2 + [0.0, 1.0, 0.0, 1.0, 0.5]
            )
        },
        index=index
    )
    dom_prices.forecasting_coefficients[Headers.MANUAL_PRICE_CONTROL_COEFFICIENTS] = pd.DataFrame(
        0.8,
        index=index,
        columns=[Headers.MANUAL_PRICE_CONTROL_COEFFICIENTS]
    )

    expected_df_1 = pd.DataFrame(
        {
            Headers.CHOSEN_PRICE_CONTROL_COEFFICIENT: (
                [1.0, 0.5] + [1.0] * 18 + [0.0] + [1.0] * 2 + [0.0, 1.0, 0.0, 1.0, 0.5]
            )
        },
        index=index
    )
    gov_energy_price_controls = 'Bucketed' # ['Bucketed', 'Manual', 'None']
    set_chosen_price_control_coefficient(gov_energy_price_controls)
    pd.testing.assert_frame_equal(
        dom_prices.forecasting_coefficients[Headers.CHOSEN_PRICE_CONTROL_COEFFICIENT],
        expected_df_1
    )

    expected_df_2 = pd.DataFrame(
        0.8,
        index=index,
        columns=[Headers.CHOSEN_PRICE_CONTROL_COEFFICIENT]
    )
    gov_energy_price_controls = 'Manual'
    set_chosen_price_control_coefficient(gov_energy_price_controls)
    pd.testing.assert_frame_equal(
        dom_prices.forecasting_coefficients[Headers.CHOSEN_PRICE_CONTROL_COEFFICIENT],
        expected_df_2
    )

    expected_df_3 = pd.DataFrame(
        1.0,
        index=index,
        columns=[Headers.CHOSEN_PRICE_CONTROL_COEFFICIENT]
    )
    gov_energy_price_controls = 'None'
    set_chosen_price_control_coefficient(gov_energy_price_controls)
    pd.testing.assert_frame_equal(
        dom_prices.forecasting_coefficients[Headers.CHOSEN_PRICE_CONTROL_COEFFICIENT],
        expected_df_3
    )

    gov_energy_price_controls = 'Incorrect str'
    with pytest.raises(
        TypeError,
        match="Wrong value for gov_energy_price_controls."
    ):
        set_chosen_price_control_coefficient(gov_energy_price_controls)
