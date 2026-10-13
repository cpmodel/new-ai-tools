import pandas as pd

import cpat_model.constants as c

from cpat_model.components.prices.international_prices import (
    IntPrices
)


def test_set_inv_deflator_growth():
    deflator = pd.DataFrame(
        [
            [1.01, 1.0, 0.99],
            [1.02, 1.0, 0.98]
        ],
        columns=[str(year) for year in range(2022, 2025)],
        index=['USA', 'MWI']
    )
    deflator.index.name = c.COUNTRY_CODE
    int_prices = IntPrices.__new__(IntPrices)
    set_inv_deflator_growth = int_prices._IntPrices__set_inv_deflator_growth # pylint: disable=protected-access
    set_inv_deflator_growth(deflator)

    expected_df_1 = pd.DataFrame(
        [
            [0.990099, 1.0, 1.010101],
            [0.980392, 1.0, 1.020408]
        ],
        columns=[str(year) for year in range(2022, 2025)],
        index=['USA', 'MWI']
    )
    expected_df_1.index.name = c.COUNTRY_CODE

    pd.testing.assert_frame_equal(
        int_prices._IntPrices__inv_deflator, # pylint: disable=protected-access
        expected_df_1
    )

    expected_df_2 = pd.DataFrame(
        [
            [0.01, 0.010101],
            [0.02, 0.020408]
        ],
        columns=[str(year) for year in range(2023, 2025)],
        index=['USA', 'MWI']
    )
    expected_df_2.index.name = c.COUNTRY_CODE
    pd.testing.assert_frame_equal(
        int_prices._IntPrices__inv_deflator_growth, # pylint: disable=protected-access
        expected_df_2
    )

def test_extend_by_country():
    prices = pd.DataFrame(
        {
            c.FUEL_CODE: [c.OIL, c.COA] + [c.NGA] * 3,
            'Region': [None, None, 'EU', 'US', 'LNG'],
            '2020': [float(i) for i in range(1, 6)],
            '2021': [float(i) for i in range(6, 11)],
            '2022': [float(i) for i in range(11, 16)]
        }
    ).set_index([c.FUEL_CODE, 'Region'])
    selected_countries = ["USA", 'MWI', 'DEU']
    int_prices = IntPrices.__new__(IntPrices)
    extend_by_country = int_prices._IntPrices__extend_by_country # pylint: disable=protected-access

    expected_df = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 5 + ['MWI'] * 5 + ['USA'] * 5,
            c.FUEL_CODE: ([c.COA] + [c.NGA] * 3 + [c.OIL]) * 3,
            'Region': [None, 'EU', 'LNG', 'US', None] * 3,
            '2020': [2.0, 3.0, 5.0, 4.0, 1.0] * 3,
            '2021': [7.0, 8.0, 10.0, 9.0, 6.0] * 3,
            '2022': [12.0, 13.0, 15.0, 14.0, 11.0] * 3
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE, 'Region'])
    pd.testing.assert_frame_equal(
        extend_by_country(prices, selected_countries),
        expected_df
    )


def test_cut_years():
    prices = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 5 + ['MWI'] * 5,
            c.FUEL_CODE: ([c.COA] + [c.NGA] * 3 + [c.OIL]) * 2,
            'Region': [None, 'EU', 'LNG', 'US', None] * 2,
            '2029': [2.0, 3.0, 5.0, 4.0, 1.0] * 2,
            '2030': [
                2.024, 3.036, 5.06, 4.048, 1.012,
                2.044, 3.066, 5.11, 4.088, 1.022
            ],
            '2031': [None] * 5 + [0.1] * 5,
            '2032': [None] * 10
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE, 'Region'])
    simulation_years = (2029, 2030)
    int_prices = IntPrices.__new__(IntPrices)
    cut_years = int_prices._IntPrices__cut_years # pylint: disable=protected-access

    expected_df = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 5 + ['MWI'] * 5,
            c.FUEL_CODE: ([c.COA] + [c.NGA] * 3 + [c.OIL]) * 2,
            'Region': [None, 'EU', 'LNG', 'US', None] * 2,
            '2029': [2.0, 3.0, 5.0, 4.0, 1.0] * 2,
            '2030': [
                2.024, 3.036, 5.06, 4.048, 1.012,
                2.044, 3.066, 5.11, 4.088, 1.022
            ],
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE, 'Region'])
    pd.testing.assert_frame_equal(
        cut_years(prices, simulation_years),
        expected_df
    )


def test_add_nga_global():
    prices = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 5 + ['MWI'] * 5,
            c.FUEL_CODE: ([c.COA] + [c.NGA] * 3 + [c.OIL]) * 2,
            'Region': [None, 'EU', 'LNG', 'US', None] * 2,
            '2029': [2.0, 3.0, 5.0, 4.0, 1.0] * 2,
            '2030': [
                2.024, 3.036, 5.06, 4.048, 1.012,
                2.044, 3.066, 5.11, 4.088, 1.022
            ],
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE, 'Region'])

    int_prices = IntPrices.__new__(IntPrices)
    add_nga_global = int_prices._IntPrices__add_nga_global # pylint: disable=protected-access

    expected_df = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 6 + ['MWI'] * 6,
            c.FUEL_CODE: ([c.COA] + [c.NGA] * 4 + [c.OIL]) * 2,
            'Region': [None, 'EU', 'Global', 'LNG', 'US', None] * 2,
            '2029': [2.0, 3.0, 4.0, 5.0, 4.0, 1.0] * 2,
            '2030': [
                2.024, 3.036, 4.048, 5.06, 4.048, 1.012,
                2.044, 3.066, 4.088, 5.11, 4.088, 1.022
            ],
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE, 'Region'])
    pd.testing.assert_frame_equal(
        add_nga_global(prices),
        expected_df
    )


def test_forecast_early_coa():
    int_prices = IntPrices.__new__(IntPrices)
    simulation_years = (2028, 2032)

    # pylint: disable=protected-access
    forecast_early_coa = int_prices._IntPrices__forecast_early_coa

    int_prices._IntPrices__inv_deflator_growth = pd.DataFrame(
        [
            [-0.01, 0.0, 0.02, 0.01],
            [-0.02, 0.0, 0.01, 0.02],
        ],
        columns=[str(year) for year in range(2029, 2032 + 1)],
        index=['USA', 'MWI']
    )
    int_prices._IntPrices__inv_deflator_growth.index.name = c.COUNTRY_CODE
    # pylint: enable=protected-access
    year_values = [
        2.1, 3.0, 4.0, 5.0, 4.0, 1.1,
        2.2, 3.0, 4.0, 5.1, 4.8, 1.2
    ]
    prices = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['USA'] * 6 + ['MWI'] * 6,
            c.FUEL_CODE: ([c.COA] + [c.NGA] * 4 + [c.OIL]) * 2,
            'Region': [None, 'EU', 'Global', 'LNG', 'US', None] * 2,
            '2029': [2.0, 3.0, 4.0, 5.0, 4.0, 1.0] * 2,
            '2030': year_values,
            '2031': year_values,
            '2032': year_values,
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE, 'Region'])

    expected_df = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['USA'] * 6 + ['MWI'] * 6,
            c.FUEL_CODE: ([c.COA] + [c.NGA] * 4 + [c.OIL]) * 2,
            'Region': [None, 'EU', 'Global', 'LNG', 'US', None] * 2,
            '2029': [2.0, 3.0, 4.0, 5.0, 4.0, 1.0] * 2,
            '2030': year_values,
            '2031': [2.142] + year_values[1:6] + [2.222] + year_values[7:],
            '2032': [2.16342] + year_values[1:6] + [2.26644] + year_values[7:]
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE, 'Region'])
    pd.testing.assert_frame_equal(
        forecast_early_coa(prices, simulation_years),
        expected_df
    )

def test_apply_fuel_adjustment():
    int_prices = IntPrices.__new__(IntPrices)
    # pylint: disable=protected-access
    apply_fuel_adjustment = int_prices._IntPrices__apply_fuel_adjustment
    int_prices._IntPrices__fist_other_source_year = 2024
    # pylint: enable=protected-access
    simulation_years = (2024, 2026)
    prices = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 3 + ['MWI'] * 3,
            c.FUEL_CODE: [c.COA, c.NGA, c.OIL] * 2,
            '2024': [float(year) for year in range(1, 4)] * 2,
            '2025': [float(year) for year in range(1, 4)] * 2,
            '2026': [float(year) for year in range(1, 4)] * 2,
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE])

    int_prices.prices = prices.copy()
    int_energy_price_adjustment = 'High'
    expected_df_1 = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 3 + ['MWI'] * 3,
            c.FUEL_CODE: [c.COA, c.NGA, c.OIL] * 2,
            '2024': [float(year) for year in range(1, 4)] * 2,
            '2025': [1.25, 2.5, 4.5] * 2,
            '2026': [1.25, 2.5, 4.5] * 2,
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE])
    apply_fuel_adjustment(
        int_energy_price_adjustment, simulation_years
    )
    pd.testing.assert_frame_equal(
        int_prices.prices,
        expected_df_1
    )

    int_prices.prices = prices.copy()
    int_energy_price_adjustment = 'Low'
    expected_df_2 = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 3 + ['MWI'] * 3,
            c.FUEL_CODE: [c.COA, c.NGA, c.OIL] * 2,
            '2024': [float(year) for year in range(1, 4)] * 2,
            '2025': [0.75, 1.5, 1.5] * 2,
            '2026': [0.75, 1.5, 1.5] * 2,
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE])
    apply_fuel_adjustment(
        int_energy_price_adjustment, simulation_years
    )
    pd.testing.assert_frame_equal(
        int_prices.prices,
        expected_df_2
    )



def test_apply_deflator():
    int_prices = IntPrices.__new__(IntPrices)
    apply_deflator = int_prices._IntPrices__apply_deflator # pylint: disable=protected-access
    int_prices.prices = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 3 + ['MWI'] * 3,
            c.FUEL_CODE: [c.COA, c.NGA, c.OIL] * 2,
            '2020': [float(year) for year in range(1, 4)] * 2,
            '2021': [float(year) for year in range(1, 4)] * 2,
            '2022': [float(year) for year in range(1, 4)] * 2,
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE])
    deflator = pd.DataFrame(
        [
            [1.1, 1.0, 0.9],
            [1.2, 1.0, 0.8]
        ],
        columns=[str(year) for year in range(2020, 2022 + 1)],
        index=['DEU', 'MWI']
    )
    deflator.index.name = c.COUNTRY_CODE

    expected_df = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 3 + ['MWI'] * 3,
            c.FUEL_CODE: [c.COA, c.NGA, c.OIL] * 2,
            '2020': [1.1, 2.2, 3.3, 1.2, 2.4, 3.6],
            '2021': [float(year) for year in range(1, 4)] * 2,
            '2022': [0.9, 1.8, 2.7, 0.8, 1.6, 2.4],
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE])
    apply_deflator(deflator)
    pd.testing.assert_frame_equal(
        int_prices.prices,
        expected_df
    )


def test_conv_to_gj():
    int_prices = IntPrices.__new__(IntPrices)
    int_prices.prices = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 3 + ['MWI'] * 3,
            c.FUEL_CODE: [c.COA, c.NGA, c.OIL] * 2,
            '2024': [float(year) for year in range(1, 4)] * 2,
            '2025': [float(year) for year in range(1, 4)] * 2,
            '2026': [float(year) for year in range(1, 4)] * 2,
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE])
    conv_to_gj = int_prices._IntPrices__conv_to_gj # pylint: disable=protected-access

    expected_df = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 3 + ['MWI'] * 3,
            c.FUEL_CODE: [c.COA, c.NGA, c.OIL] * 2,
            '2024': [0.039777, 1.895634, 0.490196] * 2,
            '2025': [0.039777, 1.895634, 0.490196] * 2,
            '2026': [0.039777, 1.895634, 0.490196] * 2,
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE])
    conv_to_gj()
    pd.testing.assert_frame_equal(
        int_prices.prices,
        expected_df
    )
