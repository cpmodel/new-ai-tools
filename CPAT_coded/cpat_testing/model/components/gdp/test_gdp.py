import pytest
import pandas as pd

from cpat_model.components.gdp.gdp import (
    GDP
)
import cpat_model.constants as c

def test_set_population_growth():
    gdp = GDP.__new__(GDP)
    gdp.population = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['ASM', 'VIR'],
            "2020": [0.047, 0.101],
            "2021": [0.048, 0.2],
            "2022": [0.050, 0.21],
            "2023": [0.051, 0.23],
            "2024": [0.051, 0.24],
            "2025": [0.043, 0.2]
        }
    ).set_index(c.COUNTRY_CODE)
    set_population_growth = gdp._GDP__set_population_growth # pylint: disable=protected-access

    expected_df = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['ASM', 'VIR'],
            "2021": [0.0212765957446807, 0.98019801980198],
            "2022": [0.0416666666666667, 0.05],
            "2023": [0.02, 0.0952380952380953],
            "2024": [0.0, 0.0434782608695651],
            "2025": [-0.156862745098039, -0.166666666666666]
        }
    ).set_index(c.COUNTRY_CODE)
    set_population_growth()
    pd.testing.assert_frame_equal(
        gdp.population_growth,
        expected_df
    )


def test_project_ngdp_d():
    gdp = GDP.__new__(GDP)
    project_ngdp_d = gdp._GDP__project_ngdp_d # pylint: disable=protected-access
    gdp.ngdp_d = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['AAA', 'BBB'],
            "2019": [199.0, 299.2],
            "2020": [1.0, 2.2],
            "2021": [1.1, 2.25],
            "2022": [1.2, 2.3],
            "2023": [1.3, 2.35],
            "2024": [1.4, 2.3],
            "2025": [1.5, 2.25]
        }
    ).set_index(c.COUNTRY_CODE)
    project_ngdp_d((2023, 2030))

    expected_df = pd.DataFrame(
        [
            [1.10, 1.2, 1.30, 1.4, 1.50, 1.6, 1.70, 1.8, 1.90, 2.0],
            [2.25, 2.3, 2.35, 2.3, 2.25, 2.2, 2.15, 2.1, 2.05, 2.0]
        ],
        columns=[str(year) for year in range(2021, 2031)],
        index=["AAA", 'BBB']
    )
    expected_df.index.name = c.COUNTRY_CODE
    pd.testing.assert_frame_equal(
        gdp.ngdp_d,
        expected_df
    )

def test_set_ngdp_d_world():
    gdp = GDP.__new__(GDP)
    set_ngdp_d_world = gdp._GDP__set_ngdp_d_world # pylint: disable=protected-access
    ngdp_d_world = pd.DataFrame(
        [
            [299.2, 2.2, 2.25, 2.3, 2.35, 2.3, 2.25]
        ],
        columns=[str(year) for year in range(2019, 2025 + 1)],
        index=["USA"]
    )
    ngdp_d_world.index.name = c.COUNTRY_CODE

    set_ngdp_d_world(ngdp_d_world, (2023, 2030))

    expected_df = pd.DataFrame(
        [
            [2.25, 2.3, 2.35, 2.3, 2.25, 2.2, 2.15, 2.1, 2.05, 2.0]
        ],
        columns=[str(year) for year in range(2021, 2030 + 1)],
        index=["USA"]
    )
    expected_df.index.name = c.COUNTRY_CODE
    pd.testing.assert_frame_equal(
        gdp.ngdp_d_world,
        expected_df
    )

def test_set_deflator():
    gdp = GDP.__new__(GDP)
    set_deflator = gdp._GDP__set_deflator # pylint: disable=protected-access
    gdp.ngdp_d = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['AAA', 'BBB'],
            "2019": [1000.0, 1999.0],
            "2020": [98.0, 199.0],
            "2021": [100.0, 200.0],
            "2022": [106.0, 202.0],
            "2023": [108.0, 203.0],
            "2024": [110.0, 204.0],
            "2025": [112.0, 205.0]
        }
    ).set_index(c.COUNTRY_CODE)
    gdp.ngdp_d_world = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['USA'],
            "2019": [1011.0],
            "2020": [101.0],
            "2021": [100.0],
            "2022": [102.0],
            "2023": [103.0],
            "2024": [104.0],
            "2025": [105.0]
        }
    ).set_index(c.COUNTRY_CODE)

    set_deflator(2021, 'Country')
    expected_df_1 = pd.DataFrame(
        [
            [1.020408, 1.0, 0.943396, 0.925926, 0.909091, 0.892857],
            [1.005025, 1.0, 0.990099, 0.985222, 0.980392, 0.975610]
        ],
        columns=[str(year) for year in range(2020, 2026)],
        index=["AAA", 'BBB']
    )
    expected_df_1.index.name = c.COUNTRY_CODE
    pd.testing.assert_frame_equal(
        gdp.deflator,
        expected_df_1
    )

    set_deflator(2021, 'World')
    case_2_values = [0.990099, 1.0, 0.980392, 0.970874, 0.961538, 0.952381]
    expected_df_2 = pd.DataFrame(
        [case_2_values, case_2_values],
        columns=[str(year) for year in range(2020, 2026)],
        index=["AAA", 'BBB']
    )
    expected_df_2.index.name = c.COUNTRY_CODE
    pd.testing.assert_frame_equal(
        gdp.deflator,
        expected_df_2
    )

def test_project_ngdp_rpch():
    gdp = GDP.__new__(GDP)
    project_ngdp_rpch = gdp._GDP__project_ngdp_rpch # pylint: disable=protected-access
    init_ngdp_rpch = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['AAA', 'BBB', 'CCC'],
            "2017": [999.0, 999.0, 999.0],
            "2018": [999.0, 999.0, 999.0],
            "2019": [999.0, 999.0, 999.0],
            "2020": [0.050, 0.11, 13.8],
            "2021": [0.051, 0.111, 13.9],
            "2022": [0.044, 0.0999, 14.0],
            "2023": [0.044, 0.099, 1.2],
            "2024": [0.045, 0.098, 1.1],
            "2025": [0.046, 0.096, 1.0]
        }
    ).set_index(c.COUNTRY_CODE)
    gdp.ngdp_rpch = init_ngdp_rpch.copy()

    # # simulation_year[1] > last db year
    project_ngdp_rpch((2022, 2030))
    expected_df_1 = pd.DataFrame(
        [
            [0.05, 0.051, 0.044, 0.044, 0.045, 0.046, 0.045, 0.045, 0.045, 0.045, 0.045],
            [0.11, 0.111, 0.0999, 0.099, 0.098, 0.096, 0.094, 0.092, 0.09, 0.088, 0.086],
            [13.8, 13.9, 14.0, 1.2, 1.1, 1.0, 2.716667, 4.433333, 6.15, 7.866667, 9.583333]
        ],
        columns=[str(year) for year in range(2020, 2031)],
        index=["AAA", 'BBB', 'CCC']
    )
    expected_df_1.index.name = c.COUNTRY_CODE
    pd.testing.assert_frame_equal(
        gdp.ngdp_rpch,
        expected_df_1
    )

    # simulation_year[1] == last db year
    gdp.ngdp_rpch = init_ngdp_rpch.copy()
    expected_df_2 = pd.DataFrame(
        [
            [0.05, 0.051, 0.044, 0.044, 0.045, 0.046],
            [0.11, 0.111, 0.0999, 0.099, 0.098, 0.096],
            [13.8, 13.9, 14.0, 1.2, 1.1, 1.0]
        ],
        columns=[str(year) for year in range(2020, 2025 + 1)],
        index=["AAA", 'BBB', 'CCC']
    )
    expected_df_2.index.name = c.COUNTRY_CODE
    project_ngdp_rpch((2022, 2025))
    pd.testing.assert_frame_equal(
        gdp.ngdp_rpch,
        expected_df_2
    )

    # simulation_year[1] < last db year
    gdp.ngdp_rpch = init_ngdp_rpch.copy()
    project_ngdp_rpch((2022, 2023))
    expected_df_3 = pd.DataFrame(
        [
            [0.05, 0.051, 0.044, 0.044],
            [0.11, 0.111, 0.0999, 0.099],
            [13.8, 13.9, 14.0, 1.2]
        ],
        columns=[str(year) for year in range(2020, 2023 + 1)],
        index=["AAA", 'BBB', 'CCC']
    )
    expected_df_3.index.name = c.COUNTRY_CODE
    pd.testing.assert_frame_equal(
        gdp.ngdp_rpch,
        expected_df_3
    )


def test_set_d_gdp_at_const_prices():
    gdp = GDP.__new__(GDP)
    set_d_gdp_at_const_prices = gdp._GDP__set_d_gdp_at_const_prices # pylint: disable=protected-access
    gdp.ngdp_rpch = pd.DataFrame(
        [
            [float(i) for i in range(1, 7)],
            [float(i) for i in range(11, 17)],
            [float(i) for i in range(21, 27)]
        ],
        columns=[str(year) for year in range(2020, 2025 + 1)],
        index=["AAA", 'BBB', 'CCC']
    )
    gdp.ngdp_rpch.index.name = c.COUNTRY_CODE


    set_d_gdp_at_const_prices('Base')
    pd.testing.assert_frame_equal(
        gdp.d_gdp_at_const_prices,
        gdp.ngdp_rpch
    )

    set_d_gdp_at_const_prices('High')
    expected_df_2 = pd.DataFrame(
        [
            [float(i * 1.5) for i in range(1, 7)],
            [float(i * 1.5) for i in range(11, 17)],
            [float(i * 1.5) for i in range(21, 27)]
        ],
        columns=[str(year) for year in range(2020, 2025 + 1)],
        index=["AAA", 'BBB', 'CCC']
    )
    expected_df_2.index.name = c.COUNTRY_CODE
    pd.testing.assert_frame_equal(
        gdp.d_gdp_at_const_prices,
        expected_df_2
    )

    set_d_gdp_at_const_prices('Low')
    expected_df_3 = pd.DataFrame(
        [
            [float(i * 0.5) for i in range(1, 7)],
            [float(i * 0.5) for i in range(11, 17)],
            [float(i * 0.5) for i in range(21, 27)]
        ],
        columns=[str(year) for year in range(2020, 2025 + 1)],
        index=["AAA", 'BBB', 'CCC']
    )
    expected_df_3.index.name = c.COUNTRY_CODE
    pd.testing.assert_frame_equal(
        gdp.d_gdp_at_const_prices,
        expected_df_3
    )

    with pytest.raises(
        KeyError,
        match="Loww'"
    ):
        set_d_gdp_at_const_prices('Loww')
