import pandas as pd
import numpy as np

from cpat_model.inputs.dashboard_inputs import DashboardInputsDict
from cpat_model.components.policies.policies import (
    Policies,
    get_fuel_carbon_price_inclusion, get_sectors_carbon_price_inclusion
)

from cpat_model.inputs.input_data import InputData

import cpat_model.constants as c

# TODO: sort out this:
APPLY_TAX_F_CODES = c.FOSSIL_FUELS + [c.BIO, c.REN, c.JFU]
P_COV_F_CODES = sorted(APPLY_TAX_F_CODES)

APPLY_TAX_S_CODES = [
    c.POW, c.TRA, c.ROD, c.RAL, c.AVI, c.NAV, c.RES, c.FOO, c.SRV,
    c.MCH, c.IRN, c.NFM, c.MAC, c.CEM, c.OMN, c.CST, c.FTR, c.OEN, c.IND
]
P_COV_S_CODES = [
    c.AVI, c.CEM, c.CST, c.FOO, c.FTR, c.IND, c.IRN, c.MAC, c.MCH,
    c.NAV, c.NFM, c.OEN, c.OMN, c.POW, c.RAL, c.RES, c.ROD, c.SRV, 'trs'
]

def test_get_p_based_policies_cov_s_f():
    obj = Policies.__new__(Policies)
    get_p_based_policies_cov_s_f = obj._Policies__get_p_based_policies_cov_s_f # pylint: disable=protected-access
    obj.p_cov_f = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 10 + ['UKR'] * 10,
            c.FUEL_CODE: P_COV_F_CODES * 2,
            "2024": list(np.arange(21, 40 + 1, dtype=float)),
            "2025": list(np.arange(1, 20 + 1, dtype=float)),
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE])
    obj.p_cov_s = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 19 + ['UKR'] * 19,
            c.SECTOR_CODE: P_COV_S_CODES * 2,
            "2024": list(np.arange(31, 68 + 1, dtype=float)),
            "2025": list(np.arange(31, 68 + 1, dtype=float)),
        }
    ).set_index([c.COUNTRY_CODE, c.SECTOR_CODE])

    expected_df_1 = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 12 + ['UKR'] * 12,
            c.SECTOR_CODE: ([c.ALL] * 5 + [c.IND] * 3 + [c.POW] * 2 + [c.RES] * 2) * 2,
            c.FUEL_CODE: [
                c.BIO, c.DIE, c.GSO, c.KER, c.LPG,
                c.COA, c.NGA, c.OOP,
                c.COA, c.NGA,
                c.COA, c.NGA
            ] * 2,
            "2025": [
                1.0, 3.0, 4.0, 6.0, 7.0, 96.0, 384.0, 432.0, 88.0, 352.0, 92.0, 368.0,
                11.0, 13.0, 14.0, 16.0, 17.0, 804.0, 1206.0, 1273.0, 756.0, 1134.0, 780.0, 1170.0
            ]
        }
    ).set_index(c.ID_COL_NAMES)
    pd.testing.assert_frame_equal(
        get_p_based_policies_cov_s_f(2025),
        expected_df_1
    )


def test_get_p_cov_f():
    obj = Policies.__new__(Policies)
    get_p_cov_f = obj._Policies__get_p_cov_f # pylint: disable=protected-access

    obj.apply_tax_f = pd.DataFrame(
        {
            c.FUEL_CODE: APPLY_TAX_F_CODES,
            c.CARBON_TAX: [True] * 7 + [False] * 3,
        }
    ).set_index(c.FUEL_CODE)
    selected_countries = ['DEU', 'UKR']
    obj._Policies__get_apply_tax_f_extended(selected_countries) # pylint: disable=protected-access

    expected_df_1 = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 10 + ['UKR'] * 10,
            c.FUEL_CODE: P_COV_F_CODES * 2,
            "2020": [0] * 20,
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE])
    pd.testing.assert_frame_equal(
        # year < cp_intro
        get_p_cov_f(c.CARBON_TAX, 2025, 2020, True, 2025, 5),
        expected_df_1
    )

    # same countries
    expected_df_2 = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 10 + ['UKR'] * 10,
            c.FUEL_CODE: P_COV_F_CODES * 2,
            "2025": [
                0, 1, 1, 1, 0, 1, 1, 1, 1, 0,
                0, 1, 1, 1, 0, 1, 1, 1, 1, 0
            ],
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE])
    pd.testing.assert_frame_equal(
        # exempt_phaseout = False
        get_p_cov_f(c.CARBON_TAX, 2025, 2025, False, 2025, 5),
        expected_df_2
    )

    selected_countries = ['DEU']
    obj._Policies__get_apply_tax_f_extended(selected_countries) # pylint: disable=protected-access
    expected_df_3 = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['DEU'] * 10,
            c.FUEL_CODE: P_COV_F_CODES,
            "2025": [0.2, 1.0, 1.0, 1.0, 0.2, 1.0, 1.0, 1.0, 1.0, 0.2],
        }
    ).set_index([c.COUNTRY_CODE, c.FUEL_CODE])
    pd.testing.assert_frame_equal(
        # exempt_phaseout = True and 1 country
        get_p_cov_f(c.CARBON_TAX, 2025, 2025, True, 2025, 5),
        expected_df_3
    )


def test_get_existing_ets_extended():
    obj = Policies.__new__(Policies)
    get_existing_ets_extended = obj._Policies__get_existing_ets_extended # pylint: disable=protected-access

    existing_ets_2025 = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['ALB'] * 4 + ['MWI'] * 4,
            c.SECTOR_CODE: [c.IND, c.POW, c.RES, 'trs'] * 2,
            '2025': [float(i) for i in range(1, 8 + 1)]
        }
    ).set_index([c.COUNTRY_CODE, c.SECTOR_CODE])
    expected_df = pd.DataFrame(
        {
            c.COUNTRY_CODE: ['ALB'] * 19 + ['MWI'] * 19,
            c.SECTOR_CODE: P_COV_S_CODES * 2,
            '2025': [
                4.0, 1.0, 1.0, 3.0, *([1.0] * 5), 4.0, 1.0, 0.0, 1.0, 2.0, 4.0, 3.0, 4.0, 3.0, 4.0,
                8.0, 5.0, 5.0, 7.0, *([5.0] * 5), 8.0, 5.0, 0.0, 5.0, 6.0, 8.0, 7.0, 8.0, 7.0, 8.0
            ]
        }
    ).set_index([c.COUNTRY_CODE, c.SECTOR_CODE])
    pd.testing.assert_frame_equal(
        get_existing_ets_extended(existing_ets_2025),
        expected_df
    )

def test_set_apply_tax_f():
    policies = Policies.__new__(Policies)
    set_apply_tax_f = policies._Policies__set_apply_tax_f # pylint: disable=protected-access
    policies.fuel_carbon_price_inclusion = pd.DataFrame(
        {
            c.FUEL_CODE: APPLY_TAX_F_CODES,
            c.CARBON_TAX: [True] * 7 + [False] * 3,
        }
    ).set_index(c.FUEL_CODE)

    set_apply_tax_f(c.CARBON_TAX)
    expected_df_1 = pd.DataFrame(
        {
            c.FUEL_CODE: APPLY_TAX_F_CODES,
            c.CARBON_TAX: [True] * 7 + [False] * 3, # explicit pricing
        }
    ).set_index(c.FUEL_CODE)
    pd.testing.assert_frame_equal(
        policies.apply_tax_f,
        expected_df_1
    )

    policies.fuel_carbon_price_inclusion = pd.DataFrame(
        {
            c.FUEL_CODE: APPLY_TAX_F_CODES,
            c.ENERGY_EFFICIENCY_REGULATIONS: [True] * 7 + [False] * 3,
        }
    ).set_index(c.FUEL_CODE)

    set_apply_tax_f(c.ENERGY_EFFICIENCY_REGULATIONS)
    expected_df_2 = pd.DataFrame(
        {
            c.FUEL_CODE: APPLY_TAX_F_CODES,
            c.ENERGY_EFFICIENCY_REGULATIONS: [False] * 10, # shadow pricing
        }
    ).set_index(c.FUEL_CODE)
    pd.testing.assert_frame_equal(
        policies.apply_tax_f,
        expected_df_2
    )

    policies.fuel_carbon_price_inclusion = pd.DataFrame(
        {
            c.FUEL_CODE: APPLY_TAX_F_CODES,
            c.FEEBATES: [True] * 7 + [False] * 3,
        }
    ).set_index(c.FUEL_CODE)

    set_apply_tax_f(c.FEEBATES)
    expected_df_3 = pd.DataFrame(
        {
            c.FUEL_CODE: APPLY_TAX_F_CODES,
            c.FEEBATES: [True] * 10, # feebates scenario
        }
    ).set_index(c.FUEL_CODE)
    pd.testing.assert_frame_equal(
        policies.apply_tax_f,
        expected_df_3
    )


def test_set_apply_tax_s():
    policies = Policies.__new__(Policies)
    set_apply_tax_s = policies._Policies__set_apply_tax_s # pylint: disable=protected-access
    policies.sectors_carbon_price_inclusion = pd.DataFrame(
        {
            c.SECTOR_CODE: APPLY_TAX_S_CODES,
            c.INDUSTRIAL_EFFICIENCY_REGULATIONS: [False] * 7 + [True] * 11 + [False],
        }
    ).set_index(c.SECTOR_CODE)

    set_apply_tax_s(c.INDUSTRIAL_EFFICIENCY_REGULATIONS)
    expected_df_1 = pd.DataFrame(
        {
            c.SECTOR_CODE: APPLY_TAX_S_CODES,
            c.INDUSTRIAL_EFFICIENCY_REGULATIONS: [False] * 19, # shadow pricing
        }
    ).set_index(c.SECTOR_CODE)
    pd.testing.assert_frame_equal(
        policies.apply_tax_s,
        expected_df_1
    )

    policies.sectors_carbon_price_inclusion = pd.DataFrame(
        {
            c.SECTOR_CODE: APPLY_TAX_S_CODES,
            c.ROAD_FUEL_TAX: [False, False, True] + [False] * 16,
        }
    ).set_index(c.SECTOR_CODE)

    set_apply_tax_s(c.ROAD_FUEL_TAX)
    expected_df_2 = pd.DataFrame(
        {
            c.SECTOR_CODE: APPLY_TAX_S_CODES,
            c.ROAD_FUEL_TAX: [False, False, True] + [False] * 16, # explicit pricing
        }
    ).set_index(c.SECTOR_CODE)
    pd.testing.assert_frame_equal(
        policies.apply_tax_s,
        expected_df_2
    )

    policies.sectors_carbon_price_inclusion = pd.DataFrame(
        {
            c.SECTOR_CODE: APPLY_TAX_S_CODES,
            c.POWER_FEEBATE: [True] + [False] * 18,
        }
    ).set_index(c.SECTOR_CODE)

    set_apply_tax_s(c.POWER_FEEBATE)
    expected_df_3 = pd.DataFrame(
        {
            c.SECTOR_CODE: APPLY_TAX_S_CODES,
            c.POWER_FEEBATE: [True] + [False] * 18, # shadow pricing and True for pow
        }
    ).set_index(c.SECTOR_CODE)
    pd.testing.assert_frame_equal(
        policies.apply_tax_s,
        expected_df_3
    )


def test_set_carbon_tax_inclusion():
    policies = Policies.__new__(Policies)
    set_carbon_tax_inclusion = policies._Policies__set_carbon_tax_inclusion # pylint: disable=protected-access

    input_data = InputData.__new__(InputData)
    input_data.fuel_carbon_price_inclusion = get_fuel_carbon_price_inclusion()
    input_data.sectors_carbon_price_inclusion = get_sectors_carbon_price_inclusion()

    set_carbon_tax_inclusion(input_data, c.INDUSTRIAL_EFFICIENCY_REGULATIONS)
    expected_df_s = pd.DataFrame(
        {
            c.SECTOR_CODE: APPLY_TAX_S_CODES,
            c.INDUSTRIAL_EFFICIENCY_REGULATIONS: [
                False, False, False, False, False, False, False, True, True,
                True, True, True, True, True, True, True, True, True, False
            ],
        }
    ).set_index(c.SECTOR_CODE)
    pd.testing.assert_frame_equal(
        policies.sectors_carbon_price_inclusion,
        expected_df_s
    )

    set_carbon_tax_inclusion(input_data, c.CARBON_TAX)
    expected_df_f = pd.DataFrame(
        {
            c.FUEL_CODE: APPLY_TAX_F_CODES,
            c.CARBON_TAX: [True, True, True, True, True, True, True, False, False, False],
        }
    ).set_index(c.FUEL_CODE)
    pd.testing.assert_frame_equal(
        policies.fuel_carbon_price_inclusion,
        expected_df_f
    )

def test_set_cp_trajectory():
    policies = Policies.__new__(Policies)
    set_cp_trajectory = policies._Policies__set_cp_trajectory # pylint: disable=protected-access

    simulation_years = (2022, 2035)
    d: DashboardInputsDict = {
        'ets_adj': 1.1,
        'tax_pathway': 'nom',

        'start_cp_year': 2025,
        'target_cp_year': 2030,
        'start_cp': 100,
        'target_cp': 250,

        'cp_trajectory_type': 'exp',
        'ct_exp_rate': 1.1,
        'per_tax_inc': 0.05
    }
    deflator = pd.DataFrame(
        [
            [1.02, 1.0, 0.94, 0.92, 0.90, 0.89, *[0.85] * 9],
            [1.01, 1.0, 0.99, 0.98, 0.98, 0.97, *[0.95] * 9]
        ],
        columns=[str(year) for year in range(simulation_years[0] - 1, simulation_years[1] + 1)],
        index=pd.Index(["AAA", "BBB"], name=c.COUNTRY_CODE)
    )
    columns = np.arange(simulation_years[0], simulation_years[1] + 1).astype(str)

    # Baseline -> all 0s
    pd.testing.assert_frame_equal(
        set_cp_trajectory(simulation_years, c.BASELINE, d, deflator),
        pd.DataFrame(
            0.0,
            columns=columns,
            index=pd.Index(["AAA", "BBB"], name=c.COUNTRY_CODE)
        )
    )

    # exponential
    pd.testing.assert_frame_equal(
        set_cp_trajectory(simulation_years, c.CARBON_TAX, d, deflator),
        pd.DataFrame(
            [[*[0.0] * 3, 100.0, 110.0, 121.0, 133.1, 146.41, *[161.051] * 6]] * 2,
            columns=columns,
            index=pd.Index(["AAA", "BBB"], name=c.COUNTRY_CODE)
        )
    )

    # constant after target_cp_year
    d['cp_trajectory_type'] = 'cst'
    pd.testing.assert_frame_equal(
        set_cp_trajectory(simulation_years, c.CARBON_TAX, d, deflator),
        pd.DataFrame(
            [[*[0.0] * 3, 100.0, 130.0, 160.0, 190.0, 220.0, *[250.0] * 6]] * 2,
            columns=columns,
            index=pd.Index(["AAA", "BBB"], name=c.COUNTRY_CODE)
        )
    )

    # linear
    d['cp_trajectory_type'] = 'lin'
    pd.testing.assert_frame_equal(
        set_cp_trajectory(simulation_years, c.CARBON_TAX, d, deflator),
        pd.DataFrame(
            [[
                *[0.0] * 3, 100.0, 130.0, 160.0, 190.0, 220.0,
                250.0, 280.0, 310.0, 340.0, 370.0, 400.0
            ]] * 2,
            columns=columns,
            index=pd.Index(["AAA", "BBB"], name=c.COUNTRY_CODE)
        )
    )

    # percentage increase
    d['cp_trajectory_type'] = 'per'
    d['ct_exp_rate'] = 12.5/1.05 # to have 5% increase with target_cp=250
    pd.testing.assert_frame_equal(
        set_cp_trajectory(simulation_years, c.CARBON_TAX, d, deflator),
        pd.DataFrame(
            [[
                *[0.0] * 3, 100.0, 130.0, 160.0, 190.0, 220.0,
                250.0, 262.5, 275.625, 289.40625, 303.876563, 319.070391
            ]] * 2,
            columns=columns,
            index=pd.Index(["AAA", "BBB"], name=c.COUNTRY_CODE)
        )
    )

    # ETS
    d['cp_trajectory_type'] = 'lin'
    pd.testing.assert_frame_equal(
        set_cp_trajectory(simulation_years, c.ETS, d, deflator),
        pd.DataFrame(
            [[
                *[0.0] * 3, 110.0, 143.0, 176.0, 209.0, 242.0,
                275.0, 308.0, 341.0, 374.0, 407.0, 440.0
            ]] * 2,
            columns=columns,
            index=pd.Index(["AAA", "BBB"], name=c.COUNTRY_CODE)
        )
    )

    # Real
    d['cp_trajectory_type'] = 'lin'
    d['tax_pathway'] = 'rea'
    pd.testing.assert_frame_equal(
        set_cp_trajectory(simulation_years, c.CARBON_TAX, d, deflator),
        pd.DataFrame(
            [
                [
                    *[0.0] * 3, 90.0, 115.7, 136.0, 161.5, 187.0,
                    212.5, 238.0, 263.5, 289.0, 314.5, 340.0
                ],
                [
                    *[0.0] * 3, 98.0, 126.1, 152.0, 180.5, 209.0,
                    237.5, 266.0, 294.5, 323.0, 351.5, 380.0
                ]
            ],
            columns=columns,
            index=pd.Index(["AAA", "BBB"], name=c.COUNTRY_CODE)
        )
    )
