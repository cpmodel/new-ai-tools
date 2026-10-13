import pandas as pd
import numpy as np

from cpat_model.components.policies.phaseouts import Phaseouts

import cpat_model.constants as c


def test_get_phaseout_producer():
    simulation_years = (2021, 2032)
    obj = Phaseouts.__new__(Phaseouts)  # creates instance without calling __init__
    # pylint: disable=protected-access
    obj._Phaseouts__years = np.arange(simulation_years[0] - 1, simulation_years[1] + 1)
    obj._Phaseouts__columns = obj._Phaseouts__years.astype(str)
    get_phaseout_producer = obj._Phaseouts__get_phaseout_producer
    # pylint: enable=protected-access
    columns = np.arange(simulation_years[0] - 1, simulation_years[1] + 1).astype(str)
    energy_sector_reform = {
        'is_ffs_phaseout_prod': False,
        'ffs_phaseout_prod': 5,
        'ffs_start_prod': 2025,
        'ffs_phaseout_share_prod': 0.5  
    }

    pd.testing.assert_frame_equal(
        get_phaseout_producer(energy_sector_reform),
        pd.DataFrame([np.ones(len(columns), dtype=float)], columns=columns)
    )

    energy_sector_reform['is_ffs_phaseout_prod'] = True
    expected_df_1 = pd.DataFrame(
        [[*[1.0] * 6, 0.9, 0.8, 0.7, 0.6, *[0.5] * 3]],
        columns=columns
    )
    pd.testing.assert_frame_equal(
        get_phaseout_producer(energy_sector_reform),
        expected_df_1
    )


def test_get_phaseout_consumer():
    simulation_years = (2021, 2032)
    obj = Phaseouts.__new__(Phaseouts)  # creates instance without calling __init__
    # pylint: disable=protected-access
    obj._Phaseouts__years = np.arange(simulation_years[0] - 1, simulation_years[1] + 1)
    obj._Phaseouts__columns = obj._Phaseouts__years.astype(str)
    get_phaseout_consumer = obj._Phaseouts__get_phaseout_consumer
    # pylint: enable=protected-access
    columns = np.arange(simulation_years[0] - 1, simulation_years[1] + 1).astype(str)
    energy_sector_reform = {
        'is_ffs_phaseout_cons': False,
        'ffs_phaseout_cons': 5,
        'ffs_start_cons': 2025,
        'ffs_phaseout_share_cons': 0.5  
    }

    pd.testing.assert_frame_equal(
        get_phaseout_consumer(energy_sector_reform),
        pd.DataFrame([np.ones(len(columns), dtype=float)], columns=columns)
    )

    energy_sector_reform['is_ffs_phaseout_cons'] = True
    expected_df_1 = pd.DataFrame(
        [[*[1.0] * 6, 0.9, 0.8, 0.7, 0.6, *[0.5] * 3]],
        columns=columns
    )
    pd.testing.assert_frame_equal(
        get_phaseout_consumer(energy_sector_reform),
        expected_df_1
    )


def test_get_phaseout_subsidy_tax():
    simulation_years = (2021, 2032)
    obj = Phaseouts.__new__(Phaseouts)  # creates instance without calling __init__
    # pylint: disable=protected-access
    obj._Phaseouts__years = np.arange(simulation_years[0] - 1, simulation_years[1] + 1)
    obj._Phaseouts__columns = obj._Phaseouts__years.astype(str)
    get_phaseout_subsidy_tax = obj._Phaseouts__get_phaseout_subsidy_tax
    # pylint: enable=protected-access
    columns = np.arange(simulation_years[0] - 1, simulation_years[1] + 1).astype(str)
    energy_sector_reform = {
        'is_pc_phaseout': True,
        'pc_phaseout': 5,
        'pc_start': 2025,
        'ffs_phaseout_share_cons': 0.5  
    }
    scenario_type = c.BASELINE
    is_pc_phaseout_baseline = False

    # 'is_pc_phaseout': True, Baseline scenario, is_pc_phaseout_baseline = False
    pd.testing.assert_frame_equal(
        get_phaseout_subsidy_tax(
            scenario_type, is_pc_phaseout_baseline, energy_sector_reform
        ),
        pd.DataFrame([np.ones(len(columns), dtype=float)], columns=columns)
    )

    is_pc_phaseout_baseline = True
    energy_sector_reform['is_pc_phaseout'] = False
    # 'is_pc_phaseout': False, Baseline scenario, is_pc_phaseout_baseline = True
    pd.testing.assert_frame_equal(
        get_phaseout_subsidy_tax(
            scenario_type, is_pc_phaseout_baseline, energy_sector_reform
        ),
        pd.DataFrame([np.ones(len(columns), dtype=float)], columns=columns)
    )

    is_pc_phaseout_baseline = True
    energy_sector_reform['is_pc_phaseout'] = False
    # 'is_pc_phaseout': False, Baseline scenario, is_pc_phaseout_baseline = True
    pd.testing.assert_frame_equal(
        get_phaseout_subsidy_tax(
            scenario_type, is_pc_phaseout_baseline, energy_sector_reform
        ),
        pd.DataFrame([np.ones(len(columns), dtype=float)], columns=columns)
    )

    is_pc_phaseout_baseline = False
    scenario_type = c.CARBON_TAX
    energy_sector_reform['is_pc_phaseout'] = False
    # 'is_pc_phaseout': False, not Baseline scenario, is_pc_phaseout_baseline = True
    pd.testing.assert_frame_equal(
        get_phaseout_subsidy_tax(
            scenario_type, is_pc_phaseout_baseline, energy_sector_reform
        ),
        pd.DataFrame([np.ones(len(columns), dtype=float)], columns=columns)
    )

    expected_df_1 = pd.DataFrame(
        [[*[1.0] * 6, 0.9, 0.8, 0.7, 0.6, *[0.5] * 3]],
        columns=columns
    )

    energy_sector_reform['is_pc_phaseout'] = True
    # 'is_pc_phaseout': True, not Baseline scenario
    pd.testing.assert_frame_equal(
        get_phaseout_subsidy_tax(
            scenario_type, is_pc_phaseout_baseline, energy_sector_reform
        ),
        expected_df_1
    )

    scenario_type = c.BASELINE
    is_pc_phaseout_baseline = True
    energy_sector_reform['is_pc_phaseout'] = True
    # 'is_pc_phaseout': True, Baseline scenario, is_pc_phaseout_baseline = True
    pd.testing.assert_frame_equal(
        get_phaseout_subsidy_tax(
            scenario_type, is_pc_phaseout_baseline, energy_sector_reform
        ),
        expected_df_1
    )
