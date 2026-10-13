import pandas as pd
import numpy as np

from cpat_model.components.prices.ntx import (
    NTX
)

def test_get_is_additional_ntx():
    # pylint: disable=protected-access
    assert NTX._get_is_additional_ntx('no tax', True) is True
    assert NTX._get_is_additional_ntx('pigouvian', False) is True
    assert NTX._get_is_additional_ntx('pigouvian', True) is True
    assert NTX._get_is_additional_ntx('FFS', False) is False
    # pylint: enable=protected-access


def test_get_pigouvian_phase_in_coeff():
    simulation_years = (2020, 2030)
    columns = np.arange(simulation_years[0], simulation_years[1] + 1).astype(str)

    pd.testing.assert_frame_equal(
        NTX._get_pigouvian_phase_in_coeff( # pylint: disable=protected-access
            5, 2023, simulation_years
        ),
        pd.DataFrame(
            [[*[0.0] * 3, 0.2, 0.4, 0.6, 0.8, *[1.0] * 4 ]],
            columns=columns
        )
    )
