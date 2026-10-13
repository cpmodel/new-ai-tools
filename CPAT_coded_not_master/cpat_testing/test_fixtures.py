import pandas as pd

from conftest import YEAR

def test_sample_df_single_year(sample_df_single_year):
    assert isinstance(sample_df_single_year, pd.DataFrame)
    assert sample_df_single_year.shape[0] == 600  # 2×20×15 = 600 rows
    assert list(sample_df_single_year.columns) == [str(YEAR)]
