import sys
import os

import itertools
import pytest
import pandas as pd
import numpy as np

# Add the root directory to the sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# pylint: disable=wrong-import-position
import cpat_model.constants as c

EXPECTED_DATA_PATH = 'cpat_testing/model/data_expected'

COUNTRIES = ['USA', 'KAZ']
INDEX_COMBINATIONS = list(itertools.product(COUNTRIES, c.SECTORS_UNIQUE, c.FUELS_UNIQUE))
YEAR = 2020
YEAR_VALUES = np.arange(0.0, 0.0 + len(INDEX_COMBINATIONS) * 1, 1.0) # 0, 1, 2, ...


@pytest.fixture
def sample_df_single_year():
    """
    Fixture that creates a sample pandas DataFrame with all unique combinations
    defined in INDEX_COMBINATIONS. Contains only a single year column with fixed range values.
    """
    df = pd.DataFrame(INDEX_COMBINATIONS, columns=c.ID_COL_NAMES)
    df[str(YEAR)] = YEAR_VALUES
    df.set_index(c.ID_COL_NAMES, inplace=True)
    return df

@pytest.fixture
def sample_df_single_year_with_zeros():
    """
    Fixture that creates a sample pandas DataFrame with all unique combinations
    defined in INDEX_COMBINATIONS. Contains only a single year column with fixed range values.
    Every 7th value is changed to 0.
    """
    df = pd.DataFrame(INDEX_COMBINATIONS, columns=c.ID_COL_NAMES)
    year_values = YEAR_VALUES.copy()
    year_values[::7] = 0
    df[str(YEAR)] = year_values
    df.set_index(c.ID_COL_NAMES, inplace=True)
    return df

@pytest.fixture
def sample_df_sg_fossil_fuels():
    """
    Fixture that creates a sample pandas DataFrame with all unique combinations
    for COUNTRIES, SECTOR_GROUPS and FOSSIL_FUELS_SORTED.
    Contains only a single year column with fixed values. Every 7th value is changed to 0.
    """
    index_combinations = list(
        itertools.product(COUNTRIES, c.SECTOR_GROUPS, c.FOSSIL_FUELS_SORTED)
    )
    df = pd.DataFrame(index_combinations, columns=c.ID_COL_NAMES)
    year_values = [9.0] * len(index_combinations)
    year_values[::7] = [0.0] * (len(year_values) // 7)
    df[str(YEAR)] = year_values
    df.set_index(c.ID_COL_NAMES, inplace=True)
    return df

@pytest.fixture
def sample_df_sg_fossil_fuels_and_bio():
    """
    Fixture that creates a sample pandas DataFrame with all unique combinations
    for COUNTRIES, SECTOR_GROUPS and FOSSIL_FUELS_AND_BIO.
    Contains only a single year column with fixed values. Every 6th value is changed to 0.
    """
    index_combinations = list(
        itertools.product(COUNTRIES, c.SECTOR_GROUPS, c.FOSSIL_FUELS_AND_BIO)
    )
    df = pd.DataFrame(index_combinations, columns=c.ID_COL_NAMES)
    year_values = [7.0] * len(index_combinations)
    year_values[::8] = [0.0] * (len(year_values) // 8)
    df[str(YEAR)] = year_values
    df.set_index(c.ID_COL_NAMES, inplace=True)
    return df

@pytest.fixture
def csv_df_index(request):
    """
    Fixture to read a CSV file for a given file name.
    Reads index columns from params and sets it on the returned df.
    """
    file_name, index_columns = request.param  # Get the parameter passed to the fixture
    return pd.read_csv(f'cpat_data/new_data/{file_name}').set_index(index_columns)

@pytest.fixture
def csv_dfs_index(request):
    """
    Fixture to read one or more CSV files.
    Reads index columns from params and sets them on the returned DataFrames.
    Returns a single DataFrame if one file is provided.
    Returns a tuple of DataFrames if multiple files are provided.
    """
    file_params = request.param  # Expecting a list of (file_name, index_columns) tuples

    # Load all DataFrames
    dfs = tuple(pd.read_csv(f"cpat_data/new_data/{file_name}").set_index(index_columns) 
                for file_name, index_columns in file_params)

    # If only one file is passed, return just the DataFrame (not as a tuple)
    return dfs[0] if len(dfs) == 1 else dfs

@pytest.fixture
def csv_df(request):
    """Fixture to read a CSV file for a given file name."""
    file_name = request.param  # Get the parameter passed to the fixture
    return pd.read_csv(f'cpat_data/new_data/{file_name}')
