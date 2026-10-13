import pandas as pd

import cpat_model.constants as c
from config import DATA_PATH


ENERGY_DATA_FILE_NAME = 'energy_consumption'


TOTAL_ENERGY_USE_COL = 'total_energy_use'

def load_energy_consumptions(
        selected_countries: list[str]
        ) -> pd.DataFrame:
    """
    [Data Loading]
    'Energy Consumption'
    CPAT Excel table strts row 559 (v407)

    Input file preprocessed from Energy Balances input table
    in get_energy_balances_csv function.

    return dims (c, s), (f)
    """
    # TODO: pkl
    df: pd.DataFrame = pd.read_csv(f'{DATA_PATH}/{ENERGY_DATA_FILE_NAME}.csv')
    df = df[df[c.COUNTRY_CODE].isin(selected_countries)].set_index([c.COUNTRY_CODE, c.SECTOR_CODE])

    return df
