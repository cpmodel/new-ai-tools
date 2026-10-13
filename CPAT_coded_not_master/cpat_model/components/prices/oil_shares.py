import pandas as pd

import cpat_model.constants as c
from config import DATA_PATH

PRICES_DATA_FILE_NAME = 'oil_product_shares'


def load_oil_product_shares(
        selected_countries: list[str]
        ) -> pd.DataFrame:
    """
    [Data Loading]
    'Portion used for residential consumption'
    Loading data based on 'OilProductShares' tab (v407).

    Manual changes to 'oil_product_shares.csv':
    - removed columns: 'countryname', 'incomelevel', 'region'
        and columns with '_road_share' postfix
    - renamed columns: 'countrycode' -> 'CountryCode' 

    return dims (c), '{f}_share' data columns
    """
    oil_product_shares: pd.DataFrame = (
        # TODO: change to pkl
        pd.read_csv(f'{DATA_PATH}/{PRICES_DATA_FILE_NAME}.csv')
    )
    oil_product_shares = (
        oil_product_shares[oil_product_shares[c.COUNTRY_CODE].isin(selected_countries)]
    )
    oil_product_shares = oil_product_shares.melt(
        id_vars=c.COUNTRY_CODE,
        var_name=c.FUEL_CODE,
        value_name="Value"
    )
    oil_product_shares[c.FUEL_CODE] = (
        oil_product_shares[c.FUEL_CODE].str.replace("_share", "", regex=False)
    )
    oil_product_shares[c.SECTOR_CODE] = c.ALL

    oil_product_shares.set_index(c.ID_COL_NAMES, inplace=True)

    return oil_product_shares.sort_index()
