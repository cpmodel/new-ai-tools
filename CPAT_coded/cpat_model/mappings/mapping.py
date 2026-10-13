import pandas as pd

import cpat_model.constants as c

MAPPINGS_PATH = 'cpat_model/mappings/mappings.xlsx'

df = pd.read_excel(MAPPINGS_PATH, sheet_name='Countries')
ALL_COUNTRIES = df.set_index(c.COUNTRY_CODE)['Country'].to_dict()


regions_df = pd.read_excel(MAPPINGS_PATH, sheet_name='Regions')
country_to_region_df = pd.read_excel(MAPPINGS_PATH, sheet_name='Country to Region')
COUNTRY_TO_REGION = (
    country_to_region_df
    .merge(regions_df, on='Region')
    .set_index(c.COUNTRY_CODE)['RegionCode']
    .to_dict()
)

# Country to IFS codes
COUNTRY_TO_IFS_CODE = (
    pd.read_excel(MAPPINGS_PATH, sheet_name='Country to IFS codes')
)

# Used only for elasticities, based on Elasticities tab, does not match Countries tab mapping
# HIC High income, LIC Low income, LMIC Lower middle income, UMIC Upper middle income
COUNTRY_TO_INCOME_GROUP = (
    pd.read_excel(MAPPINGS_PATH, sheet_name='Income Groups')
    .set_index(c.COUNTRY_CODE)['Income group']
    .to_dict()
)
