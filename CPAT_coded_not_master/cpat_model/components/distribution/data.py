"""
[Data Loading] for the Distribution module.

Mirrors the load_*/get_* functions in the other components (e.g.
energy_consumption/energy_data.py, power/data.py): each function reads one
CSV from DATA_PATH and returns it filtered for the selected countries.
InputData.__init__ wires these in the same way it wires every other
component's loaders.

Source: cpat_excel/Distribution/data_standardized/CPAT_DistributionalData.xlsx
(the six DATA_DISTN sheets, plus the sector/country mapping tables), exported
to CSV by cpat_excel/scripts/build_distribution_model_data.py. Run that
script (from the repo root) whenever data_standardized is regenerated.

Every loader here returns its sheet's native long/tidy shape (one row per
observation), not pivoted -- the different tables have different dimension
sets (fuel x decile x sample x statistic for HHSurvey; gtap_sector x metric
for IO_GTAP; quintile x program for ASPIRE; ...), so pivoting is left to the
calculation modules (price_changes.py, budget_shares.py, ...) that know what
shape they each need. The one normalization every loader does is a
`CountryCode` column in a consistent (uppercase ISO3) form, since the source
sheets are inconsistent about it (IO_GTAP uses lowercase 'egy', HHSurvey
uses 'EGY', ASPIRE has a differently-named column entirely).
"""
import pandas as pd

import cpat_model.constants as c
from config import DATA_PATH


IO_GTAP_FILE_NAME = 'distn_io_gtap'
HH_SURVEY_FILE_NAME = 'distn_hh_survey'
HH_ELAST_FILE_NAME = 'distn_hh_elast'
ASPIRE_FILE_NAME = 'distn_aspire'
WHO_COOKING_FILE_NAME = 'distn_who_cooking'
GDP_RATIOS_FILE_NAME = 'distn_gdp_ratios'
SECTOR_CROSSWALK_FILE_NAME = 'distn_mapping_sector_crosswalk'
COUNTRIES_TO_GTAP10_FILE_NAME = 'distn_mapping_countries_to_gtap10'


def load_io_gtap(selected_countries: list[str]) -> pd.DataFrame:
    """
    GTAP10 input-output table: sectoral energy intensities (Leontief
    coefficients: '<fuel>.leontiefs'), and household-demand/export/import/
    output/domestic-consumption aggregates ('hhd', 'exports', 'imports',
    'output', 'domcons', 'totd'), by GTAP sector. Feeds price_changes.py
    (Step 1: deriving indirect price changes for CPAT consumption categories).

    CPAT Excel: 'Distribution' sheet §C.I (rows 489-736); source sheet 'IO_GTAP'.
    Real 2014 US$bn (GTAP base year) -- a scale/currency/year different from
    the rest of the model; only used for *relative* (within-country) weights
    here, never mixed with LCU-denominated figures. See price_changes.py.

    columns: dis ('<iso3>.<metric>[.<fuel>]'), CountryCode, year, sector,
    gtap_sector, value
    """
    df = pd.read_csv(f'{DATA_PATH}/{IO_GTAP_FILE_NAME}.csv')
    df[c.COUNTRY_CODE] = df['iso3'].str.upper()
    return df[df[c.COUNTRY_CODE].isin(selected_countries)].reset_index(drop=True)


def load_hh_survey(selected_countries: list[str]) -> pd.DataFrame:
    """
    Household budget survey extract: budget shares (e.g. 'lpg_share',
    'food_share'), population ('popw'), per-capita consumption
    ('cons_pc_acrent'), infrastructure access indices (e.g. 'ely_acs_share'),
    and PIT incidence shares ('pit_share_income'), each broken out by
    decile/basket, sample (Overall/Urban/Rural) and statistic
    (mean/median/p25/p75). Feeds budget_shares.py (Step 3), rebasing.py
    (Step 7) and recycling.py (Step 8).

    CPAT Excel: 'Distribution' sheet §C.III, §C.IV, §C.VIII, §C.XI;
    source sheet 'HHSurvey'.

    columns: code, CountryCode, year, sample, type ('Basket'/'Deciles'),
    stat_type ('mean'/'median'/'p25'/'p75'), quant_cons (1-10, or 9999 for
    'Basket'), variable, value
    """
    df = pd.read_csv(f'{DATA_PATH}/{HH_SURVEY_FILE_NAME}.csv')
    df[c.COUNTRY_CODE] = df['iso3'].str.upper()
    return df[df[c.COUNTRY_CODE].isin(selected_countries)].reset_index(drop=True)


def load_hh_elast(selected_countries: list[str]) -> pd.DataFrame:
    """
    Own-price elasticities of demand by fuel/category, decile-specific
    (Overall sample, mean statistic only -- no urban/rural or
    median/p25/p75 breakdown is available in this data). Feeds
    budget_shares.py (Step 2: behavioural/DWL adjustments).

    CPAT Excel: 'Distribution' sheet §C.II; source sheet 'HH_Elast'.

    columns: same shape as load_hh_survey.
    """
    df = pd.read_csv(f'{DATA_PATH}/{HH_ELAST_FILE_NAME}.csv')
    df[c.COUNTRY_CODE] = df['iso3'].str.upper()
    return df[df[c.COUNTRY_CODE].isin(selected_countries)].reset_index(drop=True)


def load_aspire(selected_countries: list[str]) -> pd.DataFrame:
    """
    World Bank ASPIRE database: average per-capita transfer by *quintile*
    (not decile) and program type (e.g. Cash Transfers, Contributory
    Pensions). An alternate incidence source for recycling.py's targeted-
    transfer / current-spending channels, for use instead of the rules-based
    synthetic transfer when a country's scenario opts into an existing
    program's incidence pattern. Not used in the rules-based path (see
    recycling.py) -- kept for that future use.

    CPAT Excel: 'Distribution' sheet §C.X; source sheet 'ASPIRE'.

    columns: Series_Code, CountryCode, Year, Value, indicator_name,
    Sub_Topic5 (program group), Sub_Topic6 (program)
    """
    df = pd.read_csv(f'{DATA_PATH}/{ASPIRE_FILE_NAME}.csv')
    df[c.COUNTRY_CODE] = df['ISO-3 Country Code'].str.upper()
    df = df.rename(columns={'(firstnm) Value': 'Value', '(firstnm) indicator_name': 'indicator_name',
                             '(firstnm) Sub_Topic5': 'Sub_Topic5', '(firstnm) Sub_Topic6': 'Sub_Topic6'})
    return df[df[c.COUNTRY_CODE].isin(selected_countries)].reset_index(drop=True)


def load_who_cooking(selected_countries: list[str]) -> pd.DataFrame:
    """
    WHO household energy database: the most-used cooking fuel and its
    budget/usage share, by sample. Feeds effects.py's cooking-fuel
    exemption logic (Step 4).

    CPAT Excel: 'Distribution' sheet §C.III (cooking fuel rows); source
    sheet 'WHOCooking'.

    columns: code, CountryCode, year, sample, variable ('main_fuel' or
    'max_share'), value
    """
    df = pd.read_csv(f'{DATA_PATH}/{WHO_COOKING_FILE_NAME}.csv')
    df[c.COUNTRY_CODE] = df['iso3'].str.upper()
    return df[df[c.COUNTRY_CODE].isin(selected_countries)].reset_index(drop=True)


def load_gdp_ratios(selected_countries: list[str]) -> pd.DataFrame:
    """
    World Bank indicator time series: household consumption-to-GDP ratio,
    plus government education/health expenditure-to-GDP ratios (used as a
    rough current-spending incidence proxy). Feeds rebasing.py (Step 7).

    CPAT Excel: 'Distribution' sheet §B.I ('Household Consumption'),
    §C.VIII; source sheet 'GDPRatios'.

    columns: Country Name, CountryCode, Indicator Name, Indicator Code,
    Latest Year, year, value
    """
    df = pd.read_csv(f'{DATA_PATH}/{GDP_RATIOS_FILE_NAME}.csv')
    df[c.COUNTRY_CODE] = df['Country Code'].str.upper()
    return df[df[c.COUNTRY_CODE].isin(selected_countries)].reset_index(drop=True)


def load_gtap_cpat_sector_crosswalk() -> pd.DataFrame:
    """
    GTAP10 -> CPAT consumption-category crosswalk (plus WIOT/EORA/IEA/ISIC
    codes for reference). Not country-specific. Feeds price_changes.py
    (Step 1.6: aggregating GTAP sectors up to the ~14 CPAT indirect
    consumption categories).

    CPAT Excel: source sheet Mapping_SectorCrosswalk (see also
    Mapping_CPATSectorsToISIC, Mapping_IEAFlowsToISIC, Mapping_ISICToCPAT,
    not loaded here as nothing in this module uses them yet).

    Only the columns price_changes.py actually reads are kept:
    'GTAP 10 codes' (the IO_GTAP gtap_sector code) and
    'CPAT, consumption items (3-letter)' (the target category, 'NA' if the
    sector has no household-consumption-category mapping, or a
    comma-separated list for the refined-petroleum sector, which spans
    several direct fuels rather than one indirect category).

    return dims (GtapSectorCode), columns: CPATCategory
    """
    df = pd.read_csv(f'{DATA_PATH}/{SECTOR_CROSSWALK_FILE_NAME}.csv')
    df = df[['GTAP 10 codes', 'CPAT, consumption items (3-letter)']].rename(
        columns={'GTAP 10 codes': c.GTAP_SECTOR_CODE, 'CPAT, consumption items (3-letter)': 'CPATCategory'}
    )
    df = df.dropna(subset=[c.GTAP_SECTOR_CODE]).drop_duplicates(subset=[c.GTAP_SECTOR_CODE])
    return df.set_index(c.GTAP_SECTOR_CODE)


def load_countries_to_gtap10(selected_countries: list[str]) -> pd.DataFrame:
    """
    Country -> GTAP10 region crosswalk (most countries map to a multi-
    country GTAP10 aggregate region; larger economies, Egypt included, are
    their own GTAP10 region). Not used directly yet (this Egypt-only
    dataset only ever needs Egypt's own region), kept for when the pipeline
    covers more countries.

    CPAT Excel: source sheet Mapping_CountriesToGTAP10.

    return dims (CountryCode), columns: GtapCode
    """
    df = pd.read_csv(f'{DATA_PATH}/{COUNTRIES_TO_GTAP10_FILE_NAME}.csv')
    df[c.COUNTRY_CODE] = df['iso3'].str.upper()
    df = df[df[c.COUNTRY_CODE].isin(selected_countries)]
    return df.set_index(c.COUNTRY_CODE)[['gtapcode']].rename(columns={'gtapcode': 'GtapCode'})
