from typing import TYPE_CHECKING
import pandas as pd

import cpat_model.constants as c
from config import DATA_PATH

if TYPE_CHECKING:
    from cpat_model.inputs.input_data import InputData


class ExistingCT:
    """
    Existing carbon taxes from
    'Existing carbon pricing mechanisms (ETSs and carbon taxes)' section (v407)
    CPAT Excel: section start 1725
    """
    # Public attributes:
    existing_ct_rate: pd.DataFrame
    existing_ct_f_rate: pd.DataFrame
    existing_ct_f: pd.DataFrame
    existing_ct_s: pd.DataFrame
    do_ct_forecast: pd.DataFrame
    single_price: list[str]

    def __init__(
            self,
            selected_countries: list[str],
            simulation_years: tuple[int, int],
            # Preloaded input data:
            input_data: 'InputData',
            # Dashboard inputs:
            existing_ct_apply: bool,
            existing_ct_growth: float,
            ) -> None:
        last_simulation_year = simulation_years[1]

        self.__set_single_price(input_data.do_ct_ets_exist)
        self.__set_do_ct_forecast(existing_ct_apply)

        self.__set_existing_ct_s(
            selected_countries, input_data.existing_ct_s_raw, last_simulation_year
        )
        self.__set_existing_ct_f(
            selected_countries, input_data.existing_ct_f_raw, last_simulation_year
        )

        self.__set_existing_ct_rate(
            selected_countries, last_simulation_year, existing_ct_growth,
            input_data.existing_ct_rate_raw
        )
        self.__set_existing_ct_f_rate(
            selected_countries, existing_ct_growth,
            input_data.existing_ct_f_rate_raw, last_simulation_year
        )

    def __set_single_price(
            self,
            do_ct_ets_exist: pd.DataFrame
            ) -> None:
        """
        'Existing carbon pricing', 'Single price?' (v407)
        CPAT Excel: K81

        single_price: list[str]
        """
        self.single_price = list(
            do_ct_ets_exist[
                do_ct_ets_exist['Single price?'] == 'Yes'
            ].index.get_level_values(c.COUNTRY_CODE).unique()
        )


    def __set_existing_ct_s(
            self,
            selected_countries: list[str],
            existing_ct_s_raw: pd.DataFrame,
            last_simulation_year: int
            ) -> None:
        """
        'Existing carbon tax: sectoral coverage' (v407)
        CPAT Excel: 1737:1740

        existing_ct_s dims (c, s), all t <simulation_years[0] - 1, simulation years[1]>
        """
        existing_ct_s = existing_ct_s_raw.copy()

        # check if some countries are missing in existing_ct_s
        existing_countries = set(existing_ct_s.index.get_level_values(c.COUNTRY_CODE))
        missing_countries = set(selected_countries) - existing_countries

        if missing_countries:
            if existing_ct_s.empty:
                unique_sectors = pd.Index([c.POW, c.IND, 'trs', c.RES], name=c.SECTOR_CODE)
            else:
                unique_sectors = existing_ct_s.index.get_level_values(c.SECTOR_CODE).unique()
            missing_rows_indexes = pd.MultiIndex.from_product(
                [missing_countries, unique_sectors],
                names=[c.COUNTRY_CODE, c.SECTOR_CODE]
            )

            existing_ct_missing_countries = pd.DataFrame(
                0.0, index=missing_rows_indexes, columns=existing_ct_s_raw.columns
            )
            if existing_ct_s.empty:
                existing_ct_s = existing_ct_missing_countries.copy()
            else:
                existing_ct_s = pd.concat([existing_ct_s, existing_ct_missing_countries])

        # add the remaining years based on the last year values
        last_data_year = int(existing_ct_s_raw.columns[-1])
        new_years = range(last_data_year + 1, last_simulation_year + 1)
        for year in new_years:
            existing_ct_s[str(year)] = existing_ct_s[str(last_data_year)]

        self.existing_ct_s = existing_ct_s.sort_index()


    def __set_existing_ct_f(
            self,
            selected_countries: list[str],
            existing_ct_f_raw: pd.DataFrame,
            last_simulation_year: int
            ) -> None:
        """
        'Existing carbon tax: fuel coverage' (v407)
        CPAT Excel: 1741:1749

        existing_ct_f dims (c, f), all t
        """
        existing_ct_f = existing_ct_f_raw.copy()

        # check if some countries are missing in existing_ct_f
        existing_countries = set(existing_ct_f.index.get_level_values(c.COUNTRY_CODE))
        missing_countries = set(selected_countries) - existing_countries

        if missing_countries:
            if existing_ct_f.empty:
                unique_sectors = pd.Index(
                    c.EXISTING_ETS_CT_F_FUELS,
                    name=c.FUEL_CODE
                )
            else:
                unique_sectors = existing_ct_f.index.get_level_values(c.FUEL_CODE).unique()
            missing_rows_indexes = pd.MultiIndex.from_product(
                [missing_countries, unique_sectors],
                names=[c.COUNTRY_CODE, c.FUEL_CODE]
            )

            existing_ct_missing_countries = pd.DataFrame(
                0.0, index=missing_rows_indexes, columns=existing_ct_f.columns
            )
            if existing_ct_f.empty:
                existing_ct_f = existing_ct_missing_countries.copy()
            else:
                existing_ct_f = pd.concat([existing_ct_f, existing_ct_missing_countries])

        # add the remaining years based on the last year values
        last_data_year = int(existing_ct_f.columns[-1])
        new_years = range(last_data_year + 1, last_simulation_year + 1)
        for year in new_years:
            existing_ct_f[str(year)] = existing_ct_f[str(last_data_year)]

        self.existing_ct_f = existing_ct_f.sort_index()


    def __set_existing_ct_f_rate(
            self,
            selected_countries: list[str],
            existing_ct_growth: float,
            existing_ct_f_rate_raw: pd.DataFrame,
            last_simulation_year: int
            ) -> None:
        """
        'Existing carbon tax: fuel rate' (v407)
        CPAT Excel: 1728:1736

        For missing values 'Existing carbon tax rate: single rate'
        used as a fallback (not 0.0s), so it's easier later on when calculating ctx

        existing_ct_f_rate dims (c, f), all t
        """
        existing_ct_f_rate = existing_ct_f_rate_raw.droplevel(c.SECTOR_CODE)
        last_data_year = int(existing_ct_f_rate.columns[-1])

        # check if some countries are missing in existing_ct_f
        existing_countries = set(existing_ct_f_rate.index.get_level_values(c.COUNTRY_CODE))
        missing_countries = set(selected_countries) - existing_countries

        if existing_countries: # fill up remaining fuels
            countries_list = sorted(list(existing_countries))
            full_index = pd.MultiIndex.from_product(
                [countries_list, c.EXISTING_ETS_CT_F_FUELS],
                names=existing_ct_f_rate.index.names
            )
            existing_ct_f_rate = existing_ct_f_rate.reindex(full_index)

            existing_ct_rate_extended = self.existing_ct_rate[
                self.existing_ct_rate.index.get_level_values(c.COUNTRY_CODE).isin(countries_list)
            ]
            broadcast_index = pd.MultiIndex.from_product(
                [countries_list, c.EXISTING_ETS_CT_F_FUELS],
                names=[c.COUNTRY_CODE, c.FUEL_CODE]
            )

            existing_ct_rate_extended = (
                existing_ct_rate_extended.reindex(broadcast_index, level=c.COUNTRY_CODE)
            )
            existing_ct_f_rate = existing_ct_f_rate.combine_first(existing_ct_rate_extended)

        if missing_countries: # all fuels with fallback values
            countries_list = sorted(list(missing_countries))
            existing_ct_rate_extended = self.existing_ct_rate[
                self.existing_ct_rate.index.get_level_values(c.COUNTRY_CODE).isin(countries_list)
            ]
            broadcast_index = pd.MultiIndex.from_product(
                [countries_list, c.EXISTING_ETS_CT_F_FUELS],
                names=[c.COUNTRY_CODE, c.FUEL_CODE]
            )
            existing_ct_rate_extended = (
                existing_ct_rate_extended.reindex(broadcast_index, level=c.COUNTRY_CODE)
            )
            if existing_ct_f_rate.empty:
                existing_ct_f_rate = existing_ct_rate_extended.copy()
            else:
                existing_ct_f_rate = pd.concat([existing_ct_f_rate, existing_ct_rate_extended])

        # add the remaining years based on the last year values
        new_years = range(last_data_year + 1, last_simulation_year + 1)
        for year in new_years:
            existing_ct_f_rate[str(year)] = (
                existing_ct_f_rate[str(year - 1)]
                * (1 + existing_ct_growth)
            )

        self.existing_ct_f_rate = existing_ct_f_rate.copy()


    def __set_do_ct_forecast(
            self,
            existing_ct_apply: bool
            ) -> None:
        """
        'Existing carbon tax' (v407)
        CPAT Excel: C1725

        Redundant, kept for consitency with Excel

        do_ct_forecast: bool
        """
        self.do_ct_forecast = existing_ct_apply


    def __set_existing_ct_rate(
            self,
            selected_countries: list[str],
            last_simulation_year: int,
            existing_ct_growth: float,
            existing_ct_rate_raw: pd.DataFrame
            ) -> None:
        """
        'Existing carbon tax rate: single rate' (v407)
        CPAT Excel: 1727

        existing_ct_rate dims (c), <simulation_years[0] - 1, simulation_years[1]>
        """
        existing_ct_rate = existing_ct_rate_raw.droplevel(c.SECTOR_CODE).droplevel(c.FUEL_CODE)
        # check if some countries are missing in existing_ct_f
        existing_countries = set(existing_ct_rate.index.get_level_values(c.COUNTRY_CODE))
        missing_countries = set(selected_countries) - existing_countries

        if missing_countries:
            missing_countries_df = pd.DataFrame(
                0.0,
                index=pd.Index(sorted(missing_countries), name=c.COUNTRY_CODE),
                columns=existing_ct_rate.columns
            )
            if existing_ct_rate.empty:
                existing_ct_rate = missing_countries_df.copy()
            else:
                existing_ct_rate = pd.concat([existing_ct_rate, missing_countries_df])

        # add the remaining years based on the last year values
        last_data_year = int(existing_ct_rate.columns[-1])
        new_years = range(last_data_year + 1, last_simulation_year + 1)

        if self.do_ct_forecast:
            for year in new_years:
                existing_ct_rate[str(year)] = (
                    existing_ct_rate[str(year - 1)]
                    * (1 + existing_ct_growth)
                )
        else: # fill with 0.0s
            for year in new_years:
                existing_ct_rate[str(year)] = 0.0

        self.existing_ct_rate = existing_ct_rate.copy()


def get_existing_ct_raw(
        selected_countries: list[str],
        simulation_years: tuple[int, int]
        ) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    [Data Loading]
    'Sectoral coverage: carbon taxes' -> existing_ct_s_raw
        is derived from ECP!$B$428:$P$556 (v407)
    Manual changes:
    - columns removed: 'indicator code', 'country'
    - columns renamed: 'countrycode' -> 'CountryCode', 'sector/fuel' -> 'SectorCode'
    - only years 2020 - 2025 used

    'Fuel coverage: carbon taxes' -> existing_ct_f_raw
        is derived from ECP!$B$136:$P$424 (v407)
    Manual changes:
    - columns removed: 'indicator code', 'country'
    - columns renamed: 'countrycode' -> 'CountryCode', 'sector/fuel' -> 'FuelCode'
    - only years 2020 - 2025 used

    'Carbon tax rates, $/CO2 t (nominal)' -> existing_ct_rate_raw
        is derived from ECP!$B$68:$P$131 (v407)
    Manual changes:
    - columns removed: 'indicator code', 'country'
    - columns renamed: 'countrycode' -> 'CountryCode'
    - 'Fuel' renamed to 'FuelCode'; CAREFUL! Some fuel codes in the Excel end with space ' '.
        This spaces need to be removed
    - column 'SectorCode' added populated with 'all' values for all 'Carbon tax rates, ...' rows
    - only years 2020 - 2025 used

    returns dims:
    - existing_ct_s_raw: (c, s), <simulation_years[0] - 1, last db year>
    - existing_ct_f_raw: (c, f), <simulation_years[0] - 1, last db year>
    - existing_ct_f_rate_raw: (c, f), <simulation_years[0] - 1, last db year>
    - existing_ct_rate_raw: (c, s, f), <simulation_years[0] - 1, last db year>
    """
    existing_ct_raw: pd.DataFrame = (
        # TODO: pkl
        # pd.read_pickle(f'{DATA_PATH}/existing_ct.pkl.bz2', compression="bz2")
        pd.read_csv(f'{DATA_PATH}/existing_ct.csv')
    )

    last_db_year = max(map(int, existing_ct_raw.columns[3:]))
    cols_to_keep = [str(year) for year in range(simulation_years[0] - 1, last_db_year + 1)]
    existing_ct_raw = existing_ct_raw[c.ID_COL_NAMES + cols_to_keep]

    # check for duplicated entries for (c, s, f) combination
    dupes = existing_ct_raw[existing_ct_raw.duplicated(subset=c.ID_COL_NAMES, keep=False)]
    assert dupes.empty, f"Found duplicated rows in existing_ct_raw:\n{dupes}"

    existing_ct_raw = existing_ct_raw[
        existing_ct_raw[c.COUNTRY_CODE].isin(selected_countries)
    ]

    existing_ct_rate_raw = (
        existing_ct_raw[
            existing_ct_raw[c.SECTOR_CODE].notna() & existing_ct_raw[c.FUEL_CODE].notna()
        ]
        .fillna(0.0)
    )
    existing_ct_f_rate_raw = (
        existing_ct_rate_raw[
            existing_ct_rate_raw[c.FUEL_CODE] != c.ALL
        ]
        .set_index(c.ID_COL_NAMES)
    )
    existing_ct_rate_raw = (
        existing_ct_rate_raw[
            existing_ct_rate_raw[c.FUEL_CODE] == c.ALL
        ]
        .set_index(c.ID_COL_NAMES)
    )
    existing_ct_raw_s = (
        existing_ct_raw[existing_ct_raw[c.FUEL_CODE].isna()]
        .set_index([c.COUNTRY_CODE, c.SECTOR_CODE])
        .drop(columns=[c.FUEL_CODE])
    ).fillna(0.0)
    existing_ct_raw_f = (
        existing_ct_raw[existing_ct_raw[c.SECTOR_CODE].isna()]
        .set_index([c.COUNTRY_CODE, c.FUEL_CODE])
        .drop(columns=[c.SECTOR_CODE])
    ).fillna(0.0)

    return existing_ct_raw_s, existing_ct_raw_f, existing_ct_f_rate_raw, existing_ct_rate_raw

def load_do_ct_ets_exist(
        selected_countries: list[str]
        ) -> pd.DataFrame:
    """
    [Data Loading]
    v407

    Manual changes:
    - 'Country' column removed
    - 'Country code' column renamed to 'CountryCode'
    - for 2 level header columns, lower ones were used as column names

    returns dims (c), mechanisms
    """
    do_ct_ets_exist: pd.DataFrame = (
        # TODO: pkl
        pd.read_csv(f'{DATA_PATH}/do_ct_ets_exist.csv')
    )

    do_ct_ets_exist = do_ct_ets_exist[
        do_ct_ets_exist[c.COUNTRY_CODE].isin(selected_countries)
    ]

    do_ct_ets_exist.set_index(c.COUNTRY_CODE, inplace=True)

    return do_ct_ets_exist
