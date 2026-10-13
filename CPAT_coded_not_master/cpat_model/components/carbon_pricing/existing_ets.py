from typing import TYPE_CHECKING

import pandas as pd

import cpat_model.constants as c
from config import DATA_PATH

if TYPE_CHECKING:
    from cpat_model.inputs.input_data import InputData


class ExistingETS:
    """
    Existing ETS from
    'Existing carbon pricing mechanisms (ETSs and carbon taxes)' section (v412)    
    CPAT Excel: section start 1675

    Please mind helper columns are off in v412.
    """
    eu_ets: list[str]
    existing_ets_permit_price_nat: pd.DataFrame
    existing_ets_f_nat: pd.DataFrame
    existing_ets_s_nat: pd.DataFrame
    existing_ets_permit_price_reg: pd.DataFrame
    existing_ets_f_reg: pd.DataFrame
    existing_ets_s_reg: pd.DataFrame
    do_ets_forecast: pd.DataFrame
    do_ets_forecast: pd.DataFrame

    def __init__(
            self,
            selected_countries: list[str],
            simulation_years: tuple[int, int],
            # Preloaded input data:
            input_data: 'InputData',
            # Dashboard inputs:
            existing_ets_apply: bool,
            existing_ets_growth: float,
            ) -> None:
        last_simulation_year = simulation_years[1]

        self.__set_eu_ets(input_data.do_ct_ets_exist)
        self.__set_do_ets_forecast(existing_ets_apply)

        self.existing_ets_s_nat = self.__get_existing_ets_s(
            input_data.existing_ets_s_nat_raw, selected_countries, last_simulation_year
        )
        self.existing_ets_f_nat = self.__get_existing_ets_f(
            input_data.existing_ets_f_nat_raw, selected_countries, last_simulation_year
        )
        self.__set_existing_ets_permit_price_nat(
            selected_countries, last_simulation_year, existing_ets_growth,
            input_data.existing_ets_permit_price_nat_raw
        )

        self.existing_ets_s_reg = self.__get_existing_ets_s(
            input_data.existing_ets_s_reg_raw, selected_countries, last_simulation_year
        )
        self.existing_ets_f_reg = self.__get_existing_ets_f(
            input_data.existing_ets_f_reg_raw, selected_countries, last_simulation_year
        )
        self.__set_existing_ets_permit_price_reg(
            selected_countries, last_simulation_year, existing_ets_growth,
            input_data.existing_ets_permit_price_reg_raw
        )


    def __set_eu_ets(
            self,
            do_ct_ets_exist: pd.DataFrame
            ) -> None:
        """
        'Existing carbon pricing', 'EU ETS?' (v412)
        CPAT Excel: L81

        'Is it EU ETS?' is 'Yes' or 'Both', have the same effect in CPAT (as per v412),
        therefore stored in a list instead of dict.

        eu_ets: list[str]
        """
        self.eu_ets = list(
            do_ct_ets_exist[
                (do_ct_ets_exist['Is it EU ETS?'] == 'Yes')
                | (do_ct_ets_exist['Is it EU ETS?'] == 'Both')
            ].index.get_level_values(c.COUNTRY_CODE).unique()
        )

    def __get_existing_ets_s(
            self,
            existing_ets_s_raw: pd.DataFrame,
            selected_countries: list[str],
            last_simulation_year: int
            ) -> pd.DataFrame:
        """
        'Existing ETS: sectoral coverage' (v412)
        'National ETS' CPAT Excel: 1681:1684
        'Regional ETS' CPAT Excel: 1696:1699

        returns (c, s), <simulation_years[0] - 1, last db year>
        """
        existing_ets = existing_ets_s_raw.copy()

        # check if some countries are missing in existing_ets_s_nat_raw
        existing_countries = set(existing_ets_s_raw.index.get_level_values(c.COUNTRY_CODE))
        missing_countries = set(selected_countries) - existing_countries

        # create rows for countries not included in existing_ets_s_nat_raw and fill with 0s
        if missing_countries:
            unique_sectors = pd.Index([c.POW, c.IND, 'trs', c.RES], name=c.SECTOR_CODE)
            missing_rows_indexes = pd.MultiIndex.from_product(
                [missing_countries, unique_sectors],
                names=[c.COUNTRY_CODE, c.SECTOR_CODE]
            )

            existing_ets_missing_countries = pd.DataFrame(
                0.0, index=missing_rows_indexes, columns=existing_ets_s_raw.columns
            )
            if existing_ets.empty:
                existing_ets = existing_ets_missing_countries.copy()
            else:
                existing_ets = pd.concat([existing_ets, existing_ets_missing_countries])

        # add the remaining years based on the last year values
        last_data_year = int(existing_ets_s_raw.columns[-1])
        new_years = range(last_data_year + 1, last_simulation_year + 1)
        for year in new_years:
            existing_ets[str(year)] = existing_ets[str(last_data_year)]

        return existing_ets.sort_index()


    def __get_existing_ets_f(
            self,
            existing_ets_f_raw: pd.DataFrame, # (c, f)
            selected_countries: list[str],
            last_simulation_year: int
            ) -> pd.DataFrame:
        """
        'Existing ETS: fuel coverage' (v412)
        'National ETS' CPAT Excel: 1685:1693
        'Regional ETS' CPAT Excel: 1700:1708

        returns dims (c, f), <simulation_years[0] - 1, last db year>
        """
        existing_ets = existing_ets_f_raw.copy()

        # check if some countries are missing in existing_ets_f_nat_raw
        existing_countries = set(existing_ets_f_raw.index.get_level_values(c.COUNTRY_CODE))
        missing_countries = set(selected_countries) - existing_countries

        if missing_countries:
            unique_sectors = pd.Index(c.EXISTING_ETS_CT_F_FUELS, name=c.FUEL_CODE)
            missing_rows_indexes = pd.MultiIndex.from_product(
                [missing_countries, unique_sectors],
                names=[c.COUNTRY_CODE, c.FUEL_CODE]
            )

            existing_ets_missing_countries = pd.DataFrame(
                0.0, index=missing_rows_indexes, columns=existing_ets.columns
            )
            if existing_ets.empty:
                existing_ets = existing_ets_missing_countries.copy()
            else:
                existing_ets = pd.concat([existing_ets, existing_ets_missing_countries])

        # add the remaining years based on the last year values
        last_data_year = int(existing_ets_f_raw.columns[-1])
        new_years = range(last_data_year + 1, last_simulation_year + 1)
        for year in new_years:
            existing_ets[str(year)] = existing_ets[str(last_data_year)]


        return existing_ets.sort_index()


    def __set_do_ets_forecast(
            self,
            existing_ets_apply: bool
            ) -> None:
        """
        'Existing ETS' (v412)
        CPAT Excel: C1677

        Redundant, kept for consitency with Excel

        do_ets_forecast: bool
        """
        self.do_ets_forecast = existing_ets_apply


    def __set_existing_ets_permit_price_nat(
            self,
            selected_countries: list[str],
            last_simulation_year: int,
            existing_ets_growth: float,
            existing_ets_permit_price_nat_raw: pd.DataFrame
            ) -> None:
        """
        'Existing ETS permit price' (v412)
        CPAT Excel: 1680

        existing_ets_permit_price_nat dims (c), <simulation_years[0] - 1, last db year>
        """
        existing_ets = (
            existing_ets_permit_price_nat_raw
            .droplevel(c.FUEL_CODE)
            .droplevel(c.SECTOR_CODE)
        )

        # check if some countries are missing in existing_ets_f_nat_raw
        existing_countries = set(
            existing_ets_permit_price_nat_raw.index.get_level_values(c.COUNTRY_CODE)
        )
        missing_countries = set(selected_countries) - existing_countries

        if missing_countries:
            missing_countries_df = pd.DataFrame(
                0.0,
                index=pd.Index(sorted(missing_countries), name=c.COUNTRY_CODE),
                columns=existing_ets_permit_price_nat_raw.columns
            )
            if existing_ets.empty:
                existing_ets = missing_countries_df.copy()
            else:
                existing_ets = pd.concat([existing_ets, missing_countries_df])

        # add the remaining years based on the last year values
        last_data_year = int(existing_ets_permit_price_nat_raw.columns[-1])
        new_years = range(last_data_year + 1, last_simulation_year + 1)

        if self.do_ets_forecast:
            for year in new_years:
                existing_ets[str(year)] = (
                    existing_ets[str(year - 1)]
                    * (1 + existing_ets_growth)
                )
        else: # fill with 0.0s
            for year in new_years:
                existing_ets[str(year)] = 0.0


        self.existing_ets_permit_price_nat = existing_ets.sort_index()


    def __set_existing_ets_permit_price_reg(
            self,
            selected_countries: list[str],
            last_simulation_year: int,
            existing_ets_growth: float,
            existing_ets_permit_price_reg_raw: pd.DataFrame
            ) -> None:
        """
        'Existing ETS permit price' (v412)
        CPAT Excel: 1695

        existing_ets_permit_price_reg dims (c), <simulation_years[0] - 1, last db year>
        """
        existing_ets = (
            existing_ets_permit_price_reg_raw
            .droplevel(c.FUEL_CODE)
            .droplevel(c.SECTOR_CODE)
        )
        missing_countries = set(selected_countries) - set(self.eu_ets)

        # self.eu_ets countries EU row ->
        if self.eu_ets:
            eu_values = existing_ets.loc['EU']

            existing_ets = pd.DataFrame(
                [eu_values.values] * len(self.eu_ets),
                index=pd.Index(self.eu_ets, name=c.COUNTRY_CODE),
                columns=existing_ets.columns
            )

        if missing_countries:
            if 'EU' in existing_ets.index.get_level_values(c.COUNTRY_CODE):
                existing_ets = existing_ets.drop('EU')
            missing_df = pd.DataFrame(
                0.0,
                index=pd.Index(list(missing_countries), name=c.COUNTRY_CODE),
                columns=existing_ets.columns
            )
            existing_ets = pd.concat([existing_ets, missing_df])


        # add the remaining years based on the last year values
        last_data_year = int(existing_ets_permit_price_reg_raw.columns[-1])
        new_years = range(last_data_year + 1, last_simulation_year + 1)

        if self.do_ets_forecast:
            for year in new_years:
                existing_ets[str(year)] = (
                    existing_ets[str(year - 1)]
                    * (1 + existing_ets_growth)
                )
        else: # fill with 0.0s
            for year in new_years:
                existing_ets[str(year)] = 0.0


        self.existing_ets_permit_price_reg = existing_ets.sort_index()


def get_existing_ets_nat_raw(
        selected_countries: list[str],
        simulation_years: tuple[int, int],
        ) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    [Data Loading]
    'NationalETS'

    'Sector coverage: ETS' -> existing_ets_s_nat_raw
        is derived from ECP!$B$1041:$P$1092 (v412)
    Manual changes:
    - columns removed: 'Indicator code', 'country'
    - columns renamed: 'countrycode' -> 'CountryCode', 'sector/fuel' -> 'SectorCode'
    - years before 2020 were cut

    'Fuel coverage: ETS' -> existing_ets_f_nat_raw
        is derived from ECP!$B$620:$P$736 (v412)
    Manual changes:
    - columns removed: 'Indicator code', 'country'
    - columns renamed: 'countrycode' -> 'CountryCode', 'sector/fuel' -> 'FuelCode'
    - years before 2020 were cut

    'ETS permit price, $/CO2 t (nominal)' -> existing_ets_permit_price_nat_raw
        is derived from ECP!$B$601:$P$613 (v412)
    Manual changes:
    - columns removed: 'Indicator', 'Indicator code', 'country'
    - columns renamed: 'countrycode' -> 'CountryCode'
    - years before 2020 were cut
    - column 'sector/fuel' extended to 2 columns -> 'SectorCode' and 'FuelCode'
        (populated with 'all' values)

    returns dims:
    - existing_ets_s_nat_raw: (c, s), <simulation_years[0] - 1, last db year>
    - existing_ets_f_nat_raw: (c, f), <simulation_years[0] - 1, last db year>
    - existing_ets_permit_price_nat_raw: (c, s, f), <simulation_years[0] - 1, last db year>
    """
    existing_ets_raw: pd.DataFrame = (
        # TODO: pkl
        # pd.read_pickle(f'{DATA_PATH}/existing_ets_nat.pkl.bz2', compression="bz2")
        pd.read_csv(f'{DATA_PATH}/existing_ets_nat.csv')
    )

    last_db_year = max(map(int, existing_ets_raw.columns[3:]))
    cols_to_keep = [str(year) for year in range(simulation_years[0] - 1, last_db_year + 1)]
    existing_ets_raw = existing_ets_raw[c.ID_COL_NAMES + cols_to_keep]

    # check for duplicated entries for (c, s, f) combination
    dupes = existing_ets_raw[existing_ets_raw.duplicated(subset=c.ID_COL_NAMES, keep=False)]
    assert dupes.empty, f"Found duplicated rows in existing_ets_raw:\n{dupes}"


    existing_ets_raw = existing_ets_raw[
        existing_ets_raw[c.COUNTRY_CODE].isin(selected_countries)
    ]

    existing_ets_permit_price_nat_raw = (
        existing_ets_raw[
            existing_ets_raw[c.SECTOR_CODE].notna() & existing_ets_raw[c.FUEL_CODE].notna()
        ]
        .set_index(c.ID_COL_NAMES)
        .fillna(0.0)
    )
    existing_ets_s_nat_raw = (
        existing_ets_raw[existing_ets_raw[c.FUEL_CODE].isna()]
        .set_index([c.COUNTRY_CODE, c.SECTOR_CODE])
        .drop(columns=[c.FUEL_CODE])
        .fillna(0.0)
    )
    existing_ets_f_nat_raw = (
        existing_ets_raw[existing_ets_raw[c.SECTOR_CODE].isna()]
        .set_index([c.COUNTRY_CODE, c.FUEL_CODE])
        .drop(columns=[c.SECTOR_CODE])
        .fillna(0.0)
    )

    return existing_ets_s_nat_raw, existing_ets_f_nat_raw, existing_ets_permit_price_nat_raw

def get_existing_ets_reg_raw(
        selected_countries: list[str],
        simulation_years: tuple[int, int]
        ) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    [Data Loading]
    'RegionalETS'

    'Sector coverage: ETS' -> existing_ets_s_reg_raw
        is derived from ECP!$B$1094:$P$1221 (v412)
    Manual changes:
    - columns removed: 'Indicator code', 'country'
    - columns renamed: 'countrycode' -> 'CountryCode', 'sector/fuel' -> 'SectorCode'
    - years before 2020 were cut

    'Fuel coverage: ETS' -> existing_ets_f_reg_raw
        is derived from ECP!$B$738:$P$1034 (v412)
    Manual changes:
    - columns removed: 'Indicator code', 'country'
    - columns renamed: 'countrycode' -> 'CountryCode', 'sector/fuel' -> 'FuelCode'
    - years before 2020 were cut

    'ETS permit price, $/CO2 t (nominal)' -> existing_ets_permit_price_reg_raw
        is derived from ECP!$B$615:$P$615 (v412)
    Manual changes:
    - columns removed: 'Indicator', 'Indicator code', 'country'
    - columns renamed: 'countrycode' -> 'CountryCode'
    - years before 2020 were cut
    - column 'sector/fuel' extended to 2 columns -> 'SectorCode' and 'FuelCode'
        (populated with 'all' values)

    returns dims:
    - existing_ets_s_reg_raw: (c, s), <simulation_years[0] - 1, last db year>
    - existing_ets_f_reg_raw: (c, f), <simulation_years[0] - 1, last db year>
    - existing_ets_permit_price_reg_raw: (c, s, f), <simulation_years[0] - 1, last db year>
    """
    existing_ets_raw: pd.DataFrame = (
        # TODO: pkl
        # pd.read_pickle(f'{DATA_PATH}/existing_ets_reg.pkl.bz2', compression="bz2")
        pd.read_csv(f'{DATA_PATH}/existing_ets_reg.csv')
    )
    last_db_year = max(map(int, existing_ets_raw.columns[3:]))
    cols_to_keep = [str(year) for year in range(simulation_years[0] - 1, last_db_year + 1)]
    existing_ets_raw = existing_ets_raw[c.ID_COL_NAMES + cols_to_keep]

    # check for duplicated entries for (c, s, f) combination
    dupes = existing_ets_raw[existing_ets_raw.duplicated(subset=c.ID_COL_NAMES, keep=False)]
    assert dupes.empty, f"Found duplicated rows in existing_ets_raw:\n{dupes}"

    existing_ets_raw = existing_ets_raw[
        existing_ets_raw[c.COUNTRY_CODE].isin(selected_countries + ['EU'])
    ]

    existing_ets_permit_price_reg_raw = (
        existing_ets_raw[
            existing_ets_raw[c.SECTOR_CODE].notna() & existing_ets_raw[c.FUEL_CODE].notna()
        ]
        .set_index(c.ID_COL_NAMES)
        .fillna(0.0)
    )

    existing_ets_raw = existing_ets_raw[
        # permit price should be EU only
        # remove EU as it is not used for existing_ets_f
        existing_ets_raw[c.COUNTRY_CODE] != 'EU'
    ]
    existing_ets_s_reg_raw = (
        existing_ets_raw[existing_ets_raw[c.FUEL_CODE].isna()]
        .set_index([c.COUNTRY_CODE, c.SECTOR_CODE])
        .drop(columns=[c.FUEL_CODE])
        .fillna(0.0)
    )
    existing_ets_f_reg_raw = (
        existing_ets_raw[existing_ets_raw[c.SECTOR_CODE].isna()]
        .set_index([c.COUNTRY_CODE, c.FUEL_CODE])
        .drop(columns=[c.SECTOR_CODE])
        .fillna(0.0)
    )

    return existing_ets_s_reg_raw, existing_ets_f_reg_raw, existing_ets_permit_price_reg_raw
