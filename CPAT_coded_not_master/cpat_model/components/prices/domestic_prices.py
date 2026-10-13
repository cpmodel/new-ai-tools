from typing import Literal, TYPE_CHECKING
from enum import Enum

import pandas as pd
import numpy as np

import cpat_model.constants as c
from config import DATA_PATH

if TYPE_CHECKING:
    from cpat_model.inputs.input_data import InputData

PRICES_DATA_FILE_NAME = 'prices_dom'

class Headers:
    """
    Constants for header names in data from 'Price_dom' tab.
    """
    RETAIL_PRICE = 'Retail price'
    SUPPLY_COSTS = 'Supply costs'
    EXCISE_AND_OTHER_TAXES = 'Excise and other taxes'
    VAT_RATE = 'VAT rate'
    MARGIN_ON_TOP_OF_SUPPLY_COST = 'Margin on top of supply costs'
    RAW_PRICE_CONTROL_COEFFICIENT = 'Raw price control coefficient'
    DOMESTIC_PRODUCTION_COST = 'Domestic production cost'
    BUCKETED_PRICE_CONTROL_COEFFICIENT = 'Bucketed price control coefficient'
    MANUAL_PRICE_CONTROL_COEFFICIENTS = 'Manual price control coefficients'
    CHOSEN_PRICE_CONTROL_COEFFICIENT = 'Chosen price control coefficient'
    PRODUCER_SIDE_SUBSIDY = 'Producer-side subsidy'

PRICE_TO_CODE_MAPPING = {
    Headers.RETAIL_PRICE: 'mit.rp.',
    Headers.SUPPLY_COSTS: 'mit.sp.',
    Headers.EXCISE_AND_OTHER_TAXES: 'mit.txo.',
    Headers.VAT_RATE: 'mit.vatrate.',
    Headers.MARGIN_ON_TOP_OF_SUPPLY_COST: 'mit.mar.',
    Headers.RAW_PRICE_CONTROL_COEFFICIENT: 'mit.ps.',
    Headers.DOMESTIC_PRODUCTION_COST: ['mit.coa.prod.cost', 'mit.nga.prod.cost'],
    Headers.PRODUCER_SIDE_SUBSIDY: 'mit.pros.'
}

class ForecastingCoefficientsType(str, Enum):
    """
    Type for forecasting_coefficients
    """
    MARGIN_ON_TOP_OF_SUPPLY_COST = Headers.MARGIN_ON_TOP_OF_SUPPLY_COST
    RAW_PRICE_CONTROL_COEFFICIENT = Headers.RAW_PRICE_CONTROL_COEFFICIENT
    DOMESTIC_PRODUCTION_COST = Headers.DOMESTIC_PRODUCTION_COST
    BUCKETED_PRICE_CONTROL_COEFFICIENT = Headers.BUCKETED_PRICE_CONTROL_COEFFICIENT
    MANUAL_PRICE_CONTROL_COEFFICIENTS = Headers.MANUAL_PRICE_CONTROL_COEFFICIENTS
    CHOSEN_PRICE_CONTROL_COEFFICIENT = Headers.CHOSEN_PRICE_CONTROL_COEFFICIENT

INDEX_PAIRS = (
    [(c.ALL, fuel) for fuel in [c.BIO, c.JFU, c.OOP]]
    + [(c.AVI, c.JFU)]
    + [
        (sector, fuel)
        for sector in [c. IND, c.POW, c.RES, c.ROD]
        for fuel in [c.COA, c.DIE, c.GSO, c.KER, c.LPG, c.NGA, c.OOP]
    ]
)
MANUAL_PRICES = pd.DataFrame( # TODO: input 'Manual inputs'!J115:J128
    0.8,
    index=pd.MultiIndex.from_tuples(
        [
            *[(sector, fuel) for sector in [c.IND, c.POW, c.RES] for fuel in [c.COA, c.NGA]],
            (c.IND, c.ELE), (c.RES, c.ELE),
            *[(c.ALL, fuel) for fuel in [c.BIO, c.DIE, c.GSO, c.KER, c.LPG, c.OOP]]
        ],
        names=[c.SECTOR_CODE, c.FUEL_CODE]
    ),
    columns=[Headers.MANUAL_PRICE_CONTROL_COEFFICIENTS]
)

class DomPrices:
    """
    Combines:
    'Historical prices' table from v361 Mitigation!B912:L921
    and
    'Energy price forecasting coefficients' table from v361 Mitigation!B714:R728.
    """
    rp: pd.DataFrame
    sp: pd.DataFrame
    txo: pd.DataFrame
    vatrate: pd.DataFrame
    ps: pd.DataFrame
    forecasting_coefficients: dict[ForecastingCoefficientsType, pd.DataFrame]
    general_vatrate: pd.DataFrame # VAT_WEO

    def __init__(
            self,
            selected_countries: list[str],
            simulation_years: tuple[int, int], # TODO: filter years to simulation_years[0] - 1, assert last hist year is the same always
            input_data: 'InputData',
            gov_energy_price_controls: Literal['Bucketed', 'Manual', 'None'],
            manual_prices: pd.DataFrame = MANUAL_PRICES
            ) -> None:
        """
        'Historical prices' table from v361 Mitigation!B912:L921
        and
        'Energy price forecasting coefficients' table from v361 Mitigation!B714:R728.

        Columns not included in 'Energy price forecasting coefficients':
        - 'Produced domestically'
        - 'Override subsidy'

        'Historical prices' TODO rows from 731?
        

        """
        self.rp = pd.DataFrame()
        self.sp = pd.DataFrame()
        self.txo = pd.DataFrame()
        self.vatrate = pd.DataFrame()
        self.ps = pd.DataFrame()
        self.forecasting_coefficients = {}

        time_series_prices = {
            Headers.RETAIL_PRICE: 'rp',
            Headers.SUPPLY_COSTS: 'sp',
            Headers.EXCISE_AND_OTHER_TAXES: 'txo',
            Headers.VAT_RATE: 'vatrate',
            Headers.PRODUCER_SIDE_SUBSIDY: 'ps'
        }
        self.__set_general_vatrate(input_data.prices_dom, simulation_years)

        for price_name, code in PRICE_TO_CODE_MAPPING.items():
            df = self.__get_time_series_df(input_data, code)

            # decomposing code to Fuel and Sector
            if price_name == Headers.DOMESTIC_PRODUCTION_COST:
                self.__get_domestic_production_cost(simulation_years, df)
            else:
                # operating directly on df, not a copy so assignment a bit redundat
                df = self.__get_sector_fuel_cols_df(df, code)
                if price_name in time_series_prices:
                    # covered in test_dom_prices
                    df = self.__cut_years_prior_to_start_year(simulation_years, df)
                    df = df.rename(index={"ecy": c.ELE}, level=c.FUEL_CODE)
                    df.sort_index(inplace=True)
                    setattr(self, time_series_prices[price_name], df.fillna(0.0))

                if price_name in [
                    Headers.MARGIN_ON_TOP_OF_SUPPLY_COST,
                    Headers.RAW_PRICE_CONTROL_COEFFICIENT
                ]:
                    self.__set_single_year_coeffitients(
                        selected_countries, simulation_years, df, price_name
                    )

        self.__set_bucketed_price_control_coefficient()
        self.__set_manual_price_control_coefficients(selected_countries, manual_prices)
        self.__set_chosen_price_control_coefficient(gov_energy_price_controls)

        self.__assert_indexes(selected_countries)


    def __set_general_vatrate(
            self,
            prices_dom: pd.DataFrame,
            simulation_years: tuple[int, int]
        ) -> None:
        """
        'VAT rate', 'VAT_WEO' from 'General data' table
        CPAT Excel: row 705 (v407)

        general_vatrate dims (c), t in <simulation_years[0] - 1, simulation_years[1]>
        """
        # TODO test
        self.general_vatrate = prices_dom[[c.COUNTRY_CODE, 'year', 'VAT_WEO']]
        self.general_vatrate = (
            self.general_vatrate
            .pivot(index=c.COUNTRY_CODE, columns='year', values='VAT_WEO')
        )
        self.general_vatrate.columns.name = None
        # year col names to str:
        self.general_vatrate = self.general_vatrate.rename(columns=str)

        last_hist_year = max(map(int, self.general_vatrate.columns))
        columns_to_stay = [str(year) for year in range(simulation_years[0] - 1, last_hist_year + 1)]
        self.general_vatrate = self.general_vatrate[columns_to_stay]

        columns = [str(year) for year in range(last_hist_year + 1, simulation_years[1] + 1)]
        for col in columns:
            self.general_vatrate[col] = self.general_vatrate[str(last_hist_year)]


    def __get_time_series_df(
            self,
            input_data: 'InputData',
            code: str | list[str]
            ) -> pd.DataFrame:
        """
        Helper function.
        Filters columns for given price code and pivots it to time series format.
        """
        exact_names_to_keep = [c.COUNTRY_CODE, 'year']
        df = input_data.prices_dom.copy()
        # filtering codes
        columns_to_keep = [
            col for col in input_data.prices_dom.columns
            if (col in exact_names_to_keep or self.__get_keep_col_condiotion(code, col))
        ]

        # and pivoting
        df = (
            df[columns_to_keep]
            .melt(id_vars=[c.COUNTRY_CODE, 'year'], var_name=c.FUEL_CODE, value_name='Value')
            .pivot(index=[c.COUNTRY_CODE, c.FUEL_CODE], columns='year', values='Value')
            .reset_index()
        )
        df.columns.name = None

        return df

    def __get_sector_fuel_cols_df(
            self,
            df: pd.DataFrame,
            code: str
        ) -> pd.DataFrame:
        """
        Helper function.
        Extracts sector and fuel codes from CPAT code.
        Sets int year columns to str type.
        """
        df[c.FUEL_CODE] = (
            df[c.FUEL_CODE].str.removeprefix(code)
        )
        df[[c.FUEL_CODE, c.SECTOR_CODE]] = (
            df[c.FUEL_CODE].str.split('.', n=1, expand=True)
        )
        df.set_index(c.ID_COL_NAMES, inplace=True)

        # column names (years) to str
        df.columns = df.columns.astype(str)

        return df

    def __cut_years_prior_to_start_year(
            self,
            simulation_years: tuple[int, int],
            df: pd.DataFrame
            ) -> pd.DataFrame:
        """
        Removes years prior to simulation_years[0] - 1.
        Applies to time series domestic price attributes.
        """
        return df[[col for col in df.columns if int(col) >= (simulation_years[0] - 1)]]


    def __get_domestic_production_cost(
            self,
            simulation_years: tuple[int, int],
            df: pd.DataFrame
            ) -> None:
        """
        'Domestic production cost'
        dims (c, s, f)
        where:
        f in [coa, nga], s in [ind, pow , res] -> 2x3 = 6 per country
        """
        price_name = Headers.DOMESTIC_PRODUCTION_COST
        df[c.FUEL_CODE] = (
            df[c.FUEL_CODE].str[4:7] # leave just the fuel code
        )
        df.set_index([c.COUNTRY_CODE, c.FUEL_CODE], inplace=True)
        df = (
            # columns not str type yet!
            df[[simulation_years[0]]]
            .rename(columns={simulation_years[0]: price_name})
        )
        # extending by sectors:
        sector_values = [c.IND, c.POW, c.RES]
        df = (
            df
            .reindex(df.index.repeat(len(sector_values)))
        )
        df.index = pd.MultiIndex.from_tuples(
            [
                (country, sector, fuel)
                for (country, fuel) in df.index[::len(sector_values)]
                for sector in sector_values
            ],
            names=c.ID_COL_NAMES
        )
        df.sort_index(inplace=True)
        df.columns = df.columns.astype(str)

        self.forecasting_coefficients[price_name] = df.copy()

    def __set_single_year_coeffitients(
            self,
            selected_countries: list[str],
            simulation_years: tuple[int, int],
            df: pd.DataFrame,
            price_name: str
        ) -> None:
        """
        Sets MARGIN_ON_TOP_OF_SUPPLY_COST and RAW_PRICE_CONTROL_COEFFICIENT
        dims (c, s, f)
        where:
        for 'Raw price control coefficient' (s, f) in:
            [coa, nga]x[pow, ind, res] + [gso, die, lpg, ker, oop, bio]x[all] + [ele]x[res, ind]
        for 'Margin on top of supply costs' (s, f) same as above BUT NO ele
        """
        self.forecasting_coefficients[price_name] = (
            df[[str(simulation_years[0])]]
            .rename(columns={str(simulation_years[0]): price_name})
        )
        if price_name == Headers.MARGIN_ON_TOP_OF_SUPPLY_COST: # TODO test this and next one
            # oop added to mit.mar data, we still hardcode it as 0.0
            self.forecasting_coefficients[price_name] = (
                self.forecasting_coefficients[price_name][
                    self.forecasting_coefficients[price_name][price_name]
                    .index.get_level_values(c.FUEL_CODE) != c.OOP
                ]
            )
            # gso, die. lpg, ker
            # oop, ecy

        if price_name == Headers.RAW_PRICE_CONTROL_COEFFICIENT:
            # oop, ecy added to mit.ps data, we still hardcode it as 1.0
            self.forecasting_coefficients[price_name] = (
                self.forecasting_coefficients[price_name][
                    ~self.forecasting_coefficients[price_name][price_name]
                    .index.get_level_values(c.FUEL_CODE).isin([c.OOP, 'ecy'])
                ]
            )
        # adding oop and bio as 0s for MARGIN_ON_TOP_OF_SUPPLY_COST
        # and ele, oop, bio as 1s for RAW_PRICE_CONTROL_COEFFICIENT
        new_rows_index_pairs = [
            *[(country_code, c.ALL, c.OOP) for country_code in selected_countries],
            *[(country_code, c.ALL, c.BIO) for country_code in selected_countries],
            *(
                [
                    *[(country_code, c.IND, c.ELE) for country_code in selected_countries],
                    *[(country_code, c.RES, c.ELE) for country_code in selected_countries]
                ]
                if price_name == Headers.RAW_PRICE_CONTROL_COEFFICIENT
                else []
            )
        ]
        new_index = pd.MultiIndex.from_tuples(
            new_rows_index_pairs,
            names=self.forecasting_coefficients[price_name].index.names
        )
        new_rows_value = 0.0 if price_name == Headers.MARGIN_ON_TOP_OF_SUPPLY_COST else 1.0
        new_rows = pd.DataFrame(
            new_rows_value,
            index=new_index,
            columns=self.forecasting_coefficients[price_name].columns
        )
        self.forecasting_coefficients[price_name] = pd.concat(
            [self.forecasting_coefficients[price_name], new_rows]
        ).sort_index()


    def __set_bucketed_price_control_coefficient(
            self
            ) -> None:
        """
        'Bucketed price control coefficient'

        dims (c, s, f)
        where:
        (s, f) in:
            [coa, nga]x[pow, ind, res] + [gso, die, lpg, ker, oop, bio]x[all] + [ele]x[res, ind]
        """
        price_name = Headers.RAW_PRICE_CONTROL_COEFFICIENT
        # buckets and tresholds:
        bins = [-np.inf, 0.25, 0.5, np.inf]
        labels = [0.0, 0.5, 1.0]

        new_col = pd.cut(
            self.forecasting_coefficients[price_name][price_name],
            bins=bins,
            labels=labels
        ).astype(float)

        self.forecasting_coefficients[Headers.BUCKETED_PRICE_CONTROL_COEFFICIENT] = (
            pd.DataFrame(
                {Headers.BUCKETED_PRICE_CONTROL_COEFFICIENT: new_col},
                index=self.forecasting_coefficients[price_name].index
            )
        )


    def __set_manual_price_control_coefficients(
            self,
            selected_countries: list[str],
            df: pd.DataFrame
            ) -> None:
        """
        'Manual price control coefficients'
        """
        manual_prices = df.copy()
        manual_prices = (
            manual_prices.reindex(manual_prices.index.repeat(len(selected_countries)))
        )
        manual_prices.index = pd.MultiIndex.from_tuples(
            [
                (country, sector, fuel)
                for (sector, fuel) in manual_prices.index[::len(selected_countries)]
                for country in selected_countries
            ],
            names=c.ID_COL_NAMES
        )

        manual_prices = manual_prices.sort_index()

        self.forecasting_coefficients[Headers.MANUAL_PRICE_CONTROL_COEFFICIENTS] = (
            manual_prices.copy()
        )


    def __set_chosen_price_control_coefficient(
            self,
            gov_energy_price_controls: Literal['Bucketed', 'Manual', 'None']
            ) -> None:
        """
        'Chosen price control coefficient'
        """
        price_name = Headers.CHOSEN_PRICE_CONTROL_COEFFICIENT
        if gov_energy_price_controls == 'Bucketed':
            self.forecasting_coefficients[price_name] = (
                self.forecasting_coefficients[Headers.BUCKETED_PRICE_CONTROL_COEFFICIENT]
                .rename(columns={
                    Headers.BUCKETED_PRICE_CONTROL_COEFFICIENT: price_name
                })
            )
        elif gov_energy_price_controls == 'Manual':
            self.forecasting_coefficients[price_name] = (
                self.forecasting_coefficients[Headers.MANUAL_PRICE_CONTROL_COEFFICIENTS]
                .rename(columns={
                    Headers.MANUAL_PRICE_CONTROL_COEFFICIENTS: price_name
                })
            )
        elif gov_energy_price_controls == 'None':
            chosen_price_control_coefficient = (
                self.forecasting_coefficients[Headers.BUCKETED_PRICE_CONTROL_COEFFICIENT]
                .rename(columns={
                    Headers.BUCKETED_PRICE_CONTROL_COEFFICIENT: price_name
                })
            )
            chosen_price_control_coefficient.loc[:, :] = 1.0
            self.forecasting_coefficients[price_name] = chosen_price_control_coefficient
        else:
            raise TypeError("Wrong value for gov_energy_price_controls.")


    def __get_keep_col_condiotion(self, code: str | list[str], col: str) -> bool:
        """
        Helper function
        """
        if isinstance(code, str):
            return col.startswith(code)
        if isinstance(code, list):
            return any(col.startswith(c) for c in code)
        raise TypeError("code must be a string or list of strings, in DomPrices")


    def __assert_indexes(self, selected_countries: list[str]) -> None:
        """
        Helper to assert that all indexes are complete
        TODO: test
        """
        countries_sorted = sorted(selected_countries)

        price_idx_pairs = [
            *[(sector, fuel) for sector in [c.IND, c.POW, c.RES] for fuel in [c.COA, c.NGA]],
            (c.IND, c.ELE), (c.RES, c.ELE),
            *[(c.ALL, fuel) for fuel in [c.BIO, c.DIE, c.GSO, c.KER, c.LPG, c.OOP]]
        ]
        price_idx = pd.MultiIndex.from_tuples(
            [
                (country, sector, fuel)
                for country in countries_sorted
                for sector, fuel in price_idx_pairs
            ],
            names=c.ID_COL_NAMES
        ).sort_values()

        ps_idx = pd.MultiIndex.from_product(
            [countries_sorted, [c.ALL], [c.BIO, c.COA, c.ELE, c.NGA, c.OIL]],
            names=c.ID_COL_NAMES
        )

        margin_idx = pd.MultiIndex.from_tuples(
            (country, sector, fuel)
            for country in countries_sorted
            for sector, fuel in [
                *[(c.ALL, fuel) for fuel in [c.BIO, c.DIE, c.GSO, c.KER, c.LPG, c.OOP]],
                *[(sector, fuel) for sector in [c.IND, c.POW, c.RES] for fuel in [c.COA, c.NGA]]
            ]
        )

        coa_nga_idx = pd.MultiIndex.from_product(
            [countries_sorted, [c.IND, c.POW, c.RES], [c.COA, c.NGA]],
            names=c.ID_COL_NAMES
        )

        assert self.rp.index.equals(price_idx)
        assert self.sp.index.equals(price_idx)
        assert self.txo.index.equals(price_idx)
        assert self.vatrate.index.equals(price_idx)
        assert self.ps.index.equals(ps_idx)

        assert (
            self.forecasting_coefficients[
                Headers.MARGIN_ON_TOP_OF_SUPPLY_COST
            ].index.equals(margin_idx)
        )
        assert ( # different order
            set(
                self.forecasting_coefficients[Headers.RAW_PRICE_CONTROL_COEFFICIENT].index
            ) == set(price_idx)
        )
        assert (
            self.forecasting_coefficients[
                Headers.DOMESTIC_PRODUCTION_COST
            ].index.equals(coa_nga_idx)
        )
        assert ( # different order
            set(
                self.forecasting_coefficients[Headers.BUCKETED_PRICE_CONTROL_COEFFICIENT].index
            ) == set(price_idx)
        )
        assert ( # different order
            set(
                self.forecasting_coefficients[Headers.MANUAL_PRICE_CONTROL_COEFFICIENTS].index
            ) == set(price_idx)
        )
        assert ( # different order
            set(
                self.forecasting_coefficients[Headers.CHOSEN_PRICE_CONTROL_COEFFICIENT].index
            ) == set(price_idx)
        )


def load_prices_dom(
        selected_countries: list[str]
        ) -> pd.DataFrame:
    """
    [Data Loading]
    Data from 'Prices_dom' tab (v361).

    Raw data 'prices_dom.csv' was initially preprocessed to a lighter form
    using 'get_prices_dom_csv' from `cpat_processing/prices_dom_format.py`.

    'prices_dom' used in DomPrices class for:
    'Historical prices' table from v361 Mitigation!B912:L921
    and
    'Energy price forecasting coefficients' table from v361 Mitigation!B714:R728.
    """
    # TODO pkl
    prices_raw: pd.DataFrame = (
        pd.read_csv(f'{DATA_PATH}/{PRICES_DATA_FILE_NAME}.csv')
        .drop(columns=['countryname'])
    )
    prices_raw = prices_raw[prices_raw[c.COUNTRY_CODE].isin(selected_countries)]
    prices_raw['year'] = prices_raw['year'].astype(int)

    return prices_raw
