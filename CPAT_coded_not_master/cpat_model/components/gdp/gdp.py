from typing import Literal, TypeAlias, TYPE_CHECKING

import pandas as pd

import cpat_model.constants as c
from config import DATA_PATH

from cpat_model.inputs.dashboard_inputs import DashboardInputsDict
if TYPE_CHECKING:
    from cpat_model.inputs.input_data import InputData


GDPSource: TypeAlias = Literal['wdi', 'weo', 'un_pop']

# 'CPAT forecast starts from' D379 (v361):
LAST_WEO_WDI_YEAR = 2029
# WEO NaNs:
# - LE for many countries
# - NGDP_R_PPP_PC for MSR, SYR

# WDI NaNs:
# - countries with many NaNs: PRK, GIB, VGB
# - countries with NaNs jsut in enda: CUB, PSE

class GDP:
    """
    TODO
    TODO: finish moving from functions
    TODO: filter year earlier?
    """
    population: pd.DataFrame
    population_growth: pd.DataFrame

    gdp_per_capita: pd.DataFrame
    gdp_growth_per_capita: pd.DataFrame

    ngdp: pd.DataFrame
    enda: pd.DataFrame

    ngdp_rpch: pd.DataFrame
    d_gdp_at_const_prices: pd.DataFrame

    ngdp_d: pd.DataFrame
    ngdp_d_world: pd.DataFrame
    deflator: pd.DataFrame

    inflation_index: pd.DataFrame
    inflation_index_100: pd.DataFrame

    def __init__(
            self,
            simulation_years: tuple[int, int],
            d: DashboardInputsDict,
            input_data: 'InputData'
            ) -> None:
        """
        TODO
        """
        self.__set_population(input_data.gdp_data_country_filtered, simulation_years)
        self.__set_population_growth()
        self.__set_inflation_index(input_data.pcpi_world, simulation_years)
        self.__set_inflation_index_100(d['results_year'])

        self.__set_ngdp_d(input_data.gdp_data_country_filtered, input_data.gdp_countries_in_data)
        self.__project_ngdp_d(simulation_years)
        self.__set_ngdp_d_world(input_data.ngdp_d_world, simulation_years)
        self.__set_deflator(d['results_year'], d['deflator_selection'])

        self.__init_ngdp_rpch(
            input_data.gdp_data_country_filtered, input_data.gdp_countries_in_data
        )
        self.__project_ngdp_rpch(simulation_years)
        self.__set_d_gdp_at_const_prices(d['gdp_adj'])

    def __set_population(
            self,
            gdp_data_country_filtered: dict[str, pd.DataFrame],
            simulation_years: tuple[int, int]
            ) -> None:
        """
        'Population' (v361)
        CPAT Excel: 394

        'Popn' not added as 'UNPop' covers all of the countries in 'Popn'
        If we add {'WAV', 'WMR'}, use WDI (these are not present in UNPop)

        returns dims (c), all t TODO: years
        """
        start_year = simulation_years[0]
        end_year = simulation_years[1]

        self.population = (
            gdp_data_country_filtered['un_pop'].loc[:, str(start_year - 1):str(end_year)]
        ).copy()


    def __set_population_growth(self) -> None:
        """
        'Population' (v361)
        CPAT Excel: 395

        returns dims (c), all t TODO: years
        """
        years = sorted(int(col) for col in self.population.columns)
        self.population_growth = pd.DataFrame({
            str(year): (self.population[str(year)] / self.population[str(year - 1)] - 1)
            for year in years[1:] # skip start_year - 1
        })


    def __set_inflation_index(
            self,
            pcpi_world: pd.DataFrame,
            simulation_years: tuple[int, int]
            ) -> None:
        """
        'Inflation index' (v361)
        CPAT Excel: 424

        PCPI for USA.

        returns dims t in <simulation_years[0]-2, simulation_years[1]>
        """
        start_year = simulation_years[0]
        end_year = simulation_years[1]
        last_weo_year = max(map(int, pcpi_world.columns))

        inflation_index = pcpi_world.copy()

        if end_year < last_weo_year:
            inflation_index = (
                inflation_index.loc[:, str(start_year - 2):str(end_year)]
            )
        else:
            inflation_index = (
                inflation_index.loc[:, str(start_year - 2):str(last_weo_year)]
            )
            for year in range(last_weo_year + 1, end_year + 1):
                inflation_index[str(year)] = inflation_index[str(year - 1)]

        self.inflation_index = inflation_index.reset_index(drop=True)


    def __set_inflation_index_100(
            self,
            results_year: int
            ) -> None:
        """
        'Inflation index, {year} = 100' (v361)
        CPAT Excel: 425

        returns dims t in <simulation_years[0]-1, simulation_years[1]>
        """
        # drop 1st column
        inflation_index_100 = self.inflation_index.drop(self.inflation_index.columns[0], axis=1)
        self.inflation_index_100 = (
            self.inflation_index.loc[0, str(results_year)] / inflation_index_100
        )


    def __set_ngdp_d(
            self,
            gdp_data_country_filtered: dict[GDPSource, pd.DataFrame],
            gdp_countries_in_data: dict[GDPSource, set[str]]
            ) -> None:
        """
        'GDP Deflator' (v361) NGDP_D (For selected countries)
        CPAT Excel: 383

        Load years from DB, rest of the years are forecasted in __set_ngdp_d_projection

        returns dims (c), all t in <first DB, last DB year>
        """
        # TODO test
        # if country in 'WEO': ngdp_d from 'WEO'
        self.ngdp_d = pd.DataFrame()
        if 'weo' in gdp_countries_in_data.keys():
            weo = gdp_data_country_filtered['weo']
            self.ngdp_d = (
                weo[
                    weo.index.get_level_values('Indicator').isin(['NGDP_D'])
                ]
                .droplevel('Indicator')
            )
        if 'wdi' in gdp_countries_in_data.keys():
            wdi = gdp_data_country_filtered['wdi']
            wdi_ngdp_d = (
                wdi[
                    wdi.index.get_level_values('code').isin(['ngdp_d'])
                ]
                .droplevel('code')
            )

            # if no countries from 'WEO'
            if self.ngdp_d.empty:
                self.ngdp_d = wdi_ngdp_d.copy().drop_duplicates()
            else:
                self.ngdp_d = pd.concat([self.ngdp_d, wdi_ngdp_d]).drop_duplicates()


    def __project_ngdp_d(
            self,
            simulation_years: tuple[int, int]
            ) -> None:
        """
        'GDP Deflator' (v361) NGDP_D (For selected countries)
        CPAT Excel: 383

        Forecast for years not in the DB
        and removes years prior to simulation_years[0]-2

        returns dims (c), t in <simulation_years[0]-2, simulation_years[1]>
        """
        first_gdp_projection_year = max(int(col) for col in self.ngdp_d.columns) + 1

        self.ngdp_d = self.ngdp_d[[
            col for col in self.ngdp_d.columns
            if not col.isdigit() or int(col) >= simulation_years[0] - 2
        ]]

        for year in range(first_gdp_projection_year, simulation_years[1] + 1):
            self.ngdp_d[str(year)] = (
                2 * self.ngdp_d.loc[:, str(year - 1)]
                - self.ngdp_d.loc[:, str(year - 2)]
            )


    def __set_ngdp_d_world(
            self,
            ngdp_d_world: pd.DataFrame,
            simulation_years: tuple[int, int]
            ) -> None:
        """
        'GDP Deflator (2012=100)' (v361) NGDP_D (World - USA)
        CPAT Excel: 421

        returns dims (c = 'USA'), t in <simulation_years[0]-2, simulation_years[1]>
        """
        self.ngdp_d_world = ngdp_d_world.copy()
        first_gdp_projection_year = max(int(col) for col in self.ngdp_d_world.columns) + 1

        self.ngdp_d_world = self.ngdp_d_world[[
            col for col in self.ngdp_d_world.columns
            if not col.isdigit() or int(col) >= simulation_years[0] - 2
        ]]

        for year in range(first_gdp_projection_year, simulation_years[1] + 1):
            self.ngdp_d_world[str(year)] = (
                2 * self.ngdp_d_world.loc[:, str(year - 1)]
                - self.ngdp_d_world.loc[:, str(year - 2)]
            )

    def __set_deflator(
            self,
            results_year: int,
            deflator_selection: Literal['World', 'Country'] = 'World',
            ) -> pd.DataFrame:
        """
        'Deflator- {results_year} - used hereafter - World/Country' (v361)
        CPAT Excel: 436

        returns dims (c), t in <simulation_years[0] - 1, simulation_years[1]>
        """
        if deflator_selection == 'World':
            self.deflator = pd.DataFrame(
                {col: self.ngdp_d_world[str(results_year)] for col in self.ngdp_d_world.columns},
                index=self.ngdp_d_world.index
            )
            self.deflator /= self.ngdp_d_world

            # adjusting dims for country specific (still 'World' deflator for each country)
            self.deflator = pd.DataFrame(
                [self.deflator.iloc[0].values] * len(self.ngdp_d), # TODO: what if c not in WEO
                columns=self.deflator.columns
            )
            self.deflator.index = self.ngdp_d.index

        if deflator_selection == 'Country':
            self.deflator = pd.DataFrame(
                {col: self.ngdp_d[str(results_year)] for col in self.ngdp_d.columns},
                index=self.ngdp_d.index
            )
            self.deflator /= self.ngdp_d

        # drop simulation_years[0] - 2
        self.deflator = self.deflator.iloc[:, 1:]

    def __init_ngdp_rpch(
            self,
            gdp_data_country_filtered: dict[GDPSource, pd.DataFrame],
            countries_in_data: dict[GDPSource, set[str]]
            ) -> pd.DataFrame:
        """
        'GDP growth - real (constant prices) % change' (v361) NGDP_RPCH
        CPAT Excel: 386

        Load years from DB, rest of the years are forecasted in get_ngdp_rpch_projection

        ngdp_rpch dims (c), t in <first DB, last DB year>
        """
        # if country in 'WEO': ngdp and enda from 'WEO'
        ngdp_rpch = pd.DataFrame()
        if 'weo' in countries_in_data.keys():
            weo = gdp_data_country_filtered['weo']
            ngdp_rpch = (
                weo[
                    weo.index.get_level_values('Indicator').isin(['NGDP_RPCH'])
                ]
                .droplevel('Indicator')
            )
        if 'wdi' in countries_in_data.keys():
            wdi = gdp_data_country_filtered['wdi']
            wdi_ngdp_rpch = (
                wdi[
                    wdi.index.get_level_values('code').isin(['ngdp_rpch'])
                ]
                .droplevel('code')
            )

            # if no countries from 'WEO'
            if ngdp_rpch.empty:
                ngdp_rpch = wdi_ngdp_rpch.copy()
            else:
                ngdp_rpch = pd.concat([ngdp_rpch, wdi_ngdp_rpch])

        ngdp_rpch = ngdp_rpch / 100 # adjust % to decimal

        self.ngdp_rpch = ngdp_rpch.copy()


    def __project_ngdp_rpch(
            self,
            simulation_years: tuple[int, int]
            ) -> pd.DataFrame:
        """
        'GDP growth - real (constant prices) % change' (v361) NGDP_RPCH
        CPAT Excel: 386

        Forecast for years not in the DB

        returns dims (c), t in <simulation_years[0] - 2, simulation_years[1]>
        """
        last_db_gdp_year = max(int(col) for col in self.ngdp_rpch.columns)
        first_db_gdp_year = min(int(col) for col in self.ngdp_rpch.columns)
        first_year_to_keep = simulation_years[0] - 2

        first_gdp_projection_year = last_db_gdp_year + 1
        last_gdp_projection_year = simulation_years[1]
        columns_to_keep = [
            str(y) for y in range(first_year_to_keep, last_gdp_projection_year + 1)
        ]

        # if last_db_gdp_year > last_gdp_projection_year:
        #     self.ngdp_rpch = self.ngdp_rpch[columns_to_keep]

        if last_db_gdp_year < last_gdp_projection_year:
            # add new years with 0.0
            for year in range(first_gdp_projection_year, last_gdp_projection_year + 1):
                self.ngdp_rpch[str(year)] = 0.0

            for country in self.ngdp_rpch.index:
                value_at_t0_minus_2 = self.ngdp_rpch.at[country, str(first_gdp_projection_year - 2)]
                value_at_t0_minus_3 = self.ngdp_rpch.at[country, str(first_gdp_projection_year - 3)]
                if value_at_t0_minus_2 < value_at_t0_minus_3:
                    for year in range(first_gdp_projection_year, last_gdp_projection_year + 1):
                        avg_value = (
                            value_at_t0_minus_2
                            + value_at_t0_minus_3
                            + self.ngdp_rpch.at[country, str(first_gdp_projection_year - 4)]
                        ) / 3 / 2
                        trend_value = (
                            2 * self.ngdp_rpch.at[country, str(year - 1)]
                            - self.ngdp_rpch.at[country, str(year - 2)]
                        )
                        self.ngdp_rpch.at[country, str(year)] = max(avg_value, trend_value)
                else:
                    self.ngdp_rpch.loc[
                        country,
                        str(first_gdp_projection_year):str(last_gdp_projection_year)
                    ] = value_at_t0_minus_2


        if first_db_gdp_year < first_year_to_keep:
            self.ngdp_rpch = self.ngdp_rpch[columns_to_keep]


    def __set_d_gdp_at_const_prices(
            self,
            gdp_adj: Literal['High', 'Base', 'Low']
            ) -> None:
        """
        'Change in country total real GDP at constant prices' (v361)
        CPAT Excel: 449

        Skips the option to input gdp manually.
        TODO: include manual gdp (in v361 row 432)

        returns dims (c), t in <simulation_years[0] - 2, simulation_years[1]>
        """
        # E253 'GDP growth'
        factor = {'High': 1.5, 'Base': 1.0, 'Low': 0.5}[gdp_adj]
        self.d_gdp_at_const_prices = self.ngdp_rpch * factor


def get_gdp_data_country_filtered(
        selected_countries: list[str],
        ) -> tuple[
            dict[GDPSource, pd.DataFrame],
            dict[GDPSource, set[str]],
            pd.DataFrame, 
            pd.DataFrame
        ]:
    """
    [Data Loading] TODO handle NaNs, missing countries?
    Performs country filtering for raw data taken from (v407) CPAT Excel tabs,
    with the following manual changes:
    - 'WEO' (columns: 'ISO' column renamed to 'CountryCode', 'Country',
        'Indicator', 'Scale', year columns 2017-2029; row 'Euro area' 'ENDA' dropped)
    - 'WDI' (rows 278-453; columns: 'Country code' renamed to 'CountryCode',
        'Country', 'code', all year columns: 2018-2029),
        WAV and WMR for ngdprpc and lp rows filled with values from v361
    - 'UNPop' ('ISO' column renamed to 'CountryCode';
        dropped columns: 'CPAT code', 'Population')

    returns tuple of 2 dicts and 1 pd.DataFrame
        2 dicts:
            1st with filtered data
            2nd with countries in filtered data
        pd.DataFrame with NGDP_D for USA for World deflator
        pd.DataFrame with PCPI for USA for Inflation index
    """
    data_raw: dict[str, pd.DataFrame] = {}
    data_filtered = {}
    countries_in_data = {}

    weo_indicators = ['NGDP', 'ENDA','NGDP_RPCH', 'NGDP_D', 'PCPI']
    wdi_indicators = ['ngdp', 'enda','ngdp_rpch', 'ngdp_d']

    for data_name, index in {
        'weo': [c.COUNTRY_CODE, 'Country', 'Indicator', 'Scale'],
        'wdi': [c.COUNTRY_CODE, 'Country', 'code'],
        'un_pop': [c.COUNTRY_CODE]
    }.items():
        data_raw[data_name] = (
            # pd.read_pickle(f'{DATA_PATH}/{data_name}.pkl.bz2', compression="bz2") # TODO pkl
            pd.read_csv(f'{DATA_PATH}/{data_name}.csv')
            .set_index(index)
        )

    # WEO or WDI
    weo_unique_countries = set(data_raw['weo'].index.get_level_values(c.COUNTRY_CODE).unique())
    wdi_unique_countries = set(data_raw['wdi'].index.get_level_values(c.COUNTRY_CODE).unique())
    selected_countries = set(selected_countries)

    countries_in_weo = weo_unique_countries & selected_countries
    countries_in_wdi = wdi_unique_countries & selected_countries

    countries_not_in_weo_wdi = selected_countries - weo_unique_countries - wdi_unique_countries
    assert (
        not countries_not_in_weo_wdi
    ), f"Error: Countries: {countries_not_in_weo_wdi} not in WEO or WDI data."

    # to silence linter
    ngdp_d_world = pd.DataFrame()
    pcpi_world = pd.DataFrame()

    if len(countries_in_weo) != 0:
        weo_country_filtered = (
            data_raw['weo'][
                data_raw['weo'].index.get_level_values(c.COUNTRY_CODE).isin(countries_in_weo)
                & data_raw['weo'].index.get_level_values('Indicator').isin(weo_indicators)
            ]
            .droplevel('Country')
            .droplevel('Scale')
            .drop_duplicates()
        )
        data_filtered['weo'] = weo_country_filtered
        countries_in_data['weo'] = countries_in_weo

        if 'USA' in countries_in_weo:
            weo_usa = weo_country_filtered[
                weo_country_filtered.index.get_level_values(c.COUNTRY_CODE).isin(['USA'])
            ]
            ngdp_d_world = (
                weo_usa[
                    weo_usa.index.get_level_values('Indicator').isin(['NGDP_D'])
                ]
                .droplevel('Indicator')
                .drop_duplicates()
            )
            pcpi_world = (
                weo_usa[
                    weo_usa.index.get_level_values('Indicator').isin(['PCPI'])
                ]
                .droplevel('Indicator')
                .drop_duplicates()
            )
    # we need to extract NGDP_D and PCPI for USA anyway
    if ngdp_d_world.empty:
        weo_usa = (
            data_raw['weo'][
                data_raw['weo'].index.get_level_values(c.COUNTRY_CODE).isin(['USA'])
            ]
            .droplevel('Country')
            .droplevel('Scale')
        )
        ngdp_d_world = (
            weo_usa[
                weo_usa.index.get_level_values('Indicator').isin(['NGDP_D'])
            ]
            .droplevel('Indicator')
            .drop_duplicates()
        )
        pcpi_world = (
            weo_usa[
                weo_usa.index.get_level_values('Indicator').isin(['PCPI'])
            ]
            .droplevel('Indicator')
            .drop_duplicates()
        )

    # get wdi data only if unavailable in weo
    if len(countries_in_weo) != len(selected_countries):
        wdi_country_filtered = (
            data_raw['wdi'][
                data_raw['wdi'].index.get_level_values(c.COUNTRY_CODE).isin(countries_in_wdi)
                & data_raw['wdi'].index.get_level_values('code').isin(wdi_indicators)
            ]
            .droplevel('Country')
        )
        data_filtered['wdi'] = wdi_country_filtered
        countries_in_data['wdi'] = countries_in_wdi

    # un_pop
    un_pop_unique_countries = (
        set(data_raw['un_pop'].index.get_level_values(c.COUNTRY_CODE).unique())
    )
    countries_in_un_pop = un_pop_unique_countries & selected_countries
    countries_not_un_pop = selected_countries - un_pop_unique_countries
    assert (
        not countries_not_un_pop
    ), f"Error: Countries: {countries_not_un_pop} not in UNPop."

    if len(countries_in_un_pop) != 0:
        un_pop_country_filtered = (
            data_raw['un_pop'][
                data_raw['un_pop'].index.get_level_values(c.COUNTRY_CODE).isin(countries_in_un_pop)
            ]
        ) / 1e3 # un_pop /1e3 according to CPAT Excel
        data_filtered['un_pop'] = un_pop_country_filtered.copy()
        countries_in_data['un_pop'] = countries_in_un_pop

    return data_filtered, countries_in_data, ngdp_d_world, pcpi_world # TODO gdp_countries_in_data
