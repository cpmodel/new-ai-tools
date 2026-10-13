from typing import Literal, TYPE_CHECKING

import pandas as pd

import cpat_model.constants as c
from config import DATA_PATH

from cpat_model.inputs.dashboard_inputs import DashboardInputsDict

if TYPE_CHECKING:
    from cpat_model.inputs.input_data import InputData

PRICES_DATA_FILE_NAME = 'prices_int'
REG_ASSUMPTIONS_FILE_NAME = 'int_prices_reg_assumptions'
GLOBAL_DEMAND_FILE_NAME = 'global_fuel_demand'

# db sources:
IMF = 'IMF'
WB = 'WB'
IEA = 'IEA'
IMF_WB = 'IMF-WB'
# not supported yet TODO:
EIA = 'EIA'
AVG = 'AVG'
MANUAL = 'Manual'
IMF_IEA = 'IMF-IEA'

# db years:
FIRST_DB_YEAR = 2018
LAST_DB_YEARS = {
    IMF: 2030,
    WB: 2026,
    IEA: 2040
}

# 'Uranium Fuel Cost' row 689 (v407)
URANIUM_FUEL_COST = 0.004

# define types:
PriceScenarioType = Literal['IMF', 'WB', 'IEA', 'IMF-WB']
PriceAdjustmentType = Literal['High', 'Base', 'Low']
GlobalEnergyDemandScenarioType = Literal["Stated Policies", "Announced Pledges", "Net Zero"]

CONV_TO_GJ = {
    # 'Conversion to GJ' CPAT Excel G696:698 (v361)
    c.OIL: 6.12,
    c.COA: 6 * 4.19,
    c.NGA: 1.055056
}

class IntPrices:
    """
    TODO:
    """
    prices: pd.DataFrame
    uranium_fuel_cost: float
    global_fuel_demand: pd.DataFrame
    __inv_deflator: pd.DataFrame
    __inv_deflator_growth: pd.DataFrame

    def __init__(
            self,
            simulation_years: tuple[int, int],
            selected_countries: list[str],
            input_data: 'InputData',
            d: DashboardInputsDict, # TODO: check tests
            deflator: pd.DataFrame
            ) -> None:
        """
        TODO: docstring
        TODO: add 'EIA', 'AVG', 'Manual', 'IMF-IEA' for global_energy_demand_scenario

        prices dims (c), t in <simulation_years[0] - 1, simulation_years[1]>
        """
        self.__set_uranium_fuel_cost()
        self.__set_global_fuel_demand(
            input_data.global_fuel_demand, d['global_energy_demand_scenario']
        )
        # # TODO test for __set_int_prices_from_source and integration test
        self.__set_inv_deflator_growth(deflator)

        self.__fist_other_source_year = 2024 # CPAT Excel: E680 (v407)

        if d['price_scenario_used'] in [IMF, WB, IEA]:
            self.prices = self.__set_int_prices_from_source(
                input_data.international_prices, selected_countries,
                simulation_years, d['price_scenario_used']
            )
        elif d['price_scenario_used'] == IMF_WB:
            self.prices = (
                self.__set_int_prices_from_source(
                    input_data.international_prices, selected_countries,
                    simulation_years, IMF
                )
                + self.__set_int_prices_from_source(
                    input_data.international_prices, selected_countries,
                    simulation_years, WB
                )
            ) / 2 # avg from both

        elif d['price_scenario_used'] in [EIA, AVG, MANUAL, IMF_IEA]:
            raise NotImplementedError("Sources: EIA, AVG, Manual, IMF-IEA not yet supported.")
        else:
            raise ValueError(
                f"Wrong international price source provided: {d['price_scenario_used']}."
            )

        # CPAT Excel 683:685 (v407) (inputs included)
        self.__apply_regional_assumtions(
            input_data.int_prices_regional_assumptions, simulation_years,
            d['int_energy_price_adjustment']
        )

        # CPAT Excel 686:688 (v407)
        self.__apply_deflator(deflator)
        # CPAT Excel 710:712 (v407)
        self.__conv_to_gj()


    def __set_uranium_fuel_cost(self) -> None:
        """
        'Uranium Fuel Cost'
        CPAT Excel: row 689 (407)

        Unit: $/kwhe of final electricity

        Set as float as it is hardcoded in Excel as 1 value for all years
        """
        self.uranium_fuel_cost = URANIUM_FUEL_COST

    def __set_global_fuel_demand(
            self,
            global_fuel_demand: pd.DataFrame,
            global_energy_demand_scenario: GlobalEnergyDemandScenarioType
            ) -> None:
        """
        'Global demand'
        CPAT Excel 696:698, also baseline 6134:6136, scenario 10905:10907 (v407)
        BUG in excel 694: no 'Sustainable Development' in data -> selector bug TODO:

        return dims (f), <simulation_years[0] - 1, simulation_years[1]>
            """
        scenario_col = 'Scenario'
        self.global_fuel_demand = (
            global_fuel_demand[
                global_fuel_demand[scenario_col] == (
                    global_energy_demand_scenario + ' ' + scenario_col
                )
            ]
            .drop(columns=[scenario_col])
        )


    def __set_inv_deflator_growth(
            self,
            deflator: pd.DataFrame
            ) -> None:
        """
        CPAT Excel: rows 609 and 610 (v407)
        """
        # '1/denominator', CPAT Excel: row 609 (v407)
        self.__inv_deflator = 1 / deflator
        # growth,  CPAT Excel: row 610 (v407)
        self.__inv_deflator_growth = (
            self.__inv_deflator.iloc[:, 1:]
            .div(self.__inv_deflator.iloc[:, :-1].values)
            .sub(1)
        )

    def __set_int_prices_from_source(
            self,
            international_prices: pd.DataFrame,
            selected_countries: list[str],
            simulation_years: tuple[int, int],
            price_scenario_used: Literal['IMF', 'WB', 'IEA'] # no avg scenarios
            ) -> pd.DataFrame:
        """
        Returns prices form the selected price_scenario_used.
        'IMF' CPAT Excel: rows 613:618 (v407)
        'WB' CPAT Excel: rows 621:626 (v407)
        'IEA' CPAT Excel: rows 629:634 (v407)
        """
        prices = international_prices[
            international_prices.index.get_level_values('Source') == price_scenario_used
        ].droplevel('Source')
        prices = self.__extend_by_country(prices, selected_countries)
        prices = self.__forecast_years(prices, price_scenario_used, simulation_years)
        prices = self.__cut_years(prices, simulation_years)
        prices = self.__add_nga_global(prices)
        if price_scenario_used == IEA:
            prices = self.__forecast_early_coa(prices, simulation_years)
        return prices


    def __forecast_early_coa(
            self,
            prices: pd.DataFrame,
            simulation_years: tuple[int, int],
            ) -> pd.DataFrame:
        """
        Forecasts IEA coa values from 2031
        CPAT Excel row 630 (v407)
        """
        coa_forecast_first_year = 2031
        if coa_forecast_first_year <= simulation_years[1]:
            idx = pd.IndexSlice
            for year in range(coa_forecast_first_year, simulation_years[1] + 1):
                y = str(year)
                prices.loc[idx[:, c.COA, :], y] = (
                    prices.loc[idx[:, c.COA, :], str(year - 1)]
                    .mul(1 + self.__inv_deflator_growth[y], level=c.COUNTRY_CODE)
            )
        return prices


    def __extend_by_country(
            self,
            price: pd.DataFrame,
            selected_countries: list[str]
            ) -> pd.DataFrame:
        """
        Extends index by selected_countries.
        """
        price = price.reindex(price.index.repeat(len(selected_countries)))
        price.index = pd.MultiIndex.from_tuples(
            [
                (country, fuel, region)
                for (fuel, region) in price.index[::len(selected_countries)]
                for country in selected_countries
            ],
            names=[c.COUNTRY_CODE, c.FUEL_CODE, 'Region']
        )
        return price.sort_index()

    def __forecast_years(
            self,
            prices: pd.DataFrame,
            price_scenario_used: PriceScenarioType,
            simulation_years: tuple[int, int]
            ) -> pd.DataFrame:
        """
        Forecasts years past LAST_DB_YEARS[price_scenario_used].
        """
        for year in range(LAST_DB_YEARS[price_scenario_used] + 1, simulation_years[1] + 1):
            y = str(year)
            prices[y] = (
                prices[str(year - 1)]
                * (1 + self.__inv_deflator_growth[y].reindex(
                    prices.index.get_level_values(c.COUNTRY_CODE)).values
                )
            )
        return prices

    def __cut_years( #TODO
            self,
            prices: pd.DataFrame,
            simulation_years: tuple[int, int]
            ) -> pd.DataFrame:
        """
        In case of simulation_years[1] < LAST_DB_YEARS[IEA],
        drops columns where year>simulation_years[1].
        """
        # only IEA data has values until 2040, and determines db last year
        if simulation_years[1] < LAST_DB_YEARS[IEA]:
            return prices.loc[:, prices.columns.astype(int) <= simulation_years[1]]
        return prices

    def __add_nga_global(self, prices: pd.DataFrame) -> pd.DataFrame:
        """
        Calculates NGA Globals as avg from 'EU', 'LNG' and 'US'.
        """
        new_rows = []
        for country_code in prices.index.get_level_values(c.COUNTRY_CODE).unique():
            nga_country = prices.loc[(country_code, c.NGA), :]
            nga_avg = nga_country.sum(axis=0) / 3 # 3 sources of nga price
            new_rows.append(
                pd.DataFrame(
                    [nga_avg.values],
                    index=pd.MultiIndex.from_tuples(
                        [(country_code, c.NGA, 'Global')],
                        names=prices.index.names
                    ),
                    columns=prices.columns
                )
            )
        prices = pd.concat([prices] + new_rows).sort_index()
        return prices

    def __apply_regional_assumtions(
            self,
            int_prices_regional_assumptions: pd.DataFrame,
            simulation_years: tuple[int, int],
            int_energy_price_adjustment: PriceAdjustmentType
            ) -> None:
        """
        'Global prices' in nom (v407)
        CPAT Excel 683:685 (v407) (inputs included)
        """
        # choose nga source
        eu_shift_to_lng = self.__get_eu_shift_to_lng(
            int_prices_regional_assumptions, simulation_years
        )

        nga_assumptions_col = 'Natural gas market assumption'
        new_rows = []
        for country_code in self.prices.index.get_level_values(c.COUNTRY_CODE).unique():
            country_assumtion = (
                int_prices_regional_assumptions.loc[country_code, nga_assumptions_col]
            )
            # if EU -> apply shift values, select row with the correct assumption otherwise
            if country_assumtion == 'EU':
                shift_vals = eu_shift_to_lng.iloc[0, :]
                new_row = (
                    self.prices.loc[(country_code, c.NGA, 'EU')] * (1.0 - shift_vals)
                    + self.prices.loc[(country_code, c.NGA, 'LNG')] * shift_vals
                )
                new_rows.append(
                    pd.DataFrame(
                        [new_row],
                        index=pd.MultiIndex.from_tuples(
                            [(country_code, c.NGA, None)],
                            names=self.prices.index.names
                        ),
                        columns=self.prices.columns
                    )
                )
            else:
                new_row = self.prices.loc[(country_code, c.NGA, country_assumtion)]
                new_rows.append(
                    pd.DataFrame(
                        [new_row],
                        index=pd.MultiIndex.from_tuples(
                            [(country_code, c.NGA, None)],
                            names=self.prices.index.names
                        ),
                        columns=self.prices.columns
                    )
                )
        self.prices = pd.concat([self.prices] + new_rows).sort_index()

        self.prices = self.prices[self.prices.index.get_level_values('Region').isna()]
        self.prices = self.prices.droplevel('Region')

        # if int_energy_price_adjustment == 'Base' we do nothing
        if int_energy_price_adjustment != 'Base':
            self.__apply_fuel_adjustment(int_energy_price_adjustment, simulation_years)

    def __get_eu_shift_to_lng(
            self,
            int_prices_regional_assumptions: pd.DataFrame,
            simulation_years: tuple[int, int]
            ) -> pd.DataFrame:
        """
        'Europe shift to LNG adjustment factor (% of gas from outside Russia)' (v407)
        CPAT Excel: row 678

        Minimun for simulation_years[0] is 2022 - curve_value start at 2022

        returns dims (c), t in <simulation_years[0] - 1, simulation_years[1]>
        """
        # values hardcoded in Excel:
        curve_value = {2022: 0.58, 2023: 0.603, 2024: 0.734, 2025: 0.834}
        nga_assumptions_col = 'Natural gas market assumption'
        # last year for hardcoded calculations:
        last_calc_year = max(curve_value)
        eu_shift_to_lng = int_prices_regional_assumptions.copy()
        eu_shift_to_lng = eu_shift_to_lng[[nga_assumptions_col]]

        # add 0s in simulation_years[0] - 1 for easier calc later on
        eu_shift_to_lng[str(simulation_years[0] - 1)] = 0.0
        # up to 2025 calculate values:
        for year in range(simulation_years[0], last_calc_year + 1):
            y = str(year)
            eu_shift_to_lng[y] = (
                (eu_shift_to_lng[nga_assumptions_col] == 'EU').astype(float)
                * curve_value[year]
            )

        # rest of the years -> last year values
        for year in range(last_calc_year + 1, simulation_years[1] + 1):
            eu_shift_to_lng[str(year)] = eu_shift_to_lng[str(year - 1)]
        eu_shift_to_lng.drop(columns=[nga_assumptions_col], inplace=True)
        # nga_region -> F83 -> Prices_int!$I$370:$I$589

        return eu_shift_to_lng

    def __apply_fuel_adjustment(
            self,
            int_energy_price_adjustment: PriceAdjustmentType,
            simulation_years: tuple[int, int]
            ) -> None:
        """
        'Global prices' in nom (v407)
        CPAT Excel: rows 683:685

        Only applies adjustment for 'Oil' and 'Gas & coal'. (last part of equation)

        prices dims (c), t in <simulation_years[0] - 1, simulation_years[1]>
        """
        # CPAT Excel: E668
        oil_values = {'Low': 0.5, 'High': 1.5}
        # CPAT Excel: E669
        nga_coa_values = {'Low': 0.75, 'High': 1.25}

        # only for years
        cols_to_modify = [
            str(year) for year in range(self.__fist_other_source_year + 1, simulation_years[1] + 1)
        ]

        idx = pd.IndexSlice
        self.prices.loc[idx[:, c.OIL], cols_to_modify] *= oil_values[int_energy_price_adjustment]
        self.prices.loc[
            idx[:, [c.NGA, c.COA]], cols_to_modify
        ] *= nga_coa_values[int_energy_price_adjustment]


    def __apply_deflator(self, deflator:pd.DataFrame) -> None:
        """
        'Global prices' in real {results_year} (v407)
        CPAT Excel: rows 686:688

        prices dims (c), t in <simulation_years[0] - 1, simulation_years[1]>
        """
        self.prices = self.prices.mul(deflator, level=c.COUNTRY_CODE)

    def __conv_to_gj(self) -> None:
        """
        'Global prices' in '$/GJ real' (v407)
        CPAT Excel: rows 710:712

        prices dims (c), t in <simulation_years[0] - 1, simulation_years[1]>
        """
        idx = pd.IndexSlice
        for fuel, value in CONV_TO_GJ.items():
            self.prices.loc[idx[:, fuel], :] /= value

def load_international_prices(
        simulation_years: tuple[int, int]
        ) -> pd.DataFrame:
    """
    [Data Loading]
    Raw data loading, used later for:
    'International energy price and demand forecasts' tables from rows 602:664 (v361)

    'prices_int.csv' file:
    - based on 'Prices_int' tab (v361)
    - colums added manually: 'Source', 'FuelCode', 'Region', 'Unit', are followed by year columns
    - year columns used: 2018-last available year
    - 'IMF' data Prices_int! rows 165:173
    - 'WB' data Prices_int! rows 11:17
    - 'IEA' data Prices_int! rows 273:278 298:303

    return dims (f), t in <simulation_years[0] -1, last db year> 
    """
    if FIRST_DB_YEAR < simulation_years[0]:
        # keep 1 year back:
        years_to_drop = [str(y) for y in range(FIRST_DB_YEAR, simulation_years[0] - 1)]
    prices_raw: pd.DataFrame = (
        # TODO pkl
        pd.read_csv(f'{DATA_PATH}/{PRICES_DATA_FILE_NAME}.csv')
        .drop(columns=['Unit', *years_to_drop])
        .set_index(['Source', c.FUEL_CODE, 'Region'])
    )
    return prices_raw

def load_int_prices_regional_assumptions(
        selected_countries: list[str]
        ) -> pd.DataFrame:
    """
    [Data Loading]
    Raw data loading:
    1st table from 'Regional assumptions' section in Prices_int tab rows 415:635 (v407)

    'int_prices_reg_assumptions.csv' file:
    - based on 'Prices_int' tab (v407), 'Regional assumptions'
    - processed by 'get_int_prices_reg_assumptions_csv' script
        from 'cpat_processing/int_prices_reg_assumptions_format.py'

    return dims (c),
        'Baseline taxes are ad-valorem or fixed?' and 'Natural gas market assumption' columns
    """
    # TODO pkl
    df: pd.DataFrame = (
        pd.read_csv(f'{DATA_PATH}/{REG_ASSUMPTIONS_FILE_NAME}.csv')
        .set_index(c.COUNTRY_CODE)
    )
    df = df[df.index.get_level_values(c.COUNTRY_CODE).isin(selected_countries)]
    return df

def load_global_fuel_demand(
        simulation_years: tuple[int, int]
        ) -> pd.DataFrame:
    """
    [Data Loading]
    Loads data 'Global fuel demand' from Pices_int tab rows 676:687 (v407)

    Manual adjustments:
    - columns renamed: 'Fuel' -> 'FuelCode'
    - columns removed: 'Indicator', empty year columns: '2018', '2019'
    - values in 'FuelCode' changed according to {'Coal': 'coa', 'Oil': 'oil', 'Natural gas': 'nga'}

    Data years 2020 - 2050

    return dims (f), 'Scenario', <simulation_years[0] - 1, simulation_years[1]>
    """
    # TODO: pkl
    df: pd.DataFrame = pd.read_csv(f'{DATA_PATH}/{GLOBAL_DEMAND_FILE_NAME}.csv')
    df = df.drop(columns=['Unit']).set_index(c.FUEL_CODE)

    year_columns = [str(year) for year in range(simulation_years[0] - 1, simulation_years[1] + 1)]
    return df.loc[:, ['Scenario'] + year_columns]
