import pandas as pd
import numpy as np

from cpat_model.components.power.data import PowData

import cpat_model.constants as c

class PowInvestment:
    """
    Simplified Power Investment

    Inlcudes tables from 'Retirement and Capacity Needed' - section K
    and calculations for simplified 'Proportions of New Investment' L3-L5 tables
    """
    nameplate_retirements: pd.DataFrame
    total_nameplate_retirements: pd.DataFrame
    effective_retirements: pd.DataFrame
    total_effective_retirements: pd.DataFrame

    nameplate_capacity: pd.DataFrame
    total_nameplate_capacity: pd.DataFrame
    effective_capacity: pd.DataFrame
    total_effective_capacity: pd.DataFrame

    new_investment_permitted: pd.Series
    proportions_of_investment: pd.DataFrame

    eff_investment: pd.DataFrame
    total_eff_investment: pd.DataFrame
    def __init__(
            self,
            simulation_years: tuple[int, int],
            selected_countries: list[str],
            data: PowData
        ) -> None:
        self.__base_year: int = simulation_years[0]

        # TODO: check assigning .loc[...] for dfs created with this method:
        self.nameplate_retirements = self.__get_empty_frame(simulation_years, data.pow_index)
        self.effective_retirements = self.__get_empty_frame(simulation_years, data.pow_index)
        self.__init_total_effective_retirements()

        self.__init_nameplate_capacity(simulation_years, data.capacity)
        self.__init_total_nameplate_capacity()
        self.__init_effective_capacity(simulation_years, data.elec_for_power)
        self.__init_total_effective_capacity()

        self.__set_new_investment_permitted()
        self.proportions_of_investment = self.__get_empty_frame(simulation_years, data.pow_index)

        self.eff_investment = self.__get_empty_frame(simulation_years, data.pow_index)
        self.__init_total_eff_investment(simulation_years, selected_countries)

    def __get_empty_frame(
            self,
            simulation_years: tuple[int, int],
            pow_index: pd.MultiIndex,
            ) -> pd.DataFrame:
        """
        Helper function

        Creates df filled with 0.0s.
        TODO unify with the same method in variable_cost file

        dims (c, f), t in <simulation_years[0], simulation_years[1]>
            f in POW_FUELS
        """
        return pd.DataFrame(
            0.0,
            index=pow_index,
            columns=[str(y) for y in range(simulation_years[0], simulation_years[1] + 1)]
        ).sort_index()


    def __init_total_effective_retirements(self) -> None:
        """
        Similar to 'Total' in 'Retirements of Effective Capacity (MW)' K6 table (v407)

        total_effective_retirements dims (c), t in <simulation_years[0], simulation_years[1]>
        """
        self.total_effective_retirements = self.effective_retirements.groupby(c.COUNTRY_CODE).sum()


    def __init_nameplate_capacity(
            self,
            simulation_years: tuple[int, int],
            capacity: pd.DataFrame
            ) -> None:
        """
        Simplified 'Nameplate Capacity (MW)' K1 table (v407)

        Sets base year capacity as 'Capacity (MW) Used' from A4 table.
        Inits rest of the years with 0.0s.

        nameplate_capacity dims (c, f), t in <simulation_years[0], simulation_years[1]>
            f in POW_FUELS
        """
        self.nameplate_capacity = (
            capacity
            .rename(columns={'Value': str(simulation_years[0])})
        )
        self.nameplate_capacity[
            [str(year) for year in range(simulation_years[0] + 1, simulation_years[1] + 1)]
        ] = 0.0


    def __init_effective_capacity(
            self,
            simulation_years: tuple[int, int],
            elec_for_power: pd.DataFrame
            ) -> None:
        """
        Simplified 'Effective Capacity (MW)' K3 table (v407)

        Sets base year capacity as 'Power Output MWy/y' from A4 table,
        which is equal to 'Capacity (MW) Used'/'Capacity Factor Used'.
        Inits rest of the years with 0.0s.

        effective_capacity dims (c, f), t in <simulation_years[0], simulation_years[1]>
            f in POW_FUELS
        """
        self.effective_capacity = (
            elec_for_power
            .rename(columns={'Value': str(simulation_years[0])})
        )
        self.effective_capacity[
            [str(year) for year in range(simulation_years[0] + 1, simulation_years[1] + 1)]
        ] = 0.0


    def __init_total_nameplate_capacity(self) -> None:
        """
        Inits total nameplate capacity.

        total_nameplate_capacity dims (c), t in <simulation_years[0], simulation_years[1]>
        """
        self.total_nameplate_capacity = self.nameplate_capacity.groupby(c.COUNTRY_CODE).sum()


    def __init_total_effective_capacity(self) -> None:
        """
        Inits total effective capacity.

        total_effective_capacity dims (c), t in <simulation_years[0], simulation_years[1]>
        """
        self.total_effective_capacity = self.effective_capacity.groupby(c.COUNTRY_CODE).sum()


    def __init_total_eff_investment(
            self,
            simulation_years: tuple[int, int],
            selected_countries: list[str],
            ) -> None:
        """
        Similar to 'Total Effective Generation Investment Needed (MW)' from K7 table.

        Inits total required investment with 0.0s for all years.

        total_eff_investment dims (c), t in <simulation_years[0], simulation_years[1]>
        """
        country_index = pd.Index(selected_countries, name=c.COUNTRY_CODE)
        self.total_eff_investment = pd.DataFrame(
            0.0,
            index=country_index,
            columns=[str(y) for y in range(simulation_years[0], simulation_years[1] + 1)]
        ).sort_index()


    def __set_new_investment_permitted(self) -> None:
        """
        Simplified 'Non-Manual New Investment Permitted' L3 table

        Series can be changed to pd.DataFrame with t, after implemention of
        'If yes, generation starts when?'

        new_investment_permitted dims (c, f), value series
            f in POW_FUELS
        """
        # All values on init
        y = str(self.__base_year)
        self.new_investment_permitted = (self.nameplate_capacity[y] > 0.0).astype(float)


    def __calculate_proportions_of_investment(
            self,
            year: int,
            k_investment: float,
            lcoe: pd.DataFrame
            ) -> None:
        """
        Simplified 'Proportions of New Investment' L5 table

        proportions_of_investment dims (c, f), t in <simulation_years[0], simulation_years[1]>
            f in POW_FUELS
        """
        y = str(year)
        # Simplified 'Logit Function Nominator (of fraction)' L4 table:
        self.proportions_of_investment.loc[:, y] = (
            self.new_investment_permitted
            / (
                1 + np.exp(k_investment * lcoe[y])
            )
        )

        # 'Proportions of New Investment' L5 table:
        total = self.proportions_of_investment.groupby(c.COUNTRY_CODE)[y].transform('sum')
        self.proportions_of_investment.loc[:, y] = self.proportions_of_investment.loc[:, y] / total


    def calculate_investment_year(
            self,
            year: int,
            k_investment: float,
            lcoe_tmp: pd.DataFrame,
            aggregated_ele_supply: pd.DataFrame,
            data: PowData
            ) -> None:
        """
        Update values according to formulas:

        - proportions_of_investment for t:
            see __calculate_proportions_of_investment(...)
        
        - nameplate_retirements for t-1:
            nameplate_retirements(f, t-1) = nameplate_capacity(f, t-1) / total_lifetime(f)
        - effective_retirements for t-1:
            effective_retirements(f, t-1) = effective_capacity(f, t-1) / total_lifetime(f)
        - total_retirements for t-1
    
        - total_eff_investment for t-1:
            total_eff_investment(t-1) = (
                max(
                    0,
                    (
                        aggregated_ele_supply(t)
                        + total_effective_retirements(t-1)
                        - total_effective_capacity(t-1)
                    )
                )
            )
        - eff_investment for t-1:
            eff_investment(f, t-1) = proportions_of_investment(f, t-1) * total_eff_investment(t-1)

        - effective_capacity for t-1:
            effective_capacity(f, t) = (
                effective_capacity(f, t-1) + eff_investment(f, t-1) - effective_retirements(f, t-1)
            )
        - total_effective_capacity for t
        """
        y = str(year)
        prev_y = str(year - 1)

        self.__calculate_proportions_of_investment(year, k_investment, lcoe_tmp)

        if year > self.__base_year:
            self.nameplate_retirements.loc[:, prev_y] = (
                self.nameplate_capacity[prev_y] / data.total_lifetime.iloc[:, -1]
            )
            self.effective_retirements.loc[:, prev_y] = (
                self.effective_capacity[prev_y] / data.total_lifetime.iloc[:, -1]
            )
            self.total_effective_retirements[prev_y] = (
                self.effective_retirements.loc[:, prev_y].groupby(c.COUNTRY_CODE).sum()
            )

            # Similar to K7 'Total Effective Generation Investment Needed (MW)'
            self.total_eff_investment.loc[:, prev_y] = (
                aggregated_ele_supply[y]
                + self.total_effective_retirements[prev_y]
                - self.total_effective_capacity[prev_y]
            ).clip(lower=0)

            # similar to L7 'Raw New Effective Investment  (MWavg)'
            self.eff_investment.loc[:, prev_y] = (
                self.proportions_of_investment.loc[:, prev_y]
                .mul(self.total_eff_investment.loc[:, prev_y], level=c.COUNTRY_CODE)
            )

            self.effective_capacity[y] = (
                self.effective_capacity[prev_y]
                + self.eff_investment.loc[:, prev_y]
                - self.effective_retirements.loc[:, prev_y]
            )
            self.nameplate_capacity[y] = (
                self.effective_capacity[y] / data.capacity_factor.iloc[:, -1]
            )
            self.total_effective_capacity[y] = (
                self.effective_capacity.loc[:, y].groupby(c.COUNTRY_CODE).sum()
            )
