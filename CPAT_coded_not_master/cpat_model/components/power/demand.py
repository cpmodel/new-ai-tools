import pandas as pd

from cpat_model.components.elasticities.elasticities import Elasticities

import cpat_model.constants as c

EMISSION_INTENSITY = 0.5 * 1e-3

# represents (1 + net exports pct + transmission losses pct)
SCALING_FACTOR = 1.2


class PowDemand:
    """
    Power Demand
    CPAT Excel: H section (v407)
    """
    demand: pd.DataFrame
    __rp_demand: pd.DataFrame

    aggregated_ele_demand: pd.DataFrame
    aggregated_ele_supply: pd.DataFrame
    def __init__(
            self,
            simulation_years: tuple[int, int],
            rp: pd.DataFrame, # energy_prices.ele_prices['rp']
            cp_trajectory: pd.DataFrame,
            ec_input: pd.DataFrame
        ) -> None:
        """
        Inits simplified 'Power Demand', Total Demand and 'Total Required Supply'
        """
        self.__init_year: int = simulation_years[0]
        self.__init_rp_demand(rp, cp_trajectory, simulation_years)
        self.__init_demand(ec_input, simulation_years)
        self.__init_aggregated_ele_demand()
        self.__init_aggregated_ele_supply()


    def __init_rp_demand(
            self,
            rp: pd.DataFrame,
            cp_trajectory: pd.DataFrame,
            simulation_years: tuple[int, int]
            ) -> None:
        """
        Simplified Power Prices

        Historical retail prices used,
        for other years values are filled with fixed values equal
        to the values from the last available historical year.
        Product of 'Carbon price trajectory used' and hardcoded EMISSION_INTENSITY
        is added on top of fixed prices.

        In the future should be calculated as in G1 and G2 tables.

        t in <simulation_years[0], simulation_years[1]>
        """
        last_rp_db_year = max(map(int, rp.columns))

        # it is all ele, so we can drop FuelCode index level
        self.__rp_demand =  rp.copy().droplevel(c.FUEL_CODE)

        # calculate tra as (ind + res)/2
        rp_tra = (
            self.__rp_demand.groupby(c.COUNTRY_CODE).mean()
            .assign(**{c.SECTOR_CODE: c.TRA})
            .set_index(c.SECTOR_CODE, append=True)
        )
        # add srv as a copy of ind
        rp_srv = (
            self.__rp_demand[self.__rp_demand.index.get_level_values(c.SECTOR_CODE) == c.IND]
            .droplevel(c.SECTOR_CODE)
            .assign(**{c.SECTOR_CODE: c.SRV})
            .set_index(c.SECTOR_CODE, append=True)
        )

        self.__rp_demand = pd.concat([self.__rp_demand, rp_tra, rp_srv]).sort_index()

        # fixed prices, past last hist year
        for year in range(last_rp_db_year + 1, simulation_years[1] + 1):
            self.__rp_demand[str(year)] = self.__rp_demand[str(last_rp_db_year)]

        columns = [str(year) for year in range(simulation_years[0], simulation_years[1] + 1)]
        self.__rp_demand = self.__rp_demand[columns]

        trajectory = cp_trajectory[columns] * EMISSION_INTENSITY

        self.__rp_demand = self.__rp_demand.add(trajectory, level=c.COUNTRY_CODE)


    def __init_demand(
            self,
            ec_input: pd.DataFrame,
            simulation_years: tuple[int, int]
            ) -> None:
        """
        Simplified 'Power Demand' from H3 table (v407)

        Unit: ktoe/y

        Inits with simulation_years[0] values. Values for other years are filled with 0.0s.

        demand dims: (c, s), t in <simulation_years[0], simulation_years[1]>
        """
        # Setting base year values of demand using values from Energy Consumption input
        # Same as 'Power Dem. (ktoe)' from H3 table
        self.demand = ec_input[[c.ELE]].rename(columns={c.ELE: str(simulation_years[0])})
        self.demand = self.demand[
            # dropping unecessary sectors from input ec table:
            ~self.demand.index.get_level_values(c.SECTOR_CODE)
            .isin([c.POW, c.ELEC, c.FTR, 'total_energy_use'])
        ]

        columns = [str(year) for year in range(simulation_years[0] + 1, simulation_years[1] + 1)]
        self.demand[columns] = 0.0


    def __init_aggregated_ele_demand(self) -> None:
        """
        Aggregated demand, similar to
        'Total Demand (excluding below)' from H3 table (v407)

        Unit: ktoe/y

        Inits with simulation_years[0] values. Values for other years are 0.0s.
        aggregated_ele_demand dims: (c), t in <simulation_years[0], simulation_years[1]>
        """
        self.aggregated_ele_demand = self.demand.groupby(c.COUNTRY_CODE).sum()


    def __init_aggregated_ele_supply(self) -> None:
        """
        Simplified 'Total Required Supply' from H3 table (v407)
        We use hardcoded SCALING_FACTOR for now.

        Unit: MWy/y

        Inits with simulation_years[0] values. Values for other years are 0.0s.
        aggregated_ele_supply dims: (c), t in <simulation_years[0], simulation_years[1]>
        """
        self.aggregated_ele_supply = self.aggregated_ele_demand.copy()
        self.aggregated_ele_supply[str(self.__init_year)] *= SCALING_FACTOR
        # convert ktoe/y to GWh/y and then to MWy/y
        self.aggregated_ele_supply[str(self.__init_year)] *= c.KTOE_TO_GWH * c.GWH_TO_MWY


    def calculate_demand_year(
            self,
            year: int,
            d_gdp_at_const_prices: pd.DataFrame,
            elasticities: Elasticities
            ) -> None:
        """
        Calculate from simulation_years[0] + 1 onwards
        t in <simulation_years[0], simulation_years[1]>
        """
        # TODO: split code to methods add unit tests
        if year > self.__init_year:
            y = str(year)
            prev_y = str(year - 1)
            self.demand[y] = self.demand[prev_y]

            idx = pd.IndexSlice
            # Similar to '{s} Price Effect' from table H1
            price_effect = (
                (self.__rp_demand[y] / self.__rp_demand[prev_y])
                .pow(
                    elasticities.ele_e[c.EL_EFF]
                    .loc[idx[:, [c.TRA, c.RES, c.IND, c.SRV], :], :]
                    .droplevel(c.FUEL_CODE)
                    .iloc[:, 0]
                )
            )

            # Similar to 'GDP Demand Change, {s}' from table H1
            growth_effect = (
                (d_gdp_at_const_prices[y] / d_gdp_at_const_prices[prev_y])
                .pow(
                    elasticities.ele_e[c.EL_INC]
                        .loc[idx[:, [c.TRA, c.RES, c.IND, c.SRV], :], :].iloc[:, 0],
                    level=c.COUNTRY_CODE
                )
            ).droplevel(c.FUEL_CODE)

            # Simplified 'Overall Multiple, {s}' form H1 table
            overall_multiple = price_effect * growth_effect

            # Updating demand for y:
            countries = list(self.demand.index.get_level_values(c.COUNTRY_CODE).unique())
            for country_code in countries:
                self.demand.loc[
                    idx[country_code, [c.ROD, c.RAL, c.AVI, c.NAV]], y
                    ] *= overall_multiple.loc[(country_code, c.TRA)]
                self.demand.loc[
                    idx[country_code, [c.RES]], y
                    ] *= overall_multiple.loc[(country_code, c.RES)]
                self.demand.loc[
                    idx[
                        country_code,
                        [c.FOO, c.MCH, c.IRN, c.NFM, c.MAC, c.CEM, c.OMN, c.CST, c.OEN]
                    ], y
                    ] *= overall_multiple.loc[(country_code, c.IND)]
                self.demand.loc[
                    idx[country_code, [c.SRV]], y
                    ] *= overall_multiple.loc[(country_code, c.SRV)]

            # update aggregated_ele_demand and supply:
            self.aggregated_ele_demand[y] = self.demand[y].groupby(c.COUNTRY_CODE).sum()
            self.aggregated_ele_supply[y] = (
                self.aggregated_ele_demand[y] * SCALING_FACTOR * c.KTOE_TO_GWH * c.GWH_TO_MWY
            )
