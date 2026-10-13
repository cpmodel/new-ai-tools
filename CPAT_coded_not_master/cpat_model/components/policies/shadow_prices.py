import pandas as pd

from cpat_model.inputs.dashboard_inputs import DashboardInputsDict
from cpat_model.components.policies.policies import Policies, EXPLICIT_PRICING_VARIANT

import cpat_model.constants as c

SHADOW_PRICES_CALC_IDX_PAIRS = [
    (c.ALL, c.BIO), (c.ALL, c.OOP), (c.ROD, c.DIE), (c.ROD, c.GSO), (c.ROD, c.LPG), (c.IND, c.KER),
    (c.POW, c.COA), (c.RES, c.COA), (c.IND, c.COA), (c.POW, c.NGA), (c.RES, c.NGA), (c.IND, c.NGA)
]
SHADOW_PRICES_IDX_PAIRS = [
    (c.ALL, c.BIO), (c.ALL, c.OOP), (c.ALL, c.DIE), (c.ALL, c.GSO), (c.ALL, c.LPG), (c.ALL, c.KER),
    (c.POW, c.COA), (c.RES, c.COA), (c.IND, c.COA), (c.POW, c.NGA), (c.RES, c.NGA), (c.IND, c.NGA)
]

class ShadowPrices:
    """
    Shadow Prices from 'Policies - {scenario_type}' section (v407)
    CPAT Excel: baseline 2049:2102, scenario: 6820:6873
    """
    shadow_prices: pd.DataFrame
    shadow_prices_incr: pd.DataFrame
    percent_of_shadow_p: pd.Series

    # helper:
    __shadow_prices_incr_sum: pd.Series

    # TODO: test all methods

    def __init__(
            self,
            d: DashboardInputsDict,
            policies: Policies,
            ef_tco2_per_volume_unit: pd.DataFrame
            ) -> None:
        self.__shadow_pricing = EXPLICIT_PRICING_VARIANT[d['scenario_type']]
        self.__first_calc_year = min(map(int, policies.cp_trajectory.columns))

        self.__set_shadow_prices(
            d['scenario_type'], policies.cp_trajectory, ef_tco2_per_volume_unit
        )
        self.__init_shadow_prices_incr()
        self.__set_percent_of_shadow_p(d, policies.sectors_carbon_price_inclusion)


    def __set_shadow_prices(
            self,
            scenario_type: str,
            cp_trajectory: pd.DataFrame,
            ef_tco2_per_volume_unit: pd.DataFrame
            ) -> None:
        """
        'Shadow prices:' (v407)
        CPAT Excel: baseline 2058:2070, scenario: 6829:6841

        For fuels other than coa and nga, sectors changed to all.
        
        shadow_prices dims (c, s, f), t in <simulation_years[0], simulation_years[1]>
            (s, f) in [pow, res, ind]x[coa, nga] and [all]x[gso, die, lpg, ker, oop, bio]
        """
        if ((scenario_type == c.ETS) or not self.__shadow_pricing):
            countries = list(cp_trajectory.index.get_level_values(c.COUNTRY_CODE).unique())
            self.shadow_prices = pd.DataFrame(
                0.0,
                index=pd.MultiIndex.from_tuples(
                    [
                        (country, sector, fuel)
                        for country in countries
                        for sector, fuel in SHADOW_PRICES_IDX_PAIRS
                    ],
                    names=c.ID_COL_NAMES
                ),
                columns=cp_trajectory.columns
            ).sort_index()
        else:
            mask = (
                ef_tco2_per_volume_unit.index.droplevel(c.COUNTRY_CODE)
                .isin(SHADOW_PRICES_CALC_IDX_PAIRS)
            )
            self.shadow_prices = (
                cp_trajectory
                .mul(ef_tco2_per_volume_unit.loc[mask].iloc[:, -1], axis=0, level=c.COUNTRY_CODE)
            )

            # change sector to all if not in [coa, nga]
            idx = self.shadow_prices.index.to_frame()
            mask = ~idx[c.FUEL_CODE].isin([c.COA, c.NGA])
            idx.loc[mask, c.SECTOR_CODE] = c.ALL
            self.shadow_prices.index = pd.MultiIndex.from_frame(idx)
            self.shadow_prices.sort_index(inplace=True)


    def __init_shadow_prices_incr(self) -> None:
        """
        'Shadow prices (annual percentage point increase if were a carbon tax):' (v407)
        CPAT Excel: baseline 2071:2083, scenario: 6842:6854

        Inits with all 0.0s

        For fuels other than coa and nga, sectors changed to all.

        shadow_prices dims (c, s, f), t in <simulation_years[0], simulation_years[1]>
            (s, f) in [pow, res, ind]x[coa, nga] and [all]x[gso, die, lpg, ker, oop, bio]
        """
        self.shadow_prices_incr = self.shadow_prices.copy()
        self.shadow_prices_incr.loc[:, :] = 0.0

        self.__shadow_prices_incr_sum = self.shadow_prices_incr[str(self.__first_calc_year)]


    def __update_shadow_prices_incr(self, year: int, pbc: pd.DataFrame) -> None:
        """
        'Shadow prices (annual percentage point increase if were a carbon tax):' (v407)
        CPAT Excel: baseline 2071:2083, scenario: 6842:6854
        """
        y = str(year)
        if self.__shadow_pricing:
            self.shadow_prices_incr[y] = (
                self.shadow_prices[y]
                / (self.shadow_prices[y] + pbc[y])
            ).fillna(0.0)

            if year != self.__first_calc_year:
                self.shadow_prices_incr[y] -= self.__shadow_prices_incr_sum

        self.__shadow_prices_incr_sum += self.shadow_prices_incr[y]


    def __set_percent_of_shadow_p(
            self,
            d: DashboardInputsDict,
            sectors_carbon_price_inclusion: pd.DataFrame
            ) -> None:
        """
        '% of shadow price impacting efficiency by sector' (v407)
        CPAT Excel: baseline 2084:2102, scenario: 6855:6873
        
        percent_of_shadow_p dims (s), data series
        """
        self.percent_of_shadow_p = sectors_carbon_price_inclusion[
            sectors_carbon_price_inclusion.index.get_level_values(c.SECTOR_CODE) != c.IND
        ].iloc[:, 0] # change to series
        scenarios_shadow_price = [
            c.ENERGY_EFFICIENCY_REGULATIONS, c.VEHICLE_FUEL_ECONOMY,
            c.RESIDENTIAL_EFFICIENCY_REGULATIONS, c.INDUSTRIAL_EFFICIENCY_REGULATIONS, c.FEEBATES
        ]

        if (
            (not self.__shadow_pricing) # logic in E column in CPAT Excel
            # 'not in Baseline' logic in F column in CPAT Excel
            # '% of shadow price on efficiency margin' only for scenarios_shadow_price
            # see CPAT Excel baseline C2053 and scenario C6824
            or (d['scenario_type'] not in scenarios_shadow_price)
        ):
            # all 0.0s
            self.percent_of_shadow_p = self.percent_of_shadow_p.astype(float)
            self.percent_of_shadow_p[:] = 0.0

        else:
            value = d['adj_eff_margins_shadow_pricing'][d['scenario_type']]
            # if True -> value, if False -> 0.0
            self.percent_of_shadow_p = self.percent_of_shadow_p.astype(float) * value

    def calculate_year(self, year: int, pbc: pd.DataFrame) -> None:
        """
        Calculates values in shadow_prices_incr for a given year
        """
        self.__update_shadow_prices_incr(year, pbc)
