import pandas as pd

from cpat_model.inputs.dashboard_inputs import DashboardInputsDict
from cpat_model.components.prices.prices import EnergyPrices

from cpat_model.components.power.data import PowData

import cpat_model.constants as c

class PowVarCost:
    """
    Current Variable Costs
    CPAT Excel: D section (v407)

    At the moment primarly used to calculate Total Variable Costs D6.
    """
    vc: pd.DataFrame # variable cost D6

    def __init__(
            self,
            simulation_years: tuple[int, int],
            pow_index: pd.MultiIndex
            ) -> None:
        self.__init_vc(simulation_years, pow_index)


    def __init_vc(
            self,
            simulation_years: tuple[int, int],
            pow_index: pd.MultiIndex
            ) -> None:
        """
        Simplified 'Total Variable Costs' D6 table (v407)

        Inits with 0.0s.

        TODO: unify with __get_empty_frame from investment file

        vc dims (c, f), t in <simulation_years[0], simulation_years[1]>
            f in POW_FUELS
        """
        self.vc = pd.DataFrame() # only for pylint E1137
        self.vc = pd.DataFrame(
            0.0,
            index=pow_index,
            columns=[str(y) for y in range(simulation_years[0], simulation_years[1] + 1)]
        ).sort_index()


    def calculate_vc_year(
            self,
            year: int,
            data: PowData,
            prices: EnergyPrices,
            p_cov_s: pd.DataFrame,
            uranium_fuel_cost: float,
            d: DashboardInputsDict
            ) -> None:
        """
        Calculates values in range of simulation_years

        If any of the local variables are needed outside of this method,
        they can be changed to an attr of PowVarCost.
        """
        # TODO split to methods and add unit tests
        y = str(year)
        idx = pd.IndexSlice

        ### for D2
        # 'Fuel Price before any new carbon tax/ETS/price control' A10 table
        a10 = (
            # for (pow, coa), (pow, nga), (all, oop), (all, bio)
            prices.pbc.loc[idx[:, [c.POW, c.ALL], [c.COA, c.NGA, c.OOP, c.BIO]], y]
            .droplevel(c.SECTOR_CODE)
            .rename(
                index={c.OOP: c.OIL},
                level=c.FUEL_CODE
            )
        )

        # 'Moving Average Fuel Prices' A18 table
        if d['use_spot_fuel_prices_power']:
            a18 = a10 # just address
        else:
            raise NotImplementedError(
                "Moving Average Fuel Prices not implemented not yet supported."
            )

        # 'Fuel costs before carbon tax per kwh electricity produced' D2 table
        d2 = (
            a18
            / c.KWH_TO_GJ
            / data.thermal_efficiency.loc[idx[:, [c.COA, c.NGA, c.OIL, c.BIO]], :].iloc[:, -1]
        )
        d2.loc[idx[:, [c.OIL]]] /= c.BARREL_TO_GJ
        # add nuc as uranium_fuel_cost
        s_nuc = pd.Series(
            uranium_fuel_cost,
            index=pd.MultiIndex.from_product(
                [d2.index.get_level_values(c.COUNTRY_CODE).unique(), [c.NUC]],
                names=d2.index.names
            )
        )
        d2 = pd.concat([d2, s_nuc]).sort_index()

        # 'Coal implicit cost' D3 table TODO: implement D3 and add to D4

        # 'Total Variable Costs before Carbon Tax' D4 table
        d4 = data.var_om_cost.iloc[:, -1].copy()
        d4.name = None
        d4.loc[idx[:, [c.COA, c.NGA, c.OIL, c.BIO, c.NUC]]] += d2


        ### for D5:
        # 'Total new policy and price controls (including any sectoral exemptions)' A12 table
        a12 = (
            # for (pow, coa), (pow, nga), (all, oop)
            prices.nce.loc[idx[:, [c.POW, c.ALL], [c.COA, c.NGA, c.OOP]], y]
            .droplevel(c.SECTOR_CODE)
            .mul(
                p_cov_s.loc[
                    p_cov_s.index.get_level_values(c.SECTOR_CODE) == c.POW
                ].droplevel(c.SECTOR_CODE)[y],
                level=c.COUNTRY_CODE
            )
            .rename(
                index={c.OOP: c.OIL},
                level=c.FUEL_CODE
            )
        )

        # 'Carbon Tax ' D5 table - coa, nga, oil only, rest is 0.0s in Excel:
        d5 = (
            a12
            / c.KWH_TO_GJ
            / data.thermal_efficiency.loc[idx[:, [c.COA, c.NGA, c.OIL]], :].iloc[:, -1]
        )
        d5.loc[idx[:, [c.OIL]]] /= c.BARREL_TO_GJ

        # D6: vc = d4 + d5
        self.vc[y] = d4.copy()
        self.vc.loc[idx[:, [c.COA, c.NGA, c.OIL]], y] += d5.loc[idx[:, [c.COA, c.NGA, c.OIL]]]
