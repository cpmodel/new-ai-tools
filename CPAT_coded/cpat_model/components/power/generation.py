import pandas as pd
import numpy as np

from cpat_model.inputs.dashboard_inputs import DashboardInputsDict
from cpat_model.components.power.investment import PowInvestment

import cpat_model.constants as c

class PowGeneration:
    """
    Simplified Power Generation

    Inlcudes tables from 'Dispatch Decision' - section I
    """
    g: pd.DataFrame # genaration

    def __init__(
            self,
            effective_capacity: pd.DataFrame
            ) -> None:
        self.__init_generation(effective_capacity)

    def __init_generation(
            self,
            effective_capacity: pd.DataFrame
            ) -> None:
        """
        Simplified 'Generation (MWy)' I3 table

        Inits with copied initial effective_capacity.
        (simulation_years[0] values with elec_for_power values, rest of the years with 0.0s)

        g dims (c, f), t in <simulation_years[0], simulation_years[1]>
            f in POW_FUELS
        """
        self.g = effective_capacity.copy()

    def calculate_generation_year(
            self,
            year: int,
            investment: PowInvestment,
            aggregated_ele_supply: pd.DataFrame,
            var_cost: pd.DataFrame,
            d: DashboardInputsDict,
            ) -> None:
        """
        Simplified 'Generation (MWy)' I3 table

        Calculates values in range of simulation_years
        """
        # TODO: split code to methods add unit tests
        y = str(year)
        coa_nga = [c.COA, c.NGA]

        mask = ~self.g.index.get_level_values(c.FUEL_CODE).isin(coa_nga)
        self.g.loc[mask, y] = investment.effective_capacity.loc[mask, y]

        # 'Total non-gas non-coal' (coa, nga are initialized as 0.0 so they won't count):
        total_no_nga_no_coa = self.g[y].groupby(c.COUNTRY_CODE).sum()

        # Simplified 'Required Generation from Coal and Natural Gas' from I1 table
        required_g = (aggregated_ele_supply[y] - total_no_nga_no_coa).clip(lower=0)

        # 'Maximum Coal' and 'Maximum Gas' from I2 table
        max_coa_nga = (
            investment.nameplate_capacity.loc[
                investment.nameplate_capacity.index.get_level_values(c.FUEL_CODE).isin(coa_nga),
                y
            ]
        )
        max_coa_nga.loc[
            max_coa_nga.index.get_level_values(c.FUEL_CODE).isin([c.COA])
        ] *= d['max_coa_cf']
        max_coa_nga.loc[
            max_coa_nga.index.get_level_values(c.FUEL_CODE).isin([c.NGA])
        ] *= d['max_nga_cf']

        # Simplified 'Minimum Coal' and 'Minimum Gas' from I2 table
        min_coa_nga = (required_g.sub(max_coa_nga, level=c.COUNTRY_CODE)).clip(lower=0)
        # 'Remainder after Minimas' from I2 table
        min_coa_nga_total = min_coa_nga.groupby(c.COUNTRY_CODE).sum()

        # 'Total Variable Cost - Coal' and 'Total Variable Cost - Nat Gas' from I2 table
        mask = var_cost.index.get_level_values(c.FUEL_CODE).isin(coa_nga)
        var_cost_coa_nga = var_cost.loc[mask, y].copy()
        var_cost_coa_nga_min = var_cost_coa_nga.groupby(c.COUNTRY_CODE).min()

        # 'Logit Function Coal' and 'Logit Function Nat Gas' from I2 table
        proportion_coa_nga = np.exp(
            - d['k_dispatch']
            * var_cost_coa_nga.div(var_cost_coa_nga_min, level=c.COUNTRY_CODE)
        )
        # 'Proportion of Coal in remaining gen mix'
        # and 'Proportion of Natural Gas in remaining gen mix' from I2 table
        total_proportion = proportion_coa_nga.groupby(c.COUNTRY_CODE).sum()
        proportion_coa_nga = proportion_coa_nga.div(total_proportion, level=c.COUNTRY_CODE)

        # 'Total Coal Generation' and 'Total Gas Generation'
        mask = self.g.index.get_level_values(c.FUEL_CODE).isin(coa_nga)
        self.g.loc[mask, y] = (
            min_coa_nga
            + (required_g - min_coa_nga_total).mul(proportion_coa_nga, level=c.COUNTRY_CODE)
        )
