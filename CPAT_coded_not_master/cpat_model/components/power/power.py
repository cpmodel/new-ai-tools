from typing import TYPE_CHECKING

import pandas as pd

from cpat_model.inputs.dashboard_inputs import DashboardInputsDict
from cpat_model.components.elasticities.elasticities import Elasticities
from cpat_model.components.prices.prices import EnergyPrices

from cpat_model.components.power.variable_cost import PowVarCost
from cpat_model.components.power.demand import PowDemand
from cpat_model.components.power.investment import PowInvestment
from cpat_model.components.power.generation import PowGeneration
from cpat_model.components.power.data import PowData

if TYPE_CHECKING:
    from cpat_model.inputs.input_data import InputData


class Power:
    """
    Simplified 'Technoeconomic ('engineer') power model'
    CPAT Excel section starts: baseline row 2701, scenario row 7501 (v412)
    """
    data: PowData
    var_cost: PowVarCost
    demand: PowDemand
    investment: PowInvestment
    generation: PowGeneration

    def __init__(
            self,
            simulation_years: tuple[int, int],
            selected_countries: list[str],
            d: DashboardInputsDict,
            rp: pd.DataFrame, # energy_prices.ele_prices['rp']
            cp_trajectory: pd.DataFrame,
            input_data: 'InputData'
            ) -> None:
        """
        Inits instances for all Power components,
        please see each component for more detailed information.
        """
        self.data = PowData(selected_countries, d, input_data)
        self.var_cost = PowVarCost(simulation_years, self.data.pow_index)
        self.demand = PowDemand(simulation_years, rp, cp_trajectory, input_data.ec_input)
        self.investment = PowInvestment(
            simulation_years, selected_countries, self.data
        )
        self.generation = PowGeneration(self.investment.effective_capacity)


    def calculate_power_year(
            self,
            year: int,
            input_data: 'InputData',
            d_gdp_at_const_prices: pd.DataFrame,
            elasticities: Elasticities,
            d: DashboardInputsDict,
            prices: EnergyPrices,
            p_cov_s: pd.DataFrame,
            uranium_fuel_cost: float
        ) -> None:
        """
        Updates values for all Power components
        """
        self.var_cost.calculate_vc_year(year, self.data, prices, p_cov_s, uranium_fuel_cost, d)
        self.demand.calculate_demand_year(year, d_gdp_at_const_prices, elasticities)
        self.investment.calculate_investment_year(
            year, d['k_investment'], input_data.lcoe_tmp,
            self.demand.aggregated_ele_supply, self.data
        )
        self.generation.calculate_generation_year(
            year, self.investment, self.demand.aggregated_ele_supply,
            self.var_cost.vc, d
        )
        