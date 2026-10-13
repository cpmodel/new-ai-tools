import config

from cpat_model.inputs.dashboard_inputs import DashboardInputs
from cpat_model.inputs.distribution_inputs import DistributionInputs
from cpat_model.inputs.input_data import InputData
from cpat_model.scenario_results import ScenarioResults

from cpat_model.components.gdp.gdp import GDP
from cpat_model.components.policies.phaseouts import Phaseouts
from cpat_model.components.policies.policies import Policies
from cpat_model.components.policies.shadow_prices import ShadowPrices
from cpat_model.components.carbon_pricing.existing_ct import ExistingCT
from cpat_model.components.carbon_pricing.existing_ets import ExistingETS
from cpat_model.components.efs.co2 import EFsCO2
from cpat_model.components.elasticities.elasticities import Elasticities

from cpat_model.components.prices.domestic_prices import DomPrices
from cpat_model.components.prices.international_prices import IntPrices
from cpat_model.components.prices.prices import EnergyPrices

from cpat_model.components.energy_consumption.ec import EC
from cpat_model.components.emissions.em import CO2Emissions

from cpat_model.components.power.power import Power
from cpat_model.components.distribution.distribution import Distribution
from cpat_model.components.distribution import price_changes as distn_price_changes

import cpat_model.constants as c


def run_model():
    selected_countries = ['EGY']
    last_simulation_year = 2030

    base_year = 2022
    simulation_years = (base_year, last_simulation_year)
    input_data = InputData(
        simulation_years, selected_countries
    )

    # Distribution module is single-year/single-country; defaults to
    # analysing the last simulated year. See distribution/README.md and
    # distribution/docs/CPAT_Distribution_Module_Pseudocode.docx.
    distribution_inputs = DistributionInputs({'analysis_year': last_simulation_year})

    # Collects every scenario's finished component outputs so Distribution
    # can compare a policy scenario against the baseline scenario once all
    # scenarios have been run (see cpat_model.scenario_results.ScenarioResults
    # for why this cross-scenario comparison is needed).
    scenario_results: dict[str, ScenarioResults] = {}

    for scenario_name, dashboard_inputs in config.SCENARIOS.items():

        di = DashboardInputs(dashboard_inputs)

        elasticities = Elasticities(selected_countries, di.d, input_data)
        gdp = GDP(simulation_years, di.d, input_data)
        efs_co2 = EFsCO2(selected_countries, input_data, di.d['efs_selected'])
        policies = Policies(
            simulation_years, selected_countries, di.d['scenario_type'],
            input_data, di.d, gdp.deflator
        )
        phaseouts = Phaseouts(
            simulation_years, di.d['scenario_type'],
            di.d['energy_sector_reform'], di.d['is_pc_phaseout_baseline']
        )
        existing_ct = ExistingCT(
            selected_countries, simulation_years, input_data,
            di.d['existing_ct_apply'], di.d['existing_ct_growth']
        )
        existing_ets = ExistingETS(
            selected_countries, simulation_years, input_data,
            di.d['existing_ets_apply'], di.d['existing_ets_growth']
        )

        dom_prices = DomPrices(
            selected_countries, simulation_years, input_data, di.d['gov_energy_price_controls']
        )
        int_prices = IntPrices(
            simulation_years, selected_countries, input_data, di.d, gdp.deflator
        )
        energy_prices = EnergyPrices(
            simulation_years, di.d['scenario_type'], selected_countries, gdp, dom_prices, phaseouts,
            existing_ct, existing_ets, efs_co2, int_prices, input_data, di.d
        )

        shadow_prices = ShadowPrices(di.d, policies, efs_co2.ef_tco2_per_volume_unit)
        ec = EC(simulation_years, input_data, elasticities, shadow_prices)

        power = Power(
            simulation_years, selected_countries, di.d,
            energy_prices.ele_prices['rp'], policies.cp_trajectory, input_data
        )
        em = CO2Emissions(ec.ec_fossil_fuels_by_sector, efs_co2)

        for year in range(simulation_years[0], simulation_years[1] + 1):
            policies.calculate_policies_year(year, existing_ets.existing_ets_s_nat, di.d)
            energy_prices.calculate_prices_year(
                year, di.d['scenario_type'], policies, efs_co2,
                dom_prices, phaseouts
            )
            power.calculate_power_year(
                year, input_data,
                gdp.d_gdp_at_const_prices, elasticities, di.d,
                energy_prices, policies.p_cov_s, int_prices.uranium_fuel_cost
            )
            shadow_prices.calculate_year(year, energy_prices.pbc)
            ec.calculate_ec_year(
                year, gdp.d_gdp_at_const_prices,
                energy_prices, policies.p_cov_s, elasticities, input_data.lcoe_tmp,
                di.d, policies.exogenous_shock_on_ec,
                shadow_prices
            )
            em.calculate_em_year(year, ec.ec_fossil_fuels_by_sector)

        # Example on how to save output variables
        # power.generation.g.to_csv(f'power_generation_{scenario_name}.csv')

        scenario_results[scenario_name] = ScenarioResults(
            scenario_name=scenario_name,
            dashboard_inputs_d=di.d,
            gdp=gdp,
            policies=policies,
            energy_prices=energy_prices,
            ec=ec,
            em=em,
        )

    # Distributional analysis: compare each policy scenario against the
    # baseline scenario for the configured analysis year. See
    # distribution/README.md and
    # distribution/docs/CPAT_Distribution_Module_Pseudocode.docx.
    baseline_name = next(
        (
            name for name, results in scenario_results.items()
            if results.dashboard_inputs_d['scenario_type'] == c.BASELINE
        ),
        None
    )
    if baseline_name is None:
        raise RuntimeError(
            "No scenario with scenario_type == c.BASELINE found in config.SCENARIOS; "
            "the Distribution module needs a baseline run to compare policy scenarios against."
        )
    baseline = scenario_results[baseline_name]

    # Distribution operates on one country at a time (see distribution.py's
    # class docstring); loop over every selected country for each policy
    # scenario.
    distributions: dict[tuple[str, str], Distribution] = {}
    for scenario_name, policy in scenario_results.items():
        if scenario_name == baseline_name:
            continue
        for country in selected_countries:
            price_change_direct = distn_price_changes.derive_price_change_direct_from_scenarios(
                country, baseline, policy, distribution_inputs.d['analysis_year']
            )
            # only the 8 priced fuels feed Distribution's constructor; the
            # biomass fuels it adds internally have no Mitigation-side price
            price_change_direct_priced = price_change_direct.drop(c.COOKING_BIOMASS_FUELS)
            cp_revenue = distn_price_changes.derive_cp_revenue_from_scenarios(
                country, policy, distribution_inputs.d['analysis_year']
            )
            # TODO: national_totals=None here falls back to the household
            # survey's own implied population/consumption totals (see
            # rebasing.py module docstring), rather than a real national-
            # accounts figure. Wiring a real one needs
            # policy.gdp.population (population, already available) and a
            # household-consumption-in-LCU series derived from
            # policy.gdp.ngdp x rebasing.household_consumption_to_gdp_ratio
            # (input_data.distn_gdp_ratios) (consumption, not currently
            # exposed by GDP in a directly usable form) -- worth wiring up
            # once GDP's real-terms consumption series is available.
            national_totals = None

            distributions[(scenario_name, country)] = Distribution(
                distribution_inputs.d, country, input_data,
                price_change_direct_priced, cp_revenue, national_totals
            )

    return distributions

if __name__ == '__main__':
    run_model()
