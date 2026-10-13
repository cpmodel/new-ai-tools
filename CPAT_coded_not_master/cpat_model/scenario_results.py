from dataclasses import dataclass

from cpat_model.inputs.dashboard_inputs import DashboardInputsDict
from cpat_model.components.gdp.gdp import GDP
from cpat_model.components.policies.policies import Policies
from cpat_model.components.prices.prices import EnergyPrices
from cpat_model.components.energy_consumption.ec import EC
from cpat_model.components.emissions.em import CO2Emissions


@dataclass
class ScenarioResults:
    """
    Bundles one scenario run's key component outputs together so they can be
    compared *across* scenarios once run_model's main per-year loop has
    finished computing all of them.

    Every existing component (GDP, Policies, EnergyPrices, EC, CO2Emissions, ...)
    is scoped to a single scenario run: run_model() builds a fresh instance of
    each per scenario_name in config.SCENARIOS, and the year loop mutates those
    specific instances in place. Nothing in the codebase previously needed to
    look at two scenarios at once.

    The Distribution module does: it reports a policy scenario's household
    impact *relative to* the baseline (no-policy) scenario for the same
    country and analysis year -- it needs both scenarios' EnergyPrices (to
    derive the price change per fuel) and CO2Emissions/Policies (to derive
    carbon-pricing revenue), not just the policy scenario alone. See
    distribution/docs/CPAT_Distribution_Module_Pseudocode.docx §2 and Step 1,
    and cpat_model.components.distribution.distribution.Distribution.
    """
    scenario_name: str
    dashboard_inputs_d: DashboardInputsDict
    gdp: GDP
    policies: Policies
    energy_prices: EnergyPrices
    ec: EC
    em: CO2Emissions
