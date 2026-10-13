from typing import TypedDict, Literal

import cpat_model.constants as c

### define types:
# TODO: fix better way to define ScenarioType and ShandowPricingScenariosType (using c. throws a warning)
ScenarioType = Literal[
    'Baseline', 'Carbon tax', 'ETS', 'Feebates', 'Energy efficiency regulations',
    'TCP', 'Coal excise', 'Road fuel tax', 'Electricity emissions tax', 'Power feebate',
    'Electricity excise', 'Vehicle fuel economy', 'Residential efficiency regulations',
    'Industrial efficiency regulations'
]
ShandowPricingScenariosType = Literal[
    'Energy efficiency regulations', 'Vehicle fuel economy',
    'Residential efficiency regulations', 'Industrial efficiency regulations', 'Feebates'
]

ElasticititesSourcesType = Literal['Manual', 'Simple']
ElasticititesAdjustmentType = Literal['VHigh', 'High', 'Base', 'Low', 'VLow']

# prices
# TODO: PriceScenarioType add 'EIA', 'AVG', 'Manual', 'IMF-IEA'
PriceScenarioType = Literal['IMF', 'WB', 'IEA', 'IMF-WB']
PriceAdjustmentType = Literal['High', 'Base', 'Low']
GlobalEnergyDemandScenarioType = Literal["Stated Policies", "Announced Pledges", "Net Zero"]

# ntx
# TODO: PigouvianTaxType 'pigouvian' not implemented!
PigouvianTaxType = Literal['pigouvian', 'FFS', 'no tax']


class EnergySectorReform(TypedDict):
    """
    Type for 'Energy sector reform' (v361)
    TODO: check input fields again
    CPAT Excel: 230:237
    """
    is_ffs_phaseout_prod: bool # Phase out fossil fuel subsidies (producer)?
    # baseline =MTInputs!F57 -> Dashboard!$V$66; scenario: =MTInputs!F53 -> always True

    ffs_phaseout_prod: int # FFS (producer) phaseout period?
    # baseline =MTInputs!$F$58 -> Dashboard!$V$67, scenario: =MTInputs!F55 -> Dashboard!$Q$3

    ffs_start_prod: int # Year to begin FFS (producer) phaseout?
    # always =MTInputs!F54 -> Dashboard!$P$3

    ffs_phaseout_share_prod: float # Share of subsidies to phase-out
    # baseline =MTInputs!$F$59 -> Dashboard!$V$68, scenario: =MTInputs!$F$56 -> Dashboard!$V$65


    is_ffs_phaseout_cons: bool # Phase out fossil fuel subsidies (consumer)?
    # always =MTInputs!$F$61 -> always True

    ffs_phaseout_cons: int # FFS (consumer) phaseout period?
    # baseline =MTInputs!$F$66 -> Dashboard!$V$72, scenario: =MTInputs!F63 -> Dashboard!$Q$4

    ffs_start_cons: int # Year to begin FFS (consumer) phaseout?
    # always =MTInputs!F62 -> Dashboard!$P$4

    ffs_phaseout_share_cons: float # Share of subsidies to phase-out
    # baseline =MTInputs!$F$67 -> Dashboard!$V$73, scenario: =MTInputs!$F$64 -> Dashboard!$V$70


    is_pc_phaseout: bool # Phaseout price control?
    pc_start: int # Year to introduce price control phaseout
    pc_phaseout: int # Price control phaseout period (years)


class DashboardInputsDict(TypedDict):
    """
    Dashboard inputs types defined.
    Based on CPAT Excel (v361) if not stated otherwise
    """
    scenario_type: ScenarioType # Dashboard!$E$3 (v407)

    # elasticities
    source_income_elasticities: ElasticititesSourcesType # Dashboard!$L$21 (v407)
    source_price_elasticities: ElasticititesSourcesType # Dashboard!$L$20 (v407)
    elasticities_adjustment_income: ElasticititesAdjustmentType # Dashboard!$Q$21 (v407)
    elasticities_adjustment_price: ElasticititesAdjustmentType # Dashboard!$Q$20 (v407)

    # gdp
    deflator_selection: Literal['World', 'Country'] # Dashboard!$G$66
    results_year: int # Dashboard!$G$62 Default 2021
    gdp_adj: Literal['High', 'Base', 'Low'] # Dashboard!$Q$19

    # efs
    efs_selected: Literal['IIASA', 'IEA'] # Dashboard!$L$22

    ### Policies
    ct_tax_complimentary_to_ets: bool # Dashboard!$Q$76
    exempt_phaseout: bool # MTInputs!$F$49 -> hardcoded True
    year_pha: int # Dashboard!$P$2
    exempt_phaseout_period: int # Dashboard!$Q$2

    # 'Adjustment to efficiency margins for shadow pricing policies:' (v407)
    # Dashboard!$L$68 Dashboard!$L$69 Dashboard!$L$70 Dashboard!$L$71 Dashboard!$L$72
    adj_eff_margins_shadow_pricing: dict[ShandowPricingScenariosType, float]

    ets_adj: float # Dashboard!$Q$78
    tax_pathway: Literal['rea', 'nom'] # Dashboard!$G$20 # TODO: 'rea' not implemented
    # Dashboard!$G$19 and Dashboard!$L$78 - they do the same thing but $L$78 overrides $G$19:
    cp_trajectory_type: Literal['exp', 'lin', 'cst', 'per']
    start_cp_year: int # Dashboard!$I$4 , TODO the same as cp_intro in some parts of the code
    target_cp_year: int # Dashboard!$I$7
    start_cp: int # Dashboard!$I$5
    target_cp: int # Dashboard!$I$6
    ct_exp_rate: float # Dashboard!$L$79

    # Exisiting carbon prices mechanisms
    existing_ct_apply: bool # Dashboard!$Q$71
    existing_ct_growth: float # Dashboard!$Q$72
    existing_ets_apply: bool # Dashboard!$Q$74
    existing_ets_growth: float # Dashboard!$Q$75

    # Phaseouts
    is_pc_phaseout_baseline: bool # Dashboard!$V$76
    energy_sector_reform: EnergySectorReform


    ### prices
    price_scenario_used: PriceScenarioType # Dashboard!$L$18
    int_energy_price_adjustment: PriceAdjustmentType # Dashboard!$Q$18
    gov_energy_price_controls: Literal['Bucketed', 'Manual', 'None'] # Dashboard!$V$75
    global_energy_demand_scenario: GlobalEnergyDemandScenarioType # Dashboard!$G$73

    # ntx
    pigouvian_phase_in_t: int # Dashboard!$V$23
    add_efficient_pigouvian_tax: PigouvianTaxType # Dashboard!$V$22
    add_additional_excise_tax: bool # Dashboard!$V$24


    ### Power
    cf_override_if_above: float # Dashboard!$G$129 CFOverrideIfAbove (v407)
    cf_override_if_below_fos_oth: float # Dashboard!$G$127 CFOverrideIfBelowFosOth (v407)
    cf_override_if_below_wnd_sol: float # Dashboard!$G$128 CFOverrideIfBelowWndSol (v407)
    minimum_thermal_efficiency: float # Dashboard!G125 MinimumThermalEfficiency (v407)
    k_investment: float # Dashboard!$L$133 kInvestment (v407)
    k_dispatch: float # Dashboard!$G$121 kDispatch (v407)
    max_coa_cf: float # Dashboard!$G$123 MaximumCoalCF (v407)
    max_nga_cf: float # Dashboard!$G$124 MaximumGasCF (v407)
    use_spot_fuel_prices_power: bool # Dashboard!$G$122 UseSpotFuelPricesInEngineerModel (v407)

    ### EC
    always_exempt_res_lpg_ker: bool # Dashboard!$V$19 AlwaysExemptResLPGKer (v407)
    additional_eff_gains: dict[ # (v407)
        Literal[
            'pow', # Dashboard!$L$61 AdditionalEfficiencyPower
            'tra', # Dashboard!$L$62 AdditionalEfficiencyRoadVehicle
            'res', # Dashboard!$L$63 AdditionalEfficiencyResidential
            'ind' # Dashboard!$L$64 AdditionalEfficiencyIndustrial
        ],
        float
    ] # Additional policy-induced annual efficiency gains



class DashboardInputs:
    """
    Contains all Dashboard inputs in d: DashboardInputsDict
    Hardcoded here for now -> TODO: UI
    """
    d: DashboardInputsDict

    def __init__(self, config_input: dict | None = None) -> None:
        # default
        self.d = {
            'scenario_type': c.BASELINE,

            'source_income_elasticities': 'Simple',
            'source_price_elasticities': 'Simple',
            'elasticities_adjustment_income': 'Base',
            'elasticities_adjustment_price': 'Base',

            'global_energy_demand_scenario': 'Stated Policies',

            'deflator_selection': 'World',
            'results_year': 2022,
            'gdp_adj': 'Base',

            'efs_selected': 'IIASA',

            'is_pc_phaseout_baseline': True,
            'energy_sector_reform': {
                'is_ffs_phaseout_prod': True,
                'ffs_phaseout_prod': 5,
                'ffs_start_prod': 2025,
                'ffs_phaseout_share_prod': 0.5,

                'is_ffs_phaseout_cons': False,
                'ffs_phaseout_cons': 5,
                'ffs_start_cons': 2025,
                'ffs_phaseout_share_cons': 0.5,

                'is_pc_phaseout': True,
                'pc_phaseout': 5,
                'pc_start': 2025,
            },

            'existing_ct_apply': True,
            'existing_ct_growth': 0.1,
            'existing_ets_apply': True,
            'existing_ets_growth': 0.1,

            'gov_energy_price_controls': 'Bucketed',
            'price_scenario_used': 'IMF-WB',
            'int_energy_price_adjustment': 'Base',

            'start_cp_year': 2025,
            'target_cp_year': 2030,
            'start_cp': 0,
            'target_cp': 25,

            'exempt_phaseout': True,
            'year_pha': 2025,
            'exempt_phaseout_period': 5,
            'ct_tax_complimentary_to_ets': False,

            'adj_eff_margins_shadow_pricing': {
                c.ENERGY_EFFICIENCY_REGULATIONS: 0.7,
                c.VEHICLE_FUEL_ECONOMY: 0.7,
                c.RESIDENTIAL_EFFICIENCY_REGULATIONS: 0.7,
                c.INDUSTRIAL_EFFICIENCY_REGULATIONS: 0.7,
                c.FEEBATES: 1.0
            },

            'pigouvian_phase_in_t': 5,
            'add_efficient_pigouvian_tax': 'no tax',
            'add_additional_excise_tax': False,

            'ets_adj': 0.9,
            'tax_pathway': 'nom',
            'cp_trajectory_type': 'lin',
            'ct_exp_rate': 0.00,

            'cf_override_if_above': 1.0,
            'cf_override_if_below_fos_oth': 0.01,
            'cf_override_if_below_wnd_sol': 0.1,
            'minimum_thermal_efficiency': 0.1,
            'k_investment': 2.0,
            'k_dispatch': 2.0,
            'max_coa_cf': 0.9,
            'max_nga_cf': 0.9,
            # TODO: in Excel False as default, here True as moving prices not implemented:
            'use_spot_fuel_prices_power': True,

            'always_exempt_res_lpg_ker': False,
            'additional_eff_gains': {
                'pow': 0.0,
                'tra': 0.0,
                'res': 0.0,
                'ind': 0.0
            }
        }

        if config_input:
            self.d.update(config_input)
