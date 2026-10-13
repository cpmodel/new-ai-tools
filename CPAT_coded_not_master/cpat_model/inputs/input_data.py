import pandas as pd

from cpat_model.components.energy_consumption.energy_data import load_energy_consumptions
from cpat_model.components.elasticities.elasticities import load_elasticities
from cpat_model.components.gdp.gdp import get_gdp_data_country_filtered, GDPSource
from cpat_model.components.carbon_pricing.existing_ets import (
    get_existing_ets_nat_raw, get_existing_ets_reg_raw
)
from cpat_model.components.carbon_pricing.existing_ct import get_existing_ct_raw, load_do_ct_ets_exist
from cpat_model.components.efs.co2 import (
    get_co2_efs_global, load_ef_ghg, load_cal_val_and_air_data,
    EfGhgKeys
)
from cpat_model.components.policies.policies import (
    get_fuel_carbon_price_inclusion, get_sectors_carbon_price_inclusion, load_ghg_adjustments
)

from cpat_model.components.prices.ntx import load_co_benefits_ffs
from cpat_model.components.prices.oil_shares import load_oil_product_shares
from cpat_model.components.prices.international_prices import (
    load_global_fuel_demand,
    load_int_prices_regional_assumptions,
    load_international_prices
)
from cpat_model.components.prices.domestic_prices import load_prices_dom

from cpat_model.components.power.data import load_mpp, load_ic, load_lcoe_tmp

from cpat_model.components.distribution.data import (
    load_io_gtap, load_hh_survey, load_hh_elast, load_aspire, load_who_cooking,
    load_gdp_ratios, load_gtap_cpat_sector_crosswalk, load_countries_to_gtap10
)

class InputData:
    """
    TODO
    """
    ec_input: pd.DataFrame

    elasticities_input: pd.DataFrame

    gdp_data_country_filtered: dict[GDPSource, pd.DataFrame]
    gdp_countries_in_data: dict[GDPSource, set[str]]
    ngdp_d_world: pd.DataFrame
    pcpi_world: pd.DataFrame

    do_ct_ets_exist: pd.DataFrame
    existing_ets_s_nat_raw: pd.DataFrame
    existing_ets_f_nat_raw: pd.DataFrame
    existing_ets_permit_price_nat_raw: pd.DataFrame
    existing_ets_s_reg_raw: pd.DataFrame
    existing_ets_f_reg_raw: pd.DataFrame
    existing_ets_permit_price_reg_raw: pd.DataFrame
    existing_ct_s_raw: pd.DataFrame
    existing_ct_f_raw: pd.DataFrame
    existing_ct_f_rate_raw: pd.DataFrame
    existing_ct_rate_raw: pd.DataFrame

    co2_efs_global: pd.DataFrame
    ef_ghg: dict[EfGhgKeys, pd.DataFrame]
    cal_val_and_air_data: pd.DataFrame

    fuel_carbon_price_inclusion: pd.DataFrame
    sectors_carbon_price_inclusion: pd.DataFrame

    ghg_adjustments: pd.DataFrame

    co_benefits_ffs: pd.DataFrame
    oil_product_shares: pd.DataFrame
    global_fuel_demand: pd.DataFrame
    int_prices_regional_assumptions: pd.DataFrame
    international_prices: pd.DataFrame
    prices_dom: pd.DataFrame

    # Power
    mpp: pd.DataFrame
    ic: pd.DataFrame
    ic_base_year: pd.DataFrame
    lcoe_tmp: pd.DataFrame

    # Distribution
    distn_io_gtap: pd.DataFrame
    distn_hh_survey: pd.DataFrame
    distn_hh_elast: pd.DataFrame
    distn_aspire: pd.DataFrame
    distn_who_cooking: pd.DataFrame
    distn_gdp_ratios: pd.DataFrame
    distn_gtap_cpat_crosswalk: pd.DataFrame
    distn_countries_to_gtap10: pd.DataFrame

    def __init__(
            self,
            simulation_years: tuple[int, int],
            selected_countries: list[str]
            ) -> None:
        """
        Loads new data. TODO:
        """
        # energy balances input table
        self.ec_input = load_energy_consumptions(selected_countries)

        # elasticities input table
        self.elasticities_input = load_elasticities()

        # gdp
        (
            self.gdp_data_country_filtered,
            self.gdp_countries_in_data,
            self.ngdp_d_world,
            self.pcpi_world
        ) = get_gdp_data_country_filtered(selected_countries)

        # cabon pricing:
        self.do_ct_ets_exist = load_do_ct_ets_exist(selected_countries)
        (
            self.existing_ets_s_nat_raw,
            self.existing_ets_f_nat_raw,
            self.existing_ets_permit_price_nat_raw
        ) = get_existing_ets_nat_raw(selected_countries, simulation_years)
        (
            self.existing_ets_s_reg_raw,
            self.existing_ets_f_reg_raw,
            self.existing_ets_permit_price_reg_raw
        ) = get_existing_ets_reg_raw(selected_countries, simulation_years)
        (
            self.existing_ct_s_raw,
            self.existing_ct_f_raw,
            self.existing_ct_f_rate_raw,
            self.existing_ct_rate_raw
        ) = get_existing_ct_raw(selected_countries, simulation_years)

        # EFs
        self.co2_efs_global = get_co2_efs_global()
        self.ef_ghg = load_ef_ghg(selected_countries)
        self.cal_val_and_air_data = load_cal_val_and_air_data(selected_countries)

        # Policies
        self.fuel_carbon_price_inclusion = get_fuel_carbon_price_inclusion()
        self.sectors_carbon_price_inclusion = get_sectors_carbon_price_inclusion()

        # Policies / GHGs
        self.ghg_adjustments = load_ghg_adjustments(selected_countries)

        # Prices
        self.co_benefits_ffs = load_co_benefits_ffs(selected_countries, simulation_years)
        self.oil_product_shares = load_oil_product_shares(selected_countries)

        self.global_fuel_demand = load_global_fuel_demand(simulation_years)
        self.int_prices_regional_assumptions = load_int_prices_regional_assumptions(
            selected_countries
        )
        self.international_prices = load_international_prices(simulation_years)
        self.prices_dom = load_prices_dom(selected_countries)

        # Power
        self.mpp = load_mpp(selected_countries)
        self.ic, self.ic_base_year = load_ic(selected_countries, simulation_years)
        self.lcoe_tmp = load_lcoe_tmp(selected_countries, simulation_years)

        # Distribution
        self.distn_io_gtap = load_io_gtap(selected_countries)
        self.distn_hh_survey = load_hh_survey(selected_countries)
        self.distn_hh_elast = load_hh_elast(selected_countries)
        self.distn_aspire = load_aspire(selected_countries)
        self.distn_who_cooking = load_who_cooking(selected_countries)
        self.distn_gdp_ratios = load_gdp_ratios(selected_countries)
        self.distn_gtap_cpat_crosswalk = load_gtap_cpat_sector_crosswalk()
        self.distn_countries_to_gtap10 = load_countries_to_gtap10(selected_countries)
