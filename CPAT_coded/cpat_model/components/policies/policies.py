from typing import TYPE_CHECKING

import pandas as pd
import numpy as np

from cpat_model.inputs.dashboard_inputs import DashboardInputsDict
import cpat_model.constants as c
from config import DATA_PATH

if TYPE_CHECKING:
    from cpat_model.inputs.input_data import InputData


PKL_DATA_PATH = 'cpat_data/new_data_pkl'

# 'Explicit pricing variant?'
# CPAT Excel: Mitigation row 235 (v407)
EXPLICIT_PRICING_VARIANT = {
    c.BASELINE: False, c.CARBON_TAX: True, c.ETS: True, c.FEEBATES: False,
    c.ENERGY_EFFICIENCY_REGULATIONS: False, c.TCP: True, c.COAL_EXCISE: True,
    c.ROAD_FUEL_TAX: True, c.ELECTRICITY_EMISSION_TAX: True, c.POWER_FEEBATE: False,
    c.ELECTRICITY_EXCISE: True, c.VEHICLE_FUEL_ECONOMY: False,
    c.RESIDENTIAL_EFFICIENCY_REGULATIONS: False, c.INDUSTRIAL_EFFICIENCY_REGULATIONS: False
}

class Policies:
    """
    Existing carbon taxes from TODO
    'Existing carbon pricing mechanisms (ETSs and carbon taxes)' section (v361)
    CPAT Excel: section start 1665 #TODO time in methods
    """
    # Scenario specific input data
    sectors_carbon_price_inclusion: pd.DataFrame
    fuel_carbon_price_inclusion: pd.DataFrame
    explicit_pricing_variant: pd.DataFrame

    # Public attributes:
    apply_tax_f: pd.DataFrame
    apply_tax_s: pd.DataFrame
    apply_tax_f_extended: pd.DataFrame
    apply_tax_s_extended: pd.DataFrame

    p_cov_s: pd.DataFrame
    p_cov_f: pd.DataFrame
    p_based_policies_cov_s_f: pd.DataFrame

    cp_trajectory: pd.DataFrame

    exogenous_shock_on_ec: pd.DataFrame

    def __init__(
            self,
            simulation_years: tuple[int, int],
            selected_countries: list[str],
            scenario_type: str,
            # Preloaded input data:
            input_data: 'InputData',
            # Dashboard inputs:
            d: DashboardInputsDict,
            deflator: pd.DataFrame
            ) -> None:
        self.__set_exogenous_shock_on_ec(selected_countries, input_data.ghg_adjustments)
        self.__set_carbon_tax_inclusion(input_data, scenario_type)
        self.__set_apply_tax_s(scenario_type)
        self.__set_apply_tax_f(scenario_type)

        self.__get_apply_tax_f_extended(selected_countries)
        self.__get_apply_tax_s_extended(selected_countries)

        self.cp_trajectory = self.__set_cp_trajectory(simulation_years, scenario_type, d, deflator)

        # init empty dfs
        self.p_cov_s = pd.DataFrame()
        self.p_cov_f = pd.DataFrame()
        self.p_based_policies_cov_s_f = pd.DataFrame()


    def calculate_policies_year(
            self,
            year: int,
            existing_ets_s: pd.DataFrame, # ExistingETS.existing_ets_s
            d: DashboardInputsDict
        ) -> None:
        """
        Calculates p_cov_s, p_cov_f and p_based_policies_cov_s_f values
        for years in range of simulation years.
        """
        y = str(year)

        # TODO: make this cleaner: init dfs and update years, clean d[] params
        # TODO: modify 1->1.0, 0->0.0 in p_cov_s, p_cov_f, p_based_policies_cov_s_f
        self.p_cov_s = pd.concat(
            [
                self.p_cov_s,
                self.__get_p_cov_s(
                    d['scenario_type'], d['ct_tax_complimentary_to_ets'], existing_ets_s[[y]],
                    d['start_cp_year'], year, d['exempt_phaseout'],
                    d['year_pha'], d['exempt_phaseout_period']
                )
            ],
            axis=1
        )

        self.p_cov_f = pd.concat(
            [
                self.p_cov_f,
                self.__get_p_cov_f(
                    d['scenario_type'], d['start_cp_year'], year,
                    d['exempt_phaseout'], d['year_pha'], d['exempt_phaseout_period']
                )
            ],
            axis=1
        )

        self.p_based_policies_cov_s_f = pd.concat(
            [
                self.p_based_policies_cov_s_f,
                self.__get_p_based_policies_cov_s_f(year)
            ],
            axis=1
        )


    def __set_exogenous_shock_on_ec(
            self,
            selected_countries: list[str],
            ghg_adjustments: pd.DataFrame
            ) -> None:
        """
        'Exogenous shock on energy consumption' (v407)
        CPAT Excel: baseline row 2122, scenario row 6893

        same as
        'Exogenous shocks' (v407)
        CPAT Excel: row 462

        exogenous_shock_on_ec dims (c), t in <2023, 2024>
        """
        # TODO test
        db_countries = set(ghg_adjustments.index.get_level_values(c.COUNTRY_CODE).unique())
        missing_countries = set(selected_countries) - db_countries

        self.exogenous_shock_on_ec = ghg_adjustments.copy()
        if missing_countries:
            index = pd.Index(list(missing_countries), name=c.COUNTRY_CODE)
            missing_countries_df = pd.DataFrame(
                0.0, index=index, columns=['2023', '2024']
            )
            self.exogenous_shock_on_ec = pd.concat(
                [self.exogenous_shock_on_ec, missing_countries_df]
            )

        self.exogenous_shock_on_ec.sort_index(inplace=True)


    def __set_carbon_tax_inclusion(
            self,
            input_data: 'InputData',
            scenario_type: str
            ) -> None:
        """
        'Fuel Carbon Price Inclusion' (v361)
        CPAT Excel: D195:AB204

        'Sectors Carbon Price Inclusion' (v361)
        CPAT Excel: D206:AB223

        Filters scenario for both Carbon Price Inclusion tables.
        Data is loaded in get_sectors_carbon_price_inclusion
        and get_fuel_carbon_price_inclusion

        TODO: allow to modify from dashboard inputs:
        if values[scenario_type] != inputs -> overwrite

        returns dims (s), scenario_type
        """
        self.fuel_carbon_price_inclusion = (
            input_data.fuel_carbon_price_inclusion.loc[:, [scenario_type]]
        )
        self.sectors_carbon_price_inclusion = (
            input_data.sectors_carbon_price_inclusion.loc[:, [scenario_type]]
        )


    def __set_apply_tax_s(
            self,
            scenario_type: str
            ) -> pd.DataFrame:
        """
        'Apply tax' for Sectors (v361)
        CPAT Excel: baseline E1874:E1892, scenario E6864:E6882

        returns dims (s), scenario_type
        """
        apply_tax_s = self.sectors_carbon_price_inclusion.copy()

        if not EXPLICIT_PRICING_VARIANT[scenario_type]:
            apply_tax_s.loc[:, scenario_type] = False
            if scenario_type in [c.FEEBATES, c.POWER_FEEBATE]:
                apply_tax_s.loc[c.POW, scenario_type] = True

        self.apply_tax_s = apply_tax_s.copy()


    def __set_apply_tax_f(
            self,
            scenario_type: str
            ) -> pd.DataFrame:
        """
        'Apply tax' for Fuels (v361)
        CPAT Excel: baseline E1863:E1872, scenario E6853:E6862

        returns dims (f), scenario_type
        """
        apply_tax_f = self.fuel_carbon_price_inclusion.copy()

        if not EXPLICIT_PRICING_VARIANT[scenario_type]:
            if scenario_type in [c.FEEBATES, c.POWER_FEEBATE]:
                # TODO: this looks suspicious in Excel, I think it is a BUG
                apply_tax_f.loc[:, scenario_type] = True
            else:
                apply_tax_f.loc[:, scenario_type] = False
        self.apply_tax_f = apply_tax_f.copy()


    def __get_apply_tax_s_extended(
            self,
            selected_countries: list[str]
            ) -> pd.DataFrame:
        """
        Extends apply_tax_s by a country dim.
        """
        apply_tax_s = self.apply_tax_s.reset_index()
        apply_tax_s_per_country = []
        for country_code in selected_countries:
            apply_tax_s_per_country.append(
                apply_tax_s.assign(
                    **{c.COUNTRY_CODE: country_code}
                ).set_index([c.COUNTRY_CODE, c.SECTOR_CODE])
            )
        apply_tax_s = pd.concat(apply_tax_s_per_country).sort_index()

        self.apply_tax_s_extended = apply_tax_s.copy()


    def __get_apply_tax_f_extended(
            self,
            selected_countries: list[str]
            ) -> None:
        """
        Extends apply_tax_f by a country dim.
        """
        apply_tax_f = self.apply_tax_f.reset_index()
        apply_tax_f_per_country = []
        for country_code in selected_countries:
            apply_tax_f_per_country.append(
                apply_tax_f.assign(
                    **{c.COUNTRY_CODE: country_code}
                ).set_index([c.COUNTRY_CODE, c.FUEL_CODE])
            )
        apply_tax_f = pd.concat(apply_tax_f_per_country).sort_index()

        self.apply_tax_f_extended = apply_tax_f.copy()

    def __get_existing_ets_extended(
            self,
            existing_ets: pd.DataFrame
            ) -> pd.DataFrame:
        """
        Helper function

        Extends index from [ind, pow, res, trs] to p_cov_s index
        """
        mapping = {
            'trs': [c.ROD, c.RAL, c.AVI, c.NAV],
            c.RES : [c.FOO, c.SRV],
            c.IND: [c.MCH, c.IRN, c.NFM, c.MAC, c. CEM, c.OMN, c.CST, c.FTR],
        }
        new_rows = []
        for (country, sector_group), row in existing_ets.iterrows():
            if sector_group in mapping:
                for sector in mapping[sector_group]:
                    new_rows.append((country, sector, row.copy()))

        existing_ets_expanded = pd.DataFrame(
            [row for _, _, row in new_rows],  # Extract row data
            index=pd.MultiIndex.from_tuples(
                [(country, sector) for country, sector, _ in new_rows],
                names=[c.COUNTRY_CODE, c.SECTOR_CODE]
            )
        )

        # # also add oen with 0s only
        unique_col1 = existing_ets.index.get_level_values(c.COUNTRY_CODE).unique()
        existing_ets_oen = pd.DataFrame(
            0.0,  # Fill all columns with 0.0
            index=pd.MultiIndex.from_product(
                [unique_col1, [c.OEN]],
                names=[c.COUNTRY_CODE, c.SECTOR_CODE]
            ),
            columns=existing_ets.columns
        )
        existing_ets = pd.concat(
            [existing_ets, existing_ets_expanded, existing_ets_oen]
        ).sort_index()
        return existing_ets


    def __get_p_cov_s(
            self,
            scenario_type: str,
            ct_tax_complimentary_to_ets: bool,
            existing_ets: pd.DataFrame, # ExistingETS.existing_ets_s_nat
            cp_intro: int,
            year: int,
            exempt_phaseout: bool,
            year_pha: int,
            exempt_phaseout_period: int
            ) -> pd.DataFrame:
        """
        'Price-based policies coverage (% of the new price), by sector' (v361)
        CPAT Excel: baseline 1874:1892, scenario 6864:6882

        returns dims (c, s), t
        """
        if ct_tax_complimentary_to_ets and scenario_type != c.BASELINE:
            p_cov_s = 1 - self.__get_existing_ets_extended(existing_ets)
            if scenario_type != c.BASELINE and year >= cp_intro:
                year_apply = (year + 1 - year_pha) / exempt_phaseout_period
                p_cov_s.loc[pd.IndexSlice[:, c.OEN], str(year)] = (
                    self.apply_tax_s_extended.loc[pd.IndexSlice[:, c.OEN], scenario_type]
                    .apply(
                        lambda x: 1 if x
                        else (
                            max(min(year_apply, 1), 0) if exempt_phaseout
                            else 0
                        )
                    )
                )
            else:
                p_cov_s.loc[pd.IndexSlice[:, c.OEN], str(year)] = 0
            return p_cov_s

        p_cov_s = self.apply_tax_s_extended.copy()

        if scenario_type != c.BASELINE and year >= cp_intro:
            year_apply = (year + 1 - year_pha) / exempt_phaseout_period
            p_cov_s[scenario_type] = (
                self.apply_tax_s_extended[scenario_type]
                .apply(
                    lambda x: 1 if x
                    else (
                        max(min(year_apply, 1), 0) if exempt_phaseout
                        else 0
                    )
                )
            )
        else:
            p_cov_s.loc[:, scenario_type] = False

        return p_cov_s.rename(columns={scenario_type: str(year)})


    def __get_p_cov_f(
            self,
            scenario_type: str,
            cp_intro: int,
            year: int,
            exempt_phaseout: bool,
            year_pha: int,
            exempt_phaseout_period: int
            ) -> pd.DataFrame:
        """
        'Price-based policies coverage (% of the new price), by fuel' (v361)
        CPAT Excel: baseline 1863:1872, scenario 6853:6862

        returns dims (c, f), t
        """
        # TODO: bio, ren, jfu has a BUG in Excel, adjust this after it is resolved
        p_cov_f = self.apply_tax_f_extended.copy()

        if scenario_type != c.BASELINE and year >= cp_intro:
            year_apply = (year + 1 - year_pha) / exempt_phaseout_period
            p_cov_f[scenario_type] = (
                self.apply_tax_f_extended[scenario_type]
                .apply(
                    lambda x: 1 if x
                    else (
                        max(min(year_apply, 1), 0) if exempt_phaseout
                        else 0
                    )
                )
            )
        else:
            p_cov_f.loc[:, scenario_type] = False

        return p_cov_f.rename(columns={scenario_type: str(year)})


    def __get_p_based_policies_cov_s_f(
            self,
            year: int
            ) -> pd.DataFrame:
        """
        'Price-based policies coverage (% total sectoral emissions), by fuel and sector' (v361)
        CPAT Excel: baseline 1895:1906, scenario 6885:6896

        returns dims (c, s, f), t
        """
        coa_nga_oop = [c.COA, c.NGA, c.OOP]
        coa_nga = [c.COA, c.NGA]
        pow_res = [c.POW, c.RES]

        p_cov_f = self.p_cov_f[[str(year)]]
        p_cov_s = self.p_cov_s[[str(year)]]

        # take only f values for these fuels and extend index by sector 'all':
        only_fuel_values = [c.BIO, c.GSO, c.DIE, c.LPG, c.KER]
        p_based_policies_cov_s_f = p_cov_f[
            p_cov_f.index.get_level_values(c.FUEL_CODE).isin(only_fuel_values)
        ]
        p_based_policies_cov_s_f.index = pd.MultiIndex.from_arrays(
            [
                p_based_policies_cov_s_f.index.get_level_values(c.COUNTRY_CODE),
                [c.ALL] * len(p_based_policies_cov_s_f),
                p_based_policies_cov_s_f.index.get_level_values(c.FUEL_CODE)
            ],
            names=c.ID_COL_NAMES
        )

        # selects for [coa, nga, oop] and [coa, nga]
        p_cov_f_coa_nga_oop = p_cov_f[
            p_cov_f.index.get_level_values(c.FUEL_CODE).isin(coa_nga_oop)
        ]
        p_cov_f_coa_nga = p_cov_f_coa_nga_oop[
            p_cov_f_coa_nga_oop.index.get_level_values(c.FUEL_CODE).isin(coa_nga)
        ]

        ### ind for coa, nga, oop
        # max from ind sectors for coa, nga, gso in ind
        ind_sectors = [c.FOO, c.SRV, c.MCH, c.IRN, c.NFM, c.MAC, c.CEM, c.OMN, c.CST, c.FTR, c.OEN]
        p_cov_s_ind = (
            p_cov_s[p_cov_s.index.get_level_values(c.SECTOR_CODE).isin(ind_sectors)]
            .groupby(c.COUNTRY_CODE)
            .max()
        )
        p_cov_s_ind = (
            p_cov_s_ind.set_index(
                pd.Index([c.IND] * len(p_cov_s_ind), name=c.SECTOR_CODE), append=True
            )
        )
        # extend index by coa, nga, gso
        p_cov_s_ind = p_cov_s_ind.reindex(p_cov_s_ind.index.repeat(len(coa_nga_oop)))
        p_cov_s_ind.index = pd.MultiIndex.from_tuples(
            [
                (country, sector, fuel)
                for (country, sector) in p_cov_s_ind.index[::len(coa_nga_oop)]
                for fuel in coa_nga_oop
            ],
            names=c.ID_COL_NAMES
        )
        # extend index by a constant c.IND
        p_cov_f_coa_nga_oop.index = pd.MultiIndex.from_arrays(
            [
                p_cov_f_coa_nga_oop.index.get_level_values(c.COUNTRY_CODE),
                [c.IND] * len(p_cov_f_coa_nga_oop),
                p_cov_f_coa_nga_oop.index.get_level_values(c.FUEL_CODE)
            ],
            names=c.ID_COL_NAMES
        )

        ### pow, res for coa, nga:
        p_cov_s_pow_res = (
            p_cov_s[p_cov_s.index.get_level_values(c.SECTOR_CODE).isin(pow_res)]
        )
        # extend index by coa, nga
        p_cov_s_pow_res = p_cov_s_pow_res.reindex(p_cov_s_pow_res.index.repeat(len(coa_nga)))
        p_cov_s_pow_res.index = pd.MultiIndex.from_tuples(
            [
                (country, sector, fuel)
                for (country, sector) in p_cov_s_pow_res.index[::len(coa_nga)]
                for fuel in coa_nga
            ],
            names=c.ID_COL_NAMES
        )
        # extend index by pow, res
        p_cov_f_coa_nga = p_cov_f_coa_nga.reindex(p_cov_f_coa_nga.index.repeat(len(pow_res)))
        p_cov_f_coa_nga.index = pd.MultiIndex.from_tuples(
            [
                (country, sector, fuel)
                for (country, fuel) in p_cov_f_coa_nga.index[::len(pow_res)]
                for sector in pow_res
            ],
            names=c.ID_COL_NAMES
        )

        # append all:
        p_based_policies_cov_s_f = pd.concat([
            p_based_policies_cov_s_f,
            p_cov_f_coa_nga_oop * p_cov_s_ind,
            p_cov_s_pow_res * p_cov_f_coa_nga
        ]).sort_index()

        return p_based_policies_cov_s_f


    def __set_cp_trajectory(
            self,
            simulation_years: tuple[int, int],
            scenario_type: str,
            d: DashboardInputsDict,
            deflator: pd.DataFrame
            ) -> pd.DataFrame: # (just t)
        """
        'Carbon price trajectory used' (v361)
        CPAT Excel: baseline 6188 -> 1855, scenario 11178 -> 6845

        cp_trajectory dims <simulation_years[0], simulation_years[1]>
        """
        # TODO: move tests, clean inputs

        # template with 0s
        years = np.arange(simulation_years[0], simulation_years[1] + 1)
        columns = years.astype(str)
        countries = list(deflator.index.get_level_values(c.COUNTRY_CODE).unique())
        cp_trajectory = pd.DataFrame([np.zeros(len(columns), dtype=float)], columns=columns)

        if scenario_type == c.BASELINE:
            # extend by country
            cp_trajectory = cp_trajectory.reindex(countries, fill_value=0.0)
            cp_trajectory.index.name = c.COUNTRY_CODE
            # return all 0s
            return cp_trajectory


        # exponential trajectory
        if d['cp_trajectory_type'] == 'exp':
            cp_trajectory.values[0, years >= d['start_cp_year']] = (
                # cp_trajectory(t) = start_cp * ct_exp_rate ** (t - start_cp_year)
                d['start_cp']
                * d['ct_exp_rate'] ** np.maximum(
                    0, years[years >= d['start_cp_year']] - d['start_cp_year']
                )
            )
            # after target year we keep trajectory constant based on value at target_cp_year
            cp_trajectory.values[0, years > d['target_cp_year']] = (
                cp_trajectory.values[0, np.where(years == d['target_cp_year'])[0][0]]
            )
            cp_trajectory = (
                cp_trajectory.loc[cp_trajectory.index.repeat(len(countries))]
                .reset_index(drop=True)
            )
            cp_trajectory.index = pd.Index(countries, name=c.COUNTRY_CODE)

        else: # all linear trajectories
            tax_inc_value = (
                (d['target_cp'] - d['start_cp'])
                / (d['target_cp_year'] - d['start_cp_year'])
            )
            target_range = (years >= d['start_cp_year']) & (years <= d['target_cp_year'])
            post_target_range = years > d['target_cp_year']

            cp_trajectory.loc[:, years == d['start_cp_year']] = d['start_cp']

            # Between start_cp_year and target_cp_year
            cp_trajectory.loc[:, target_range] = (
                d['start_cp'] + (years[target_range] - d['start_cp_year']) * tax_inc_value
            )

            if d['cp_trajectory_type'] == 'per':
                per_tax_inc = d['ct_exp_rate']/(d['target_cp'] - d['ct_exp_rate'])
                # After target_cp_year, apply the percentage increase
                for y in years[post_target_range]:
                    cp_trajectory.loc[:, str(y)] += (
                        cp_trajectory.loc[:, str(y - 1)] * (1 + per_tax_inc)
                    )
            elif d['cp_trajectory_type'] == 'lin':
                # After target_cp_year, continue linear increment
                for y in years[post_target_range]:
                    cp_trajectory.loc[:, str(y)] += cp_trajectory.loc[:, str(y - 1)] + tax_inc_value
            elif d['cp_trajectory_type'] == 'cst':
                # After target_cp_year, carry forward previous value
                cp_trajectory.loc[:, post_target_range] = (
                    cp_trajectory[str(d['target_cp_year'])].iloc[0]
                )

            # d['cp_trajectory_type'] != 'exp' included in else: for all linear trajectories
            if d['tax_pathway'] == 'rea':
                cp_trajectory = deflator.iloc[:, 1:] * cp_trajectory.iloc[0]
            else:
                # just extend the index
                cp_trajectory = (
                    cp_trajectory.loc[cp_trajectory.index.repeat(len(countries))]
                    .reset_index(drop=True)
                )
                cp_trajectory.index = pd.Index(countries, name=c.COUNTRY_CODE)

        if scenario_type == c.ETS:
            cp_trajectory = cp_trajectory * d['ets_adj']

        return cp_trajectory


def get_fuel_carbon_price_inclusion() -> pd.DataFrame:
    """
    [Data Loading]
    'Fuel Carbon Price Inclusion' (v407)
    CPAT Excel: D205:AC214

    Manual changes to 'Fuel Carbon Price Inclusion':
    - fuels mapped to fuel codes and located in FuelCode column
    - all 'No' changed to FALSE and 'Yes' to TRUE

    Scenario filtered later in __set_carbon_tax_inclusion

    returns dims (f), all scenarios
    """
    fuel_carbon_price_inclusion: pd.DataFrame = (
        pd.read_pickle(f'{PKL_DATA_PATH}/fuel_carbon_price_inclusion.pkl.bz2', compression="bz2")
    )
    fuel_carbon_price_inclusion.set_index(c.FUEL_CODE, inplace=True)

    assert (
            set(fuel_carbon_price_inclusion.columns) == set(c.ALL_SCENARIOS)
        ), "Invalid scenario in Fuel Carbon Price Inclusion"

    return fuel_carbon_price_inclusion


def get_sectors_carbon_price_inclusion() -> pd.DataFrame:
    """
    [Data Loading]
    'Sectors Carbon Price Inclusion' (v407)
    CPAT Excel: D216:AC233

    Manual changes to 'Sectors Carbon Price Inclusion':
    - sectors mapped to sector codes and located in SectorCode column
    - all 'No' changed to FALSE and 'Yes' to TRUE
    - 'ind' row with all FALSE values added to match apply tax dims
        (in Excel 'ind' is blank)

    Scenario filtered later in __set_carbon_tax_inclusion

    returns dims (s), all scenarios
    """
    sectors_carbon_price_inclusion: pd.DataFrame = (
        pd.read_pickle(f'{PKL_DATA_PATH}/sectors_carbon_price_inclusion.pkl.bz2', compression="bz2")
    )
    sectors_carbon_price_inclusion.set_index(c.SECTOR_CODE, inplace=True)

    assert (
            set(sectors_carbon_price_inclusion.columns) == set(c.ALL_SCENARIOS)
        ), "Invalid scenario in Sectors Carbon Price Inclusion"

    return sectors_carbon_price_inclusion


def load_ghg_adjustments(selected_countries: list[str]) -> pd.DataFrame:
    """
    [Data Loading]
    'GHG adjustments (manually inferred based on observed emissions)' (v407)
    CPAT Excel:  GHGs tab, rows 2078:2265

    Manual adjustments:
    - 'Code' renamed to 'CountryCode'
    - 'EU', 'Region', 'Country', 'Notes', '#', 'GHG MATCH' columns omitted,
        both 'Targets for energy CO2' columns omitted
    - 'Exogenous adjustments' columns named just as '2023', '2024'

    Only 'Exogenous adjustments' loaded!
    Add other columns if needed in the future!

    returns dims (c), t in <2023, 2024>
    """
    # TODO: pkl
    ghg_adjustments: pd.DataFrame = pd.read_csv(f'{DATA_PATH}/ghg_adjustments.csv')

    ghg_adjustments = ghg_adjustments[
        ghg_adjustments[c.COUNTRY_CODE].isin(selected_countries)
    ].set_index(c.COUNTRY_CODE).fillna(0.0)

    return ghg_adjustments
