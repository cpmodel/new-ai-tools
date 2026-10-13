from typing import TYPE_CHECKING
import pandas as pd
import numpy as np

from cpat_model.inputs.dashboard_inputs import DashboardInputsDict

import cpat_model.constants as c
from config import DATA_PATH

if TYPE_CHECKING:
    from cpat_model.inputs.input_data import InputData

MPP_DATA_FILE_NAME = 'mpp'
IC_DATA_FILE_NAME = 'ic'
LCOE_DATA_FILE_NAME = 'lcoe_tmp'

POW_FUELS = [c.COA, c.NGA, c.OIL, c.NUC, c.WND, c.SOL, c.HYD, c.ORE, c.BIO]
OIL_FUELS = [c.OOP, c.GSO, c.DIE, c.KER, c.LPG, c.JFU]

# columns
MPP_COLS = { # TODO: change these inputs into just 1 table?
    'lifetime': 'TotalLifetime',
    'capacity_f': 'CapacityFactor',
    'termal_eff': 'Efficiency.NCV',
    'var_om_cost': 'OpexVariable.USD_kWh'
}

class PowData:
    """
    Input Data for Power model
    CPAT Excel: A section (v407)
    """
    pow_index: pd.MultiIndex
    elec_for_power: pd.DataFrame
    total_lifetime: pd.DataFrame
    capacity_factor: pd.DataFrame
    capacity: pd.DataFrame
    thermal_efficiency: pd.DataFrame
    var_om_cost: pd.DataFrame

    def __init__(
            self,
            selected_countries: list[str],
            d: DashboardInputsDict,
            input_data: 'InputData'
            ) -> None:
        self.__init_pow_index_frame(selected_countries)
        self.__init_elec_for_power(input_data.ec_input)
        self.__init_pow_for_power(input_data.ec_input)
        self.__init_total_lifetime(input_data)
        self.__init_capacity_factor(d, input_data)
        self.__init_capacity()
        self.__init_thermal_efficiency(d, input_data)
        self.__init_var_om_cost(input_data)


    def __init_elec_for_power(
            self,
            ec_input: pd.DataFrame
            ) -> None:
        """
        'Power Output MWy/y' from 'Energy Balance Data Aggregated' A4 table
        CPAT Excel: M2669:M2678 (v407)

        Includes transformations form:
        'Electricity output (Gwh)' from 'Energy Balance Data For Power' A1 table
        CPAT Excel: M2629:M2643 (v407)

        'All oil products' row as 'oil', oil components row dropped

        elec_for_power dims (c, f), 'Value' column
        """
        elec_for_power = (
            ec_input[ec_input.index.get_level_values(c.SECTOR_CODE) == c.ELEC]
            .droplevel(c.SECTOR_CODE)
        )
        elec_for_power[c.OIL] = elec_for_power[OIL_FUELS].sum(axis=1)
        elec_for_power = elec_for_power[POW_FUELS]
        self.elec_for_power = (
            elec_for_power
            .stack()
            .rename('Value')
            .rename_axis(index=[elec_for_power.index.name, c.FUEL_CODE])
            .to_frame()
        ) * c.GWH_TO_MWY
        # TODO: assert index with pow_index_frame?


    def __init_pow_for_power(
            self,
            ec_input: pd.DataFrame
            ) -> None:
        """
        'Primary Energy MWy/y ' from 'Energy Balance Data Aggregated' A4 table
        CPAT Excel: L2669:L2678 (v407)

        Includes transformations form:
        'Primary Energy (ktoe)' from 'Energy Balance Data For Power' A1 table
        CPAT Excel: L2629:L2643 (v407)

        'All oil products' row as 'oil', oil components row dropped

        __pow_for_power dims (c, f), 'Value' column
        """
        pow_for_power = (
            ec_input[ec_input.index.get_level_values(c.SECTOR_CODE) == c.POW]
            .droplevel(c.SECTOR_CODE)
        )
        pow_for_power[c.OIL] = pow_for_power[OIL_FUELS].sum(axis=1)
        pow_for_power = pow_for_power[POW_FUELS]
        self.__pow_for_power = (
            pow_for_power
            .stack()
            .rename('Value')
            .rename_axis(index=[pow_for_power.index.name, c.FUEL_CODE])
            .to_frame()
        ) * c.KTOE_TO_MWY
        # TODO: assert index with pow_index_frame?


    def __init_pow_index_frame(
            self,
            selected_countries: list[str],
            ) -> None:
        """
        Helper

        __pow_index pd.DataFrame with c, f as columns
            for f in POW_FUELS
        """
        self.pow_index = pd.MultiIndex.from_product(
            [selected_countries, POW_FUELS],
            names=[c.COUNTRY_CODE, c.FUEL_CODE]
        )
        # build frame:
        self.__pow_index_frame = pd.DataFrame(index=self.pow_index).reset_index()


    def __init_total_lifetime(
            self,
            input_data: 'InputData'
            ) -> None:
        """
        'TotalLifetime' from A2 table
        CPAT Excel: T2647:T2656 (v407)

        total_lifetime dims (c, f), 'TotalLifeline'
            f in POW_FUELS
        """
        total_lifetime = input_data.mpp[[c.COUNTRY_CODE, c.FUEL_CODE, 'TotalLifetime']]

        wld = (
            total_lifetime[total_lifetime[c.COUNTRY_CODE] == 'WLD']
            .rename(columns={'TotalLifetime': 'wld_value'})
            [[c.FUEL_CODE, 'wld_value']]
        )
        total_lifetime = self.__pow_index_frame.merge(
            total_lifetime[total_lifetime[c.COUNTRY_CODE] != 'WLD'],
            on=[c.COUNTRY_CODE, c.FUEL_CODE],
            how='left'
        )

        total_lifetime = total_lifetime.merge(wld, on=c.FUEL_CODE, how='left')
        total_lifetime['TotalLifetime'] = (
            total_lifetime['TotalLifetime']
            .fillna(total_lifetime['wld_value'])
        )
        self.total_lifetime = (
            total_lifetime
            .drop(columns='wld_value')
            .set_index([c.COUNTRY_CODE, c.FUEL_CODE])
        )


    def __init_capacity_factor(
            self,
            d: DashboardInputsDict,
            input_data: 'InputData'
            ) -> None:
        """
        'Capacity Factor Used' from A4 table
        CPAT Excel: W2669:W2678 (v407)

        capacity_factor dims (c, f), 'CapacityFactor'
            f in POW_FUELS
        """
        # TODO: split the code and test
        # 'Capacity Factor Studies'  from A4 table R2669:R2678 -> N2647:N2656 (v407)
        # TODO: unify this and total_lifetime logic
        capacity_factor_studies = input_data.mpp[[c.COUNTRY_CODE, c.FUEL_CODE, 'CapacityFactor']]

        wld = (
            capacity_factor_studies[capacity_factor_studies[c.COUNTRY_CODE] == 'WLD']
            .rename(columns={'CapacityFactor': 'wld_value'})
            [[c.FUEL_CODE, 'wld_value']]
        )
        capacity_factor_studies = self.__pow_index_frame.merge(
            capacity_factor_studies[capacity_factor_studies[c.COUNTRY_CODE] != 'WLD'],
            on=[c.COUNTRY_CODE, c.FUEL_CODE],
            how='left'
        )

        capacity_factor_studies = capacity_factor_studies.merge(wld, on=c.FUEL_CODE, how='left')
        capacity_factor_studies['CapacityFactor'] = (
            capacity_factor_studies['CapacityFactor']
            .fillna(capacity_factor_studies['wld_value'])
        )
        capacity_factor_studies = (
            capacity_factor_studies
            .drop(columns='wld_value')
            .set_index([c.COUNTRY_CODE, c.FUEL_CODE])
        )

        # 'Capacity Factor Actual'  from A4 table Q2669:Q2678 (v407)
        # should rise AssertionError if index does not match
        pd.testing.assert_index_equal(
            input_data.ic_base_year.index.sort_values(),
            self.elec_for_power.index.sort_values(),
            obj='Capacity Actual (MW) and Power Output MWy/y index (order ignored)'
        )
        # should rise AssertionError if cpolumns does not match
        pd.testing.assert_index_equal(
            input_data.ic_base_year.columns, self.elec_for_power.columns
        )
        capacity_factor_actual = (
            (self.elec_for_power / input_data.ic_base_year)
            .rename(columns={'Value': 'CapacityFactor'})
            .replace([np.inf, -np.inf], 0.0)
            .fillna(0.0)
        )

        # use actual values if given range, otherwise use studies values as fallback
        # should rise AssertionError if index does not match
        pd.testing.assert_index_equal(
            capacity_factor_actual.index.sort_values(),
            capacity_factor_studies.index.sort_values(),
            obj='Actual and Studies capacity factor dfs index (order ignored)'
        )
        fuel = capacity_factor_actual.index.get_level_values(c.FUEL_CODE)

        mask = (
            (capacity_factor_actual['CapacityFactor'] > d['cf_override_if_above'])
            | (
                fuel.isin([c.WND, c.SOL])
                & (capacity_factor_actual['CapacityFactor'] < d['cf_override_if_below_wnd_sol'])
            )
            | (
                ~fuel.isin([c.WND, c.SOL])
                & (capacity_factor_actual['CapacityFactor'] < d['cf_override_if_below_fos_oth'])
            )
        )

        self.capacity_factor = capacity_factor_actual.assign(
            **{
                'CapacityFactor': (
                    capacity_factor_actual['CapacityFactor']
                    .where(~mask, capacity_factor_studies['CapacityFactor'])
                )
            }
        )


    def __init_capacity(self) -> None:
        """
        'Capacity (MW) Used' from A4 table
        CPAT Excel: Y2669:Y2678 (v407)

        capacity dims (c, f), 'Value'
            f in POW_FUELS
        """
        self.capacity = self.elec_for_power.copy()
        self.capacity.iloc[:, -1] /= self.capacity_factor.iloc[:, -1]


    def __init_thermal_efficiency(
            self,
            d: DashboardInputsDict,
            input_data: 'InputData'
            ) -> None:
        """
        'Thermal Efficiency Used' from A4 table
        CPAT Excel: V2669:V2678 (v407)

        thermal_efficiency dims (c, f), 'Efficiency.NCV'
            f in POW_FUELS
        """
        # TODO: split to methods and add unit tests
        # TODO assertion on index?
        # 'Thermal Efficiency Studies' R2669:R2678 -> 'Efficiency.NCV' P2647:P2656 (v407)
        # TODO: unify this and total_lifetime logic
        thermal_efficiency_studies = input_data.mpp[[c.COUNTRY_CODE, c.FUEL_CODE, 'Efficiency.NCV']]

        wld = (
            thermal_efficiency_studies[thermal_efficiency_studies[c.COUNTRY_CODE] == 'WLD']
            .rename(columns={'Efficiency.NCV': 'wld_value'})
            [[c.FUEL_CODE, 'wld_value']]
        )
        thermal_efficiency_studies = self.__pow_index_frame.merge(
            thermal_efficiency_studies[thermal_efficiency_studies[c.COUNTRY_CODE] != 'WLD'],
            on=[c.COUNTRY_CODE, c.FUEL_CODE],
            how='left'
        )

        thermal_efficiency_studies = (
            thermal_efficiency_studies
            .merge(wld, on=c.FUEL_CODE, how='left')
        )
        thermal_efficiency_studies['Efficiency.NCV'] = (
            thermal_efficiency_studies['Efficiency.NCV']
            .fillna(thermal_efficiency_studies['wld_value'])
        )
        thermal_efficiency_studies = (
            thermal_efficiency_studies
            .drop(columns='wld_value')
            .set_index([c.COUNTRY_CODE, c.FUEL_CODE])
        )

        # additionally set c.WND, c.SOL, c.HYD, c.ORE to 1.0
        mask = (
            thermal_efficiency_studies.index
            .get_level_values(c.FUEL_CODE).isin([c.WND, c.SOL, c.HYD, c.ORE])
        )
        thermal_efficiency_studies.loc[mask, 'Efficiency.NCV'] = 1.0

        # actual
        thermal_efficiency_actual = (
            (self.elec_for_power/ self.__pow_for_power)
            .rename(columns={'Value': 'Efficiency.NCV'})
            .replace([np.inf, -np.inf], 0.0)
            .fillna(0.0)
        )

        mask = thermal_efficiency_actual['Efficiency.NCV'] < d['minimum_thermal_efficiency']

        self.thermal_efficiency = thermal_efficiency_actual.assign(
            **{
                'Efficiency.NCV': (
                    thermal_efficiency_actual['Efficiency.NCV']
                    .where(~mask, thermal_efficiency_studies['Efficiency.NCV'])
                )
            }
        )

    def __init_var_om_cost(
            self,
            input_data:'InputData'
            ) -> None:
        """
        'OpexVariable.USD_kWh' from A2 table
        CPAT Excel: R2647:R2656 (v407)

        total_lifetime dims (c, f), 'OpexVariable.USD_kWh'
            f in POW_FUELS
        """
        var_om_cost = input_data.mpp[[c.COUNTRY_CODE, c.FUEL_CODE, 'OpexVariable.USD_kWh']]

        wld = (
            var_om_cost[var_om_cost[c.COUNTRY_CODE] == 'WLD']
            .rename(columns={'OpexVariable.USD_kWh': 'wld_value'})
            [[c.FUEL_CODE, 'wld_value']]
        )
        var_om_cost = self.__pow_index_frame.merge(
            var_om_cost[var_om_cost[c.COUNTRY_CODE] != 'WLD'],
            on=[c.COUNTRY_CODE, c.FUEL_CODE],
            how='left'
        )

        var_om_cost = var_om_cost.merge(wld, on=c.FUEL_CODE, how='left')
        var_om_cost['OpexVariable.USD_kWh'] = (
            var_om_cost['OpexVariable.USD_kWh']
            .fillna(var_om_cost['wld_value'])
        )
        self.var_om_cost = (
            var_om_cost
            .drop(columns='wld_value')
            .set_index([c.COUNTRY_CODE, c.FUEL_CODE])
        )


def load_mpp(
        selected_countries: list[str]
        ) -> pd.DataFrame:
    """
    [Data Loading]
    Use for 'Main Parameters' A2 table
    CPAT Excel table stats row 2647 (v407)

    Input file preprocessed from raw input table
    in get_mpp_csv function.

    return dims c, f, data columns
    """
    # TODO: pkl
    df: pd.DataFrame = pd.read_csv(f'{DATA_PATH}/{MPP_DATA_FILE_NAME}.csv')
    df = df[df[c.COUNTRY_CODE].isin([*selected_countries, 'WLD'])]
    df = df[df[c.FUEL_CODE].isin(POW_FUELS)]

    return df


def load_ic(
        selected_countries: list[str],
        simulation_years: tuple[int, int]
        ) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    [Data Loading]
    'Installed capacity per country'

    Table from 'Power' tab (v407), rows 355:1635.

    Columns used:
    - 'CountryCode'
    - 'FuelType', renamed manually to 'FuelCode'
    - year columns 2021-2023

    Used for 'Capacity Actual (MW)' from A4 table N2669:N2678 (v407)

    return 
        ic dims (c, f), t in <2021, 2023>
        ic_base_year  (c, f), 'Value'
    """
    # TODO: pkl
    ic: pd.DataFrame = pd.read_csv(f'{DATA_PATH}/{IC_DATA_FILE_NAME}.csv')
    ic = (
        ic[ic[c.COUNTRY_CODE].isin(selected_countries)]
        .set_index([c.COUNTRY_CODE, c.FUEL_CODE])
    )
    ic_base_year = (
        ic[[str(simulation_years[0])]]
        .rename(columns={str(simulation_years[0]): 'Value'})
    )
    return ic, ic_base_year

def load_lcoe_tmp(
        selected_countries: list[str],
        simulation_years: tuple[int, int]
        ) -> pd.DataFrame:
    """
    [Data Loading]
    Temporary
    'Levelised Total Investment Costs (Levelised Fixed Costs + Levelised Variable Costs)'
    E1 table 

    CPAT Excel rows 3107:3116 (v407),
    fixed values for Egypt

    Columns used:
    - 'FuelCode' column added with fuel codes based on 1st column
    - all year columns, starting form 2022.


    return ic dims (c, f), t in <simulation_years[0], simulation_years[1]>
    """
    # TODO: pkl
    lcoe: pd.DataFrame = (
        pd.read_csv(f'{DATA_PATH}/{LCOE_DATA_FILE_NAME}.csv')
        .set_index(c.FUEL_CODE)
    )
    # years
    last_data_year = max(map(int, lcoe.columns))
    if last_data_year < simulation_years[1]:
        for year in range(last_data_year + 1, simulation_years[1] + 1):
            lcoe[str(year)] = lcoe[str(year - 1)]

    lcoe = lcoe[[str(year) for year in range(simulation_years[0], simulation_years[1] + 1)]]

    # add countries
    lcoe = pd.concat(
        [lcoe] * len(selected_countries),
        keys=selected_countries,
        names=[c.COUNTRY_CODE]
    )
    return lcoe
