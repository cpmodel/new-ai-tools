from typing import TYPE_CHECKING, Literal
import pandas as pd

from cpat_model.mappings.mapping import COUNTRY_TO_INCOME_GROUP
from cpat_model.inputs.dashboard_inputs import DashboardInputsDict

import cpat_model.constants as c
from config import DATA_PATH

if TYPE_CHECKING:
    from cpat_model.inputs.input_data import InputData


DATA_FILE_NAME = 'elasticities'
INCOME_GROUPS = ['LIC', 'HIC', 'LMIC', 'UMIC']

ElasticitiesSectorsType = Literal['tra', 'res', 'ind', 'srv']
ElasticitiesType = Literal['el_inc', 'el_dem', 'el_eff', 'eff_imp']

class Elasticities:
    """
    'Elasticities and rates of improvement'

    Calculates 'Elasticity used' from 'Elasticities' tab (v407), used for:
    - el_inc - 'Income elasticities - energy use'
    - el_dem - 'Own-price elasticities of demand - intensive margin (usage of energy-using capital)'
    - el_eff - 'Own-price elasticities of demand - efficiency
        and extenstive margin (fuel economy of energy-using capital and ownership)'
    - eff_imp - 'Autonomous efficiency improvements'

    Values based on income group with 'global' as fallback.
    TODO: Manual elasticities not implemented yet.
    
    """
    # TODO: split to methods and decribe vars in docstrings

    # elasticities with where FuelCode == 'ele':
    ele_e: dict[ElasticitiesType, pd.DataFrame]
    # elasticities filtered by sector, where FuelCode != 'ele':
    e: dict[
        ElasticitiesSectorsType, dict[
            ElasticitiesType, pd.DataFrame
        ]
    ]

    def __init__(
            self,
            selected_countries: list[str],
            d: DashboardInputsDict,
            input_data: 'InputData'
            ) -> None:
        # TODO: split code into methods and test!
        self.ele_e = {}
        self.e = {c.TRA: {}, c.RES: {}, c.SRV: {}, c.IND: {}}
        income = pd.DataFrame()
        own_price_usage = pd.DataFrame()
        own_price_efficiency = pd.DataFrame()
        autonomous_efficiency = pd.DataFrame()

        for country_code in selected_countries:
            c_elasticities = input_data.elasticities_input.copy()
            income_group = COUNTRY_TO_INCOME_GROUP[country_code]
            columns = ['Type', c.SECTOR_CODE, c.FUEL_CODE, 'Standard deviation']
            if income_group not in INCOME_GROUPS:
                income_group = 'Global'
            c_elasticities = c_elasticities[[*columns, income_group]]

            c_elasticities[c.COUNTRY_CODE] = country_code
            c_elasticities.set_index(c.ID_COL_NAMES, inplace=True)
            c_elasticities.rename(columns={income_group: 'Value'}, inplace=True)

            income = pd.concat(
                [
                    income,
                    c_elasticities[c_elasticities['Type'] == 'Income'].drop(columns=['Type'])
                ]
            )
            own_price_usage = pd.concat(
                [
                    own_price_usage,
                    c_elasticities[
                        c_elasticities['Type'] == 'Own-price, usage'
                    ].drop(columns=['Type'])
                ]
            )
            own_price_efficiency = pd.concat(
                [
                    own_price_efficiency,
                    c_elasticities[
                        c_elasticities['Type'] == 'Own-price, efficiency'
                    ].drop(columns=['Type'])
                ]
            )
            autonomous_efficiency = pd.concat(
                [
                    autonomous_efficiency,
                    c_elasticities[
                        c_elasticities['Type'] == 'Autonomous efficiency'
                    ].drop(columns=['Type', 'Standard deviation'])
                ]
            )

        if d['source_income_elasticities'] == 'Manual':
            raise NotImplementedError("Manual income elasticities not yet supported.")
        if d['source_price_elasticities'] == 'Manual':
            raise NotImplementedError("Manual price elasticities not yet supported.")

        adj_mapping = {'VHigh': 2.0, 'High': 1.0, 'Base': 0.0, 'Low': -1.0, 'VLow': -2.0}
        income_adj = adj_mapping[d['elasticities_adjustment_income']]
        price_adj = adj_mapping[d['elasticities_adjustment_price']]

        income['Value'] += income['Standard deviation'] * income_adj
        own_price_usage['Value'] += own_price_usage['Standard deviation'] * price_adj
        own_price_efficiency['Value'] += (
            own_price_efficiency['Standard deviation'] * price_adj
        )

        income.drop(columns=['Standard deviation'], inplace=True)
        own_price_usage.drop(columns=['Standard deviation'], inplace=True)
        own_price_efficiency.drop(columns=['Standard deviation'], inplace=True)

        # only these 2 elasticities are used in power (calculating demand)
        self.ele_e[c.EL_INC] = income[
            income.index.get_level_values(c.FUEL_CODE) == c.ELE
        ]
        self.ele_e[c.EL_EFF] = own_price_efficiency[
            own_price_efficiency.index.get_level_values(c.FUEL_CODE) == c.ELE
        ]

        e_name_to_df = {
            c.EL_INC: income, c.EL_DEM: own_price_usage,
            c.EL_EFF: own_price_efficiency, c.EFF_IMP: autonomous_efficiency
        }
        for sector in [c.TRA, c.RES, c.SRV, c.IND]:
            for elast in [c.EL_INC, c.EL_DEM, c.EL_EFF, c.EFF_IMP]:
                self.e[sector][elast] = e_name_to_df[elast][
                    (e_name_to_df[elast].index.get_level_values(c.FUEL_CODE) != c.ELE)
                    & (e_name_to_df[elast].index.get_level_values(c.SECTOR_CODE) == sector)
                ]

                self.e[sector][elast] = self.e[sector][elast].droplevel(c.SECTOR_CODE)



def load_elasticities() -> pd.DataFrame:
    """
    [Data Loading]

    'Elasticities' tab

    Preprocessed in get_elasticities_csv function from:
    cpat_processing/mitigation/elasticities_format.py
    """
    # TODO: pkl
    df: pd.DataFrame = pd.read_csv(f'{DATA_PATH}/{DATA_FILE_NAME}.csv')

    return df
