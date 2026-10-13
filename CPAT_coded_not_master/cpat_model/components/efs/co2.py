from typing import Literal, TYPE_CHECKING

import pandas as pd
import numpy as np

import cpat_model.constants as c
from config import DATA_PATH
from cpat_model.mappings.mapping import COUNTRY_TO_REGION

if TYPE_CHECKING:
    from cpat_model.inputs.input_data import InputData

GLOBAL_VALUES_DATA_FILE_NAME = 'co2_efs_global'
EF_GHG_FILE_NAME = 'ef_ghg'
CAL_VAL_FILE_NAME = 'cal_val'
AIR_DATA_FILE_NAME = 'air_data' # NetCV and density only

INDEX_PAIRS = (
    [(c.ALL, fuel) for fuel in [c.BIO, c.JFU, c.OOP]]
    + [(c.AVI, c.JFU)]
    + [
        (sector, fuel)
        for sector in [c. IND, c.POW, c.RES, c.ROD]
        for fuel in [c.COA, c.DIE, c.GSO, c.KER, c.LPG, c.NGA, c.OOP]
    ]
)


FUELS_FOR_CALORIFIC_VALUES_AND_DENSITIES = [
    c.COA, c.NGA, c.GSO, c.DIE, c.LPG, c.KER, c.OOP, c.BIO, c.JFU
]
KCAL_KG = 'kcal/kg'
GJ_TONNE = 'GJ/tonne'
TONNE_M3 = 'tonne/m3'
GJ_M3 = 'GJ/m3'
GJ_LITRE = 'GJ/litre'
CONVERSIONS_TO_PHYSICAL_UNITS = 'Conversions To Physical Units'

CONV_BARREL_TO_GJ = 6.12

EfGhgKeys = Literal['Countries', 'Regions', 'World']
CalValAirDataKeys = Literal[
    'CalVal Countries', 'CalVal Regions',
    'AirData density', 'AirData NetCV'
]

class EFsCO2:
    """
    Tables from
    'Emissions factors (EFs) - CO2' section (v361)
    CPAT Excel: section start 906
    """
    # Public attributes:
    ef_global_iea: pd.DataFrame
    iiasa_ef_tco2: pd.DataFrame
    fuel_calorific_values_and_densities: pd.DataFrame
    conv_factor_volume_unit_to_gj: pd.DataFrame
    ef_iiasa_per_gj: pd.DataFrame
    ef_tco2_per_volume_unit: pd.DataFrame
    ef_tco2_per_gj: pd.DataFrame
    ef_tco2_per_ktoe: pd.DataFrame
    co2_adjustment_factor: pd.DataFrame

    def __init__(
            self,
            selected_countries: list[str],
            # Preloaded input data:
            input_data: 'InputData',
            # Dashboard inputs:
            efs_selected: Literal['IIASA', 'IEA']
            ) -> None:
        self.ef_global_iea = self.__get_ef_global_iea(input_data.co2_efs_global)
        self.__assert_ef_global_iea_index()
        self.iiasa_ef_tco2 = self.__get_iiasa_ef_tco2(selected_countries, input_data.ef_ghg)

        self.fuel_calorific_values_and_densities = self.__get_fuel_calorific_values_and_densities(
            selected_countries, input_data.cal_val_and_air_data
        )
        self.conv_factor_volume_unit_to_gj = self.__get_conv_factor_volume_unit_to_gj()
        self.ef_iiasa_per_gj = self.__get_ef_iiasa_per_gj()
        self.ef_tco2_per_volume_unit = self.__get_ef_tco2_per_volume_unit(
            selected_countries, efs_selected
        )
        self.__set_ef_tco2_per_gj()
        self.__set_ef_tco2_per_ktoe()
        self.__set_co2_adjustment_factor()

    def __set_co2_adjustment_factor(self) -> None:
        """
        'Adjustment factor' (v412)
        CPAT Excel: D967

        co2_adjustment_factor dims (c), 'Adjustment factor' 
        """
        self.co2_adjustment_factor = 1.0
        # TODO: implement, should be pd.DataFrame (c) specific


    def __set_ef_tco2_per_ktoe(self) -> None:
        """
        'EF - tCO2/ktoe' (v412)
        CPAT Excel: Q970:Q1002

        ef_tco2_per_ktoe dims (c, s, f), columns: 'EF - tCO2/ktoe'
        """
        # TODO: test
        self.ef_tco2_per_ktoe = (
            self.ef_tco2_per_gj.rename(columns={'EF - tCO2/GJ': 'EF - tCO2/ktoe'})
            * c.KTOE_TO_GJ
        )


    def __set_ef_tco2_per_gj(self) -> None:
        """
        'EF - tCO2/GJ' (v412)
        CPAT Excel: P970:P1002

        ef_tco2_per_gj dims (c, s, f), columns: 'EF - tCO2/GJ'
        """
        # TODO: test
        assert (
            self.ef_tco2_per_volume_unit.index.symmetric_difference(
                self.conv_factor_volume_unit_to_gj.index
            ).empty
        ), "Indexes of ef_tco2_per_volume_unit and conv_factor_volume_unit_to_gj differ"
        self.ef_tco2_per_gj = (
            self.ef_tco2_per_volume_unit.iloc[:, -1]
            / self.conv_factor_volume_unit_to_gj.iloc[:, -1]
        ).to_frame(name='EF - tCO2/GJ')


    @staticmethod
    def __get_ef_global_iea(
            co2_efs_global: pd.DataFrame
            ) -> pd.DataFrame:
        """
        'EF-GlobalIEA' (v361)
        CPAT Excel: M960:M992

        returns dims (s, f), columns: ['EF-GlobalIEA'] 
        """
        ef_physical_units = co2_efs_global[['EF-physical units']]

        fossil_fuels = [c.COA, c.NGA, c.GSO, c.DIE, c.LPG, c.KER, c.OOP]
        sectors = [c.POW , c.ROD, c.RES, c.IND]

        ef_global_iea = ef_physical_units.loc[fossil_fuels, :]
        ef_global_iea_all_sector = ef_physical_units.loc[[c.OOP, c.BIO, c.JFU], :]
        ef_global_iea_jfu_avi = ef_physical_units.loc[[c.JFU], :]

        # extend index by sector for fossil fuels
        extended_index = pd.MultiIndex.from_product(
            [ef_global_iea.index, sectors],
            names=[c.FUEL_CODE, c.SECTOR_CODE]
        )
        ef_global_iea = (
            ef_global_iea
            .reindex(
                ef_global_iea.index.repeat(len(sectors))
            )
        )
        ef_global_iea.index = extended_index
        # change index order
        ef_global_iea = ef_global_iea.swaplevel(c.SECTOR_CODE, c.FUEL_CODE)

        # add sector 'all' for [c.OOP, c.BIO, c.JFU]
        ef_global_iea_all_sector.index = pd.MultiIndex.from_arrays(
            [[c.ALL] * len(ef_global_iea_all_sector), ef_global_iea_all_sector.index],
            names=[c.SECTOR_CODE, c.FUEL_CODE]
        )

        # add jfu avi
        ef_global_iea_jfu_avi.index = pd.MultiIndex.from_arrays(
            [[c.AVI] * len(ef_global_iea_jfu_avi), ef_global_iea_jfu_avi.index],
            names=[c.SECTOR_CODE, c.FUEL_CODE]
        )

        ef_global_iea = (
            pd.concat([ef_global_iea, ef_global_iea_all_sector, ef_global_iea_jfu_avi])
            .rename(columns={'EF-physical units': 'EF-GlobalIEA'})
        )

        return ef_global_iea


    def __assert_ef_global_iea_index(self) -> None:
        """
        Assertion for ef_global_iea index values and index names.
        """
        assert set(self.ef_global_iea.index) == set(INDEX_PAIRS), (
            'ef_global_iea index values do not match expected values.'
        )
        assert self.ef_global_iea.index.names == [c.SECTOR_CODE, c.FUEL_CODE], (
            'ef_global_iea index level names are incorrect.'
        )


    def __get_iiasa_ef_tco2(
            self,
            selected_countries: list[str],
            ef_ghg: dict[EfGhgKeys, pd.DataFrame]
            ) -> pd.DataFrame:
        """
        'IIASA EF tCO2' (v361)
        CPAT Excel: F960:F992

        Manual changes to EF_GHG -> ef_ghg.csv:
        - columns removed: 'Code', 'Series name'
        - columns renamed: 'Region/country' -> 'CountryCode','Atribute 1' -> 'SectorCode',
                            'Atribute 2' -> 'FuelCode', '#VALUE!' -> 'Value'

        returns dims (c, s, f), data coulum
        """
        world_code = 'world'
        column_name = self.ef_global_iea.columns[0]
        iiasa_ef_tco2 = pd.DataFrame()

        skip_indexes = [
            (c.ALL, c.BIO), (c.ROD, c.OOP), (c.ROD, c.KER), (c.ROD, c.LPG),
            (c.ROD, c.COA), (c.POW, c.KER), (c.POW, c.LPG)
        ]

        for country in selected_countries:
            country_lower = country.lower()
            region = ''
            if country in COUNTRY_TO_REGION:
                region = COUNTRY_TO_REGION[country].lower()
            country_df = self.ef_global_iea.copy() # copy (s, f) structure
            country_df.iloc[:, -1] = np.nan # empty value column

            for idx in country_df.index.difference(skip_indexes):
                s, f = idx # sector, fuel
                if (country_lower, s, f) in ef_ghg['Countries'].index:
                    country_df.at[idx, column_name] = (
                        ef_ghg['Countries'].loc[(country_lower, s, f)].item()
                    )
                elif (region, s, f) in ef_ghg['Regions'].index:
                    country_df.at[idx, column_name] = ef_ghg['Regions'].loc[(region, s, f)].item()
                elif (world_code, s, f) in ef_ghg['World'].index:
                    country_df.at[idx, column_name] = ef_ghg['World'].loc[(world_code, s, f)].item()
                else:
                    raise KeyError(f"{country}, {s}, {f} not found in EF_GHG.")

            # skipped indexes
            country_df.at[(c.ALL, c.BIO), column_name] = 0
            country_df.at[(c.ROD, c.OOP), column_name] = country_df.at[(c.ROD, c.DIE), column_name]
            country_df.at[(c.ROD, c.KER), column_name] = country_df.at[(c.RES, c.KER), column_name]
            country_df.at[(c.ROD, c.LPG), column_name] = country_df.at[(c.RES, c.LPG), column_name]
            country_df.at[(c.ROD, c.COA), column_name] = country_df.at[(c.POW, c.COA), column_name]
            country_df.at[(c.POW, c.KER), column_name] = country_df.at[(c.RES, c.KER), column_name]
            country_df.at[(c.POW, c.LPG), column_name] = country_df.at[(c.IND, c.LPG), column_name]

            country_df /= 1e3

            country_df = (
                country_df
                .set_index(
                    pd.Index([country] * len(country_df), name=c.COUNTRY_CODE),
                    append=True
                )
                .reorder_levels(c.ID_COL_NAMES)
                .sort_index()
            )

            if iiasa_ef_tco2.empty:
                iiasa_ef_tco2 = country_df.copy()
            else:
                iiasa_ef_tco2 = pd.concat([iiasa_ef_tco2, country_df])

        iiasa_ef_tco2.rename(columns={column_name: 'IIASA EF tCO2'}, inplace=True)
        return iiasa_ef_tco2


    @staticmethod
    def __get_fuel_calorific_values_and_densities(
            selected_countries: list[str],
            cal_val_and_air_data: dict[CalValAirDataKeys, pd.DataFrame]
            ) -> pd.DataFrame:
        """
        'Fuels calorific values and densities - {Country}' table from v361 Mitigation!B926:L936

        Uses data: 'CalVal' and 'Air_2'.

        Changes in 'Fuels calorific values and densities - {Country}':
        - 'Sector' column not used - all fuels are unique, makes the table easier to handle
        - 'ren' not used (in Excel version it is a blank placeholder)
        - skipped columns:
            'Desired Physical Units in EF',
            'Conversions To Physical Units - Chosen',
            'l/ktoe', 'PK/kl'

        returns dims (c, f), data columns
        """
        country_dfs = []
        for country in selected_countries:
            region = ''
            country_lower = country.lower()
            if country in COUNTRY_TO_REGION:
                region = COUNTRY_TO_REGION[country].lower()

            # kcal/kg column
            country_df = pd.DataFrame({
                c.FUEL_CODE: FUELS_FOR_CALORIFIC_VALUES_AND_DENSITIES,
                KCAL_KG: [0.0] * len(FUELS_FOR_CALORIFIC_VALUES_AND_DENSITIES)
            }).set_index(c.FUEL_CODE)

            for f in [c.COA, c.NGA, c.OOP, c.BIO]:
                if (country_lower, c.POW, f) in cal_val_and_air_data['CalVal Countries'].index:
                    country_df.at[(f), KCAL_KG] = (
                        cal_val_and_air_data['CalVal Countries'].loc[(country_lower, c.POW, f)].item()
                    )
                elif (region, c.POW, f) in cal_val_and_air_data['CalVal Regions'].index:
                    country_df.at[(f), KCAL_KG] = (
                        cal_val_and_air_data['CalVal Regions'].loc[(region, c.POW, f)].item()
                    )
                else:
                    raise KeyError(f"{country}, {f} not found in CalVal.")

            for f in [c.LPG, c.KER]:
                if (country_lower, c.RES, f) in cal_val_and_air_data['CalVal Countries'].index:
                    country_df.at[(f), KCAL_KG] = (
                        cal_val_and_air_data['CalVal Countries'].loc[(country_lower, c.RES, f)].item()
                    )
                elif (region, c.RES, f) in cal_val_and_air_data['CalVal Regions'].index:
                    country_df.at[(f), KCAL_KG] = (
                        cal_val_and_air_data['CalVal Regions'].loc[(region, c.RES, f)].item()
                    )
                else:
                    raise KeyError(f"{country}, {f} not found in CalVal.")

            for f in [c.GSO, c.DIE, c.JFU]:
                if (country_lower, f) in cal_val_and_air_data['AirData NetCV'].index:
                    country_df.at[(f), KCAL_KG] = (
                        cal_val_and_air_data['AirData NetCV'].loc[(country_lower, f)].item()
                    )
                else:
                    country_df.at[(f), KCAL_KG] = (
                        cal_val_and_air_data['AirData NetCV'].loc[(c.ALL, f)].item()
                    )

            # GJ/tonne column
            country_df[GJ_TONNE] = country_df[KCAL_KG] * 4.19 / 1e3

            # tonne/m3 column
            country_df[TONNE_M3] = np.nan
            for f in [c.GSO, c.DIE, c.LPG, c.KER, c.OOP, c.JFU]:
                if (country_lower, f) in cal_val_and_air_data['AirData density'].index:
                    country_df.at[(f), TONNE_M3] = (
                        cal_val_and_air_data['AirData density'].loc[(country_lower, f)].item()
                    )
                else:
                    country_df.at[(f), TONNE_M3] = (
                        cal_val_and_air_data['AirData density'].loc[(c.ALL, f)].item()
                    )

            country_df = (
                country_df
                .set_index(
                    pd.Index([country] * len(country_df), name=c.COUNTRY_CODE),
                    append=True
                )
                .reorder_levels([c.COUNTRY_CODE, c.FUEL_CODE])
                .sort_index()
            )

            country_dfs.append(country_df.copy())

        # results df:
        df = pd.concat(country_dfs)

        # rest of the columns:
        df[GJ_M3] = df[GJ_TONNE] * df[TONNE_M3]
        df[GJ_LITRE] = df[GJ_M3] / 1e3

        df[CONVERSIONS_TO_PHYSICAL_UNITS] = 1.0
        tco2_liter_fuels = [c.GSO, c.DIE, c.LPG, c.KER, c.JFU]
        idx = pd.IndexSlice
        df.loc[
            idx[:, tco2_liter_fuels], CONVERSIONS_TO_PHYSICAL_UNITS
        ] = df.loc[idx[:, tco2_liter_fuels], GJ_LITRE]
        df.loc[idx[:, c.OOP], CONVERSIONS_TO_PHYSICAL_UNITS] = CONV_BARREL_TO_GJ

        return df


    def __get_conv_factor_volume_unit_to_gj(
            self,
            ) -> pd.DataFrame:
        """
        'Conversion factor (volume unit to GJ)' (v361)
        CPAT Excel: I960:I992

        returns dims (c, s, f), data column
        """
        all_sectors = [c.POW, c.ROD, c.RES, c.IND]
        idx_expand_mapping = {
            c.COA: all_sectors,
            c.NGA: all_sectors,
            c.GSO: all_sectors,
            c.DIE: all_sectors,
            c.KER: all_sectors,
            c.LPG: all_sectors,
            c.OOP: all_sectors + [c.ALL],
            c.BIO: [c.ALL],
            c.JFU: [c.AVI, c.ALL],
        }
        df = (
            self.fuel_calorific_values_and_densities[[CONVERSIONS_TO_PHYSICAL_UNITS]]
            .rename(
                columns={CONVERSIONS_TO_PHYSICAL_UNITS: 'Conversion factor (volume unit to GJ)'}
            )
            .reset_index()
            .assign(**{c.SECTOR_CODE:lambda d: d[c.FUEL_CODE].map(idx_expand_mapping)})
            .explode(c.SECTOR_CODE)
            .set_index(c.ID_COL_NAMES)
        )
        return df


    def __get_ef_iiasa_per_gj(
            self,
            ) -> pd.DataFrame:
        """
        'EF, IIASA, per GJ' (v361)
        CPAT Excel: L960:L992

        returns dims (c, s, f), data column
        """
        if set(self.iiasa_ef_tco2.index) != set(self.conv_factor_volume_unit_to_gj.index):
            raise ValueError(
                "iiasa_ef_tco2 and conv_factor_volume_unit_to_gj do not have the same index values."
            )
        # different column names, same indexes
        return (
            self.iiasa_ef_tco2.iloc[:, 0] * self.conv_factor_volume_unit_to_gj.iloc[:, 0]
        ).to_frame("EF, IIASA, per GJ")


    def __get_ef_tco2_per_volume_unit(
            self,
            selected_countries: list[str],
            efs_selected: Literal['IIASA', 'IEA']
            ) -> pd.DataFrame:
        """
        'EF - tCO2 per volume unit' (v361)
        CPAT Excel: O960:O992

        returns dims (c, s, f), ['EF - tCO2 per volume unit']
        """
        col_name = 'EF - tCO2 per volume unit'
        if efs_selected == 'IIASA':
            return (
                self.ef_iiasa_per_gj
                .rename(columns={'EF, IIASA, per GJ': col_name})
            )
        elif efs_selected == 'IEA':
            return pd.concat(
                {country: self.ef_global_iea for country in selected_countries},
                names=[c.COUNTRY_CODE]
            ).rename(columns={'EF-GlobalIEA': col_name})
        else:
            raise ValueError(f"Unknown EF source selected: {efs_selected}")


def load_cal_val_and_air_data(
        selected_countries: list[str]
        ) -> dict[CalValAirDataKeys, pd.DataFrame]:
    """
    [Data Loading]
    Table inputs: 'CalVal' and 'Air_Data' (v407).
    
    Used later in get_fuel_calorific_values_and_densities function,
    for
    'Fuels calorific values and densities - {Country}' table v361 Mitigation!B926:L936

    Manual changes to 'CalVal':
    - 'Code', 'Series' and 'Unit' columns removed
    - columns renamed: 'Country code'->'CountryCode', 'Subsector'->'SectorCode', 'Fuel'->'FuelCode'

    Manual changes to 'Air_Data':
    - only rows 3004:3154 - NetCV and density
    - 'Code' and 'Atrib 1' columns removed
    - columns renamed: 'Country'->'CountryCode', 'Atrib 2'->'FuelCode'
    - duplicated den.lpg.all removed

    Important: CountryCode values are iso codes in lower case!!!
    returns {
        'CalVal Countries': dims (c, s, f), 'Value' column
        'CalVal Regions': (c, s, f), 'Value' column
        'AirData density': dims (c, s, f), 'Value' column
        'AirData NetCV': dims (c, s, f), 'Value' column
    }
    dims (c, f), data columns
    """
    regions = [r.lower() for c, r in COUNTRY_TO_REGION.items() if c in selected_countries]
    selected_countries = [s.lower() for s in selected_countries]

    # load CalVal
    cal_val: pd.DataFrame = (
        # TODO: pkl
        # pd.read_pickle(f'{DATA_PATH}/{CAL_VAL_FILE_NAME}.pkl.bz2', compression="bz2")
        pd.read_csv(f'{DATA_PATH}/{CAL_VAL_FILE_NAME}.csv')
    )
    cal_val_countries = (
        cal_val[cal_val[c.COUNTRY_CODE].isin(selected_countries)]
    )
    cal_val_regions = (
        cal_val[cal_val[c.COUNTRY_CODE].isin(regions)]
    )
    # rename 'oil' to 'oop'
    cal_val_countries.loc[cal_val[c.FUEL_CODE] == c.OIL, c.FUEL_CODE] = c.OOP
    cal_val_regions.loc[cal_val[c.FUEL_CODE] == c.OIL, c.FUEL_CODE] = c.OOP

    cal_val_countries.set_index(c.ID_COL_NAMES, inplace=True)
    cal_val_regions.set_index(c.ID_COL_NAMES, inplace=True)

    # load Air_2
    air_2: pd.DataFrame = (
        # TODO: pkl
        # pd.read_pickle(f'{DATA_PATH}/{AIR_2_FILE_NAME}.pkl.bz2', compression="bz2")
        pd.read_csv(f'{DATA_PATH}/{AIR_DATA_FILE_NAME}.csv')
    )
    air_2 = air_2[air_2[c.COUNTRY_CODE].isin(selected_countries + [c.ALL])]

    air_2_den = (
        air_2[air_2['Series Name'] == 'density']
        .drop(columns=['Series Name'])
    )
    air_2_den.loc[air_2_den[c.FUEL_CODE] == c.OIL, c.FUEL_CODE] = c.OOP
    air_2_den.set_index([c.COUNTRY_CODE, c.FUEL_CODE], inplace=True)
    air_2_ncv = (
        air_2[air_2['Series Name'] == 'NetCV']
        .drop(columns=['Series Name'])
        # no need to rename oil->oop as we use oop nvc from CalVal
        .set_index([c.COUNTRY_CODE, c.FUEL_CODE])
    )

    return {
        'CalVal Countries': cal_val_countries,
        'CalVal Regions': cal_val_regions,
        'AirData density': air_2_den,
        'AirData NetCV': air_2_ncv
    }


def load_ef_ghg(
        selected_countries: list[str]
        ) -> dict[EfGhgKeys, pd.DataFrame]:
    """
    [Data Loading]
    'EF_GHG' tab data (v407)

    Used later in get_iiasa_ef_tco2 function
    to obtain 'IIASA EF tCO2'.

    Manual changes to EF_GHG -> ef_ghg.csv:
    - columns removed: 'Code', 'Series name'
    - columns renamed: 'Region/country' -> 'CountryCode','Atribute 1' -> 'SectorCode',
                        'Atribute 2' -> 'FuelCode', 'FACTOR' -> 'Value'

    Important: CountryCode values are iso codes in lower case!!!
    returns {
        'Countries': dims (c, s, f), 'Value' column
        'Regions': dims (c, s, f), 'Value' column
        'World': dims (c, s, f), 'Value' column
    }
    """
    ef_ghg = {}
    world_code = 'world'
    regions = [r.lower() for c, r in COUNTRY_TO_REGION.items() if c in selected_countries]
    selected_countries_lower = [s.lower() for s in selected_countries]

    ef_ghg_raw: pd.DataFrame = (
        # TODO: pkl
        # pd.read_pickle(f'{DATA_PATH}/{EF_GHG_FILE_NAME}.pkl.bz2', compression="bz2")
        pd.read_csv(f'{DATA_PATH}/{EF_GHG_FILE_NAME}.csv')
        .set_index(c.ID_COL_NAMES)
    )

    ef_ghg['Countries'] = (
        ef_ghg_raw[
            ef_ghg_raw.index.get_level_values(c.COUNTRY_CODE).isin(selected_countries_lower)
            ]
    )
    ef_ghg['Regions'] = (
        ef_ghg_raw[
            ef_ghg_raw.index.get_level_values(c.COUNTRY_CODE).isin(regions)
            ]
    )
    ef_ghg['World'] = (
        ef_ghg_raw[
            ef_ghg_raw.index.get_level_values(c.COUNTRY_CODE).isin([world_code])
            ]
    )

    return ef_ghg


def get_co2_efs_global() -> pd.DataFrame:
    """
    [Data Loading]
    'Fuels calorific values and densities - Global' table from v407 Mitigation!B922:L931
    and 'CO2 EFs - Global' table from v407 Mitigation!B951:L960.

    Manual changes to 'Fuels calorific values and densities - Global':
    - 'ren' row removed (in Excel version it is a blank placeholder)
    - both 'Sector' column removed, 'Fuel Description' column removed
    - columns renamed: 'Fuel Code' -> 'FuelCode'
    - empty column removed

    Columns used from 'CO2 EFs - Global' are:
    - EF- tC/GJ
    - EF-tCO2/GJ
    - EF-physical units
    'ren' row removed (in Excel version it is a blank placeholder).
    'CO2 EFs - Global' table merged into
        'Fuels calorific values and densities - Global'. 


    returns dims (f), data coulms
    """
    df: pd.DataFrame = (
        # TODO: pkl
        # pd.read_pickle(f'{DATA_PATH}/{GLOBAL_VALUES_DATA_FILE_NAME}.pkl.bz2', compression="bz2")
        pd.read_csv(f'{DATA_PATH}/{GLOBAL_VALUES_DATA_FILE_NAME}.csv')
    )

    # rename 'oil' to 'oop'
    df.loc[df[c.FUEL_CODE] == c.OIL, c.FUEL_CODE] = c.OOP

    df = (
        df.set_index([c.FUEL_CODE])
        .fillna(0.0)
    )
    return df
