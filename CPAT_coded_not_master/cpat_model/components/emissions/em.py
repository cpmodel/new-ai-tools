import pandas as pd

from cpat_model.components.efs.co2 import EFsCO2

import cpat_model.constants as c

class CO2Emissions:
    """
    Logic from 'Results - Energy-related CO2 emissions - {scenario}' section.

    'Adjustment factor' not implemented.
    """
    em: pd.DataFrame
    ef: pd.DataFrame
    em_by_fuel: pd.DataFrame
    total_em: pd.DataFrame

    def __init__(
            self,
            ec_fossil_fuels_by_sector: pd.DataFrame,
            efs_co2: EFsCO2
            ) -> None:
        self.__init_em(ec_fossil_fuels_by_sector)
        self.__init_em_by_fuel()
        self.__init_total_em()
        self.__set_ef(efs_co2)


    def __init_em(
            self,
            ec_fossil_fuels_by_sector: pd.DataFrame
            ) -> None:
        """
        'CO2 emissions by fuel and sector'

        CPAT Excel: scenario rows 4897:4939, baseline: rows 9692:9734 (v416)

        Init with all 0.0s.
        
        em dims (c, s, f), t in <simulation_years[0], simulation_years[1]>
        """
        self.em = ec_fossil_fuels_by_sector.copy()

        # runtime guard:
        if not (self.em.to_numpy() == 0.0).all():
            raise ValueError("CO2Emissions must contain only 0.0 at init")


    def __init_em_by_fuel(self) -> None:
        """
        'CO2 emissions by fuel'
        from 'CO2 emissions by fuel and sector'

        CPAT Excel: scenario rows 4897:4939, baseline: rows 9692:9734 (v416)

        Init with all 0.0s.
        
        em_by_fuel dims (c, f), t in <simulation_years[0], simulation_years[1]>
        """
        self.em_by_fuel = (
            self.em
            .groupby(level=[c.COUNTRY_CODE, c.FUEL_CODE])
            .sum()
        )


    def __init_total_em(self) -> None:
        """
        'Total CO2 emissions from fossil fuels'

        CPAT Excel: scenario row 4940, baseline: row 9735 (v416)

        Init with all 0.0s.
        
        total_em dims (c), t in <simulation_years[0], simulation_years[1]>
        """
        self.total_em = (
            self.em_by_fuel
            .groupby(level=[c.COUNTRY_CODE])
            .sum()
        )


    def __set_ef(
            self,
            efs_co2: EFsCO2
            ) -> None:
        """
        'EF (ton CO2e/ktoe)'

        CPAT Excel: scenario B4897:B4939, baseline: B9692:B9734 (v416)

        Important: /= 1e6 performed here not in the 'CO2 emissions by fuel' calculations
        
        ef dims (c, s, f), data column
        """
        self.ef = efs_co2.ef_tco2_per_ktoe.loc[
            (
                efs_co2.ef_tco2_per_ktoe.index
                .get_level_values(c.SECTOR_CODE).isin([c.POW, c.RES, c.ROD, c.IND])
            )
            & efs_co2.ef_tco2_per_ktoe.index.get_level_values(c.FUEL_CODE).isin(c.FOSSIL_FUELS)
        ].iloc[:, -1].copy()

        # rename to ec sectors:
        self.ef = self.ef.rename(index={c.RES: c.BLD, c.ROD: c.TRA}, level=c.SECTOR_CODE)

        # in Excel we use ind for oen (apart from oop where we use pow)
        oen_ef = self.ef.loc[
            (
                (self.ef.index.get_level_values(c.SECTOR_CODE) == c.IND)
                & (self.ef.index.get_level_values(c.FUEL_CODE) != c.OOP)
            ) | (
                (self.ef.index.get_level_values(c.SECTOR_CODE) == c.POW)
                & (self.ef.index.get_level_values(c.FUEL_CODE) == c.OOP)
            )
        ].copy()
        oen_ef.index = pd.MultiIndex.from_arrays(
            [
                oen_ef.index.get_level_values(c.COUNTRY_CODE),
                [c.OEN] * len(oen_ef),
                oen_ef.index.get_level_values(c.FUEL_CODE),
            ],
            names=c.ID_COL_NAMES
        )

        self.ef = pd.concat([self.ef, oen_ef]).sort_index()
        # TODO: implement co2_adjustment_factor
        # self.em = self.em.mul(efs_co2.co2_adjustment_factor, level=c.COUNTRY_CODE)

        # /= 1e6 applied here, can be moved if 'EF (ton CO2e/ktoe)' is used somewhere else:
        self.ef /= 1e6


    def calculate_em_year(
            self,
            year: int,
            ec_fossil_fuels_by_sector: pd.DataFrame
            ) -> None:
        """
        'CO2 emissions by fuel and sector'
        and
        'CO2 emissions by fuel'
        and
        'Total CO2 emissions from fossil fuels'

        CPAT Excel: scenario rows 4897:4940, baseline: rows 9692:9735 (v416)

        Updates values for a given year.

        Important: /= 1e6 performed here not in the 'CO2 emissions by fuel' calculations
        
        em dims (c, s, f), t in <simulation_years[0], simulation_years[1]>
        em_by_fuel dims (c, f), t in <simulation_years[0], simulation_years[1]>
        total_em dims (c), t in <simulation_years[0], simulation_years[1]>
        """
        y = str(year)

        self.em[y] = (
            ec_fossil_fuels_by_sector[y]
            * self.ef
        )

        self.em_by_fuel[y] = (
            self.em[y]
            .groupby(level=[c.COUNTRY_CODE, c.FUEL_CODE])
            .sum()
        )

        self.total_em[y] = (
            self.em_by_fuel[y]
            .groupby(level=[c.COUNTRY_CODE])
            .sum()
        )
