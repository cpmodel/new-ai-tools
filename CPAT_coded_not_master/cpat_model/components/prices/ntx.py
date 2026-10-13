from typing import Literal, TYPE_CHECKING

import pandas as pd
import numpy as np

from cpat_model.inputs.dashboard_inputs import DashboardInputsDict, PigouvianTaxType
import cpat_model.constants as c
from config import DATA_PATH

if TYPE_CHECKING:
    from cpat_model.inputs.input_data import InputData

CO_BENEFITS_FFS_FILE_NAME = 'co_benefits_ffs'

EXCISE_REFORM_IDX_PAIRS = [
    (c.POW, c.COA), (c.RES, c.COA), (c.IND, c.COA),
    (c.POW, c.NGA), (c.RES, c.NGA), (c.IND, c.NGA),
    (c.RES, c.ELE), (c.IND, c.ELE),
    (c.ALL, c.GSO), (c.ALL, c.DIE), (c.ALL, c.LPG),
    (c.ALL, c.KER), (c.ALL, c.OOP), (c.ALL, c.BIO)
]
EXCISE_REFORM = pd.DataFrame(
    [[0.0] * 31 for _ in range(14)], # 31 is the amount of cols
    index=pd.MultiIndex.from_tuples(
        EXCISE_REFORM_IDX_PAIRS,
        names=[c.SECTOR_CODE, c.FUEL_CODE]
        ),
    columns=np.arange(2020, 2050 + 1).astype(str)
) # 'Manual inputs'!J98:AM111 (v361) TODO: input table/input csv in the UI, double check years!


class NTX:
    """
    Static class that calcualtes 'new excise tax' for years in
    <simulation_years[0] - 1, simulation_years[1]>
    
    It can be changed so it holds all the mid step variables as attributes.
    """

    @staticmethod
    def calculate_ntx(
            scenario_type: str,
            selected_countries: list[str],
            simulation_years: tuple[int, int],
            deflator: pd.DataFrame,
            # Preloaded input data:
            input_data: 'InputData',
            # Dashboard inputs:
            d: DashboardInputsDict
            ) -> pd.DataFrame:
        """
        Returns ntx for <simulation_years[0] - 1, simulation_years[1]>
        """
        externalities_ffs = NTX._get_externalities_ffs(
            selected_countries, simulation_years, input_data.co_benefits_ffs
        )
        pigouvian_phase_in_coeff = NTX._get_pigouvian_phase_in_coeff(
            d['pigouvian_phase_in_t'], d['start_cp_year'], simulation_years
        )
        is_additional_ntx = NTX._get_is_additional_ntx(
            d['add_efficient_pigouvian_tax'], d['add_additional_excise_tax']
        )
        return NTX._get_ntx(
            scenario_type, selected_countries, simulation_years, is_additional_ntx,
            d['add_efficient_pigouvian_tax'], externalities_ffs['ap_externalities'],
            externalities_ffs['tra_externalities'], deflator, pigouvian_phase_in_coeff
        )

    @staticmethod
    def _get_ntx(
            scenario_type: str,
            selected_countries: list[str],
            simulation_years: tuple[int, int],
            is_additional_ntx: bool,
            add_efficient_pigouvian_tax: PigouvianTaxType,
            ap_externalities_ffs: pd.DataFrame,
            tra_externalities_ffs: pd.DataFrame,
            deflator: pd.DataFrame,
            pigouvian_phase_in_coeff: pd.DataFrame,
            excise_reform: pd.DataFrame = EXCISE_REFORM, # TODO input
            ap_externalities_ap: pd.DataFrame = pd.DataFrame(), # TODO: implement
            tra_externalities_tra: pd.DataFrame = pd.DataFrame(), # TODO: implement
            ) -> pd.DataFrame:
        """
        'new excise tax' (v361)
        CPAT Excel: 170:184

        returns dims (c, s, f), t in <simulation_years[0] - 1 , simulation_years[1]>
        """
        # drop simulation_years[0] - 1 timestep from defletor
        deflator_no_first_year = deflator.loc[:, deflator.columns[1:]]
        # drop columns that are above simulation years limit and extend idx by countries
        excise_reform = (
            excise_reform.loc[
                :,
                [str(y) for y in range(simulation_years[0], simulation_years[1] + 1)]
            ]
        )
        excise_reform = excise_reform.reset_index()
        excise_reform_per_country = []
        for country_code in selected_countries:
            excise_reform_per_country.append(
                excise_reform.assign(
                    **{c.COUNTRY_CODE: country_code}
                ).set_index(c.ID_COL_NAMES)
            )
        excise_reform = pd.concat(excise_reform_per_country).sort_index()

        if not all(
            set(excise_reform.columns) == set(df.columns)
            for df in [
                pigouvian_phase_in_coeff, deflator_no_first_year,
                tra_externalities_ffs, ap_externalities_ffs
            ]
        ):
            raise ValueError("Columns of ntx inputs are not the same!")

        if not is_additional_ntx or scenario_type == c.BASELINE:
            ntx = excise_reform.loc[:, :].copy()
            ntx.loc[:, :] = 0.0
            # add simulation_years[0] - 1 columns with 0s
            ntx.insert(0, str(simulation_years[0] - 1), 0.0)
            return ntx

        # not baseline and either additional excise tax or pigouvian tax is selected:
        externalities = pd.DataFrame()
        oop_externalities_all = []
        bio_externalities_all = []
        gso_die_externalities_all = []
        elec_externalities_all = []

        # TODO: implement after connecting with AP module and getting externalities from Transport
        # should be very similar to == 'FFS' case
        # TODO: and TEST!
        # Air pollution externalities from Air Pollution and Transport module
        if add_efficient_pigouvian_tax == 'pigouvian':
            raise NotImplementedError("Pigouvian tax not yet supported.")
            # pylint: disable=unreachable
            externalities = ap_externalities_ap[
                ap_externalities_ap.index.get_level_values(c.FUEL_CODE).isin(
                    [c.COA, c.NGA, c.ELE, c.LPG, c.KER]
                )
            ]

            # for gso and die we sum all 3 tra externalities
            f_tra_externalities = tra_externalities_tra[
                tra_externalities_tra.index.get_level_values(c.FUEL_CODE).isin([c.GSO, c.DIE])
            ].groupby([c.COUNTRY_CODE, c.FUEL_CODE]).sum()
            for country_code in selected_countries:
                # (oop, all) from (oop, pow)
                oop_externalities = ap_externalities_ap.loc[[(country_code, c.POW, c.OOP)]].copy()
                oop_externalities.index = pd.MultiIndex.from_tuples(
                    [(country_code, c.ALL, c.OOP)], names=c.ID_COL_NAMES
                )
                oop_externalities_all.append(oop_externalities.copy())

                # bio from tra_externalities_tra acc, gso
                bio_externalities = tra_externalities_tra.loc[[(country_code, c.ACC, c.GSO)]].copy()
                bio_externalities.index = pd.MultiIndex.from_tuples(
                    [(country_code, c.ALL, c.GSO)], names=c.ID_COL_NAMES
                )
                bio_externalities_all.append(bio_externalities.copy())

                # gso, die
                for fuel in [c.GSO, c.DIE]:
                    f_externalities = ap_externalities_ap.loc[[(country_code, c.ALL, fuel)]].copy()
                    f_tra_externalities_country = f_tra_externalities[
                        (
                            f_tra_externalities.index.get_level_values(c.COUNTRY_CODE)
                            .isin([country_code])
                        )
                        & f_tra_externalities.index.get_level_values(c.FUEL_CODE).isin([fuel])
                    ]
                    f_tra_externalities_country.index = pd.MultiIndex.from_tuples(
                        [(country_code, c.ALL, fuel)], names=c.ID_COL_NAMES
                    )
                    f_externalities += f_tra_externalities_country
                    gso_die_externalities_all.append(f_externalities.copy())
            # pylint: enable=unreachable


        # Air pollution and transportation externalities from FFS
        elif add_efficient_pigouvian_tax == 'FFS':
            externalities = ap_externalities_ffs[
                ap_externalities_ffs.index.get_level_values(c.FUEL_CODE).isin(
                    [c.COA, c.NGA, c.ELE, c.LPG, c.KER]
                )
            ]

            # for gso and die we sum all 3 tra externalities
            f_tra_externalities = tra_externalities_ffs[
                tra_externalities_ffs.index.get_level_values(c.FUEL_CODE).isin([c.GSO, c.DIE])
            ].groupby([c.COUNTRY_CODE, c.FUEL_CODE]).sum()
            for country_code in selected_countries:
                # (oop, all) from (oop, pow)
                oop_externalities = ap_externalities_ffs.loc[[(country_code, c.POW, c.OOP)]].copy()
                oop_externalities.index = pd.MultiIndex.from_tuples(
                    [(country_code, c.ALL, c.OOP)], names=c.ID_COL_NAMES
                )
                oop_externalities_all.append(oop_externalities.copy())

                # no bio when using FFS, adding rows with 0.0s
                bio_externalities = pd.DataFrame(
                    [[0.0] * tra_externalities_ffs.shape[1]],
                    index=pd.MultiIndex.from_tuples(
                        [(country_code, c.ALL, c.BIO)], names=c.ID_COL_NAMES
                    ),
                    columns=tra_externalities_ffs.columns
                )
                bio_externalities_all.append(bio_externalities.copy())

                # gso, die
                for fuel in [c.GSO, c.DIE]:
                    f_externalities = ap_externalities_ffs.loc[[(country_code, c.ALL, fuel)]].copy()
                    f_tra_externalities_country = f_tra_externalities[
                        (
                            f_tra_externalities.index.get_level_values(c.COUNTRY_CODE)
                            .isin([country_code])
                        )
                        & f_tra_externalities.index.get_level_values(c.FUEL_CODE).isin([fuel])
                    ]
                    f_tra_externalities_country.index = pd.MultiIndex.from_tuples(
                        [(country_code, c.ALL, fuel)], names=c.ID_COL_NAMES
                    )
                    f_externalities += f_tra_externalities_country
                    gso_die_externalities_all.append(f_externalities.copy())


        # connect
        externalities = pd.concat([
            externalities, *oop_externalities_all, *bio_externalities_all,
            *elec_externalities_all, *gso_die_externalities_all
        ])
        # pigouvian_phase_in_coeff not country specific:
        ntx = excise_reform + externalities.mul(pigouvian_phase_in_coeff.iloc[0], axis=1)
        # deflator is country specific so we need to reindex:
        ntx = ntx.mul(
            deflator_no_first_year.reindex(ntx.index.get_level_values(c.COUNTRY_CODE)).values,
            axis=1
        )

        # add simulation_years[0] - 1 columns with 0s
        ntx.insert(0, str(simulation_years[0] - 1), 0.0)
        return ntx


    @staticmethod
    def _get_is_additional_ntx(
            add_efficient_pigouvian_tax: PigouvianTaxType,
            add_additional_excise_tax: bool
            ) -> bool:
        """
        'Additional excise tax' (v361)
        CPAT Excel: D163 

        returns bool
        """
        if add_efficient_pigouvian_tax == 'pigouvian' or add_additional_excise_tax:
            return True
        return False


    @staticmethod
    def _get_pigouvian_phase_in_coeff(
            pigouvian_phase_in_t: int,
            start_cp_year: int,
            simulation_years: tuple[int, int]
            ) -> pd.DataFrame:
        """
        'phase-in coefficient -->' (v361)
        (for Pigouvian tax)
        CPAT Excel: 169

        returns dims all t
        """
        years = np.arange(simulation_years[0], simulation_years[1] + 1)
        columns = years.astype(str)
        pigouvian_phase_in_coeff = pd.DataFrame(
            [np.zeros(len(columns), dtype=float)], columns=columns
        )

        pigouvian_phase_in_coeff.loc[0] = np.where(
            years >= start_cp_year,
            np.minimum((years - start_cp_year + 1) / pigouvian_phase_in_t, 1),
            0
        )

        return pigouvian_phase_in_coeff


    @staticmethod
    def _get_externalities_ffs(
            selected_countries: list[str],
            simulation_years: tuple[int, int],
            co_benefits_ffs: pd.DataFrame
            ) -> dict[Literal['ap_externalities', 'tra_externalities'], pd.DataFrame]:
        """
        'Externalities from Fossil Fuel Subsidies' (v361)
        'Air pollution externalities' CPAT Excel: rows 130:138
        'Transportation externalities (lagged by 1 year)' CPAT Excel: rows 133:138

        co_benefits_ffs preloaded using `load_co_benefits_ffs` function.

        returns dict[..., pd.DataFrame[dims (c, s, f), all t]]
        """
        ap_columns = (
            ['year'] + [col for col in co_benefits_ffs.columns if col.startswith('cor_lap_')]
        )
        ap_externalities = co_benefits_ffs[ap_columns].copy().reset_index()
        ap_externalities = ap_externalities.melt(
            id_vars=[c.COUNTRY_CODE, 'year'],
            var_name='Code',
            value_name='Value'
        )
        ap_externalities = ap_externalities.pivot_table(
            index=[c.COUNTRY_CODE, 'Code'],
            columns='year',
            values='Value'
        ).reset_index()
        ap_externalities[c.FUEL_CODE] = ap_externalities['Code'].str[8:11]
        ap_externalities[c.SECTOR_CODE] = ap_externalities['Code'].str[-3:]
        ap_externalities = (
            ap_externalities
            .drop(columns=['Code'])
            .set_index(c.ID_COL_NAMES)
        )
        ap_externalities.columns = ap_externalities.columns.astype(str)
        # adding data after max t in db
        ap_last_year = max(int(col) for col in ap_externalities.columns)
        for y in range(ap_last_year + 1, simulation_years[1] + 1):
            ap_externalities[str(y)] = (
                (ap_externalities[str(y-1)] ** 2)
                / ap_externalities[str(y-2)]
            )
        # adding pow, oop row with proxy values from all, die
        oop_rows = ap_externalities.xs((c.ALL, c.DIE), level=[c.SECTOR_CODE, c.FUEL_CODE])
        oop_rows_index = pd.MultiIndex.from_tuples(
            [(country, c.POW, c.OOP) for country in oop_rows.index],
            names=ap_externalities.index.names
        )
        oop_rows = pd.DataFrame(
            oop_rows.values, index=oop_rows_index, columns=ap_externalities.columns
        )
        # adding missing res and ind for ele, filled with 0s
        ele_rows = pd.DataFrame(
            0.0,
            index=pd.MultiIndex.from_product(
                [selected_countries, [c.RES, c.IND], [c.ELE]],
                names=c.ID_COL_NAMES
            ),
            columns=ap_externalities.columns
        )
        ap_externalities = pd.concat([ap_externalities, oop_rows, ele_rows]).sort_index()
        ap_externalities.columns.name = None

        tra_columns = (
            ['year']
            + [col for col in co_benefits_ffs.columns
            if col.startswith('cor_acc') or col.startswith('cor_con') or col.startswith('cor_rdm')
            ]
        )
        tra_externalities = co_benefits_ffs[tra_columns].copy().reset_index()
        tra_externalities = tra_externalities.melt(
            id_vars=[c.COUNTRY_CODE, 'year'],
            var_name='Code',
            value_name='Value'
        )
        tra_externalities = tra_externalities.pivot_table(
            index=[c.COUNTRY_CODE, 'Code'],
            columns='year',
            values='Value'
        ).reset_index()
        tra_externalities[c.FUEL_CODE] = tra_externalities['Code'].str[8:11]
        # instead of sector w select externality type {acc, rdm, con}
        tra_externalities[c.SECTOR_CODE] = tra_externalities['Code'].str[4:7]
        tra_externalities = (
            tra_externalities
            .drop(columns=['Code'])
            .set_index(c.ID_COL_NAMES)
        )
        tra_externalities.columns = tra_externalities.columns.astype(str)
        # adding data after max t in db
        tra_last_year = max(int(col) for col in tra_externalities.columns)
        for y in range(tra_last_year + 1, simulation_years[1] + 1):
            tra_externalities[str(y)] = (
                (tra_externalities[str(y-1)] ** 2)
                / tra_externalities[str(y-2)]
            )
        # adding missing rdm, gso, filled with 0s
        rdm_gso = pd.DataFrame(
            0.0,
            index=pd.MultiIndex.from_product(
                [selected_countries, [c.RDM], [c.GSO]],
                names=c.ID_COL_NAMES
            ),
            columns=tra_externalities.columns
        )
        tra_externalities = pd.concat([tra_externalities, rdm_gso]).sort_index()
        tra_externalities.columns.name = None

        return {'ap_externalities': ap_externalities, 'tra_externalities': tra_externalities}


def load_co_benefits_ffs(
        selected_countries: list[str],
        simulation_years: tuple[int, int],
        ) -> pd.DataFrame:
    """
    [Data Loading]
    Loads data from 'CoBenefitsFFS' tab. (v407)

    Used later for 'Externalities from Fossil Fuel Subsidies' (v361). #TODO
    
    'CoBenefitsFFS' data preprocessed in `get_co_benefits_ffs_csv`
    from `cpat_processing/co_benefits_ffs_format.py`
    and saved to co_benefits_ffs.csv # TODO preprocess for better range years

    returns dims (c), data columns
    """
    co_benefits_ffs: pd.DataFrame = (
        # TODO: pkl
        # pd.read_pickle(f'{DATA_PATH}/{CO_BENEFITS_FFS_FILE_NAME}.pkl.bz2', compression="bz2")
        pd.read_csv(f'{DATA_PATH}/{CO_BENEFITS_FFS_FILE_NAME}.csv')
    )

    co_benefits_ffs = co_benefits_ffs[
        co_benefits_ffs[c.COUNTRY_CODE].isin(selected_countries)
        & (co_benefits_ffs['year'] >= simulation_years[0])
    ].set_index([c.COUNTRY_CODE]).drop(columns=['countryname'])

    assert (
        set(co_benefits_ffs.index.unique()) == set(selected_countries)
    ), f"Error: One of: {selected_countries} not in CoBenefitsFFS data."

    return co_benefits_ffs
