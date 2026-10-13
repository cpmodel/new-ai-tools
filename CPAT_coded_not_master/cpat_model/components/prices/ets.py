import pandas as pd

from cpat_model.components.carbon_pricing.existing_ets import ExistingETS

import cpat_model.constants as c

SECTOR_FUEL_INDEX_PAIRS = [
    *[(sector, fuel) for fuel in [c.COA, c.NGA] for sector in [c.POW, c.RES, c.IND]],
    (c.POW, c.OOP),
    *[(c.IND, fuel) for fuel in [c.GSO, c.DIE, c.LPG, c.KER]],
    *[(c.ALL, fuel) for fuel in [c.BIO, c.JFU]]
]


class ETS:
    """
    Static class that calcualtes 'Effective permit price' for years in
    <simulation_years[0] - 1, simulation_years[1]>
    """

    @staticmethod
    def get_ets(
            existing_ets: ExistingETS,
            ef_tco2_per_volume_unit: pd.DataFrame
            ) -> pd.DataFrame:
        """
        'Effective permit price' (v412)
        CPAT Excel: rows 1710:1723

        Changes:
        - apart from fuel in [coa, nga], all sectors changed to all
            for easier calculations later on

        returns dims (c, s, f), t in <simulation_years[0] - 1, simulation_years[1]>
        """
        coa_nga = [c.COA, c.NGA]
        transport_fuels = [c.GSO, c.DIE, c.KER, c.LPG, c.JFU, c.BIO]
        pow_res_ind = [c.POW, c.RES, c.IND]

        ets_coa_nga_dfs: list[pd.DataFrame] = []
        ets_other_fuels_dfs: list[pd.DataFrame] = []
        for s_data, f_data, permit_price in [
            (
                existing_ets.existing_ets_s_nat,
                existing_ets.existing_ets_f_nat,
                existing_ets.existing_ets_permit_price_nat
            ),
            (
                existing_ets.existing_ets_s_reg,
                existing_ets.existing_ets_f_reg,
                existing_ets.existing_ets_permit_price_reg
            )
        ]:
            ### coa, nga
            # existing_ets_s
            s_data_ind_pow_res = s_data[s_data.index.get_level_values(c.SECTOR_CODE) != 'trs']
            s_data_ind_pow_res = (
                s_data_ind_pow_res.reindex(s_data_ind_pow_res.index.repeat(len(coa_nga)))
            )
            s_data_ind_pow_res.index = pd.MultiIndex.from_tuples(
                [
                    (country, sector, fuel)
                    for (country, sector) in s_data_ind_pow_res.index[::len(coa_nga)]
                    for fuel in coa_nga
                ],
                names=c.ID_COL_NAMES
            )

            # existing_ets_f
            f_data_coa_nga = f_data[f_data.index.get_level_values(c.FUEL_CODE).isin(coa_nga)]
            f_data_coa_nga = f_data_coa_nga.reindex(f_data_coa_nga.index.repeat(len(pow_res_ind)))
            f_data_coa_nga.index = pd.MultiIndex.from_tuples(
                [
                    (country, sector, fuel)
                    for (country, fuel) in f_data_coa_nga.index[::len(pow_res_ind)]
                    for sector in pow_res_ind
                ],
                names=c.ID_COL_NAMES
            )
            ets_coa_nga_dfs.append(
                (s_data_ind_pow_res * f_data_coa_nga)
                .mul(permit_price, level=c.COUNTRY_CODE)
            )


            ### other fuels
            # existing_ets_s
            s_data_trs = s_data[s_data.index.get_level_values(c.SECTOR_CODE) == 'trs']
            s_data_pow = s_data[s_data.index.get_level_values(c.SECTOR_CODE) == c.POW]
            s_data_trs = s_data_trs.reindex(s_data_trs.index.repeat(len(transport_fuels)))
            s_data_trs.index = pd.MultiIndex.from_tuples(
                [
                    (country, sector, fuel)
                    for (country, sector) in s_data_trs.index[::len(transport_fuels)]
                    for fuel in transport_fuels
                ],
                names=c.ID_COL_NAMES
            )

            s_data_pow = (
                s_data_pow
                .assign(**{c.FUEL_CODE: c.OOP})
                .set_index(c.FUEL_CODE, append=True)
            )
            s_data_other_fuels = pd.concat([s_data_trs, s_data_pow])

            # existing_ets_f
            f_data_other_fuels = f_data[
                ~f_data.index.get_level_values(c.FUEL_CODE).isin(coa_nga)
            ]
            f_data_other_fuels.index = pd.MultiIndex.from_arrays(
                [
                    f_data_other_fuels.index.get_level_values(c.COUNTRY_CODE),
                    ['trs'] * len(f_data_other_fuels),
                    f_data_other_fuels.index.get_level_values(c.FUEL_CODE)
                ],
                names=c.ID_COL_NAMES
            )
            # rename trs to pow for oop
            idx_df = f_data_other_fuels.index.to_frame()
            mask = idx_df[c.FUEL_CODE] == c.OOP
            idx_df.loc[mask, c.SECTOR_CODE] = c.POW
            f_data_other_fuels.index = pd.MultiIndex.from_frame(idx_df)

            ets_other_fuels_dfs.append(
                f_data_other_fuels
                * s_data_other_fuels
                .mul(permit_price, level=c.COUNTRY_CODE)
            )

        ets_coa_nga: pd.DataFrame = sum(ets_coa_nga_dfs)
        ets_other_fuels: pd.DataFrame = sum(ets_other_fuels_dfs)
        ets = pd.concat([ets_coa_nga, ets_other_fuels]).sort_index()

        # # get only desired ef rows
        ef_filtered: pd.DataFrame = (
            ef_tco2_per_volume_unit[
                ef_tco2_per_volume_unit.index
                .droplevel(c.COUNTRY_CODE)
                .isin(SECTOR_FUEL_INDEX_PAIRS)
            ]
        )

        # # rename sectors to match ets sector values
        idx_df = ef_filtered.index.to_frame(index=False)
        mask = idx_df[c.FUEL_CODE].isin(transport_fuels)
        idx_df.loc[mask, c.SECTOR_CODE] = 'trs'
        ef_filtered.index = pd.MultiIndex.from_frame(idx_df)

        ets = ets.mul(ef_filtered.squeeze(), axis=0)

        # # if fuel is not in [coa, nga], rename sector to 'all'
        idx_df = ets.index.to_frame(index=False)
        mask = ~idx_df[c.FUEL_CODE].isin(coa_nga)
        idx_df.loc[mask, c.SECTOR_CODE] = c.ALL
        ets.index = pd.MultiIndex.from_frame(idx_df)

        return ets
