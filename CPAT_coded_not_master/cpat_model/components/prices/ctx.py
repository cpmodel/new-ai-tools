import pandas as pd

from cpat_model.components.carbon_pricing.existing_ct import ExistingCT

import cpat_model.constants as c

SECTOR_FUEL_INDEX_PAIRS = [
    *[(sector, fuel) for fuel in [c.COA, c.NGA] for sector in [c.POW, c.RES, c.IND]],
    (c.ALL, c.OOP),
    *[(c.IND, fuel) for fuel in [c.GSO, c.DIE, c.LPG, c.KER]],
    *[(c.ALL, fuel) for fuel in [c.BIO, c.JFU]]
]


class CTX:
    """
    Static class that calcualtes 'Effective carbon tax' for years in
    <simulation_years[0] - 1, simulation_years[1]>
    """

    @staticmethod
    def get_ctx(
            existing_ct: ExistingCT,
            ef_tco2_per_volume_unit: pd.DataFrame
            ) -> pd.DataFrame:
        """
        'Effective carbon tax' (v407)
        CPAT Excel: rows 1751:1764

        Changes:
        - apart from fuel in [coa, nga], all sectors changed to all
            for easier calculations later on

        returns dims (c, s, f), t in <simulation_years[0] - 1, simulation_years[1]>
        """
        # columns = [str(year) for year in range(simulation_years[0] - 1, simulation_years[1] + 1)]
        idx = pd.IndexSlice
        sp_countries = set(existing_ct.single_price)
        all_cointries = set(
            existing_ct.existing_ct_rate.index.get_level_values(c.COUNTRY_CODE).unique()
        )
        other_countries = list(all_cointries - sp_countries)
        sp_countries = list(sp_countries)

        coa_nga = [c.COA, c.NGA]
        transport_fuels = [c.GSO, c.DIE, c.KER, c.LPG, c.JFU, c.BIO]
        no_sp_fuels = [c.DIE, c.LPG, c.KER] # single price does not apply
        sp_fuels = [c.GSO, c.JFU, c.BIO, c.OOP] # single price applies
        pow_res_ind = [c.POW, c.RES, c.IND]

        # get rid off years 2015 - simulation_years[0] - 1 TODO: move up?
        existing_ct_s_filtered = existing_ct.existing_ct_s.copy()
        existing_ct_s_pow_res_ind = existing_ct_s_filtered[
            existing_ct_s_filtered.index.get_level_values(c.SECTOR_CODE) != 'trs'
        ]
        existing_ct_s_ind = existing_ct_s_pow_res_ind[
            existing_ct_s_pow_res_ind.index.get_level_values(c.SECTOR_CODE) == c.IND
        ]
        existing_ct_s_trs = existing_ct_s_filtered[
            existing_ct_s_filtered.index.get_level_values(c.SECTOR_CODE) == 'trs'
        ]

        existing_ct_f_filtered = existing_ct.existing_ct_f.copy()
        existing_ct_rate_filtered = existing_ct.existing_ct_rate.copy()
        existing_ct_f_rate_filtered = existing_ct.existing_ct_f_rate.copy()

        ### coa, nga
        # existing_ets_s
        existing_ct_s_pow_res_ind = (
            existing_ct_s_pow_res_ind
            .reindex(existing_ct_s_pow_res_ind.index.repeat(len(coa_nga)))
        )
        existing_ct_s_pow_res_ind.index = pd.MultiIndex.from_tuples(
            [
                (country, sector, fuel)
                for (country, sector) in existing_ct_s_pow_res_ind.index[::len(coa_nga)]
                for fuel in coa_nga
            ],
            names=c.ID_COL_NAMES
        )

        # existing_ets_f
        existing_ct_f_coa_nga = existing_ct_f_filtered[
            existing_ct_f_filtered.index.get_level_values(c.FUEL_CODE).isin(coa_nga)
        ]
        existing_ct_f_coa_nga = (
            existing_ct_f_coa_nga
            .reindex(existing_ct_f_coa_nga.index.repeat(len(pow_res_ind)))
        )
        existing_ct_f_coa_nga.index = pd.MultiIndex.from_tuples(
            [
                (country, sector, fuel)
                for (country, fuel) in existing_ct_f_coa_nga.index[::len(pow_res_ind)]
                for sector in pow_res_ind
            ],
            names=c.ID_COL_NAMES
        )

        # multiply
        ctx_coa_nga = existing_ct_s_pow_res_ind * existing_ct_f_coa_nga

        if existing_ct.single_price:
            # always single price for all fuels
            ctx_coa_nga.loc[idx[sp_countries, :, :], :] = (
                ctx_coa_nga.loc[idx[sp_countries, :, :], :]
                .mul(existing_ct_rate_filtered.loc[idx[sp_countries], :], level=c.COUNTRY_CODE)
            )
        if other_countries:
            for fuel in coa_nga:
                f_rate = (
                    existing_ct_f_rate_filtered.loc[idx[other_countries, [fuel]], :]
                    .droplevel(c.FUEL_CODE)
                )
                ctx_coa_nga.loc[idx[other_countries, :, [fuel]], :] = (
                    ctx_coa_nga.loc[idx[other_countries, :, [fuel]], :]
                    .mul(f_rate, level=c.COUNTRY_CODE)
                )

        ### other fuels
        # existing_ets_s
        existing_ct_s_trs = (
            existing_ct_s_trs
            .reindex(existing_ct_s_trs.index.repeat(len(transport_fuels)))
        )
        existing_ct_s_trs.index = pd.MultiIndex.from_tuples(
            [
                (country, sector, fuel)
                for (country, sector) in existing_ct_s_trs.index[::len(transport_fuels)]
                for fuel in transport_fuels
            ],
            names=c.ID_COL_NAMES
        )

        existing_ct_s_ind = (
            existing_ct_s_ind
            .assign(**{c.FUEL_CODE: c.OOP})
            .set_index(c.FUEL_CODE, append=True)
        )
        existing_ct_s_other_fuels = pd.concat([existing_ct_s_trs, existing_ct_s_ind])

        # existing_ets_f
        existing_ct_f_other_fuels = existing_ct_f_filtered[
            ~existing_ct_f_filtered.index.get_level_values(c.FUEL_CODE).isin(coa_nga)
        ]
        existing_ct_f_other_fuels.index = pd.MultiIndex.from_arrays(
            [
                existing_ct_f_other_fuels.index.get_level_values(c.COUNTRY_CODE),
                ['trs'] * len(existing_ct_f_other_fuels),
                existing_ct_f_other_fuels.index.get_level_values(c.FUEL_CODE)
            ],
            names=c.ID_COL_NAMES
        )

        # rename trs to ind for oop
        idx_df = existing_ct_f_other_fuels.index.to_frame()
        mask = idx_df[c.FUEL_CODE] == c.OOP
        idx_df.loc[mask, c.SECTOR_CODE] = c.IND
        existing_ct_f_other_fuels.index = pd.MultiIndex.from_frame(idx_df)

        ctx_other_fuels = existing_ct_f_other_fuels* existing_ct_s_other_fuels

        if existing_ct.single_price:
            # always single price for all f
            ctx_other_fuels.loc[idx[sp_countries, :, :], :] = (
                ctx_other_fuels.loc[idx[sp_countries, :, :], :]
                .mul(existing_ct_rate_filtered.loc[idx[sp_countries], :], level=c.COUNTRY_CODE)
            )
        if other_countries:
            for fuel in sp_fuels:
                f_rate = (
                    existing_ct_f_rate_filtered.loc[idx[other_countries, [fuel]], :]
                    .droplevel(c.FUEL_CODE)
                )
                ctx_other_fuels.loc[idx[other_countries, :, [fuel]], :] = (
                    ctx_other_fuels.loc[idx[other_countries, :, [fuel]], :]
                    .mul(f_rate, level=c.COUNTRY_CODE)
                )

            # always single price for these fuels
            ctx_other_fuels.loc[idx[other_countries, :, no_sp_fuels], :] = (
                ctx_other_fuels.loc[idx[other_countries, :, no_sp_fuels], :]
                .mul(existing_ct_rate_filtered.loc[idx[other_countries], :], level=c.COUNTRY_CODE)
            )

        ctx = pd.concat([ctx_coa_nga, ctx_other_fuels]).sort_index()

        # get only desired ef rows
        ef_tco2_per_volume_unit_filtered: pd.DataFrame = (
            ef_tco2_per_volume_unit[
                ef_tco2_per_volume_unit.index
                .droplevel(c.COUNTRY_CODE)
                .isin(SECTOR_FUEL_INDEX_PAIRS)
            ]
        )

        # rename sectors to match ctx sector values
        idx_df = ef_tco2_per_volume_unit_filtered.index.to_frame(index=False)
        idx_df[c.SECTOR_CODE] = idx_df.apply(
            lambda row: (
                'trs' if row[c.FUEL_CODE] in transport_fuels
                else c.IND if row[c.FUEL_CODE] == c.OOP
                else row[c.SECTOR_CODE]
            ),
            axis=1
        )
        ef_tco2_per_volume_unit_filtered.index = pd.MultiIndex.from_frame(idx_df)

        ctx = ctx.mul(ef_tco2_per_volume_unit_filtered.squeeze(), axis=0)

        # if fuel is not in [coa, nga], rename sector to 'all'
        idx_df = ctx.index.to_frame(index=False)
        mask = ~idx_df[c.FUEL_CODE].isin(coa_nga)
        idx_df.loc[mask, c.SECTOR_CODE] = c.ALL
        ctx.index = pd.MultiIndex.from_frame(idx_df)

        return ctx
