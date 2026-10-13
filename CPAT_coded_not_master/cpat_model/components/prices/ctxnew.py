import pandas as pd

import cpat_model.constants as c

from cpat_model.components.policies.policies import EXPLICIT_PRICING_VARIANT

# TODO: rename
SECTOR_FUEL_IDX_PAIRS = [
    (c.ALL, c.BIO), (c.ALL, c.OOP), (c.ROD, c.DIE), (c.ROD, c.GSO), (c.ROD, c.LPG), (c.IND, c.KER),
    (c.POW, c.COA), (c.RES, c.COA), (c.IND, c.COA), (c.POW, c.NGA), (c.RES, c.NGA), (c.IND, c.NGA)
]

PRICES_SECTOR_FUEL_IDX_PAIRS = [
    (c.ALL, c.BIO), (c.ALL, c.OOP), (c.ALL, c.DIE), (c.ALL, c.GSO), (c.ALL, c.LPG), (c.ALL, c.KER),
    (c.POW, c.COA), (c.RES, c.COA), (c.IND, c.COA), (c.POW, c.NGA), (c.RES, c.NGA), (c.IND, c.NGA)
]

class CTXNew:
    """
    Static class that calcualtes 'Implied carbon price by fuel'
    for one year.
    """

    @staticmethod
    def get_ctxnew(
            year: int,
            scenario_type: str,
            cp_trajectory: pd.DataFrame,
            p_based_policies_cov_s_f: pd.DataFrame,
            ef_tco2_per_volume_unit: pd.DataFrame
            ) -> pd.DataFrame:
        """
        'Implied carbon price by fuel' (v361)
        CPAT Excel: baseline 1912:1923, scenario 6902:6913

        Changes:
        - for all fuels apart from 'coa' and 'nga' sectors are changed to 'all'
            in order to make further transformations easier.

        returns dims (c, s, f), t
        """
        # change gso, ind to gso, all
        idx_df = p_based_policies_cov_s_f.index.to_frame()
        idx_df.loc[
            (idx_df[c.SECTOR_CODE] == c.IND)
            & (idx_df[c.FUEL_CODE] == c.OOP),
            c.SECTOR_CODE
        ] = c.ALL
        p_based_policies_cov_s_f.index = pd.MultiIndex.from_frame(idx_df)

        if (
            # no need to calc further as all factors are multiplied
            (cp_trajectory[str(year)] == 0.0).all()
            # same here, lines 1910 and 1911 are 0ed in Excel (v361)
            or not ((scenario_type == c.ETS) or (EXPLICIT_PRICING_VARIANT[scenario_type]))
        ):
            if scenario_type == c.POWER_FEEBATE:
                mask = (
                    ef_tco2_per_volume_unit.index
                    .droplevel(c.COUNTRY_CODE)
                    .isin([(c.POW, c.COA), (c.POW, c.NGA)])
                )
                ef_tco2_per_volume_unit = ef_tco2_per_volume_unit[mask]

                # rename column to year
                ef_tco2_per_volume_unit = (
                    ef_tco2_per_volume_unit.rename(columns={'EF - tCO2 per volume unit': str(year)})
                )
                return (
                    p_based_policies_cov_s_f
                    * ef_tco2_per_volume_unit
                ).fillna(0.0).mul(cp_trajectory.loc[:, [str(year)]], level=c.COUNTRY_CODE)
            ctxnew = p_based_policies_cov_s_f.copy().astype(float)
            ctxnew.loc[:, :] = 0.0
            return ctxnew

        # filter and rename index in ef_tco2_per_volume_unit
        # so it is the same as in p_based_policies_cov_s_f
        mask = ef_tco2_per_volume_unit.index.droplevel(c.COUNTRY_CODE).isin(SECTOR_FUEL_IDX_PAIRS)
        ef_tco2_per_volume_unit = ef_tco2_per_volume_unit[mask]

        idx_df = ef_tco2_per_volume_unit.index.to_frame()
        mask = ~idx_df[c.FUEL_CODE].isin([c.COA, c.NGA])
        idx_df.loc[mask, c.SECTOR_CODE] = c.ALL
        ef_tco2_per_volume_unit.index = pd.MultiIndex.from_frame(idx_df)

        assert (
            ef_tco2_per_volume_unit.index.sort_values()
            .equals(p_based_policies_cov_s_f.index.sort_values())
        ), "Indexes differ for ef_tco2_per_volume_unit and p_based_policies_cov_s_f."

        # rename column to year
        ef_tco2_per_volume_unit = (
            ef_tco2_per_volume_unit.rename(columns={'EF - tCO2 per volume unit': str(year)})
        )

        return (
            ef_tco2_per_volume_unit * p_based_policies_cov_s_f
        ).mul(cp_trajectory.loc[:, [str(year)]], level=c.COUNTRY_CODE)
