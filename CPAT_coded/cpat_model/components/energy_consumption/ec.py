from typing import TYPE_CHECKING, Literal

import pandas as pd

from cpat_model.inputs.dashboard_inputs import DashboardInputsDict
from cpat_model.components.prices.prices import EnergyPrices
from cpat_model.components.elasticities.elasticities import Elasticities, ElasticitiesSectorsType
from cpat_model.components.policies.shadow_prices import ShadowPrices

import cpat_model.constants as c
if TYPE_CHECKING:
    from cpat_model.inputs.input_data import InputData


# mapping made according to elasticities used not sectors:subsectors aggregation!:
EC_SECTORS = {
    c.TRA: c.SECTORS_AGG[c.TRA],
    c.RES: [c.RES],
    c.SRV: [c.SRV],
    c.IND: c.SECTORS_AGG[c.IND]+ c.SECTORS_AGG[c.OEN] + [c.FOO]
}
SUBSECTORS_TO_SECTORS = {
    subsector: sector
    for sector, subsectors in EC_SECTORS.items()
    for subsector in subsectors
}
EC_FUELS = c.FOSSIL_FUELS + [c.BIO, c.REN]

# for totals:
ALL_SUBSECTORS = list(SUBSECTORS_TO_SECTORS.keys())


EcSectorsType = Literal[
    'rod', 'ral', 'avi', 'nav',
    'res',
    'srv',
    'mch', 'irn', 'nfm', 'mac', 'cem', 'omn', 'cst', 'ftr', 'oen', 'foo'
]

class EC:
    """
    Energy Consumption with simplified: GDP growth and 'Residential' sector
    'Energy Use'
    and
    'After Tax Price (taking account of any exemptions)'
    from 'Transport sector', 'Buildings sector', 'Industrial sector'
    and 'Other - energy use' sections (v416)

    CPAT Excel: baseline 4149:4526, scenario 8816:9321

    TODO: Not all totals implemented yet ('Total Energy Use by Fuel Type',
    and all 'Total' and 'Total {sector}')


    General 'Energy Use' formula:
    ec(c, s, f, t) = (
        ec(c, s, f, t-1)
        * (1 + d_gdp_at_const_prices(c, t)) ^ el_inc(c, s_e, f_e)
        * (1 + exogenous_shock_on_ec(c, t))
        * (post_tax_p(c, s, f, t) / post_tax_p(c, s, f, t-1)) ^ el_dem(c, s_e, f_e)
        * (
            (
                (post_tax_p(c, s, f, t) + shadow_prices(c, s, f, t) * percent_of_shadow_p(s)) 
                / (post_tax_p(c, s, f, t-1) + shadow_prices(c, s, f, t-1) * percent_of_shadow_p(s)) 
            ) ^ el_eff(c, s_e, f_e)
            / (1 + eff_imp(c, s_e, f_e) + additional_eff_gains(s))
        ) ^ (1 + el_dem(c, s_e, f_e))

    Where:
    - c - Country, s- Sector, f - Fuel, t - year, s_e - elasticity sector,  f_e - elasticity fuel
    - ec - 'Energy Use'
    - d_gdp_at_const_prices - 'Change in country total real GDP at constant prices'
    - elasticities:
        - el_inc - 'Income elasticities - energy use'
        - el_dem - 'Own-price elasticities of demand - intensive margin
            (usage of energy-using capital)'
        - el_eff - 'Own-price elasticities of demand - efficiency and extenstive margin
            (fuel economy of energy-using capital and ownership)'
        - eff_imp - 'Autonomous efficiency improvments in energy-consuming capital
            (cars, buildings, factories etc)'
        - elasticity sectors (s_e) are mapped to ec sectors (f) in EC_SECTORS
        - apart from oil all of elasticity fuels (f_e) are mapped 1:1 with ec fuels (f).
            For el_inc, el_dem and el_eff oil mapps to lpg, ker, oop.
            For eff_imp oil mapps to gso, die, lpg, ker, oop.
    - exogenous_shock_on_ec - 'Exogenous shock on energy consumption'
    - shadow_prices - 'Shadow prices'
    - percent_of_shadow_p - '% of shadow price impacting efficiency by sector'
    - additional_eff_gains - 'Additional policy-induced annual efficiency gains'

    Exemptions to the formula:
    - in all sectors for 'Self-Generated Renewables' (ren) there is no
        '+ shadow_prices * percent_of_shadow_p' (t and t-1)
    - in 'Road' (rod) 'Biomass' (bio) is calculated as a sum of biofuels:
        - 'Bioethanol' (bgs) 
        - 'Biodiesel' (bdi)
        - 'Other biofuels' (obf)
        - all of the biofuels use elasticities for oil,
            bgs and obf use 'Gasoline' (gso) prices and bdi use 'Diesel' (die) prices
    - in 'Other energy use' (oen) there is no '+ additional_eff_gains'

    Additional calculations:
    - in 'Domestic aviation' (avi) we additionally calculate energy use for 'Jet Fuel' (jfu)
        and it should be included in totals in aggregations 
    

    TODOs:
    - rest of totals
    - d_post_pol_gdp 'Post-policy GDP growth' should replace d_gdp_at_const_prices
    - d_factor_elast 'Discount factor (proportion of estimated to average
        calibrated income elasticity)' should be included in the gdp_component
    - ec for 'Residential' (res) sector should be calculated differently
        (for fuels: nga, lpg, ker, bio)
    """
    ec: dict[EcSectorsType, pd.DataFrame]
    ec_rod_bio: pd.DataFrame # bgs, bdi, obf
    ec_avi_jfu: pd. DataFrame # (avi, jfu)

    ec_by_subsector: pd.DataFrame
    ec_fossil_fuels_by_sector: pd.DataFrame

    thermal_efficiency: pd.DataFrame

    # 'Post Tax Price (taking account of any exemptions)'
    post_tax_p: dict[EcSectorsType, pd.DataFrame]
    post_tax_p_rod_bio: pd.DataFrame

    # Helper for elasticities so we don't need to filter them at every time step
    __e: dict[
        Literal['el_dem', 'el_eff'],
        dict[
            ElasticitiesSectorsType,
            dict[
                Literal['oil', 'all'],
                pd.Series
            ]
        ]
    ]
    # Helper for shadow prices
    __shadow_prices_mul: dict[EcSectorsType, pd.DataFrame]
    # BUG in Excel (v416): in avi for coa, nga additional *exogenous_shock_component
    # BUG in Excel (v416): rod for obf and bdi instead of
    #   + 'Additional policy-induced annual efficiency gains' we use
    #   'Price-based policies coverage (% of the new price), by fuel and sector'
    #   for nga and coa correspondingly
    # BUG in Excel (v416): rod for obf el_dem instead of el_eff in the power

    def __init__(
            self,
            simulation_years: tuple[int, int],
            input_data: 'InputData',
            elasticities: Elasticities,
            shadow_prices: ShadowPrices
            ) -> None:
        # TODO: test all the methods!
        # for linter only:
        self.post_tax_p_rod_bio = pd.DataFrame()

        self.__init_year: int = simulation_years[0]
        self.__columns: list[str] = [
            str(i) for i in range(simulation_years[0], simulation_years[1] + 1)
        ]
        self.__init_ec(input_data)
        self.__countries = self.ec[c.ROD].index.get_level_values(c.COUNTRY_CODE).unique()
        self.__init_thermal_efficiency(input_data)

        self.__init_post_tax_p()

        self.__init_e(elasticities)
        self.__init_shadow_prices_mul(shadow_prices)

        self.__init_ec_by_subsector()
        self.__init_ec_fossil_fuels_by_sector()


    def __init_post_tax_p(self) -> None:
        """
        'Post Tax Price (taking account of any exemptions)'
        for all subsectors.
        Prices for biofuels stored separately in post_tax_p_rod_bio.

        CPAT Excel: sector sections in scenario rows 4021:4526, baseline: rows 8816:9321 (v416)

        post_tax_p[sector] dims (c, s, f), t in <simulation_years[0], simulation_years[1]>
        post_tax_p_rod_bio dims (c, s, f), t in <simulation_years[0], simulation_years[1]>
        """
        self.post_tax_p = {}

        for sector in SUBSECTORS_TO_SECTORS:
            index = pd.MultiIndex.from_product(
                [self.__countries, [sector], EC_FUELS],
                names=c.ID_COL_NAMES
            )
            self.post_tax_p[sector] = pd.DataFrame(
                0.0, index=index, columns=self.__columns
            ).sort_index()

        index_bio = pd.MultiIndex.from_product(
            [self.__countries, [c.ROD], c.BIO_COMPONENTS],
            names=c.ID_COL_NAMES
        )
        self.post_tax_p_rod_bio = pd.DataFrame(
            0.0, index=index_bio, columns=self.__columns
        ).sort_index()


    def __init_ec(
            self,
            input_data: 'InputData'
            ) -> None:
        """
        'Energy Use'
        for all subsectors.
        Energy use for biofuels and for ('Domestic aviation', 'Jet Fuel') stored separately
        in ec_rod_bio and ec_avi_jfu respectively.

        CPAT Excel: sector sections in scenario rows 4021:4526, baseline: rows 8816:9321 (v416)

        ec[sector] dims (c, s, f), t in <simulation_years[0], simulation_years[1]>
        ec_rod_bio dims (c, s, f), t in <simulation_years[0], simulation_years[1]>
        ec_avi_jfu dims (c, s, f), t in <simulation_years[0], simulation_years[1]>
        """
        self.ec = {}

        # biofuels
        ec_bio = input_data.ec_input[c.BIO_COMPONENTS].copy()
        ec_bio = ec_bio[ ec_bio.index.get_level_values(c.SECTOR_CODE) == c.ROD]
        # jfu
        ec_jfu = input_data.ec_input[c.JFU].copy()
        ec_jfu = ec_jfu[ec_jfu.index.get_level_values(c.SECTOR_CODE) == c.AVI]

        ec = input_data.ec_input[EC_FUELS].copy()

        for sector in SUBSECTORS_TO_SECTORS:
            ec_sector = ec[ec.index.get_level_values(c.SECTOR_CODE) == sector]
            self.ec[sector] = self.__get_ec_input_formatted(ec_sector)
            self.ec[sector][self.__columns[1:]] = 0.0

        self.ec_rod_bio = self.__get_ec_input_formatted(ec_bio)
        self.ec_avi_jfu = self.__get_ec_input_formatted(ec_jfu)
        self.ec_rod_bio[self.__columns[1:]] = 0.0
        self.ec_avi_jfu[self.__columns[1:]] = 0.0


    def __init_e(self, elasticities: Elasticities) -> None:
        """
        Hold values for el_dem and el_eff elasticities, grouped and filtered
        separately for oil and no oil (all).

        _e[elasticity][sector_agg][oil] dims (c), data series
        _e[elasticity][sector_agg][all] dims (c, f), data series
        """
        self.__e = {c.EL_DEM: {}, c.EL_EFF: {}}
        for sector_agg in EC_SECTORS:
            self.__e[c.EL_DEM][sector_agg] = {c.OIL: {}, c.ALL: {}}
            self.__e[c.EL_EFF][sector_agg] = {c.OIL: {}, c.ALL: {}}

            oil_el_dem_mask = (
                elasticities.e[sector_agg][c.EL_DEM].index
                .get_level_values(c.FUEL_CODE) == c.OIL
            )
            self.__e[c.EL_DEM][sector_agg][c.OIL] = (
                elasticities.e[sector_agg][c.EL_DEM].loc[oil_el_dem_mask]
                .squeeze(axis=1)
                .droplevel(c.FUEL_CODE)
            )
            self.__e[c.EL_DEM][sector_agg][c.ALL] = (
                elasticities.e[sector_agg][c.EL_DEM].loc[~oil_el_dem_mask]
                .squeeze(axis=1)
            )

            oil_el_eff_mask = (
                elasticities.e[sector_agg][c.EL_EFF].index
                .get_level_values(c.FUEL_CODE) == c.OIL
            )

            self.__e[c.EL_EFF][sector_agg][c.OIL] = (
                elasticities.e[sector_agg][c.EL_EFF].loc[oil_el_eff_mask]
                .squeeze(axis=1)
                .droplevel(c.FUEL_CODE)
            )
            self.__e[c.EL_EFF][sector_agg][c.ALL] = (
                elasticities.e[sector_agg][c.EL_EFF].loc[~oil_el_eff_mask]
                .squeeze(axis=1)
            )


    def __init_shadow_prices_mul(
            self,
            shadow_prices: ShadowPrices
            ) -> None:
        """
        Holds values for:
        shadow_prices(c, s, f, t) * percent_of_shadow_p(s)

        shadow_prices_mul[sector] dims (c, s, f), t in <simulation_years[0], simulation_years[1]>
        """
        self.__shadow_prices_mul = {}

        for sector in SUBSECTORS_TO_SECTORS:
            shadow_prices_mul = self.__calculate_shadow_prices_mul_for_sector(
                str(self.__init_year), sector, shadow_prices
            )
            self.__shadow_prices_mul[sector] = shadow_prices_mul.to_frame()
            self.__shadow_prices_mul[sector][self.__columns[1:]] = 0.0


    def __update_shadow_prices_mul(
            self,
            y: str,
            shadow_prices: ShadowPrices
            ) -> None:
        """
        Updates values for given year in:
        shadow_prices(c, s, f, t) * percent_of_shadow_p(s)

        shadow_prices_mul[sector] dims (c, s, f), t in <simulation_years[0], simulation_years[1]>
        """
        for sector in SUBSECTORS_TO_SECTORS:
            shadow_prices_mul = self.__calculate_shadow_prices_mul_for_sector(
                y, sector, shadow_prices
            )
            self.__shadow_prices_mul[sector][y] = shadow_prices_mul.copy()


    def __init_ec_by_subsector(self) -> None:
        """
        'Total Energy Use by Subsector'

        Inits with 0.0s

        ec_by_subsector[sector] dims (c, s), t in <simulation_years[0], simulation_years[1]>
        """
        index = pd.MultiIndex.from_product(
            [self.__countries, ALL_SUBSECTORS],
            names=[c.COUNTRY_CODE, c.SECTOR_CODE]
        )
        self.ec_by_subsector = pd.DataFrame(
            0.0, index=index, columns=self.__columns
        ).sort_index()


    def __init_ec_fossil_fuels_by_sector(self) -> None:
        """
        'Fossil fuel consumption by sector'
        from
        'Total energy consumption (TEC)'

        CPAT Excel: baseline rows 4817:4859, scenario rows 9612:9654 (v416)

        Inits with 0.0s

        ec_fossil_fuels_by_sector[sector] dims (c, s, f),
            t in <simulation_years[0], simulation_years[1]>
        """
        index = pd.MultiIndex.from_product(
            [self.__countries, c.SECTOR_GROUPS, c.FOSSIL_FUELS],
            names=c.ID_COL_NAMES
        )
        self.ec_fossil_fuels_by_sector = pd.DataFrame(
            0.0, index=index, columns=self.__columns
        ).sort_index()


    def __calculate_shadow_prices_mul_for_sector(
            self,
            y: str,
            sector: EcSectorsType,
            shadow_prices: ShadowPrices
            ) -> pd.Series:
        """
        Calculates values for given year in:
        shadow_prices(c, s, f, t) * percent_of_shadow_p(s)

        Used in init and update
        """
        idx = pd.IndexSlice
        # shadow prices sector for coa, nga at list index 0 (res or ind):
        if sector in [c.ROD, c.RAL, c.AVI, c.NAV, c.RES]:
            # rod, ral, avi, nav, res
            s_list = [c.RES, c.ALL]
        else:
            # sector in [foo, srv, mch, irn, nfm, mac, cem, omn, cst, ftr, oen]
            s_list = [c.IND, c.ALL]

        shadow_prices_mul: pd.Series = (
            shadow_prices.shadow_prices.loc[idx[:, s_list, :], y]
            * shadow_prices.percent_of_shadow_p.loc[sector]
        )
        shadow_prices_mul.index = pd.MultiIndex.from_arrays(
            [
                shadow_prices_mul.index.get_level_values(c.COUNTRY_CODE),
                [sector] * len(shadow_prices_mul),
                shadow_prices_mul.index.get_level_values(c.FUEL_CODE),
            ],
            names=c.ID_COL_NAMES
        )
        return shadow_prices_mul


    def __get_ec_input_formatted(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Helper function

        Formats ec_input into (c, s, f), simulation_year[0] table.
        """
        return (
            df
            .reset_index()
            .melt(
                id_vars=[c.COUNTRY_CODE, c.SECTOR_CODE],
                var_name=c.FUEL_CODE,
                value_name=str(self.__init_year)
            )
            .set_index(c.ID_COL_NAMES)
            .sort_index()
        )


    def __get_series_with_sector_code(self, sector: EcSectorsType, s: pd.Series) -> pd.Series:
        """
        Helper function

        Appends SectorCode index level with 1 index value to (c, f) index.
        """
        return (
            s
            .to_frame()
            .assign(**{c.SECTOR_CODE: sector})
            .set_index(c.SECTOR_CODE, append=True)
            .reorder_levels(c.ID_COL_NAMES)
            .iloc[:, 0]
        )


    def __init_thermal_efficiency(self, input_data: 'InputData') -> None:
        """
        'Thermal Efficiency'

        CPAT Excel: O592:O602 (v407)

        thermal_efficiency dims (c, f), 'Value' column
            f in [coa, nga, oop]
        """
        # filter:
        other_oop_fuels = [c.GSO, c.DIE, c.KER, c.LPG, c.JFU]
        thermal_efficiency = input_data.ec_input[[c.COA, c.NGA, c.OOP, *other_oop_fuels]].copy()
        thermal_efficiency = thermal_efficiency[
            thermal_efficiency.index.get_level_values(c.SECTOR_CODE).isin([c.ELEC, c.POW])
        ]

        # aggregate oop:
        thermal_efficiency[c.OOP] += thermal_efficiency[other_oop_fuels].sum(axis=1)
        thermal_efficiency = thermal_efficiency.drop(columns=other_oop_fuels)

        # elec / KTOE_TO_GWH / pow
        thermal_efficiency = thermal_efficiency.unstack(c.SECTOR_CODE)
        thermal_efficiency = (
            thermal_efficiency.xs(c.ELEC, axis=1, level=c.SECTOR_CODE)
            / thermal_efficiency.xs(c.POW, axis=1, level=c.SECTOR_CODE)
        )
        thermal_efficiency /= c.KTOE_TO_GWH

        # fill NaNs (div by 0) with default values:
        default_efficiencies = {c.COA: 0.35, c.NGA: 0.4, c.OOP: 0.35}
        self.thermal_efficiency = thermal_efficiency.fillna(value=default_efficiencies)
        # TODO: unstack to (c, f)?


    def __calculate_post_tax_p(
            self,
            year: int,
            energy_prices: EnergyPrices,
            p_cov_s: pd.DataFrame,
            lcoe: pd.DataFrame,
            always_exempt_res_lpg_ker: bool
            ) -> None:
        """
        'Post Tax Price (taking account of any exemptions)'
        for all subsectors.
        Prices for biofuels stored separately in post_tax_p_rod_bio.

        CPAT Excel: sector sections in scenario rows 4021:4526, baseline: rows 8816:9321 (v416)

        General 'Post Tax Price (taking account of any exemptions)' formula:
            post_tax_p(c, s, f, t) = (
                pbc(c, s, f) + nce(c, s, f) * p_cov_s(c, s)
            )

        For f == ren:
            post_tax_p(c, s, ren, t) = lcoe(c, sol, t)
        as 'Total new carbon tax/ETS (bef. sect. exempts.)' for ren is always 0.

        post_tax_p[sector] dims (c, s, f), t in <simulation_years[0], simulation_years[1]>
        post_tax_p_rod_bio dims (c, s, f), t in <simulation_years[0], simulation_years[1]>
        """
        y = str(year)

        pbc_y = energy_prices.pbc[y]
        nce_y = energy_prices.nce[y]
        p_cov_s_y = p_cov_s[y]

        p_cov_s_filtered = p_cov_s_y[
            ~p_cov_s_y.index.get_level_values(c.SECTOR_CODE).isin([c.POW, c.IND, c.TRA])
        ].sort_index() # TODO: sort index earlier?
        p_cov_s_rod_res = p_cov_s_filtered[
            p_cov_s_filtered.index.get_level_values(c.SECTOR_CODE).isin([c.ROD, c.RES])
        ]
        p_cov_s_rod = p_cov_s_rod_res[
            p_cov_s_rod_res.index.get_level_values(c.SECTOR_CODE) == c.ROD
        ].droplevel(c.SECTOR_CODE)

        pbc_other = pbc_y[
            pbc_y.index.get_level_values(c.SECTOR_CODE).isin([c.ALL, c.IND])
        ].droplevel(c.SECTOR_CODE)
        nce_other = nce_y[
            nce_y.index.get_level_values(c.SECTOR_CODE).isin([c.ALL, c.IND])
        ].droplevel(c.SECTOR_CODE)

        pbc_res = pbc_y[
            pbc_y.index.get_level_values(c.SECTOR_CODE) == c.RES
        ].droplevel(c.SECTOR_CODE)
        nce_res = nce_y[
            nce_y.index.get_level_values(c.SECTOR_CODE) == c.RES
        ].droplevel(c.SECTOR_CODE)


        # biofuels
        self.post_tax_p_rod_bio[y] = self.__calculate_post_tax_p_biofuels(
            pbc_other, nce_other, p_cov_s_rod
        )


        # calculating all values where f != ren
        for sector in SUBSECTORS_TO_SECTORS:
            mask = self.post_tax_p[sector].index.get_level_values(c.FUEL_CODE) != c.REN
            self.post_tax_p[sector].loc[mask, y] = self.__calculate_post_tax_p_other(
                pbc_other, nce_other, p_cov_s_filtered, always_exempt_res_lpg_ker, sector
            )

        # ren
        self.__calculate_post_tax_p_ren(y, lcoe[[y]])

        # overwrite [coa, nga] for rod and res
        self.__calculate_post_tax_rod_res_coa_nga(y, pbc_res, nce_res, p_cov_s_rod_res)


    def __calculate_post_tax_p_other(
            self,
            pbc_other: pd.Series,
            nce_other: pd.Series,
            p_cov_s_filtered: pd.Series,
            always_exempt_res_lpg_ker: bool,
            sector: EcSectorsType
            ) -> pd.Series:
        """
        'Post Tax Price (taking account of any exemptions)'
        for all fuels except ren and all sectors.

        [rod, res] x [coa, nga] are later overwritten
        """
        idx = pd.IndexSlice
        pbc_other.index = pd.MultiIndex.from_arrays(
            [
                pbc_other.index.get_level_values(c.COUNTRY_CODE),
                [sector] * len(pbc_other),
                pbc_other.index.get_level_values(c.FUEL_CODE)
            ],
            names=c.ID_COL_NAMES
        )
        nce_other.index = pd.MultiIndex.from_arrays(
            [
                nce_other.index.get_level_values(c.COUNTRY_CODE),
                [sector] * len(nce_other),
                nce_other.index.get_level_values(c.FUEL_CODE)
            ],
            names=c.ID_COL_NAMES
        )

        if always_exempt_res_lpg_ker and (sector == c.RES):
            nce_other.loc[idx[:, :, [c.KER, c.LPG]]] = 0.0

        p_cov_s_sector = p_cov_s_filtered[
            p_cov_s_filtered.index.get_level_values(c.SECTOR_CODE) == sector
        ].droplevel(c.SECTOR_CODE)
        nce_other = nce_other.mul(p_cov_s_sector, level=c.COUNTRY_CODE)

        return pbc_other + nce_other


    def __calculate_post_tax_p_biofuels(
            self,
            pbc_y: pd.Series,
            nce_y: pd.Series,
            p_cov_s_rod: pd.Series
            ) -> pd.DataFrame:
        """
        'Post Tax Price (taking account of any exemptions)' for fuels: bgs, bdi, obf.
        Used only for sector: rod
        """
        mapping = {c.GSO: c.BGS, c.DIE: c.BDI}
        pbc_for_rod_bio = pbc_y[pbc_y.index.get_level_values(c.FUEL_CODE).isin([c.GSO, c.DIE])]
        nce_for_rod_bio = nce_y[nce_y.index.get_level_values(c.FUEL_CODE).isin([c.GSO, c.DIE])]

        post_tax_p_bio = (
            pbc_for_rod_bio + nce_for_rod_bio.mul(p_cov_s_rod, level=c.COUNTRY_CODE)
        ).reset_index()
        post_tax_p_bio[c.SECTOR_CODE] = c.ROD
        post_tax_p_bio[c.FUEL_CODE] = post_tax_p_bio[c.FUEL_CODE].replace(mapping)

        obf_rows = post_tax_p_bio[post_tax_p_bio[c.FUEL_CODE] == c.BGS].copy()
        obf_rows[c.FUEL_CODE] = c.OBF
        return pd.concat([post_tax_p_bio, obf_rows], ignore_index=True).set_index(c.ID_COL_NAMES)


    def __calculate_post_tax_p_ren(
            self,
            y: str,
            lcoe: pd.DataFrame
            ) -> pd.DataFrame:
        """
        'Post Tax Price (taking account of any exemptions)' for fuel == ren in all sectors

        nce(ren) prices are always 0.0 in Excel, so post_tax_p are equal to lcoe(sol)
        """
        idx = pd.IndexSlice
        lcoe_y = lcoe[lcoe.index.get_level_values(c.FUEL_CODE) == c.SOL].droplevel(c.FUEL_CODE)


        for sector in SUBSECTORS_TO_SECTORS:
            lcoe_sector = lcoe_y.copy()
            lcoe_sector.index = pd.MultiIndex.from_product(
                [lcoe_y.index, [sector], [c.REN]],
                names=c.ID_COL_NAMES
            )
            self.post_tax_p[sector].loc[idx[:, :, [c.REN]], y] = lcoe_sector.copy()


    def __calculate_nce_rod_res_coa_nga(
            self,
            nce_res: pd.DataFrame,
            p_cov_s_rod_res: pd.DataFrame
            ) -> pd.DataFrame:
        """
        nce * p_cov_s component for
        'Post Tax Price (taking account of any exemptions)'
        only for fuels: coa and nga, in sectors: rod and res

        Calculated separately as for rod and res we use res prices for coa and nga.
        """
        rod_res = [c.ROD, c.RES]
        coa_nga = [c.COA, c.NGA]

        nce_res = nce_res.reindex(nce_res.index.repeat(len(rod_res)))
        nce_res.index = pd.MultiIndex.from_tuples(
            [
                (country, sector, fuel)
                for (country, fuel) in nce_res.index[::len(rod_res)]
                for sector in rod_res
            ],
            names=c.ID_COL_NAMES
        )
        p_cov_s_rod_res = p_cov_s_rod_res.reindex(p_cov_s_rod_res.index.repeat(len(coa_nga)))
        p_cov_s_rod_res.index = pd.MultiIndex.from_tuples(
            [
                (country, sector, fuel)
                for (country, sector) in p_cov_s_rod_res.index[::len(coa_nga)]
                for fuel in coa_nga
            ],
            names=c.ID_COL_NAMES
        )
        return nce_res * p_cov_s_rod_res


    def __calculate_post_tax_rod_res_coa_nga(
            self,
            y: str,
            pbc_res: pd.DataFrame,
            nce_res: pd.DataFrame,
            p_cov_s_rod_res: pd.DataFrame
            ) -> pd.DataFrame:
        """
        'Post Tax Price (taking account of any exemptions)' for [rod, res]x[coa, nga]
        """
        idx = pd.IndexSlice
        nce_res = self.__calculate_nce_rod_res_coa_nga(nce_res, p_cov_s_rod_res)
        for sector in [c.ROD, c.RES]:
            self.post_tax_p[sector].loc[idx[:, :, [c.COA, c.NGA]], y] = 0.0
            for country in self.__countries:
                pbc_res_country = pbc_res[
                    pbc_res.index.get_level_values(c.COUNTRY_CODE) == country
                ].droplevel(c.COUNTRY_CODE)
                self.post_tax_p[sector].loc[idx[[country], :, [c.COA, c.NGA]], y] = (
                    self.post_tax_p[sector].loc[idx[[country], :, [c.COA, c.NGA]], y]
                    .add(pbc_res_country, level=c.FUEL_CODE)
                )
            self.post_tax_p[sector].loc[idx[:, :, [c.COA, c.NGA]], y] += (
                nce_res.loc[idx[:, [sector], :]]
            )


    def __calculate_gdp_component(
            self,
            year: int,
            sector_agg: ElasticitiesSectorsType,
            gdp_component_c: pd.DataFrame,
            exogenous_shock_component: pd.DataFrame,
            elasticities: Elasticities
            ) -> pd.Series:
        """
        Calculates:
        (1 + d_gdp_at_const_prices(c, t)) ^ el_inc(c, s_e, f_e)
        multiplied with:
        (1 + exogenous_shock_on_ec(c, t))

        Please mind gdp_component_c = 1 + d_gdp_at_const_prices

        TODO: after Post-policy GDP is ready, include:
        'Discount factor (proportion of estimated to average calibrated income elasticity)'
        in gdp_component_c.pow(...). It will have dims (c) and unique value for each year,
        so we can just do it like: el_inc.mul(factor(t), level=c.COUNTRY_CODE)
        """
        gdp_component = (
            gdp_component_c
            .pow(
                elasticities.e[sector_agg][c.EL_INC].iloc[:, 0],
                axis=0,
                level=c.COUNTRY_CODE
            )
        )
        # we can multiply exogenous_shock_component with gdp_component at this stage:
        if year in [2023, 2024]:
            gdp_component = (
                gdp_component
                .mul(exogenous_shock_component, level=c.COUNTRY_CODE)
            )

        return gdp_component


    def __calculate_price_component(
            self,
            y: str,
            prev_y: str,
            sector: EcSectorsType,
            sector_agg: ElasticitiesSectorsType
            ) -> pd.Series:
        """
        Calculates:
        (post_tax_p(c, s, f, t) / post_tax_p(c, s, f, t-1)) ^ el_dem(c, s_e, f_e)
        """
        price_component = (
            self.post_tax_p[sector][y] / self.post_tax_p[sector][prev_y]
        ).droplevel(c.SECTOR_CODE)
        oop_ker_lpg_mask = (
            price_component.index
            .get_level_values(c.FUEL_CODE).isin([c.OOP, c.KER, c.LPG])
        )
        price_component.loc[oop_ker_lpg_mask] = (
            price_component.loc[oop_ker_lpg_mask]
            .pow(self.__e[c.EL_DEM][sector_agg][c.OIL], level=c.COUNTRY_CODE)
        )

        price_component.loc[~oop_ker_lpg_mask] = (
            price_component.loc[~oop_ker_lpg_mask]
            .pow(self.__e[c.EL_DEM][sector_agg][c.ALL])
        )
        return price_component


    def __calculate_last_component_numerator(
            self,
            y: str,
            prev_y: str,
            sector: EcSectorsType,
            sector_agg: ElasticitiesSectorsType
            ) -> pd.Series:
        """
        Calculates:
        (
            (post_tax_p(c, s, f, t) + shadow_prices(c, s, f, t) * percent_of_shadow_p(s)) 
            / (post_tax_p(c, s, f, t-1) + shadow_prices(c, s, f, t-1) * percent_of_shadow_p(s)) 
        ) ^ el_eff(c, s_e, f_e)
        """
        price_mask = (
            self.post_tax_p[sector].index
            .get_level_values(c.FUEL_CODE) != c.REN
        )

        last_component = (
            (
                self.post_tax_p[sector].loc[price_mask, y]
                + self.__shadow_prices_mul[sector][y]
            )
            / (
                self.post_tax_p[sector].loc[price_mask, prev_y]
                + self.__shadow_prices_mul[sector][prev_y]
            )
        )
        last_component = pd.concat([
            last_component,
            (
                self.post_tax_p[sector].loc[~price_mask, y]
                / self.post_tax_p[sector].loc[~price_mask, prev_y]
            )
        ]).droplevel(c.SECTOR_CODE)

        oop_ker_lpg_mask = (
            last_component.index
            .get_level_values(c.FUEL_CODE).isin([c.OOP, c.KER, c.LPG])
        )
        last_component.loc[oop_ker_lpg_mask] = (
            last_component.loc[oop_ker_lpg_mask]
            .pow(self.__e[c.EL_EFF][sector_agg][c.OIL], level=c.COUNTRY_CODE)
        )
        last_component.loc[~oop_ker_lpg_mask] = (
            last_component.loc[~oop_ker_lpg_mask]
            .pow(self.__e[c.EL_EFF][sector_agg][c.ALL])
        )

        return last_component


    def __calculate_last_component_denominator(
            self,
            sector: EcSectorsType,
            d: DashboardInputsDict,
            last_component_denominator_agg: pd.Series
            ) -> pd.Series:
        """
        Calculates:
        (1 + eff_imp(c, s_e, f_e) + additional_eff_gains(s))

        Please mind last_component_denominator_agg = 1 + 1 + eff_imp
        """
        if sector in c.SECTORS_AGG[c.TRA]:
            return (
                last_component_denominator_agg + d['additional_eff_gains'][c.TRA]
            )
        elif sector == c.RES:
            return (
                last_component_denominator_agg + d['additional_eff_gains'][c.RES]
            )
        elif (sector in c.SECTORS_AGG[c.IND]) or (sector in [c.FOO, c.SRV]):
            return (
                last_component_denominator_agg + d['additional_eff_gains'][c.IND]
            )
        else: # only for oen
            return last_component_denominator_agg.copy()


    def __calculate_last_component(
            self,
            sector_agg: ElasticitiesSectorsType,
            last_component: pd.Series,
            last_component_denominator: pd.Series
            ) -> tuple[pd.Series, pd.Series]:
        """
        Calculates:
        (last_component_numerator / last_component_denominator) ^ (1 + el_dem(c, s_e, f_e))

        Returns also last_component_denominator_oil used later for (rod, bio)
        """
        oil_denominator_mask = (
            last_component_denominator.index
            .get_level_values(c.FUEL_CODE) == c.OIL
        )
        last_component_denominator_oil = (
            last_component_denominator.loc[oil_denominator_mask]
            .droplevel(c.FUEL_CODE)
        )
        oil_fuels_mask = (
            last_component.index
            .get_level_values(c.FUEL_CODE).isin([c.OOP, c.KER, c.LPG, c.GSO, c.DIE])
        )

        last_component.loc[oil_fuels_mask] = (
            last_component.loc[oil_fuels_mask]
            .div(last_component_denominator_oil, level=c.COUNTRY_CODE)
        )
        last_component.loc[~oil_fuels_mask] = (
            last_component.loc[~oil_fuels_mask]
            .div(last_component_denominator.loc[~oil_denominator_mask])
        )

        oop_ker_lpg_mask = (
            last_component.index
            .get_level_values(c.FUEL_CODE).isin([c.OOP, c.KER, c.LPG])
        )
        last_component.loc[oop_ker_lpg_mask] = (
            last_component.loc[oop_ker_lpg_mask]
            .pow(self.__e[c.EL_DEM][sector_agg][c.OIL], level=c.COUNTRY_CODE)
        )
        last_component.loc[~oop_ker_lpg_mask] = (
            last_component.loc[~oop_ker_lpg_mask]
            .pow(self.__e[c.EL_DEM][sector_agg][c.ALL])
        )

        return last_component, last_component_denominator_oil


    def __mul_ec_diff_with_gdp_component(
            self,
            ec_diff: pd.Series,
            gdp_component: pd.Series
            ) -> pd.Series:
        """
        Calculates:
        ec_diff * gdp_component
        
        where ec_diff = price_component * last_component
        """
        oop_ker_lpg_mask = (
            ec_diff.index
            .get_level_values(c.FUEL_CODE).isin([c.OOP, c.KER, c.LPG])
        )
        oil_gdp_mask = (
            gdp_component.index.get_level_values(c.FUEL_CODE) == c.OIL
        )
        ec_diff.loc[oop_ker_lpg_mask] = (
            ec_diff.loc[oop_ker_lpg_mask]
            .mul(
                gdp_component.loc[oil_gdp_mask]
                .droplevel(c.FUEL_CODE),
                level=c.COUNTRY_CODE
            )
        )
        ec_diff.loc[~oop_ker_lpg_mask] = (
            ec_diff.loc[~oop_ker_lpg_mask]
            .mul(gdp_component.loc[~oil_gdp_mask])
        )
        return ec_diff


    def __calculate_ec_avi_jfu(
            self,
            y: str,
            prev_y: str, 
            ec_diff: pd.Series,
            gdp_component: pd.Series
            ) -> None:
        """
        'Enegy Use' for Jet Fuel in Domestic aviation

        Components used:
        ec_diff (price_component * last_component) for (avi, ker)
        gdp_component for (avi, die)
        """
        ker_mask = (
            ec_diff.index
            .get_level_values(c.FUEL_CODE) == c.KER
        )
        ec_diff_ker = ec_diff.loc[ker_mask].droplevel(c.FUEL_CODE)

        die_mask = (
            gdp_component.index
            .get_level_values(c.FUEL_CODE) == c.DIE
        )
        gdp_component_die = gdp_component.loc[die_mask].droplevel(c.FUEL_CODE)
        self.ec_avi_jfu[y] = (
            self.ec_avi_jfu[prev_y]
            .mul(ec_diff_ker, level=c.COUNTRY_CODE)
            .mul(gdp_component_die, level=c.COUNTRY_CODE)
        )


    def __calculate_ec_rod_bio(
            self,
            y: str,
            prev_y: str,
            sector: EcSectorsType,
            gdp_component: pd.Series,
            last_component_denominator_oil: pd.Series
            ) -> None:
        """
        'Enegy Use' for biofuels in Road
        """
        # CAPT Excel BUG? _el_dem_oil_rod in obf instead of _el_eff_oil_rod TODO email IMF
        # BUG? Price-based policies coverage (% of the new price), by fuel and sector
        # instead of additional_eff_gains - 'Additional policy-induced annual efficiency gains'
        # shadow prices rod, bio
        # bgs, bdi, obf gdp_component like oop # TODO move comments
        price_component_bio = (
            (self.post_tax_p_rod_bio[y] / self.post_tax_p_rod_bio[prev_y])
            .droplevel(c.SECTOR_CODE)
            .pow(self.__e[c.EL_DEM][c.TRA][c.OIL], level=c.COUNTRY_CODE)
        )
        price_component_bio = self.__get_series_with_sector_code(sector, price_component_bio)
        gdp_component_bio = gdp_component.loc[
            gdp_component.index.get_level_values(c.FUEL_CODE) == c.OIL
        ].droplevel(c.FUEL_CODE)

        # shadow prices for rod, bio
        shadow_prices_mul_y_bio = (
            self.__shadow_prices_mul[sector][y].loc[
                self.__shadow_prices_mul[sector][y].index
                .get_level_values(c.FUEL_CODE) == c.BIO
            ].droplevel(c.SECTOR_CODE).droplevel(c.FUEL_CODE)
        )
        shadow_prices_mul_prev_y_bio = (
            self.__shadow_prices_mul[sector][prev_y].loc[
                self.__shadow_prices_mul[sector][prev_y].index
                .get_level_values(c.FUEL_CODE) == c.BIO
            ].droplevel(c.SECTOR_CODE).droplevel(c.FUEL_CODE)
        )

        last_component_bio = (
            (self.post_tax_p_rod_bio[y] + shadow_prices_mul_y_bio)
            / (self.post_tax_p_rod_bio[prev_y] + shadow_prices_mul_prev_y_bio)
        ).pow(self.__e[c.EL_EFF][c.TRA][c.OIL], level=c.COUNTRY_CODE)

        # last_component_denominator_oil
        last_component_bio = (
            last_component_bio
            .div(last_component_denominator_oil, level=c.COUNTRY_CODE)
            .pow(self.__e[c.EL_DEM][c.TRA][c.OIL], level=c.COUNTRY_CODE)
        )

        self.ec_rod_bio[y] = (
            self.ec_rod_bio[prev_y].mul(gdp_component_bio, level=c.COUNTRY_CODE)
            * price_component_bio * last_component_bio
        )


    def __overwrite_ec_rod_bio(self, y: str) -> None:
        """
        Overwries 'Enegy Use' for (Biomass Road) as a sum of biofuels.
        """
        idx = pd.IndexSlice
        bio_idx = idx[:, :, [c.BIO]]


        ec_rod_bio_sum = (
            self.ec_rod_bio[y]
            .groupby(level=[c.COUNTRY_CODE])
            .sum(numeric_only=True)
            .reindex(
                self.ec[c.ROD].loc[bio_idx, y].index,
                level=c.COUNTRY_CODE
            )
        )

        self.ec[c.ROD].loc[bio_idx, y] = ec_rod_bio_sum.copy()


    def __update_ec_by_subsector(self, y: str) -> None:
        """
        'Total Energy Use by Subsector'

        ec_by_subsector[sector] dims (c, s), t in <simulation_years[0], simulation_years[1]>
        """
        idx = pd.IndexSlice
        for subsector in SUBSECTORS_TO_SECTORS:
            self.ec_by_subsector.loc[idx[:, [subsector]], y] = (
                self.ec[subsector][y]
                .groupby(level=[c.COUNTRY_CODE, c.SECTOR_CODE])
                .sum(numeric_only=True)
            )
            if subsector == c.AVI:
                self.ec_by_subsector.loc[idx[:, [c.AVI]], y] += (
                    self.ec_avi_jfu[y].droplevel(c.FUEL_CODE)
                )


    def __update_ec_fossil_fuels_by_sector(self, y: str) -> None:
        """
        'Fossil fuel consumption by sector'
        from
        'Total energy consumption (TEC)'

        CPAT Excel: baseline rows 4817:4859, scenario rows 9612:9654 (v416)

        ec_fossil_fuels_by_sector[sector] dims (c, s, f),
            t in <simulation_years[0], simulation_years[1]>
        """
        idx = pd.IndexSlice

        for sector, subsectors in c.SECTORS_AGG_TMP.items():
            if sector not in [c.POW, c.OEN]:
                s_list: list[pd.Series] = []
                for subsector in subsectors:
                    # print(self.ec[subsector])
                    s_list.append(self.ec[subsector].loc[idx[:, :, c.FOSSIL_FUELS], y])

                s = (
                    pd.concat(s_list)
                    .groupby(level=[c.COUNTRY_CODE, c.FUEL_CODE])
                    .sum()
                )
                s.index = pd.MultiIndex.from_arrays(
                    [
                        s.index.get_level_values(c.COUNTRY_CODE),
                        [sector] * len(s),
                        s.index.get_level_values(c.FUEL_CODE),
                    ],
                    names=c.ID_COL_NAMES,
                )
                self.ec_fossil_fuels_by_sector.loc[idx[:, [sector], :], y] += s

        self.ec_fossil_fuels_by_sector.loc[idx[:, [c.OEN], :], y] += (
            self.ec[c.OEN].loc[idx[:, :, c.FOSSIL_FUELS], y]
        )


    def calculate_ec_year(
            self,
            year: int,
            d_gdp_at_const_prices: pd.DataFrame, # GDP.d_gdp_at_const_prices
            energy_prices: EnergyPrices,
            p_cov_s: pd.DataFrame, # Policies.p_cov_s
            elasticities: Elasticities,
            lcoe: pd.DataFrame, # TODO: for now we use lcoe_tmp
            d: DashboardInputsDict,
            exogenous_shock_on_ec: pd.DataFrame,
            shadow_prices: ShadowPrices
        ) -> None:
        """
        Calcualtes values for a given year.
        """
        y = str(year)
        prev_y = str(year - 1)

        self.__calculate_post_tax_p(
            year, energy_prices, p_cov_s, lcoe, d['always_exempt_res_lpg_ker']
        )

        if year > self.__init_year:# + 1: # TODO: change to > and remove +1

            self.__update_shadow_prices_mul(y, shadow_prices)

            # these do not depend only on country (and year):
            gdp_component_c = 1 + d_gdp_at_const_prices[y]
            if year in [2023, 2024]:
                exogenous_shock_component = 1 + exogenous_shock_on_ec[y]
            else:
                exogenous_shock_component = pd.DataFrame() # TODO:

            for sector_agg, sectors in EC_SECTORS.items():
                gdp_component = self.__calculate_gdp_component(
                    year, sector_agg, gdp_component_c, exogenous_shock_component, elasticities
                )

                last_component_denominator_agg = (
                    1 + elasticities.e[sector_agg][c.EFF_IMP].iloc[:, 0]
                )

                for sector in sectors:
                    price_component = self.__calculate_price_component(
                        y, prev_y, sector, sector_agg
                    )

                    ### last component:
                    last_component = self.__calculate_last_component_numerator(
                        y, prev_y, sector, sector_agg
                    )
                    last_component_denominator = self.__calculate_last_component_denominator(
                        sector, d, last_component_denominator_agg
                    )
                    (
                        last_component,
                        last_component_denominator_oil # used later for rod, bio
                    ) = self.__calculate_last_component(
                        sector_agg, last_component, last_component_denominator
                    )

                    ec_diff = price_component * last_component

                    if sector == c.AVI:
                        self.__calculate_ec_avi_jfu(y, prev_y, ec_diff, gdp_component)

                    ec_diff = self.__mul_ec_diff_with_gdp_component(ec_diff, gdp_component)

                    # adding SectorCode again:
                    ec_diff = self.__get_series_with_sector_code(sector, ec_diff)

                    self.ec[sector][y] = (
                        self.ec[sector][prev_y]
                        * ec_diff
                    )


                    if sector == c.ROD:
                        self.__calculate_ec_rod_bio(
                            y, prev_y, sector, gdp_component,
                            last_component_denominator_oil
                        )
                        self.__overwrite_ec_rod_bio(y)

        self.__update_ec_by_subsector(y)
        self.__update_ec_fossil_fuels_by_sector(y)
