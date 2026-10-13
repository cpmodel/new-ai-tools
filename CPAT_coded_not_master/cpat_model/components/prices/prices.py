from typing import Literal

import pandas as pd
import numpy as np

from cpat_model.inputs.dashboard_inputs import DashboardInputsDict
from cpat_model.inputs.input_data import InputData
from cpat_model.components.energy_consumption.energy_data import TOTAL_ENERGY_USE_COL

from cpat_model.components.prices.domestic_prices import Headers, DomPrices
from cpat_model.components.prices.international_prices import IntPrices
from cpat_model.components.gdp.gdp import GDP
from cpat_model.components.prices.ntx import NTX
from cpat_model.components.prices.ctx import CTX
from cpat_model.components.prices.ets import ETS
from cpat_model.components.prices.ctxnew import CTXNew

from cpat_model.components.policies.policies import Policies
from cpat_model.components.policies.phaseouts import Phaseouts
from cpat_model.components.efs.co2 import EFsCO2
from cpat_model.components.carbon_pricing.existing_ct import ExistingCT
from cpat_model.components.carbon_pricing.existing_ets import ExistingETS

import cpat_model.constants as c

# TODO: Unify conv
CONV_KTOE_TO_GJ = 41_868
CONV_KTOE_LITER_OIL = 1_108_792

ElePricesType = Literal['rp', 'sp', 'txo', 'vatrate']

class EnergyPrices:
    """
    EnergyPrices
    'Retail energy prices' section (v361)
    CPAT Excel: baseline 2009:2272, scenario 6999:7262

    Methods:
    __init_{price} - inits for historical years and sets 0.0
        for rest of the simulation years
    __set_{price} - inits for historical years and calculates
        for rest of the simulation years
    __update_{price} - updates years in calculate_prices_year

    Changes from Excel:
    - only nce and pbs from 'Info:' rows
    """
    rp: pd.DataFrame
    sp: pd.DataFrame
    fixsp: pd.DataFrame
    fltsp: pd.DataFrame
    ps: pd.DataFrame
    vatrate: pd.DataFrame
    vat: pd.DataFrame
    txo: pd.DataFrame
    txo_hist: pd.DataFrame
    fadtx: pd.DataFrame
    fixs: pd.DataFrame
    cs: pd.DataFrame
    ctx: pd.DataFrame
    ets: pd.DataFrame
    ctxnew: pd.DataFrame
    ntx: pd.DataFrame
    nce: pd.DataFrame
    pbc: pd.DataFrame
    ele_prices: dict[ElePricesType, pd.DataFrame]

    def __init__(
            self,
            simulation_years: tuple[int, int],
            scenario_type: str,
            selected_countries: list[str],
            gdp: GDP,
            dom_prices: DomPrices,
            phaseouts: Phaseouts,
            existing_ct: ExistingCT,
            existing_ets: ExistingETS,
            efs_co2: EFsCO2,
            int_prices: IntPrices,
            input_data: InputData,
            d: DashboardInputsDict
            ) -> None:
        """
        Calculates values up to last year in historical prices for all attributes.
        Calculates rest of the years for attributes not dependent on scenario type/scenario inputs.
        """
        self.__set_helpers(simulation_years, dom_prices.rp)
        self.__set_total_territorial_ec(input_data.ec_input)

        self.sp = self.__init_sp(dom_prices.sp, gdp.inflation_index_100)
        self.txo_hist = self.__set_txo_hist(dom_prices.txo, gdp.inflation_index_100)
        self.vatrate = self.__set_vatrate(dom_prices.vatrate) # TODO: vatrate oop 0?
        self.rp = self.__init_rp(
            dom_prices.rp, gdp.inflation_index_100, input_data.oil_product_shares
        )

        self.ps = self.__set_ps(
            phaseouts.phaseout_producer, dom_prices.ps, gdp.inflation_index_100,
            int_prices.global_fuel_demand # TODO check test
        )
        self.fixsp = self.__set_fixsp(dom_prices)
        self.fltsp = self.__set_fltsp(int_prices.prices)

        self.vat = self.__init_vat()
        self.txo = self.__init_txo()

        self.ntx = NTX.calculate_ntx(
            scenario_type, selected_countries, simulation_years,
            gdp.deflator, input_data, d
        )
        self.ctx = CTX.get_ctx(existing_ct, efs_co2.ef_tco2_per_volume_unit)
        self.ets = ETS.get_ets(existing_ets, efs_co2.ef_tco2_per_volume_unit)
        self.__filter_ntx_ctx_ets() # TODO: test

        # TODO: test all below
        self.__init_ctxnew()
        self.cs = self.__init_cs(
            dom_prices.forecasting_coefficients[Headers.CHOSEN_PRICE_CONTROL_COEFFICIENT]
        )
        self.__init_fadtx_and_fixs()
        self.__no_floating_subsidy = pd.DataFrame() # just for linter

        self.__init_ele_prices(dom_prices, gdp.inflation_index_100)

        # 'Info' prices:
        self.__init_nce()
        self.__init_pbc()


    def __set_helpers(
            self,
            simulation_years: tuple[int, int],
            rp: pd.DataFrame
            ) -> None:
        """
        Sets helper attributes.
        """
        self.__hist_cols = rp.columns.copy()
        self.__last_hist_year = max(map(int, rp.columns))
        self.__first_hist_year = min(map(int, rp.columns))
        self.__new_years_cols = [
            str(y) for y in range(self.__last_hist_year + 1, simulation_years[1] + 1)
        ]
        self.__base_year_to_new_years_cols = [
            str(y) for y in range(simulation_years[0], self.__last_hist_year + 1)
        ]
        self.__year_prior_to_base_year = str(simulation_years[0] - 1)


    def __set_total_territorial_ec(
            self,
            ec_input: pd.DataFrame
            ) -> None:
        """
        Extracting 'Total energy-use territorial consumption' values for
        bio, coa, nga and ele from Energy Consumption input table.
        Used to calculate ps values.
        """
        idx = pd.IndexSlice
        total_territorial_ec = (
            ec_input.loc[idx[:, [TOTAL_ENERGY_USE_COL]], [c.BIO, c.COA, c.NGA, c.ELE]]
            .reset_index()
        )
        total_territorial_ec[c.SECTOR_CODE] = c.ALL
        self.__total_territorial_ec = pd.melt(
            total_territorial_ec,
            id_vars=[c.COUNTRY_CODE, c.SECTOR_CODE],
            value_vars=[c.BIO, c.COA, c.NGA, c.ELE],
            var_name=c.FUEL_CODE,
            value_name='Value'
        ).set_index(c.ID_COL_NAMES).sort_index()


    def __init_rp(
            self,
            rp_hist: pd.DataFrame,
            inflation_index_100: pd.DataFrame,
            oil_product_shares: pd.DataFrame
            ) -> pd.DataFrame:
        """
        'Retail price', historical years
        """
        rp = rp_hist.copy()
        # ele in ele_prices
        rp = rp[rp.index.get_level_values(c.FUEL_CODE) != c.ELE]

        # coa, nga, bio: * inflation_index_100
        idx = pd.IndexSlice
        mask = idx[:, :, [c.COA, c.NGA, c.BIO]]
        rp.loc[mask, :] = rp.loc[mask, :].mul(inflation_index_100.iloc[0], axis=1)

        # gso, die, lpg, ker -> (sp+txo)*(1+vatrate*share),
        # oop -> (sp+txo)*(1+vatrate);
        # (vatrate for oop is hardcoded as 0 in v361 so no *(1+vatrate))
        mask = idx[:, :, [c.GSO, c.DIE, c.KER, c.LPG, c.OOP]]
        rp.loc[mask, :] = self.sp.loc[mask, :] + self.txo_hist.loc[mask, :]

        # apply vatrate*share
        mask = idx[:, :, [c.GSO, c.DIE, c.KER, c.LPG]]
        vatrate_share = (
            self.vatrate.loc[mask, self.__hist_cols]
            .mul(oil_product_shares['Value'], axis=0)
        )
        rp.loc[mask, :] = rp.loc[mask, :] * (1 + vatrate_share)

        rp[self.__new_years_cols] = 0.0
        return rp

    def __init_sp(
            self,
            sp_hist: pd.DataFrame,
            inflation_index_100: pd.DataFrame
            ) -> pd.DataFrame:
        """
        'Retail price', historical years
        """
        sp = sp_hist.copy()
        # ele in ele_prices
        sp = sp[sp.index.get_level_values(c.FUEL_CODE) != c.ELE]
        sp = sp.mul(inflation_index_100.iloc[0], axis=1)
        sp[self.__new_years_cols] = 0.0

        return sp

    def __set_txo_hist(
            self,
            txo: pd.DataFrame,
            inflation_index_100: pd.DataFrame
            ) -> pd.DataFrame:
        """
        'Excise and other taxes', historical years
        from 'Historical prices' table:
        CPAT Excel 731:792 (v361)

        Used only to calculate historical year for rp.
        """
        txo_hist = txo.copy()
        # ele in ele_prices
        txo_hist = txo_hist[txo_hist.index.get_level_values(c.FUEL_CODE) != c.ELE]
        txo_hist = txo_hist.mul(inflation_index_100.loc[0, self.__hist_cols], axis=1)

        return txo_hist

    def __set_ps(
            self,
            phaseout_producer: pd.DataFrame,
            ps_hist: pd.DataFrame,
            inflation_index_100: pd.DataFrame,
            global_fuel_demand: pd.DataFrame
            ) -> pd.DataFrame:
        """
        '- in which: producer-side subsidies', all years
        """
        # TODO: check test, correct nga,coa all
        ps = ps_hist.copy()
        # ele in ele_pricess
        ps = ps[ps.index.get_level_values(c.FUEL_CODE) != c.ELE]

        ps = - ps.mul(inflation_index_100.loc[0, self.__hist_cols], axis=1)

        idx = pd.IndexSlice
        # f in [coa, nga] -> /= (
        #           (Global demand(f, t)*_conv_ktoe_to_gj/10^9*0.5)
        #           + (Total energy-use territorial consumption(f)*_conv_ktoe_to_gj/10^9*0.5)
        # )
        mask = idx[:, :, [c.COA, c.NGA]]
        # calculating denominator values:
        coa_nga_demand = (
            global_fuel_demand.loc[[c.COA, c.NGA], self.__hist_cols]
            * CONV_KTOE_TO_GJ / 1e9 * 0.5
        )
        coa_nga_demand = coa_nga_demand.reindex(ps.loc[mask, :].index, level=c.FUEL_CODE)
        coa_nga_ec = (
            self.__total_territorial_ec.loc[mask, ['Value']]
            * CONV_KTOE_TO_GJ / 1e9 * 0.5
        )
        coa_nga_demand = coa_nga_demand.add(coa_nga_ec['Value'], axis=0)

        ps.loc[mask, :] = (
            ps.loc[mask, :]
            .div(coa_nga_demand, axis=1)
        )

        # oil -> /= (Global demand(oil, t)*_conv_ktoe_liter_oil/10^9)
        mask = idx[:, :, [c.OIL]]
        ps.loc[mask, :] = (
            ps.loc[mask, :]
            .div(
                global_fuel_demand.loc[c.OIL, self.__hist_cols] * CONV_KTOE_LITER_OIL / 1e9,
                axis=1
            )
        )

        # bio -> /= (Total energy-use territorial consumption*_conv_ktoe_to_gj/10^9)
        mask = idx[:, :, [c.BIO]]
        ps.loc[mask, :] = (
            ps.loc[mask, :]
            .div(
                self.__total_territorial_ec.loc[mask, ['Value']].values * CONV_KTOE_TO_GJ / 1e9,
                axis=0
            )
        )

        # cast coa, nga to 3 sectors
        new_rows = []
        for fuel in [c.COA, c.NGA]:
            # Step 1: Select rows to transform
            mask = (
                (ps.index.get_level_values(c.SECTOR_CODE) == c.ALL)
                & (ps.index.get_level_values(c.FUEL_CODE) == fuel)
            )
            rows_to_copy = ps[mask]

            new_sectors = [c.IND, c.POW, c.RES]
            for sector in new_sectors:
                new_index = [
                    (country, sector, fuel)
                    for (country, _, _) in rows_to_copy.index
                ]
                new_row = rows_to_copy.copy()
                new_row.index = pd.MultiIndex.from_tuples(new_index, names=ps.index.names)
                new_rows.append(new_row)

        # drop original nga, coa:
        mask = (
                (ps.index.get_level_values(c.SECTOR_CODE) == c.ALL)
                & (ps.index.get_level_values(c.FUEL_CODE).isin([c.COA, c.NGA]))
            )
        ps = ps[~mask]

        # cast oil to [c.GSO, c.DIE, c.KER, c.LPG, c.OOP]
        mask = (
            (ps.index.get_level_values(c.SECTOR_CODE) == c.ALL)
            & (ps.index.get_level_values(c.FUEL_CODE) == c.OIL)
        )
        rows_to_copy = ps[mask]
        new_fuels = [c.GSO, c.DIE, c.KER, c.LPG, c.OOP]

        new_index = pd.MultiIndex.from_tuples(
            [
                (country_code, sector, fuel)
                for (country_code, sector, _) in rows_to_copy.index
                for fuel in new_fuels
            ],
            names=ps.index.names
        )
        new_oil_rows = pd.DataFrame(
            rows_to_copy.to_numpy().repeat(len(new_fuels), axis=0),
            index=new_index,
            columns=ps.columns
        )

        # drop original oil:
        ps = ps[~mask]
        ps = pd.concat([ps, *new_rows, new_oil_rows]).sort_index()

        for col in self.__new_years_cols:
            ps[col] = ps[str(self.__last_hist_year)] * phaseout_producer.loc[0, str(col)]

        return ps

    def __set_fixsp(
            self,
            dom_prices: DomPrices
            ) -> pd.DataFrame:
        """
        'fixed portion of supply costs', all years
        TODO: can be just 1 column,
        and then we can do (...).sub(fixsp.iloc[:, -1], level=c.COUNTRY_CODE) (or sum(...))
        """
        fixsp = (
            dom_prices.forecasting_coefficients[Headers.MARGIN_ON_TOP_OF_SUPPLY_COST]
            .rename(columns={Headers.MARGIN_ON_TOP_OF_SUPPLY_COST: str(self.__first_hist_year)})
        )
        # only for coa, nga
        idx = pd.IndexSlice
        mask = idx[:, :, [c.COA, c.NGA]]
        fixsp.loc[mask, str(self.__first_hist_year)] = (
            fixsp.loc[mask, str(self.__first_hist_year)]
            + dom_prices.forecasting_coefficients[Headers.BUCKETED_PRICE_CONTROL_COEFFICIENT].loc[
                mask, Headers.BUCKETED_PRICE_CONTROL_COEFFICIENT
            ]
            * dom_prices.forecasting_coefficients[Headers.DOMESTIC_PRODUCTION_COST].loc[
                mask, Headers.DOMESTIC_PRODUCTION_COST
            ]
        )

        # rest of historic years and simulation years
        for col in self.__hist_cols:
            if col != str(self.__first_hist_year):
                fixsp[col] = fixsp[str(self.__first_hist_year)]
        for col in self.__new_years_cols:
            fixsp[col] = fixsp[str(self.__last_hist_year)]
        return fixsp


    def __set_fltsp(self, int_prices: pd.DataFrame) -> pd.DataFrame:
        """
        'floating portion of supply costs', all years
        """
        # TODO: write test
        fltsp = self.sp[self.__hist_cols] - self.fixsp[self.__hist_cols] - self.ps[self.__hist_cols]
        idx = pd.IndexSlice

        countries = fltsp.index.get_level_values(c.COUNTRY_CODE).unique()
        oil_fuels = [c.GSO, c.DIE, c.KER, c.LPG, c.OOP]
        oil_index = pd.MultiIndex.from_product(
            [countries, [c.ALL], oil_fuels],
            names=c.ID_COL_NAMES
        )
        coa_index = pd.MultiIndex.from_product(
            [countries, [c.IND, c.POW, c.RES], [c.COA]],
            names=c.ID_COL_NAMES
        )
        nga_index = pd.MultiIndex.from_product(
            [countries, [c.IND, c.POW, c.RES], [c.NGA]],
            names=c.ID_COL_NAMES
        )

        for col in self.__new_years_cols:
            previous_year = str(int(col) - 1)

            # bio copied from last year
            fltsp[col] = fltsp[previous_year]

            oil_values = (
                int_prices.loc[idx[:, c.OIL], col] / int_prices.loc[idx[:, c.OIL], previous_year]
            ).to_frame().droplevel(c.FUEL_CODE)
            oil_values = oil_values.reindex(oil_index, level=c.COUNTRY_CODE)
            oil_values.columns = [col]

            coa_values = (
                int_prices.loc[idx[:, c.COA], col] / int_prices.loc[idx[:, c.COA], previous_year]
            ).to_frame().droplevel(c.FUEL_CODE)
            coa_values = coa_values.reindex(coa_index, level=c.COUNTRY_CODE)
            coa_values.columns = [col]

            nga_values = (
                int_prices.loc[idx[:, c.NGA], col] / int_prices.loc[idx[:, c.NGA], previous_year]
            ).to_frame().droplevel(c.FUEL_CODE)
            nga_values = nga_values.reindex(nga_index, level=c.COUNTRY_CODE)
            nga_values.columns = [col]

            fltsp.loc[idx[:, :, c.COA], [col]] *= coa_values.loc[:, [col]]
            fltsp.loc[idx[:, :, c.NGA], [col]] *= nga_values.loc[:, [col]]
            fltsp.loc[idx[:, :, oil_fuels], [col]] *= oil_values.loc[:, [col]]


        return fltsp

    def __set_vatrate(
            self,
            vatrate_hist: pd.DataFrame
            ) -> pd.DataFrame:
        """
        'VAT rate', all years
        """
        vatrate = vatrate_hist.copy()
        # ele in ele_prices
        vatrate = vatrate[vatrate.index.get_level_values(c.FUEL_CODE) != c.ELE]
        for col in self.__new_years_cols:
            vatrate[col] = vatrate[str(self.__last_hist_year)]
        return vatrate

    def __init_vat(self) -> pd.DataFrame:
        """
        'VAT payment', historical years
        """
        vat = self.rp / (1 + self.vatrate[self.__hist_cols]) * self.vatrate[self.__hist_cols]
        vat[self.__new_years_cols] = 0.0
        return vat

    def __init_txo(self) -> pd.DataFrame:
        """
        'Excise and other taxes', historical years
        """
        # TODO test
        txo = self.rp - self.sp - self.vat
        txo[self.__new_years_cols] = 0.0
        return txo

    def __init_ctxnew(self):
        """
        Init ctxnew for simulation_years[0] - 1 with all 0s.
        """
        # TODO: test
        self.ctxnew = pd.DataFrame(
            0.0,
            index=self.rp.index,
            columns=[
                self.__year_prior_to_base_year,
                *self.__base_year_to_new_years_cols,
                *self.__new_years_cols
            ]
        )


    # TODO: all below chenge to loop
    def __filter_ntx_ctx_ets(self) -> None:
        """
        TODO: and test
        """
        self.ntx = self.ntx[self.ntx.index.get_level_values(c.FUEL_CODE) != c.ELE]
        self.ctx = self.ctx[self.ctx.index.get_level_values(c.FUEL_CODE) != c.JFU]
        self.ets = self.ets[self.ets.index.get_level_values(c.FUEL_CODE) != c.JFU]

    def __init_fadtx_and_fixs(self) -> None:
        """
        TODO
        """
        # TODO: test
        y = self.__year_prior_to_base_year
        fixed_portion = (
            self.txo[[y]] - (
                self.cs[[y]] + self.ctx[[y]] + self.ets[[y]]
                + self.ctxnew[[y]] + self.ntx[[y]]
            )
        )

        self.fadtx = fixed_portion.clip(lower=0) # max(value, 0)
        self.fixs = fixed_portion.clip(upper=0) # min(value, 0)

        self.fadtx[self.__base_year_to_new_years_cols + self.__new_years_cols] = 0.0
        self.fixs[self.__base_year_to_new_years_cols + self.__new_years_cols] = 0.0


    def __init_cs(
            self,
            chosen_price_control_coefficient: pd.DataFrame
            ) -> pd.DataFrame:
        """
        '- in which: floating portion of subsidy/tax', historical years
        """
        # TODO test
        y = self.__year_prior_to_base_year
        cs = (
            self.txo[[y]] - self.ctx[[y]] - self.ets[[y]]
            - self.ctxnew[[y]] - self.ntx[[y]]
        )
        # TODO: do it more efficient
        mask = chosen_price_control_coefficient.index.get_level_values(c.FUEL_CODE) != c.ELE
        cs = cs.mul(
            (
                1 - chosen_price_control_coefficient[mask]
            ).iloc[:, 0],
            axis=0
        )
        cs[self.__base_year_to_new_years_cols + self.__new_years_cols] = 0.0

        return cs

    def __set_no_floating_subsidy(self) -> None:
        """
        'no floating subsidy'
        """
        #TODO: test
        year = str(self.__last_hist_year)
        self.__no_floating_subsidy = self.cs[[year]].copy()
        # set to True if >= 0, set False otherwise
        self.__no_floating_subsidy[year] = self.__no_floating_subsidy[year] >= 0


    def __init_ele_prices(
            self,
            dom_prices: DomPrices,
            inflation_index_100: pd.DataFrame
        ) -> None:
        """
        'Historical prices' for Electricity
        CPAT Excel: rows 757:764 (v361)
        TODO test
        """
        self.ele_prices = {}
        self.ele_prices['rp'] = dom_prices.rp[
            dom_prices.rp.index.get_level_values(c.FUEL_CODE) == c.ELE
        ].mul(inflation_index_100[dom_prices.rp.columns].iloc[0], axis=1)

        self.ele_prices['sp'] = dom_prices.sp[
            dom_prices.sp.index.get_level_values(c.FUEL_CODE) == c.ELE
        ].mul(inflation_index_100[dom_prices.sp.columns].iloc[0], axis=1)

        self.ele_prices['txo'] = dom_prices.txo[
            dom_prices.txo.index.get_level_values(c.FUEL_CODE) == c.ELE
        ].mul(inflation_index_100[dom_prices.txo.columns].iloc[0], axis=1)

        self.ele_prices['vatrate'] = dom_prices.vatrate[
            dom_prices.vatrate.index.get_level_values(c.FUEL_CODE) == c.ELE
        ]


    def __init_nce(self) -> None:
        """
        'Info: total new policy (bef. sect. exempts.)'
        (v407)

        Inits with 0s as ntx and ctxnew new are always 0s at simulation_years[0] - 1
        """
        self.nce = self.ctxnew.copy()


    def __init_pbc(self) -> None:
        """
        'Info: price before any new policies (carbon tax/ETS)'
        (v407)
        """
        # TODO: test
        self.pbc = self.nce.copy()

        # to - 1 year
        first_y = str(self.__first_hist_year)
        self.pbc[first_y] += self.rp.loc[:, first_y]


    def __update_nce(self, y: str) -> None:
        """
        Calculates from simulation_years[0]
        TODO test
        """
        self.pbc[y] = (self.ctxnew[y] + self.ntx[y]) * (1 + self.vatrate[y])

    def __update_pbc(self, y: str) -> None:
        """
        Calculates from simulation_years[0]
        TODO test
        """
        self.pbc[y] = self.nce[y] + self.rp[y]


    def __update_sp(self, y: str) -> None:
        """
        TODO
        """
        self.sp[y] = self.fixsp[y] + self.fltsp[y] + self.ps[y]


    def __update_cs(
            self,
            year: int,
            phaseout_subsidy_tax: pd.DataFrame,
            chosen_price_control_coefficient: pd.DataFrame
            ) -> None:
        """
        TODO
        """
        y = str(year)

        if year <= self.__last_hist_year:
            self.cs[y] = (
                self.txo[y] - self.ctx[y] - self.ets[y]
                - self.ctxnew[y] - self.ntx[y]
            )
            # TODO: do it more efficient
            mask = chosen_price_control_coefficient.index.get_level_values(c.FUEL_CODE) != c.ELE
            self.cs[y] = self.cs[y].mul(
                (
                    1 - chosen_price_control_coefficient[mask]
                ).iloc[:, 0],
                axis=0
            )
            if year == self.__last_hist_year:
                self.__set_no_floating_subsidy()
        else:
            last_y = str(self.__last_hist_year)
            self.cs[y] = (
                (self.sp[last_y] - self.sp[y] + self.cs[last_y])
                .mul( # Check TODO
                    (1 - chosen_price_control_coefficient).iloc[:, 0],
                    axis=0
                )
            )
            self.cs[y] = np.where(
                # condition: True/False from no_floating_subsidy:
                self.__no_floating_subsidy[last_y],
                # True -> take max of last_hist_year and y:
                self.cs[[last_y, y]].max(axis=1),
                self.cs[y] * phaseout_subsidy_tax.loc[0, y] # False -> multiply by phaseout
            )


    def __update_fadtx_and_fixs(
            self,
            year: int,
            phaseout_consumer: pd.DataFrame
            ) -> None:
        """
        '- in which: fixed portion of tax' TODO and test
        '- in which: fixed portion of subsidy'
        """
        y = str(year)
        if year <= self.__last_hist_year:
            fixed_portion = (
                self.txo[[y]] - (
                    self.cs[[y]] + self.ctx[[y]] + self.ets[[y]]
                    + self.ctxnew[[y]] + self.ntx[[y]]
                )
            )

            self.fadtx[y] = fixed_portion.clip(lower=0) # max(value, 0)
            self.fixs[y] = fixed_portion.clip(upper=0) # min(value, 0)
        else:
            self.fadtx[y] = self.fadtx[str(self.__last_hist_year)]
            self.fixs[y] = self.fixs[str(self.__last_hist_year)] * phaseout_consumer.loc[0, str(y)]

    def __update_txo(self, y: str) -> None:
        """
        TODO
        """
        self.txo[y] = (
            self.fadtx[y] + self.fixs[y] + self.cs[y]
            + self.ctx[y] + self.ets[y] + self.ctxnew[y] + self.ntx[y]
        )


    def __update_vat(self, y: str) -> None:
        """
        TODO
        """
        self.vat[y] = (self.txo[y] + self.sp[y]) * self.vatrate[y]


    def __update_rp(self, y: str) -> None:
        """
        TODO
        """
        # # max (value, 0.01):
        self.rp[y] = (self.sp[y] + self.vat[y] + self.txo[y]).clip(lower=0.01)


    def calculate_prices_year(
            self,
            year: int,
            scenario_type: str,
            policies: Policies,
            efs_co2: EFsCO2,
            dom_prices: DomPrices,
            phaseouts: Phaseouts
        ) -> None:
        """
        Calculates ctxnew, cs, fadtx, fixs for each of the simulation year.
        Calculates values for sp, txo, vat and rp, for years
        after last historical year.
        """
        y = str(year)
        self.ctxnew[y] = CTXNew.get_ctxnew(
            year, scenario_type, policies.cp_trajectory,
            policies.p_based_policies_cov_s_f[[y]],
            efs_co2.ef_tco2_per_volume_unit
        )
        if year > self.__last_hist_year:
            self.__update_sp(y)

        self.__update_cs(
            year, phaseouts.phaseout_subsidy_tax,
            dom_prices.forecasting_coefficients[Headers.CHOSEN_PRICE_CONTROL_COEFFICIENT]
        )
        self.__update_fadtx_and_fixs(year, phaseouts.phaseout_consumer)

        if year > self.__last_hist_year:
            self.__update_txo(y)
            self.__update_vat(y)
            self.__update_rp(y)

        # 'Info'
        self.__update_nce(y)
        self.__update_pbc(y)
