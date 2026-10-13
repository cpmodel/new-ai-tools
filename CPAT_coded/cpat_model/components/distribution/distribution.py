from typing import TYPE_CHECKING

import pandas as pd

import cpat_model.constants as c
from cpat_model.inputs.distribution_inputs import DistributionInputsDict
from cpat_model.inputs.input_data import InputData

from cpat_model.components.distribution import price_changes, effects, rebasing, recycling, welfare, sectoral
from cpat_model.components.distribution import budget_shares as bsh_mod

if TYPE_CHECKING:
    from cpat_model.scenario_results import ScenarioResults


DECILES = list(range(1, 11))


class Distribution:
    """
    Household distributional (incidence) analysis of a carbon-pricing policy
    scenario, for one analysis year and one country.

    CPAT Excel: 'Distribution' sheet (sections A-F, rows 1-6732).
    Full spec: distribution/docs/CPAT_Distribution_Module_Pseudocode.docx.
    Per-step implementation: price_changes.py (Step 1), budget_shares.py
    (Steps 2-3), effects.py (Steps 4-6), rebasing.py (Step 7), recycling.py
    (Step 8), welfare.py (Steps 9-12), sectoral.py (Step 13). This class is
    Step 14: it wires the others together and assembles the results table.

    Unlike every other component in cpat_model.components, Distribution is
    not stepped year-by-year inside run_model's main loop: it is a single-
    year, single-country incidence analysis. See
    cpat_model.scenario_results.ScenarioResults, and
    price_changes.derive_price_change_direct_from_scenarios /
    derive_cp_revenue_from_scenarios, for how its two real inputs
    (price_change_direct, cp_revenue) are produced from a Mitigation-module
    run when one is available; run_model.py wires that up after its main
    per-year loop, once per non-baseline scenario.

    Known data gaps (see each step module's docstring for the specific
    ones): sectoral pass-through coefficients and IEA/GAINS emissions
    recalibration (Step 1.4/1.5), the Pm/Pd revenue reconciliation (Step
    1.3), and absolute (LCU) PIT/national-accounts baselines for the
    revenue-recycling caps (Step 8) and rebasing (Step 7) are Mitigation-
    module data not present in cpat_excel/Distribution/data_standardized.
    Where a formula needs one of these, it either takes the missing figure
    as an optional parameter (national_totals below) or falls back to a
    share-based approximation documented in the relevant step module.
    """
    country: str
    analysis_year: int
    sample: str
    stat_type: str

    price_change_direct: pd.Series
    price_change_indirect: pd.Series
    cp_revenue: float

    budget_shares: pd.DataFrame
    elasticities: pd.DataFrame
    direct_effect: pd.DataFrame
    indirect_effect: pd.DataFrame
    total_effect: pd.DataFrame

    rebased: pd.DataFrame
    amount_recycled: pd.Series
    net_effect: pd.Series
    post_policy: pd.DataFrame
    compensation: pd.DataFrame

    gini: dict
    delta_gini_excl_recycling: float
    delta_gini_incl_recycling: float

    def __init__(
            self,
            d: DistributionInputsDict,
            country: str,
            input_data: InputData,
            price_change_direct_priced: pd.Series,
            cp_revenue: float,
            national_totals: dict | None = None,
            sample: str = 'Overall',
            ) -> None:
        """
        price_change_direct_priced: Series indexed by the 8 priced fuel
        codes (c.COA/ELE/NGA/OOP/GSO/DIE/KER/LPG), fraction (0.5 = 50%).
        Either price_changes.derive_price_change_direct_from_scenarios(...)
        output, or a known value (e.g. read from the Excel model directly,
        for validation -- see cpat_testing).
        cp_revenue: total carbon-pricing revenue in the analysis year, LCU.
        Either price_changes.derive_cp_revenue_from_scenarios(...) output,
        or a known/assumed value.
        national_totals: optional {'population': float, 'consumption': float}
        for Step 7's rebasing target (see rebasing.py module docstring for
        why this isn't derived internally). None falls back to the
        household survey's own implied totals (population.sum(),
        (population * per_capita_consumption).sum()) -- i.e. no national-
        accounts adjustment is applied, just the survey's raw scale.
        sample: 'Overall', 'Urban' or 'Rural'.
        """
        self.d = d
        self.country = country
        self.analysis_year = d['analysis_year']
        self.sample = sample
        self.stat_type = d['statistic_output']
        self.cp_revenue = cp_revenue

        self.__init_price_changes(input_data, price_change_direct_priced)
        self.__init_budget_shares_and_elasticities(input_data)
        self.__init_effects(d)
        self.__init_rebasing(input_data, national_totals)
        self.__init_recycling(d, input_data)
        self.__init_welfare(d)
        self.__init_sectoral(input_data)


    def __init_price_changes(self, input_data: InputData, price_change_direct_priced: pd.Series) -> None:
        self.price_change_direct = price_changes.full_direct_price_change(price_change_direct_priced)

        basket_shares = bsh_mod.get_budget_shares(input_data.distn_hh_survey, self.sample, 'mean').loc[0]
        pc_weights = basket_shares.reindex(price_changes.PC_GROUP_FUELS).fillna(0.0)

        self.price_change_indirect = price_changes.aggregate_price_change_to_categories(
            price_change_direct_priced, pc_weights, input_data.distn_io_gtap, input_data.distn_gtap_cpat_crosswalk
        )


    def __init_budget_shares_and_elasticities(self, input_data: InputData) -> None:
        self.budget_shares = bsh_mod.get_budget_shares(
            input_data.distn_hh_survey, self.sample, self.stat_type
        ).reindex(DECILES)
        self.elasticities = bsh_mod.get_elasticities(input_data.distn_hh_elast).reindex(DECILES)


    def __init_effects(self, d: DistributionInputsDict) -> None:
        behavior_adjustment = d['behavioral_response_adj_factor'] if d['adjust_for_behavioral_change'] else 1.0

        dwl_direct = dwl_indirect = None
        if d['adjust_for_deadweight_losses']:
            dwl_direct = bsh_mod.compute_deadweight_loss(
                self.budget_shares, self.elasticities, self.price_change_direct
            )
            dwl_indirect = bsh_mod.compute_deadweight_loss(
                self.budget_shares, self.elasticities, self.price_change_indirect
            )

        direct = effects.direct_effect(
            self.budget_shares, self.price_change_direct, dwl_direct, behavior_adjustment
        )
        if d['exempt_cooking_fuel']:
            direct = effects.apply_cooking_fuel_exemption(
                direct, d['exempt_cooking_fuel_code'], d['exempt_cooking_fuel_share'],
                d['exempt_cooking_fuel_bottom_deciles']
            )
        self.direct_effect = direct

        self.indirect_effect = effects.indirect_effect(
            self.budget_shares, self.price_change_indirect, dwl_indirect, behavior_adjustment
        )
        self.total_effect = effects.total_effect(self.direct_effect, self.indirect_effect)


    def __init_rebasing(self, input_data: InputData, national_totals: dict | None) -> None:
        survey = bsh_mod.select_hh_cells(
            input_data.distn_hh_survey, self.sample, 'mean',
            {'population': 'popw', 'per_capita_consumption': 'cons_pc_acrent'}
        ).reindex(DECILES)

        if national_totals is not None:
            national_population = national_totals['population']
            national_consumption = national_totals['consumption']
        else:
            national_population = survey['population'].sum()
            national_consumption = (survey['population'] * survey['per_capita_consumption']).sum()

        self.rebased = rebasing.rebase_to_national_accounts(
            survey['population'], survey['per_capita_consumption'],
            national_population, national_consumption
        )


    def __init_recycling(self, d: DistributionInputsDict, input_data: InputData) -> None:
        hh_survey = input_data.distn_hh_survey
        pit_raw_share = bsh_mod.select_hh_cells(
            hh_survey, self.sample, 'mean', {'pit_share': 'pit_share_income'}
        ).reindex(DECILES)['pit_share'].fillna(0.0)

        access = bsh_mod.select_hh_cells(
            hh_survey, self.sample, 'mean', {'access': d['public_investment_access_type']}
        ).reindex(DECILES)['access'].fillna(0.0)

        population = self.rebased['population']

        revenue_labor_tax = self.cp_revenue * d['share_labor_tax_reduction']
        revenue_transfer = self.cp_revenue * d['share_targeted_transfer']
        revenue_investment = self.cp_revenue * d['share_public_investment']
        revenue_spending = self.cp_revenue * d['share_current_spending']

        pit_shares = recycling.pit_reduction_shares(
            pit_raw_share, d['labor_tax_reduction_method'], d['labor_tax_exempt_bottom_deciles']
        )
        transfer_shares = recycling.targeted_transfer_shares(
            population, d['transfer_targeted_percentile'], d['transfer_coverage_rate'], d['transfer_leakage_rate']
        )
        investment_shares = recycling.public_investment_shares(access)
        spending_shares = recycling.current_spending_shares(population)

        self.pit_amount = pit_shares * revenue_labor_tax
        self.transfer_amount = transfer_shares * revenue_transfer
        self.investment_amount = investment_shares * revenue_investment
        self.spending_amount = spending_shares * revenue_spending

        self.amount_recycled = self.pit_amount + self.transfer_amount + self.investment_amount + self.spending_amount


    def __init_welfare(self, d: DistributionInputsDict) -> None:
        consumption = self.rebased['total_consumption']

        # total_effect is negative (a cost, matching the Excel model's own
        # sign convention -- see effects.direct_effect); amount_recycled is
        # a positive benefit, so it's added back, not subtracted.
        self.net_effect = self.total_effect['total'] + (self.amount_recycled / consumption * 100.0)

        self.post_policy = welfare.post_policy_consumption(
            consumption, self.total_effect['total'], self.amount_recycled
        )

        population_share = self.rebased['population'] / self.rebased['population'].sum()
        self.gini = {
            'baseline': welfare.gini_from_shares(population_share, self.post_policy['baseline_share']),
            'excl_recycling': welfare.gini_from_shares(population_share, self.post_policy['post_excl_recycling_share']),
            'incl_recycling': welfare.gini_from_shares(population_share, self.post_policy['post_incl_recycling_share']),
        }
        self.delta_gini_excl_recycling = self.gini['excl_recycling'] - self.gini['baseline']
        self.delta_gini_incl_recycling = self.gini['incl_recycling'] - self.gini['baseline']

        self.compensation = welfare.compensation_shares(
            self.total_effect['total'], self.post_policy['baseline_share']
        )


    def __init_sectoral(self, input_data: InputData) -> None:
        io_gtap = input_data.distn_io_gtap
        pc_weights = bsh_mod.get_budget_shares(input_data.distn_hh_survey, self.sample, 'mean').loc[0].reindex(
            price_changes.PC_GROUP_FUELS
        ).fillna(0.0)
        price_change_direct_priced = self.price_change_direct.reindex(
            [c.COA, c.ELE, c.NGA, c.OOP, c.GSO, c.DIE, c.KER, c.LPG]
        )
        price_increase_by_sector = price_changes.indirect_price_increase_by_gtap_sector(
            price_change_direct_priced, pc_weights, io_gtap
        )
        hhd_by_sector = io_gtap[io_gtap['dis'].str.endswith('.hhd')].set_index('gtap_sector')['value']
        self.sector_ranking = sectoral.rank_affected_sectors(price_increase_by_sector, hhd_by_sector)


    def results_table(self) -> pd.DataFrame:
        """
        Step 14: assembles a tidy results table, one row per indicator,
        matching the shape of the Excel model's own MTOutputs 'distn' block
        (CPATCode/Variable/Unit/Value) documented in
        distribution/docs/CPAT_Distribution_Module_Pseudocode.docx §6 and
        traced in §8. Deliberately not identical column-for-column to
        MTOutputs (no year-columns pivot, no Include?/Scenario/SubScenario
        bookkeeping columns) -- this is the model-facing shape;
        reshaping it into the MTOutputs export shape is a thin follow-on
        step once the module is wired into a UI/export layer, not done here.
        """
        rows = []

        def add(indicator, code, value, unit='%'):
            rows.append({'Indicator': indicator, 'Code': code, 'Value': value, 'Unit': unit})

        add('Total effect', 'tef.tot', self.total_effect['total'].mean())
        add('Direct effect', 'def.tot', self.total_effect['direct'].mean())
        add('Indirect effect', 'ief.tot', self.total_effect['indirect'].mean())
        for fuel in self.direct_effect.columns:
            add(f'Direct effect: {fuel}', f'def.{fuel}', self.direct_effect[fuel].mean())
        for cat in self.indirect_effect.columns:
            add(f'Indirect effect: {cat}', f'ief.{cat}', self.indirect_effect[cat].mean())
        for item in self.budget_shares.columns:
            add(f'Budget share: {item}', f'bsh.{item}', self.budget_shares.reindex(DECILES)[item].mean())
        add('Net effect (post-recycling)', 'nef.tot', self.net_effect.mean())
        add('Total recycled', 'rec.tot', (self.amount_recycled / self.rebased['total_consumption'] * 100).mean())
        add('Gini, excl. recycling (Δ)', 'gini.excl', self.delta_gini_excl_recycling, unit='Gini points')
        add('Gini, incl. recycling (Δ)', 'gini.incl', self.delta_gini_incl_recycling, unit='Gini points')
        add('Share of CP revenue to compensate all deciles', 'pctrev_comp',
            self.compensation['cumulative_share_of_revenue'].iloc[-1] * 100)

        return pd.DataFrame(rows)
