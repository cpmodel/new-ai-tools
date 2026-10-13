"""
End-to-end integration test for the Distribution module, against the real
Egypt data in cpat_data/new_data/distn_*.csv (exported from
cpat_excel/Distribution/data_standardized by
cpat_excel/scripts/build_distribution_model_data.py -- run that script
first if these files are missing).

This does not use InputData (which also loads Mitigation-side CSVs this
repo doesn't ship, see cpat_data/README.txt) -- just the Distribution
loaders directly, bundled into a lightweight stand-in with the same
attribute names InputData would have.

price_change_direct and cp_revenue below are not derived from a Mitigation
run (none is available); they're read off the CPAT Excel model's own
already-computed MTOutputs sheet for Egypt's standard $0-in-2026,
$50-in-2030 carbon tax scenario (see the session's prior turn / distn_*
artifact), the closest thing to ground truth this repo has. The assertions
are deliberately structural (signs, exact zeros, internal consistency),
not numeric matches to those Excel figures: several formulas here are
documented approximations (rebasing.py, recycling.py, price_changes.py all
say where and why), so an exact match isn't expected -- see
distribution/README.md for how close a validation run actually came.
"""
from types import SimpleNamespace

import pytest
import pandas as pd

from cpat_model.components.distribution.data import (
    load_io_gtap, load_hh_survey, load_hh_elast, load_gtap_cpat_sector_crosswalk
)
from cpat_model.components.distribution.distribution import Distribution
from cpat_model.inputs.distribution_inputs import DistributionInputs
import cpat_model.constants as c


COUNTRY = 'EGY'

# Egypt, $0-in-2026 -> $50-in-2030 carbon tax: direct fuel price increases,
# as read from the Excel model's own MTOutputs sheet.
PRICE_CHANGE_DIRECT = pd.Series({
    c.COA: 0.3781202596689215, c.ELE: 0.25453525713581588, c.NGA: 0.5943929712377744,
    c.OOP: 1.039578829680615, c.GSO: 0.19170744285566332, c.DIE: 0.28067927548660254,
    c.KER: 0.30101780782047417, c.LPG: 0.5867220010621488,
})

# Same source: 'Total Recycled - Overall' (1.00127%) x 'Total consumption'
# (4,334,711,952,138 LCU), both for 2030 -- recycling shares sum to 100% of
# CP revenue in this scenario (30% PIT + 50% investment + 20% transfers),
# so this backs out an implied cp_revenue without needing the Mitigation
# module's own revenue figure.
CP_REVENUE_2030 = 0.010012709841958751 * 4_334_711_952_138.174
NATIONAL_TOTALS_2030 = {'population': 10_876_684.950810278, 'consumption': 4_334_711_952_138.174}


@pytest.fixture(scope='module')
def input_data():
    return SimpleNamespace(
        distn_io_gtap=load_io_gtap([COUNTRY]),
        distn_hh_survey=load_hh_survey([COUNTRY]),
        distn_hh_elast=load_hh_elast([COUNTRY]),
        distn_gtap_cpat_crosswalk=load_gtap_cpat_sector_crosswalk(),
    )


@pytest.fixture(scope='module')
def distribution(input_data):
    d = DistributionInputs({
        'analysis_year': 2030,
        'share_labor_tax_reduction': 0.30,
        'share_targeted_transfer': 0.20,
        'share_public_investment': 0.50,
        'share_current_spending': 0.0,
        'labor_tax_reduction_method': 'Personal Allowance',
        'transfer_targeted_percentile': 0.40,
        'transfer_coverage_rate': 1.0,
        'transfer_leakage_rate': 0.0,
    }).d
    return Distribution(
        d, COUNTRY, input_data, PRICE_CHANGE_DIRECT, CP_REVENUE_2030, NATIONAL_TOTALS_2030
    )


def test_coal_has_zero_budget_share_and_zero_direct_effect(distribution):
    # Egypt households don't buy coal directly (IO_GTAP's household-demand
    # figure for the coal sector is ~0) -- matches the Excel ground truth
    # exactly (coal direct effect = 0.0 in every year the sheet reports).
    assert (distribution.budget_shares[c.COA] == 0).all()
    assert (distribution.direct_effect[c.COA] == 0).all()


def test_total_effect_is_negative_every_decile(distribution):
    assert (distribution.total_effect['total'] < 0).all()
    assert (distribution.total_effect['direct'] < 0).all()
    assert (distribution.total_effect['indirect'] < 0).all()


def test_indirect_categories_cover_all_fourteen(distribution):
    assert set(distribution.price_change_indirect.index) == set(c.DISTN_INDIRECT_CATEGORIES)
    assert (distribution.price_change_indirect >= 0).all()


def test_recycling_reduces_the_burden(distribution):
    # net_effect = total_effect - (amount recycled / consumption); every
    # decile should be better off net of recycling than gross of it, since
    # some revenue is always recycled back and nothing here removes value.
    assert (distribution.net_effect > distribution.total_effect['total']).all()
    assert (distribution.amount_recycled > 0).all()


def test_compensation_shares_sum_to_full_revenue_at_the_top_decile(distribution):
    # cumulative_share_of_revenue is a running sum across *all* deciles with
    # a loss; with every decile losing here, it should reach 1.0 (100%) by
    # the richest decile.
    assert distribution.compensation['cumulative_share_of_revenue'].iloc[-1] == pytest.approx(1.0)


def test_gini_baseline_is_a_plausible_consumption_gini(distribution):
    # Egypt's actual consumption Gini is in the 0.3 range; this is a sanity
    # bound, not a precision check.
    assert 0.1 < distribution.gini['baseline'] < 0.5


def test_progressive_recycling_narrows_inequality_more_than_the_tax_alone_widens_it(distribution):
    # This scenario recycles 100% of revenue via PIT relief (Personal
    # Allowance, capped/redistributed toward the poor), infrastructure
    # investment and targeted transfers -- all progressive-leaning channels
    # -- so inequality including recycling should end up lower than the
    # inequality from the tax alone (excl. recycling), which itself is
    # expected to rise only slightly.
    assert distribution.gini['incl_recycling'] < distribution.gini['excl_recycling']


def test_results_table_has_no_missing_values(distribution):
    table = distribution.results_table()
    assert not table.empty
    assert table['Value'].notna().all()
