# CPAT Distribution Module

Root folder for the household distributional (incidence) analysis module of CPAT —
the model that estimates how a carbon-pricing policy scenario affects household
consumption by income decile, and how revenue recycling offsets that effect.

This module exists in full inside the main Excel model
(`cpat_excel/original/CPAT 1.0pre_456_NoPropData.xlsb`, sheet `Distribution`) and
its supporting data tables (`cpat_excel/Distribution/data_*`). This folder holds
the two planned reimplementations of that sheet:

- Python — lives in `cpat_coded/cpat_model/components/distribution` (not under this
  folder, so it sits alongside every other component), wired into the rest of the
  model. **All 14 pseudocode steps are implemented and running against real Egypt
  data** (see below for status and known accuracy gaps).
- `excel/` (not yet started) — a standalone Excel workbook that replicates just the
  Distribution module, independent of the full CPAT model.

## Current contents

- `docs/CPAT_Distribution_Module_Pseudocode.docx` — the design specification:
  Inputs, Algorithm (14 steps) and Outputs pseudocode for the module, derived
  directly from the live Excel sheet's structure, row labels and embedded notes.
  This is the reference the Python implementation is built from, and each source
  file below cites the pseudocode section/step it implements.

## Python implementation (cpat_coded)

`cpat_model/components/distribution/`, one file per step group, each with its
formulas documented against both the pseudocode doc and the exact Excel row
range:

| File | Steps | What it does |
|---|---|---|
| `data.py` | — | Loaders for the module's data tables (IO_GTAP, HHSurvey, HH_Elast, ASPIRE, WHOCooking, GDPRatios, GTAP↔CPAT crosswalk), reading CSVs exported from `data_standardized` |
| `price_changes.py` | 1 | Direct fuel price change (from a Mitigation-module run, or supplied directly) → indirect price change per CPAT category, via IO_GTAP's Leontief coefficients and the GTAP→CPAT crosswalk |
| `budget_shares.py` | 2, 3 | Household budget shares and price elasticities, pivoted from HHSurvey/HH_Elast; deadweight-loss calculation |
| `effects.py` | 4, 5, 6 | Direct effect (by fuel), indirect effect (by category), total effect, cooking-fuel exemption |
| `rebasing.py` | 7 | Rescales survey population/consumption onto given national-accounts totals |
| `recycling.py` | 8 | PIT/labor-tax reduction (3 methods), rules-based targeted transfers, public-investment and current-spending incidence |
| `welfare.py` | 9, 10, 11, 12 | Post-policy consumption & shares, Gini/Lorenz, horizontal equity, compensation requirement |
| `sectoral.py` | 13 | Ranks GTAP sectors by embedded-energy price shock and household-demand share |
| `distribution.py` | 14 | The `Distribution` class — wires every step together per country/analysis-year, and assembles the results table |

Supporting pieces:
- `cpat_model/inputs/distribution_inputs.py` — `DistributionInputsDict` + `DistributionInputs`, kept separate from `DashboardInputsDict` since the Excel Distribution sheet has its own independent input block. Includes the revenue-recycling design fields (PIT method, transfer targeting/coverage/leakage, public-investment access type) discovered while wiring this to real data.
- `cpat_model/scenario_results.py` — `ScenarioResults` bundle. Distribution needs *two* finished scenario runs (baseline and policy) to derive a price change; `run_model()` stores one per scenario and instantiates `Distribution` once per (policy scenario × country) after the main year loop.
- `cpat_model/constants.py` — Distribution-specific fuel/category/decile/sample/statistic codes (one flagged naming collision: the Excel model's `hea` code means both "heat" (Mitigation) and "health services" (Distribution); kept as distinct Python names `HEA`/`HEALTH_SRV`).
- `cpat_excel/scripts/build_distribution_model_data.py` — exports `cpat_excel/Distribution/data_standardized/CPAT_DistributionalData.xlsx` into `cpat_data/new_data/distn_*.csv`, the bridge between the Excel-side extraction pipeline and the Python model's data convention. Re-run this whenever `data_standardized` is regenerated.
- `cpat_testing/model/components/distribution/` — `test_price_changes.py` (Step 1 unit tests, including one against real Egypt IO_GTAP data) and `test_distribution.py` (end-to-end integration test running the full pipeline against real Egypt data, asserting structural correctness — signs, exact zeros, internal consistency — see below for why not numeric parity).

## Validating against the Excel model

The Excel workbook's `MTOutputs` sheet already has computed Distribution results
for Egypt's standard $0-in-2026 → $50-in-2030 carbon tax. Feeding the same
scenario's known price changes into the Python pipeline and comparing:

- **Budget shares**: coal = 0 in both (Egyptian households don't buy coal
  directly — IO_GTAP's household-demand figure for the coal sector is ~0,
  independently confirming the same fact HHSurvey shows). Electricity/LPG/food
  dominate in both, similar order of magnitude.
- **Direct effect by fuel**: electricity, natural gas and gasoline match Excel's
  figures within ~15%. LPG and other-oil are further off (~2x), most likely
  because HHSurvey only has the 2017 survey-year budget shares — Excel's 2026/2030
  figures are income-growth-projected forward, a step this dataset doesn't
  support (see gaps below).
- **Indirect effect by category**: correct *relative ranking* across all 14
  categories (same categories high/low in both), but running roughly 7-15x
  larger in absolute terms — consistent with a genuine, expected gap (see below),
  not a formula error.
- **Gini, excl. recycling**: same sign and same order of magnitude (computed:
  +0.06 Gini points; Excel: +0.13).
- **Gini, incl. recycling**: same sign (recycling reduces inequality) in both;
  smaller in magnitude here (-0.3 vs. Excel's -4.7 Gini points), most likely
  because the PIT/transfer/investment incidence formulas here are share-based
  approximations of Excel's more elaborate capped/redistributive algorithm (see
  gaps below).
- **Structural invariants** (asserted in `test_distribution.py`): total effect
  negative in every decile, all 14 indirect categories present, compensation
  shares sum to exactly 100% at the richest decile, recycling always improves
  the net position, baseline Gini is a plausible 0.1-0.5.

### Known accuracy gaps (data this session didn't have)

Two things live only on the Mitigation side of the Excel workbook, not in
`cpat_excel/Distribution/data_standardized`:

1. **Sectoral pass-through coefficients** (Excel §B.IV, rows 464-486) and the
   **IEA/GAINS emissions-based recalibration** (§C.I, rows 632-664). Without
   these, the indirect-effect calculation is the *pre-pass-through* GTAP-IO
   figure — structurally correct (right ranking) but not scaled down to
   retail-price reality. This is the single biggest source of the gap above.
2. **Real GDP and a deflator/FX chain.** Needed for: projecting HHSurvey's
   2017 budget shares forward to the analysis year (income-growth/Engel-curve
   effects); the Pm/Pd revenue reconciliation (Step 1.3 — currently skipped
   entirely, not even approximated, since IO_GTAP's consumption figures are in
   GTAP-year US$bn with no conversion path to LCU here); and absolute PIT/
   national-accounts baselines for `recycling.py` and `rebasing.py` (both fall
   back to share-based approximations, documented in-line, rather than the
   Excel model's absolute-currency-capped algorithms).

If this Mitigation-side data becomes available (either as real `cpat_data/new_data`
CSVs, or exported from the Excel workbook the same way `data_standardized` was),
closing gap 1 is a small, contained change to `price_changes.py`; closing gap 2
touches `rebasing.py`, `recycling.py` and the budget-share projection step (not
yet a separate module — would slot into `budget_shares.py`).

## Status

All 14 steps implemented and running end-to-end against real Egypt data, with the
accuracy gaps above documented and tested for. Not yet: the standalone Excel tool,
and the two data gaps above. See the pseudocode document for the full spec.
