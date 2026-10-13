# Quick start and guide to the repository (CPAT-AI-Mitigation-MVP v2.2)

## What it is

CPAT-AI-Mitigation-MVP is the mitigation module of CPAT (the IMF/World Bank Climate Policy Assessment Tool) rebuilt as one short, readable Excel workbook with Egypt data. It follows the legacy CPAT chain: policies (carbon tax, new ETS with a cap, feebates, fuel price reform, existing carbon pricing) change retail energy prices; prices change fuel use in 16 subsectors and 8 fuels; a power model turns electricity demand into dispatch, retirement, investment and electricity prices; the results are fuel use, revenues and CO2 to 2040, in real US dollars of 2026. Every formula is short (at most 100 characters) and one formula serves a whole block, so the workbook can be extended by copying and dragging. The workbook and every document were generated with AI and checked by an independent Python recomputation.

## Quick start (about 15 minutes)

1. **Open** `CPAT-AI-Mitigation-MVP-v2.2.xlsm` (Excel 2019 or later; in the folder `cpat_excel_new/working_version/` or `cpat_excel_new/mitigation_copypaste/`). If the file was downloaded, unblock it (file properties) and enable macros.
2. **Test once:** press Alt+F8 and run `CheckBatchRun`. The sheet `MacroCheck` must show Result PASS. It checks that the macros find every sheet, row and column and that Excel reproduces the stored results.
3. **Look at the results:** the tab `Mitigation` opens rolled up. Each scenario is a column group (scenario 1 = baseline, scenario 2 = the live policy scenario). Use the outline buttons (1 / 2 / 3, top left) to open sections: 1 Policies, 2 Prices, 3 Power, 5-8 sector blocks, 11 Energy, 12 Revenues, 13 Emissions. `MTOutputs` lists 58 key results by output code; `Charts` plots them.
4. **Change a policy:** on `MTInputs`, column K holds the live scenario (rows A:H describe each input). Change, for example, `CPLevelTarget` (carbon price in the target year); Mitigation group 2 recalculates.
5. **Run many scenarios:** copy the last scenario definition on `MTInputs` (column N) one column to the right, edit it, set row 4 Run? = Yes and run the macro `RunAllScenarios`. The results are stored as values in `StoredResults`; `ScenarioCompare` compares up to 8 stored scenarios with the baseline for one year.
6. **Read more in the workbook:** `ReadMe` (how the sheet works), `LegacyDiff` (every difference with legacy CPAT), `Variables` (every variable code with its formula and source), `Lambdas` (optional named LAMBDA formulas).
7. **Read more outside it:** this documentation; the register of bugs, issues and upgrades `CPAT-AI-Mitigation-MVP_BugsIssuesUpgrades_v2.2.docx`; the deck `CPAT-AI-Mitigation-MVP_v2.2.pptx`.

**Rebuild from source (Python, LibreOffice):** in `cpat_excel_new/mitigation_copypaste/` run `python build_v2_2.py` (workbook), `python check_v2_2.py` (all checks; writes `check_report_v2.2.md`), `python make_docs_v2_2.py` and `python make_issues_doc_v2_2.py` (Word documents), `python presentation/make_screenshots_v2_2.py` and `node presentation/build_deck_v2_2.js` (deck). The data step scripts (`extract_data_v2_2.py`, `extract_power_v2_2.py`, `extract_ecp_v2_2.py`, `update_prices_egypt_v2_2.py`) are needed only when source data change.

## Guide to the repository

The repository `CPAT-AI-Local` holds the AI-assisted rebuilds of CPAT. Top level:

| Folder or file | What it holds |
|---|---|
| `README.md`, `AGENTS.md`, `CLAUDE.md`, `SyncRepo.bat` | Overview of the repository; instructions for AI assistants working in it; local sync helper |
| `NORMS.md` | Working norms (versioning, formulas at most 100 characters, bookkeeping) |
| `CAVEATS.md` | One entry per task: what was asked, built, checked and what to watch (append only) |
| `TODO.md` | Open decisions and next steps |
| `templates/` | `MTInputs_template.xlsx`, the legacy scenario-input layout read by the MVP |
| `cpat_excel_original/` | The legacy CPAT workbook (`CPAT 1.0pre_456_NoPropData.xlsb`), the source of the data and of the legacy tests |
| `cpat_excel_new/` | Excel prototypes made with AI (below) |
| `cpat_coded/` | CPAT ported to Python (`cpat_model/`, data, tests, developer docs) |
| `_research/` | Source literature (IPCC, CBAM guidance, sector roadmaps) |

The Egypt case (`egypt/`, `egypt-final/`) was removed from this repository on 2026-10-09 and is in the git history.

In `cpat_excel_new/`:

| Folder | What it holds |
|---|---|
| `mitigation_copypaste/` | **The mitigation MVP**: builder, checks, data, method notes, documents (below) |
| `working_version/` | Hand-over copies of the current MVP files (workbook, Word documents, deck, method notes, VBA module); earlier copies in `Old/` |
| `distribution/` | Distribution module prototype |
| `cbam_tool/` | CBAM policy tool guide |
| `standalone_initial_prototypes/`, `tecp_and_validation/` | First-generation prototypes; carbon-price and elasticity reference data |
| `old/` | Archive, including the Egypt industry kernel (`standalone_working_version/`) |

In `cpat_excel_new/mitigation_copypaste/` (every current file has version 2.2; earlier versions are in `Old/`):

| File | Role |
|---|---|
| `CPAT-AI-Mitigation-MVP-v2.2.xlsm` | The workbook (macro-enabled) |
| `build_v2_2.py` | Builds the workbook from `data/` and `templates/MTInputs_template.xlsx` (Python, openpyxl; LibreOffice for the baseline CO2 data and the stored results) |
| `check_v2_2.py`, `check_report_v2.2.md` | All checks (LibreOffice recalculation, independent Python recomputation, scenario tests, regression, consistency) and their report |
| `data/` | Extracted source data (CSV) and `source/` (user-supplied Egypt price block) |
| `extract_data_v2_2.py`, `extract_power_v2_2.py`, `extract_ecp_v2_2.py`, `update_prices_egypt_v2_2.py` | Data step: legacy workbook and other sources to `data/*.csv` |
| `power_model_ref_v2_2.py` | Python reference of the power model (reproduces legacy, switches for the ten legacy bugs) |
| `ets_goalseek_v2_2.py` | ETS permit-price goal seek (same algorithm as the macro) |
| `legacy_compare_v2_2.py`, `legacy_comparison_v2_2.md` | Comparison of the baseline with the cached legacy run |
| `CPATScenarios_v2_2.bas`, `vba_project_v2_2.py`, `lo_run_macro_v2_2.py` | VBA module (embedded in the .xlsm), the writer of the VBA project, a LibreOffice macro runner for the checks |
| `QuickStart_v2_2.md`, `README.md` | This quick start and the model overview |
| `PriceProjection_Method_v2_2.md`, `ETS_Method_v2_2.md`, `Scenarios_Method_v2_2.md`, `PowerPrices_Method_v2_2.md`, `PowerModel_Method_v2_2.md` | Method notes (current state) |
| `ETS_Legacy_Algorithm_v2_2.md`, `ETS_Cap_Design_Options_v2_2.md` | Background notes on the ETS design |
| `issues_register_v2_2.py`, `make_issues_doc_v2_2.py`, `render_docx_v2_2.js`, `CPAT-AI-Mitigation-MVP_BugsIssuesUpgrades_v2.2.docx` | Register of bugs, known issues and upgrades and its Word document |
| `make_docs_v2_2.py`, `CPAT-AI-Mitigation-MVP_Documentation_v2.2.docx` | This Word documentation |
| `presentation/` | Deck `CPAT-AI-Mitigation-MVP_v2.2.pptx`, its builder `build_deck_v2_2.js`, screenshots `make_screenshots_v2_2.py`, `img/`; earlier decks in `Old/` |
| `Old/` | Every earlier version (workbooks v0.1-v1.09, builders, checks, reports, notes) |

## Where to find what in the workbook

| Sheet | Content |
|---|---|
| `Mitigation` | The model: one column group per scenario; sections 1 Policies, 2 Prices, 3 Power (sub-tables A-M), 5-8 sectors, 11 Energy, 12 Revenues, 13 Emissions |
| `MTInputs` | Scenario inputs (legacy layout): J baseline, K live scenario, L onwards scenario definitions |
| `MTOutputs`, `Charts`, `StoredResults`, `ScenarioCompare`, `MacroCheck` | Results by output code, charts, stored scenario results, comparison, first-run test |
| `ReadMe`, `LegacyDiff`, `Settings`, `Lambdas`, `Variables` | Guide, differences with legacy, settings and version log, optional LAMBDA library, variable dictionary |
| `Inputs`, `Inputs_prices`, `Inputs_power`, `Mapping` | Data step: every lookup used by Mitigation |
| `Countries`, `Elasticities`, `Prices_dom`, `Prices_int`, `PriceAssump`, `EnergyCons`, `WEO`, `EF_GHG`, `PowerTech`, `PowerPaths`, `PowerParams`, `PowerDemand`, `ECP` | Data sheets |
