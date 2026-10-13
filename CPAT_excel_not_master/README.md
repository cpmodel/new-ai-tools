# CPAT-AI-Local

Local workspace for the **Climate Policy Assessment Tool (CPAT)** — a World
Bank model for assessing the fiscal, economic, energy and distributional
impacts of carbon pricing and related climate policy reforms. This repo
contains both the original Excel-based model and an in-progress Python
reimplementation, plus supporting data-conversion tooling.

## Repository layout

The repo serves two parallel purposes: **Coded CPAT** (Python) and **Excel-AI** (Excel-based prototypes built/maintained with AI assistance). All Excel work follows [`NORMS.md`](NORMS.md) (column, colour, input and versioning norms); **every task**, Excel or not, ends by logging task/inputs/outputs/caveats in [`CAVEATS.md`](CAVEATS.md) (see [When a task is done](#when-a-task-is-done)).

```
cpat_coded/            Coded CPAT: Python reimplementation of the model
cpat_excel_original/   Legacy CPAT workbook (CPAT 1.0pre_456_NoPropData.xlsb) - the reference (older docs: original_cpat_excel/)
cpat_excel_new/        Excel-AI: new Excel prototypes replicating legacy CPAT modules
templates/             Master templates (MTInputs_template.xlsx)
_research/             Source literature: IPCC AR6 WGIII Ch11/Ch12, CBAM regulation and guidance, sector roadmaps (PDF + extracted text)
NORMS.md               Excel column, colour, input and versioning norms; task-completion process (section 6)
CAVEATS.md             Append-only log of completed tasks: task, inputs, outputs, caveats
TODO.md                Open work only (checks, modelling) and the per-task conventions
SyncRepo.bat           Local sync helper
```

> **Removed folders.** The Egypt case (`egypt/`, `egypt-final/`) was removed from this repository on 2026-10-09 (commit f070c20, "Add CBAM Policy Tool Guide"); it is available in the git history. The industry-kernel working version moved to `cpat_excel_new/old/standalone_working_version/`. References to these paths in `NORMS.md`, `AGENTS.md`, `TODO.md` and `CAVEATS.md` point to that history.

### Workflow
1. Replicate legacy CPAT modules as Excel prototypes in `cpat_excel_new/` (same columns/colours, validated against `cpat_excel_original/`).
2. Build country-specific functionality (Egypt so far; the Egypt case folder is in the git history).
3. Port validated logic to `cpat_coded/`; codes stay identical across Excel and Python.
4. Log every finished task (below).

### When a task is done
Follow [`NORMS.md` section 6](NORMS.md#6-task-completion-caveats-log-and-bookkeeping). In short:
1. **Append an entry to [`CAVEATS.md`](CAVEATS.md)** (repo root, newest at the bottom, never edit earlier entries) recording the **task**, its **inputs**, its **outputs** and any **caveats** - placeholders, assumptions, known gaps, deferred decisions, explained regression differences. The entry template is at the top of that file.
2. For workbooks: add the version-log row in `Settings`/`ReadMe` and run the `Check`-sheet regression (NORMS section 5).
3. Tick `TODO.md` (the Egypt task inventory in `egypt/instructions/` is in the git history, see above).
4. Do not commit unless asked.

### Key documents

| Document | Role |
|---|---|
| [`NORMS.md`](NORMS.md) | Norms for all Excel work (sections 1-5) and the task-completion process for all work (section 6) |
| [`CAVEATS.md`](CAVEATS.md) | Log of completed tasks - task, inputs, outputs, caveats. Read before building on earlier work |
| [`TODO.md`](TODO.md) | Open work and the conventions every kernel increment must follow |
| [`cpat_excel_new/distribution/README.md`](cpat_excel_new/distribution/README.md), [`LESSONS_LEARNED.md`](cpat_excel_new/distribution/LESSONS_LEARNED.md), [`REGENERATING_THE_DISTRIBUTIONAL_WORKBOOK.md`](cpat_excel_new/distribution/REGENERATING_THE_DISTRIBUTIONAL_WORKBOOK.md) | Distribution module spec/status, LAMBDA-workbook lessons, regeneration steps |
| [`cpat_excel_new/mitigation_copypaste/README.md`](cpat_excel_new/mitigation_copypaste/README.md) | CPAT-AI-Mitigation-MVP (copy-pasteable mitigation module): design, files, rebuild, how to add and batch-run scenarios; method notes `PriceProjection_Method_v0.4.md`, `ETS_Method_v2_3.md`, `Scenarios_Method_v0_1.md`, `PowerPrices_Method_v0_1.md`, `PowerModel_Method_v0_1.md`; Word documentation `CPAT-AI-Mitigation-MVP_Documentation_v1.06.docx` and the legacy power bugs report `CPAT-AI-Mitigation-MVP_BugsIssuesUpgrades_v2.3.docx`; decks in `presentation/` |
| `cpat_excel_original/CPAT 1.0pre_456_NoPropData.xlsb` | Legacy CPAT - ground truth. Read-only |

### `cpat_excel_original/`
The source workbook (`CPAT 1.0pre_456_NoPropData.xlsb`). Read-only reference; do not edit. Older documents refer to this folder as `cpat_excel_original/`.

### `cpat_excel_new/` - Excel-AI prototypes
- `distribution/` - Distribution module: `docs/` (pseudocode), `data_*` staged data exports, `scripts/` (conversion pipeline), `standalone/` (standalone Egypt workbook + builder; `Old/` holds earlier versions), plus README, lessons learned and regeneration notes
- `mitigation_copypaste/` - **CPAT-AI-Mitigation-MVP v2.3** (last year 2041 calculated with named LAMBDAs; moving-average fuel prices for power as legacy), the copy-pasteable mitigation module (Egypt data): workbook `CPAT-AI-Mitigation-MVP-v2.3.xlsm`, builder `build_v2_3.py`, checks `check_v2_3.py`, data step, method notes, Word documentation (opens with a quick start and a guide to this repository, `QuickStart_v2_3.md`), register of bugs, issues and upgrades, deck (`presentation/`); every current file has version 2.3, earlier versions in `Old/`. Contents: policies (carbon tax, cap-based new ETS, feebates, fuel price reform, existing carbon pricing), domestic price projection in real USD, fuel use for 16 subsectors x 8 fuels, power engineer model (demand, dispatch, retirement, investment, electricity prices; ten legacy bugs fixed), fuel use, revenues and CO2 with power in the totals, many scenarios with a VBA batch run; see its README
- `working_version/` - current hand-over copies of the mitigation MVP v2.3 (workbook, Word documentation, register of bugs, issues and upgrades, deck, method notes, VBA module); earlier copies in `Old/`
- `standalone_initial_prototypes/` - first-generation standalone workbooks (distribution v0.4, prices module v1.12, industry kernel v0.1/v0.1b LAMBDA, energy-kernel mock-up, South Africa inputs, mitigation-equations PoC)
- `tecp_and_validation/` - total effective carbon price data and price-elasticity references
- `cbam_tool/` - CBAM policy tool guide (`CBAM_Policy_Tool_Guide_v3.0.docx`)
- `old/` - archived material: data from an earlier conversion effort (`Coded_Conversion_Data/`) and the industry-kernel working version (`standalone_working_version/`, `CPAT_Industry_Kernel_Egypt_v1.6.xlsx` and its builders)
- `readme.xlsx`

### `egypt/` - Egypt case (removed)
Removed from this repository on 2026-10-09 (commit f070c20); see the git history for the final Egypt set v1.6 (kernel, methodology, Table 2, caveats) and its supporting work.

### `cpat_coded/` - Python model

A component-based reimplementation of the CPAT model, driven by
`run_model.py` and configured via `config.py` (selected countries, scenarios,
data path).

- `cpat_model/` - the model: `components/` (`gdp`, `policies`, `carbon_pricing`, `efs`, `elasticities`, `prices`, `energy_consumption`, `emissions`, `power`, `distribution`), `inputs/`, `mappings/`, `constants.py`, `scenario_results.py`
- `cpat_data/` - Egypt (EGY), MENA and global input data
- `cpat_testing/` - pytest suite mirroring `cpat_model/`
- `cpat_documentation/developer/` - developer docs
- `environment.yml` / `requirements.txt` - dependencies (Python 3.11, pandas, numpy, openpyxl, pyxlsb, pytest)

The Distribution module is the most complete component; see `cpat_excel_new/distribution/README.md` for its spec, status and known gaps versus the Excel model.

## Getting started

```powershell
# Create the Conda environment
conda env create -f cpat_coded\environment.yml
conda activate cpat_sisepuede

# Or install with pip
pip install -r cpat_coded\requirements.txt

# Run the model
cd cpat_coded
python run_model.py
```

Run the test suite with:

```powershell
cd cpat_coded
pytest
```


