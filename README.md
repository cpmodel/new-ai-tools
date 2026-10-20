# Fiscal Climate AI Tools

A collection of working papers, data workbooks, and process files for fiscal and
climate-policy analysis, organized by project.

See **[SETUP.md](SETUP.md)** for environment requirements, build instructions, and
run-process specifications consolidated from the subfolder documentation.

## Repository structure

```
fiscal-climate-ai-tools/
├── SETUP.md                     Consolidated requirements, build steps, run process
├── CPAT_excel/                  CPAT Excel workspace (legacy model + Excel-AI prototypes);
│                                 has its own README/AGENTS/NORMS/CAVEATS/TODO
├── CPAT_coded/                  CPAT Python reimplementation
├── CBAMPolicyTool/              CBAM Policy Tool (global workbook, guide, Egypt case)
├── ExciseDiagnostic/            Excise tax diagnostic working paper and data
├── MacFiscEnergyScenarioProto/   Macro-Fiscal-Energy-Scenario Tool (prototype): one Dashboard for
│   │                             macro-fiscal and energy scenarios; its source models sit inside it:
│   ├── InvestmentNeeds/          Mitigation (net zero) investment needs model, by country,
│   │                             sector and cost-of-capital lens
│   ├── MacroFiscalModels/        GIMF-lite, DIGNAD/DIGNAR-lite, LIC-DSF-lite and the
│   │                             integrated run (dashboard panel, NMIM, PriceFeasibility)
│   └── PriceFeasibility/         Power-sector carbon tax vs feebate metrics under
│                                 political-feasibility limits (standalone Python)
├── TaxMacroFiscalDashboard/     Global Fiscal Indicators dashboard (the master) and other
│   └── build/                   dashboards; raw data, config and scripts for the indicator
│                                 panel (Python + R build pipeline), which the tool reads
├── templates/                   Master templates
└── _research/                   Source literature (IPCC, CBAM regulation, roadmaps)
```

## CPAT (Excel)

CPAT-Excel work is documented separately inside [`CPAT_excel/`](CPAT_excel/). For
any CPAT-Excel task, start there rather than here:

- [`CPAT_excel/README.md`](CPAT_excel/README.md) — CPAT workspace map, workflow,
  key documents
- [`CPAT_excel/AGENTS.md`](CPAT_excel/AGENTS.md) — instructions for AI agents on
  CPAT tasks
- [`CPAT_excel/NORMS.md`](CPAT_excel/NORMS.md) — Excel column, colour, input and
  versioning norms; task-completion process
- [`CPAT_excel/CAVEATS.md`](CPAT_excel/CAVEATS.md) — append-only log of completed
  tasks and their caveats
- [`CPAT_excel/TODO.md`](CPAT_excel/TODO.md) — open CPAT work

## ExciseDiagnostic

File set for the excise diagnostic working paper (current as of September 2026).

- **`Excise_Diagnostic_WorkingPaper_v0_9.docx`** — Current working paper. Section 5
  covers whole vs. five-part runs.
- **`Excise_Diagnostic_CrossCountry.xlsx`** — Cross-country workbook (Sierra Leone,
  Uganda, Liberia), tagged to the five-part map, with `Part_Map` and `Part_Summary`
  sheets.
- **`process_files/`** — Process specifications used to run the diagnostic:
  - `excise_diagnostic_process.md` — Master process specification (v1.1)
  - `excise_part1_fuel.md` — Part 1: fuel (E1–E6)
  - `excise_part2_vehicles.md` — Part 2: vehicles (E8–E9)
  - `excise_part3_health.md` — Part 3: health (T1–T3, A1–A5, S1–S4, I3)
  - `excise_part4_other.md` — Part 4: other (E7, G1, D1–D2, F1, I1, I2, I4, L1)
  - `excise_part5_overarching.md` — Part 5: overarching framework and aggregation
    (run last)
- **`country_results/`** — Country-level outputs (`Liberia/`, `SierraLeone/`)
- **`main_working_paper/`** — Supporting working-paper material, including the
  cross-country workbook, externalities/pathways summary, and an `Additional/` and
  `integrated_version/` subfolder
- **`archive/`** — Prior versions of the working paper (v0.7, v0.8)
- **`ExciseFiscalReportsProposal_v1.0.pptx`** — Proposal deck

Run modes:
- **Whole run**: master process file only (stages A1–A6).
- **Sectoral run**: master file plus one part file per run; Parts 1–4 in any
  order, Part 5 last.

## PriceFeasibility

Standalone metrics comparing a carbon tax/ETS, a feebate and hybrids in the power
sector: abatement subject to limits on visible tariff rises, bottom-40% burden and
losses to coal and oil generators, with per-country calibration from the dashboard
panel and Ember. See [`MacFiscEnergyScenarioProto/PriceFeasibility/README.md`](MacFiscEnergyScenarioProto/PriceFeasibility/README.md).
Also as a light three-tab Excel screening workbook,
`MacFiscEnergyScenarioProto/PriceFeasibility/PriceFeasibility_Screen.xlsx`, which reads the dashboard's long data format.

## MacFiscEnergyScenarioProto

The **Macro-Fiscal-Energy-Scenario Tool** (prototype, formerly the Integrated MF-C tool),
`MacFiscEnergyScenarioProto/Macro_Fiscal_Energy_Scenario_Tool_v0.7.xlsx`. One core tab, the **Dashboard**: pick a country
and a scenario (revenue increase and public investment increase in pp of GDP, the climate investment
programme on/off, power-sector carbon pricing - a carbon tax, a feebate or the most effective
politically feasible mix) and read headline results and the key time series: public debt, revenue
and expenditure, growth, green investment, the carbon price, debt sustainability. Everything else sits
behind a **Detail ->>** divider: Fiscal outlook, Macro & debt (the integrated run, the live DIGNAD-R
model and LIC-DSF-lite), Green investment (NMIM v0.2 live; its assumptions drive the fiscal outlook),
Carbon pricing, Policy settings and the model sheets. At the far right, behind **Not master ->>**, a
reference copy of the Global Fiscal Indicators country profile; the master is
`TaxMacroFiscalDashboard/`. Built by `MacFiscEnergyScenarioProto/build/build_mfc_tool.py`. The source models
(InvestmentNeeds, MacroFiscalModels, PriceFeasibility) are subfolders of it; it reads the dashboard
panel from `TaxMacroFiscalDashboard/build/`. See [`MacFiscEnergyScenarioProto/README.md`](MacFiscEnergyScenarioProto/README.md).

## InvestmentNeeds

Mitigation (net zero) investment needs model, built on the CDPM-based global
method (October 2026; CDPM was formerly SiSePuede). See [`MacFiscEnergyScenarioProto/InvestmentNeeds/README.md`](MacFiscEnergyScenarioProto/InvestmentNeeds/README.md).

- **`Mitigation_Investment_Needs_v1.1.xlsx`** — 2030 total and incremental investment
  needs, and incremental cost, for 195 economies × 7 sectors, under three lenses: market
  cost of capital, World Bank rate (6%), and World Bank rate with local co-benefits.
  Engineering-based savings for power and transport; Sensitivity sheet. Uses the
  dashboard's cost-of-capital and PIMA indicators; has a Damodaran input sheet for wider
  WACC coverage.
- **`Mitigation_Investment_Needs_v1.1_summary.png`** — one-page summary of the results
- **`national_model/`** — NMIM v0.2, a simplified national mitigation investment model:
  186 countries, 2024-2050, baseline and net zero, CPAT norms (codes, layout, data,
  activity equation), consistent with the CBAM Policy Tool; Python engine + CPAT-layout
  workbook with live formulas; method note and check report. v0.2 adds a typology of who
  pays (P1 pure public ... P5 pure private) by country and sector, driven by power market
  structure (PSRT, PPI), electricity cost recovery, WACC and subsidies
- **`Mitigation_Investment_Needs_Methodology.md` / `.docx`** — methodology (v1.1 workbook, sector by sector, who pays: investment typology and power market structure, sensitivity, national model and data needs)
- **`CAVEATS.md`**, **`Old/`** — task log and superseded versions
- **`build/`** — `build_investment_model.py` (workbook), `run_sensitivity.py`
  (Sensitivity sheet), `make_summary_image.py` (summary PNG)
- **`raw_data/`** — optional inputs (Damodaran `ctryprem.xlsx`)

## MacroFiscalModels

Simplified Python versions of the IMF GIMF, DIGNAD and DIGNAR models and the World Bank-IMF
LIC-DSF. `integrated.py` chains the dashboard, NMIM (investment needs), PriceFeasibility
(feasible power-sector carbon pricing, whose revenue finances part of the budget need), the
macro model and the DSF for every country, and its results feed the dashboard's Appendix D.
The standalone single-country tool is `MacFiscEnergyScenarioProto/MacroFiscalModels/MacroFiscal_CountryTool_v0.3.xlsx`. See
[`MacFiscEnergyScenarioProto/MacroFiscalModels/README.md`](MacFiscEnergyScenarioProto/MacroFiscalModels/README.md) and the method note
`MacroFiscalModels_Method_v0_1.md` (equations, data needs, validation); pitch deck and user
guides in `MacFiscEnergyScenarioProto/MacroFiscalModels/Docs/`. Calibrations are illustrative.

## TaxMacroFiscalDashboard

Dashboards and data on tax, macro-fiscal, and carbon/climate policy indicators.

- **`CBAM Dashboard/`** — Carbon Border Adjustment Mechanism (CBAM) dashboard,
  emissions factors, and policy guides (e.g. `CBAM_dash_v1.6.xlsx`,
  `Egypt_CBAM_EmissionsFactors_v1.1.xlsx`, `cbam_global_data.xlsx`)
- **`Fiscal Risk Dashboard/`** — Climate fiscal risk dashboard (`CFRiskDashboardV1.5.xlsx`),
  exemplars, and a statistical annex
- **`Political Economy/`** — Political economy metric documentation
- **`TECP + Validation/`** — Total Effective Carbon Price data, elasticities, and
  validation materials (`Total_Effective_Carbon_Price.pdf`, `CPATPriceElasticities.xlsx`)
- **`build/`** — Self-contained Python/R pipeline that builds the Global Fiscal
  Indicators workbook (80 indicators, 12 blocks, including macro-model and debt-sustainability
  inputs, the integrated scenario from MacroFiscalModels and fiscal risk from the Climate Fiscal
  Risk Dashboard v1.5) from raw source files; see
  [SETUP.md](SETUP.md) for requirements and build steps
- **`Global_fiscal_indicators_v1.3.xlsx`** — The dashboard workbook (current version; v1.0 and
  later superseded versions in `Archive/`, history in `VersionNotes.md`)
- **`TaxMacroFiscalDashboard_v1.0.pptx`** — Dashboard overview deck

