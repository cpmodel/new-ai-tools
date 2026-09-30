# Fiscal Climate AI Tools

A collection of working papers, data workbooks, and process files for fiscal and
climate-policy analysis, organized by project.

See **[SETUP.md](SETUP.md)** for environment requirements, build instructions, and
run-process specifications consolidated from the subfolder documentation.

## Repository structure

```
fiscal-climate-ai-tools/
├── SETUP.md                     Consolidated requirements, build steps, run process
├── ExciseDiagnostic/            Excise tax diagnostic working paper and data
└── TaxMacroFiscalDashboard/     Tax, macro-fiscal, and climate policy dashboards
    └── build/                   Raw data, config, and scripts for the Global Fiscal
                                  Indicators dashboard (Python + R build pipeline)
```

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
  Indicators workbook (53 indicators, 9 blocks) from raw source files; see
  [SETUP.md](SETUP.md) for requirements and build steps
- **`Tax+Macro-Fiscal-Energy_indicators.xlsx`** — Combined tax, macro-fiscal, and
  energy indicators workbook
- **`TaxMacroFiscalDashboard_v1.0.pptx`** — Dashboard overview deck

