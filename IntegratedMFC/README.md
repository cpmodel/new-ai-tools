# IntegratedMFC — Integrated MF-C tool (macro-fiscal climate)

**The tool:** `Integrated_MFC_Tool_v0.1.xlsx`. Pick a country on the **Dashboard** sheet and it
shows the Global Fiscal Indicators dashboard and the country's mitigation investment needs on one
page. No Python needed to use it.

It is built in modules, so more can be added. Now:

| Module | Blocks | Source | Coverage |
|---|---|---|---|
| Macro-fiscal dashboard | Fiscal, PIMA, Finance 1-2, Context 1-3, Appendix A-D (65 indicators) | `TaxMacroFiscalDashboard/` panel, v1.2 | 217 |
| Investment needs | **Investment 1**: 2030 need, by sector, and incremental cost under three cost-of-capital lenses (15) | `InvestmentNeeds/Mitigation_Investment_Needs_v1.1.xlsx` | 194 |
| | **Investment 2**: who pays (P1-P5), budget need, private share, path to 2050 (11) | `InvestmentNeeds/national_model/results/nmim_summary_v0_2.csv` (NMIM v0.2) | 186 |

The investment blocks sit after Context 3, before the appendices, and read High / Medium / Low
(magnitudes, like the Context blocks), with 1-10 scores on the same 5th-95th percentile rescale as
every other indicator. All values are in % of the model's own GDP for the year, except the cost
per tonne (USD/tCO2e) and the private share (%). The year column is the scenario year (2030, 2050
or 2025-2050). Taiwan (TWN) is in v1.1 but not in the dashboard country list, so it is left out.

The workbook has the dashboard's sheets: README, Dashboard, Data_wide, Scores, Spec, Sources,
Data_long. Data_long names the source file for every value; investment rows say
"model output".

## How it is built

The tool does not have its own layout code. `build/build_mfc_tool.py`:

1. runs each module's extractor, which turns a model's results into rows of the dashboard's long
   format (`iso3, country, region, income, is_ccdr, indicator, value, year, source`) and writes
   them to `build/output/<module>_long.csv`;
2. splices the module's indicator rows (`build/config/<module>_indicators.csv`, the dashboard's
   `indicators.csv` columns plus `block_kind`) into the dashboard config after the block named in
   `MODULES`, and adds its sources (`build/config/<module>_sources.csv`);
3. writes the merged panel and config to `build/output/merged/` (not committed) and runs the
   dashboard's own builder, `TaxMacroFiscalDashboard/build/scripts/build_dashboard.py`, on them,
   with the title, file name, version and README text from `build/config/tool.json`.

So the integrated tool scores, colours and reads exactly like the dashboard, and a change to the
dashboard (a new indicator, a fix) reaches it on the next build. The build stops if a module's
config names an indicator its extractor does not emit, or reuses a dashboard code.

```bash
pip install pandas numpy openpyxl
python3 IntegratedMFC/build/build_mfc_tool.py     # -> IntegratedMFC/Integrated_MFC_Tool_v<version>.xlsx
```

It reads repository files only. Rebuild after any of its inputs changes:

- the dashboard panel: `TaxMacroFiscalDashboard/build/scripts/build_global_indicators.py`;
- v1.1: `InvestmentNeeds/build/build_investment_model.py`. The extractor reads the values
  cached in the workbook (Results_Country and Calc), so the workbook must have been recalculated
  and saved (Excel, or LibreOffice headless) after a rebuild;
- NMIM: `InvestmentNeeds/national_model/nmim_v0_2.py`.

### Adding a module

Write an extract function in `build_mfc_tool.py` that returns `iso3, indicator, value, year,
source` rows, add `build/config/<name>_indicators.csv` and `build/config/<name>_sources.csv`,
register it in `MODULES` with the block it follows, and update `tool.json` (README text, version).
Use `block_kind` = `assessment` (Good / Medium / Bad), `context` (High / Medium / Low) or
`appendix` (reference, no block verdict).

## Caveats

- **Model output, not data.** The investment blocks are scenario results with the caveats of
  their models: v1.1 is a Tier 3 allocation of group totals (IHLEG, IEA-IFC) to countries; NMIM's
  typology shares are judgements applied to country data. See `InvestmentNeeds/README.md`,
  `InvestmentNeeds/CAVEATS.md` and `InvestmentNeeds/national_model/NMIM_Method_v0_2.md`.
- **Two models, two answers.** v1.1 (2023 prices, allocated totals) and NMIM (2024 prices,
  bottom-up stocks) give different 2030 needs for the same country. Both are shown, labelled;
  neither is the reference.
- **Overlap with Appendix D.** Appendix D's budget need is NMIM's *incremental* budget need passed
  through MacroFiscalModels; Investment 2's is the budget need of the *whole* net zero programme.
- **Scores rank, they do not judge.** A High investment need is a magnitude, not a verdict; do
  not average the investment blocks with the assessment blocks (see the dashboard's README sheet).

Versions: `VersionNotes.md`. Task log: `CAVEATS.md`.
