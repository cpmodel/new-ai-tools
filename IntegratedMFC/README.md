# IntegratedMFC — Integrated MF-C tool (macro-fiscal climate)

**The tool:** `Integrated_MFC_Tool_v0.2.xlsx`. Pick a country on the **Dashboard** sheet and it
shows the Global Fiscal Indicators dashboard and the country's mitigation investment needs,
including its public investment needs, on one page. No Python needed to use it.

It is built in modules, so more can be added. Now:

| Module | Blocks | Source | Coverage |
|---|---|---|---|
| Macro-fiscal dashboard | Fiscal, PIMA, Finance 1-2, Context 1-3, Appendix A-D (65 indicators) | `TaxMacroFiscalDashboard/` panel, v1.2 | 217 |
| Investment needs | **Investment 1**, need and cost: net zero and incremental investment 2030, incremental by sector, incremental 2050, incremental cost under three cost-of-capital lenses, cost-of-capital penalty, cost per tonne, abatement, PV of cost 2025-2050 (18) | NMIM v0.2, `InvestmentNeeds/national_model/results/nmim_summary_v0_2.csv` | 186 |
| | **Investment 2**, public investment needs: direct budget need (P1 + public gap of P4) and its parts, incremental budget need, PIMA-adjusted budget need, plus regulated utilities, all investment needing public money, private capital and its parts, private share (11) | NMIM v0.2; PIMA adjustment also uses the dashboard's PEFA PI-11 | 186 |

The investment blocks sit after Context 3, before the appendices, and read High / Medium / Low
(magnitudes, like the Context blocks), with 1-10 scores on the same 5th-95th percentile rescale as
every other indicator. Values are in % of NMIM's GDP for the year, except the cost per tonne
(USD/tCO2e), abatement (% of baseline emissions) and the private share (%). The year column is the
scenario year (2030, 2050 or 2025-2050).

**Why NMIM.** InvestmentNeeds has two models. The v1.1 workbook allocates published group totals
(IHLEG, IEA-IFC) to countries for 2030 and applies one public share per country group, so it has
no country estimate of public needs. NMIM builds the need country by country to 2050, splits it by
who pays, and is the model MacroFiscalModels runs (Appendix D). The tool uses NMIM only; v1.1 stays
in `InvestmentNeeds/` as the global reference.

**Public investment needs, narrow to wide** (Investment 2):

| Measure | Definition |
|---|---|
| Direct budget need | P1 pure public + 30% of P4 viability gap (NMIM `ty_p4_gap`): what the budget spends |
| incremental | the same, net zero less baseline; equals Appendix D's budget need |
| PIMA-adjusted | budget need / (1 - efficiency gap); gap from PEFA PI-11 where available (40% at 1 to 13% at 4), else IMF (2015) by income (AE 13%, low income 40%, others 27%), as in the v1.1 workbook |
| Plus regulated utilities | + P2: utility investment repaid by tariffs, a contingent liability where tariffs do not recover costs |
| All needing public money or a decision | P1 + P2 + all of P4 |
| Private capital | P3 + P5 + 70% of P4 |

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
- NMIM: `InvestmentNeeds/national_model/nmim_v0_2.py`.

### Adding a module

Write an extract function in `build_mfc_tool.py` that returns `iso3, indicator, value, year,
source` rows, add `build/config/<name>_indicators.csv` and `build/config/<name>_sources.csv`,
register it in `MODULES` with the block it follows, and update `tool.json` (README text, version).
Use `block_kind` = `assessment` (Good / Medium / Bad), `context` (High / Medium / Low) or
`appendix` (reference, no block verdict).

## Caveats

- **Model output, not data.** The investment blocks are NMIM scenario results with NMIM's
  caveats (`InvestmentNeeds/national_model/NMIM_Method_v0_2.md`, `InvestmentNeeds/CAVEATS.md`):
  no price response, power follows target dates, buildings and industry unit costs and the
  incremental need there may be overstated.
- **Who pays is a judgement.** The P1-P5 shares, the 30% public gap of P4 and the pressure
  indices are assumptions applied to country data, not observed splits. Power market structure
  rests on the Power Sector Reform Tracker (ends 2013), with 60 countries set or updated by
  judgement; power sector staff should review them.
- **The typology splits investment, not incremental cost.**
- **PIMA adjustment** treats the whole budget need as public investment subject to the
  efficiency gap; PEFA PI-11 covers 82 countries, the rest use income-group gaps.
- **Scores rank, they do not judge.** A High investment need is a magnitude, not a verdict; do
  not average the investment blocks with the assessment blocks (see the dashboard's README sheet).

Versions: `VersionNotes.md`. Task log: `CAVEATS.md`.
