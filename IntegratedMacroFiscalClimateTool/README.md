# IntegratedMacroFiscalClimateTool — Integrated MF-C tool (macro-fiscal climate)

**The tool:** `Integrated_MFC_Tool_v0.6.xlsx`. No Python needed to use it. Tabs, top-down:

| Tab | What it is |
|---|---|
| **Dashboard** | The front page and the only place you drive it: the country selector, the main levers (revenue increase, public investment increase, the climate investment programme on/off, the carbon pricing mode), headline results for 2030 and 2050, and six charts: public debt, revenue and expenditure, GDP vs no new measures, green investment, the carbon price in force, debt sustainability |
| **Indicators** | The Global Fiscal Indicators country profile (formerly the Dashboard tab): 16 blocks, with fiscal risk, the investment needs blocks, the macro blocks and the CDPM-method cross-check |
| **Fiscal outlook** | The scenario item by item, 2024-2050: revenue, expenditure, balance, debt, growth vs trend, with charts |
| **Macro & debt** | The integrated run (pre-computed) beside the live simple model (DIGNAD-R), the other model family as a cross-check, and LIC-DSF-lite debt sustainability, with charts |
| **Green investment** | The National Mitigation Investment Model (NMIM v0.2) live for the country: investment and incremental investment by sector, who pays (P1-P5), the budget need, incremental cost under three lenses, emissions and power, with charts |
| **Carbon pricing** | PriceFeasibility in formulas: carbon tax vs feebate, the feasibility grid, the best feasible mixes |
| **Policy settings** | The secondary levers: start years, ramp-up, share of new taxes on capital, other spending, fiscal rule, feasibility limits, carbon revenue use |
| GI_Inputs, Macro_assumptions | Model assumptions (yellow): NMIM's, and the macro model's and pricing priors |
| Data and engine sheets (grey) | Data_wide, Scores, Spec, Sources, Data_long; GI_Model, GI_Data (NMIM engine); Macro_model, Macro_inputs (DIGNAD-R) |

This folder also holds the tool's source models: `TaxMacroFiscalDashboard/` (the indicator panel
and dashboard), `InvestmentNeeds/` (NMIM and the v1.1 workbook), `MacroFiscalModels/` (GIMF-lite,
DIGNAD/DIGNAR-lite, LIC-DSF-lite, the integrated run) and `PriceFeasibility/` (carbon tax vs feebate
under political-feasibility limits). Paths in this README are relative to this folder.

## The scenario: levers, green investment and carbon pricing

The policy scenario is the no-new-measures baseline plus the levers (main ones on the Dashboard, the
rest on Policy settings):

| Lever | What it does in the model |
|---|---|
| Revenue increase: pp of GDP a year, ceiling (Dashboard); first year, share on capital (Policy settings) | Revenue. The share on capital lowers desired private capital as the base tax does; the rest (VAT-like) only cuts household income |
| Public investment increase, pp of GDP (Dashboard) | Builds public capital at the country's efficiency (PEFA PI-11), which raises output by the return on public capital |
| Other new spending, pp of GDP | Spending only; first year and ramp-up shared with public investment |
| Fiscal rule strength | Scales the speed at which transfers fall as debt rises (1 = as calibrated, 0 = off) |
| Climate investment programme, On/Off (Dashboard) | NMIM's incremental budget need, live from GI_Model, with grants and concessional loans |
| Carbon pricing mode (Dashboard) | None, Manual (your tax and feebate rates), or the best feasible carbon tax, feebate or hybrid |
| Limits F1-F3 | First-year tariff rise, net loss of the bottom 40%, net payment by coal and oil generators (PriceFeasibility's defaults) |
| Phase-in, first year | The rate rises in equal steps; phasing in also lowers F1 |
| Revenue use | Shares rebated to the bottom 40% (lowers F2) and spent on green public investment (builds public capital); the rest goes to the budget |

Green investment is NMIM v0.2 live: its engine sheets (Inputs, Mitigation, Data from
`InvestmentNeeds/national_model/NMIM_v0.2.xlsx`) are copied in by `build/nmim_live.py` as GI_Inputs,
GI_Model and GI_Data, with the country taken from the Dashboard. The climate programme's capital and
spending parts (as `MacroFiscalModels/nmim_bridge.py` splits them), power CO2 and GDP feed the macro
model from GI_Model, so a change on GI_Inputs moves the fiscal outlook: Kenya, raising the public
share of viability-gap investment (`p_ty_p4_gap`) from 30% to 60% lifts the 2030 budget need from
0.49% to 0.69% of GDP and 2050 debt from 75.2% to 79.0%. At NMIM's defaults the Green investment tab
equals `nmim_summary_v0_2.csv`.

Carbon pricing is PriceFeasibility v0.1 in formulas. For the selected country the Carbon pricing sheet
computes abatement and F1-F3 for any tax and feebate, shows a tax-by-feebate grid with the
feasible cells in green, and finds the best feasible mix live: tax only and feebate only exactly
(each limit is linear in the rate), the hybrid by searching the F3 limit line. At the default
limits it matches PriceFeasibility's own optimiser to its grid step (Kenya: tax 88.9 vs 88 USD/t).
Revenue = tax rate x fossil power CO2 / NMIM GDP, on NMIM's net zero path when the climate
programme is on (its abatement is already in that path), otherwise NMIM's baseline less the
pricing abatement. A feebate raises nothing and, in this model, costs no growth.

Everything is live (DIGNAD-R, Macro_model), for the 186 countries of the integrated run.
`build/outlook.py --check ISO3 '{"lever": value}'` computes the same scenario in Python (programme from
the NMIM engine): for Kenya, Indonesia, South Africa and India under several lever settings, with the
programme on and off, the workbook (recalculated in LibreOffice) and Python agree to 1e-14.

Charts (`build/xlcharts.py`) share one style: years on a numeric axis every five years, legend below,
colours by role (policy, baseline, integrated run, other model, limits), hidden chart data plotted.

Limits: DIGNAD-R has no trend growth, so the baseline holds the 2024 position (debt, revenue and
spending ratios) except that fossil rents fall, and growth is shown above or below trend. It is a
reference point for comparing packages, not a WEO forecast. The integrated run (Macro & debt, Macro
blocks) is pre-computed with PriceFeasibility's default hybrid and does not respond to the levers.

## Modules

It is built in modules, so more can be added. Now:

| Module | Blocks | Source | Coverage |
|---|---|---|---|
| Macro-fiscal dashboard | Fiscal, PIMA, Finance 1-2, Context 1-3, Appendix A-C and E (74 indicators; Appendix D replaced by the macro-fiscal module). Appendix C has the model and DSF inputs, Appendix E fiscal risk from the Climate Fiscal Risk Dashboard v1.5 | `TaxMacroFiscalDashboard/` panel, v1.3 | 217 |
| Investment needs | **Investment 1**, need and cost: net zero and incremental investment 2030, incremental by sector, incremental 2050, incremental cost under three cost-of-capital lenses, cost-of-capital penalty, cost per tonne, abatement, PV of cost 2025-2050 (18) | NMIM v0.2, `InvestmentNeeds/national_model/results/nmim_summary_v0_2.csv` | 186 |
| | **Investment 2**, public investment needs: direct budget need (P1 + public gap of P4) and its parts, incremental budget need, PIMA-adjusted budget need, plus regulated utilities, all investment needing public money, private capital and its parts, private share (11) | NMIM v0.2; PIMA adjustment also uses the dashboard's PEFA PI-11 | 186 |
| Macro-fiscal | **Macro 1**, financing: financing gap = incremental budget need less feasible power carbon-pricing revenue, 2030; pricing share of the need 2025-2050; feasible carbon tax; abatement (6) | MacroFiscalModels integrated run (`results/integrated_results.csv`, PriceFeasibility inside) | 186 |
| | **Macro 2**, growth and debt: GDP vs today's path 2030 and 2050, GDP vs no-programme baseline 2050, public debt 2030 and 2050, debt added by the programme, cross-check with the other model family (GDP and debt 2050), LIC-DSF Composite Indicator, LIC-DSF-lite risk before and with the programme (11) | integrated run (`integrated_results.csv`, `integrated_paths.csv`) | 184 (DSF: 66 LIC/LMIC) |
| | **Macro & debt tab**: the integrated run's GDP and debt paths, the other model family's, and **DIGNAD-R**, the simple recursive DIGNAD-lite / DIGNAR-lite in formulas, for the Dashboard's country, with charts and yellow assumptions | `MacroFiscalModels/ExcelDIGNAD/build_dignad_excel.py` (equations, assumptions, inputs) | 186 |
| | **Macro & debt tab, debt sustainability**: LIC-DSF-lite's five indicators against the thresholds for the country's debt-carrying capacity, 2025-2034, before and with the programme, breaches, ratings and charts | integrated run (`integrated_paths.csv`) | 66 LIC/LMIC |
| CDPM-method cross-check | **Appendix F**: total and incremental need, incremental cost (6%), and public need at a flat group share, 2030, from the CDPM-based global method (4) | `InvestmentNeeds/Mitigation_Investment_Needs_v1.1.xlsx` (cached values) | 194 |

The investment blocks sit after Context 3, the macro blocks after them, then the appendices. The investment blocks read High / Medium / Low
(magnitudes, like the Context blocks), with 1-10 scores on the same 5th-95th percentile rescale as
every other indicator. Values are in % of NMIM's GDP for the year, except the cost per tonne
(USD/tCO2e), abatement (% of baseline emissions) and the private share (%). The year column is the
scenario year (2030, 2050 or 2025-2050).

**Why NMIM.** InvestmentNeeds has two models. The v1.1 workbook allocates published group totals
(IHLEG, IEA-IFC) to countries for 2030 and applies one public share per country group, so it has
no country estimate of public needs. NMIM builds the need country by country to 2050, splits it by
who pays, and is the model MacroFiscalModels runs (Appendix D). The tool's investment blocks use
NMIM; v1.1, the CDPM-based method, appears only as the Appendix F cross-check. The CDPM model itself
(CDPM / SiSePuede pathways) is not in these repositories, so its own outputs cannot be added yet.

**Macro-fiscal: two models, on purpose.**

- *Cross-check* (pre-computed): the integrated run also solves the other model family on the
  same programme - DIGNAD-lite where GIMF-lite was chosen, GIMF-lite otherwise (no resource
  block, so for DIGNAR-lite countries it leaves out the fall in rents). A large gap between the two
  says the answer depends on the model, not only on the data.
- *Integrated run* (pre-computed, Python, `MacroFiscalModels/integrated.py`): one model per
  country by `calibration.py`'s rule - GIMF-lite (135 countries, mostly upper-middle and high
  income, market finance), DIGNAD-lite (35, low income and concessional borrowers), DIGNAR-lite
  (16, fossil rents of 10% of GDP or more) - solved with perfect foresight; then LIC-DSF-lite for
  low and lower-middle income. Two countries are not financeable under the fiscal rule and have no
  macro results. This is what the Macro blocks show.
- *DIGNAD-R* (live, formulas, Macro & debt tab): the same DIGNAD-lite / DIGNAR-lite equations made
  recursive, so they run in Excel and respond to the yellow assumptions (disaster, grant and
  concessional shares, pricing earmark, fiscal rule). It is DIGNAD-type for every country, has
  one fiscal instrument and no energy savings or private investment push, so it sits below the
  integrated run for GIMF-lite countries and where the fiscal rule had to tighten. Its equations,
  assumptions and input rules are imported from `ExcelDIGNAD/build_dignad_excel.py`, so the two
  workbooks are one model; this one computes the selected country only, for all 186.

**Public investment needs, narrow to wide** (Investment 2):

| Measure | Definition |
|---|---|
| Direct budget need | P1 pure public + 30% of P4 viability gap (NMIM `ty_p4_gap`): what the budget spends |
| incremental | the same, net zero less baseline; equals Appendix D's budget need |
| PIMA-adjusted | budget need / (1 - efficiency gap); gap from PEFA PI-11 where available (40% at 1 to 13% at 4), else IMF (2015) by income (AE 13%, low income 40%, others 27%), as in the v1.1 workbook |
| Plus regulated utilities | + P2: utility investment repaid by tariffs, a contingent liability where tariffs do not recover costs |
| All needing public money or a decision | P1 + P2 + all of P4 |
| Private capital | P3 + P5 + 70% of P4 |

The workbook has the dashboard's sheets: README, Indicators (the dashboard's Dashboard, renamed), Data_wide, Scores, Spec, Sources,
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
   with the title, file name, version and README text from `build/config/tool.json`. A module can
   replace dashboard blocks (macro-fiscal replaces Appendix D);
4. adds the sheets the dashboard builder does not make: `build/macro_live.py` writes Macro,
   Macro_assumptions, Macro_model and Macro_inputs and the Macro & debt tab; `build/nmim_live.py` copies
   NMIM's engine in (GI_Inputs, GI_Model, GI_Data) and writes Green investment; `build/outlook.py`
   writes the Dashboard (the selector every tab follows), Fiscal outlook, Carbon pricing and Policy
   settings. It renames the dashboard's Dashboard tab to Indicators.

So the integrated tool scores, colours and reads exactly like the dashboard, and a change to the
dashboard (a new indicator, a fix) reaches it on the next build. The build stops if a module's
config names an indicator its extractor does not emit, or reuses a dashboard code.

```bash
pip install pandas numpy openpyxl
python3 build/build_mfc_tool.py     # -> Integrated_MFC_Tool_v<version>.xlsx (run from this folder or anywhere)
```

It reads repository files only. Rebuild after any of its inputs changes:

- NMIM's workbook (`InvestmentNeeds/national_model/build_workbook_v0_2.py`), whose engine sheets are copied in;

- PriceFeasibility: `PriceFeasibility/price_feasibility.py` (parameters and reference optima);

- the dashboard panel: `TaxMacroFiscalDashboard/build/scripts/build_global_indicators.py`;
- NMIM: `InvestmentNeeds/national_model/nmim_v0_2.py`;
- the v1.1 workbook (Appendix F): `InvestmentNeeds/build/build_investment_model.py`, then
  recalculate and save it (the extractor reads cached values);
- the integrated run: `MacroFiscalModels/integrated.py --all` (about 3 minutes; after NMIM and
  the dashboard panel).

`python3 build/macro_live.py --check KEN AGO` prints DIGNAD-R's results from its
Python copy of the equations, to compare with the Macro & debt tab after recalculation.

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
- **Debt sustainability is simplified.** LIC-DSF-lite takes debt-carrying capacity from the
  Composite Indicator where the dashboard has a CPIA score (IDA countries) and medium otherwise;
  exports (all external ratios' denominators) and imports (reserve cover) are still assumptions,
  because no country trade series is in the repository sources; no MAC-DSA for market-access
  countries.
- **Macro calibrations are illustrative** (MacroFiscalModels method note, section 6): four
  DIGNAD-R inputs (public investment, return on public capital, hand-to-mouth share, grants) are
  income-group defaults; DSF coefficients and thresholds still need checking against the Guidance
  Note; the DSF baseline ratios are placeholders.
- **The Macro & debt tab's two columns differ by construction** (see above). Read the integrated run
  for the number, the simple model for what moves it.
- **PIMA adjustment** treats the whole budget need as public investment subject to the
  efficiency gap; PEFA PI-11 covers 82 countries, the rest use income-group gaps.
- **Scores rank, they do not judge.** A High investment need is a magnitude, not a verdict; do
  not average the investment blocks with the assessment blocks (see the dashboard's README sheet).

## Handover: rebuilding everything, and open decisions

Full rebuild after a change to raw data, in order, from this folder (about 6 minutes in all):

```bash
python3 TaxMacroFiscalDashboard/build/scripts/build_global_indicators.py   # panel
python3 PriceFeasibility/price_feasibility.py                              # pricing optima
python3 MacroFiscalModels/integrated.py --all                               # about 3 minutes
python3 TaxMacroFiscalDashboard/build/scripts/build_global_indicators.py   # panel again (Appendix D)
python3 TaxMacroFiscalDashboard/build/scripts/build_dashboard.py           # dashboard workbook
python3 MacroFiscalModels/ExcelDIGNAD/build_dignad_excel.py --check        # DIGNAD-R Excel
python3 MacroFiscalModels/build_country_tool.py                             # country tool
python3 build/build_mfc_tool.py                                             # this tool
```

Requirements beyond the dashboard's: `scipy` (macro models), `pyxlsb` and `xlsxwriter` (panel).
R is needed only to check the R pipeline. Rebuild NMIM or the v1.1 workbook first if their inputs
changed (and recalculate v1.1). Bump versions and archive per each folder's `VersionNotes.md`.

Decisions waiting for the owner (details in `MacroFiscalModels/TODO.md` and `CAVEATS.md`):

- **Model choice for concessional borrowers:** keep the IDS multilateral share as the
  concessional proxy, or switch to the concessional share of external debt now in Appendix C
  (moves countries between DIGNAD-lite and GIMF-lite).
- **Which channels to credit:** DIGNAD-lite and GIMF-lite differ most on energy savings and
  private investment (Kenya 2050: +0.3% vs +4.6% of GDP); see TODO 1, 10 and 36.
- **Hand-to-mouth proxy:** use Findex account ownership or keep the income-group default.
- **Trade data:** a manual WDI download (exports and imports of goods and services) would replace
  the DSF's two biggest assumptions (TODO 11); for Kenya it can move capacity from medium to strong.

A worked country example: [`briefings/Kenya_briefing.md`](briefings/Kenya_briefing.md) (Word: `briefings/Kenya_briefing.docx`, built with `pandoc Kenya_briefing.md -o Kenya_briefing.docx --reference-doc=../../Presentations/Global_Investment_Needs_SISEPUEDE.docx`).

Versions: `VersionNotes.md`. Task log: `CAVEATS.md`.
