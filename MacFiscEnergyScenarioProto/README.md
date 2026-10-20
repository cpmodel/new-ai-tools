# IntegratedMacroFiscalClimateTool — Integrated MF-C tool (macro-fiscal climate)

**The tool:** `Integrated_MFC_Tool_v0.5.xlsx`. It opens on the **Fiscal outlook**: pick a country,
set the policy levers, and read the revenue and expenditure outlook, balance, debt, growth and
green public investment to 2050 (see *Fiscal outlook and carbon pricing* below). The **Pricing**
sheet finds the most effective politically feasible carbon tax and feebate mix. The **Dashboard**
follows the same country and shows the Global Fiscal Indicators dashboard (with fiscal risk), the country's mitigation
investment needs (including its public investment needs, and a CDPM-method cross-check), and
what the programme does to growth, public debt and debt risk. The **Macro** sheet has the same
country's paths, with the other model family as a cross-check and a live simple macro model you
can change; the **DSA** sheet has the debt-sustainability indicators against their thresholds.
No Python needed to use it.

This folder also holds the tool's source models: `TaxMacroFiscalDashboard/` (the indicator panel
and dashboard), `InvestmentNeeds/` (NMIM and the v1.1 workbook), `MacroFiscalModels/` (GIMF-lite,
DIGNAD/DIGNAR-lite, LIC-DSF-lite, the integrated run) and `PriceFeasibility/` (carbon tax vs feebate
under political-feasibility limits). Paths in this README are relative to this folder.

## Fiscal outlook and carbon pricing

The front page. The policy scenario is the no-new-measures baseline plus:

| Lever | What it does in the model |
|---|---|
| New taxes: pp of GDP a year, first year, ceiling | Revenue. The share on capital lowers desired private capital as the base tax does; the rest (VAT-like) only cuts household income |
| New public investment, pp of GDP | Builds public capital at the country's efficiency (PEFA PI-11), which raises output by the return on public capital |
| Other new spending, pp of GDP | Spending only; first year and ramp-up shared with public investment |
| Fiscal rule strength | Scales the speed at which transfers fall as debt rises (1 = as calibrated, 0 = off) |
| Climate programme (`p_programme`, Macro_assumptions) | NMIM's incremental budget need, with grants and concessional loans, as before |
| Carbon pricing mode | None, Manual (your tax and feebate rates), or the best feasible carbon tax, feebate or hybrid |
| Limits F1-F3 | First-year tariff rise, net loss of the bottom 40%, net payment by coal and oil generators (PriceFeasibility's defaults) |
| Phase-in, first year | The rate rises in equal steps; phasing in also lowers F1 |
| Revenue use | Shares rebated to the bottom 40% (lowers F2) and spent on green public investment (builds public capital); the rest goes to the budget |

Carbon pricing is PriceFeasibility v0.1 in formulas. For the selected country the Pricing sheet
computes abatement and F1-F3 for any tax and feebate, shows a tax-by-feebate grid with the
feasible cells in green, and finds the best feasible mix live: tax only and feebate only exactly
(each limit is linear in the rate), the hybrid by searching the F3 limit line. At the default
limits it matches PriceFeasibility's own optimiser to its grid step (Kenya: tax 88.9 vs 88 USD/t).
Revenue = tax rate x fossil power CO2 / NMIM GDP, on NMIM's net zero path when the climate
programme is on (its abatement is already in that path), otherwise NMIM's baseline less the
pricing abatement. A feebate raises nothing and, in this model, costs no growth.

Everything is live (DIGNAD-R, Macro_model), for the 186 countries of the integrated run.
`build/outlook.py --check ISO3 '{"lever": value}'` computes the same scenario in Python: for Kenya,
Indonesia, South Africa and Nigeria under several lever settings the workbook (recalculated in
LibreOffice) and Python agree to 1e-14.

Limits: DIGNAD-R has no trend growth, so the baseline holds the 2024 position (debt, revenue and
spending ratios) except that fossil rents fall, and growth is shown above or below trend. It is a
reference point for comparing packages, not a WEO forecast. The integrated run (Macro sheet, Macro
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
| | **Macro sheet**: the integrated run's GDP and debt paths, the other model family's, and **DIGNAD-R**, the simple recursive DIGNAD-lite / DIGNAR-lite in formulas, for the Dashboard's country, with charts and yellow assumptions | `MacroFiscalModels/ExcelDIGNAD/build_dignad_excel.py` (equations, assumptions, inputs) | 186 |
| | **DSA sheet**: LIC-DSF-lite's five indicators against the thresholds for the country's debt-carrying capacity, 2025-2034, before and with the programme, breaches, ratings and charts | integrated run (`integrated_paths.csv`) | 66 LIC/LMIC |
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
- *DIGNAD-R* (live, formulas, Macro sheet): the same DIGNAD-lite / DIGNAR-lite equations made
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
   with the title, file name, version and README text from `build/config/tool.json`. A module can
   replace dashboard blocks (macro-fiscal replaces Appendix D);
4. adds the sheets the dashboard builder does not make: `build/macro_live.py` writes Macro,
   Macro_assumptions, Macro_model and Macro_inputs, and `build/outlook.py` writes Fiscal outlook
   (the front page, with the country picker the Dashboard now follows) and Pricing.

So the integrated tool scores, colours and reads exactly like the dashboard, and a change to the
dashboard (a new indicator, a fix) reaches it on the next build. The build stops if a module's
config names an indicator its extractor does not emit, or reuses a dashboard code.

```bash
pip install pandas numpy openpyxl
python3 build/build_mfc_tool.py     # -> Integrated_MFC_Tool_v<version>.xlsx (run from this folder or anywhere)
```

It reads repository files only. Rebuild after any of its inputs changes:

- PriceFeasibility: `PriceFeasibility/price_feasibility.py` (parameters and reference optima);

- the dashboard panel: `TaxMacroFiscalDashboard/build/scripts/build_global_indicators.py`;
- NMIM: `InvestmentNeeds/national_model/nmim_v0_2.py`;
- the v1.1 workbook (Appendix F): `InvestmentNeeds/build/build_investment_model.py`, then
  recalculate and save it (the extractor reads cached values);
- the integrated run: `MacroFiscalModels/integrated.py --all` (about 3 minutes; after NMIM and
  the dashboard panel).

`python3 build/macro_live.py --check KEN AGO` prints DIGNAD-R's results from its
Python copy of the equations, to compare with the Macro sheet after recalculation.

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
- **The Macro sheet's two columns differ by construction** (see above). Read the integrated run
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
