# MacroFiscalModels: GIMF-lite, DIGNAD-lite, DIGNAR-lite and LIC-DSF-lite

**The tool:** `MacroFiscal_CountryTool_v0.2.xlsx` — open it, pick a country, see the whole chain. No Python needed.

Simplified Python versions of three IMF / World Bank tools, linked to the National Mitigation
Investment Model (`../InvestmentNeeds/national_model/`, NMIM v0.2), so a country's net zero investment
programme can be run through a macro model and a debt sustainability test.

| File | What it is |
|---|---|
| `gimf_lite.py` | One-country GIMF: Blanchard-Yaari OLG (non-Ricardian) and liquidity-constrained households, sticky prices and wages, Taylor rule, UIP, public capital, fiscal rule; climate add-on for public and private investment and energy savings |
| `dignad_lite.py` | DIGNAD: savers and hand-to-mouth households, standard and adaptation public capital, PIMA efficiency, disasters, grant / concessional / commercial financing; with `rents0 > 0` it is DIGNAR-lite (resource rents, government take, a falling rent path) |
| `lic_dsf.py` | LIC-DSF (2018): Composite Indicator, thresholds, loan-cohort PV, five indicators, stress tests, mechanical rating |
| `nmim_bridge.py` | NMIM → models: public budget need, private investment and energy savings (% GDP, 2025-2050); macro paths → DSF |
| `calibration.py` | One calibration of the shared settings from the dashboard panel; model choice rule; cross-check of the two models |
| `integrated.py` | Integrated country run: dashboard -> NMIM -> PriceFeasibility -> GIMF / DIGNAD / DIGNAR-lite -> LIC-DSF-lite -> `results/integrated_results.csv` (dashboard Appendix D) |
| `build_country_tool.py` | Builds the standalone single-country Excel tool, `MacroFiscal_CountryTool_v<version>.xlsx` (see `CountryTool/README.md`) |
| `ExcelDIGNAD/` | DIGNAD-R Excel: a recursive, formula-only DIGNAD-lite / DIGNAR-lite for 12 countries, dashboard-form inputs, long-form outputs for a master tool (proof of concept; see `ExcelDIGNAD/README.md`) |
| `pricing_bridge.py` | PriceFeasibility's feasible power-sector mix -> carbon-pricing revenue earmarked for the budget need |
| `pf_solver.py` | Perfect-foresight Newton solver (sparse Jacobian) shared by the two macro models |
| `run_demo.py` | Disaster scenarios, the NMIM programme for one country in GIMF-lite and DIGNAD-lite, and the DSF |
| `MacroFiscalModels_Method_v0_1.md` | Method: equations of the full and the lite models, data needs, validation, open issues |
| `VersionNotes.md` | Version notes |
| `CAVEATS.md` | Task log |
| `TODO.md` | Open work, including the plan for mitigation investment effectiveness |
| `Docs/` | Pitch deck and user guides (getting started, GIMF-lite, DIGNAD-lite, LIC-DSF-lite, integrated run); see `Docs/README.md` |

## Run

```bash
pip install numpy scipy pandas openpyxl matplotlib
cd MacroFiscalModels
python3 run_demo.py KEN          # any ISO3 code in NMIM; about 15 seconds
python3 gimf_lite.py             # 1% of GDP public investment test
python3 dignad_lite.py           # disaster test
python3 lic_dsf.py               # DSF on placeholder data
python3 calibration.py KEN SLE   # dashboard calibration, model choice, cross-check
python3 integrated.py KEN SLE    # the whole chain for some countries
python3 integrated.py --all      # all NMIM countries -> results/integrated_{results,paths,settings}.csv (~10 min)
python3 build_country_tool.py    # the standalone country tool -> MacroFiscal_CountryTool_v0.2.xlsx
```

Outputs go to `results/`: `gimf_<ISO>.png`, `dsf_<ISO>.png`, `dignad_disaster.png`, `paths_<ISO>.csv`.

## Use from code

```python
from nmim_bridge import programme, run_gimf
from gimf_lite import Params
prog = programme("SEN")                         # NMIM paths, % of GDP
out, ss0, ssT, cal = run_gimf(prog, Params(psi=0.4, s=0.5, dfs=0.5))
out["y"], out["debt_gdp"], out["pi"], out["q"]  # annual paths from 2025
```

## Caveats

The macro calibrations and the DSF baseline ratios are illustrative placeholders; only the investment
programme is country data. DSF coefficients and thresholds must be checked against the current
Guidance Note. Full list: method note section 6.
