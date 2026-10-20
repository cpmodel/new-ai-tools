# CAVEATS - log of completed tasks (InvestmentNeeds)

Append-only log of completed work in `InvestmentNeeds/`, following the CPAT norm (`CPAT_excel_not_master/NORMS.md` section 6): what was asked, what went in, what came out, and what the next person must know. **Newest entry at the bottom.** Never edit or delete an earlier entry; to resolve a caveat, add a new entry that names it.

## Entry template

```markdown
## YYYY-MM-DD - <task / short name> (<version>)
- **Task:**
- **Inputs:**
- **Outputs:**
- **Caveats:**
```

---

## Entries

## 2026-10-19 - Mitigation investment needs workbook (v1.0)
- **Task:** Excel model of 2030 total and incremental mitigation investment and incremental cost by country (195) and sector (7) under market WACC, World Bank 6% and 6% + local co-benefits, using the dashboard's cost-of-capital and PIMA indicators.
- **Inputs:** Global Investment Needs SISEPUEDE method note (Oct 2026); dashboard panel (`global_data_wide.csv`), CPAT 2024 panel, WEO, Steffen et al. WACC database; IHLEG (2024), IEA-IFC (2023), McKinsey (2022), CCDR synthesis.
- **Outputs:** `Old/Mitigation_Investment_Needs_v1.0.xlsx`, `Old/Mitigation_Investment_Needs_v1.0_summary.png`, `build/build_investment_model.py` (v1.0 state in git history).
- **Caveats:** Tier 3 allocation of group totals; one savings ratio (b = 0.133, SISEPUEDE LAC) set the sign of every lens; Damodaran sheet empty (host blocked); PIMA scores not available (PEFA mapping used).

## 2026-10-20 - Engineering benefits for power and transport (v1.1)
- **Task:** Strengthen the weakest parts of v1.0: physical savings, co-benefits and abatement for power and transport; no benefits for grids; cost-per-tonne for fuels and AFOLU; one-at-a-time sensitivity.
- **Inputs:** v1.0; Ember generation data; CPAT supply and air-pollution costs; IPCC 2006 Table 2.2 factors; IPCC AR6 SPM C.9.1 and C.12.
- **Outputs:** `Mitigation_Investment_Needs_v1.1.xlsx` (with Sensitivity sheet), `Mitigation_Investment_Needs_v1.1_summary.png`, `build/run_sensitivity.py`, `build/make_summary_image.py`; methodology doc (shared Claude doc).
- **Caveats:** Buildings and industry still use the calibrated b and c; power is conservative (no credit for fossil plants avoided; 2024 fuel prices); EV premium, end-use split and public shares are assumptions; Sensitivity sheet holds static values (re-run after changes).

## 2026-10-20 - National Mitigation Investment Model (NMIM v0.1)
- **Task:** Simplified national model, 186 countries, annual 2024-2050, baseline and net zero, built to CPAT norms and consistent with CPAT's activity equation, the CBAM Policy Tool and the v1.1 workbook.
- **Inputs:** CPAT 2024 panel (energy, prices, air pollution, emissions); CPAT income-group elasticities (mitigation MVP data); Ember; WEO; CBAM Policy Tool global v1.0 Country data; v1.1 WACC; IEA World Energy Investment 2024 (calibration); IRENA 2024 costs; CPAT PowerTech; IPCC 2006 factors.
- **Outputs:** `national_model/` - `extract_data_v0_1.py`, `nmim_v0_1.py`, `build_workbook_v0_1.py`, `check_v0_1.py`, `NMIM_v0.1.xlsx`, `NMIM_Method_v0_1.md`, `VersionNotes_NMIM.md`, `check_report_v0_1.md`, `PLAN_v0_1.md`, `data/*.csv`, `results/nmim_summary_v0_1.csv` (the 8 MB long results file is regenerated, not committed).
- **Caveats:** (1) No price response or carbon pricing; real prices fixed at 2024. (2) Power follows target dates, not least cost; WACC does not change the mix. (3) Vehicles are car-equivalent per GJ. (4) Transformation dates and buildings/industry unit costs are assumptions (†). (5) Calibration term (IEA 2024 by group) sits in both scenarios: levels tie to IEA/IHLEG, incremental is bottom-up and may be overstated in buildings and industry. (6) 10 Gt residual GHG in 2050 (no removals modelled). (7) 193 of 3,936 Mitigation formulas exceed 100 characters. (8) Mitigation horizon extends the norm's 2045 to 2050; Inputs follow the MTInputs column structure but not its rows (new parameters). (9) Damodaran still not loaded.

## 2026-10-20 - Investment typology and power market structure (NMIM v0.2)
- **Task:** Typology of mitigation investment by sector and type of payer (pure public, regulated utility, contracted revenue, viability gap, pure private), integrated into the national model by country, and power market structure data.
- **Inputs:** Power Sector Reform Tracker v1 (dashboard raw data, ends 2013); World Bank PPI 2024; IMF CPAT 2024 electricity retail prices, supply costs and subsidies; dashboard RISE and explicit fossil subsidies; v1.1 WACC; IEA-IFC (2023) for the private-share benchmark.
- **Outputs:** `national_model/` v0.2 - `extract_data_v0_2.py` (adds `data/power_market.csv`), `nmim_v0_2.py` (section 12), `build_workbook_v0_2.py`, `check_v0_2.py`, `NMIM_v0.2.xlsx` (Typology, PowerMarket sheets), `NMIM_Method_v0_2.md`, `check_report_v0_2.md`; v0.1 workbook, plan, check report and summary moved to `national_model/Old/`. Methodology (`.md`, `.docx`, shared Claude doc) gains a *Who pays* chapter; figures `typology_spectrum.png`, `typology_by_group.png`. Typology doc (Claude doc) kept as the stand-alone version.
- **Caveats:** (1) Type shares and pressure-index spans are judgements (†), applied to country data, not observed splits. (2) PSRT ends in 2013; 48 countries curated and 12 updated by judgement (notes on PowerMarket sheet) and need review by power sector staff. (3) PPI misses domestic private investment and advanced economies; used only to raise structure 1 to 2. (4) Cost recovery is average retail price over CPAT supply cost; ignores collection losses and arrears. (5) The typology splits investment, not incremental cost; the budget need counts P1 plus 30% of P4.

