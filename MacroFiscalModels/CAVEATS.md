# CAVEATS - log of completed tasks (MacroFiscalModels)

Append-only log, following the CPAT norm (`../CPAT_excel_not_master/NORMS.md` section 6). Newest entry at the bottom.

## Entries

## 2026-10-20 - GIMF-lite, DIGNAD-lite, LIC-DSF-lite and the NMIM link (v0.1)
- **Task:** Main equations and data needs of GIMF, DIGNAD and the LIC-DSF; simplified country-level versions; link to the investment needs model.
- **Inputs:** Kumhof, Laxton, Muir and Mursula (2010), IMF WP/10/34; Marto, Papageorgiou and Klyuev (2018); Buffie et al. (2012); World Bank-IMF LIC-DSF Guidance Note (2018), from memory, not re-checked against the documents; NMIM v0.2 engine.
- **Outputs:** `gimf_lite.py`, `dignad_lite.py`, `lic_dsf.py`, `nmim_bridge.py`, `pf_solver.py`, `run_demo.py`, `MacroFiscalModels_Method_v0_1.md`, `README.md`, `VersionNotes.md`, `results/` (Kenya example).
- **Caveats:** (1) Calibrations are illustrative, not country calibrations. (2) DSF CI coefficients, thresholds and shock sizes need checking against the current Guidance Note and template. (3) Energy savings enter GIMF-lite as a value-added gain, which drives about two-thirds of the long-run GDP gain in the Kenya example. (4) The DSF demo baseline uses placeholder ratios scaled to NMIM GDP. (5) GIMF-lite is one country and one good; no oil block, financial accelerator, zero lower bound or carbon pricing.

## 2026-10-20 - Pitch deck and user guides (v0.1)
- **Task:** A pitch deck and user guides for the three models, in a subfolder of the repository.
- **Inputs:** The models and results of the previous entry; house style of `Docs and PPTs/AIClimate-FiscalTools_v2.5.pptx` (navy, teal, orange; Cambria and Calibri); `Docs and PPTs/Global_Investment_Needs_SISEPUEDE.docx` as the Word reference document.
- **Outputs:** `Docs/MacroFiscalModels_Pitch_v0_1.pptx` (11 slides, native charts, speaker notes), `Docs/UserGuide_0_GettingStarted`, `_1_GIMF_lite`, `_2_DIGNAD_lite`, `_3_LIC_DSF_lite` (`.md` sources and `.docx`), `Docs/build_pitch.js`, `Docs/export_deck_data.py`, `Docs/deck_data.json`, `Docs/build_docs.sh`, `Docs/README.md`.
- **Caveats:** (1) The deck's Kenya numbers carry every caveat of the previous entry: illustrative calibration, placeholder DSF baseline. (2) The deck was checked in LibreOffice renders, not in PowerPoint. (3) Every code snippet in the guides was run; parameter source suggestions are pointers, not calibrations.

## 2026-10-20 - Shared calibration of GIMF-lite and DIGNAD-lite (option 1)
- **Task:** Harmonise the settings the two macro models share, from one calibration; a rule for choosing the model; a cross-check run.
- **Inputs:** Dashboard panel (`global_data_wide.csv`): gross debt, tax revenue, PEFA PI-11, IDS multilateral share and private creditor rate, income group; PIMA mapping from InvestmentNeeds v1.1; IMF (2015) efficiency gaps.
- **Outputs:** `calibration.py` (`shared`, `to_gimf`, `to_dignad`, `choose_model`, `cross_check`, CLI); `nmim_bridge.programme` returns `spend_share`; method note section 4a; README; user guide 0 (section on calibration, Word rebuilt); TODO items 5-9.
- **Caveats:** (1) Return on public capital, public investment, constrained share, grants, net foreign debt and the tax split are income-group assumptions, flagged as such. (2) The IDS multilateral share is of PPG external debt, used as the concessional share of all public debt. (3) The borrowing cost is raised for 7 low-income countries so GIMF-lite has a discount factor below 1. (4) With shared settings GIMF-lite still gives about 1.6 times DIGNAD-lite's GDP gain, explained by labour supply (checked with `wage_el` = 50). (5) Non-CO2 and land-use spending are treated as productive capital: Liberia's results (+14-22% GDP, debt 117-119% by 2050) are not meaningful until TODO item 7 is done. (6) Model defaults are unchanged, so the demo, deck and guides keep their illustrative numbers.
