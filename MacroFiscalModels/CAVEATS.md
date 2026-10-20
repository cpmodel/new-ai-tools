# CAVEATS - log of completed tasks (MacroFiscalModels)

Append-only log, following the CPAT norm (`../CPAT_excel_not_master/NORMS.md` section 6). Newest entry at the bottom.

## Entries

## 2026-10-20 - GIMF-lite, DIGNAD-lite, LIC-DSF-lite and the NMIM link (v0.1)
- **Task:** Main equations and data needs of GIMF, DIGNAD and the LIC-DSF; simplified country-level versions; link to the investment needs model.
- **Inputs:** Kumhof, Laxton, Muir and Mursula (2010), IMF WP/10/34; Marto, Papageorgiou and Klyuev (2018); Buffie et al. (2012); World Bank-IMF LIC-DSF Guidance Note (2018), from memory, not re-checked against the documents; NMIM v0.2 engine.
- **Outputs:** `gimf_lite.py`, `dignad_lite.py`, `lic_dsf.py`, `nmim_bridge.py`, `pf_solver.py`, `run_demo.py`, `MacroFiscalModels_Method_v0_1.md`, `README.md`, `VersionNotes.md`, `results/` (Kenya example).
- **Caveats:** (1) Calibrations are illustrative, not country calibrations. (2) DSF CI coefficients, thresholds and shock sizes need checking against the current Guidance Note and template. (3) Energy savings enter GIMF-lite as a value-added gain, which drives about two-thirds of the long-run GDP gain in the Kenya example. (4) The DSF demo baseline uses placeholder ratios scaled to NMIM GDP. (5) GIMF-lite is one country and one good; no oil block, financial accelerator, zero lower bound or carbon pricing.
