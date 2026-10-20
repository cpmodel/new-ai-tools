# Macro-fiscal country tool

The current tool, `../MacroFiscal_CountryTool_v<version>.xlsx` (in `MacroFiscalModels/`), is a standalone Excel workbook: pick a country on the
**Country** sheet and it shows the whole integrated chain for it, with charts. No Python is needed
to use it.

| Section | Source |
|---|---|
| The country: income group, model used and why, run status, how many settings are defaults | `calibration.py` |
| 1. How much: incremental budget need (capital and spending), private investment, energy savings | NMIM v0.2 |
| 2. How it can be paid: feasible carbon tax and feebate, binding limit, F1-F3, revenue, financing gap | PriceFeasibility v0.1 |
| 3. What it does: GDP vs today's path, public debt 2024 / 2030 / 2050, fossil rents | GIMF-lite / DIGNAD-lite / DIGNAR-lite |
| 4. Is debt sustainable: risk rating before and with the programme | LIC-DSF-lite (low and lower-middle income) |
| Settings: each calibration value and whether it is dashboard data or a default | `calibration.py` |
| Charts, two columns: the result (programme and financing; GDP and debt, programme and baseline; PV of public debt against the DSF threshold) and the difference from the baseline (financing gap; GDP, pp; debt, pp; PV of debt, scenario less DSF baseline) | annual paths, incl. a no-programme baseline from the same macro model |

Results are pre-computed; changing an assumption needs a re-run:

```bash
cd MacroFiscalModels
python3 integrated.py --all          # results/integrated_{results,paths,settings}.csv (about 3 minutes)
python3 build_country_tool.py        # -> MacroFiscal_CountryTool_v<VERSION>.xlsx (in MacroFiscalModels/)
```

Versions: `VersionNotes.md`; the current workbook sits in `MacroFiscalModels/`, superseded ones in `Archive/` here. Calibrations are illustrative;
see `../Docs/UserGuide_4_Integrated` for the chain and its caveats.
