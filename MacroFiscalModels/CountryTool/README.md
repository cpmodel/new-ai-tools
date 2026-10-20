# Macro-fiscal country tool

`MacroFiscal_CountryTool_v<version>.xlsx` is a standalone Excel workbook: pick a country on the
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
| Charts: programme and financing; GDP; debt; PV of public debt against the DSF threshold | annual paths |

Results are pre-computed; changing an assumption needs a re-run:

```bash
cd MacroFiscalModels
python3 integrated.py --all          # results/integrated_{results,paths,settings}.csv (about 10 minutes)
python3 build_country_tool.py        # -> CountryTool/MacroFiscal_CountryTool_v<VERSION>.xlsx
```

Versions: `VersionNotes.md`; superseded workbooks go to `Archive/`. Calibrations are illustrative;
see `../Docs/UserGuide_4_Integrated` for the chain and its caveats.
