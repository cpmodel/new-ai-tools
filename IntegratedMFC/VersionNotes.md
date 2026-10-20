# Version notes: Integrated MF-C tool

The shipped workbook is `Integrated_MFC_Tool_v<version>.xlsx` in this folder, built by
`build/build_mfc_tool.py`; the version is set in `build/config/tool.json`. To ship a new version:
bump `version`, move the previous workbook to `Archive/`, rebuild, and add a line here.

| Version | Date | What changed |
|---|---|---|
| 0.1 | 2026-10-20 | First version: the Global Fiscal Indicators dashboard (v1.2 panel, 65 indicators, 11 blocks) plus the investment needs module, two blocks after Context 3: Investment 1 (v1.1 workbook, 2030: total and incremental need, incremental need by 7 sectors, incremental cost under lenses a/b/c, cost per tonne, cost-of-capital penalty, PIMA-adjusted need; 194 economies) and Investment 2 (NMIM v0.2: net zero investment by who pays P1-P5, direct budget need, private share, incremental investment 2030 and 2050, PV of incremental cost; 186 economies). 217 countries, 91 indicators, 13 blocks |
