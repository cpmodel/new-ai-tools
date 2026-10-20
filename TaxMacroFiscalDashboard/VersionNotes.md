# Version notes: Tax and Macro-Fiscal Dashboard (Global Fiscal Indicators workbook)

The shipped workbook is `Global_fiscal_indicators_v<version>.xlsx` in this folder, built by
`build/scripts/build_dashboard.py` (the version is set there). Earlier versions are in `Archive/`.
To ship a new version: bump `VERSION`, move the previous workbook to `Archive/`, rebuild, and add a
line here.

| Version | Date | File | What changed |
|---|---|---|---|
| 1.0 | 2026-09-30 | `Archive/Tax+Macro-Fiscal-Energy_indicators_v1.0.xlsx` | First shipped workbook: 217 countries, 53 indicators, 9 blocks (Fiscal, PIMA, Finance 1-2, Context 1-3, Appendix A-B) |
| 1.1 | 2026-10-20 | `Global_fiscal_indicators_v1.1.xlsx` | 64 indicators, 11 blocks: Appendix C (macro-model inputs: WEO revenue, primary balance, investment, current account, 5-year growth) and Appendix D (integrated scenario from MacroFiscalModels, model output); version shown on the README sheet; the file is versioned and committed |
