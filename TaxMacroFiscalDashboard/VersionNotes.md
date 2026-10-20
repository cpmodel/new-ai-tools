# Version notes: Tax and Macro-Fiscal Dashboard (Global Fiscal Indicators workbook)

The shipped workbook is `Global_fiscal_indicators_v<version>.xlsx` in this folder, built by
`build/scripts/build_dashboard.py` (the version is set there). Earlier versions are in `Archive/`.
To ship a new version: bump `VERSION`, move the previous workbook to `Archive/`, rebuild, and add a
line here.

| Version | Date | File | What changed |
|---|---|---|---|
| 1.0 | 2026-09-30 | `Archive/Tax+Macro-Fiscal-Energy_indicators_v1.0.xlsx` | First shipped workbook: 217 countries, 53 indicators, 9 blocks (Fiscal, PIMA, Finance 1-2, Context 1-3, Appendix A-B) |
| 1.1 | 2026-10-20 | `Archive/Global_fiscal_indicators_v1.1.xlsx` | 64 indicators, 11 blocks: Appendix C (macro-model inputs: WEO revenue, primary balance, investment, current account, 5-year growth) and Appendix D (integrated scenario from MacroFiscalModels, model output); version shown on the README sheet; the file is versioned and committed |
| 1.2 | 2026-10-20 | `Archive/Global_fiscal_indicators_v1.2.xlsx` | 65 indicators: adds interest payments (net), general govt (% GDP) to the Fiscal block, derived from WEO as primary balance less overall balance (the numerator of interest-to-revenue); 189 economies, negative where interest received exceeds interest paid (e.g. Kuwait, Norway) |
| 1.3 | 2026-10-20 | `Global_fiscal_indicators_v1.3.xlsx`. 80 indicators, 12 blocks: Appendix C gains the macro-model and debt-sustainability inputs (public investment from IMF ICSD; external PPG debt and its share of public debt from IDS; concessional share of external debt, CPIA, remittances, reserves and Findex account ownership from the Climate Fiscal Risk Dashboard v1.5); new Appendix E, fiscal risk (sovereign rating, debt maturing within 12 months, total external debt, short-term external debt over reserves, climate in PIM, budget tagging, national climate fund). Appendix D rebuilt from the re-run integrated run. Existing values unchanged. v1.2 in `Archive/` |
