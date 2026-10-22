# Prompts - Add tax indicators to the dashboard

Copy and paste one at a time. Read each answer first.

## Prompt 1 - plan

```text
Read TaxMacroFiscalDashboard/build/README.md. I want four new indicators in the Fiscal block, indented under tax_revenue_pct_gdp, from the IMF WoRLD file already in raw_data (world-imf2026.xlsx, already % of GDP): pit_revenue_pct_gdp = TaxIncI, cit_revenue_pct_gdp = TaxIncC, vat_revenue_pct_gdp = TaxSalG, excise_revenue_pct_gdp = TaxSalExc. Labels start "of which:". Tell me your plan in a short list. Do not edit yet.
```

## Prompt 2 - do it

```text
Go ahead: change the Python script, the R script and config/indicators.csv as planned. Then rebuild (build_global_indicators.py, then build_dashboard.py) and show me Indonesia's five tax rows.
```

## Prompt 3 - save

```text
Save this as version 1.4 following VersionNotes.md. Then commit and push.
```
