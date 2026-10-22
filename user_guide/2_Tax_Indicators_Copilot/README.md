# Add tax indicators to the dashboard - with GitHub Copilot

Copilot adds four rows under Tax revenue in the Global Fiscal Indicators dashboard: personal income tax, corporate income tax, VAT and excises (% of GDP). The data are already in the repo.

## What you need (we assume you have this)

- The new-ai-tools repo is downloaded on your laptop and up to date.
- GitHub Copilot is installed and signed in (if not: Presentations/GitHub_Copilot_AI_User_Guide.pptx).
- Python is set up for the dashboard (TaxMacroFiscalDashboard/build/0_setup.bat, run once).

## Steps

1. **Open the repo.** Open the new-ai-tools folder in VS Code. Click Sync to get the latest.
2. **Open Copilot Chat.** Pick Agent mode and the model Claude Sonnet.
3. **Paste Prompt 1.** Copilot reads the files and tells you its plan. Nothing changes yet.
4. **Paste Prompt 2.** Copilot makes the change and rebuilds. Click Keep. Check the numbers (last page).
5. **Paste Prompt 3.** Copilot saves version 1.4 and pushes it to GitHub. Done.

## Prompts (copy and paste, one at a time)

**Prompt 1 - plan**

```text
Read TaxMacroFiscalDashboard/build/README.md. I want four new indicators in the Fiscal block, indented under tax_revenue_pct_gdp, from the IMF WoRLD file already in raw_data (world-imf2026.xlsx, already % of GDP): pit_revenue_pct_gdp = TaxIncI, cit_revenue_pct_gdp = TaxIncC, vat_revenue_pct_gdp = TaxSalG, excise_revenue_pct_gdp = TaxSalExc. Labels start "of which:". Tell me your plan in a short list. Do not edit yet.
```

**Prompt 2 - do it**

```text
Go ahead: change the Python script, the R script and config/indicators.csv as planned. Then rebuild (build_global_indicators.py, then build_dashboard.py) and show me Indonesia's five tax rows.
```

**Prompt 3 - save**

```text
Save this as version 1.4 following VersionNotes.md. Then commit and push.
```

## Check: Indonesia, % of GDP (2024)

| | |
|---|---|
| Tax revenue | 10.08 |
| Personal income tax | 1.79 |
| Corporate income tax | 2.65 |
| VAT | 3.52 |
| Excises | 1.11 |
