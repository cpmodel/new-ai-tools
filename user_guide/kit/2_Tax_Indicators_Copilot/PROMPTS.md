# Example 2 - Add tax indicators to the dashboard (GitHub Copilot)

Goal: add four tax indicators to the Fiscal block of the Global Fiscal Indicators
dashboard (`TaxMacroFiscalDashboard/`), under "Tax revenue (% GDP)":

| New indicator | Code | WoRLD column |
|---|---|---|
| of which: personal income tax (% GDP) | `pit_revenue_pct_gdp` | `TaxIncI` |
| of which: corporate income tax (% GDP) | `cit_revenue_pct_gdp` | `TaxIncC` |
| of which: VAT / general sales tax (% GDP) | `vat_revenue_pct_gdp` | `TaxSalG` |
| of which: excises (% GDP) | `excise_revenue_pct_gdp` | `TaxSalExc` |

No download is needed: the data are already in `build/raw_data/world-imf2026.xlsx`.

Work in the repo on your laptop, in VS Code (or the Copilot app), with Copilot Chat in
**Agent** mode and model **Claude Sonnet**. Pull first. One prompt per message.

---

## Prompt 1 - Plan only (no edits)

```text
Read AGENTS.md, TaxMacroFiscalDashboard/build/README.md and
TaxMacroFiscalDashboard/VersionNotes.md. I want to add four indicators to the Fiscal
block, indented under tax_revenue_pct_gdp, all from the IMF WoRLD file already in
raw_data (world-imf2026.xlsx, already % of GDP):
pit_revenue_pct_gdp = TaxIncI, "of which: personal income tax (% GDP)"
cit_revenue_pct_gdp = TaxIncC, "of which: corporate income tax (% GDP)"
vat_revenue_pct_gdp = TaxSalG, "of which: VAT / general sales tax (% GDP)"
excise_revenue_pct_gdp = TaxSalExc, "of which: excises (% GDP)"
Tell me which files you will change and how, in a short list. Do not edit anything yet.
```

## Prompt 2 - Make the change

```text
Go ahead. Follow the build README ("To add an indicator"):
1. In scripts/build_global_indicators.py, in the IMF WoRLD section, emit the four
   indicators next to tax_revenue_pct_gdp, same years (2010-2024) and same source file.
2. Make the same change in scripts/build_global_indicators.R so the two routes match.
3. In config/indicators.csv add four rows right after tax_revenue_pct_gdp: copy its
   block columns, area "Revenue capacity", direction "High = better", eval_scale GMB,
   score_flag 1 (as its parent), indent 1, headline 0,
   source_label "IMF WoRLD 2026".
Do not change any other indicator. Never invent data.
```

## Prompt 3 - Build and check

```text
Run the build: python build_global_indicators.py, then python build_dashboard.py
(from TaxMacroFiscalDashboard/build/scripts). Confirm "missing indicators: none".
Then show me the five tax rows for Indonesia (IDN) from output/global_data_long.csv.
```

Expected for Indonesia (2024): tax revenue 10.08, PIT 1.79, CIT 2.65, VAT 3.52,
excises 1.11 (% of GDP). On Windows you can instead double-click
`1_build_database.bat`, then `2_build_tool.bat`.

## Prompt 4 - New version and notes

```text
Ship this as version 1.4 following VersionNotes.md: bump the version, move v1.3 of the
workbook to Archive/, rebuild, and add a v1.4 line to VersionNotes.md saying the four
tax indicators were added. Update the indicator count in the build README. Then show me
the list of changed files.
```

## Prompt 5 - Save to GitHub

```text
Commit these changes with a clear message and push them.
```

(Or double-click `SyncRepo.bat` if your repo has it.)

---

## Check before you push

- Open `Global_fiscal_indicators_v1.4.xlsx`, pick Indonesia: the four rows sit under
  Tax revenue, indented.
- The four parts add up to less than total tax revenue (trade, property and other taxes
  are the rest). For Indonesia: 1.79 + 2.65 + 3.52 + 1.11 = 9.07 < 10.08.
- Nothing else in the workbook changed (`git diff` touches only the files above).
