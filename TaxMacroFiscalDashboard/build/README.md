# Global Fiscal Indicators — raw data, config and build scripts

Self-contained. No network access, no API keys, no hardcoded paths.

```
config/
  indicators.csv        THE definition: 53 indicators, 9 blocks, one row each
  sources.csv           source metadata (coverage is derived, not stored)
raw_data/               23 source files the scripts read
raw_data_not_used/      12 downloaded but superseded or too thin
scripts/
  build_global_indicators.py    preprocessing -> the panel      (Python)
  build_global_indicators.R     identical pipeline              (R)
  build_dashboard.py            the tool: panel -> Excel workbook
0_setup.bat             Windows: install Python packages (run once)
1_build_database.bat    Windows: run the preprocessing
2_build_tool.bat        Windows: build the workbook
output/                 created on first run
requirements.txt        pip install -r requirements.txt
r_dependencies.R        Rscript r_dependencies.R
MANIFEST.csv            every file with size and MD5
```

## Running it on Windows

Double-click, in order. Each pauses at the end so you can read the output.

| File | What it does |
|---|---|
| `0_setup.bat` | Installs the Python packages. Run once. |
| `1_build_database.bat` | Preprocessing: raw files -> `output\global_data_*.csv` |
| `2_build_tool.bat` | The tool: panel -> `output\Global_fiscal_indicators.xlsx` |

`1_build_database.bat ccdr` restricts to the 98 published-CCDR economies.
Step 1 also writes the CPAT `.xlsx` the R route needs, so you can run R afterwards
without a separate command.

Step 2 reads no raw data, so re-run it on its own after editing
`config\indicators.csv` — to move an indicator between blocks, relabel one, or change
its direction.

## Running it anywhere else

Paths default to the folders above, resolved relative to the script. From `scripts/`:

```bash
python build_global_indicators.py        # -> output/global_data_{long_all_years,long,wide}.csv
python build_dashboard.py                # -> output/Global_fiscal_indicators.xlsx
```

Override with flags or environment variables if your layout differs:

```bash
python build_global_indicators.py --raw /path/raw --out /path/out --config /path/config
RAW_DIR=... OUT_DIR=... CONFIG_DIR=... Rscript build_global_indicators.R
```

`--ccdr-only` (Python) or `CCDR_ONLY=1` (R) restricts to the 98 published-CCDR
economies instead of all 217.

### The R route needs one pre-step

R has no reliable `.xlsb` reader, so run this once before the R script:

```bash
python build_global_indicators.py --emit-cpat-xlsx
```

It writes the three CPAT sheets to `.xlsx` with explicit `c0..cN` headers. The headers
matter: `readxl` silently drops leading all-empty columns, which would shift every
positional index. Doing it by hand in Excel works too — Save As `.xlsx`, named
`imffossilfuelsubsidiesdata_sheets.xlsx`, into `output/`.

## Preprocessing and tool are separate

They share exactly one thing — `config/indicators.csv` — and communicate only through
`output/global_data_long.csv`.

**To add an indicator:** add the extraction to the pipeline, then add a row to
`config/indicators.csv`. Nothing else. Block membership, label, direction, evaluation
scale, scoring flag, indentation and source label all come from that row, in both the
panel and the workbook.

**The contract is enforced.** If the config names an indicator the pipeline never
emits, the run stops with a `PIPELINE/CONFIG MISMATCH` error naming it. It does not
warn and continue — that failure mode previously let indicators go missing silently.

**Source coverage is derived**, counted from the panel at build time rather than typed
into the Sources sheet, so it cannot drift.

## Output

| Stage | File | Rows |
|---|---|---|
| 1 | `global_long_all_years.csv` — every source, every year | 105,072 |
| 2 | `global_data_long.csv` — latest observation per country × indicator | 8,057 |
| 3 | `global_data_wide.csv` — one row per country, plus `__year` columns | 217 × 53 |

Python writes `global_*`, R writes `global_*_R`, so you can diff them. Verified
identical in a clean-room run: same row count, zero value mismatches, zero year
mismatches, identical coverage on all 53 indicators.

## Ground rules baked in

- **Nothing is imputed.** Missing stays missing; the dashboard flags it per row and
  counts it per block.
- **Every value carries its observation year.** Vintages span 2013 to 2025.
- **Every value names its source file** in the long format.
- **Three traps are guarded explicitly**, each commented in the code: `INDEX()` on an
  empty cell returns 0 rather than blank; Ember carries a country's whole record
  forward where it has no observations, producing a spurious 0.0pp change; and a WACC
  evidence count of zero is an absence, not an observation.

## Not used, and why

| File | Reason |
|---|---|
| `dataset_*ICSD*.csv` | IMF capital stock — ends 2019, never made the indicator set |
| `CRDF-PP_2024.xlsx` | OECD provider perspective — recipient perspective used instead |
| `CRDF-RP_all_years_2000-2024.xlsx` | 25-year panel — held for trend work |
| `UNUWIDERGRD_2025.xlsx` | WoRLD's excise and resource columns covered it |
| `assessments_1787749635.csv` | PEFA Climate framework — 20 assessments, too thin |
| `P_Data_Extract_..._1_.xlsx` | First IDS pull — one series, two years; superseded |
| `rethinking_*.xlsx` (6) | 15 countries plus 3 Indian states. The cost-recovery workbook holds a genuine quasi-fiscal deficit series (% GDP, decomposed) but only 2010–2012 |

Three files in `raw_data/` contribute transcribed tables rather than being parsed: the
IRENA WACC appendix and the UNDP review are PDFs, and `CofCObservatoryData.xlsx`
supplies a 9-country list. Each transcription is inline in the scripts, commented and
attributed.
