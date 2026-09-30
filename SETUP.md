# Setup, Requirements, and Process

This file consolidates the environment requirements, build instructions, and run
processes that were previously spread across subfolder `README` files
(`TaxMacroFiscalDashboard/build/README.md`, `ExciseDiagnostic/README.txt`) and their
associated `requirements.txt` / `r_dependencies.R` files.

---

## 1. TaxMacroFiscalDashboard/build — Global Fiscal Indicators pipeline

Self-contained. No network access required to build (beyond the optional source
refresh), no API keys, no hardcoded paths.

```
TaxMacroFiscalDashboard/build/
  config/
    indicators.csv        THE definition: 53 indicators, 9 blocks, one row each
    sources.csv            source metadata (coverage is derived, not stored)
  raw_data/                23 source files the scripts read
  raw_data_not_used/       12 downloaded but superseded or too thin
  raw_data_downloaded/     where download_sources.py puts fresh copies (never raw_data/)
  scripts/
    build_global_indicators.py    preprocessing -> the panel      (Python)
    build_global_indicators.R     identical pipeline              (R)
    build_dashboard.py            the tool: panel -> Excel workbook
    download_sources.py           optional: re-fetch sources that have a direct URL
  0_setup.bat              Windows: install Python packages (run once)
  1_build_database.bat     Windows: run the preprocessing
  2_build_tool.bat         Windows: build the workbook
  output/                  created on first run
  requirements.txt         pip install -r requirements.txt
  r_dependencies.R         Rscript r_dependencies.R
  MANIFEST.csv             every file with size and MD5
```

### Requirements

**Python** (`TaxMacroFiscalDashboard/build/requirements.txt`):

```
pandas>=2.0
numpy>=1.24
openpyxl>=3.1
pyxlsb>=1.0.10
xlsxwriter>=3.1
```

Install with:

```bash
pip install -r TaxMacroFiscalDashboard/build/requirements.txt
```

**R** (optional, `TaxMacroFiscalDashboard/build/r_dependencies.R`) — an identical
preprocessing pipeline for cross-checking the Python output:

```
data.table, readxl, writexl, haven
```

Install with:

```bash
Rscript TaxMacroFiscalDashboard/build/r_dependencies.R
```

### Running it on Windows

From `TaxMacroFiscalDashboard/build/`, double-click in order. Each pauses at the end
so you can read the output.

| File | What it does |
|---|---|
| `0_setup.bat` | Installs the Python packages. Run once. |
| `0b_download_sources.bat` | Optional. Re-fetches the 6 sources with a direct URL into `raw_data_downloaded\`, and prints a checklist for the 13 that need manual download. |
| `1_build_database.bat` | Preprocessing: raw files -> `output\global_data_*.csv` |
| `2_build_tool.bat` | The tool: panel -> `output\Global_fiscal_indicators.xlsx` |

`1_build_database.bat ccdr` restricts to the 98 published-CCDR economies. Step 1 also
writes the CPAT `.xlsx` the R route needs, so you can run R afterwards without a
separate command.

Step 2 reads no raw data, so re-run it on its own after editing `config\indicators.csv`
— to move an indicator between blocks, relabel one, or change its direction.

### Running it anywhere else

Paths default to the folders above, resolved relative to the script. From
`TaxMacroFiscalDashboard/build/scripts/`:

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

#### The R route needs one pre-step

R has no reliable `.xlsb` reader, so run this once before the R script:

```bash
python build_global_indicators.py --emit-cpat-xlsx
```

It writes the three CPAT sheets to `.xlsx` with explicit `c0..cN` headers. The headers
matter: `readxl` silently drops leading all-empty columns, which would shift every
positional index. Doing it by hand in Excel works too — Save As `.xlsx`, named
`imffossilfuelsubsidiesdata_sheets.xlsx`, into `output/`.

### Refreshing the sources

`config/sources.csv` carries a `download_kind` column.

**6 are `direct`** — the six World Bank WDI indicator files, the IRENA appendix PDF,
the UNDP review PDF, and the PSRT archive via the Dataverse API. `download_sources.py`
fetches these.

**13 are `manual`.** The publisher offers only a landing page or an interactive query
builder: the IMF portals render their download links in JavaScript, and IDS, PEFA and
OECD all require you to build a query first. The `download_note` column records the
settings to pick — for IDS, for instance, Country=all, Counterpart-Area=World, the 19
`DT.*` series, 2015-2024. The script prints these as a checklist rather than
pretending to handle them.

```bash
python scripts/download_sources.py --list     # show the plan
python scripts/download_sources.py            # fetch the direct ones
```

Downloads go to `raw_data_downloaded/`, never `raw_data/`. **Diff before swapping any
in.** Publishers revise series, and an unnoticed change in an input is how a rebuild
quietly stops matching published numbers. `MANIFEST.csv` has an MD5 for every file in
`raw_data/` so you can tell what actually changed.

### Preprocessing and tool are separate

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

### Output

| Stage | File | Rows |
|---|---|---|
| 1 | `global_long_all_years.csv` — every source, every year | 105,072 |
| 2 | `global_data_long.csv` — latest observation per country × indicator | 8,057 |
| 3 | `global_data_wide.csv` — one row per country, plus `__year` columns | 217 × 53 |

Python writes `global_*`, R writes `global_*_R`, so you can diff them. Verified
identical in a clean-room run: same row count, zero value mismatches, zero year
mismatches, identical coverage on all 53 indicators.

### Ground rules baked in

- **Nothing is imputed.** Missing stays missing; the dashboard flags it per row and
  counts it per block.
- **Every value carries its observation year.** Vintages span 2013 to 2025.
- **Every value names its source file** in the long format.
- **Three traps are guarded explicitly**, each commented in the code: `INDEX()` on an
  empty cell returns 0 rather than blank; Ember carries a country's whole record
  forward where it has no observations, producing a spurious 0.0pp change; and a WACC
  evidence count of zero is an absence, not an observation.

### Not used, and why

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

---

## 2. ExciseDiagnostic — run process

File set current as of September 2026. There is no software build step here; the
"process" is a specification for running the diagnostic with an AI assistant (e.g.
Claude), one stage at a time, with a human reviewer checking output between stages.

### File set

- **`Excise_Diagnostic_WorkingPaper_v0_9.docx`** — Working paper (current). Section 5:
  whole vs five-part runs.
- **`Excise_Diagnostic_CrossCountry.xlsx`** — Cross-country workbook (Sierra Leone,
  Uganda, Liberia), re-tagged to the five-part map; `Part_Map` and `Part_Summary`
  sheets.
- **`process_files/`** — give the master spec to the assistant at the start of every
  run:
  - `excise_diagnostic_process.md` — master process specification (v1.1)
  - `excise_part1_fuel.md` — Part 1: fuel (E1–E6)
  - `excise_part2_vehicles.md` — Part 2: vehicles (E8–E9)
  - `excise_part3_health.md` — Part 3: health (T1–T3, A1–A5, S1–S4, I3)
  - `excise_part4_other.md` — Part 4: other (E7, G1, D1–D2, F1, I1, I2, I4, L1)
  - `excise_part5_overarching.md` — Part 5: overarching — framework and aggregation
    (run last)
- **`archive/`** — prior working-paper versions (v0.7, v0.8) and an earlier copy of
  `process_files/` kept for historical reference

### Run modes

- **Whole run**: master process file only (stages A1–A6).
- **Sectoral run**: master file plus one part file per run; Parts 1–4 in any order,
  Part 5 last.

### Master spec summary (`excise_diagnostic_process.md`, v1.1)

1. Choose a run mode (WHOLE or SECTORAL).
2. Give the master process file to the assistant as context at the start of every run
   (Claude Project upload, or pasted as the first message). In a sectoral run, also
   give the part file for the part being run.
3. Supply the inputs defined in the spec.
4. If the assistant cannot fetch a source, it runs the data step: it stops and asks
   for a set of downloads, as a list of links plus a `.bat` file.
5. Run the stages in order, one message per stage, using the spec's prompt templates
   (or the part file). A human reviewer checks each stage's output before the next
   stage starts.
6. Every run produces a report and a workbook (in the cross-country structure).

Rule of precedence: the master process file governs over any conflicting instruction,
unless a human reviewer overrides it in writing. Part files add detail; they do not
override the master file.
