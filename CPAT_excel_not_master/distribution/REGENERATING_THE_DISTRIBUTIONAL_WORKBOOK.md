# Regenerating the standalone Distribution workbook

Concrete, step-by-step instructions to rebuild
`CPAT_Distribution_Standalone_Egypt_vN.N.xlsx` + `RebuildDistributionModule.bas`
from source, from a clean checkout, end to end. Follow this whenever a
formula, the Tests sheet, or anything else in `build_workbook.py` changes.

## 0. Prerequisites

- Python 3.11+ (tested on 3.11.15)
- `pip install openpyxl pyxlsb` (only `pyxlsb` is needed if you're re-running
  `extract_egypt.py`; `build_workbook.py` itself only needs `openpyxl`)
- A real copy of Excel (desktop, Windows or Mac) to run the VBA macro and
  save the populated workbook — there is no way to populate the LAMBDA
  names/formulas or produce the final cached-value workbook without
  actually running the macro in real Excel. This cannot be done headlessly
  as part of the Python build.
- Working directory: `cpat_excel/Distribution/standalone/`

## 1. (Only if the underlying Egypt data changed) Re-extract `egypt_data.pkl`

Skip this step if you're only changing formulas, labels, or the Tests
sheet — `egypt_data.pkl` is already checked into the repo and
`build_workbook.py` reads it directly.

Only re-run this if the source `.xlsb` or the `data_tabular` tables changed:

```bash
python3 extract_egypt.py
```

This reads `cpat_excel/original/CPAT 1.0pre_456_NoPropData.xlsb` and
`cpat_excel/Distribution/data_tabular/CPAT_DistributionalData.xlsx`, and
overwrites `egypt_data.pkl`. Check the printed summary (row counts, price
changes, elasticity-adjustment-factor count) looks sane before continuing.

## 2. Make your source changes

Edit `build_workbook.py` (sheet layout, the `LAMBDAS` dict, the `Tests`
sheet's `TEST_CASES`, etc.) and/or `reference_calc.py` (the independent
Python re-implementation used both to validate formula design and to
compute the `Tests` sheet's expected values).

If you change a LAMBDA's formula, also update the matching function in
`reference_calc.py` and re-derive its expected value by hand (or by running
`reference_calc.py` directly) — the two are meant to be independent checks
of each other, so don't just copy one into the other.

## 3. Run the build

```bash
python3 build_workbook.py
```

Read the printed output top to bottom; it is the whole validation story for
this step:

```
Data tabs written: [...]
Distribution_Inputs written, last row ...
Section B formula upgrades applied.
39 LAMBDA functions defined.
Checked 39 LAMBDA formulas for known array-broadcasting bug patterns: none found.
C sections (II-VII) written up to row ...
Distribution_Outputs written, last row ...
Tests sheet written: 20 test cases, last row ...
Captured NNNN formula cells and 83 named ranges for the VBA rebuild.
Checked 83 names against 15 table names: no collisions.
Wrote .../RebuildDistributionModule.bas
Saved .../CPAT_Distribution_Standalone_Egypt.xlsx
```

If the script raises `AssertionError` instead of printing all of the above,
**do not work around it** — it means one of the build-time guards caught a
real problem (a name/table collision, a `MIN`/`MAX`/`INDEX` array-broadcast
pattern, a VBA string round-trip failure). Fix the underlying formula; see
`LESSONS_LEARNED.md` for what each of these guards is protecting against.

This writes two files into the current directory, both with the **bare**
(unversioned) name:
- `CPAT_Distribution_Standalone_Egypt.xlsx` — all formula cells and defined
  names intentionally blank, so it opens with no repair prompt.
- `RebuildDistributionModule.bas` — the VBA macro that populates everything.

## 4. Version the output

This repo's convention: the bare-named `.xlsx` is a *build artifact*, never
committed directly. Decide the next version number (one more than the
highest `_vN.N` already at the top level — check with `ls`), then:

```bash
mv CPAT_Distribution_Standalone_Egypt.xlsx CPAT_Distribution_Standalone_Egypt_v0.6.xlsx
```

`RebuildDistributionModule.bas` keeps its bare name at the top level (it
always overwrites the previous top-level one directly — only the `.xlsx`
gets a version suffix). If you're keeping the *previous* top-level version
pair around for reference, move it into `Old/` first:

```bash
mv CPAT_Distribution_Standalone_Egypt_v0.5.xlsx Old/
```

(`RebuildDistributionModule.bas`'s previous content is already preserved in
git history, so it typically doesn't need a manual `Old/` copy unless you
specifically want one alongside an archived `.xlsx`.)

## 5. Run the macro in real Excel

1. Open the new versioned `.xlsx` in desktop Excel.
2. `Alt+F11` to open the VBA editor.
3. `Insert > Module`, then either paste in `RebuildDistributionModule.bas`'s
   contents, or `File > Import File...` and pick the `.bas` file directly
   (import keeps the `Attribute VB_Name` line meaningful; if you paste
   manually instead, delete that first line — it only does anything via a
   real file import).
4. Run `RebuildDistributionModule` (F5, or Macros list). It defines all the
   named LAMBDA functions and plain named ranges, writes every formula cell
   in `Distribution_Inputs`, `Distribution_Outputs` and `Tests`, then forces
   a full recalculation (`Application.CalculateFullRebuild`).
5. Wait for the completion `MsgBox` (names count + formula count) — if
   Excel doesn't show it, something stalled; check the Immediate window /
   VBA editor for a runtime error before continuing.

## 6. Verify

1. Open the **Tests** sheet. Read cell `C5`.
   - `"ALL 20 TESTS PASS"` → every LAMBDA, called with the full
     `DECILE_ARRAY`, agrees with the independent Python model
     (`reference_calc.py`) to the stated tolerance, and the 3 live Gini
     cross-checks agree too. Done, move to step 7.
   - Anything else (`"k TEST(S) FAILED -- see below"`) → scroll down to the
     failing block(s). Each block shows: the live Actual row, the Expected
     row (from Python), the Max Abs Diff, and which Status cell is `FAIL`.
     Fix the Excel formula or the Python reference (whichever is wrong),
     go back to step 2.
2. Spot-check the Name Manager: should show 83 names (39 LAMBDA + plain
   named ranges), none of them accidentally missing (a name silently
   failing to register — e.g. from a case-insensitive collision with a
   Table name — is exactly the bug class the build-time check in step 3 is
   supposed to prevent; the Tests sheet is the final real-Excel backstop
   for it).
3. Spot check a couple of cells against expectation yourself if you touched
   anything the Tests sheet doesn't directly cover (new sheet labels,
   formatting, a brand-new section).

## 7. Re-save to bake in cached values

Once the Tests sheet shows all-pass, save the workbook (`Ctrl+S`, keep
`.xlsx` format — do not save as `.xlsm`, the macro doesn't need to persist
in the file, see note below). This replaces the blank-formula file on disk
with one that has live formulas *and* cached values, matching the pattern
of the last confirmed-working version (`CPAT_Distribution_Standalone_Egypt_v0.4.xlsx`,
committed directly from a real Excel save, not from this Python script's
raw output).

> Why not `.xlsm`: the whole point of the macro is to populate the
> workbook once; nothing about using the workbook afterward needs VBA to
> still be present. Saving as `.xlsx` (which silently drops the VBA
> project) keeps the deliverable a plain, macro-free spreadsheet. If you
> specifically want a reusable macro-enabled copy for convenience, that's a
> separate, clearly-labeled `.xlsm` artifact (see `Old/*_v0.1.xlsm`,
> `Old/*_v0.3.xlsm`) — don't let it become the main deliverable.

## 8. Commit

Follow the repo's established pattern — one commit per artifact, so each
file's history stays easy to read:

```bash
git add build_workbook.py reference_calc.py   # whatever source files changed
git commit -m "..."

git add CPAT_Distribution_Standalone_Egypt_v0.6.xlsx
git commit -m "Add CPAT_Distribution_Standalone_Egypt_v0.6.xlsx"

git add RebuildDistributionModule.bas
git commit -m "Update RebuildDistributionModule.bas for ..."

# if you archived the previous version:
git add Old/CPAT_Distribution_Standalone_Egypt_v0.5.xlsx
git commit -m "Archive v0.5"
```

Push to the feature branch you're working on; only push to `main` if
explicitly asked to.

## Quick reference: full command sequence

```bash
cd cpat_excel/Distribution/standalone
# python3 extract_egypt.py          # only if source data changed
python3 build_workbook.py
mv CPAT_Distribution_Standalone_Egypt.xlsx CPAT_Distribution_Standalone_Egypt_vX.Y.xlsx
# -> open in real Excel, run the macro, check Tests!C5, save, close
git add -A   # review with `git status` first; commit in the granular steps above instead if preferred
```
