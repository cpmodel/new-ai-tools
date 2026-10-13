# Multiple scenarios: definitions, batch run, stored results, comparison

How CPAT-AI-Mitigation-MVP runs many scenarios. Each run uses two calculated scenarios: the baseline and one live policy scenario. Results are stored as values and compared afterwards. This follows legacy CPAT's approach, where a macro copies each scenario column into the column used for calculation and stores the outputs.

## Where things are

| Place | What it holds |
|---|---|
| MTInputs column J | Scenario 1, the **baseline** (Mitigation group 1). Row 4: "Baseline (stored once)". |
| MTInputs column K | Scenario 2, the **live policy scenario** (Mitigation group 2). Row 4: "Live (batch target)". The batch run overwrites it with each definition and restores it at the end. |
| MTInputs columns L onwards | **Scenario definitions**, one full MTInputs column each: row 4 Run? (Yes/No), row 5 number (previous + 1), row 6 name, rows 8-415 the inputs. Red cells differ from the template. To add one, copy the last definition one column to the right and edit it. |
| StoredResults | Results as **values only**: one block of the 25 MTOutputs rows per scenario, under its ID. ID 1 is the baseline; a definition keeps its MTInputs number. Columns: ID, name, code stem, output code, description, unit, time stored, then 2022-2040. |
| ScenarioCompare | Key results for **one year** (C4) of the baseline and up to 8 stored scenarios (IDs in row 6): levels, differences from the baseline, and % differences. Its formulas read StoredResults only. |
| VBA module CPATScenarios | The macros below (Alt+F8). The workbook is an .xlsm; the source is also in `CPATScenarios_v0_2.bas`. |
| MacroCheck | Report of the first-run test macro CheckBatchRun: one row per check, Result in row 6. |

## Shipped definitions

| No. | Name | Inputs changed from the template |
|---|---|---|
| 3 | Carbon price $20/tCO2 from 2027 | The live scenario's settings |
| 4 | Package A: carbon tax to $50 by 2030 + feebates (transport, industry) | Carbon tax $10 (2027) to $50 (2030). Feebates 2027-2030: transport $10 → $50, industry $5 → $25. Feebate coverage adds road, mining & chemicals, and iron & steel |
| 5 | Package B: ETS on power and industry + carbon tax $25 elsewhere | Carbon tax $25 from 2027. New ETS from 2027: cap −5% (2027) to −20% (2035) of baseline covered emissions, constant afterwards. Template ETS coverage (power, mining & chemicals, iron & steel, non-ferrous metals, cement), default benchmarks and volatility |

Paths continue after their target year as set in the model: the carbon price follows its MTInputs switch, and feebates continue linearly.

## Macros

- **RunAllScenarios**
  1. Stores the baseline under ID 1, but only if no baseline is stored yet.
  2. For every definition with Run? = Yes:
     1. copies its rows 8-415 into column K, except the formula rows 10 and 12, and copies its name;
     2. clears the ETS override row of group 2 and recalculates;
     3. runs the ETS goal seek if the definition applies a new ETS and does not set "Override ETS price" = Yes;
     4. stores the MTOutputs block of scenario 2 as values, under the definition's number and name. An existing block with that ID is overwritten.
  3. Restores column K and the override row, and recalculates.
- **StoreBaseline:** replaces the stored baseline with the current scenario 1 results, after a confirmation.
- **StoreLiveScenario:** stores the live scenario 2 as it stands, under an ID you choose (the default is the next free ID from 100).
- **SolveETSLive:** ETS goal seek for the live scenario only. It fills the override row `ets.ovr` and sets the override switch in column K.
- **ClearStoredScenarios:** deletes every stored scenario except the baseline.
- **CheckBatchRun:** the first-run test (see below). Reruns the batch, so the stored scenarios are recalculated by the application that runs it.

The ETS goal seek is legacy's damped log-space iteration, as in `ets_goalseek_v0_2.py`:
- It starts from the override row, or from the fast estimate if the row is empty.
- Each step is p × (`ets.next` / p) ^ alpha, mixed 0.3 with the previous iterate, with 5% smoothing.
- Alpha adapts: it halves when the error grows (minimum 0.0625) and rises ×1.2 when the error more than halves (maximum 1).
- It stops when the worst |covered / cap − 1| is below 0.5%, or after 20 iterations.

## First run in Excel

The VBA project and the stored results were produced without Excel. They are tested in LibreOffice: the module loads, the macros reproduce the Python emulation, and CheckBatchRun passes. To confirm the same in Excel:

1. If the file was downloaded, unblock it first (file Properties > Unblock), or Excel opens it with macros disabled.
2. Open it and choose Enable Content.
3. Run the macro **CheckBatchRun** (Alt+F8). It takes a minute or two. It:
   - checks the 12 lookups the macros rely on (sheets, rows by code, scenario columns, MTOutputs blocks);
   - compares Excel's own recalculation of the baseline with the stored baseline. Excel evaluates the LAMBDA column natively; LibreOffice needed it expanded;
   - reruns RunAllScenarios, including the ETS goal seek, and compares every stored value with the value stored before the run.
4. Sheet MacroCheck shows **Result PASS** when no value differs by more than 1e-6 (relative). CHECK lists how many values differ; FAIL names the lookup that was not found.
5. If Excel reports that it cannot load the VBA project, or Alt+F8 lists no macros:
   1. remove the project if one is listed;
   2. import `CPATScenarios_v0_2.bas` (Alt+F11 > File > Import);
   3. save as .xlsm and run CheckBatchRun again.

## Comparing scenarios

- Choose the year in ScenarioCompare C4 and the scenario IDs in row 6.
- Each value is a SUMIFS on StoredResults by ID and code stem, in the column of the chosen year. A blank cell means the ID or stem is not stored.
- The change rows that are already relative to scenario 1 (fuel use %, CO2 change, revenue change) are left out; the difference blocks take their place.

## Keeping results

- StoredResults holds values only, so results from other files can be pasted in as values under new IDs.
- The baseline does not change unless StoreBaseline is run.
- ScenarioCompare finds pasted rows as long as the column layout is kept (ID in A, stem in C, years from H).

## Limits

- The batch run calculates one policy scenario at a time. Mitigation scenario groups 3+ (copied groups) are still possible, but the macros do not use them.
- MTOutputs (25 indicators) defines what is stored. To store more, add rows to the MTOutputs blocks first.
- The VBA project is written by `vba_project_v0_1.py` without Excel. It is tested in LibreOffice: the module loads and the macro reproduces the Python emulation. CheckBatchRun is the test for its first run in Excel. If Excel does not accept the project, import `CPATScenarios_v0_2.bas` (Alt+F11 > File > Import) and save as .xlsm.
