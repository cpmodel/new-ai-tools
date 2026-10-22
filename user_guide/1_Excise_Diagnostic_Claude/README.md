# Excise diagnostic for Indonesia - with Claude (deep research)

Claude researches Indonesia's excise laws and writes the diagnostic: a Word report and an Excel workbook.

## What you need (we assume you have this)

- The new-ai-tools repo is downloaded on your laptop and up to date.
- You have a Claude account with Projects and Research (deep research) switched on.

> **Deep research uses a lot of credits.** Go one prompt at a time. Read each answer before you send the next. If an answer goes wrong, stop and fix it before going on. Never ask for the whole report in one go.

## Steps

1. **Make a Project.** In Claude: Projects > New project. Call it "Excise - Indonesia".
2. **Add two files.** excise_diagnostic_process.md (in the repo: ExciseDiagnostic/process_files/) and INDONESIA_INPUTS.md (in this folder).
3. **Start a chat.** New chat inside the Project. Switch on Research.
4. **Paste Prompt 1.** Read the answer. Is every rate quoted, with its date and a link? Then paste Prompt 2, and so on to Prompt 6.
5. **Save the results.** Put the Word report and the Excel file in ExciseDiagnostic/country_results/Indonesia/.

## Prompts (copy and paste, one at a time)

**Prompt 1 - the laws**

```text
Use the excise diagnostic process file and INDONESIA_INPUTS.md. Run stage A1 for Indonesia, 2020-01-01 to 2026-09-30: find the excise laws and regulations and list every rate as printed, with its article, date and a link. If you cannot get a source, give me a list of links to download. Stop after A1.
```

**Prompt 2 - the schedule**

```text
Run stage A2. Stop after A2.
```

**Prompt 3 - the numbers and workbook**

```text
Run stage A3. No comparator. Stop after A3.
```

**Prompt 4 - the findings**

```text
Run stage A4. Stop after A4.
```

**Prompt 5 - the review**

```text
Run stage A5. Stop after A5.
```

**Prompt 6 - the final files**

```text
Run stage A6. Give me the report as a Word file and the workbook as an Excel file.
```

## What good looks like

| | |
|---|---|
| Rates | quoted as printed, with date |
| Sources | a working link for each |
| Gaps | marked data_gap, never guessed |
| Bills | pending ones listed, not applied |

## Subsidy diagnostic instead?

Use the subsidy prompt at the end of PROMPTS.md in place of Prompt 1. There is no subsidy process file yet, so treat the result as a first draft.

**Subsidy - Prompt 1**

```text
Using the structure and quality rules of the excise diagnostic process file, start a fossil fuel SUBSIDY diagnostic for Indonesia at 2026-09-30. Step 1 only: list each fuel and electricity subsidy and compensation payment, its legal basis, and the amounts budgeted and spent for the last five years (IDR and % of GDP), with a link for every source. Mark gaps data_gap; never guess. Stop after step 1 and suggest the next steps.
```
