# Example 1 - Excise diagnostic for Indonesia (Claude)

Copy-paste prompts. Send ONE prompt per message. Read and check each answer before
sending the next one. Replace nothing: the prompts are already filled in for Indonesia.

Before you start, your Claude Project must hold these files (Project knowledge):

- `excise_diagnostic_process.md` (the master process file - always)
- `Excise_Diagnostic_CrossCountry.xlsx` (shows Claude the workbook layout)
- `INDONESIA_INPUTS.md` (this kit)

Turn on **Research** (deep research) and **Web search** in the chat box.

---

## Prompt 0 - Check the set-up (optional, 1 minute)

```text
Read the project files. In five bullet points, tell me: what the excise diagnostic is,
which run mode you will use (WHOLE), the six stages A1-A6, what inputs you still need
for Indonesia, and what you will do if you cannot fetch a source. Do not start the run.
```

## Prompt A1 - Legislative history

```text
Using the excise diagnostic process specification (master) and the inputs in
INDONESIA_INPUTS.md, run stage A1 for Indonesia, from 2020-01-01 to 2026-09-30.
Research the Indonesian legal texts yourself first (laws, government regulations,
Ministry of Finance regulations - PMK - and budget documents).
Output the legislative history instrument by instrument, quoting every rate as printed
with its provision and effective date. Mark unreadable tables "not legible". List
customs-tariff changes separately. List pending bills and announced-but-not-in-force
excises (for example on sugar-sweetened beverages and plastics) - record them, do not
apply them. Give the link for every source.
If you cannot fetch a source, run the data step: give me the links list and the .bat file.
```

## Prompt A2 - Current schedule

```text
Run stage A2. Build the current schedule at 2026-09-30 as schedule_line records
(section 5), tagged by part, then as a Level 3 table with a Ref. column. Add the rate
trajectory table.
```

## Prompt A3 - Parameters, standard schedule, workbook

```text
Run stage A3. Fill the parameter sheet (section 5) from the data provided and flag gaps.
Map every line to standard schedule v1.0 and compute the three measures, showing each
calculation. Add collections. Populate the workbook (section 11) and recalculate it.
No comparator; omit comparator columns and ratios.
```

## Prompt A4 - Findings (Level 1)

```text
Run stage A4. Write the framework section and Level 1, grouped by part. Trace every
number to Level 4 or the workbook. Fill Country_Metrics.
```

## Prompt A5 - Peer review

```text
Run stage A5. Build the peer-review register, write statuses into the workbook, and
apply every correction to Levels 1 to 4.
```

## Prompt A6 - Final report

```text
Run stage A6. Assemble the report in the section 8.1 order, with this specification as
the last annex, and deliver the final report as a Word document and the final workbook
as an Excel file.
```

Save the two files in `ExciseDiagnostic/country_results/Indonesia/` as
`Indonesia_Excise_Diagnostic_v0.1.docx` and `Indonesia_Excise_Diagnostic_v0.1.xlsx`.

---

## Short on time? Do one part only (fuel)

Add `excise_part1_fuel.md` to the Project, then use the prompts in that file instead of
A1-A6 above. Parts 2-4 (vehicles, health, other) work the same way. Part 5 runs last.

## Variant - a fossil fuel SUBSIDY diagnostic for Indonesia

There is no subsidy process file in the repo yet, so this output is less standard than
the excise one. Treat it as a first draft. Upload the same master file (Claude reuses
its structure and quality rules) and send:

```text
Using the structure and quality rules of the excise diagnostic process file (four
levels, every number traced to a source, data gaps marked data_gap and never guessed,
a peer-review register), run a fossil fuel SUBSIDY diagnostic for Indonesia at
2026-09-30. Cover: (1) the legal and budget basis of each energy subsidy and
compensation payment (fuel - Pertalite, Solar/diesel, LPG 3 kg - and electricity), with
the instrument and year; (2) amounts budgeted and realised for the last five years, in
IDR and % of GDP, from the budget documents (APBN, Nota Keuangan, LKPP);
(3) regulated retail prices vs. a cost-recovery reference price; (4) who benefits;
(5) compare your total with the IMF Fossil Fuel Subsidies figure in the dashboard
(Indonesia 2024: explicit 2.98% of GDP, implicit 9.20% of GDP) and explain any gap.
Give a link for every source. Do one section per message; start with (1) and wait.
```
