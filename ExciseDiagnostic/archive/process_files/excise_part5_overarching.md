# EXCISE DIAGNOSTIC - PART 5: OVERARCHING (SECTORAL RUN, ALWAYS LAST)

```yaml
part_id: "5 Overarching"
file: excise_part5_overarching.md
spec_version: "1.1"                     # read together with excise_diagnostic_process.md (MASTER)
standard_schedule_version: "1.0"
codes: []                               # no codes of its own
register_prefix: "PR-X"
run_order: "LAST. Start only when the four part workbooks exist and have been peer reviewed."
inputs_from_parts:
  - "<ISO3>_excise_part1_fuel.xlsx and report"
  - "<ISO3>_excise_part2_vehicles.xlsx and report"
  - "<ISO3>_excise_part3_health.xlsx and report"
  - "<ISO3>_excise_part4_other.xlsx and report"
outputs:
  report: "<ISO3>_excise_diagnostic_report (docx): the full diagnostic, whole-run layout (master 8.1)"
  workbook: "<ISO3>_excise_diagnostic.xlsx: the master workbook in the cross-country structure"
```

## 0. HOW TO USE THIS FILE

1. Give Claude the MASTER specification (excise_diagnostic_process.md) and this file.
2. Upload the four part workbooks and part reports, and each part's list of framework items.
3. Supply the framework inputs in section 2. If Claude cannot fetch a source, it runs the data step
   (section 4) and asks you for a set of downloads: a links list plus a .bat file.
4. Send the prompts in section 8 one at a time. Check each output before sending the next.

## 1. WHAT THIS PART DOES

Some material belongs to no single part. Without an overarching part it would be repeated four times
or lost. Part 5 does two jobs:

- **Framework.** Aggregate revenue; the legal framework; the indexation rule; the records
  reconciliation (section 3).
- **Aggregation.** Consolidate the four part workbooks into one master workbook, and the four part
  reports into one report with the whole-run layout: framework first, then Levels 1-3 grouped by part,
  then a shared Level 4 and a single peer-review register.

## 2. INPUTS

```yaml
part_outputs: "Four part workbooks (sheets: Common, Summary, Country_Metrics, Std_Schedule,
               Std_Schedule_Code, Parameters, DQ_Scale, Sources, Register) and four part reports"
framework_legal_texts: "The excise law's general provisions: charging sections, definitions,
                        effective-date rules, stamps, exemptions, indexation clauses; any
                        'Current' and 'Proposed' rate columns and the commencement provision that
                        decides which applies"
collections: "Total excise collections and, where published, by group, from every available record
              (revenue authority, finance ministry fiscal tables, budget, IMF reports), latest two
              years, with nominal GDP"
```

## 3. FRAMEWORK CONTENT

```yaml
aggregate_revenue:
  - "Total excise collected (% of GDP), latest two years; excise-like levies outside the law shown
     separately"
  - "Collected and implied revenue by part, from the part workbooks"
  - "Share of collections by part"
legal_framework:
  legislative_history: "The instruments list (from the Common sheet), with which parts each one
                        touched"
  current_or_proposed: "Where a law prints 'Current' and 'Proposed' columns, state which column is
                        in force and cite the commencement provision"
  effective_dates: "Commencement and transition dates; staggered rates; administrative start dates
                    (for example, when the revenue authority begins applying a rate)"
  currency: "Notations used (local, old and new units, US$) and how each was resolved"
  stamps: "Excise or tax stamp requirements and the products they cover"
  exemptions: "General exemption regimes (diplomatic, investment incentives, government) that cut
               across parts"
indexation_rule: "Any rule that adjusts specific rates for inflation or exchange rates. It usually
                  covers tobacco, alcohol and SSBs together, so it is analyzed here, not in Part 3.
                  State the index, the frequency, the lines covered, and whether it has been
                  applied."
records_reconciliation: "Compare excise collections across records (revenue authority, finance
                         ministry, budget, IMF). Then compare collected with implied revenue by
                         part. Explain each gap: exemptions, arrears, administration, or data."
```

## 4. DATA STEP (WHEN NEEDED)

If a framework input cannot be fetched, Claude stops and produces
`datastep_<ISO3>_part5_overarching_links.md` and `datastep_<ISO3>_part5_overarching.bat` using the
template in master section 12. Typical downloads: the consolidated excise law, the revenue authority's
annual report, the finance ministry's fiscal tables and the latest IMF country report.

## 5. STAGES

```yaml
O1_consolidation_checks:
  do:
    - "Check the four Common sheets are identical. If not, list the differences, stop, and ask the
       reviewer which to keep."
    - "Check codes: Parts 1-4 together hold all 30 standard codes, each exactly once (part map,
       master section 2A)."
    - "Check register ids are unique across PR-F, PR-V, PR-H and PR-O."
    - "Check each part workbook recalculated with zero errors."
  output: "Consolidation note"
O2_framework:
  do: "Write the framework section (section 3) from the part outputs and framework inputs.
       Collect the framework items each part handed over."
O3_aggregation:
  do:
    - "Master workbook: stack the four parts' Std_Schedule and Std_Schedule_Code rows (keep row ids
       and part tags); one Parameters row; merge Country_Metrics; merge Registers; add Part_Map,
       Part_Summary and Framework sheets."
    - "Report: Level 1 across parts (key metrics, headline findings, inconsistencies, peer metrics,
       issues for dialogue), then Levels 1-3 grouped by part, Level 4 shared (master section 8.1)."
    - "Recompute totals from the stacked rows; never add up numbers typed in part reports."
  check: "Part_Map checks show 30 codes and 0 mismatches. Every total traces to the stacked rows."
O4_framework_peer_review:
  do: "Master A5 for framework items only, with PR-X ids. Part registers are carried over unchanged
       unless O1-O3 found an error, which is logged as a new PR-X entry."
O5_assemble:
  do: "Assemble the overarching report (master section 8.1) and deliver the master workbook."
```

## 6. MASTER WORKBOOK

```yaml
file: "<ISO3>_excise_diagnostic.xlsx"
structure: "Cross-country workbook structure (master section 11): Summary, Country_Metrics,
            Std_Schedule, Part_Map, Part_Summary, Std_Schedule_Code, Parameters, DQ_Scale, Sources,
            plus Framework and Register"
rules:
  - "Stack part rows; do not retype them."
  - "Part column values: '1 Fuel', '2 Vehicles', '3 Health', '4 Other'."
  - "Summary and Part_Summary are formulas over the stacked rows."
  - "Recalculate with zero errors before delivery."
to_add_to_cross_country_workbook: "Append the country's Parameters row and Std_Schedule_Code rows;
                                   add the country's columns to Std_Schedule and Country_Metrics."
```

## 7. REPORT OUTLINE

```text
Front matter  Title, version note, contents, executive summary (across all parts)
Part I        Method (standard text)
Part II  7    Scope, sources, comparator, and how the five parts were run
         8    Framework: aggregate revenue; legal framework; indexation rule; records reconciliation
         9    Level 1: across parts, then by part (fuel, vehicles, health, other)
         10   Level 2: standardized schedule grouped by part
         11   Level 3: country schedule grouped by part
         12   Level 4 (shared): parameters; computations; collections; legislative history
         13   Conclusions and next steps
Annexes       References; single peer-review register (PR-F, PR-V, PR-H, PR-O, PR-X);
              original prompts (historical); process specification (MASTER, last)
```

## 8. PROMPT TEMPLATES

```text
[O1] Using the master specification and this part file, run stage O1 for {COUNTRY}: check the four
     part workbooks (Common sheets, codes, register ids, recalculation) and write the
     consolidation note.
[O2] Run stage O2: write the framework section (aggregate revenue, legal framework, indexation rule,
     records reconciliation). If you cannot fetch a framework source, give me the links list and
     the .bat file.
[O3] Run stage O3: build the master workbook from the four part workbooks, and write Levels 1-4 in
     the whole-run layout, grouped by part.
[O4] Run stage O4: the PR-X register for framework items; apply corrections.
[O5] Run stage O5: assemble the overarching report and deliver the master workbook.
```
