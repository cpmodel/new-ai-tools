# EXCISE DIAGNOSTIC - PROCESS SPECIFICATION (MASTER)

```yaml
spec_id: excise-diagnostic-process
spec_version: "1.1"
standard_schedule_version: "1.0"      # the 30 rows are unchanged
date: 2026-09
status: MASTER
supersedes: "Spec 1.0, and the original prompts (Annex E of the working paper), which are a
             historical record only."
changes_from_1_0:
  - "Two run modes: WHOLE (one run) or SECTORAL (five parts, Part 5 overarching run last)."
  - "Part map: every standard code belongs to exactly one of Parts 1-4 (section 2A)."
  - "Data step: when Claude cannot fetch a source, it asks for downloads (links list + .bat)."
  - "Workbook: every run populates a workbook in the cross-country structure (section 11)."
  - "Output order (section 8): framework section first; Levels 1-3 grouped by part; Level 4 and
     the peer-review register shared."
part_files:
  - excise_part1_fuel.md
  - excise_part2_vehicles.md
  - excise_part3_health.md
  - excise_part4_other.md
  - excise_part5_overarching.md       # run last
rule_of_precedence: "Where this file conflicts with any other instruction, this file governs,
                     unless a human reviewer overrides it in writing. Part files add detail;
                     they do not override this file."
language: "US English. Short words, short sentences."
```

## 0. HOW TO USE THIS FILE

1. Choose a run mode (section 2A): WHOLE or SECTORAL.
2. Give this file to Claude as context at the start of every run: upload it to a Claude Project, or paste
   it as the first message. In a sectoral run, also give the part file for the part being run.
3. Supply the INPUTS in section 2.
4. If Claude cannot fetch a source, it runs the DATA STEP (section 12): it stops and asks for a set of
   downloads, as a list of links plus a .bat file.
5. Run the stages in order, one message per stage, using the prompt templates (section 10, or the part
   file). A human reviewer checks each stage's output before the next stage starts.
6. Every run produces a report AND a workbook (section 11).
7. A run covers ONE country. See the COMPARATOR RULE in section 3.

## 1. WHAT THE DIAGNOSTIC IS

- A structured snapshot of one country's excise schedule at a legal cut-off date.
- Pre-analytical. It does not model behavior or say what a rate should be.
- It does five things: categorize every line; standardize it; surface erosion; surface gaps;
  surface incoherence.
- It covers every excisable line plus adjacent levies that work as excises. It concentrates on six
  priority categories: fuel, vehicles, tobacco, alcohol, sugar-sweetened beverages (SSBs), gambling.
- The report has four levels. It is READ top-down and BUILT bottom-up:
  - Level 1: top level (key metrics, findings, inconsistencies, peer metrics, issues for dialogue)
  - Level 2: standardized schedule (fixed rows, section 4.5) and comparison
  - Level 3: country schedule (every line, five fields, rate trajectory)
  - Level 4: raw data (parameter sheet, computations, collections, legislative history)

## 2. INPUTS

```yaml
required:
  country: "<name>"
  currency_code: "<ISO code and every notation used in the law, e.g. SLE; printed as Le, NLe>"
  cut_off_date: "<YYYY-MM-DD>"
  legal_texts:
    - "Consolidated excise law, if one exists"
    - "Every Finance Act or excise amendment from 2020 (or the last consolidation) to the cut-off"
    - "Any pending bill (record it; do not apply it)"
  budget_documents:
    - "Budget speeches and citizens' budgets for each year (used to confirm rates and notation)"
  collections:
    - "Excise collections by line or group for the latest two years, with nominal GDP in local
       currency"
  parameters: "See parameter schema in section 5. Missing values are marked data_gap, never guessed."
optional:
  comparator_report: "A PRE-RUN diagnostic for another country, produced with THIS spec and the same
                      standard_schedule_version. Must include its Level 2 table and Level 4
                      parameter sheet. See section 3."
  benchmarks: "International reference table. Default in section 9."
sectoral_runs_also_need:
  common_header: "See section 2B. Created by the first part run; reused unchanged by the others."
```

## 2A. RUN MODES

```yaml
WHOLE:
  what: "One run covers all 30 standard lines."
  instructions: "This file, stages A1-A6 (section 6), prompts in section 10."
  outputs: ["one report", "one workbook"]
  use_when: "A single team, a small schedule, or a first diagnostic for a country."

SECTORAL:
  what: "Five parts. Parts 1-4 each cover a fixed set of standard codes. Part 5 (overarching)
         has no codes; it builds the framework and aggregates Parts 1-4."
  instructions: "This file plus one part file per run (see part_files in the header)."
  run_order: "Parts 1-4 in any order (they can run in parallel). Part 5 ALWAYS runs last."
  outputs: "Each part produces its own report and its own workbook. Part 5 reads the four part
            workbooks as inputs and produces the overarching report and the master workbook."
  use_when: "Different specialists own different sectors; a large schedule; or a joint paper
             (Parts 1-2 line up with a fuel and vehicle tax diagnostic)."
```

Part map (every code in exactly one part; no leftovers):

```text
part          codes                                   content
1 Fuel        E1 E2 E3 E4 E5 E6                       petrol, diesel, kerosene, aviation fuel, heavy fuel
                                                      oil and other gas oils, lubricants; fuel levies
                                                      outside the excise law (road fund, regulator,
                                                      price-stabilization levies)
2 Vehicles    E8 E9                                   annual circulation levies; import, first-
                                                      registration and luxury-vehicle excises
3 Health      T1 T2 T3 A1 A2 A3 A4 A5 S1 S2 S3 S4 I3  tobacco, alcohol, sugar-sweetened beverages
                                                      and sugar
4 Other       E7 G1 D1 D2 F1 I1 I2 I4 L1              plastics, gambling, telecoms and digital,
                                                      financial transactions, cement and other
                                                      aggregates, edible oils, other industrial
                                                      inputs, luxury goods
5 Overarching (none)                                  framework and aggregation; run last
totals: {part1: 6, part2: 2, part3: 13, part4: 9, all: 30}
note: "Parts are production groups. The typology (section 4.1) still classifies lines for analysis.
       Three placements differ from the typology: plastics E7 and gambling G1 are produced in
       Part 4; sugar I3 is produced in Part 3."
```

## 2B. COMMON HEADER (SECTORAL RUNS)

```yaml
common_header:
  country: ""
  iso3: ""
  cut_off_date: ""
  currency: {code: "", notations_in_law: [], notation_resolution: "", review_ref: ""}
  fx: {value: 0, unit: "LCU per USD", as_of: "", source: "", finance_ministry_rate: 0}
  gdp_nominal: {lcu: 0, usd: 0, year: "", source: ""}
  cpi: {series: "", source: ""}                   # for erosion (Q5)
  legal_instruments:                              # every instrument in the history, all parts
    - {id: "S5", title: "", provision_style: "", effective: "", status: "in_force | pending"}
  sources: []                                     # shared source ids S1, S2 ...
rules:
  - "The first part run creates the common header and saves it as the Common sheet."
  - "Every later part takes that Common sheet as an input and does not change it. If a part finds
     an error in it, it records a review entry and flags it to Part 5; it does not edit it."
  - "Part 5 checks that all four copies are identical before it aggregates."
```

## 3. COMPARATOR RULE (HARD CONSTRAINT)

- A run of this spec produces a SINGLE-COUNTRY diagnostic.
- Do NOT produce comparator columns, ratios or comparator narrative unless `comparator_report` is
  supplied.
- NEVER fill comparator values from memory, from web search, or by analyzing a second country inside
  the same run.
- To add a comparator: (1) run this spec for country B as its own run; (2) complete peer review for B;
  (3) supply B's report and workbook as `comparator_report` in the run for country A.
- In a sectoral run, the comparator for a part is the SAME part of country B (for example, B's Part 3
  workbook for A's Part 3). Part 5 may use B's master workbook.
- The comparator must use the same `standard_schedule_version`, the same three measures and the same
  conversion conventions. If it does not, list the differences and do not compute ratios for the
  affected lines.
- Without a comparator: Level 1 benchmarks against the international references in section 9 only;
  Level 2 shows single-country columns; the ratio column is omitted.

## 4. FIXED DEFINITIONS

### 4.1 Typology (assign each line by policy rationale, not HS code)

```yaml
typology:
  - energy_environment      # fuels, lubricants, plastics, vehicle registration and circulation levies
  - health                  # tobacco, alcohol, SSBs, gambling
  - luxury                  # cosmetics, perfumes, furniture, jewelry
  - telecoms_digital        # airtime, voice, data, SMS, international calls
  - financial_transactions  # mobile money, bank and ATM fees, payment services
  - industrial_residual     # cement, sugar, edible oils, soaps, nails, other legacy lines
  - resource_extractive     # minerals, timber, other natural resources
  - procedural              # payment windows, returns, penalties (recorded, not measured)
tie_breaks:
  vehicle_levy: energy_environment
  gambling: health
```

### 4.2 Five recording fields (every line)

```yaml
fields: [category, item_as_printed, current_rate_as_printed, structure_and_differentiation,
         recent_changes]
structure: [specific, ad_valorem, hybrid_higher_of, compound]
differentiation: [origin, local_content, abv_band, sugar_content, pack_type, user, age, engine_size,
                  value, none]
recent_changes: "Every amendment since 2020 with instrument, provision and effective date"
```

### 4.3 Three comparable measures

```yaml
usd_per_standard_unit: "rate / fx, per liter (fuel), per liter of pure alcohol (alcohol),
                        per pack of 20 (cigarettes), per kg or per 50 kg bag, per liter (SSBs),
                        per vehicle or per vehicle-year (vehicles)"
share_of_consumer_price: "excise per unit / retail price per unit"
implied_revenue_pct_gdp: "rate x annual volume / nominal GDP. Upper bound when volume is
                          prevalence-based. Always report next to COLLECTED revenue where
                          available."
```

### 4.4 Conversion rules

```yaml
fx: "Working mid-rate near the cut-off. State the source. Compare with the finance ministry's rate
     and report the gap."
cigarettes: "rate per 1,000 sticks x 0.02 = rate per pack of 20"
alcohol_lpa: "rate per liter / representative ABV. Beer 0.05, opaque beer 0.04, wine 0.12,
              spirits 0.40"
hybrid: "Evaluate at a representative price; report the binding component"
ad_valorem_without_legal_base: "Apply to retail price as an upper bound; say so"
currency_notation: "Quote as printed. Resolve ambiguous notation (e.g. old vs new currency) from
                    government texts such as the budget speech. Record the resolution as a review
                    entry."
```

### 4.5 Standard schedule, version 1.0 (30 lines; do not change rows within a version)

```text
code  line                                           unit                   part
E1    Petrol                                         per liter              1 Fuel
E2    Diesel, road                                   per liter              1 Fuel
E3    Kerosene                                       per liter              1 Fuel
E4    Aviation fuel                                  per liter              1 Fuel
E5    Heavy fuel oil / other gas oils                per liter              1 Fuel
E6    Lubricants                                     per liter              1 Fuel
E7    Plastic bags and packaging                     per kg                 4 Other
E8    Vehicles: annual circulation                   per vehicle-year       2 Vehicles
E9    Vehicles: first registration / import          per vehicle            2 Vehicles
T1    Cigarettes                                     per pack of 20         3 Health
T2    Cigars and other tobacco                       per pack / per kg      3 Health
T3    E-cigarettes and novel products                per ml / per cartridge 3 Health
A1    Beer, mainstream malt (5% ABV)                 per LPA                3 Health
A2    Beer, local-content concession (5% ABV)        per LPA                3 Health
A3    Traditional / opaque beer (4% ABV)             per LPA                3 Health
A4    Wine (12% ABV)                                 per LPA                3 Health
A5    Spirits (40% ABV)                              per LPA                3 Health
S1    Carbonated and other soft drinks               per liter              3 Health
S2    Fruit and vegetable juice                      per liter              3 Health
S3    Energy drinks                                  per liter              3 Health
S4    Bottled and processed water                    per liter              3 Health
G1    Gambling, betting and lotteries                % of base              4 Other
D1    Voice and airtime                              per minute / % of fee  4 Other
D2    Data and value-added services                  % of fee               4 Other
F1    Mobile money and financial services            % of value or fee      4 Other
I1    Cement and other aggregates                    per 50 kg bag          4 Other
I2    Edible oils                                    per liter / %          4 Other
I3    Sugar                                          per kg                 3 Health
I4    Other industrial inputs                        %                      4 Other
L1    Cosmetics, furniture and other luxury          %                      4 Other
coverage_codes: {1: "in excise law", 2: "levy outside excise law", 0: "not covered"}
labels: ["Not excisable", "Not identified (confirm)", "Levy"]
```

## 5. DATA SCHEMAS (MACHINE-READABLE)

```yaml
parameter:
  id: "P-01"
  name: "exchange_rate | gdp_nominal | pump_price_petrol | volume_diesel | tobacco_users | ..."
  value: 22
  unit: "NLe per USD"
  as_of: "2026-04"
  source: "S12"                 # id in the source list
  status: "published | estimate | assumption | data_gap"
  part: "common | 1 Fuel | 2 Vehicles | 3 Health | 4 Other"
  note: "free text"

schedule_line:
  id: "L-001"
  part: "1 Fuel"
  category: "energy_environment"
  std_code: "E1"
  item_as_printed: "Petrol"
  rate: {value: 5.00, currency_as_printed: "NLe", unit: "liter", text_as_printed: "NLe5.0 per litre"}
  structure: "specific"
  differentiation: ["none"]
  history:
    - {instrument: "Finance Act 2024", provision: "s.4(k)", effective: "2024-01-01",
       rate_as_printed: "Le 2.80 per litre"}
    - {instrument: "Finance Act 2026", provision: "", effective: "2026-01-01",
       rate_as_printed: "NLe5.0 per litre"}
  refs: ["PR-F-01", "PR-F-04"]  # at least one; no reference, no line

source:
  id: "S5"
  title: "Finance Act, 2024 (No. 1 of 2024)"
  location: "URL or file name"
  accessed: "YYYY-MM-DD"
  obtained_by: "fetched | data_step_download | supplied_by_user"

review_entry:
  id: "PR-F-01"                 # whole run: PR-nn. Sectoral: PR-F (fuel), PR-V (vehicles),
                                #   PR-H (health), PR-O (other), PR-X (overarching framework)
  part: "1 Fuel"
  area: "Energy"
  claim_checked: "Petrol Le 2.80/l (FA2024)"
  sources: ["S5"]
  location: "s.4(k), p.11"
  quote: "exact words from the source"
  finding: "what the check found"
  status: "verified | verified_open_point | corrected | partly_verified | not_verified |
           not_reviewed | estimate | last_legible | assumption | data_gap"
  action: "what changed in the report"
```

Status definitions and data-quality (DQ) scores (same scale as the workbook's DQ_Scale sheet):

```text
status               DQ  meaning
verified              5  matches the primary legal text or official data, quote on file
verified_open_point   4  text checked and matches; an interpretation point is still open
corrected             4  error or omission found in peer review; corrected; now matches the source
partly_verified       3  confirmed only in an official secondary source, or only in part
not_verified          2  no accessible source confirmed it
not_reviewed          2  extracted from the law but not yet peer reviewed
estimate              2  derived from published inputs with estimated volumes or prices
last_legible          1  current rate not legible; last legible (superseded) rate shown
assumption            1  unsourced assumption or indicative figure
data_gap              1  no data; not estimated
rule: "A derived figure takes the lowest DQ of its inputs."
```

## 6. STAGES

### 6.1 Whole run (stages A1-A6)

```yaml
A1_legislative_history:
  builds: "Level 4 legislative history and raw rate data"
  do:
    - "Identify the starting position (1 January 2020, or the last consolidation)."
    - "List every amending instrument to the cut-off: provision, effective date, headings touched."
    - "For each provision, quote each rate as printed. If a table cannot be read, write 'not legible'."
    - "Record pending bills separately. Do not apply them."
    - "Separate customs-tariff changes from excise changes."
    - "If a source cannot be fetched, run the data step (section 12) before going on."
  check: "Every instrument in the inputs appears. No year is skipped without saying 'no excise
          changes found'."

A2_country_schedule:
  builds: "Level 3 schedule (every line, five fields) and rate trajectory table"
  do:
    - "Build the current schedule at the cut-off from A1."
    - "Give each line a schedule_line record with refs and its part."
    - "Where the current rate is not legible, show the last legible rate and flag it."
    - "Record adjacent levies outside the excise law."
  check: "Every line has a source. Currency notation is resolved and recorded."

A3_standardize_and_measure:
  builds: "Level 2 standardized schedule, Level 4 parameter sheet and computations, and the
           workbook sheets Std_Schedule, Std_Schedule_Code and Parameters"
  do:
    - "Map each national line to the section 4.5 codes. Show ranges for tiers."
    - "Fill the parameter sheet. Every value has a source and a status."
    - "Compute the three measures with calculations shown step by step."
    - "Add collected revenue by line or group where available."
    - "Count coverage (codes 1, 2, 0)."
    - "Populate the workbook (section 11): one Std_Schedule_Code row per rate component; USD and
       share-of-price columns by formula from Parameters."
    - "Add comparator columns ONLY if comparator_report is supplied (section 3)."
  check: "Recompute three lines by hand. Totals equal the sum of their parts. Workbook recalculates
          with zero errors."

A4_top_level:
  builds: "Framework section and Level 1"
  do:
    - "Framework: aggregate revenue, legal framework, indexation rule, records reconciliation
       (see excise_part5_overarching.md, section 3, for the content)."
    - "Key metrics table: collected and implied revenue, rate intensity, structure and base."
    - "Five or six headline findings, each with numbers traced to Level 4."
    - "Inconsistency table: issue and evidence, why it matters, refs."
    - "Peer metrics against section 9 benchmarks (and the comparator, if supplied)."
    - "Issues for policy dialogue. Do not recommend rates."
    - "Fill the workbook's Country_Metrics sheet."
  check: "Every number in Level 1 appears in Levels 2 to 4 or in the workbook."

A5_peer_review:
  builds: "Reference register annex and DQ scores in the workbook"
  do:
    - "Create one review_entry per legal line or group of lines, per parameter and per aggregate."
    - "Open each source and compare the claim with the text. Quote the text."
    - "Set status. Correct the report for every 'corrected' entry."
    - "Write each status into the workbook (review_status, peer_review_ref columns)."
    - "Tally statuses and list open items for the revenue authority."
  check: "No legal line without a ref. No 'verified' without a quote from the source."

A6_assemble:
  builds: "The report and the final workbook"
  order: "section 8"
```

### 6.2 Sectoral run

```yaml
parts_1_to_4:
  file: "excise_part<n>_<name>.md"
  stages: "P1 history, P2 schedule, P3 standardize and measure, P4 part findings, P5 peer review,
           P6 part report and part workbook. Each stage is A1-A5 restricted to the part's codes."
  hand_off_to_part_5: ["part report", "part workbook", "list of framework items found"]
part_5_overarching:
  file: "excise_part5_overarching.md"
  runs: "LAST, after all four part workbooks exist and are peer reviewed."
  stages: "O1 consolidation checks, O2 framework, O3 aggregation, O4 framework peer review,
           O5 overarching report and master workbook."
```

## 7. QUALITY RULES (EVERY STAGE, EVERY PART)

1. Compute every number from the parameter sheet. Never type a number into the narrative that Level 4
   or the workbook does not produce.
2. Every legal line carries at least one reference id. No reference, no line.
3. Quote rates as printed, with currency notation. Resolve ambiguity from government texts.
4. Report collected revenue next to implied revenue. Implied revenue from prevalence-based volumes is
   an upper bound.
5. If a table cannot be read, write "not legible" and show the last legible rate. Never guess.
6. Do not rely on memory for rates or data. Use the supplied documents or cited sources. Mark anything
   else "not verified".
7. Customs duties are not excise. Exclude them, but note any that look like excise.
8. Flag pending bills; do not apply them.
9. Say what the diagnostic does not do. It does not recommend rates.
10. Comparator content only from a supplied pre-run report (section 3).
11. If a source cannot be fetched, stop and run the data step (section 12). Do not fill the gap from
    memory.
12. In a sectoral run, stay inside the part's codes. Items that belong to no single part go to Part 5
    as framework items.

## 8. OUTPUT STRUCTURE

### 8.1 Whole run: one report

```text
Front matter: title, version note (what changed), contents, executive summary
Part I   Method (reuse the standard text; update only if this spec changes)
Part II  7  Scope, sources and comparator (state "none" if no comparator_report)
         8  Framework: aggregate revenue; legal framework (legislative history, current or
            proposed columns, effective dates, currency, stamps and exemptions); indexation rule;
            records reconciliation
         9  Level 1, grouped by part: key metrics; findings; inconsistencies; peer metrics;
            issues for policy dialogue
         10 Level 2, grouped by part: standardized schedule (section 4.5) and comparison
         11 Level 3, grouped by part: country schedule and rate trajectory
         12 Level 4 (shared): parameter sheet; computations; cross-check with collections;
            legislative history
         13 Conclusions and next steps
Annexes  Comparator schedule and history (only if comparator_report supplied)
         References and data sources
         Peer-review register (single register, grouped by part)
         Original prompts (historical)
         Process specification (this file; MASTER; always last)
Workbook One workbook in the cross-country structure (section 11)
```

### 8.2 Sectoral run: five reports and five workbooks

```text
Part 1-4 report   Scope; Level 1 for the part; Level 2 rows for the part's codes; Level 3 lines;
                  Level 4 part parameters and computations; part register (PR-F / V / H / O);
                  framework items handed to Part 5
Part 1-4 workbook Common, Summary, Country_Metrics (part metrics), Std_Schedule (part rows),
                  Std_Schedule_Code (part rows), Parameters, DQ_Scale, Sources, Register
Part 5 report     The whole-run report structure (section 8.1), built from the four part reports
                  and workbooks, plus the framework section and a consolidation note
Part 5 workbook   Master workbook: the four part workbooks consolidated, plus Part_Map,
                  Part_Summary, Framework and the merged Register
```

## 9. DEFAULT INTERNATIONAL BENCHMARKS (replace when better data are supplied)

```yaml
fuel:     {oecd_petrol_usd_per_liter: 0.55, oecd_share_of_pump_price: "30-50%",
           oecd_pct_gdp: "1.5-2.5%", south_africa_usd_per_liter: "0.20-0.25"}
excise_total_pct_gdp: {africa_average_2021: 2.4, low_income_average_2021: 2.2}
tobacco:  {who_excise_share_of_retail: ">=70%", who_total_tax_share: ">=75%",
           south_africa_pct_gdp: 0.4, eu_pct_gdp: 0.6}
alcohol:  {oecd_beer_usd_per_lpa: "15-40", oecd_spirits_usd_per_lpa: "25-50",
           oecd_pct_gdp: "0.4-0.6%"}
ssb:      {uk_sdil_usd_per_liter: "0.23-0.30 (sugar-graduated)", mexico_usd_per_liter: "~0.05 (flat)"}
```

## 10. STAGE PROMPT TEMPLATES FOR A WHOLE RUN (ONE PER MESSAGE)

```text
[A1] Using the excise diagnostic process specification (master) and the legal texts provided,
     run stage A1 for {COUNTRY}, from {START} to {CUT_OFF}. Output the legislative history
     instrument by instrument, quoting every rate as printed with its provision and effective date.
     Mark unreadable tables "not legible". List customs-tariff changes separately. List pending bills.
     If you cannot fetch a source, run the data step: give me the links list and the .bat file.

[A2] Run stage A2. Build the current schedule at {CUT_OFF} as schedule_line records (section 5),
     tagged by part, then as a Level 3 table with a Ref. column. Add the rate trajectory table.

[A3] Run stage A3. Fill the parameter sheet (section 5) from the data provided and flag gaps.
     Map every line to standard schedule v1.0 and compute the three measures, showing each
     calculation. Add collections. Populate the workbook (section 11) and recalculate it.
     {IF COMPARATOR: Use the supplied pre-run report for {COMPARATOR} as the comparator.
     ELSE: No comparator; omit comparator columns and ratios.}

[A4] Run stage A4. Write the framework section and Level 1, grouped by part. Trace every number
     to Level 4 or the workbook. Fill Country_Metrics.

[A5] Run stage A5. Build the peer-review register, write statuses into the workbook, and apply
     every correction to Levels 1 to 4.

[A6] Run stage A6. Assemble the report in the section 8.1 order, with this specification as the
     last annex, and deliver the final workbook.
```

Sectoral prompt templates are in each part file.

## 11. WORKBOOK (EVERY RUN POPULATES ONE)

```yaml
structure: "Same sheets and columns as the cross-country workbook (Excise_Diagnostic_CrossCountry.xlsx)."
sheets:
  Summary:           "Headline indicators by formula over the other sheets; read-before-use notes"
  Country_Metrics:   "One row per metric (Context X, Revenue R, Rate intensity I, Structure B,
                      Key questions Q, Open items O; sectoral parts add V for vehicles and OL for
                      other lines), with value, review status, DQ (formula) and source or ref;
                      column 'Part (spec 1.1)'"
  Std_Schedule:      "30 standard lines (human-readable): part, typology, code, line, unit, coverage,
                      statutory rate and structure, USD per unit, % of price, DQ, ref"
  Std_Schedule_Code: "Long format, one row per rate component (a higher-of hybrid has two rows with
                      the same option_group). Key columns: row_id, country_iso3, cut_off, part,
                      std_code, coverage_code, national_item, instrument, effective_date,
                      record_type (current | last_legible | not_covered), rate_component,
                      rate_value, rate_currency, rate_per_unit, text_as_printed, condition columns
                      (origin, ABV, sugar, local content), fx and USD columns by formula,
                      ref_price_per_rate_unit, share_of_ref_price, review_status, peer_review_ref,
                      dq_score (formula)"
  Parameters:        "Country row: currency, FX and status, cut-off, GDP, version. Inputs in blue."
  DQ_Scale:          "Status to DQ score (section 5)"
  Sources:           "Documents, versions, cut-off, peer-review tallies"
  Part_Map:          "Codes by part with consistency checks (master and whole-run workbooks)"
  Part_Summary:      "Coverage and DQ by part (master and whole-run workbooks)"
sectoral_part_workbook_adds:
  Common:            "The common header (section 2B), identical in all four part workbooks"
  Register:          "The part's review entries (section 5 review_entry fields)"
part_5_adds:
  Framework:         "Aggregate revenue by part (collected and implied), legal-framework table,
                      indexation rule, records reconciliation"
  Register:          "Merged register from Parts 1-4 plus PR-X entries"
formatting: "Arial 10. Inputs blue; formulas black; links to other sheets green. Recalculate and
             ship only with zero formula errors."
```

## 12. DATA STEP (WHEN CLAUDE CANNOT FETCH A SOURCE)

Sometimes Claude cannot retrieve a source itself: no web access in the environment, a host that blocks
automated access, a redirect loop, a paywall, or a PDF whose tables cannot be read as text. Then it runs
a data step instead of guessing.

```yaml
trigger: "Any required source that cannot be fetched or read."
claude_does:
  - "STOP the stage and say which sources are missing and why."
  - "Produce a links list: datastep_<ISO3>_<part>_links.md"
  - "Produce a Windows batch file: datastep_<ISO3>_<part>.bat"
  - "Wait for the files. When they arrive, check each one opens and is the right document,
     record obtained_by: data_step_download and the access date, then resume the stage."
links_list_columns: [file_id, file_name, url, what_it_is_for, stage_or_part, shared_with_parts,
                     manual_download]
file_naming: "<source id>_<short title>.<ext>, e.g. S05_Finance_Act_2024.pdf. Shared legal texts use
              the same name in every part, so the .bat skips files already downloaded."
manual_download: "Links the .bat cannot fetch (login, JavaScript page, form) are marked MANUAL in
                  the list, with instructions."
other_systems: "On Mac or Linux, ask for the same list as a .sh script."
```

Template for the .bat file (Windows 10 or later has curl.exe built in):

```bat
@echo off
setlocal
REM Excise diagnostic data step: {COUNTRY} ({ISO3}), {PART}. Generated {DATE}.
REM Downloads each source into sources\{ISO3}\ next to this file, skipping files already there.
set "OUT=%~dp0sources\{ISO3}"
if not exist "%OUT%" mkdir "%OUT%"
set "LOG=%OUT%\_download_log.txt"
echo Run %DATE% %TIME%>> "%LOG%"

call :get "S05_Finance_Act_2024.pdf" "https://example.gov/finance-act-2024.pdf"
call :get "S12_Macro_fiscal_tables.pdf" "https://example.gov/macro-tables.pdf"
REM One call line per file in the links list. Write any % in a URL as %%%% (call expands it twice).

echo.
echo Done. Upload the folder "%OUT%" to the Claude project.
echo Files marked FAILED in _download_log.txt need a manual download: see the links list.
pause
exit /b 0

:get
if exist "%OUT%\%~1" goto :skip
curl.exe -L --fail --retry 2 --connect-timeout 30 -A "Mozilla/5.0" -o "%OUT%\%~1" "%~2"
if errorlevel 1 goto :failed
echo OK      %~1>> "%LOG%"
exit /b 0
:skip
echo SKIP    %~1 - already present>> "%LOG%"
exit /b 0
:failed
echo FAILED  %~1 - %~2>> "%LOG%"
if exist "%OUT%\%~1" del "%OUT%\%~1"
exit /b 0
```
