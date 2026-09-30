# EXCISE DIAGNOSTIC - PART 4: OTHER (SECTORAL RUN)

```yaml
part_id: "4 Other"
file: excise_part4_other.md
spec_version: "1.1"                     # read together with excise_diagnostic_process.md (MASTER)
standard_schedule_version: "1.0"
codes: [E7, G1, D1, D2, F1, I1, I2, I4, L1]
register_prefix: "PR-O"
run_order: "Any order with Parts 1-3. Part 5 (overarching) runs after all four."
outputs:
  report: "<ISO3>_excise_part4_other_report (docx or md)"
  workbook: "<ISO3>_excise_part4_other.xlsx"
```

## 0. HOW TO USE THIS FILE

1. Give Claude the MASTER specification (excise_diagnostic_process.md) and this file.
2. Supply the common header (the Common sheet). If this is the first part run for the country, Claude
   creates it in stage P0 and every later part reuses it unchanged.
3. Supply the inputs in section 2. If Claude cannot fetch a source, it runs the data step (section 3)
   and asks you for a set of downloads: a links list plus a .bat file.
4. Send the prompts in section 9 one at a time. Check each output before sending the next.
5. Hand the part report, the part workbook and the framework-items list to Part 5.

## 1. SCOPE

```text
code  standard line                          what to include
E7    Plastic bags and packaging             bags, bottles, granules, single-use items
G1    Gambling, betting and lotteries        state the base: stakes, gross gaming revenue, winnings
D1    Voice and airtime                      airtime excise; per-minute surcharges; incoming
                                             international call charges; telecom surtaxes
D2    Data and value-added services          data, SMS, over-the-top levies
F1    Mobile money and financial services    withdrawal, transfer and commission levies; bank fees
I1    Cement and other aggregates            cement, clinker, concrete and other construction
                                             aggregates (per 50 kg bag or per tonne)
I2    Edible oils                            vegetable oils, palm olein, margarine, cooking fat
I4    Other industrial inputs                soaps, nails, fatty acids, iron rod and similar lines
                                             when set as excise
L1    Cosmetics, furniture and other luxury  cosmetics, perfumes, jewelry, furniture, other luxury
```

In scope: every excise line on these goods and services; levies outside the excise law that work as
excises (coverage 2), such as telecom surtaxes or cellular surcharges set in a Finance Act.

Out of scope (note, do not count): customs duties on the same goods (a common confusion: check whether
a Finance Act amends the excise schedule or the customs tariff); fees that only recover a service cost.
Sugar (I3) is in Part 3.

Framework items to hand to Part 5: general exemption regimes, stamps, currency notation, any
indexation clause that reaches these lines.

## 2. INPUTS

```yaml
common_header: "Common sheet (master section 2B)"
legal_texts: "Every provision on these lines from 2020 to the cut-off, including Finance Act
              levies outside the excise law (telecoms, mobile money)"
data:
  prices: "Cement per bag; plastic bag or bottle prices where relevant; typical airtime and data
           prices"
  volumes_or_bases: "Cement sales; telecom revenue; mobile-money values; gaming revenue"
  collections: "Collections by line or group where published, latest two years"
```

## 3. DATA STEP (WHEN NEEDED)

If any input in section 2 cannot be fetched, Claude stops and produces
`datastep_<ISO3>_part4_other_links.md` and `datastep_<ISO3>_part4_other.bat` using the template in
master section 12. Typical downloads: the Finance Acts (shared), the telecom regulator's annual report,
the gaming regulator's licensing notices, and cement industry or customs data.

## 4. PART-SPECIFIC MEASURES

```yaml
plastics: "USD per kg"
cement_and_aggregates: "USD per 50 kg bag (convert per-tonne rates: x 0.05)"
gambling: "rate and base; convert to % of gross gaming revenue where possible"
telecoms: "USD per minute, or % of fee; note which calls or services are excluded"
financial: "% of transaction value or fee; note the base"
ad_valorem_lines: "rate and legal base (ex-factory, CIF, retail)"
implied_revenue: "where a base is published: rate x base / GDP; else data_gap"
collected_revenue: "by line or group where published"
metrics_to_fill: ["OL1 plastics, USD per kg", "OL2 gambling rate and base",
                  "OL3 telecom and digital levies (text and % of fee)",
                  "OL4 financial-transaction levies (text)",
                  "OL5 cement and aggregates, USD per 50 kg bag",
                  "OL6 other-lines excise collected (% of GDP)"]
```

## 5. STAGES

```yaml
P0_common_header:  "Only if no Common sheet exists yet. See master section 2B."
P1_history:        "Master A1 for the nine codes; separate excise from customs changes."
P2_schedule:       "Master A2 for the nine codes: every line, five fields, refs, trajectory."
P3_measure:        "Master A3 for the nine codes: section 4 measures; populate the part workbook."
P4_part_findings:  "Level 1 for other lines: key metrics, 3-5 findings, inconsistencies (for example,
                    customs lines shown as excise; levies outside the excise law), issues for
                    dialogue. List framework items for Part 5."
P5_peer_review:    "Master A5 with PR-O ids. Quote every source. Write statuses into the workbook."
P6_report:         "Assemble the part report (section 7) and final part workbook (section 6)."
```

## 6. PART WORKBOOK

```yaml
file: "<ISO3>_excise_part4_other.xlsx"
sheets: [Common, Summary, Country_Metrics, Std_Schedule, Std_Schedule_Code, Parameters, DQ_Scale,
         Sources, Register]
rows: "Std_Schedule and Std_Schedule_Code hold ONLY codes E7, G1, D1, D2, F1, I1, I2, I4, L1, with
       part = '4 Other'."
parameters_added: [cement_price_per_bag, telecom_revenue, mobile_money_values, gaming_revenue,
                   other_collections]
checks: "Recalculate with zero errors. Row ids unique (<ISO3>-O-001 ...)."
```

## 7. PART REPORT OUTLINE

```text
1  Scope, sources and the laws that set these lines
2  Level 1 (other): key metrics; findings; inconsistencies; peer metrics; issues for dialogue
3  Level 2: standardized rows E7, G1, D1, D2, F1, I1, I2, I4, L1
4  Level 3: every line; trajectory since 2020; customs lines excluded (note)
5  Level 4: parameters and computations; cross-check with collections
6  Register PR-O (sources, quotes, statuses)
7  Framework items for Part 5
```

## 8. CHECKS BEFORE HANDING TO PART 5

- All nine codes present in Std_Schedule, with coverage 0, 1 or 2.
- No customs-tariff line recorded as excise.
- Common sheet unchanged from the first part run.
- Every line has a PR-O reference; every "verified" has a quote.

## 9. PROMPT TEMPLATES

```text
[P0] (First part only.) Create the common header for {COUNTRY} at {CUT_OFF} as the Common sheet.
[P1] Run stage P1 for Part 4 (other): the history of plastics, gambling, telecom, financial, cement
     and aggregates, edible oils, other inputs and luxury lines from {START} to {CUT_OFF}. Separate
     excise from customs. If you cannot fetch a source, give me the links list and the .bat file.
[P2] Run stage P2: the Level 3 schedule for the nine codes, with refs and the trajectory.
[P3] Run stage P3: fill the parameters, compute the measures, and populate the part workbook.
     {IF COMPARATOR: use the Part 4 workbook for {COMPARATOR}.}
[P4] Run stage P4: Level 1 for other lines, and the list of framework items for Part 5.
[P5] Run stage P5: the PR-O register; apply corrections; write statuses into the workbook.
[P6] Run stage P6: assemble the Part 4 report and deliver the final part workbook.
```
