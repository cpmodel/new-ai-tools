# EXCISE DIAGNOSTIC - PART 3: HEALTH (SECTORAL RUN)

```yaml
part_id: "3 Health"
file: excise_part3_health.md
spec_version: "1.1"                     # read together with excise_diagnostic_process.md (MASTER)
standard_schedule_version: "1.0"
codes: [T1, T2, T3, A1, A2, A3, A4, A5, S1, S2, S3, S4, I3]
register_prefix: "PR-H"
run_order: "Any order with Parts 1, 2 and 4. Part 5 (overarching) runs after all four."
outputs:
  report: "<ISO3>_excise_part3_health_report (docx or md)"
  workbook: "<ISO3>_excise_part3_health.xlsx"
note: "Usually the largest part: most legal changes (rate reforms, new product lines, indexation)
       fall here."
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
group          codes            what to include
Tobacco        T1 T2 T3         cigarettes (by tier, pack type, origin); cigars, loose, shisha;
                                e-cigarettes, heated tobacco and other novel products
Alcohol        A1 A2 A3 A4 A5   malt beer; local-content or concessionary beer; traditional or
                                opaque beer; wine; spirits (and other fermented beverages, ethyl
                                alcohol, noted under the nearest code)
SSBs           S1 S2 S3 S4      carbonated and other soft drinks; juice; energy drinks; bottled
                                water, including sugar-content bands
Sugar          I3               excise on sugar itself (per kg), with industrial-use exemptions
```

In scope: every rate, band and tier; ABV and sugar-content bands; local-content and origin
differentials; stamps rules that apply only to these products; the text of any indexation clause on
these lines (quoted and referenced, then handed to Part 5).

Out of scope (hand to Part 5): the indexation rule as a whole (it usually covers tobacco, alcohol and
SSBs together), general stamp regimes, currency notation, general exemptions, total collections.
Gambling is in Part 4.

## 2. INPUTS

```yaml
common_header: "Common sheet (master section 2B)"
legal_texts: "Every provision on tobacco (24.01-24.04), alcohol (22.03-22.09 and local items),
              non-alcoholic beverages (20.09, 22.01-22.02) and sugar (17.01) from 2020 to the
              cut-off; indexation clauses"
data:
  retail_prices: "Cigarette pack (by tier), beer, wine, spirits, soft drink, juice, sugar per kg"
  consumption: "Tobacco use prevalence (WHO), alcohol per capita (WHO), SSB and sugar volumes
                where published; formal-market volumes from industry or customs"
  collections: "Excise by product group where published (tobacco, alcohol, beverages), latest two
                years"
```

## 3. DATA STEP (WHEN NEEDED)

If any input in section 2 cannot be fetched, Claude stops and produces
`datastep_<ISO3>_part3_health_links.md` and `datastep_<ISO3>_part3_health.bat` using the template in
master section 12. Typical health downloads: the Finance Acts (shared), the revenue authority's excise
schedule or notices, the WHO tobacco and alcohol country profiles, and industry or customs volume data.
If a key table in a law cannot be read as text, the data step asks for a clearer copy.

## 4. PART-SPECIFIC MEASURES

```yaml
cigarettes: {usd_per_pack_20: "rate per 1,000 sticks x 0.02 / fx", share_of_retail: "per pack",
             benchmark: "WHO excise >= 70% of retail; total tax >= 75%"}
alcohol: {usd_per_lpa: "rate per liter / ABV (beer 0.05, opaque 0.04, wine 0.12, spirits 0.40)",
          benchmark: "OECD beer 15-40, spirits 25-50 USD per LPA"}
ssbs: {usd_per_liter_by_sugar_band: "one value per band", share_of_retail: "per liter",
       benchmark: "UK 0.23-0.30 (sugar-graduated); Mexico ~0.05 (flat)"}
sugar: {usd_per_kg: "rate / fx", note: "state the industrial-use exemption"}
implied_revenue: "rate x volume / GDP; prevalence-based volumes are upper bounds (DQ 'estimate')"
collected_revenue: "by product group where published; else all non-fuel excise, flagged"
inversions: "wine vs beer per LPA; local vs imported; tiers that reward lower quality"
metrics_to_fill: [R5, R6, R7, R8, I4, I5, I6, I7, I8, I9]   # Country_Metrics IDs
extra_metrics: ["H1 sugar excise, USD per kg", "H2 indexation clause present on these lines (yes/no,
                with ref)"]
```

## 5. STAGES

```yaml
P0_common_header:  "Only if no Common sheet exists yet. See master section 2B."
P1_history:        "Master A1 for tobacco, alcohol, SSB and sugar lines."
P2_schedule:       "Master A2 for the 13 codes: every line, band and tier; refs; trajectory."
P3_measure:        "Master A3 for the 13 codes: section 4 measures; populate the part workbook."
P4_part_findings:  "Level 1 for health: key metrics, 4-6 findings, inconsistencies, peer metrics,
                    issues for dialogue. List framework items (indexation, stamps) for Part 5."
P5_peer_review:    "Master A5 with PR-H ids. Quote every source. Write statuses into the workbook."
P6_report:         "Assemble the part report (section 7) and final part workbook (section 6)."
```

## 6. PART WORKBOOK

```yaml
file: "<ISO3>_excise_part3_health.xlsx"
sheets: [Common, Summary, Country_Metrics, Std_Schedule, Std_Schedule_Code, Parameters, DQ_Scale,
         Sources, Register]
rows: "Std_Schedule and Std_Schedule_Code hold ONLY the 13 health codes, with part = '3 Health'.
       One Std_Schedule_Code row per rate component: each ABV band, sugar band, tier and origin;
       two rows for a higher-of hybrid (same option_group)."
parameters_added: [retail_prices_by_product, tobacco_users, sticks_per_day, alcohol_lpa_per_adult,
                   formal_share, ssb_volumes, sugar_volumes, health_collections]
checks: "Recalculate with zero errors. Row ids unique (<ISO3>-H-001 ...). Lines shown at the last
         legible rate carry record_type 'last_legible'."
```

## 7. PART REPORT OUTLINE

```text
1  Scope, sources and the laws that set health excises
2  Level 1 (health): key metrics; findings; inconsistencies; peer metrics; issues for dialogue
3  Level 2: standardized rows T1-T3, A1-A5, S1-S4, I3
4  Level 3: tobacco, alcohol, SSB and sugar lines; trajectory since 2020
5  Level 4: health parameters and computations; cross-check with collections
6  Register PR-H (sources, quotes, statuses)
7  Framework items for Part 5 (indexation clauses, stamps, other cross-cutting rules)
```

## 8. CHECKS BEFORE HANDING TO PART 5

- All 13 codes present in Std_Schedule, with coverage 0, 1 or 2.
- Every band and tier in the law has its own Std_Schedule_Code row.
- Common sheet unchanged from the first part run.
- Every line has a PR-H reference; every "verified" has a quote.

## 9. PROMPT TEMPLATES

```text
[P0] (First part only.) Create the common header for {COUNTRY} at {CUT_OFF} as the Common sheet.
[P1] Run stage P1 for Part 3 (health): the history of tobacco, alcohol, SSB and sugar lines from
     {START} to {CUT_OFF}. Quote rates as printed. Quote any indexation clause. If you cannot fetch
     a source, give me the links list and the .bat file.
[P2] Run stage P2: the Level 3 health schedule (13 codes), every band and tier, with refs and the
     trajectory.
[P3] Run stage P3: fill the health parameters, compute the measures, and populate the part workbook.
     {IF COMPARATOR: use the Part 3 workbook for {COMPARATOR}.}
[P4] Run stage P4: Level 1 for health, and the list of framework items for Part 5.
[P5] Run stage P5: the PR-H register; apply corrections; write statuses into the workbook.
[P6] Run stage P6: assemble the Part 3 report and deliver the final part workbook.
```
