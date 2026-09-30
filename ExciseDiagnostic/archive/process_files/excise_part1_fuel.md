# EXCISE DIAGNOSTIC - PART 1: FUEL (SECTORAL RUN)

```yaml
part_id: "1 Fuel"
file: excise_part1_fuel.md
spec_version: "1.1"                     # read together with excise_diagnostic_process.md (MASTER)
standard_schedule_version: "1.0"
codes: [E1, E2, E3, E4, E5, E6]
register_prefix: "PR-F"
run_order: "Any order with Parts 2-4. Part 5 (overarching) runs after all four."
outputs:
  report: "<ISO3>_excise_part1_fuel_report (docx or md)"
  workbook: "<ISO3>_excise_part1_fuel.xlsx"
joint_paper: "Codes, measures and tables line up with Part 2 (vehicles), so Parts 1 and 2 can be
              stitched into a joint fuel and vehicle tax diagnostic."
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
code  standard line                       what to include
E1    Petrol                              gasoline, premium and regular; ethanol blends
E2    Diesel, road                        automotive gas oil; note power-generation or mining exemptions
E3    Kerosene                            illuminating and household kerosene
E4    Aviation fuel                       Jet A1, avgas; note airline or supply-contract exemptions
E5    Heavy fuel oil / other gas oils     fuel oil, marine and industrial gas oils; LPG and other
                                          fuels not listed elsewhere (note the product)
E6    Lubricants                          motor and industrial lubricating oils
```

In scope for this part:
- Fuel excise in the excise law (coverage 1).
- Fuel levies outside the excise law that work as excises, recorded under the fuel's code with
  coverage 2: road fund levies, regulator and storage levies (for example, a petroleum refining or
  storage company levy), price-stabilization levies, carbon or environmental levies on fuel.
- Exemptions and nil rates by user (power generation, airlines, mining agreements, diplomats).

Out of scope (note, do not count):
- Customs duty and VAT on petroleum. Record them in a short "other taxes on fuel" note, because
  they matter for the pump-price build-up, but they are not excise.
- Vehicle charges (Part 2). Plastics (Part 4).

Framework items to hand to Part 5 (do not analyze here): the general legal framework, currency
notation, indexation rules that apply beyond fuel, stamps, general exemption regimes, total excise
collections.

## 2. INPUTS

```yaml
common_header: "Common sheet (master section 2B)"
legal_texts: "Every provision on fuel lines (typically tariff heading 27.10) from 2020 to the cut-off;
              laws creating fuel levies outside the excise law"
pricing: "Pump-price build-up or pricing formula; is excise a fixed rate or a residual in the formula?"
data:
  pump_prices: "Petrol, diesel, kerosene at or near the cut-off (regulator; GlobalPetrolPrices)"
  volumes: "Annual sales or imports by product (regulator, national oil company, customs)"
  collections: "Fuel excise and fuel-levy collections, latest two years (finance ministry,
                revenue authority)"
  exemptions: "Volumes or shares exempt, where published"
```

## 3. DATA STEP (WHEN NEEDED)

If any input in section 2 cannot be fetched, Claude stops and produces
`datastep_<ISO3>_part1_fuel_links.md` and `datastep_<ISO3>_part1_fuel.bat` using the template in master
section 12. Typical fuel downloads: the Finance Acts (shared with other parts), the regulator's price
notices and sales statistics, the national oil company's import plan, and the finance ministry's fiscal
tables. Shared legal texts keep the same file names in every part, so the .bat skips them if already
downloaded.

## 4. PART-SPECIFIC MEASURES

```yaml
usd_per_liter: "rate / fx"
share_of_pump_price: "excise per liter / pump price per liter (state the price date; flag prices
                      taken before a rate change)"
implied_revenue: "sum over products of rate x volume / GDP; add fuel levies separately"
collected_revenue: "fuel excise and fuel levies collected / GDP, latest two years"
fuel_intensity: "liters of petrol and diesel per USD 1,000 of GDP"
differentials:
  - "diesel vs petrol (per liter)"
  - "kerosene vs diesel (blending incentive)"
  - "exempt users (aviation, power, mining): share of volume if known"
pricing_regime: "fixed specific rate | residual in a pricing formula | ad valorem"
metrics_to_fill: [R3, R4, I1, I2, I3, B1]          # Country_Metrics IDs
extra_metrics: ["F1 fuel levies outside the excise law (% of GDP)", "F2 kerosene-diesel differential",
                "F3 exempt share of fuel volume"]
```

## 5. STAGES

```yaml
P0_common_header:
  when: "Only if no Common sheet exists yet for this country."
  do: "Fill master section 2B: cut-off, currency and notation, FX, GDP, CPI, legal instruments,
       shared sources. Save as the Common sheet."
P1_history:        "Master A1, fuel lines and fuel levies only."
P2_schedule:       "Master A2 for E1-E6: every fuel line and levy, five fields, refs, trajectory."
P3_measure:        "Master A3 for E1-E6: measures in section 4; populate the part workbook."
P4_part_findings:  "Level 1 for fuel: key metrics, 3-5 findings, inconsistencies, peer metrics
                    (section 9 of the master), issues for dialogue. List framework items for Part 5."
P5_peer_review:    "Master A5 with PR-F ids. Quote every source. Write statuses into the workbook."
P6_report:         "Assemble the part report (section 7) and final part workbook (section 6)."
```

## 6. PART WORKBOOK

```yaml
file: "<ISO3>_excise_part1_fuel.xlsx"
sheets: [Common, Summary, Country_Metrics, Std_Schedule, Std_Schedule_Code, Parameters, DQ_Scale,
         Sources, Register]
rows: "Std_Schedule and Std_Schedule_Code hold ONLY codes E1-E6, with part = '1 Fuel'. Fuel levies
       outside the excise law are separate rows with coverage_code 2."
parameters_added: [pump_price_petrol, pump_price_diesel, pump_price_kerosene, volume_petrol,
                   volume_diesel, volume_kerosene, fuel_collections]
checks: "Recalculate with zero errors. Row ids unique (<ISO3>-F-001 ...)."
```

## 7. PART REPORT OUTLINE

```text
1  Scope, sources and pricing regime
2  Level 1 (fuel): key metrics; findings; inconsistencies; peer metrics; issues for dialogue
3  Level 2: standardized rows E1-E6
4  Level 3: fuel lines and levies; rate trajectory since 2020; other taxes on fuel (note)
5  Level 4: fuel parameters and computations; cross-check with collections
6  Register PR-F (sources, quotes, statuses)
7  Framework items for Part 5
```

## 8. CHECKS BEFORE HANDING TO PART 5

- All six codes present in Std_Schedule, with coverage 0, 1 or 2.
- Common sheet unchanged from the first part run.
- Every line has a PR-F reference; every "verified" has a quote.
- Collected and implied fuel revenue both reported, or the gap marked data_gap.

## 9. PROMPT TEMPLATES

```text
[P0] (First part only.) Using the master specification and this part file, create the common header
     for {COUNTRY} at {CUT_OFF} and save it as the Common sheet.
[P1] Run stage P1 for Part 1 (fuel): the legislative history of fuel lines and fuel levies from
     {START} to {CUT_OFF}. Quote rates as printed. If you cannot fetch a source, give me the links
     list and the .bat file.
[P2] Run stage P2: the Level 3 fuel schedule (E1-E6), with refs and the rate trajectory.
[P3] Run stage P3: fill the fuel parameters, compute the measures, and populate the part workbook.
     {IF COMPARATOR: use the Part 1 workbook for {COMPARATOR}. ELSE: no comparator.}
[P4] Run stage P4: Level 1 for fuel, and the list of framework items for Part 5.
[P5] Run stage P5: the PR-F register; apply corrections; write statuses into the workbook.
[P6] Run stage P6: assemble the Part 1 report and deliver the final part workbook.
```
