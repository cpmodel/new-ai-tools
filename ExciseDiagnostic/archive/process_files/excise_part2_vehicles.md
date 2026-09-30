# EXCISE DIAGNOSTIC - PART 2: VEHICLES (SECTORAL RUN)

```yaml
part_id: "2 Vehicles"
file: excise_part2_vehicles.md
spec_version: "1.1"                     # read together with excise_diagnostic_process.md (MASTER)
standard_schedule_version: "1.0"
codes: [E8, E9]
register_prefix: "PR-V"
run_order: "Any order with Parts 1, 3 and 4. Part 5 (overarching) runs after all four."
outputs:
  report: "<ISO3>_excise_part2_vehicles_report (docx or md)"
  workbook: "<ISO3>_excise_part2_vehicles.xlsx"
joint_paper: "Codes, measures and tables line up with Part 1 (fuel), so Parts 1 and 2 can be
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
code  standard line                            what to include
E8    Vehicles: annual circulation             annual road-use, circulation or licensing levies;
                                               per vehicle-year, by class (car, truck, bus,
                                               motorcycle, three-wheeler, vessel)
E9    Vehicles: first registration / import    excise on vehicle imports (by engine size, age, value
                                               or emissions); luxury-vehicle excises; first-
                                               registration charges and age surcharges that work
                                               as taxes
```

In scope for this part:
- Vehicle excise in the excise law (coverage 1).
- Registration, licensing and circulation charges set in other laws or regulations when they are
  tax-like (coverage 2): they raise revenue well above the cost of the service.
- Differentiation by class, engine size, age, value, fuel type or emissions; exemptions (diplomats,
  government, investment incentives).

Out of scope (note, do not count):
- Customs duty and VAT on vehicles: record them in an "other taxes on vehicles" note, because they
  shape the total charge at import, but they are not excise.
- Fees that only recover a service cost (number plates, inspection), unless they are tax-like.
- Fuel (Part 1).

Framework items to hand to Part 5: general exemption regimes, stamps, indexation rules, currency.

## 2. INPUTS

```yaml
common_header: "Common sheet (master section 2B)"
legal_texts: "Excise law and Finance Act provisions on vehicles; road, traffic or licensing laws
              and regulations that set registration or circulation charges; recent amendments"
data:
  fleet: "Registered vehicles by class (latest year)"
  flows: "New registrations or vehicle imports by class, with typical CIF values"
  collections: "Vehicle excise and registration or circulation levy collections, latest two years"
  import_rules: "Age limits or age-based surcharges on imported vehicles"
```

## 3. DATA STEP (WHEN NEEDED)

If any input in section 2 cannot be fetched, Claude stops and produces
`datastep_<ISO3>_part2_vehicles_links.md` and `datastep_<ISO3>_part2_vehicles.bat` using the template
in master section 12. Typical vehicle downloads: the Finance Acts (shared), the road or transport
authority's fee schedule, the licensing regulations and the customs vehicle-import statistics.

## 4. PART-SPECIFIC MEASURES

```yaml
representative_vehicles:          # use the same three in every country so results compare
  small_car:  "1.3 liter petrol, 8 years old at import, CIF USD 5,000"
  suv:        "3.0 liter diesel, 3 years old at import, CIF USD 30,000"
  motorcycle: "125 cc, new, CIF USD 1,000"
usd_per_vehicle_at_import: "all excise-type charges at import and first registration for each
                            representative vehicle (E9)"
usd_per_vehicle_year: "annual circulation or licensing levy for each representative vehicle (E8)"
share_of_value: "E9 charges / CIF value, for each representative vehicle"
implied_revenue: "E8: fleet x annual levy; E9: annual imports x charge; each / GDP"
collected_revenue: "vehicle excise and levies collected / GDP, latest two years"
differentiation: "does the charge rise with engine size, age, value or emissions? which one binds?"
metrics_to_fill: ["V1 vehicle excise and levies, collected (% of GDP)",
                  "V2 vehicle excise and levies, implied (% of GDP)",
                  "V3 charge at import, representative vehicles (USD)",
                  "V4 annual levy, representative vehicles (USD per year)",
                  "V5 basis of differentiation (text)"]
```

## 5. STAGES

```yaml
P0_common_header:  "Only if no Common sheet exists yet. See master section 2B."
P1_history:        "Master A1, vehicle lines and vehicle levies only."
P2_schedule:       "Master A2 for E8-E9: every vehicle charge, five fields, refs, trajectory."
P3_measure:        "Master A3 for E8-E9: section 4 measures; populate the part workbook."
P4_part_findings:  "Level 1 for vehicles: key metrics, 3-5 findings, inconsistencies (for example,
                    charges that fall as vehicle age rises), issues for dialogue. List framework
                    items for Part 5."
P5_peer_review:    "Master A5 with PR-V ids. Quote every source. Write statuses into the workbook."
P6_report:         "Assemble the part report (section 7) and final part workbook (section 6)."
```

## 6. PART WORKBOOK

```yaml
file: "<ISO3>_excise_part2_vehicles.xlsx"
sheets: [Common, Summary, Country_Metrics, Std_Schedule, Std_Schedule_Code, Parameters, DQ_Scale,
         Sources, Register]
rows: "Std_Schedule and Std_Schedule_Code hold ONLY codes E8-E9, with part = '2 Vehicles'.
       One Std_Schedule_Code row per vehicle class and rate component."
parameters_added: [fleet_by_class, imports_by_class, representative_vehicle_values, vehicle_collections]
checks: "Recalculate with zero errors. Row ids unique (<ISO3>-V-001 ...)."
```

## 7. PART REPORT OUTLINE

```text
1  Scope, sources and the laws that set vehicle charges
2  Level 1 (vehicles): key metrics; findings; inconsistencies; peer metrics; issues for dialogue
3  Level 2: standardized rows E8-E9
4  Level 3: vehicle charges by class; trajectory since 2020; other taxes on vehicles (note)
5  Level 4: vehicle parameters, representative-vehicle computations, cross-check with collections
6  Register PR-V (sources, quotes, statuses)
7  Framework items for Part 5
```

## 8. CHECKS BEFORE HANDING TO PART 5

- Both codes present in Std_Schedule, with coverage 0, 1 or 2.
- Representative-vehicle results shown for all three vehicles.
- Common sheet unchanged from the first part run.
- Every line has a PR-V reference; every "verified" has a quote.

## 9. PROMPT TEMPLATES

```text
[P0] (First part only.) Create the common header for {COUNTRY} at {CUT_OFF} as the Common sheet.
[P1] Run stage P1 for Part 2 (vehicles): the history of vehicle excises and levies from {START}
     to {CUT_OFF}. Quote rates as printed. If you cannot fetch a source, give me the links list
     and the .bat file.
[P2] Run stage P2: the Level 3 vehicle schedule (E8-E9) by class, with refs and the trajectory.
[P3] Run stage P3: fill the vehicle parameters, compute the representative-vehicle measures, and
     populate the part workbook. {IF COMPARATOR: use the Part 2 workbook for {COMPARATOR}.}
[P4] Run stage P4: Level 1 for vehicles, and the list of framework items for Part 5.
[P5] Run stage P5: the PR-V register; apply corrections; write statuses into the workbook.
[P6] Run stage P6: assemble the Part 2 report and deliver the final part workbook.
```
