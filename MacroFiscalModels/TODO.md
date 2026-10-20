# TODO: MacroFiscalModels

Open work, in priority order. Move items to `CAVEATS.md` when done.

## Later: effectiveness of mitigation public investment in DIGNAD-lite (agreed 2026-10-20)

1. **Energy-savings channel.** Feed NMIM's net energy savings (`sav` from `nmim_bridge.programme`) into DIGNAD-lite, as GIMF-lite already does, so investment that replaces fossil capacity earns its real return (fuel and O&M savings) rather than an infrastructure return. Most important of the four for mitigation.
2. **Sector-weighted return on public capital.** Set `Rg` from NMIM's sector mix of the budget need: higher for grids, transmission, public transport and access; low for replacement generation; near zero domestically for land use and non-CO2 (the benefit is global or outside GDP).
3. **Absorptive-capacity costs.** Unit costs of public investment rise with the speed of the scale-up, as in the full DIG / DIGNAD, so front-loaded programmes are not overstated.
4. **Efficiency-return switch.** Option to hold the output elasticity of public capital fixed instead of its return, so low efficiency comes with a higher marginal return on scarce capital (Berg et al. 2019, "Some misconceptions about public investment efficiency and growth").

## Integration (agreed 2026-10-20)

5. **Option 2, later: one model with switches.** GIMF-lite as the core, with DIGNAD-lite's financing (concessional / commercial / grants), efficiency and disaster blocks as switches.
6. **Option 3, later: DIGNAR-style resource block** for resource-rich countries (Melina, Yang and Zanna 2016), using the dashboard's fossil-rent indicators: stranded fossil revenue as a fiscal risk.
7. **Done 2026-10-20 (see CAVEATS).** ~~Spending versus capital in the bridge.~~ Non-CO2, land-use and other energy CO2 lines are annual spending in NMIM but enter the macro models as productive public capital (Liberia 90%, Kenya 53% of incremental investment 2025-2035). Pass them as government spending, or as capital with a zero domestic return.
8. **Efficiency of public investment, more fully.** Beyond the PEFA-mapped `s`: PIMA scores where they exist, absorptive capacity (item 3), the reform scenario (closing up to two-thirds of the gap, IMF 2015), and execution rates (BOOST) as a separate leak.
9. **Tariff revenue versus public return.** Utility (P2) investment is repaid by tariffs and enters the bridge as private investment, so state-owned utility debt and tariff shortfalls are outside public debt. Model the utility explicitly: tariff revenue, cost recovery (NMIM's index), utility debt as a contingent liability in the DSF (PEFA PI-10.3), and the distinction between financial return (tariffs, a transfer from users) and economic return (fuel savings, reliability, access).

## From the integration (2026-10-20)

10. **Energy savings in DIGNAD-lite and DIGNAR-lite** (item 1) now also matters for consistency: GIMF-lite countries get the savings channel, the others do not.
11. **Trade ratios for the DSF.** The WEO extract has exports and imports for country groups only; add WDI NE.EXP.GNFS.ZS and NE.IMP.GNFS.ZS to the dashboard (Appendix C) and use them in `integrated.dsf_baseline`. Still open (2026-10-20): no country trade series in any repository source, and the WDI API is not reachable from the cloud session; needs a manual download into `raw_data/`. Until then exports (DSF denominators) and imports (reserve cover in the CI) are assumptions.
12. **Done 2026-10-20 (see CAVEATS).** ~~CPIA, reserves, remittances in the dashboard for the DSF Composite Indicator.~~ Dashboard v1.3 Appendix C, from the Climate Fiscal Risk Dashboard v1.5; `integrated.dsf_capacity` computes the CI where CPIA exists (66 assessed countries: 30 weak, 25 medium, 11 strong).
13. **Done 2026-10-20.** ~~External share of public debt from IDS instead of the income-group default.~~ Dashboard v1.3 Appendix C; used in `integrated.dsf_baseline` (default only where IDS has no stock).
14. **R pipeline parity** for Appendix C, D and E is written but unverified (R was not available, again on 2026-10-20 for the v1.3 additions); run `Rscript build_global_indicators.R` and diff against the Python output.
15. **Pricing beyond power:** road fuels and industry in PriceFeasibility would raise the revenue estimate.
16. **Programme pricing interaction:** the tariff rise (F1) has an economic cost that the macro models do not see.

17. **Data issues found by the integrated run.** Solomon Islands: NMIM budget need 91% of GDP in 2030 (check NMIM land-use / non-CO2 inputs). Cabo Verde: WDI coal rents 17% of GDP with no fuel exports (ignored by `calibration.shared`; flag to the dashboard owner). Libya: rents of 61% of GDP falling 80% cannot be stabilised by the fiscal rule.
18. **Pitch deck** covers the models, not yet the integrated run; add a slide when the pilots are calibrated.

## DIGNAD-R Excel (proof of concept, 2026-10-20)

19. **Partly done 2026-10-20.** Public investment is in the dashboard's Appendix C (IMF ICSD, to 2019) and read by `calibration.py` and DIGNAD-R Excel v0.3. Still defaults: return on public capital (no cross-country source), grants (needs OECD CRS or IMF GFS), hand-to-mouth share (Findex account ownership is now in Appendix C as a candidate proxy, not used: 1 - account ownership would halve the shares and drive the multipliers, so it needs a judgement first).
20. A second fiscal instrument (consumption tax) to close the gap to Python DIGNAD-lite where the rule has to tighten.
21. Energy savings and per-country overrides on the Assumptions sheet.
22. **Done 2026-10-20 (see `IntegratedMFC/CAVEATS.md`).** ~~Wire `Outputs` into the country tool or the dashboard (Appendix D) as a live module.~~ DIGNAD-R runs live for the selected country on the Integrated MF-C tool's Macro sheet (all 186 countries), with the integrated run's paths beside it.

## Execution efficiency, light CBA and financing (agreed 2026-10-20)

From the review in `cpmodel/CPAT-CDPM`, `reports/Public spending efficiency and CBA.md` (sections 4-6, 9). Aim: place climate and energy spending in the wider fiscal picture, not a full allocative-efficiency analysis. Frontier work (Herrera et al., World Bank Fiscal Policy unit) is used as an input only. Items already listed above are cross-referenced, not repeated.

Order: first 23, 36, 32 and 29 (things that are wrong or unused today), then 27-28 and 31-34, then the rest.

**A. Execution efficiency** (extends items 3, 4 and 8)

23. **BOOST execution rate as a delay, separate from `s`.** The dashboard's `capex_execution_rate_pct` is shown but unused. Unspent appropriations stay in the budget; benefits follow actual disbursement (`I_act = e · (I_plan + backlog)`). Apply `s` once only ("money in, assets out"; report section 5, Step 3).
24. **Recurrent cost and maintenance.** O&M per dollar of public capital; if it is unfunded, `delta_g` rises. Report the maintenance gap.
25. **Cost-overrun uplift** where there are no PIM data (HM Treasury optimism-bias values, reduced on evidence), and the PIM reform scenario (close up to two-thirds of the gap, IMF 2015) as an option (item 8).
26. **Frontier interface.** An input slot for World Bank spending-efficiency scores (Herrera et al.; PFR Fundamentals: Spending Efficiency) as an alternative source of `s`, and an export of each country's `s` and its source so it can be compared with the Bank's scores. No frontier estimation here.

**B. Light cost-benefit analysis**

27. **Two numbers per measure** (NMIM line or CDPM transformation): BCR at the World Bank 6% rate with the 2024 shadow price of carbon, and value per net public dollar (benefits over the net cost to the budget). Charts show reference lines at BCR 1, about 1.2 (cost of public funds) and about 1.5-2 (rough practice under tight budgets).
28. **One VSL rule for CPAT, NMIM and CDPM.** VSL income elasticity as an input (CPAT hard-wires 1; `VslEl` is #REF!), with a floor of 20 × GNI per capita, and two variants reported: theoretical (elasticity 1, about 160 × income) and Bank/practice (elasticity 1.5, reference case, Robinson et al. 2019). NMIM lens (c) and CDPM to use CPAT's per-GJ air-pollution values instead of the LAC-average IMF factors.
29. **CDPM into the chain.** Fix CDPM discounting (power capex and opex are pre-discounted by NemoMod, other lines undiscounted) and map each transformation to a payer type (P1-P5), so CDPM costs and co-benefits can pass through `nmim_bridge` to the macro models.
30. **Sector benchmark table.** Published BCRs for health, education, roads (maintenance) and water, harmonised to one discount rate and VSL, shown next to the climate and energy BCRs. Status of each figure flagged (the report's sources were checked against abstracts only).

**C. Financing** (extends item 9)

31. **Financing route per programme, not per debt stock:** grant, concessional, commercial, PPA (P3), utility (P2). Use the dashboard's `ids_avg_maturity_years` and `mitigation_finance_received_pct_gdp` (both unused today) for terms and the concessional envelope.
32. **Viability gap from the private NPV at the country WACC** (`wacc_solar_real_aftertax_pct`), replacing NMIM's flat 30% of P4 where the private NPV can be computed.
33. **PPA and utility register as contingent liabilities in LIC-DSF-lite:** PV of take-or-pay payments and utility debt, scaled by PEFA PI-10.3 (`pefa_pi103_contingent_liabilities`, unused today).
34. **Cash-flow test by route:** annual debt service against energy savings (the small-island trap, report section 6), with guarantees as a lever that lowers the WACC.
35. **Pricing as financing beyond power** (item 15): road fuels and fossil-subsidy reform.

**D. Macro-fiscal, not yet done** (items 1, 2, 10-13 above stay as they are)

36. **Imported-fuel savings as a terms-of-trade and current-account gain** in GIMF-lite, not a value-added gain; with item 1, this closes most of the gap between GIMF-lite and DIGNAD-lite (median +7.7% against +0.7% of GDP by 2050).
37. **Budget share of energy savings.** Where fuel is subsidised or the utility is state-owned, part of the savings returns to the budget (`rev_extra` / `rev`); report it next to the BCR, not inside it.
38. **Pilot calibrations** (KEN, SLE, UGA, LBR) to replace income-group placeholders for `Rg`, grants, constrained share and the financing mix; then the CPAT revenue link.

## Speed of the integrated run (2026-10-20)

`integrated.py --all` takes about 10 minutes (2-9 s a country). Profiling one country: almost all
of the time is spent building the Newton Jacobian in `pf_solver.jacobian`, one entry at a time into
a `lil_matrix`; the model equations themselves are cheap. In order:

39. **Done 2026-10-20.** ~~Assemble the Jacobian in one go~~ (collect row, column, value arrays; build a `csc_matrix`
    once). Tested on KEN, LBR, AGO, USA: 6-13x faster, results identical (max difference 0.0).
    About a minute for all countries. No change to the method. Done: `--all` 1m41s (from about 10 min),
    results byte-for-byte identical; 3m10s with the cross-check model added.
40. **One pass, no retries, at this stage**: start each run with the fiscal-rule speed saved in
    `results/integrated_settings.csv` from the last run instead of trying 0.05, 0.2, 0.4 in turn
    (Liberia-type countries pay for up to three full solves); start the programme solve from the
    baseline path; reuse the Jacobian across Newton steps (modified Newton). Optional "linear"
    mode: one Newton step from the steady state (first-order solution, no iterations), checked
    against the full solve before it is used for anything shipped.
41. **Parallel and incremental**: run countries in parallel (`multiprocessing`); re-run only
    countries whose inputs changed (hash of panel row, NMIM paths, PriceFeasibility row).
42. **One-click run**: a `.bat` (as in the dashboard build) for integrated run -> dashboard panel ->
    Integrated MF-C tool. Not a VBA macro: the models are Python, a macro would only shell out to
    them, and `.xlsm` files are often blocked; the Excel side (DIGNAD-R on the Macro sheet) is
    already one recursive pass with no iteration and recalculates instantly.

## Other open items

- Country calibrations: `calibration.py` now sets shared settings from the dashboard; macro-model input block in the dashboard still to add (CPIA, reserves, remittances, projected debt service, net foreign assets, labour share, Findex, public capital stock, disaster damages, import share).
- Check LIC-DSF-lite coefficients, thresholds and shock sizes against the current Guidance Note and template.
- Recurrent disasters in DIGNAD-lite.
- Carbon pricing revenue from CPAT; multi-country spillovers in GIMF-lite.
- Open the pitch deck in PowerPoint (checked only in LibreOffice).
