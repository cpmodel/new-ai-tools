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
11. **Trade ratios for the DSF.** The WEO extract has exports and imports for country groups only; add WDI NE.EXP.GNFS.ZS and NE.IMP.GNFS.ZS to the dashboard (Appendix C) and use them in `integrated.dsf_baseline`.
12. **CPIA, reserves, remittances** in the dashboard for the DSF Composite Indicator (today: medium capacity for all).
13. **External share of public debt** from IDS (PPG external stock / gross debt) instead of the income-group default.
14. **R pipeline parity** for Appendix C and D is written but unverified (R was not available); run `Rscript build_global_indicators.R` and diff against the Python output.
15. **Pricing beyond power:** road fuels and industry in PriceFeasibility would raise the revenue estimate.
16. **Programme pricing interaction:** the tariff rise (F1) has an economic cost that the macro models do not see.

17. **Data issues found by the integrated run.** Solomon Islands: NMIM budget need 91% of GDP in 2030 (check NMIM land-use / non-CO2 inputs). Cabo Verde: WDI coal rents 17% of GDP with no fuel exports (ignored by `calibration.shared`; flag to the dashboard owner). Libya: rents of 61% of GDP falling 80% cannot be stabilised by the fiscal rule.
18. **Pitch deck** covers the models, not yet the integrated run; add a slide when the pilots are calibrated.

## DIGNAD-R Excel (proof of concept, 2026-10-20)

19. Move the four augmented inputs (public investment, return on public capital, hand-to-mouth share, grants) into the dashboard's Appendix C with real sources; then DIGNAD-R Excel reads the dashboard columns only.
20. A second fiscal instrument (consumption tax) to close the gap to Python DIGNAD-lite where the rule has to tighten.
21. Energy savings and per-country overrides on the Assumptions sheet.
22. Wire `Outputs` into the country tool or the dashboard (Appendix D) as a live module.

## Other open items

- Country calibrations: `calibration.py` now sets shared settings from the dashboard; macro-model input block in the dashboard still to add (CPIA, reserves, remittances, projected debt service, net foreign assets, labour share, Findex, public capital stock, disaster damages, import share).
- Check LIC-DSF-lite coefficients, thresholds and shock sizes against the current Guidance Note and template.
- Recurrent disasters in DIGNAD-lite.
- Carbon pricing revenue from CPAT; multi-country spillovers in GIMF-lite.
- Open the pitch deck in PowerPoint (checked only in LibreOffice).
