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
7. **Spending versus capital in the bridge (high priority).** Non-CO2, land-use and other energy CO2 lines are annual spending in NMIM but enter the macro models as productive public capital (Liberia 90%, Kenya 53% of incremental investment 2025-2035). Pass them as government spending, or as capital with a zero domestic return.
8. **Efficiency of public investment, more fully.** Beyond the PEFA-mapped `s`: PIMA scores where they exist, absorptive capacity (item 3), the reform scenario (closing up to two-thirds of the gap, IMF 2015), and execution rates (BOOST) as a separate leak.
9. **Tariff revenue versus public return.** Utility (P2) investment is repaid by tariffs and enters the bridge as private investment, so state-owned utility debt and tariff shortfalls are outside public debt. Model the utility explicitly: tariff revenue, cost recovery (NMIM's index), utility debt as a contingent liability in the DSF (PEFA PI-10.3), and the distinction between financial return (tariffs, a transfer from users) and economic return (fuel savings, reliability, access).

## Other open items

- Country calibrations: `calibration.py` now sets shared settings from the dashboard; macro-model input block in the dashboard still to add (CPIA, reserves, remittances, projected debt service, net foreign assets, labour share, Findex, public capital stock, disaster damages, import share).
- Check LIC-DSF-lite coefficients, thresholds and shock sizes against the current Guidance Note and template.
- Recurrent disasters in DIGNAD-lite.
- Carbon pricing revenue from CPAT; multi-country spillovers in GIMF-lite.
- Open the pitch deck in PowerPoint (checked only in LibreOffice).
