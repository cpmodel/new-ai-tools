# Cap-based ETS price: design options (for decision)

Status: the MVP's new ETS uses the carbon price path as its permit price (user decision v0.11: "as effective as a carbon tax"). Legacy CPAT derives the permit price from a cap instead (legacy Mitigation rows 1797-1966).

## How legacy does it

1. **Inputs (MTInputs 84-115):** the cap as a change relative to baseline emissions in the covered sectors (start and target year, then constant). Also the auction share, covered sectors, and an override switch.
2. **Needed reduction:** the cap path minus the contribution of other policies (the carbon tax trajectory x a weighted semi-elasticity).
3. **Price:** needed reduction / adjusted ETS semi-elasticity. The semi-elasticities are per sector, per $/tCO2: power -0.0028, transport -0.0028, residential -0.0039, industry -0.0055. They are weighted by baseline emissions and scaled by allocation effectiveness (auctioned 1.0, output-based 0.5).
4. **Iteration:** a convergence step (factor 0.5) against the realised emissions; the "Max ETS Error" row shows the residual.

## Why it cannot simply be dragged into the MVP

The price needs baseline covered emissions, which come from scenario 1. In the MVP every scenario is a column group with the same formulas, so scenario 1's ETS price row would reference scenario 1's own CO2. Excel tracks precedents by range, not by value, so it flags a circular reference even when the ETS is off. Legacy avoids this with a separate baseline section computed before the policy section.

## Options

| Option | How | Pros | Cons |
|---|---|---|---|
| A. Baseline as data (recommended) | A "baseline covered emissions" row on Settings, refreshed by copy and paste-values from scenario 1 (one documented step; a check flags when it is stale) | No circularity; transparent; keeps one formula per row | One manual refresh after baseline changes |
| B. Legacy semi-elasticity method on the scenario's own pre-ETS path | Price from the cap and the sector semi-elasticities, applied to this scenario's emissions before the ETS year (last pre-ETS year grown with GDP) | Fully automatic; no cross-scenario reference | Approximation; the cap is not exactly met; still needs the semi-elasticities (legacy values, Egypt) |
| C. Excel iterative calculation | Enable iterative calculation; the price converges on the cap | Exact cap | Hidden iteration, fragile, not auditable; LibreOffice behaves differently |
| D. Keep the current rule | Permit price = carbon price inputs (override) | Simple, already built and tested | No cap |

Recommendation: A, with D kept as the override (`D_ETSPriceOverride` = Yes, which matches the legacy switch). The legacy semi-elasticities would be stored as data for the price formula, and a "difference from cap" row would show how far the realised emissions are from the cap, as legacy does.
