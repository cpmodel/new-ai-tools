# Legacy ETS algorithm and a benchmark-based design for the MVP (v0.1)

Source: `CPAT 1.0pre_456_NoPropData.xlsb`, sheet Mitigation rows 1797–1969 and 2244–2425, sheet Manual inputs O15:AB15, and VBA Module1 (`OverrideETSFast`, `SolveFast`). Row numbers are legacy rows. Formulas that were not stored in the dump have been inferred from the row labels and the cached values; those are marked *(inferred)*.

## 1. Legacy: what goes in

| Rows | Item |
|---|---|
| 1800–1809 | Apply ETS; cap change at start and at target (% vs baseline); cap continuation; start and target years; auction proportion at start and at target; continuation; override switch |
| 1811–1813 | Price volatility 0.42, impact 0.5, policy risk 0.2 (risk premium inputs) |
| 1815–1836 | Sector table: included flag, baseline emissions, semi-elasticity of emissions to the carbon price (power −0.00279, transport −0.00283, residential −0.00386, industry −0.00549, CBAM −0.001 per $/t), generic proportions |
| 1843 | Relative price of auctioned ETS vs tax for the same emissions target: 1.1 |
| 1846–1847 | Relative effectiveness, hardcoded: auctioning **1.0**, output-based allocation (OBA/OBR) **0.5** |

The allocation choice is a single **time path of the auctioned proportion** (rows 1863–1864: auctioned share *a*, OBA share 1 − *a*). The same share applies to every covered sector.

## 2. Legacy: the fast approximate model (no iteration)

1. **Covered semi-elasticity** (row 1845): the coverage-weighted average of the sector semi-elasticities, ε = −0.00288 in the cached run. Row 1844 is the all-sector average used for carbon tax, −0.00371.
2. **Reduction left for the ETS** (rows 1858–1860): needed reduction from the cap minus the contribution of other policies. Other policies count as an effective carbon price × ε_tax (row 1859, e.g. $50/t × −0.00371 = −18.6%).
3. **Effectiveness of the allocation mix** (row 1866): e = a × 1.0 + (1 − a) × 0.5.
4. **Adjusted semi-elasticity** (row 1867): ε_adj = ε × e. With full OBA, −0.00288 × 0.5 = −0.00144, which matches the cached value.
5. **Estimated price** (row 1868, named `QuickEstimateOfETSPrices`): p ≈ (reduction left for the ETS) / ε_adj *(inferred: linear semi-elasticity form)*.

This is a one-shot, closed-form estimate. It ignores the interaction with the full fuel-demand model: the real response comes from prices × elasticities by fuel and subsector, not from one semi-elasticity.

## 3. Legacy: the override and goal seek (VBA)

- Row 1872 reads the override prices from Manual inputs O15:AB15 (`OverrideOfETSPrices`). Row 1873 uses them when the override switch is on.
- The full model then computes covered emissions with the ETS (rows 1888–1953) and compares the actual reduction with the cap (rows 1956–1963: actual, cap, difference, ratio, error per year; max error in J1963).
- The workbook proposes a corrected price per year (`PricesAdjustedByErrorFactor`): p × (target log-reduction / actual log-reduction)^0.5, with convergence factor 0.5 in row 1966.
- `OverrideETSFast` sets the dashboard switches (ETS on, override on) and calls `SolveFast`, a damped fixed-point iteration in log space:
  1. If StartFresh, seed the override row from the quick estimate.
  2. Each iteration takes the workbook's proposal r and steps p ← p × (r/p)^α.
  3. Anderson-style mixing: 0.7 new + 0.3 previous iterate.
  4. Light smoothing: 5% weight on a centred 3-year moving average.
  5. Sanitise: negative or invalid values keep the old price.
  6. Write back the row, recalculate, and take the worst-year error over all 14 years. (The named `WorstYearlyETSError` averages only 9 years, so the macro bypasses it.)
  7. Adapt α: halve it if the error got worse (minimum 0.0625); multiply it by 1.2 if the error more than halved (maximum 1).
  8. Stop when the worst error < `ETSErrorTolerance` or after `ETSMaxIterations`.
- The cached run shows an override row of about $22–28/t for 2028–2038, with jumps (1.5 in 2027, 13.8 in 2040) and a max error of 0.33. That run had not converged; the cap rows are 0 in the cached state.

## 4. Legacy: what happens after the split (the same for both parts)

Once the price p is known, it is split by the allocation shares:

| Part | Legacy rows | What it does |
|---|---|---|
| Auctioned | 1875–1879, 2252 (`egy.mit.etstraj`) | p × a enters fuel prices as a **price wedge**, like a carbon tax: fuel switching, efficiency and usage all respond. It also raises revenue. |
| OBA / OBR | 1881–1885, 2343–2349 ("Feebate plus non-auctioned ETS"), 2402–2419 | p × (1 − a) enters as a **shadow price on the efficiency margin only**, the same channel as a feebate: no pass-through to consumer prices and no revenue. |

The legacy note on rows 1875–1885 says "Arguably this should be an exp/ln function TBC": the split is linear in p.

This is also why OBA effectiveness is 0.5. Only the efficiency share of the response is triggered, about half of the total in the default elasticities.

## 5. Proposed MVP design: benchmarks instead of one auction share

The legacy input "auctioned proportion" is replaced by **benchmarks per covered subsector**. Everything after the split stays as in section 4.

**Inputs** (per subsector, in the hidden parameter columns; trajectory rows in years):
- covered flag (already `etsc`, effective coverage);
- benchmark *b*(t) as a share of the subsector's base-year emission intensity (e.g. 0.9 → free allocation at 90% of base-year intensity; a declining path tightens it).

**Shares** (one plain formula per block, the LAMBDA on the right):
- free allocation FA = b × base-year intensity × output. Output is proxied by baseline fuel use, because the MVP has no output variable.
- OBR share s = MIN(FA / covered emissions, 1); auctioned share a = 1 − s.
- To first order, s ≈ b (with intensity unchanged). The exact version uses actual covered emissions and would be circular through `ener`. **Recommendation:** use scenario-1 (baseline) covered emissions, as in option A of `ETS_Cap_Design_Options_v0_1.md`. Then s is a data-driven input and not circular.

**Then, as in legacy:**
- effectiveness e = a + 0.5 × s (per subsector, not one number);
- auctioned part: price wedge p × a in `ets` (fuel prices, revenue = p × (covered emissions − FA));
- OBR part: p × s added to the sector shadow price (`shp`, the feebate channel);
- fast estimate p = (reduction left for the ETS) / (ε × e), with ε and e weighted by covered baseline emissions;
- optional override row: in the MVP an offline goal seek (a Python script reusing the same damped log-space step) writes a pasted row of prices. There is no VBA in the workbook, and no circular reference.

**What changes vs v1.01** (only once you confirm): v1.01 treats the ETS as fully effective (`ets.p` = carbon price trajectory × coverage, as effective as a carbon tax, per your earlier decision), and the auction share only affects revenue. With benchmarks, the OBR part moves from the price wedge to the shadow price, so emissions fall less for the same p and revenue falls. A benchmark of 0 (full auctioning) reproduces v1.01 exactly.

## 6. Open points

1. Semi-elasticities vs the full model: the fast estimate uses legacy's sector semi-elasticities. The MVP could derive them from its own elasticities (one-off calibration run) to keep them consistent.
2. Legacy's linear split (p × a) vs a log form (legacy's own TBC note).
3. Legacy's relative price 1.1 (auctioned ETS vs tax) and the risk-premium inputs (rows 1811–1813) are not used by the fast estimate as cached. Their role should be checked before they are carried over.
4. CBAM sector semi-elasticity (−0.001) is not modelled in the MVP.
