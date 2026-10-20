# PriceFeasibility: notes

## What this is

A quick screening tool. It asks one narrow question: **how much power-sector
CO2 can carbon pricing cut in each country without exceeding set limits on
political cost?** It compares three instruments:

- **Carbon tax or ETS.** Raises the marginal and the average electricity
  price, so it gets both substitution and demand reduction. It also triggers
  all three political costs.
- **Feebate.** The same marginal incentive with no rise in the average price.
  Output-based allocation in an ETS and tradable performance standards work
  the same way. It gets only substitution, and its political cost falls on
  high-emitting generators.
- **Hybrid.** The tax is set as high as the price limit allows, then a feebate
  tops up the marginal incentive.

Political feasibility is measured by three numbers, each with a limit:

- **F1:** the tariff rise households see.
- **F2:** the net cost to the bottom 40% after rebates.
- **F3:** net payments by coal and oil generators.

The tool reports:

- the best mix for each country;
- which limit binds;
- the shadow price of that limit (the extra abatement from one more point of
  allowed tariff rise).

**It is not a substitute for CPAT.** It has one sector, one period and simple
response functions. Several parameters are priors (see README, "Priors to
replace"). Use it to rank countries and instruments, to find which constraint
matters where, and to frame questions for a full CPAT run. Don't quote its
numbers as estimates.

## Why it exists

CPAT reports effectiveness and distribution for each instrument separately.
It doesn't ask the constrained question: "best abatement subject to a
tariff-rise limit". It also doesn't report the shadow price of a feasibility
limit. Those two outputs are the point of this tool.

## Can it be integrated? Yes, at three levels

### 1. As a post-processor on CPAT output (recommended first step)

Keep the F1-F3 and optimiser logic here. Feed it CPAT results instead of
priors. For each country, run CPAT for `Carbon tax` and `Power feebate` at a
few rates, then read:

| Tool quantity | CPAT source (coded, `cpat_model/components/`) |
|---|---|
| `A_sub`, `A_dem` (or `s` directly, as feebate/tax abatement) | Emissions by scenario (`emissions/em.py`); power generation mix (`power/generation.py`) |
| Tariff rise, `dp` | Domestic electricity prices (`prices/domestic_prices.py`) |
| `elasticity` | `elasticities/elasticities.py` |
| `b40`, F2 | Distribution module (`distribution/budget_shares.py`, `effects.py`, `recycling.py`) |
| Generator payments, F3 | Power variable costs by technology (`power/variable_cost.py`) times the feebate rate |

The response curves `A_sub(tau)` and `A_dem(tau)` would then be interpolated
from CPAT runs rather than taken from the exponential and linear forms here.
`evaluate()` is the only function that would change.

**Check first: how CPAT implements feebates.**
- In `policies/policies.py`, `Feebates` and `Power feebate` apply the carbon
  price to the power sector only. They switch off explicit pricing.
- There is an open TODO there: fuel inclusion is set to True for all fuels,
  "looks suspicious in Excel, I think it is a BUG".
- Confirm this is a revenue-neutral, benchmark-based feebate before using its
  results as the feebate case here.

### 2. Inside CPAT, as a scenario or dashboard output

This is the bigger step. Add F1-F3 as standard CPAT outputs for every
scenario. Add a "constrained optimum" run mode that searches rates for a
tax-plus-feebate combination. It belongs in the master coded CPAT (CPAT-AI
repo, `cpat_coded/`), not in `CPAT_coded_not_master`. It should follow that
repo's norms: method note separate from version notes, and a CAVEATS entry.
Do it only after level 1 has shown the metrics are useful.

### 3. Into the other tools in this repo

- **TaxMacroFiscalDashboard.** Add a small "Pricing feasibility" block per
  country, using output from this tool run on the same panel: `s_at50`,
  feebate/tax ratio, binding constraint and F1 shadow price. The calibration
  inputs already come from that panel.
- **MacroFiscalModels.** Built (October 2026): `MacroFiscalModels/pricing_bridge.py`
  turns the feasible mix's tax rate times NMIM's net zero power CO2 into revenue
  earmarked for the public investment need, and `MacroFiscalModels/integrated.py`
  runs it through the macro models for every country (a feebate raises roughly
  none).
- **ExciseDiagnostic.** Part 1 (fuel) could reuse the F1 logic once road fuels
  are added.

## Extensions worth doing

- **Road fuels.** Feebates act on new vehicles, so they act slowly. The tax
  and excise trade-off is mostly F1 there.
- **Industry.** Carbon leakage and output-based allocation.
- **Calibration data now in the panel but unused.** Explicit subsidies
  (administered prices, so lower pass-through) and solar WACC (switching
  cost).
- **Country thresholds instead of global placeholders.** Indicators for each
  limit are listed in the README.
