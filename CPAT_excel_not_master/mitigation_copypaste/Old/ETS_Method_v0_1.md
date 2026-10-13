# New ETS: method (CPAT-AI-Mitigation-MVP)

How the MVP models a new emissions trading system. Background on the legacy algorithm and the design choice: `ETS_Legacy_Algorithm_v0_1.md` and `ETS_Cap_Design_Options_v0_1.md`. Row names are the variable codes in column A of the Mitigation sheet. Section 1 holds the inputs and the price; the sector sections hold the per-fuel wedges; section 13 holds the cap check.

## Inputs

| Input | Where | Notes |
|---|---|---|
| Apply new ETS, start and target year | MTInputs `D_NewETS`, `D_ETSIntro`, `D_ETSOutro` | |
| Cap: change vs baseline covered emissions at start and target, continuation | MTInputs `D_ETSChangeRelStart`, `D_ETSChangeRelTarget`, `D_ETSCapCont` | `Constant` keeps the cap level after the target year; any other value keeps the relative change |
| Price override switch | MTInputs `D_ETSPriceOverride` | `Yes`: the permit price comes from the override row `ets.ovr` |
| Volatility, impact of volatility, carbon tax policy risk | MTInputs `D_ETSVolatility`, `D_ETSVolImpact`, `D_ETSCTRisk` | Labels Medium / High / Low / Zero; values in the Settings table (rows 16-18, legacy ETS+LTS) |
| Sector coverage | MTInputs rows 99-115 | 17 ETS sectors: power plus the 16 subsectors |
| Benchmarks by sector group | Section 1 typed rows `etsb.s`, `etsb.t` (power, transport, residential/buildings, industry) | Share of a sector's emission intensity that is allocated free. Assumption 1.0 at the start and 0.8 at the target year |
| Semi-elasticities by ETS sector | Section 1 rows `obr`, column E (hidden) | Legacy Egypt sector table (rows 1815-1831); used only by the fast estimate |
| Baseline CO2 by ETS sector | Section 1 data rows `bco2` | Values of scenario 1's `co2.sec` rows; the build refreshes them, check row `bco2.chk` |
| Effectiveness of output-based allocation, convergence exponent | Settings C14 (0.5), C15 (0.5) | Legacy hardcodes |

## Price

- Effective coverage `etsc` = 1 for a covered sector once the ETS applies, else 0.
- Benchmark path `etsb` = linear from the start to the target year, flat afterwards. OBR share by sector `obr` = MIN(1, MAX(0, benchmark of its group)).
- Volatility adjustment `ets.vadj` = (1 + carbon tax policy risk x impact) / (1 + ETS volatility x impact). At Medium this is (1 + 0.2 x 0.5) / (1 + 0.42 x 0.5) = 1/1.1.
- Baseline covered emissions `ets.bce` = SUMPRODUCT(`bco2`, `etsc`).
- Cap `ets.cap` = `ets.bce` x (1 + change), where the change runs linearly from the start to the target year (after the target year, see `D_ETSCapCont`).
- Effective semi-elasticity `ets.se` = baseline-weighted average over covered sectors of semi-elasticity x (1 − (1 − 0.5) x OBR share).
- Fast estimate `ets.est` = LN(cap / baseline covered) / (`ets.se` x `ets.vadj`), when the cap is below the baseline; else 0.
- Permit price `ets.p` = the override row `ets.ovr` if `D_ETSPriceOverride` = Yes (a blank cell takes the carbon price path), else the fast estimate. 0 before the ETS applies.
- Tax-equivalent price `ets.pe` = `ets.p` x `ets.vadj`; revenue factor `ets.rf` = 1 / `ets.vadj`.

## Effect on fuel use

The ETS replaces the carbon tax in covered sectors (`ctxnew` x (1 − `etsc`)). The tax-equivalent price splits by the OBR share:

- **Price wedge (auctioned part):** `ets` = `ets.pe` x EF x `etsc` x (1 − OBR). It enters `nce` and the after-tax price `atp`, so usage and efficiency both respond, as for a carbon tax.
- **Shadow price (OBR part):** `shp` = (feebate shadow price x share + `ets.pe` x `etsc` x OBR) x EF. It acts only on the efficiency margin of the fuel-use equation, as feebates do.

## Revenue

`rnew` includes `ets` x `ets.rf`, which is the permit price on the auctioned part of covered emissions. The OBR part raises no revenue.

## Meeting the cap

The fast estimate uses one semi-elasticity, so the full model does not hit the cap exactly. Section 13 shows:

- `co2.sec`: CO2 by ETS sector;
- `co2.ets`: covered emissions;
- `co2.cap`: the cap;
- `co2.gap`: covered / cap − 1;
- `ets.next`: p x (LN(cap/baseline) / LN(covered/baseline)) ^ 0.5.

To meet the cap, either:

- paste `ets.next` into `ets.ovr` (with the override switch on) and repeat; or
- run `python ets_goalseek_v0_2.py WORKBOOK.xlsx <scenario>`. It iterates with legacy's damped log-space step (mixing, smoothing, adaptive step) until the worst yearly gap is below 0.5%, and writes the prices to `ets_override_s<scenario>.csv` for pasting into `ets.ovr`.

## Limits

- Power is not modelled, so covered power emissions are 0. The cap applies to the covered end-use sectors only.
- Benchmarks are relative to current emission intensity, and output is proxied by fuel use. Free allocation therefore moves with emissions, and the OBR share equals the benchmark.
- The volatility adjustment scales behaviour (the tax-equivalent price), not the permit price. VAT on new policies is levied on the tax-equivalent ETS cost.
- No banking, borrowing, price floors or ceilings, and no CBAM process emissions.
