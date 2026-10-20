---
title: "Macro-fiscal models: getting started"
subtitle: "User guide 0 of 3 · MacroFiscalModels v0.1"
---

# What the toolkit does

The toolkit takes a country's net zero investment programme from the National Mitigation Investment Model (NMIM, `InvestmentNeeds/national_model/`) and asks three questions:

| Question | Model | Guide |
|---|---|---|
| What does the programme do to GDP, consumption, inflation, the exchange rate and public debt? | GIMF-lite | Guide 1 |
| How do public investment efficiency, concessional financing and disasters change the picture in a low-income country? | DIGNAD-lite | Guide 2 |
| Does the resulting debt path change the risk of debt distress? | LIC-DSF-lite | Guide 3 |

All three are simplified versions of IMF and World Bank tools. **Their default calibrations are illustrative, not country calibrations.** Only the investment programme is country data. Use the results to explore mechanisms and orders of magnitude, not as forecasts or official assessments.

# Install

You need Python 3.10 or later and five packages:

```
pip install numpy scipy pandas openpyxl matplotlib
```

The NMIM engine reads its data from `InvestmentNeeds/national_model/data/`, which is in the repository. Nothing is downloaded.

# Run the demo

From the repository root:

```
cd MacroFiscalModels
python3 run_demo.py KEN
```

Replace `KEN` with any ISO3 code in NMIM (186 countries). The run takes about 15 seconds and prints three blocks:

- **A. DIGNAD-lite disaster scenarios** (do not depend on the country).
- **B. The NMIM programme** for the country (public budget need, private investment and energy savings, % of GDP), then GIMF-lite and DIGNAD-lite results for several financing choices.
- **C. LIC-DSF-lite**, before and after the programme.

It writes to `MacroFiscalModels/results/`:

| File | Content |
|---|---|
| `gimf_<ISO>.png` | Six panels: GDP, consumption, private investment, public debt, inflation, real exchange rate |
| `dsf_<ISO>.png` | The five debt indicators, baseline and with the programme, with thresholds |
| `dignad_disaster.png` | Disaster scenarios |
| `paths_<ISO>.csv` | Annual programme and GIMF-lite paths, 2025-2050 |

# Run your own scenario

Each model is a Python module with a `Params` dataclass. Change any parameter with `replace` and solve:

```python
from dataclasses import replace
from nmim_bridge import programme, run_gimf
from gimf_lite import Params

prog = programme("SEN")                        # NMIM paths for Senegal
p = replace(Params(), psi=0.4, s=0.5, dfs=0.5)  # more constrained households,
                                                # lower efficiency, half debt financed
out, ss0, ssT, cal = run_gimf(prog, p)
print(100 * (out["y"][25] / ss0["y"] - 1))      # GDP gain in 2050, %
```

`out` holds annual paths from 2025 (index 0); `ss0` is the starting steady state and `ssT` the long-run steady state with the programme.

# Calibrate from the dashboard

`calibration.py` sets the settings both macro models share (efficiency, return on public capital, public investment, taxes, debt, concessional share, constrained households, borrowing cost) from the dashboard panel, with income-group defaults where data are missing. It also chooses the model and runs both on the public part of the programme as a cross-check:

```
python3 calibration.py KEN SLE UGA LBR
```

From code:

```python
from calibration import shared, to_gimf, to_dignad, choose_model
sh = shared("UGA")                  # {setting: (value, source)}
p_g, p_d = to_gimf(sh), to_dignad(sh)
choose_model("UGA")                 # ('DIGNAD-lite', 'low-income country')
```

Check the warning it prints: where non-CO2 and land-use spending are a large share of the programme, the macro gains are overstated.

# How the pieces connect

1. `nmim_bridge.programme(iso)` runs NMIM for one country and returns, for 2025-2050, as shares of GDP:
   - `pub`: incremental direct budget need (P1 pure public plus the public gap of P4 viability gap);
   - `prv`: incremental private investment (total incremental less `pub`);
   - `sav`: net energy cost savings (fuel savings less extra O&M).
2. `run_gimf(prog, p)` feeds all three to GIMF-lite; `run_dignad(prog, p)` feeds `pub` to DIGNAD-lite.
3. `dsf_overlay(inputs, y_path, prim_def_path, prim_def_ss)` shifts a LIC-DSF-lite baseline by the macro model's change in growth and primary deficit; `lic_dsf.assess(...)` rates it.

# Where to read more

- `MacroFiscalModels_Method_v0_1.md`: equations of the full and lite models, data needs, validation, open issues.
- Guides 1-3: parameters, outputs and recipes for each model.
- `CAVEATS.md`: task log and caveats.

# Five rules for using the results

1. **State the calibration.** Every chart should say "illustrative calibration" until a country is calibrated.
2. **Test energy savings first.** In GIMF-lite they drive about two-thirds of the long-run GDP gain in the Kenya example. Run with `savings=False` as a lower bound.
3. **Report ranges, not points.** Vary the planning horizon (`theta`), the constrained share (`psi`) and efficiency (`s`).
4. **Use the country's own DSA baseline** in LIC-DSF-lite. The demo's baseline ratios are placeholders.
5. **Do not present a LIC-DSF-lite rating as a DSA rating.** It is the mechanics of the framework without judgement.
