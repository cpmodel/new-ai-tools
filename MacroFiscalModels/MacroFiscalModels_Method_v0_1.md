# Macro-fiscal models for the investment needs work: method (v0.1)

Three simplified, open-source re-implementations of IMF and World Bank tools, written so a country's
mitigation investment programme from NMIM (`../InvestmentNeeds/national_model/`) can be run through
a macro model and a debt sustainability test:

| Model | Original | Question it answers here |
|---|---|---|
| GIMF-lite (`gimf_lite.py`) | IMF Global Integrated Monetary and Fiscal Model (Kumhof, Laxton, Muir and Mursula 2010) | What does a net zero investment push do to GDP, consumption, inflation, the exchange rate and public debt, and does the financing choice matter? |
| DIGNAD-lite (`dignad_lite.py`) | IMF Debt-Investment-Growth and Natural Disasters model (Marto, Papageorgiou and Klyuev 2018), built on DIG (Buffie, Berg, Pattillo, Portillo and Zanna 2012) | Public investment scale-ups and disasters in low-income countries: efficiency, concessional versus commercial financing, adaptation and reconstruction. |
| LIC-DSF-lite (`lic_dsf.py`) | Joint World Bank-IMF Debt Sustainability Framework for Low-Income Countries (2018 Guidance Note) | Does the resulting debt path change the risk of debt distress? |

All three are simplifications. Their calibrations are illustrative, not country calibrations; only
the investment programme is country data (from NMIM). None replaces the official tools.

## 1. GIMF

### 1.1 The full model

GIMF is a multi-region New Keynesian DSGE model (2 to 6 regions; quarterly) with these blocks:

- **OLG households (Blanchard-Yaari).** Each household survives each period with probability θ, so its
  planning horizon is 1/(1 − θ) (about 20 years in typical calibrations). Labour productivity declines
  over the lifetime at rate χ. Aggregate consumption of these households is a fraction of total wealth,
  `C^OLG_t = Θ_t^{-1} (FW_t + HW_t)`, with human wealth discounted at about `R_t / (θχ)`:
  `HW_t = (1 − τ_L) W_t L_t − T_t + (θχ / R_t) HW_{t+1}`. Because θχ < 1, a debt-financed tax cut raises
  perceived wealth: Ricardian equivalence fails, and in steady state higher public debt crowds out
  capital and net foreign assets and raises the real interest rate.
- **Liquidity-constrained households** (share ψ, often about 25%) consume current after-tax income.
- **Firms.** Tradable and non-tradable producers with CES technology in private capital, labour and
  productive public capital; investment adjustment costs (Tobin's q); a financial accelerator
  (external finance premium on net worth) in some versions.
- **Nominal rigidities.** Sticky prices (Rotemberg) and wages; import price stickiness.
- **Monetary policy.** Inflation-forecast-based rule `i_t = ρ i_{t−1} + (1 − ρ)(r̄ + π̄ + φ_π(π_{t+k} − π̄) + φ_y ŷ_t)`,
  with options for a peg, monetary union or the zero lower bound.
- **Fiscal block.** Government consumption, public investment (public capital in production),
  general and targeted transfers, labour, capital, consumption and lump-sum taxes; a fiscal rule
  that targets the deficit-to-GDP ratio with automatic stabilisers.
- **International.** Trade in tradables, intermediate and (optionally) oil and commodities;
  international bonds with a debt-elastic risk premium; uncovered interest parity.

**Data needs for a country calibration:** national accounts shares (C, I, G, public investment,
exports, imports); fiscal revenue by tax and expenditure by type; public debt; net foreign assets;
interest rates and inflation; trade shares by partner; labour share; return on public capital;
estimates of the liquidity-constrained share (Findex, household surveys); planning horizon and
χ (usually taken from GIMF papers); price and wage stickiness and the monetary rule (from the
central bank's own model where one exists).

### 1.2 GIMF-lite equations

One country facing a fixed world (world real rate r*, world demand), annual, one good. Steady-state
GDP = 1. Variables: y, l, w, mc, r^k, π, i, r, q (real exchange rate), k, inv, Tobin's Q, V (value of
firms), HW, C^O, C^L, B (public debt), the fiscal instrument (transfers or labour tax), NFD (net
foreign debt, foreign units), NX.

```
Production        y = A (1 + sav_t) KG_{t-1}^{α_G} (k_{t-1} + ζ KX_{t-1})^α l^{1-α}
Factor prices     r^k = mc α y / K^eff_{t-1};   w = mc (1-α) y / l
Wages (sticky)    ln w = ρ_w ln w_{-1} + (1-ρ_w)[ln w̄ + ε_w ln l]
Phillips curve    π − π̄ = β*(π_{+1} − π̄) + κ_p (mc − mc̄)/mc̄
Taylor rule       i = ρ_i i_{-1} + (1-ρ_i)[(1 + r* + prem)(1 + π̄) − 1 + φ_π(π − π̄)]
Fisher, UIP       1 + i = (1 + r)(1 + π_{+1});   1 + r = (1 + r* + prem) q_{+1}/q,   prem = η_f (NFD − NFD̄)
Investment        Q = 1 + κ(inv/k_{-1} − δ);   k = (1-δ)k_{-1} + inv
                  Q = [(1-τ_K) r^k_{+1} + (1-δ)Q_{+1} − κ/2(x_{+1}−δ)² + κ(x_{+1}−δ)x_{+1}] / (1 + r)
Firm value        V = (div_{+1} + V_{+1})/(1 + r),   div = (1-τ_K)(r^k K^eff_{-1} + (1-mc)y) − inv − adj − IX
Human wealth      HW = (1-ψ)[(1-τ_L) w l + tr] + θχ/(1 + r) · HW_{+1}
OLG consumption   (1+τ_C) C^O = (1 − βθ)(FW + HW),   FW = div + V + (1+r_{-1})B_{-1} − (1+r*+prem_{-1}) q NFD_{-1}
Constrained       (1+τ_C) C^L = ψ[(1-τ_L) w l + tr]
Government        B = (1 + r_{-1})B_{-1} + G + IG + tr − τ_L w l − τ_C C − τ_K(r^k K^eff_{-1} + (1-mc)y)
Fiscal rule       B − B_{-1} = dfs · IG^extra − λ_B (B_{-1} − B̄)        (instrument: tr or τ_L)
BoP               NFD = (1 + r* + prem_{-1}) NFD_{-1} − NX/q
Trade             NX = x̄ q^{ε_x} − m̄ (absorption/absorption̄) q^{-ε_m}
Goods market      y = C^O + C^L + inv + adj + IX + G + IG + NX
Stocks (exog.)    KG = (1-δ_G)KG_{-1} + s·IG;   KX = (1-δ)KX_{-1} + IX
```

β is calibrated so the initial steady state has the chosen NFD/GDP; α_G = R_g · KG/y (return on public
capital). The model is solved under perfect foresight from the initial steady state to the steady
state with the programme's final values held forever (found numerically), with a sparse Newton
solver (`pf_solver.py`). Walras's law is checked: the OLG households' budget, which is not imposed,
holds to 1e-15.

**Climate add-on:** `ig_extra` (incremental public investment), `IX` (incremental private investment,
building capital KX that adds ζ units of effective capital; ζ = 0 by default because most of it
replaces fossil capacity), and `sav` (net energy cost saving, % of GDP, entered as a value-added
gain), plus `dfs`, the share of the extra public investment the fiscal rule lets be debt financed.

### 1.3 Default calibration (illustrative)

r* 3%, π̄ 3%, θ 0.95 (20-year horizon), χ 0.97, ψ 0.25, α 0.33, R_g 15%, δ 8%, δ_G 4%, PIMA
efficiency s 0.7, κ 4, markup 1.1, κ_p 0.25, ρ_w 0.5, φ_π 1.5, ρ_i 0.5, η_f 0.02, import share 30%;
G 12%, IG 4%, τ_L 15%, τ_C 10%, τ_K 20%, debt 50% and NFD 30% of GDP; λ_B 0.05.

## 2. DIGNAD

### 2.1 The full model

Small open economy, tradable and non-tradable sectors, annual, perfect foresight.

- **Households:** savers with access to external commercial borrowing (debt-elastic premium) and
  hand-to-mouth households (share φ). Savers: `λ_t = β(1 + r_t) λ_{t+1}`, `λ = U'(c)/(1+τ_c)`.
  Hand-to-mouth: `(1+τ_c) c^h = (1-τ_y) w L^h + tr`.
- **Production:** `y_j = A_j (k^G)^ψ k_j^{α_j} L_j^{1-α_j}`, with public capital the sum of standard
  and adaptation capital.
- **Capital:** `k = (1-δ) k_{-1}(1 − ε^k) + i`; adjustment costs give Tobin's q;
  `k^{Gi} = (1-δ_G) k^{Gi}_{-1}(1 − ε^{Gi}) + s · i^{Gi}`, with adaptation capital costlier and
  destroyed less (ε^{G2} = (1-ρ) ε^{G1}); absorptive-capacity costs on fast scale-ups.
- **Disasters:** a TFP loss with persistence and destruction of private and public capital; the
  adaptation share reduces the damage.
- **Government:** budget with concessional, domestic and external commercial debt
  (`r^c = r* + η e^{(d^c − d̄)}`), grants, and a fiscal rule that closes the gap gradually with
  consumption taxes and transfers (optionally after a period of debt financing).
- Market clearing for each sector, real exchange rate, balance of payments.

**Data needs:** national accounts and sector shares; revenue by tax, government consumption, transfers,
grants; debt by concessional, domestic and external commercial, with rates; PIMA or PIE-X efficiency;
return on public capital and depreciation; share of hand-to-mouth households (Findex); disaster
frequency and damage (EM-DAT, post-disaster needs assessments); adaptation cost premium and damage
avoided; the available reconstruction financing package (grants, concessional loans).

### 2.2 DIGNAD-lite

One good, inelastic labour, no domestic debt, a disaster as a one-off surprise at t = 1. Fourteen
equations (see the module docstring): production, savers' Euler and budget, hand-to-mouth budget,
Tobin's q and capital Euler, private capital, two public capital stocks, savers' and the government's
debt-elastic rates, government budget, and the two fiscal rules (consumption tax and transfers react
to the commercial-debt gap). Reconstruction and the NMIM programme (`prog`) are public investment
financed by grant, concessional and commercial shares; the programme tapers to zero over ten years
after it ends.

## 3. LIC-DSF

### 3.1 The framework

1. **Debt-carrying capacity**, from the Composite Indicator (10-year averages, 5 historical and 5 projected):
   `CI = 0.385·CPIA + 2.719·growth + 4.052·(reserves/imports) − 3.990·(reserves/imports)² + 2.022·(remittances/GDP) + 13.520·(world growth)`;
   weak < 2.69 ≤ medium ≤ 3.05 < strong.
2. **Thresholds** (weak / medium / strong): PV of external debt/GDP 30/40/55; PV of external
   debt/exports 140/180/240; external debt service/exports 10/15/21; external debt service/revenue
   14/18/23; benchmark for PV of total public debt/GDP 35/55/70.
3. **Debt dynamics:** `d_t = d_{t-1}(1 + i)/((1 + g)(1 + π)) + exchange-rate effect − primary balance + other flows`;
   PV = Σ debt service / 1.05^s; new borrowing by loan terms (rate, grace, maturity).
4. **Stress tests:** standardised (real growth, primary balance, exports, other flows, a 30%
   depreciation, combination, contingent liabilities) and tailored (natural disasters, commodity prices,
   market financing).
5. **Rating:** a baseline breach signals high risk; a breach only under stress, moderate; none, low.
   External and overall ratings, then judgement and realism tools.

**Data needs:** 10-20 years of projections (GDP and deflator, exports, imports, revenue, grants, primary
balance, other flows); loan-by-loan existing external debt service; terms of new financing; domestic
debt; CPIA, reserves, remittances and world growth; ten-year historical standard deviations.

The coefficients, thresholds and shock sizes are as recalled from the 2018 Guidance Note and must be
checked against the current Guidance Note and template before any official use.

### 3.2 LIC-DSF-lite

Everything in USD; PPG external debt only, by loan cohort (equal principal after grace); domestic debt
at face value, rolled over at an effective rate; the five indicators; the seven standard tests plus the
disaster test; the mechanical rating (overall = worse of external and public). No realism tools, no
judgement, no small-breach rule.

## 4. Link with NMIM

`nmim_bridge.py` runs the NMIM engine for one country and takes, for 2025-2050, as shares of NMIM's
real GDP: the incremental direct budget need (`mit.dtyp.bud`, P1 + the public gap of P4); incremental
private investment (`mit.dinv.tot` less the budget need); and net energy savings (`mit.dfuel` less
`mit.dom`). GIMF-lite takes all three; DIGNAD-lite only the public part. The macro paths then shift a
LIC-DSF-lite baseline: growth by the change in the model's GDP growth, the primary deficit by the
change in the model's primary deficit.

## 4a. Shared calibration and model choice

`calibration.py` sets the settings the two macro models share once, from the Tax and Macro-Fiscal
Dashboard panel (`TaxMacroFiscalDashboard/build/output/global_data_wide.csv`), with an
income-group default where the panel has no value. Every setting carries its source.

| Setting | GIMF-lite / DIGNAD-lite | Source |
|---|---|---|
| Efficiency `s` | `s` / `s` | 1 − efficiency gap; gap from PEFA PI-11 (score 1 → 40%, 4 → 13%, as in InvestmentNeeds v1.1), else IMF (2015) by income group |
| Return on public capital | `Rg` / `Rg` | Default by income group: 20 / 17 / 15 / 10% (LIC to HIC) |
| Public investment | `ig` / `ig1 + ig2` (15% adaptation) | Default by income group |
| Taxes | `tau_c, tau_l, tau_k` / `tau_c, tau_y` | Tax revenue (% GDP), split 40% consumption, 45% labour, 15% capital |
| Public debt | `b` / `dk + dc` | Gross debt (WEO); concessional share = IDS multilateral share of PPG external debt |
| Constrained households | `psi` / `phi` | Default by income group: 60 / 50 / 35 / 25% |
| Borrowing cost | `rstar` / `rstar` | IDS private creditor rate less 3%, between 3% and 8%; raised in 0.5-point steps where GIMF-lite's discount factor would exceed 1 (7 low-income countries) |
| Depreciation, capital share | `delta, delta_g, alpha` | 7%, 4%, 0.33 in both |
| Government consumption | `g` / `gc` | Residual that closes the starting budget with transfers of 2% of GDP (floor 5%) |

**Model choice** (`choose_model`): DIGNAD-lite for low-income countries, disaster-prone countries and
lower-middle-income countries whose multilateral share is at least 50%; GIMF-lite otherwise.

**Cross-check** (`cross_check`): both models run on the public part of the NMIM programme with the
same shared settings. With the shared settings, GIMF-lite still gives about 1.6 times DIGNAD-lite's
GDP gain (Kenya 2050: 1.85% against 1.15%). The difference is labour supply: in GIMF-lite
employment responds to higher productivity, in DIGNAD-lite it is fixed. With near-fixed labour in
GIMF-lite (`wage_el` = 50) the gain is 1.03%, close to DIGNAD-lite's. The remaining gap is a
model difference, not a calibration difference.

**Warning on the programme:** non-CO2, land-use and other energy CO2 lines are annual spending in
NMIM, but the bridge passes them to the macro models as public investment, which builds productive
capital. Where they are a large share of the programme (Liberia 90%, Kenya 53% of incremental
investment 2025-2035), the macro gains are overstated. `cross_check` prints a warning above 25%.

## 5. Validation

| Check | Result |
|---|---|
| GIMF-lite, Ricardian limit (θ → 1, χ = 1, ψ = 0) | Debt- and tax-financed public investment give the same consumption path (to 0.001%) |
| GIMF-lite, default | They differ (consumption −0.20% vs −0.52% on impact for a 1% of GDP, 5-year push): Ricardian equivalence fails as in GIMF |
| GIMF-lite, Walras's law | OLG budget, not imposed, holds to 1e-15 |
| GIMF-lite, long run | Terminal inflation = target; terminal debt = target + dfs·ΔIG/λ_B |
| DIGNAD-lite and GIMF-lite solvers | Newton converges in 3-5 iterations; results independent of horizon (T = 80 vs 150) to 1e-4 at year 40 |

## 6. Open issues

- Calibrations are illustrative; a country calibration needs the data listed above.
- GIMF-lite has one country and one good; no oil block, financial accelerator or zero lower bound.
- Energy savings enter as a value-added gain; imported fuel savings are really a terms-of-trade gain and
  domestic producers' losses are not modelled. With savings switched off, the GDP gains fall by about
  two-thirds in the Kenya example.
- Green private capital adds no output (ζ = 0) by default; carbon pricing (CPAT) is not in the models.
- DIGNAD-lite lets the programme taper to zero after 2050; GIMF-lite holds it at its 2050 level.
- Non-CO2, land-use and other spending enter the macro models as productive public capital (section 4a).
- Utility (P2) investment, repaid by tariffs, enters as private investment: state-owned utility debt and tariff shortfalls are not in public debt.
- The DSF baseline in the demo uses placeholder ratios scaled to the country's GDP.
