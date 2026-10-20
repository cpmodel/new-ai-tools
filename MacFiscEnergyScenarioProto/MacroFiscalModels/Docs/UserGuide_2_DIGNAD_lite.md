---
title: "DIGNAD-lite user guide"
subtitle: "User guide 2 of 3 · MacroFiscalModels v0.1"
---

# What it is

The same module is also **DIGNAR-lite** when the resource block is on (see Parameters).


DIGNAD-lite is a simplified version of the IMF's Debt-Investment-Growth and Natural Disasters model (Marto, Papageorgiou and Klyuev 2018), which builds on the DIG model of public investment scale-ups in low-income countries (Buffie, Berg, Pattillo, Portillo and Zanna 2012).

It has savers (who borrow abroad at a debt-elastic rate) and hand-to-mouth households; private capital with adjustment costs; **two public capital stocks**, standard and adaptation (adaptation costs more but is destroyed less by disasters); **public investment efficiency** (only a share `s` of each dollar becomes capital); **three kinds of financing**: grants, concessional loans and commercial external debt, whose rate rises with the debt stock; and a fiscal rule in which the consumption tax and transfers close the gap.

**Use it for:** low-income and disaster-prone countries; public investment scale-ups (including NMIM's budget need); the value of concessional finance and of better public investment management; disasters, reconstruction and adaptation.

**Do not use it for:** private investment pushes (use GIMF-lite), monetary policy or the exchange rate (one good, no money).

# Running it

```python
from dataclasses import replace
import dignad_lite as dg

# a disaster with reconstruction (defaults)
out, ss, cal = dg.solve(dg.Params())

# the NMIM public programme, no disaster
from nmim_bridge import programme, run_dignad
out, ss, cal = run_dignad(programme("ETH"))

# your own programme: 1% of GDP for 10 years, 50% grants
p = dg.Params(eps_k=0, eps_g=0, eps_a=0, rec=0, prog=(0.01,) * 10, sh_grant=0.5, sh_conc=0.3)
out, ss, cal = dg.solve(p)
```

The disaster is a surprise at year 1; then the path is known (perfect foresight). To run a programme without a disaster, set `eps_k`, `eps_g`, `eps_a` and `rec` to 0, as `run_dignad` does.

# Parameters

All shares are of GDP; rates are annual and real.

## Economy

| Parameter | Default | Meaning | Where to get it |
|---|---|---|---|
| `rstar` | 0.04 | World real interest rate | |
| `alpha` | 0.35 | Private capital share | National accounts |
| `delta`, `delta_g` | 0.06, 0.05 | Depreciation, private and public | IMF Investment and Capital Stock dataset |
| `kappa` | 5.0 | Private investment adjustment cost | Literature |
| `phi` | 0.5 | Share of hand-to-mouth households | Findex, household surveys |

## Public investment

| Parameter | Default | Meaning | Where to get it |
|---|---|---|---|
| `s` | 0.6 | Efficiency: share of spending that becomes capital | PIMA, PIE-X |
| `Rg` | 0.20 | Return on public capital | IMF literature: 10-25% |
| `ig1`, `ig2` | 0.06, 0.01 | Standard and adaptation public investment before the shock | BOOST, budget data |
| `mu_a` | 0.25 | Cost premium of adaptation capital | Engineering studies; 15-50% |

## Fiscal and debt

| Parameter | Default | Meaning | Where to get it |
|---|---|---|---|
| `tau_y`, `tau_c` | 0.10, 0.10 | Income and consumption tax rates | Revenue data |
| `gc`, `grants` | 0.08, 0.02 | Government consumption and grants | Budget, OECD CRS |
| `dk`, `rk` | 0.30, 0.01 | Concessional debt and its real rate | DSA, World Bank IDS |
| `dc`, `eta_g` | 0.20, 0.05 | Commercial external debt and its premium elasticity | DSA, spreads |
| `bp`, `eta_p` | 0.10, 0.05 | Private external debt and its premium elasticity | IIP |
| `lam_tau`, `lam_tr` | 0.05, 0.05 | Fiscal rule: response of the consumption tax and of transfers to the debt gap | Judgement |

## Disaster

| Parameter | Default | Meaning | Where to get it |
|---|---|---|---|
| `eps_k`, `eps_g` | 0.10, 0.10 | Share of private and standard public capital destroyed | Post-disaster needs assessments (PDNA), EM-DAT |
| `rho_res` | 0.80 | Adaptation capital is destroyed at (1 − `rho_res`) × `eps_g` | Engineering studies |
| `eps_a`, `rho_a` | 0.03, 0.6 | TFP loss and its persistence | Literature on disaster output losses |
| `theta_a` | 0.5 | How much the adaptation share reduces private and TFP damage | Judgement |

## Reconstruction and programme financing

| Parameter | Default | Meaning |
|---|---|---|
| `rec`, `rho_rec` | 0.03, 0.7 | Reconstruction spending in year 1, and its decay |
| `bbb` | 0.0 | Build back better: share of reconstruction that is adaptation capital |
| `sh_grant`, `sh_conc` | 0.3, 0.3 | Shares financed by grants and new concessional loans; the rest is commercial |
| `amort_k` | 0.03 | Repayment rate of concessional loans |
| `prog` | `()` | Programme of extra public investment, % GDP by year from year 1 |
| `prog_bbb` | 0.0 | Share of the programme that is adaptation capital |
| `prog_taper` | 10 | Years over which the programme tapers to zero after its last year |
| `prog_spend` | `()` | Programme spending that builds no capital (non-CO2, land use), same financing as `prog` |
| `rev` | `()` | Revenue earmarked for the programme (e.g. feasible carbon pricing), % GDP by year |

## Resource block (DIGNAR-lite)

| Parameter | Default | Meaning | Where to get it |
|---|---|---|---|
| `rents0` | 0.0 | Resource rents, % of non-resource GDP; 0 switches the block off | WDI oil + gas + coal rents (dashboard) |
| `gov_take` | 0.5 | Government share of rents | Resource fiscal regime, EITI |
| `rents_path` | `()` | Rents by year from year 1, held at the last value | A price and volume scenario; `calibration.to_dignar` uses an 80% fall by 2050 |

With the block on, `out["gdp"]` is total GDP (non-resource + rents) and `debt_gdp` is debt over total GDP. A permanent fall in rents needs a firmer fiscal rule (`lam_tau = lam_tr = 0.2`), which `to_dignar` sets.

# Outputs

`out` holds annual paths from year 1; `ss` is the steady state (GDP = 1):

| Key | Meaning |
|---|---|
| `y`, `C`, `cs`, `ch` | GDP; total, savers' and hand-to-mouth consumption |
| `k`, `kg1`, `kg2` | Private, standard public and adaptation public capital |
| `dc`, `dk`, `debt`, `debt_gdp` | Commercial, concessional, total public debt; debt / GDP |
| `rc`, `rp` | Government and private borrowing rates |
| `tauc`, `tr` | Consumption tax and transfers (the fiscal rule's instruments) |
| `revenue`, `prim_bal` | Revenue and primary balance |
| `ig1`, `ig2`, `grants`, `rec`, `prog` | The exogenous paths used |

# Recipes

| Question | How |
|---|---|
| Value of concessional finance | Compare `sh_grant=0.3, sh_conc=0.3` with `sh_grant=0, sh_conc=0` |
| Value of better public investment management | Compare `s=0.4` with `s=0.7` |
| Build back better | `bbb=1.0` (reconstruction as adaptation capital) |
| Adaptation before the disaster | Raise `ig2` and lower `ig1` (same total), e.g. `ig1=0.04, ig2=0.03` |
| Bigger disaster | Raise `eps_k`, `eps_g`, `eps_a` |
| Faster fiscal adjustment | Raise `lam_tau`, `lam_tr` |

# Reading the results

- Reconstruction speeds the recovery but raises debt; grants and concessional loans reduce the debt cost.
- Build back better costs more today and pays off only when the next disaster hits. With a single disaster it looks slightly worse; the "adaptation first" scenario shows the payoff.
- Hand-to-mouth consumption falls with wages and with transfers cut by the fiscal rule; it is the welfare-relevant series for the poor.

# Limits

One good; fixed labour; no domestic debt; a single surprise disaster (no recurrent disasters or insurance); the programme tapers to zero after it ends. See the method note, section 6.
