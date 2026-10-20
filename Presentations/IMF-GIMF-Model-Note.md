# IMF GIMF: the Global Integrated Monetary and Fiscal Model

Information note, 20 Oct 2026. Written from memory, without a fresh literature search: check the exact IEA wording and the calibration tables before quoting.

## What it is

The model with "Ricardian equivalence violation due to myopia" that the IEA cites is the IMF's **Global Integrated Monetary and Fiscal Model (GIMF)**. The IEA used it, with IMF staff, for the macroeconomic analysis in *Net Zero by 2050* (2021), and it is cited in later World Energy Outlook work.

**Terminology:** GIMF is a multi-region **New Keynesian DSGE** model, not a CGE model. People often call it a CGE informally; in writing, use "DSGE" or "general equilibrium macro model".

## Definition (citable)

> **GIMF** is a multi-country, micro-founded New Keynesian DSGE model developed at the IMF (Kumhof, Laxton, Muir and Mursula, 2010) for analysing fiscal and monetary policy in an open economy. **Ricardian equivalence fails** in it for two reasons. First, most households have **finite planning horizons**: they are Blanchard–Yaari overlapping-generations agents, whose "myopia" means they discount future tax liabilities more heavily than the government's cost of borrowing. Second, a fraction of households are **liquidity-constrained** and spend all their current income each period. Because of this, the timing of taxes and the path of government debt affect consumption, interest rates, the current account and output, in the short run and the long run.

## How the myopia works

- Each household faces a constant per-period probability **(1 − θ)** of "exiting" the economy. This is read as a **planning horizon**, not literal mortality. The average horizon is **1/(1 − θ)**; typical calibrations put it around **20 years**.
- Households hold annuity-type contracts, so the return they earn on their assets is grossed up by 1/θ.
- **Human wealth** is the present value of future after-tax labour income minus lump-sum taxes. Households discount it at roughly **R/θ**, higher than the government's borrowing rate **R**. Expected labour income also declines over the life cycle, at a rate χ, which raises the effective discount rate further.
- Aggregate consumption from these households is a share of total wealth: **C = (financial wealth + human wealth) / Θ**.

**Why Ricardian equivalence breaks:** suppose the government cuts taxes today, finances the cut with debt, and raises taxes later to pay it back. In present-value terms at rate R, nothing changes. But households discount the later tax rise at R/θ, so their perceived wealth rises and they spend more. Part of the later tax burden also falls on future cohorts. So in steady state, higher government debt crowds out capital and net foreign assets and raises the world real interest rate.

## Other features worth citing

- **Liquidity-constrained (hand-to-mouth) households:** a share of households, often around 25% depending on the country, consume their current after-tax income. This makes transfers and temporary measures more powerful.
- **Rich fiscal block:** government consumption, productive **public investment** (public capital enters the production function), general and targeted transfers, and labour, capital and consumption taxes. A **fiscal rule** targets the deficit-to-GDP ratio and lets automatic stabilisers operate.
- **Monetary policy:** an inflation-forecast-based interest rate rule; some uses allow for the zero lower bound or monetary accommodation.
- **Nominal and real rigidities:** sticky prices and wages, investment adjustment costs, and a financial-accelerator external finance premium.
- **Multi-region trade and finance:** usually 2 to 6 regions, with tradable and nontradable sectors and an optional oil or commodity block. This lets the IEA model global spillovers from energy investment and carbon pricing.

## Key references

- Kumhof, M., Laxton, D., Muir, D. and Mursula, S. (2010), *The Global Integrated Monetary and Fiscal Model (GIMF): Theoretical Structure*, IMF Working Paper WP/10/34. The core specification.
- Anderson, D., Hunt, B., Kortelainen, M., Kumhof, M., Laxton, D., Muir, D., Mursula, S. and Snudden, S. (2013), *Getting to Know GIMF: The Simulation Properties of the Global Integrated Monetary and Fiscal Model*, IMF WP/13/55.
- Underlying theory: Blanchard, O. (1985), "Debt, Deficits, and Finite Horizons", *Journal of Political Economy*; Yaari, M. (1965).

## Caveats

The calibrations quoted (a horizon of about 20 years, about 25% liquidity-constrained) differ by paper and by region. For exact values, use the tables in the working papers above and the IEA report's annex.
