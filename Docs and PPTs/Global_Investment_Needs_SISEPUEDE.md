---
title: "Global Investment Needs for Climate Action: A Method Based on CDPM (formerly SiSePuede)"
subtitle: "A global method for climate investment needs, built on CDPM's equations and parameterised with the most complete global estimates, evaluated at commercial rates, at World Bank rates, and at World Bank rates with co-benefits. Appendices cover CDPM, the other models and estimates, the parameters, and an annotated bibliography on financing."
date: "October 20, 2026"
---

# Executive summary

> **Name.** CDPM was formerly called SiSePuede (and CPM). Studies published before the change use the old name; file, repository and code names keep it too.

Every climate finance discussion starts with a number for investment needs. That number depends on three choices that are rarely stated:

* what is counted: total or incremental investment, capital only or net of savings;
* which discount rate, or cost of capital, turns future flows into today's money;
* whether benefits other than fuel savings are counted.

This note sets out one global method that makes all three explicit. It uses CDPM's equations and parameters wherever CDPM has them. It then applies the method to the world.

**Which model is complete.** No single model covers every country, every sector and the cost of capital. Nine models and estimates were reviewed (Table 2 and Appendix B).

* **The most complete global estimate is the IHLEG's (Bhattacharya et al. 2024).** It covers:
  * all countries, in three groups: advanced economies, China, and the other emerging market and developing countries (EMDCs);
  * all climate investment: clean energy, adaptation and resilience, loss and damage, natural capital and the just transition;
  * the cost of capital and co-benefits, which it reports.

  It is a synthesis, not a model. Energy comes from the IEA's Global Energy and Climate (GEC) model, which is proprietary. Adaptation comes from UNEP's Adaptation Gap Report.
* **The most complete open model is CDPM (RAND and IDB).** It:
  * covers every IPCC sector at country level;
  * prices costs, savings and co-benefits in one module;
  * runs about a thousand futures.

  It has been applied to 18 Latin American and Caribbean (LAC) countries. Newer applications cover Mexico, Peru and the state of Oaxaca, and, outside LAC, Uganda, Morocco and Libya. Its parameters are LAC-specific and need review (Appendix A).
* **OSeMOSYS Global** is the most complete open power-sector model, covering 163 countries. The IEA's GEC model has the richest energy detail but is closed.

The method therefore combines the strengths of each:

* **Structure:** CDPM's equation (physical change times unit cost, compared with a baseline).
* **Default quantities and totals:** the IHLEG and IEA.
* **Gaps:** Beyond the Gap for infrastructure, UNEP for adaptation.
* **Financing:** a country-and-technology cost of capital, which no open model yet has.

**The method in brief (equations 1–7).** For each country, sector and technology:

1. Investment equals the change in physical capacity, times a unit capital cost, times a local price factor, compared with a baseline. Where no physical model exists, a share of GDP by income group is used instead.
2. The capital is converted into an annual charge with a capital recovery factor at the chosen rate.
3. The annual charge is set against fuel and operating savings (the financial case).
4. Co-benefits and avoided emissions, valued at the World Bank shadow price of carbon, are then added (the economic case).

**Global needs.** Climate investment of **USD 6.3–6.7 trillion a year by 2030** (2023 prices), about 5% of world GDP:

| Group | USD trillion a year, 2030 | Share of GDP |
|---|---|---|
| Advanced economies | 2.7–2.8 | Not reported |
| China | 1.3–1.4 | Not reported |
| EMDCs other than China | 2.3–2.5 | 6.5% |

Clean energy is USD 4.8 trillion of the global total. For EMDCs other than China, the totals rise to USD 3.1–3.5 trillion by 2035. Both are total needs. The increment over what would happen anyway is smaller: 1.4–2.8% of GDP in the CCDRs and about 2% in CDPM's LAC study.

**The three lenses (Table 6).** Each figure below is for one year's investment of USD 6.5 trillion, with each item expressed as an annual flow over a 25-year asset life.

* **(a) Commercial rates.** The WACC is 5.5% in advanced economies, 6% in China and 10.5% in EMDCs.
  * The capital charge is USD 585 billion a year, against financial savings of USD 865 billion.
  * The world as a whole still gains, by about USD 280 billion.
  * In EMDCs the margin is thin: USD 45 billion on USD 2.4 trillion of investment. At 15% it turns negative.
  * The extra financing cost in EMDCs, compared with 6%, is USD 87–184 billion a year for each year's investment.
* **(b) World Bank rate of 6%.** The charge is USD 508 billion and the net financial benefit USD 356 billion, a benefit-cost ratio of 1.7. EMDCs gain USD 131 billion.
* **(c) World Bank rate plus co-benefits.** The net benefit rises to about USD 1.2 trillion a year, and to USD 1.3–1.4 trillion with avoided emissions valued at the shadow price of carbon. The benefit-cost ratio is 3.6–3.8.

The gap between (a) and (c) in EMDCs is the case for de-risking and concessional finance. For each year's investment, the private investor sees a gain of USD 45 billion, or a loss at 15%. Society sees a gain of USD 440 billion.

**What needs review.** Three parameters drive the lens results:

* the financial savings per dollar of capital (0.133 a year);
* the co-benefits per dollar (0.129 a year);
* the abatement per dollar (0.35 tCO2e a year per USD 1,000).

All three come from CDPM's LAC results and the IEA's net zero pathway. They should be replaced, country by country, by CDPM runs. The EMDC cost of capital is the most influential input. It should come from the IEA Cost of Capital Observatory or CPI, by country and technology.

CDPM itself needs five fixes before it can produce the method's outputs globally (Appendix A):

1. a dual discount rate;
2. consistent discounting of power-sector capital;
3. technology interest rates;
4. capital recovery factors in place of straight-line charges;
5. a separate gross capital output.

Appendix A covers CDPM. Appendix B covers the other models and estimates, Appendix C the full parameter set, and Appendix D the verification notes. Appendix E is the annotated bibliography on financing.


# Why the investment-needs number differs between studies

Four definitional choices explain most of the spread in published numbers. They matter more than model differences.

**Total or incremental.** Beyond the Gap and the IEA report *total* investment: all new capital needed to reach a goal. The CCDRs, Bhattacharya et al. and CDPM report *incremental* investment: the extra spending above a baseline that would happen anyway. Incremental numbers are typically a quarter to a half of totals. The CCDR synthesis makes the point itself: its USD 574 billion a year is lower than the IHLEG's USD 1.2–1.7 trillion because it counts only the climate increment.

**Capital only or net of savings.** Beyond the Gap counts capital and reports operations and maintenance separately. The IEA and McKinsey count capital only. CDPM counts technical costs that mix capital, operations and avoided costs, then sets them against fuel savings. Its "investment" line is therefore closer to a net incremental cost.

**The discount rate.** A sum of future costs depends heavily on the rate. Beyond the Gap uses 6% (World Bank). The LAC CDPM study uses 7%. The Costa Rica study used 5%. The IEA embeds technology- and country-specific WACCs. Most others state no rate.

**What counts as a benefit.** CDPM counts fuel savings, avoided air pollution, congestion, crashes, ecosystem services and diet-related health. The IHLEG (2024) puts co-benefits and avoided costs at 15–18% of global GDP in 2030. Most other studies count none.

Table 1 sets out how each of the major estimates makes these choices.

**Table 1. How the major estimates define investment needs**

| Study | What is counted | Scope | Discount rate / cost of capital | Co-benefits |
|---|---|---|---|---|
| CDPM LAC (Kalra et al. 2023) | Incremental technical cost vs baseline; set against savings | 18 LAC countries, 2025–2050 | 7% flat, applied after the model run | Yes: pollution, crashes, congestion, ecosystems, diet health |
| Beyond the Gap (Rozenberg & Fay 2019) | Total new capital; O&M separate | LMICs, infrastructure, 2015–2030 | 6% real, flat | No |
| CCDR synthesis (World Bank 2023) | Incremental climate investment | 42 economies, to 2030 | Not stated; varies by country model | No (country-level in places) |
| Bhattacharya et al. (2022) | Incremental spending above 2019 | EMDEs other than China, 4 areas | Not stated | No |
| IHLEG 3 (Bhattacharya et al. 2024) | Total climate investment | Global, by group, 2030/2035 | Reports solar WACC 9–12% EMDCs vs 5–6% AEs | Yes, 15–18% of global GDP (headline) |
| IEA NZE / IEA-IFC (2023) | Total clean energy capital | Global / EMDEs | Country- and technology-specific WACC in model | No |
| CPI GLCF (2023–2025) | Total climate finance needed, from models | Global | Not stated | No |
| UNFCCC NDR (2021, 2024) | Costed needs in NDCs, self-reported | Developing countries | Not applicable | No |
| McKinsey (2022) | Total capital on physical assets | Global, 69 countries | Not stated | No |



# Which model is complete for global investment needs

A complete model for global investment needs would do six things:

1. cover every country;
2. cover every sector: power and grids, end uses, transport vehicles and infrastructure, buildings, industry, agriculture, forestry and land use (AFOLU), and adaptation;
3. separate total from incremental investment;
4. compute capital explicitly from physical change;
5. use a cost of capital that differs by country and technology;
6. be open enough to review.

Table 2 rates the candidates on each test. Appendix B gives the detail.

**Table 2. How complete each model or estimate is for global investment needs**

G means the test is met, M partly met, P not met.

| Model or estimate | Countries | Sectors | Total vs incremental | Capital from physical change | Cost of capital by country | Open |
|---|---|---|---|---|---|---|
| IHLEG (2024) synthesis | G: world, three groups | G: energy, adaptation, loss and damage, nature, just transition | M: total | M: inherits IEA | M: reported, not applied | M: report |
| CDPM (RAND/IDB) | M: 18 LAC; Uganda, Morocco, Libya in progress | G: all IPCC sectors, with co-benefits | G: incremental | G: activity × unit cost | P: one rate | G: code and data |
| IEA GEC model | G: 29 regions | M: energy only | G: both | G: stock and capacity | G: in model | P: proprietary |
| OSeMOSYS Global | G: 163 countries | P: power only | M: gross | G: linear programme | G: by region and technology | G |
| MESSAGEix-GLOBIOM and IAMs (NGFS, McCollum et al. 2018) | M: 10–33 regions | M: energy and land | M: gross | G | P: model-specific | M |
| GCAM | M: 32 regions | M: energy, land, water | P: physical only | P: post-processed | P | G |
| Beyond the Gap (World Bank 2019) | M: LMICs | M: infrastructure, with flood protection | M: total | G: sector models | P: 6% flat | M |
| CCDRs (World Bank) | M: 42 economies | G: country-specific | G: incremental | M: varies | M: varies | P |
| CPAT (World Bank/IMF) | G: about 200 countries | M: energy | n.a. | P: no capital module | P | M |

**The conclusion.** The IHLEG estimate is the only one that is complete in coverage, so the method uses it for the global totals. CDPM is the only open model that is complete in sectors and in accounting (costs, savings and co-benefits against a baseline), so the method uses its equations and structure. Neither has a cost of capital that differs by country. The method adds one, from the IEA Cost of Capital Observatory and CPI.

Two candidates were set aside:

* **The IAM investment variables (NGFS, McCollum et al. 2018)** are a regional cross-check, not a country method.
* **CPAT** has no capital module. It values carbon pricing and fiscal effects, not investment.

# The method

The method generalises CDPM's cost equation (Appendix A, equation A1) from cost items to capital, and adds the financing and economic layers that CDPM lacks.

## Capital required

For country $c$, sector $s$, technology $j$ and year $t$, investment in scenario $S$ is:

$$
I^{S}_{c,s,t} = \sum_{j \in s} \Delta Q^{S}_{c,j,t}\; k_{j,t}\; p_{c,s} \qquad (1)
$$

where:

* $\Delta Q$ is new physical capacity: additions plus replacements, such as GW of solar, km of transmission, EVs sold, buildings retrofitted or hectares restored;
* $k_{j,t} = k_{j,0}\, g_j^{\,t-t_0}$ is the unit capital cost, with a learning or trend factor $g_j$ as in CDPM;
* $p_{c,s}$ is the country's price level for investment goods relative to the source of $k$, taken from the ICP purchasing-power parities for gross fixed capital formation.

Incremental investment is the difference from the baseline $B$, as in CDPM:

$$
\Delta I_{c,s,t} = I^{S}_{c,s,t} - I^{B}_{c,s,t} \qquad (2)
$$

Total investment is $I^S$ itself. Both are reported. Total is the number that compares with Beyond the Gap, the IEA and the IHLEG. Incremental compares with the CCDRs and CDPM.

**Where no physical model exists**, the method uses a share-of-GDP rule by income group and sector, calibrated on the models that do exist:

$$
I_{c,s,t} = \theta_{s,\,inc(c)}\; GDP_{c,t} \qquad (3)
$$

$\theta$ comes from Beyond the Gap (infrastructure), the CCDRs (incremental climate investment) and the IHLEG (totals by country group), as listed in Table 3. This is the fallback, not the method. It is only as good as the studies it is calibrated on.

## Cost of financing

Capital is converted to an annual charge with the capital recovery factor at rate $r$ over the asset life $n_j$:

$$
A_{c,j,t} = I_{c,j,t}\; \frac{r_{c,j}}{1-(1+r_{c,j})^{-n_j}} \qquad (4)
$$

For the financial case, $r$ is the weighted average cost of capital (WACC) for that country and technology:

$$
WACC_{c,j} = \frac{D}{V}\,\big(r_f + \pi_c + \pi_j\big)(1-\tau) + \frac{E}{V}\, k^{E}_{c,j} \qquad (5)
$$

where:

* $r_f$ is the risk-free rate;
* $\pi_c$ is the country risk premium (CPI's climate investment risk premium, or a sovereign spread);
* $\pi_j$ is the technology premium;
* $D/V$ and $E/V$ are the debt and equity shares, $\tau$ the tax rate and $k^E$ the cost of equity.

For the economic case, $r$ is the World Bank social rate $r_s = \rho + \eta g \approx 6\%$ (Fay et al. 2016).

## Net cost and benefit

The financial case sets the capital charge against savings in operating and fuel costs:

$$
N^{F}_{c,t} = \sum_j \Big( \Delta F_{c,j,t} - \Delta OM_{c,j,t} - A_{c,j,t}(WACC) \Big) \qquad (6)
$$

The economic case discounts at the social rate. It adds the co-benefits $CB$ (air pollution, crashes, congestion, ecosystems, health), as CDPM does. It also adds avoided emissions $\Delta E$ valued at the World Bank shadow price of carbon $P^{C}_t$:

$$
N^{E}_{c,t} = \sum_j \Big( \Delta F_{c,j,t} - \Delta OM_{c,j,t} - A_{c,j,t}(r_s) \Big) + CB_{c,t} + P^{C}_t\, \Delta E_{c,t} \qquad (7)
$$

Equations 6 and 7 are the two numbers a finance ministry needs:

* **Equation 6** says what investors and households will see.
* **Equation 7** says what society gets.

The difference between them is what public finance, guarantees and carbon pricing have to bridge.

## What each step takes from CDPM

| Step | CDPM has it | What the method adds |
|---|---|---|
| Physical change by sector, compared with a baseline | Yes, all IPCC sectors, about 1,000 futures | Defaults from the IEA and IHLEG where CDPM has not been run |
| Unit costs with trends (equation 1) | Yes, several dozen cost factors, LAC 2019 dollars | Global and regional unit capital costs; local price factors |
| Power capital | Yes, NemoMod least cost | Rate set to the country WACC; technology interest rates implemented |
| Gross capital output | No, technical costs mix capital and operations | Reported separately |
| Cost of capital (equations 4–5) | No, one rate; technology interest rates stubbed | WACC by country and technology |
| Savings and co-benefits (equations 6–7) | Yes, IMF-based co-benefit values | Carbon at the World Bank shadow price, as a separate line |
| Uncertainty | Yes, Latin Hypercube and scenario discovery | WACC and the social rate added to the sampled inputs |

# Key parameters

Table 3 lists the parameters that drive the result, with the value used, the source and a rating against international evidence. Appendix C gives the full set by technology. Values marked † were not verified against the primary source in this review and should be confirmed before publication.

**Table 3. Key parameters of the global method**

The last column rates the parameter against international evidence: G good, M medium, P poor (it needs review before use).

| Parameter | Value used | Source | Evidence | Fit |
|---|---|---|---|---|
| Global needs, 2030 | USD 6.3–6.7 trillion a year: advanced 2.7–2.8, China 1.3–1.4, EMDCs 2.3–2.5 | IHLEG (2024) | CPI 7.4–9 trillion; McKinsey 9.2 trillion gross; ETC 3.5 trillion | G: most complete |
| Clean energy, 2030 | USD 4.8 trillion a year (EMDCs 1.6) | IHLEG (2024), from IEA | IEA NZE 4.5 trillion; IEA–IFC EMDEs 2.2–2.8 trillion | G |
| Adaptation, EMDCs | USD 0.2–0.3 trillion a year | UNEP AGR (2023) | UNFCCC NDR 0.46–0.58 trillion (all needs) | M: wide range |
| Nature and agriculture, EMDCs | USD 0.3–0.4 trillion a year | IHLEG (2024) | CDPM reforestation USD 45/ha a year | M |
| Share of GDP, fallback (θ) | Total infrastructure 4.5% (2–8%) in LMICs; incremental climate 1.4% (under 1 to 10%) | Beyond the Gap; CCDR synthesis | CDPM 1–2%; IHLEG EMDCs 6.5% (total) | M: scope differs |
| Unit capital cost, utility solar | About USD 760/kW (2023) | IRENA (2024)† | IEA regional values 500–1,200 | G |
| Unit capital cost, onshore wind | About USD 1,160/kW (2023) | IRENA (2024)† | IEA regional values | G |
| Grids | About USD 0.6–1 of grid investment per USD 1 of renewables | IEA (2023) grids report† | CDPM USD 2.70/MWh (levelised) | M |
| EV premium | USD 12,000 + 1,000 charging, falling | CDPM (Baik et al. 2019) | Near parity in China by 2024 (IEA) | P: dated |
| Asset life | 25 years (12 vehicles; 40 grids) | Engineering norms | CDPM 12 years for EVs | G |
| Price level of investment goods (p) | ICP PPP for gross fixed capital formation | World Bank ICP | CDPM: one US-to-LAC scaling | M |
| WACC, commercial | Advanced 5.5%; China 6%†; EMDCs 10.5% (9–12%), high case 15% | IEA via IHLEG (2024); CPI (2023) | CPI cost of debt 2.8% (Germany) to 45% (Zambia) | M: by group, not country |
| Social discount rate | 6% (range 2–10%) | Fay et al. (2016) | Beyond the Gap 6%; CDPM 7% | G |
| Financial savings per USD of capital (b) | 0.133 a year | Calibrated on CDPM LAC (benefit-cost ratio 1.70 at 6%) | Implies break-even WACC of 12.6% | M: one region |
| Co-benefits per USD of capital (c) | 0.129 a year | Calibrated on CDPM LAC | IHLEG 15–18% of GDP; Markandya 1.4–2.45 times costs | M: one region |
| Abatement per USD 1,000 of capital (a) | 0.35 tCO2e a year | Author's calculation: about 37 Gt over about USD 110 trillion (IEA NZE)† | CDPM LAC higher (AFOLU is cheap) | P: needs review |
| Shadow price of carbon | USD 50–100/tCO2e (2030) | World Bank (2017, 2024) | US EPA about USD 190 | G |

Three parameters need the most review, and in this order:

1. **The EMDC cost of capital.** It decides whether the financial case holds. The IEA Cost of Capital Observatory and CPI give it by country and technology. The method should use those values, not one rate per group.
2. **The savings and co-benefits per dollar (b and c).** They are calibrated on one region. A CDPM run per country, or per region, would replace them.
3. **Abatement per dollar (a).** It mixes energy, which is expensive per tonne, with AFOLU, which is cheap per tonne. It should be computed by sector.

# Global investment needs

Table 4 gives the global needs by country group and sector. Table 5 compares them with the other estimates once scope is aligned. Appendix B gives the detail.

**Table 4. Global climate investment needs, USD trillion a year (2023 prices)**

| Group | 2030 total | Clean energy, 2030 | Adaptation and resilience | Loss and damage | Nature and agriculture | Just transition | 2035 total |
|---|---|---|---|---|---|---|---|
| Advanced economies | 2.7–2.8 | Within the global 4.8 | n.a. | n.a. | n.a. | n.a. | 2.6–3.1 |
| China | 1.3–1.4 | Within the global 4.8 | n.a. | n.a. | n.a. | n.a. | 1.3–1.5 |
| EMDCs other than China | 2.3–2.5 | 1.6 | 0.2–0.3 | 0.25 or more | 0.3–0.4 | 0.04–0.07 | 3.1–3.5 |
| World | 6.3–6.7 | 4.8 | | | | | 7.0–8.1 |

Source: Bhattacharya et al. (2024), the third IHLEG report. Clean energy is from the IEA. Adaptation is from UNEP. The IHLEG does not split the advanced economy and China totals by sector. By residual, about USD 1.7 trillion of the world total is outside clean energy (author's calculation).

How to read the table:

* **Total and current.** These are total needs. Global climate finance was USD 1.3 trillion in 2021–2022 (CPI), and clean energy investment USD 1.9 trillion in 2023 (IEA). Needs are about four times current flows in EMDCs other than China.
* **Share of GDP.** At an assumed world GDP of about USD 125 trillion in 2030 (author's assumption), USD 6.5 trillion is about 5% of world GDP. The IHLEG gives 6.5% of GDP for EMDCs other than China.
* **Incremental.** The CCDRs give 1.4% of GDP for incremental climate investment to 2030, and CDPM about 2% for LAC. The IHLEG totals include investment that would happen anyway.

**Table 5. The global needs compared with the other estimates, once aligned**

| Estimate | As published | Aligned to the method |
|---|---|---|
| IHLEG (2024) | USD 6.3–6.7 trillion a year, world, 2030 | The method's default |
| IEA NZE (2023) | USD 4.5 trillion a year, clean energy, early 2030s | Matches IHLEG clean energy (4.8) |
| CPI GLCF (2023, 2024) | USD 7.4–9 trillion a year to 2030 | Higher: broader scope, average of models |
| McKinsey (2022) | USD 9.2 trillion a year gross; +3.5 trillion over today | Higher: all physical assets, including baseline |
| ETC (2023) | USD 3.5 trillion a year gross to 2050 | Lower: average over 2025–2050, energy |
| Beyond the Gap (2019) | 4.5% of GDP (2–8%), LMICs, all infrastructure | Total; includes non-climate infrastructure |
| CCDR synthesis (2023) | 1.4% of GDP incremental; USD 574 billion a year, LICs and MICs other than China | Incremental share of IHLEG EMDC total (about a quarter) |
| CDPM LAC (2023) | USD 1.3 trillion present value, about 2% of GDP a year incremental | Incremental; agrees with the CCDRs |
| UNFCCC NDR 2 (2024) | USD 455–584 billion a year, costed NDC needs | Lower: partial, self-reported |

The estimates are consistent once scope is aligned:

* **Totals** (IHLEG, IEA, CPI, McKinsey) are 5–8% of GDP.
* **Increments** (CCDRs, CDPM) are 1–3% of GDP for middle-income countries.
* **Low-income and climate-exposed countries** are much higher, at 7–10% (CCDRs: Pakistan, the Sahel).

# Investment needs under three discount-rate lenses

The discount rate does two jobs that are often confused:

* It sets the **cost of financing**: what an investor must earn, the WACC (equation 5).
* It sets the **social value of time**: how much less society values a dollar of benefit in 2045 than today.

The three lenses separate these jobs.

* **(a) Commercial rates.** The WACC each group faces today:
  * central case: advanced economies 5.5%, China 6%, EMDCs 10.5% (the IHLEG and IEA range for solar is 9–12%);
  * high case: 7%, 8% and 15%.
* **(b) World Bank rate.** 6% everywhere, the World Bank default under the Ramsey rule (Fay et al. 2016). Beyond the Gap and the IMF also use 6%.
* **(c) World Bank rate plus co-benefits.** 6%, with the co-benefits that CDPM counts (air pollution, crashes, congestion, ecosystems, health), plus avoided emissions valued at the World Bank shadow price of carbon of USD 50–100/tCO2e in 2030.

The lenses are applied to one year's investment, the 2030 need of USD 6.5 trillion. Its capital charge (equation 4, 25-year life) is set against savings and co-benefits, using the parameters in Table 3. All items are annual flows over the asset life.

**Table 6. Global investment needs, 2030, under the three lenses (USD billion a year, for one year's investment of USD 6.5 trillion)**

| Lens | Rate | Capital charge | Financial savings | Net financial benefit | Co-benefits | Carbon at the shadow price | Total net benefit | Benefit-cost ratio |
|---|---|---|---|---|---|---|---|---|
| (a) Commercial, central | 5.5% / 6% / 10.5% | 585 | 865 | 280 | (not counted) | (not counted) | 280 | 1.48 |
| (a) Commercial, high | 7% / 8% / 15% | 734 | 865 | 131 | (not counted) | (not counted) | 131 | 1.18 |
| (b) World Bank | 6% | 508 | 865 | 356 | (not counted) | (not counted) | 356 | 1.70 |
| (c) World Bank plus co-benefits | 6% | 508 | 865 | 356 | 839 | (not counted) | 1,195 | 3.35 |
| (c) plus carbon | 6% | 508 | 865 | 356 | 839 | 114–227 | 1,310–1,420 | 3.6–3.8 |

The commercial rates are for advanced economies, China and EMDCs, in that order.

**Table 7. The same, by country group (USD billion a year)**

| Group | Investment, 2030 | (a) Commercial charge | (a) Net financial | (a) High: net financial | (b) World Bank charge | (b) Net financial | (c) Net with co-benefits |
|---|---|---|---|---|---|---|---|
| Advanced economies | 2,750 | 205 (5.5%) | 161 | 130 (7%) | 215 | 151 | 506 |
| China | 1,350 | 106 (6%) | 74 | 53 (8%) | 106 | 74 | 248 |
| EMDCs other than China | 2,400 | 275 (10.5%) | 45 | −52 (15%) | 188 | 131 | 441 |
| World | 6,500 | 585 | 280 | 131 | 508 | 356 | 1,195 |

The figures are the author's calculations from equations 4, 6 and 7 and Table 3. They are illustrative. The savings and co-benefit ratios are calibrated on CDPM's LAC results and applied uniformly, so the split by group reflects the cost of capital only.

**What the lenses show.**

**(a) At commercial rates, the world still gains, but EMDCs barely do.**

* At 10.5%, the USD 2.4 trillion of EMDC investment returns only USD 45 billion a year more than it costs to finance. At 15% it loses money.
* Financing costs USD 87 billion a year more than at 6% (USD 184 billion more at 15%), for each year's investment. That is the cost-of-capital penalty.
* The penalty accumulates as investment builds up. Per point of WACC, it is about USD 40 billion a year for each year's clean energy investment of USD 4.8 trillion. After about four years of investment that matches the IEA's estimate of USD 150 billion a year per point.
* At high rates an optimiser also changes the mix toward less capital-intensive options (Ameli et al. 2021), so the physical needs change too. Only a model with the cost of capital inside it, such as CDPM's NemoMod once its interest rates are implemented, can show this.

**(b) At the World Bank's 6%, the financial case holds in every group.**

* The benefit-cost ratio is 1.7.
* The investment need in capital is unchanged at USD 6.3–6.7 trillion. What changes is the cost of carrying it.
* This is the lens to use when comparing with Beyond the Gap and the CCDRs.

**(c) With co-benefits at 6%, the economic case is large.**

* The net benefit is about USD 1.2 trillion a year for each year's investment, and USD 1.3–1.4 trillion with carbon.
* Co-benefits are about as large as the fuel savings. This is consistent with Markandya et al. (2018), who find health co-benefits 1.4–2.45 times mitigation costs, and with the IHLEG's 15–18% of GDP.
* Co-benefits are public: no investor collects them.

**The gap between (a) and (c) is the case for public finance.** In EMDCs, the private investor at 10.5% sees a gain of USD 45 billion a year, and at 15% a loss. Society, at 6% and counting co-benefits, sees USD 441 billion. Three instruments close the gap:

* **Cutting the cost of capital.** CPI's Global Credit Guarantee Facility (USD 6.6 billion, 10% paid in) and FX guarantees target exactly this.
* **Pricing carbon and air pollution**, which turns co-benefits into revenue.
* **Lending on IBRD terms.** Kharas and Rivard (2022) show that financing half of a big push externally at IBRD terms (about 5.2%) keeps debt manageable for most country groups. At market rates, average debt rises from about 52% to 135% of GDP.

# Using the method

**For a country.** There are three tiers:

1. **Where CDPM has been run** (18 LAC countries; Uganda, Morocco and Libya in progress):
   * take $\Delta Q$ and the baseline from CDPM;
   * apply equation 1 with global unit costs and the country price factor;
   * report gross capital separately;
   * apply equations 4–7 at the country WACC and at 6%.
2. **Where an energy model exists but not CDPM** (OSeMOSYS Global, a CCDR model, the IEA): use its power and energy quantities, and the fallback rule (equation 3) for AFOLU, adaptation and infrastructure.
3. **Elsewhere**: use the fallback rule for all sectors, with the IHLEG group totals as a check.

**For the global total.** Sum the country results, and check them against the IHLEG group totals (Table 4) and the IAM regional ranges (NGFS, McCollum et al. 2018).

**What to build first.** Five steps, in order of their effect on the result:

1. **Country WACC by technology**, from the IEA Cost of Capital Observatory and CPI, as an open table. This is the method's main original contribution: no open model has it.
2. **CDPM fixes** (Appendix A.5): a dual discount rate; consistent power-capital discounting; technology interest rates in NemoMod; capital recovery factors in place of straight-line charges; a gross capital output.
3. **Global unit costs** (IRENA, IEA, NREL's Annual Technology Baseline), with ICP price factors, replacing the US-to-LAC scaling.
4. **CDPM runs for the largest EMDC emitters outside LAC** (India, Indonesia, South Africa, Türkiye, Viet Nam, Egypt), to replace the calibrated b, c and a with country values.
5. **A financing layer**: the split between public and private, and between domestic and external, as in the IHLEG, linked to a debt module such as MFMod or Kharas and Rivard (2022).


# Appendix A. CDPM

This appendix reviews how CDPM estimates investment needs, its core parameters, and its published results, and re-bases those results under the three lenses. It is the model whose equations the global method generalises. Equations and tables are numbered A1, A2 and so on.

## A.1 How CDPM estimates investment needs

### The model in one paragraph

CDPM represents an economy as a set of sector models: agriculture, forestry and land use; livestock; waste and wastewater; industrial processes; buildings; transport; fugitive emissions; and energy production. Each is driven by population and GDP. Users define *transformations* (more than 50 in the LAC study, such as "electrify light-duty vehicles", "produce electricity with renewables" or "halt deforestation"). These are combined into *strategies*. The model is calibrated to historical emissions (Climate Watch) by minimising mean-squared error with a genetic algorithm and particle swarm optimisation. It then runs each strategy across about a thousand futures drawn by Latin Hypercube sampling of uncertain inputs. This is Robust Decision Making in the sense of Lempert et al. Energy production is optimised by NemoMod, an open-source least-cost energy model in Julia. The cost-benefit module (`costs_benefits_ssp`) then converts physical differences between each strategy and the baseline into money.

### The core equation

For each cost or benefit item $k$ in year $t$, the module compares the strategy run $S$ with the baseline run $B$:

$$
V_{k,t} = \left(x^{S}_{k,t} - x^{B}_{k,t}\right)\, m_k\, g_k^{\max(0,\,t-8)} \qquad (A1)
$$

where $x_{k,t}$ is the driving physical variable (for example vehicle-km by electric light-duty vehicles, hectares reforested, tonnes of CO2e of fugitive emissions abated, or petajoules of fuel saved), $m_k$ is the unit cost factor in 2019 US dollars, $g_k$ is an annual change in the factor (for example a falling EV cost premium), and $t$ is the model period. Period 0 is 2015, so the factor starts changing from 2023.

Some items use one of two variants:

* *scale_variable*: cost is proportional to a level rather than a difference.
* *fraction_change*: cost is proportional to the change in a share, such as the fraction of waste recycled.

Two post-processing steps follow:

* `cb_process_interactions` rescales overlapping transformations so that, for example, efficiency savings are not counted twice when vehicles are also electrified.
* `cb_shift_costs` moves some costs forward to the year of investment.

The module itself never discounts. It returns undiscounted annual flows by item, country and future. The present value is computed afterwards:

$$
PV = \sum_{t=2025}^{2050} \sum_{k} \frac{V_{k,t}}{(1+r)^{\,t-t_0}} \qquad (A2)
$$

In the LAC study $r = 7\%$ for every flow, country and future. The report calls this a "higher bound" (citing Emmerling et al. 2019) and does not explore other rates, because "it does not have an ordinal impact" on the ranking of strategies. That is true for ranking but not for the *level* of needs or the sign of the financial case (Section A.3).

### Power-sector capital: an embedded optimisation

Electricity is different. NemoMod chooses the least-cost capacity mix to meet demand and reports two discounted variables, `vdiscountedcapitalinvestment` and `vdiscountedoperatingcost`. CDPM stores them as `nemomod_entc_discounted_capital_investment_*` and `nemomod_entc_discounted_operating_costs_*`. The cost-benefit database maps power capex and opex directly onto these discounted series.

The NemoMod discount rate comes from CDPM's analytical parameters:

* `attributes/analytical_parameters.csv` documents a default of 5%, "used for analyzing all costs and benefits … passed to NemoMod (NemoMod_DiscountRate) to find the least cost solution".
* The shipped configuration (`sisepuede.config`) sets `discount_rate: 0.0`.

This has three consequences:

1. The *capacity mix* NemoMod chooses depends on the rate. At 0% capital-intensive renewables look cheaper than they would at a commercial cost of capital. Ameli et al. (2021) show the effect can be large: using region-specific WACC instead of a uniform rate cuts green electricity in Africa by 35% in a 2°C pathway.
2. Power capex arrives in the cost-benefit module *already discounted* to NemoMod's base year at NemoMod's rate. Other cost lines are undiscounted. If equation A2 is then applied to everything, power capex is either discounted twice (at a non-zero NemoMod rate) or once at the reporting rate (at 0%). The documentation does not say which convention the LAC run used.
3. `InterestRateTechnology` and `InterestRateStorage`, NemoMod's inputs for financing costs that differ from the social discount rate, are stubbed in CDPM (their formatters return nothing). Depreciation defaults to the sinking-fund method. So there is no way to give a solar farm in Zambia a higher cost of capital than one in Chile.

### Cost factors outside power

The cost factors for everything other than power are fixed unit costs in 2019 LAC dollars. Table A1 lists the most influential ones; Section A.6 gives the full set. Three features matter for investment needs.

* **They are mostly US-derived.** EV costs come from Baik et al. (2019) and AAA. Transmission costs come from NREL (Gorman et al. 2019, USD 1–10/MWh, average USD 5, converted to USD 2.70 for LAC). Fugitive emissions costs come from ICF (2014). Heat pump and efficiency costs come from Rissman et al. The conversion to LAC prices is a single scaling, not a country-specific price level.
* **Capital is annualised inconsistently.** The EV premium (USD 12,000 plus USD 1,000 for charging) is divided by 15,000 km a year over a 12-year life, giving USD 0.039/vkm in LAC prices. That is straight-line, equivalent to a 0% interest rate. At 6% the capital recovery factor (equation A3) over 12 years is 1.43 times the straight-line charge, at 12% it is 1.94 times, and at 15% 2.21 times. Several building and industry factors are "levelised costs as presented", and the RAND appendix notes that "no discount rate is documented" for them. N2O abatement costs from EPA are undiscounted.
* **Some infrastructure is left out.** Mode shift (freight and passenger) carries no technical cost, because the infrastructure is "deeply uncertain, highly localized, and beyond this study's scope". Transmission-loss reduction uses an IDB estimate of grid upgrade costs averaged per year. Both are biased low for the capital-heavy items that drive real investment needs.

$$
CRF(r,n) = \frac{r}{1-(1+r)^{-n}} \qquad (A3)
$$

**Table A1. Core CDPM cost parameters that drive investment needs (2019 USD, LAC average)**

The last column rates the parameter against international evidence (G good, M medium, P poor) with the main reason.

| Item | Value | Source | Fit |
|---|---|---|---|
| Discount rate, reporting | 7% (LAC study); 5% Costa Rica | Kalra et al. 2023; Groves et al. 2020 | M: one rate for all; not linked to WACC |
| Discount rate, NemoMod | 0% config; 5% documented default | sisepuede.config; analytical_parameters.csv | P: inconsistent; drives power mix |
| Interest rate by technology | Not implemented | energy_production.py | P: no cost-of-capital differentiation |
| Power capex and opex | Endogenous, NemoMod least cost | NemoMod | G: method; M: inputs per country |
| New transmission | USD 2.70/MWh | NREL (Gorman et al. 2019) | M: US-based; levelised |
| Reduce transmission losses | Country-specific grid upgrade cost | IDB (Brichetti et al. 2021) | M: SDG upgrade, not losses |
| EV light-duty premium | USD 0.039/vkm capital, −0.012/vkm maintenance | Baik et al. 2019; AAA 2019 | P: straight-line capital, 0% interest |
| Electric trucks and buses | USD 0.042/vkm + 0.011/kWh capital; −0.02/vkm maintenance | Burke et al. 2022 | M: US costs |
| Rail electrification | USD 0.0013 per tkm or pkm | Popovich et al. 2021 | M |
| Maritime fuel switch | USD 0.0005/tkm | Carlo et al. 2020 | M: ammonia-based |
| Mode shift | No technical cost | Assumption | P: omits infrastructure |
| Fugitive leaks; flaring | USD 20/tCO2e; USD 2/tCO2e | ICF 2014 | G: conservative |
| CCS cement / steel / chemicals | USD 40 / 50 / 100 per tCO2 | Literature | M: low for EMDEs |
| Reforestation | USD 45/ha a year | Literature | M |
| Diet health benefit | USD 1,750 per person shifting | Springmann et al. | M: range 350–3,500 |
| Air pollution, coal / gas | USD 2.77 / 0.99 per GJ | IMF (Parry et al. 2014) | G: IMF standard; dated |
| Transport pollution, diesel / gasoline | USD 0.31 / 0.039 per litre | IMF (Parry et al. 2014) | G: dated |
| Crashes and congestion | Per vkm, IMF values | Parry et al. 2014 | G |
| Jobs | Input-output multipliers | ECLAC/ADB 2011 MRIO | M: dated table |

### Jobs and the wider economy

CDPM estimates employment effects with a multi-regional input-output table (ECLAC/ADB 2011): the cost and saving streams are mapped onto sectors and multiplied through. It has no macroeconomic feedback. Investment does not raise growth, prices do not adjust, and financing does not crowd anything out. In this it differs from the CCDRs' macro models (MFMod, Envisage, MANAGE) and from Kharas and Rivard (2022), where growth is the channel through which a big push pays for itself.

### Uncertainty

The LAC study runs each strategy across about a thousand futures, varying technology costs, fuel prices, demand drivers and the effectiveness of transformations. It reports the distribution of net benefits and finds the vulnerabilities with scenario discovery (PRIM). This is a real strength over the point estimates of most global studies. The discount rate and the cost of capital, though, are *not* among the uncertain inputs, and they are among the most influential. Costa Rica's earlier study varied 3,000 futures and found benefits exceeded costs in all but 22.

## A.2 What CDPM finds

### Latin America and the Caribbean

Table A2 summarises the published results (Kalra et al. 2023). All values are present values over 2025–2050 in 2019 US dollars at 7%.

**Table A2. CDPM results for 18 LAC countries**

| Result | Value |
|---|---|
| Baseline emissions | 4.1 GtCO2e today, rising to 6.9 GtCO2e by 2050 |
| All Actions strategy | Net zero by about 2040 |
| Additional investment needs (technical costs) | USD 1.3 trillion (range 0.5–1.5 trillion across strategies and futures) |
| Annual equivalent (author's calculation) | About USD 110 billion a year, about 2% of regional GDP |
| Fuel savings | USD 0.9 trillion |
| Avoided air pollution | USD 0.5 trillion |
| Health, safety and productivity | USD 1.0 trillion |
| Total net benefit | USD 2.7 trillion |
| Net benefit, financial flows only | USD 0.7 trillion |
| Median net benefit across about 1,000 strategies and futures | USD 1.0 trillion |
| Incremental Improvements strategy, net benefit | USD 1.8 trillion |
| Buildings: all actions / electrification only | +USD 150 billion / −USD 480 billion |
| Industry efficiency | +USD 160 billion |

The annual equivalent uses an annuity factor of 11.8 (26 years at 7%) and assumes the present value is taken to 2025. Regional GDP for the 18 countries is about USD 5.5 trillion.

The critical actions are renewable electricity, transport electrification and halting deforestation. The study excludes the social cost of carbon from the headline benefits, on the grounds that decarbonisation is a goal in itself. The waste-sector appendix nonetheless values avoided emissions at the social cost of carbon, an inconsistency the documentation should resolve.

### Countries

**Table A3. CDPM country results**

| Country | Investment needs | Share of GDP a year | Net benefit | Rate, period | Source |
|---|---|---|---|---|---|
| Brazil | USD 709 billion | 1.92% of today's GDP | USD 1.0 trillion | 7%, 2025–2050 | Kalra et al. 2023 |
| Mexico | USD 272 billion | About 1% | USD 659 billion | 7%, 2025–2050 | Kalra et al. 2023 |
| Dominican Republic | USD 30 billion | 1.6% | USD 59 billion | 7%, 2025–2050 | Kalra et al. 2023 |
| Costa Rica | Not reported as a single figure | n.a. | USD 41 billion | 5%, 2020–2050 | Groves et al. 2020 |
| Colombia, Peru, Chile | Earlier RAND/IDB studies | n.a. | USD 330 / 140 / 7 billion | Various | Cited in Kalra et al. 2023 |

The country spread (1–2% of GDP a year) is close to the CCDR average of 1.4%. The research for this note found no CCDR that names CDPM as its model. The CCDR synthesis lists Envisage, CGEBox, MANAGE, MFMod, the LTGM, CPAT and E3ME. Any World Bank use of CDPM in a CCDR should be confirmed with the country teams.

## A.3 CDPM's LAC results under the three lenses

### The three lenses

The discount rate does two jobs that are often confused:

* It sets the **cost of financing**: what an investor must earn, which is the WACC.
* It sets the **social value of time**: how much less society values a dollar of benefit in 2045 than today.

**(a) Commercial rates.** The WACC a private investor faces:

$$
WACC = \frac{D}{V}\, k_d\,(1-\tau) + \frac{E}{V}\, k_e \qquad (A4)
$$

where $D/V$ and $E/V$ are the debt and equity shares, $k_d$ the cost of debt, $k_e$ the required return on equity and $\tau$ the tax rate. The evidence on how far these rates diverge:

* CPI (Gautam et al. 2023) estimates, for solar in January 2023:
  * cost of debt: Germany 2.8%, United States 5.3%, Brazil 7.8%, India 11.4%, South Africa 20.3%, Egypt 29.5%, Zambia 45.4%;
  * required equity return: from 8.3% in Germany to 51% in Zambia;
  * climate investment risk premium: 1–29%.
* The IHLEG (2024), citing the IEA, puts solar WACC at 9–12% in EMDCs against 5–6% in advanced economies. It puts the average cost of debt at 21% in EMDCs against 6% in advanced economies.
* The IEA and IFC (2023) find the cost of capital for utility-scale solar two to three times higher in key emerging economies.
* The IEA (2021) finds financing costs "up to seven times" higher than in the United States and Europe.

This note uses **10%, 12% and 15%** to span typical commercial WACC for clean energy in middle-income EMDEs.

**(b) World Bank rate.** The World Bank's guidance (Fay et al. 2016) applies the Ramsey rule:

$$
r = \rho + \eta\, g \;\approx\; 2g \qquad (A5)
$$

where $\rho$ is pure time preference, $\eta$ the elasticity of marginal utility and $g$ the growth of consumption per person. With long-run growth of 3% a year this gives a default of **6%**. The interquartile range across countries is 2–10%, and the extremes are 0–12%. Beyond the Gap and the IMF's GFSR (2022) use 6%. For comparison:

* Stern (2006) implies about 1.4%.
* Nordhaus's DICE uses about 4–4.5%.
* The UK Green Book uses 3.5%.
* US OMB Circular A-4 (2023) uses about 2%.

**(c) World Bank rate plus co-benefits.** The same 6%, with all of CDPM's non-financial benefits (air pollution, crashes, congestion, ecosystems, diet-related health) and, optionally, avoided emissions valued at the World Bank shadow price of carbon. That price is USD 40–80/tCO2e in 2020 and 50–100 in 2030, rising 2.25% a year to about 78–156 in 2050. It has been mandatory in IBRD/IDA project appraisal since July 2017. International evidence supports counting co-benefits:

* Markandya et al. (2018) find health co-benefits 1.4–2.45 times mitigation costs.
* Vandyck et al. (2018) cite USD 2–196/tCO2, with a mean of USD 49.
* Parry and Black (2021) find efficient fuel pricing would avoid 0.9 million air-pollution deaths a year.
* Karlsson et al. (2020) review 239 studies and find air-quality co-benefits often of the same order as mitigation costs.

### Method for re-basing CDPM

CDPM's annual flows are not published, so the LAC results are re-based with two stylised profiles:

* **Costs** are front-loaded and roughly flat from 2025.
* **Benefits** ramp up linearly as the capital stock turns over.

For each rate $r$, the present value at 7% is multiplied by the ratio of the present-value factors. For lens (a), 70% of the technical cost is treated as capital and re-annuitised at the WACC with equation A3 over 20 years, to reflect the higher financing charge. These are author's calculations and are illustrative. They should be replaced by a direct re-run of CDPM when the annual flows are available.

**Table A4. Present-value factors relative to CDPM's 7%**

| Rate | Flat cost stream | Ramping benefit stream | CRF, 20 years | Annual charge per USD 1 trillion of capital |
|---|---|---|---|---|
| 0% | 2.06 | 2.81 | 0.050 | USD 50 billion |
| 3% | 1.46 | 1.75 | 0.067 | USD 67 billion |
| 6% (World Bank) | 1.09 | 1.14 | 0.087 | USD 87 billion |
| 7% (CDPM LAC) | 1.00 | 1.00 | 0.094 | USD 94 billion |
| 10% | 0.80 | 0.70 | 0.118 | USD 118 billion |
| 12% | 0.70 | 0.56 | 0.134 | USD 134 billion |
| 15% | 0.59 | 0.42 | 0.160 | USD 160 billion |

Benefits fall faster than costs as the rate rises, because they come later. That is the whole story of the commercial lens.

### Results

**Table A5. CDPM LAC All Actions re-based under the three lenses (USD trillion, present value 2025–2050, 2019 prices)**

| Rate | Lens | Cost, adjusted | Financial benefit | Net financial benefit | Total net benefit (with co-benefits) | Financial benefit-cost ratio |
|---|---|---|---|---|---|---|
| 3% | Low social | 1.51 | 3.49 | 1.98 | 5.09 | 2.31 |
| 6% | (b) World Bank | 1.34 | 2.28 | 0.94 | 3.15 | 1.70 |
| 7% | CDPM published | 1.30 | 2.00 | 0.70 | 2.70 | 1.54 |
| 10% | (a) Commercial, low | 1.21 | 1.39 | 0.18 | 1.74 | 1.15 |
| 12% | (a) Commercial, central | 1.17 | 1.12 | −0.06 | 1.33 | 0.95 |
| 15% | (a) Commercial, high | 1.14 | 0.83 | −0.31 | 0.90 | 0.73 |

At 7% the table reproduces the published figures: a net financial benefit of USD 0.7 trillion and a total net benefit of USD 2.7 trillion.

**Table A6. Lens (c): adding avoided emissions at the World Bank shadow price of carbon (USD trillion, present value)**

| Rate | Total net benefit before carbon | Value of avoided emissions (low–high price) | Total net benefit with carbon |
|---|---|---|---|
| 6% (World Bank) | 3.15 | 2.4–4.8 | 5.5–7.9 |
| 10% | 1.74 | 1.4–2.8 | 3.1–4.5 |
| 15% | 0.90 | 0.8–1.6 | 1.7–2.5 |

The carbon value assumes illustrative abatement of about 93 GtCO2e over 2025–2050, rising from zero to the gap between the baseline (6.9 Gt in 2050) and net zero.

### What the lenses show

**(a) At commercial rates the private financial case for the All Actions strategy breaks down at about 12%.** That is the WACC many EMDE clean energy projects face today. The investment need itself rises: the same assets cost more to finance. Each USD 1 trillion of capital costs USD 134–160 billion a year at 12–15%, against USD 87 billion at 6%, so financing adds USD 47–73 billion a year per trillion. The IEA's estimate that a 1 percentage point fall in the cost of capital saves about USD 150 billion a year in clean energy financing globally is consistent with this.

Commercial rates would also change the mix. At high WACC an optimiser such as NemoMod builds less capital-intensive renewable capacity (Ameli et al. 2021), so physical needs, emissions and fuel bills all change. CDPM cannot show this today, because its NemoMod rate is a single social rate.

**(b) At the World Bank's 6% the financial case is robust.** The net financial benefit is about USD 0.9 trillion and the benefit-cost ratio 1.7. The investment need is about USD 1.3 trillion in present value, or about 2% of GDP a year. This is the lens closest to Beyond the Gap and the CCDRs, and the one to use when comparing CDPM with them.

**(c) With co-benefits at 6%, the economic case is overwhelming.** The total net benefit is about USD 3.2 trillion, and USD 5.5–7.9 trillion with carbon valued at the World Bank shadow price. Even at 15%, co-benefits keep the total net benefit positive (about USD 0.9 trillion). Co-benefits are public: avoided deaths, cleaner air and fewer crashes do not appear on an investor's balance sheet.

**The gap between (a) and (c) is the case for public finance.** At 12% the private investor sees a small loss. Society, valuing the same flows at 6% and counting co-benefits, sees a gain of more than USD 3 trillion. Closing that gap is what concessional finance, guarantees and carbon pricing are for:

* CPI's Global Credit Guarantee Facility (USD 6.6 billion, 10% paid in) is designed to cut the risk premium by about 6 percentage points.
* Persaud's FX guarantee targets currency risk.
* Kharas and Rivard (2022) show the debt path depends heavily on the rate. Financing half of a 4-point-of-GDP big push externally at IBRD terms (about 5.2%) raises debt about 20 points of GDP by 2050. At market rates average debt rises from about 52% to 135% of GDP.

Two caveats apply:

* The stylised profiles overstate precision. The direction is robust; the break-even rate (11–13%) is approximate.
* The "financial benefit" here includes fuel savings to all sectors, households included, not just to the investor. A strict project-finance view would show an even weaker private case at commercial rates.

## A.4 Recommendations for CDPM

1. **Make the discount rate explicit and dual.** Report economic results at the World Bank social rate (6% by default, or 2× country growth per Fay et al. 2016) and financial results at a country-specific WACC by technology class (CPI, IEA Cost of Capital Observatory). Sensitivity at 3% and 12% should be standard output, and the discount rate should be added to the uncertain inputs in the futures.
2. **Fix the NemoMod inconsistency.** Either pass undiscounted capital investment from NemoMod and discount everything once in the cost-benefit step, or document that power capex is pre-discounted and exclude it from the second discounting. Align `sisepuede.config` (0%) with the documented default (5%).
3. **Implement `InterestRateTechnology` and `InterestRateStorage`** so that NemoMod's least-cost mix reflects the country's cost of capital, and run the "de-risked" case (WACC cut by 2–6 points) as a transformation. This would let CDPM estimate the value of guarantees directly.
4. **Annualise capital consistently.** Replace straight-line per-km charges (EVs, trucks) with capital recovery factors at the chosen rate (equation A3), and document the rate behind every levelised cost factor.
5. **Report gross capital expenditure separately** from operating costs and savings, by sector and year, so that CDPM's results can be compared directly with Beyond the Gap, the IEA and the IHLEG, and fed into financing models.
6. **Add the missing infrastructure**: mode-shift infrastructure (rail, bus rapid transit), distribution grids and storage beyond NemoMod's scope, and adaptation.
7. **Localise unit costs** with country price levels and, where possible, local capex data (IRENA cost databases, MDB project data), not a single US-to-LAC scaling.
8. **Add a financing layer**: split investment into public and private, domestic and external, and link it to a debt module (as in Kharas and Rivard 2022, or the CCDRs' MFMod), so that the model can say who pays and what it does to debt.
9. **Resolve the social-cost-of-carbon inconsistency** between the headline (excluded) and the waste appendix (included), and report the carbon value as a separate line.

## A.5 Issues found in CDPM's treatment of investment needs

| # | Issue | Where | Effect | Fix |
|---|---|---|---|---|
| 1 | Config discount rate 0% vs documented 5% | sisepuede.config; analytical_parameters.csv | Power mix and capex differ by run | Align; document |
| 2 | Power capex pre-discounted, other costs undiscounted | energy_production.py; cb_data.db mapping | Double or inconsistent discounting | Discount once, in one place |
| 3 | No country or technology cost of capital | InterestRate formatters return None | Overstates renewables in high-risk countries; no de-risking analysis | Implement; use CPI or IEA WACC |
| 4 | Straight-line capital in EV and truck factors | transport_trns_tech_costs | Capital charge understated 40–120% at 6–15% | Use capital recovery factors |
| 5 | Undocumented levelisation rates | Building and industry factors | Unknown bias | Document or re-derive |
| 6 | No mode-shift infrastructure cost | transport_trns_tech_costs | Needs biased low | Add rail and BRT capex |
| 7 | Gross capex not reported | Cost-benefit outputs | Not comparable with other estimates | Separate capex output |
| 8 | US unit costs scaled to LAC | Most cost factors | Country price bias | Localise |
| 9 | Social cost of carbon excluded in headline, used in waste | LAC report vs appendix | Inconsistent benefits | Separate carbon line |
| 10 | Discount rate not in the uncertainty ensemble | Futures design | Most influential input left fixed | Add to the Latin Hypercube sample |

## A.6 Core CDPM parameters for investment needs

**Model structure and settings**

| Parameter | Value | Source in code or documentation |
|---|---|---|
| Base period | 2015 = period 0 | CDPM time periods |
| Transformation start | 2025 | LAC study |
| Horizon | 2050 (36 periods) | Model attributes |
| Currency | 2019 USD, LAC average | costs_benefits_ssp docs |
| Cost-benefit equation | Difference × unit factor × annual change | cb_calculate.py (lines 1139–1193) |
| Variants | scale_variable; fraction_change | cb_calculate.py (1199–1330) |
| Interactions | Rescaled to avoid double counting | cb_process_interactions (905–995) |
| Cost shifting | Early costs moved forward | cb_shift_costs (997–1017) |
| Discounting in module | None | cb_calculate.py |
| Discount rate, config | 0.0 | sisepuede.config, line 8 |
| Discount rate, documented default | 0.05 | attributes/analytical_parameters.csv |
| Discount rate, LAC report | 7% | Kalra et al. 2023 |
| Discount rate, Costa Rica | 5% | Groves et al. 2020 |
| Interest rates, technology and storage | Stubbed, not implemented | energy_production.py (5876–5910) |
| Depreciation method | 1 (sinking fund) | energy_production.py |
| Power capex variable | nemomod_entc_discounted_capital_investment_* | energy_production.py (8995–9180) |
| Calibration | MSE vs Climate Watch; genetic algorithm and PSO | RAND appendices (Kalra et al. 2023) |
| Uncertainty | Latin Hypercube, about 1,000 futures; PRIM | Kalra et al. 2023 |
| Jobs | MRIO multipliers | ECLAC/ADB 2011 |

**Benefit factors**

| Benefit | Value | Source |
|---|---|---|
| Air pollution, coal | USD 2.77/GJ | IMF (Parry et al. 2014) |
| Air pollution, gas | USD 0.99/GJ | IMF (Parry et al. 2014) |
| Transport pollution, diesel | USD 0.31/litre | IMF (Parry et al. 2014) |
| Transport pollution, gasoline | USD 0.039/litre | IMF (Parry et al. 2014) |
| Crashes and congestion | IMF per-vkm values | Parry et al. 2014 |
| Diet-related health | USD 1,750 per person (range 350–3,500) | Springmann et al.; direct-cost vs value of statistical life approaches differ about 25-fold |
| Forest ecosystem services | Per hectare, literature values | RAND appendices |
| N2O abatement | EPA costs, undiscounted | RAND appendices |

Technical cost factors are listed in Table A1. The full set is in the `docs/cost_factors` folder of the costs_benefits_ssp repository. It has 21 tables covering buildings, crops, cross-cutting items, energy production, industry, land use, livestock, transport, wastewater and solid waste.


# Appendix B. Other models and estimates

## B.1 The models

**The IEA Global Energy and Climate (GEC) model.** It is the most detailed model of energy investment, covering 29 regions and every part of the energy system: supply, transformation, grids, storage and end uses.

* **How it calculates investment.** Supply-side investment is capacity additions times unit capital costs, which is gross capital. End-use investment is explicitly incremental: "the incremental spending … to acquire a piece of equipment that is more efficient than the local market average". The two concepts should not be mixed.
* **Cost of capital.** The IEA applies regional costs of capital, and publishes them for some technologies through its Cost of Capital Observatory.
* **Limits.** The model is proprietary, and its unit costs are spread across annexes of the World Energy Outlook. It does not cover AFOLU, adaptation or non-energy infrastructure.
* **Use in the method:** the source of the clean energy totals (through the IHLEG) and of unit costs and WACC. <https://www.iea.org/reports/global-energy-and-climate-model>

**OSeMOSYS Global.** An open-data generator of electricity models for up to 163 countries and 265 nodes.

* **How it calculates investment.** Investment is a decision variable in a linear programme (new capacity times capital cost, discounted at a rate that can vary by region and technology). It is the most transparent investment equation among the candidates.
* **Limits.** It covers power only; the CLEWs variants add land and water, country by country.
* **Use in the method:** an open cross-check on power capital, and a source of country cost data. <https://github.com/OSeMOSYS/osemosys_global>

**MESSAGEix-GLOBIOM and the IAM ensemble (McCollum et al. 2018; NGFS).**

* **What it is.** McCollum et al. (2018) compare energy investment needs across six integrated assessment models for Paris-consistent and SDG pathways. Beyond the Gap uses the same six models for power.
* **Where the data are.** The NGFS scenarios publish investment variables in the IAMC format (Investment|Energy Supply|…) from REMIND-MAgPIE, MESSAGEix-GLOBIOM and GCAM. McKinsey (2022) built its estimate on them.
* **Limits.** The results are regional (about 10–33 regions), not by country, and each model treats investment and discounting differently.
* **Use in the method:** a regional cross-check only. †The McCollum paper and the NGFS variable list could not be retrieved in this review.

**GCAM (PNNL/JGCRI).**

* **What it is.** Fully open, with 32 regions covering energy, land and water.
* **Limits.** Its native outputs are physical (EJ, Mt, km³). Investment has to be derived outside the model from capacity change times unit cost, which is CDPM's pattern.
* **Use in the method:** none directly. It is an alternative source of physical change for regions without a CDPM run.

**CPAT (World Bank and IMF).**

* **What it is.** A spreadsheet and Python tool covering about 200 countries, for carbon pricing, fuel subsidy reform, revenue, distribution and air pollution.
* **Limits.** It has no capital-investment module.
* **Use in the method:** none for investment, though its air-pollution values could replace the IMF 2014 factors in CDPM's co-benefits. †The public CPAT pages could not be reached in this review.

**E3ME and FTT (Cambridge Econometrics).**

* **What it is.** FTT's technology-diffusion models (power, transport, heat, steel) compute investment as capacity additions times unit cost, with technology shares that evolve endogenously. E3ME adds macroeconomic feedback, and is one of the CCDR models.
* **Limits.** It is partly open. †It was not reviewed in detail here.

**Global Infrastructure Hub Outlook (G20 GI Hub and Oxford Economics).**

* **What it is.** Country estimates of infrastructure needs against current trends for 56 countries, covering transport, energy, water and telecoms.
* **Use in the method:** a complement to Beyond the Gap for non-energy infrastructure. †The online tool could not be reached in this review.

**CDPM** is covered in Appendix A.

## B.2 The estimates

**Table B1. Leading estimates of investment needs**

The last column rates comparability with CDPM (G good, M medium, P poor).

| Study | Headline | Base | Method | Comparability |
|---|---|---|---|---|
| CDPM LAC (2023) | USD 1.3 trillion PV incremental, about 2% of GDP a year; net benefit USD 2.7 trillion | 18 LAC countries, 2025–2050, 7% | Economy-wide sector models; NemoMod; about 1,000 futures | n.a. |
| Beyond the Gap (2019) | 2–8% of GDP, 4.5% central (USD 1.5 trillion a year) | LMICs, 2015–2030, 6% | Bottom-up sector models; six IAMs for power | M: total, infrastructure only |
| CCDR synthesis (2023) | +1.4% of GDP a year (under 1 to 10%); USD 574 billion a year (2.8% of GDP) for LICs and MICs other than China | 42 economies, to 2030 | Country models (MFMod, CGE, CPAT, E3ME) | G: incremental, country-level |
| Bhattacharya et al. (2022) | 11.3% → 18.2% of GDP, 2019–2030; +USD 1.3 trillion by 2025, +3.5 trillion by 2030 | EMDEs other than China | Synthesis of sector studies | M: includes human capital |
| IHLEG 1 (2022) | USD 2.4 trillion a year by 2030 (6.5% of GDP); USD 1 trillion external | EMDCs other than China | From Bhattacharya et al., IEA, ETC | M: total, climate scope |
| IHLEG 3 (2024) | USD 6.3–6.7 trillion a year global by 2030; EMDCs other than China USD 2.3–2.5 trillion (2030), 3.1–3.5 trillion (2035) | 2023 USD | Sector synthesis on IEA and UNEP | M: total; co-benefits 15–18% of GDP |
| IEA NZE (2023) | USD 4.5 trillion a year clean energy by the early 2030s | Global | GEC model, differentiated WACC | P: energy only, total |
| IEA–IFC (2023) | USD 2.2–2.8 trillion a year EMDEs; 1.4–1.9 trillion other than China | Early 2030s | NZE and SDS scenarios | M: energy only |
| CPI GLCF (2023/2024) | USD 8.1–9 trillion (2023 ed.); 7.4 trillion (2024 ed.) a year to 2030 | Global | Average of scenario models | P: top-down, total |
| UNFCCC NDR 2 (2024) | USD 5.0–6.9 trillion to 2030; USD 455–584 billion a year | 98 Parties, NDC costed needs | Self-reported | P: inconsistent methods |
| UNEP AGR (2023/2025) | Adaptation USD 215–387 billion a year (2030); 310–365 billion (2035) | Developing countries | Models plus NAPs | P: adaptation only |
| McKinsey (2022) | USD 9.2 trillion a year, +3.5 trillion over today; 7.5% of GDP | Global, 2021–2050 | NGFS Net Zero 2050 | P: total, capex |
| ETC (2023) | USD 3.5 trillion a year gross (3 trillion net) to 2050; 1.3% of GDP | Global | Sector analysis | M: net of fossil decline |
| IMF GFSR (2022) | USD 1 trillion a year EMDE renewables by 2030 | EMDEs | Uses Beyond the Gap, 6% | P: renewables only |
| IPCC AR6 WG3 Ch. 15 | Flows must rise 3–6 times; 4–7 times in LDCs for AFOLU | Global, 2030 | Literature review | P: multiples, not levels |

Sources and verification notes are in Appendix D.

### Aligning the numbers

To compare like with like, Figure B1 restates each estimate as an annual share of GDP and labels it total or incremental.

```
Share of GDP per year, by study (T = total, I = incremental)

CCDR synthesis (avg)            I  |####                          1.4%
CDPM LAC                   I  |######                        ~2%
CDPM Brazil / DR / Mexico  I  |#####                         1-1.9%
CCDR synthesis (LIC+MIC ex-CN)  I  |########                      2.8%
Bhattacharya et al. 2025 incr.  I  |###########                   +3.8pp
Beyond the Gap (central)        T  |#############                 4.5%
IHLEG 1 (EMDCs ex-China 2030)   T  |###################           6.5%
McKinsey (global average)       T  |######################        7.5%
Beyond the Gap (range)          T  |######  to  ########################  2-8%
```

Four points stand out.

1. **On incremental investment, CDPM and the CCDRs agree.** Both find 1–2% of GDP a year for middle-income countries, with low-income and climate-exposed countries much higher (Pakistan 9–10%, Sahel 7–8%). The agreement is not by construction, since the models are completely different. It suggests that 1–3% of GDP is a robust range for the incremental climate investment of a middle-income country.
2. **Total estimates are two to four times larger** because they include investment that development requires anyway: the baseline power plants, roads and water systems. Beyond the Gap's main finding, that decarbonised paths "need not cost more" than polluting ones, is consistent with CDPM's: the incremental cost is modest, and fuel savings largely repay it.
3. **CDPM is alone in netting costs against benefits.** Its headline is a net benefit, not a cost. That is the right question for a finance ministry. It also makes the number hard to compare: a figure of 2% of GDP in investment with a net benefit of twice that is a different message from "USD 2.4 trillion a year".
4. **Only the IEA embeds a cost of capital that differs by country**, and only the IHLEG (2024) reports one. All the others, including CDPM, value a solar farm in Zambia with the same discount rate as one in Germany. Section A.3 and the main text show why that matters.

### Strengths and gaps of CDPM relative to the others

**Strengths.**

* It covers every sector, including AFOLU, waste and diets.
* It is open source and runs at country level.
* It estimates savings and co-benefits alongside costs.
* It handles uncertainty properly, across about a thousand futures with scenario discovery.
* The transformations map directly to policy.

No other estimate in Table B1 combines these.

**Gaps.**

* It reports no gross capital requirement.
* Its discount rate is ad hoc and inconsistent between NemoMod and the cost-benefit module.
* It has no country cost of capital and no financing module (who pays, public or private, domestic or external).
* It has no macro feedback.
* Many unit costs are US values scaled to LAC.

It cannot yet say how much of the USD 1.3 trillion is capital to be financed, at what rate, and by whom. That is the question the IHLEG, the IEA–IFC and Kharas and Rivard (2022) address.


# Appendix C. Parameters of the global method

Values marked † need confirmation against the primary source. The ratings follow Table 3: G good, M medium, P poor.

**Unit capital costs and asset lives**

| Technology | Unit capital cost | Life (years) | Source | Fit |
|---|---|---|---|---|
| Utility-scale solar PV | About USD 760/kW (2023 global weighted average) | 25–30 | IRENA (2024)† | G |
| Onshore wind | About USD 1,160/kW | 25 | IRENA (2024)† | G |
| Offshore wind | About USD 2,800/kW | 25 | IRENA (2024)† | G |
| Battery storage, utility | About USD 250–300/kWh installed | 15 | IRENA; BNEF† | M |
| Transmission and distribution | USD 0.6–1 per USD 1 of renewables; CDPM USD 2.70/MWh | 40 | IEA (2023)†; NREL via CDPM | M |
| Power, all technologies, by country | Endogenous least cost | By technology | NemoMod in CDPM; OSeMOSYS Global | G |
| Electric light-duty vehicle | Premium USD 12,000 + 1,000 charging, falling | 12 | CDPM (Baik et al. 2019) | P: near parity in China |
| Electric trucks and buses | USD 0.042/vkm + 0.011/kWh (levelised) | 12 | CDPM (Burke et al. 2022) | M |
| Rail and bus rapid transit | Not in CDPM; fallback 1.3% of GDP for all transport | 40 | Beyond the Gap | P: missing in CDPM |
| Buildings: efficiency and heat pumps | CDPM levelised factors (rate undocumented) | 20–30 | CDPM (Rissman et al.) | P |
| Industry: CCS for cement / steel / chemicals | USD 40 / 50 / 100 per tCO2 | 25 | CDPM | M: low |
| Fugitive emissions: leaks; flaring | USD 20 / 2 per tCO2e | 15 | CDPM (ICF 2014) | G |
| Reforestation | USD 45/ha a year | n.a. | CDPM | M |
| Adaptation, all | USD 215–387 billion a year (developing countries, 2030) | n.a. | UNEP AGR (2023) | M |
| Flood protection | 0.32% of GDP (LMICs) | 40+ | Beyond the Gap | M |
| Water and sanitation | 0.55% of GDP (LMICs) | 40 | Beyond the Gap | M |

**Fallback shares of GDP (θ, equation 3)**

| Sector | Total, LMICs (Beyond the Gap central) | Incremental climate (CCDRs) |
|---|---|---|
| Electricity | 2.2% | Included in 1.4% average |
| Transport | 1.3% | Included |
| Water and sanitation | 0.55% | Included where climate-related |
| Flood protection | 0.32% | Included |
| Irrigation | 0.13% | Included |
| All | 4.5% (2–8%) | 1.4% (under 1 to 10%); 2.8% for LICs and MICs other than China |

**Cost of capital and discount rates**

| Parameter | Value | Source | Fit |
|---|---|---|---|
| WACC, solar, advanced economies | 5–6% | IEA, via IHLEG (2024) | G |
| WACC, solar, EMDCs other than China | 9–12% | IEA, via IHLEG (2024) | G |
| WACC, China | 6% (assumed close to advanced economies) | Author's assumption† | M |
| Cost of debt, average | 6% advanced; 21% EMDCs | CPI (Gautam et al. 2023) via IHLEG | G |
| Cost of equity, average | 12% advanced; 27% EMDCs | CPI via IHLEG | G |
| Climate investment risk premium | 1% (Germany) to 29% (Zambia); average 4.6% advanced, 15% EMDCs | CPI (2023) | G |
| Social discount rate | 6% default; 2–10% interquartile range by country | Fay et al. (2016) | G |
| CDPM reporting rate | 7% (LAC); 5% (Costa Rica); 0% config; 5% documented | Appendix A | P: inconsistent |

**Capital recovery factors, 25-year life (equation 4)**

| Rate | 0% | 3% | 5.5% | 6% | 7% | 9% | 10.5% | 12% | 15% |
|---|---|---|---|---|---|---|---|---|---|
| CRF | 0.040 | 0.057 | 0.075 | 0.078 | 0.086 | 0.102 | 0.114 | 0.128 | 0.155 |

At 10.5% the annual charge per dollar is 46% higher than at 6%. At 15% it is 98% higher.

**Benefit parameters**

| Parameter | Value | Source | Fit |
|---|---|---|---|
| Financial savings per USD of capital (b) | 0.133 a year | Calibrated: CDPM LAC benefit-cost ratio 1.70 at 6% × CRF 0.078 | M |
| Co-benefits per USD of capital (c) | 0.129 a year | Calibrated: CDPM LAC co-benefits 1.65 × cost at 6% | M |
| Abatement per USD 1,000 of capital (a) | 0.35 tCO2e a year | Author's calculation from IEA NZE† | P |
| Air pollution: coal / gas | USD 2.77 / 0.99 per GJ | IMF (Parry et al. 2014), via CDPM | G: dated; CPAT has newer values |
| Transport pollution: diesel / gasoline | USD 0.31 / 0.039 per litre | IMF (Parry et al. 2014) | G: dated |
| Diet-related health | USD 1,750 per person (350–3,500) | Springmann et al., via CDPM | M |
| Shadow price of carbon | USD 40–80 (2020), 50–100 (2030), rising 2.25% a year | World Bank (2017, 2024) | G |


# Appendix D. Sources and verification notes

All figures in Tables 5 and B1 come from primary documents (reports, PDFs or official summaries). The following points were checked and should be noted when citing.

* **Beyond the Gap** (Rozenberg and Fay 2019). The 6% rate is confirmed in the technical appendix: "discounted with a 6 percent discount rate, and … annualized between 2015 and 2030". The text refers to "thousands of scenarios"; a specific total of about 3,000 could not be confirmed. The 4.5% central case breaks down as: electricity 2.2%, transport 1.3%, water and sanitation 0.55%, flood protection 0.32%, irrigation 0.13%.
* **CCDR synthesis** (World Bank 2023, *The Development, Climate, and Nature Crisis*). It is dated December 2023 and covers 38 CCDRs and 42 economies. The range is "less than 1 to 10 percent of GDP". It names Envisage, CGEBox, MANAGE, MFMod, the LTGM, CPAT and E3ME. RAND and CDPM are not mentioned. No discount rate is stated.
* **Bhattacharya et al. (2022).** The +4 points of GDP is the 2025 increment (+3.8 points). The 2030 increment is about +7 points. Stern is credited for guidance, not as an author, on the LSE version.
* **IHLEG 3 (2024).** The cost-of-capital figures (solar WACC 9–12% vs 5–6%; cost of debt 21% vs 6%) and the co-benefits figure (15–18% of global GDP in 2030) are verbatim.
* **IEA (2021).** The "200 basis points saves USD 15 trillion" figure is widely cited but was not found in the IEA's published text. Verify against the full PDF before citing. "Up to seven times higher" financing costs is verbatim.
* **IMF GFSR October 2022, chapter 2.** It says USD 1 trillion a year for EMDE renewables by 2030, not USD 2 trillion, and uses Beyond the Gap's 6%.
* **IPCC AR6 WG3, chapter 15.** The LDC multiple is 4–7 times, scoped to sectors such as AFOLU.
* **IRENA WETO** headline figures (about USD 150 trillion to 2050) could not be verified from IRENA's site and are omitted from Tables 5 and B1.
* **CPI GLCF 2025.** The needs figure could not be extracted; the 2023 and 2024 editions are used.
* **CPI (Gautam et al. 2023).** The "seven times" figure often attached to it was not found in the paper. It comes from the IEA (2021).
* **Global method parameters.** The following were not checked against the primary source and are marked † in Tables 3 and C:
  * IRENA's 2023 unit costs for solar, wind and storage;
  * the IEA's ratio of grid to renewables investment;
  * the 6% WACC for China, which is an assumption;
  * the abatement per dollar (a), which is the author's derivation from the IEA's Net Zero pathway.

  Several sources could not be reached in this review: McCollum et al. (2018), the NGFS investment-variable list, the public CPAT pages and the GI Hub Outlook tool. CDPM's country repositories outside LAC were seen on GitHub (sisepuede-framework), but no consolidated global input database was found. Confirm with the CDPM team.


# Appendix E. Annotated bibliography on financing

The bibliography is grouped by theme. Entries marked † were verified only through citation records, or have a detail that could not be confirmed. Notes say which.

## E.1 Cost of capital and de-risking

**IEA (2021). *Financing Clean Energy Transitions in Emerging and Developing Economies.* World Energy Investment Special Report, with the World Bank and the World Economic Forum. Paris: IEA.** <https://www.iea.org/reports/financing-clean-energy-transitions-in-emerging-and-developing-economies>
The reference point for the EMDE clean energy financing gap. Clean energy investment in EMDEs fell below USD 150 billion in 2020 and must rise more than sevenfold to over USD 1 trillion a year by the end of the 2020s. Nominal financing costs are "up to seven times higher" than in the United States and Europe. It frames the cost of capital (base rate plus country and technology premia) as central to the gap. †The widely quoted estimate that a 200 basis point fall in the cost of capital would cut cumulative net-zero financing costs by USD 15 trillion to 2050 is attributed to this report but was not found in its published text; check the full PDF before citing.

**IEA and IFC (2023). *Scaling Up Private Finance for Clean Energy in Emerging and Developing Economies.* Paris: IEA.** <https://www.iea.org/reports/scaling-up-private-finance-for-clean-energy-in-emerging-and-developing-economies>
Updates the 2021 report. EMDE clean energy investment must rise from USD 770 billion (2022) to USD 2.2–2.8 trillion a year by the early 2030s, and from USD 260 billion to 1.4–1.9 trillion outside China. About 60% must be private. It is the source of the finding that the cost of capital for utility-scale solar is "two or three times higher" in key emerging economies. Concessional finance of USD 80–100 billion a year is needed. Useful for the public–private split.

**Gautam, K., Purkayastha, D. and Widge, V. (2023). *Cost of Capital for Renewable Energy Investments in Developing Economies — and the Need for a Global Credit Guarantee Facility.* Discussion Paper, Climate Policy Initiative, June.** [PDF](https://www.climatepolicyinitiative.org/wp-content/uploads/2023/06/Discussion-Paper-EM-Cost-of-Capital-for-RE-and-GCGF-FINAL-Jun-2023.pdf)
CPI's own cost-of-capital study, written with the International Solar Alliance. It builds a Climate Investment Risk Score from the sovereign rating, political risk and off-taker risk, and converts it through CAPM into a country-level Climate Investment Risk Premium (1% in Germany to 29% in Zambia). From this it estimates:

* cost of debt for solar: from 2.8% in Germany to 45% in Zambia;
* required equity return: from 8.3% to 51%.

It proposes a Global Credit Guarantee Facility of USD 6.6 billion, 10% paid in and the rest callable capital from donors pro-rated by emissions, targeting high-solar-potential countries. It projects an average six-notch rating improvement and about a 6-point cut in risk premia. The best single source for lens (a). †The "seven times" figure sometimes attributed to this paper is not in it (it is the IEA's). CPI's October 2023 GCGF proposal paper could not be located separately and may be a revision of this paper.

**Persaud, A. (2023). *An FX Guarantee Mechanism for the Green Transformation in Developing Countries.* Climate Policy Initiative, June.** †
Proposes a partial foreign-exchange guarantee, funded by a modest fee and donor backing, to absorb the currency risk that makes local-currency clean energy finance expensive. It compares the cost of capital for a solar farm in developed and developing countries. It complements the GCGF (credit risk) by unbundling FX risk, as Wolf (2023) argues. Title wording varies between citing sources; the CPI page could not be loaded.

**Ameli, N., Dessens, O., Winning, M., Cronin, J., Chenet, H., Drummond, P., Calzadilla, A., Anandarajah, G. and Grubb, M. (2021). "Higher cost of finance exacerbates a climate investment trap in developing economies." *Nature Communications* 12: 4046.** <https://doi.org/10.1038/s41467-021-24305-3>
Shows that using region-specific WACC in a global energy model (TIAM-UCL), rather than the uniform rate used in most integrated assessment models, cuts green electricity in Africa by 35% in a cost-optimal 2°C pathway. Policies that lower WACC would let Africa reach net zero about ten years earlier. Coins the "climate investment trap": high risk raises the cost of capital, which deters investment and worsens climate risk. This is the key methodological critique of models, CDPM included, that use one discount rate everywhere.

**Steffen, B. (2020). "Estimating the cost of capital for renewable energy projects." *Energy Economics* 88: 104783.** <https://doi.org/10.1016/j.eneco.2020.104783> †
Compares survey, CAPM and auction-based methods for estimating technology- and country-specific costs of capital for renewables, and provides one of the most used empirical datasets. It is the methodological basis for replacing a uniform discount rate with differentiated WACC.

**Calcaterra, M., Aleluia Reis, L., Fragkos, P., Briera, T. et al. (2024). "Reducing the cost of capital to finance the energy transition in developing countries." *Nature Energy* 9: 1241–1251.** †
Builds country- and technology-specific costs of capital into integrated assessment models and quantifies how de-risking (blended finance, guarantees, concessional loans) lowers total transition costs in developing countries. A modelled, peer-reviewed counterpart to the IEA's cost-of-capital claims.

**Kling, G., Lo, Y.C., Murinde, V. and Volz, U. (2018). "Climate vulnerability and the cost of debt." Centre for Global Finance Working Paper 12/2018, SOAS. SSRN 3198093.** †
Early empirical evidence that climate vulnerability raises sovereign borrowing costs in developing countries, independent of macro fundamentals. Supports treating climate vulnerability as a distinct premium in the cost of capital. Cited by Kharas and Rivard (2022).

**IEA, World Economic Forum, ETH Zurich and Imperial College London. *Cost of Capital Observatory.*** <https://www.iea.org/reports/cost-of-capital-observatory>
Free dashboard of actual financing costs for clean energy in EMDEs, updated with 2023–2024 data in 2025. The IEA's methodology article gives real pre-tax WACC of 3–6% for utility-scale solar and onshore wind in advanced economies, and finds the cost of capital accounts for 20–50% of the levelised cost of solar. The recommended data source for country WACC in CDPM.

## E.2 Debt, creditworthiness and fiscal space

**Kharas, H. and Rivard, C. (2022). "Debt, creditworthiness, and climate: A new development dilemma." Brookings Global Working Paper 180, December.** [PDF](https://www.brookings.edu/wp-content/uploads/2022/12/Debt-creditworthiness-and-climate.pdf)
Written for the World Bank/IMF/Stern High-Level Advisory Group and presented to the Songwe–Stern IHLEG. It is the source of the FT chart on debt paths.

* *Set-up.* From an IMF October 2022 WEO baseline, a big push raises investment and deficits by 4 points of GDP and growth by 1 point. Half is financed domestically and half externally, either at IBRD terms (an effective 5.2% for a 10-year flexible loan, against a US 10-year yield of 3.82%) or at each country's market bond yield.
* *Results.* At IBRD rates, income is 32% above baseline by 2050 and debt about 20 points of GDP higher. At market rates, average developing-country debt rises from about 52% to 80% to 135% of GDP. Market rates exceed official non-concessional rates by more than 10 points in Egypt and Uganda. Only six EMDEs could borrow below 5%, and 23 paid more than 10%.
* *Ratings.* Using rating models from the literature, the big push at IBRD rates improves ratings, because income and the rule of law drive ratings (68% of variance) far more than debt to GDP (2%). The exception is the low-growth, high-deficit quarter of countries.
* *Policy asks.* IFIs should change how they assess creditworthiness, and MDBs should scale up affordable non-concessional lending. Countries should raise domestic revenue and strengthen institutions.

The key caveat is that growth is assumed, not derived from the investment. This is the main bridge between investment needs and the financing lens in the main text. Its full reference list is in E.6.

**Rogoff, K. (2022). "Emerging market sovereign debt in the aftermath of the pandemic." *Journal of Economic Perspectives* 36(4): 147–166.** <https://doi.org/10.1257/jep.36.4.147>
Argues that fiscal space in emerging markets is a binding, state-contingent constraint that worsens with higher world interest rates, falling commodity prices or global recession, and that many low-income countries are already in distress. Calls for more grants rather than loans from the Bretton Woods institutions. A caution against assuming unconstrained borrowing for climate investment.

**Cevik, S. and Jalles, J.T. (2020). "Feeling the heat: Climate shocks and credit ratings." IMF Working Paper 20/286.** <https://ideas.repec.org/p/imf/imfwpa/2020-286.html>
Finds that climate vulnerability lowers sovereign ratings and resilience raises them, after controlling for macro-fiscal determinants, with larger effects in developing countries. One of the two rating models behind Kharas and Rivard's creditworthiness result.

**Borio, C. and Packer, F. (2004). "Assessing new perspectives on country risk." *BIS Quarterly Review*, December: 47–63.** <https://www.bis.org/publications/assessing-new-perspectives-country-risk>
Examines debt intolerance, "original sin" (inability to borrow abroad in domestic currency) and currency mismatches as drivers of country risk. Finds that income and institutional quality dominate ratings. The other rating model behind Kharas and Rivard, and the basis for the claim that debt-financed investment need not worsen ratings.

**IMF (2023). *Fiscal Monitor, October 2023: Climate Crossroads — Fiscal Policies in a Warming World.*** <https://www.imf.org/en/Publications/FM/Issues/2023/10/10/fiscal-monitor-october-2023>
Frames climate policy as a trilemma between climate goals, debt sustainability and political feasibility.

* A spending-led policy mix could raise public debt by 45–50% of GDP by 2050.
* A carbon-pricing-led mix raises it by 10–15%; for EMDEs, about 15% of GDP.
* Each year of delay adds 0.8–2.0% of GDP a year.

The IMF's reference on the fiscal cost of financing investment needs.

## E.3 The scale of needs and the big push

**Bhattacharya, A., Dooley, M., Kharas, H. and Taylor, C. (2022). *Financing a Big Investment Push in Emerging Markets and Developing Economies for Sustainable, Resilient and Inclusive Recovery and Growth.* LSE Grantham Research Institute and Brookings, May.** [PDF](https://www.lse.ac.uk/granthaminstitute/wp-content/uploads/2022/05/Financing-the-big-investment-push-in-emerging-markets-and-developing-economies-for-sustainable-resilient-and-inclusive-recovery-and-growth-1.pdf)
Sets the size of the big push. Spending in EMDEs other than China on human capital, sustainable infrastructure, adaptation and natural capital must rise from 11.3% of GDP (2019) to 15.1% (2025) and 18.2% (2030). That is an extra USD 1.3 trillion by 2025 and USD 3.5 trillion by 2030. It proposes a "Grand Match" for the 2025 increment:

* domestic revenue USD 653 billion (2.7% of GDP);
* concessional finance +USD 96 billion;
* MDB non-concessional finance +USD 126 billion;
* bilateral finance +USD 35 billion;
* private finance +USD 395 billion.

The basis for Kharas and Rivard's 4-point push and for the IHLEG. Stern is credited with guidance, not as an author, in the published version; Kharas and Rivard list him as a co-author.

**Songwe, V., Stern, N. and Bhattacharya, A. (2022). *Finance for Climate Action: Scaling Up Investment for Climate and Development.* Report of the Independent High-Level Expert Group on Climate Finance (IHLEG). London: LSE Grantham Research Institute, November.** [PDF](https://www.lse.ac.uk/granthaminstitute/wp-content/uploads/2022/11/IHLEG-Finance-for-Climate-Action-1.pdf)
EMDCs other than China need about USD 2.4 trillion a year by 2030 (6.5% of GDP) for the energy transition, adaptation and nature, of which USD 1 trillion is external finance; about half can come from domestic sources. It stresses that "the USD 1 trillion is not the new USD 100 billion": it is all external finance, not only donor finance.

**Bhattacharya, A., Songwe, V., Soubeyran, E. and Stern, N. (2023). *A Climate Finance Framework: Decisive Action to Deliver on the Paris Agreement.* Second IHLEG report.** [Summary](https://www.lse.ac.uk/granthaminstitute/wp-content/uploads/2023/11/A-Climate-Finance-Framework-IHLEG-Report-2-SUMMARY.pdf)
Reaffirms USD 2.4 trillion. To 2030:

* international private finance must rise more than 15-fold;
* MDB finance must triple;
* concessional finance must rise fivefold.

Delivery, not the estimate of needs, is the binding constraint.

**Bhattacharya, A., Songwe, V., Soubeyran, E. and Stern, N. (2024). *Raising Ambition and Accelerating Delivery of Climate Finance.* Third IHLEG report, November.** [PDF](https://www.lse.ac.uk/granthaminstitute/wp-content/uploads/2024/11/Raising-ambition-and-accelerating-delivery-of-climate-finance_Third-IHLEG-report.pdf)
The latest IHLEG update. Climate investment needs are USD 6.3–6.7 trillion a year globally by 2030. For EMDCs other than China they are USD 2.3–2.5 trillion (2030) and 3.1–3.5 trillion (2035): clean energy 1.6 trillion, adaptation 0.25, loss and damage 0.25, nature 0.3, just transition 0.04. External finance is about USD 1 trillion by 2030 and 1.3 trillion by 2035.

It is the first IHLEG report to quantify the cost of capital:

* solar WACC of 9–12% in EMDCs against 5–6% in advanced economies;
* cost of debt of 21% against 6%;
* a 1-point fall in the cost of capital saves about USD 150 billion a year.

It is also the first to quantify co-benefits and avoided costs, at 15–18% of global GDP in 2030. The closest global counterpart to lens (c).

**G20 Independent Expert Group (Summers, L. and Singh, N.K., co-chairs) (2023). *Strengthening Multilateral Development Banks: The Triple Agenda.*** †
Recommends tripling MDB lending (to about USD 390 billion a year by 2030), a triple mandate covering poverty, prosperity and global public goods, and greater use of capital adequacy flexibility. The institutional route to scaling the "affordable non-concessional lending" that Kharas and Rivard call for. The primary URL and the exact USD 390 billion wording could not be confirmed.

**Bridgetown Initiative (2022; versions 2.0 and 3.0, 2023–2024). Government of Barbados.** †
Proposals for reforming the international financial architecture:

* re-channel SDRs and expand MDB capital and guarantees;
* add natural-disaster clauses to debt contracts;
* create new levies for loss and damage;
* scale up concessional and guarantee-backed climate finance.

It is the political origin of much of the MDB reform agenda. The exact content and dates of version 3.0 could not be confirmed from an official source.

**Rozenberg, J. and Fay, M., eds. (2019). *Beyond the Gap: How Countries Can Afford the Infrastructure They Need while Protecting the Planet.* Washington, DC: World Bank.** <https://openknowledge.worldbank.org/handle/10986/31291>
Low- and middle-income countries need 2–8% of GDP a year for infrastructure to 2030, with 4.5% (USD 1.5 trillion) as the central case, at a 6% discount rate. Decarbonised paths "need not cost more". Spending better (maintenance, lower construction costs) matters as much as spending more. The World Bank benchmark for total infrastructure needs.

**World Bank (2023). *The Development, Climate, and Nature Crisis: Solutions to End Poverty on a Livable Planet.* CCDR synthesis.** <https://openknowledge.worldbank.org/handle/10986/40652>
Across 42 economies, building resilience and cutting emissions 73% by 2050 needs an extra 1.4% of GDP a year to 2030, ranging from under 1% to 10%. Extrapolated to all LICs and MICs other than China, that is USD 574 billion a year (2.8% of GDP). It names the macro models used in CCDRs. The closest World Bank counterpart to CDPM's incremental estimates.

## E.4 Commentary (Financial Times)

All by Martin Wolf. Headlines, dates and standfirsts were confirmed from archived copies.

**Wolf, M. (2022). "Delay only makes climate action more urgent." *Financial Times*, 8 November.** <https://www.ft.com/content/5198e3ba-997f-4db5-8310-3a68ae91d9f2>
"A huge acceleration in investment is necessary to help emerging and developing countries." Written at COP27, and drawing on the Songwe–Stern work, it argues that delay compounds both costs and risks. The FT chart of debt paths under IBRD and market rates, credited to the IHLEG, comes from Kharas and Rivard (2022).

**Wolf, M. (2023). "The green transition won't happen without financing for developing countries." *Financial Times*, 20 June.** <https://www.ft.com/content/770aadbb-1583-40ae-b072-9ef44c27cc15>
"Multilateral development banks and the IMF should offer to pool currency risks." Argues that the cost of capital, not technology, is the bottleneck for the transition in developing countries, and that MDBs and the IMF should absorb FX risk. Closely aligned with the CPI work (Gautam et al. 2023; Persaud 2023) published the same month, and probably the Wolf column that drew on CPI's cost-of-capital estimates.

**Wolf, M. (2024). "Time is running out for the world to meet climate targets." *Financial Times*, 11 November.** <https://www.ft.com/content/aed9fbdb-0e07-423b-8028-15b40d6b1a5a>
"Human inability to act in advance of distant perils is preventing necessary action now." Written at COP29, it draws on:

* the IEA's World Energy Outlook 2024 (EMDE clean energy investment must rise about sevenfold);
* Kotz et al. (2024, *Nature*), estimating a 19% loss in global income by 2050 from damages already committed;
* Carney's "tragedy of the horizon".

## E.5 Items supplied but outside the topic

**Wolf, M. (2023). "A sensible strategy for the UK needs radical changes on pensions." *Financial Times*, 19 March.** <https://www.ft.com/content/9bcdd1b4-6c07-4edd-a3c4-1ae4ec2ab971>
UK pension and fiscal reform. Not related to climate finance.

**Wolf, M. (2023). "Citizens' juries can help fix democracy." *Financial Times*, 28 May.** <https://www.ft.com/content/81b94348-c915-4b3c-a9b2-8a26d71bb4c9>
Deliberative assemblies in Western democracies. Not related to climate finance.

## E.6 References cited in Kharas and Rivard (2022)

Grouped by type; endnote numbers in brackets. The key sources for this note are 1 (the big push, basis of the IHLEG), and 6 and 8 (the two rating models behind the creditworthiness result).

*Academic papers*

1. Bhattacharya, A., Dooley, M., Kharas, H., Taylor, C. and Stern, N. (2022). "Financing a big investment push in emerging markets and developing economies for sustainable, resilient and inclusive recovery and growth." LSE Grantham Research Institute. [3]
2. Rogoff, K. (2022). "Emerging market sovereign debt in the aftermath of the pandemic." *Journal of Economic Perspectives* 36(4): 147–166. [7]
3. Cantor, R. and Packer, F. (1996). "Determinants and impact of sovereign credit ratings." *Economic Policy Review* 2(2). [8]
4. Kling, G., Lo, Y.C., Murinde, V. and Volz, U. (2018). "Climate vulnerability and the cost of debt." SSRN 3198093. [9]
5. Kahn, M., Mohaddes, K., Ng, R., Pesaran, M.H., Raissi, M. and Yang, J.-C. (2019). "Long-term macroeconomic effects of climate change: A cross-country analysis." NBER Working Paper 26167. [10]
6. Cevik, S. and Jalles, J.T. (2020). "Feeling the heat: Climate shocks and credit ratings." SSRN 3772492. [17]
7. Afonso, A. (2003). "Understanding the determinants of sovereign debt ratings: Evidence for the two leading agencies." *Journal of Economics and Finance* 27(1): 56–74. [18]
8. Borio, C. and Packer, F. (2004). "Assessing new perspectives on country risk." *BIS Quarterly Review*, December. [19]
9. Bissoondoyal-Bheenick, E. (2005). "An analysis of the determinants of sovereign ratings." *Global Finance Journal* 15(3): 251–280. [20]
10. Mellios, C. and Paget-Blanc, E. (2006). "Which factors determine sovereign credit ratings?" *European Journal of Finance* 12(4): 361–377. [21]
11. Afonso, A., Gomes, P. and Rother, P. (2007). "What 'hides' behind sovereign debt ratings?" SSRN 954705. [22]
12. Canuto, O., Dos Santos, P.F.P. and de Sá Porto, P.C. (2012). "Macroeconomics and sovereign risk ratings." *Journal of International Commerce, Economics and Policy* 3(2). [23]
13. Bozic, V. and Magazzino, C. (2013). "Credit rating agencies: The importance of fundamentals in the assessment of sovereign ratings." *Economic Analysis and Policy* 43(2): 157–176. [24]
14. Teixeira, J.C.A., Silva, F.J.F., Ferreira, M.B.S. and Vieira, J.A.C. (2018). "Sovereign credit rating determinants under financial crises." *Global Finance Journal* 36: 1–13. [25]
15. Wüste, C. (2022). "The logics of sovereign credit ratings in developed and developing countries." *Research in International Business and Finance* 62: 101717. [26]

*Policy and institutional sources*

16. IMF (2022). *World Economic Outlook, October 2022: Countering the Cost-of-Living Crisis.* [4]
17. Georgieva, K. (2022). "The IMF's role in a more shock-prone world." Plenary speech, IMF–World Bank Annual Meetings, 14 October. [1]
18. IMF (2021). "The G20 Common Framework for debt treatments must be stepped up." IMF Blog, 2 December. [5]
19. Devarajan, S. and Kharas, H. (2022). "Is the Sri Lankan debt crisis a harbinger?" *Foreign Affairs*, 4 August. [6]
20. McCormack, T., Parker, E., Masip, M. and Shearman, J. (2022). "Sovereign rating criteria." Fitch Ratings, 11 July. [11]
21. Kimenyi, M.S. (2012). "Is Rwanda exportable?" Brookings blog. [16]
22. Stern, N. (2022). "New approaches to the economics of climate change: Urgency, scale, opportunity." Princeton lecture (YouTube). [2]

*Data and background sources*

23. Kaufmann, D. and Kraay, A. *Worldwide Governance Indicators.* World Bank (accessed 2022). [13]
24. Trading Economics. Rwanda credit rating; credit ratings by country (accessed 2022). [14, 27]
25. World Bank. *Migration and Remittances Data* (May 2021). [28]
26. Lemarchand, R. "Rwanda: Genocide and aftermath." *Britannica* (accessed 2022). [12]
27. Country Economy. Rwanda vs US homicide rate comparison (2015). [15]

Two further data sources appear only in figure notes: IMF Government Finance Statistics, and bond yields from World Government Bonds and Market Insider.


# References


Ameli, N. et al. (2021). "Higher cost of finance exacerbates a climate investment trap in developing economies." *Nature Communications* 12: 4046.

Barnes, T., Shivakumar, A., Brinkerink, M. and Niet, T. (2022). "OSeMOSYS Global, an open-source, open data global electricity system model generator." *Scientific Data* 9: 623.

Bhattacharya, A., Dooley, M., Kharas, H. and Taylor, C. (2022). *Financing a Big Investment Push in Emerging Markets and Developing Economies.* LSE Grantham Research Institute and Brookings.

Bhattacharya, A., Songwe, V., Soubeyran, E. and Stern, N. (2024). *Raising Ambition and Accelerating Delivery of Climate Finance.* Third IHLEG report. LSE Grantham Research Institute.

Climate Policy Initiative (2023, 2024). *Global Landscape of Climate Finance.* CPI.

Emmerling, J., Drouet, L., van der Wijst, K.-I., van Vuuren, D., Bosetti, V. and Tavoni, M. (2019). "The role of the discount rate for emission pathways and negative emissions." *Environmental Research Letters* 14: 104008.

Energy Transitions Commission (2023). *Financing the Transition.* ETC.

Fay, M., Hallegatte, S., Kraay, A., Vogt-Schilb, A. et al. (2016). *Discounting Costs and Benefits in Economic Analysis of World Bank Projects.* OPSPQ Guidance Note. World Bank.

Gautam, K., Purkayastha, D. and Widge, V. (2023). *Cost of Capital for Renewable Energy Investments in Developing Economies.* CPI Discussion Paper.

Groves, D.G., Syme, J., Molina-Perez, E. et al. (2020). *The Benefits and Costs of Decarbonizing Costa Rica's Economy.* RAND Corporation and IDB. RR-A633-1.

IEA (2021). *Financing Clean Energy Transitions in Emerging and Developing Economies.* IEA.

IEA (2023). *Electricity Grids and Secure Energy Transitions.* IEA.

IEA (2023). *Net Zero Roadmap: A Global Pathway to Keep the 1.5°C Goal in Reach, 2023 Update.* IEA.

IEA (2024). *Cost of Capital Observatory.* IEA.

IEA (2024). *Global Energy and Climate Model: Documentation.* IEA. <https://www.iea.org/reports/global-energy-and-climate-model>

IEA and IFC (2023). *Scaling Up Private Finance for Clean Energy in Emerging and Developing Economies.* IEA.

IMF (2022). *Global Financial Stability Report, October 2022*, Chapter 2. IMF.

IPCC (2022). *Climate Change 2022: Mitigation of Climate Change*, Working Group III, Chapter 15. Cambridge University Press.

IRENA (2024). *Renewable Power Generation Costs in 2023.* IRENA.

Kalra, N., Molina-Pérez, E., Syme, J., Esteves, F., Cortés, H., Rodríguez-Cervantes, M.T., Espinoza-Juárez, V.M., Jaramillo, M., Baron, R., Alatorre, C., Buttazzoni, M. and Vogt-Schilb, A. (2023). *The Benefits and Costs of Reaching Net Zero Emissions in Latin America and the Caribbean.* IDB and RAND. doi:10.18235/0005330. Sector appendices: RAND RR-A2847-1.

Karlsson, M., Alfredsson, E. and Westling, N. (2020). "Climate policy co-benefits: a review." *Climate Policy* 20(3): 292–316.

Kharas, H. and Rivard, C. (2022). "Debt, creditworthiness, and climate." Brookings Global Working Paper 180.

Markandya, A. et al. (2018). "Health co-benefits from air pollution and mitigation costs of the Paris Agreement." *Lancet Planetary Health* 2(3): e126–e133.

McCollum, D.L. et al. (2018). "Energy investment needs for fulfilling the Paris Agreement and achieving the Sustainable Development Goals." *Nature Energy* 3: 589–599.

McKinsey Global Institute (2022). *The Net-Zero Transition: What It Would Cost, What It Could Bring.* McKinsey & Company.

NGFS (2024). *NGFS Climate Scenarios*, Phase V. Network for Greening the Financial System.

Parry, I. and Black, S. (2021). "Still not getting energy prices right." IMF Working Paper 21/236.

Parry, I., Heine, D., Lis, E. and Li, S. (2014). *Getting Energy Prices Right: From Principle to Practice.* IMF.

Rozenberg, J. and Fay, M., eds. (2019). *Beyond the Gap.* World Bank.

Songwe, V., Stern, N. and Bhattacharya, A. (2022). *Finance for Climate Action.* IHLEG report. LSE Grantham Research Institute.

UNEP (2023, 2025). *Adaptation Gap Report.* UNEP.

UNFCCC Standing Committee on Finance (2021, 2024). *First and Second Needs Determination Reports.* UNFCCC.

Vandyck, T. et al. (2018). "Air quality co-benefits for human health and agriculture counterbalance costs to meet Paris Agreement pledges." *Nature Communications* 9: 4939.

World Bank (2023). *The Development, Climate, and Nature Crisis: Insights from Country Climate and Development Reports.* World Bank.

World Bank (2024). *International Comparison Program: Purchasing Power Parities and the Size of World Economies.* World Bank.

World Bank (2024). *Shadow Price of Carbon in Economic Analysis.* Guidance Note. World Bank.

CDPM documentation: <https://sisepuede.readthedocs.io/en/latest/>. Code: github.com/jcsyme/sisepuede; github.com/sisepuede-framework/costs_benefits_ssp.
