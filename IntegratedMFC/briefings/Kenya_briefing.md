# Kenya: what a net zero programme would cost the budget, and what it would do to debt

*A worked example from the Integrated MF-C tool v0.4 (`../Integrated_MFC_Tool_v0.4.xlsx`, pick
"Kenya" on the Dashboard). 20 October 2026. Every number below is in the tool; the investment and
macro figures are **model output with illustrative calibrations, not forecasts**.*

## The short version

- **The need is modest and mostly not energy.** Kenya's power is already 88% renewable, so the
  net zero programme is about land use and non-CO2 gases, not power plants. Net zero investment is
  2.3% of GDP in 2030 (about USD 3.6 bn), of which 1.6% of GDP is above current trends. Land use
  and non-CO2 account for 0.94 of that 1.64.
- **The budget's share is about 0.6% of GDP**, and about 0.8% once weak public investment
  management is allowed for. Private capital would carry close to 60%.
- **Carbon pricing cannot pay for it.** With a clean grid there is little to tax: feasible
  power-sector pricing raises 0.05% of GDP in 2030 and nothing by 2040, about 5% of the budget
  need. The financing gap is 0.44% of GDP a year.
- **It adds debt without adding much growth**: about 7 points of GDP of public debt by 2050
  (75% instead of 67%) for a GDP gain of a quarter of a percent. This is because most of the
  budget need is land-use and non-CO2 *spending*, which builds no productive public capital.
- **Kenya is already at high risk of debt distress before the programme**, and stays there. The
  programme is small next to that problem: interest already takes 30% of revenue.
- **The answer depends on the model.** The other model family (GIMF-lite) gives +4.6% of GDP
  by 2050 and less debt, because it credits energy savings and private investment. The gap is
  the most important uncertainty, not a rounding error.

## The worked example, step by step

### 1. How much is needed (NMIM v0.2, Investment 1)

| 2030, % of GDP | Value |
|---|---|
| Investment on a net zero path | 2.31 |
| Investment on current trends | 0.66 |
| **Incremental investment** | **1.64** |
| of which non-CO2 and land use | 0.94 |
| of which power generation + grids | 0.26 |
| of which transport, industry, buildings, other | 0.44 |

It cuts emissions by 26% against the baseline in 2030. At the World Bank's 6% rate the
incremental cost (capital charge less fuel savings) is 0.88% of GDP, about USD 43 per tonne.
Local co-benefits barely change that (0.85%), again because the grid is already clean. Market
finance at Kenya's 8.4% WACC adds 0.1% of GDP a year.

### 2. Who pays (Investment 2)

NMIM splits every investment line by who repays it. For Kenya in 2030 (% of GDP):

| Type | Value | Who repays |
|---|---|---|
| P1 pure public | 0.38 | taxes |
| P2 regulated utility | 0.35 | electricity tariffs |
| P3 contracted revenue | 0.28 | a power purchase agreement with the utility |
| P4 viability gap | 0.73 | private, with a public top-up |
| P5 pure private | 0.56 | the investor |

The **direct budget need** is P1 plus the public 30% of P4:

> 0.38 + 0.3 × 0.73 = 0.38 + 0.22 = **0.60% of GDP** (about USD 0.9 bn)

Kenya's PEFA score for public investment management is 2.5 out of 4, which maps to a 26.5%
efficiency gap (40% at a score of 1, 13% at 4):

> gap = 0.40 − (2.5 − 1)/3 × 0.27 = 0.265;  0.60 / (1 − 0.265) = **0.81% of GDP**

So weak delivery adds about a third to the bill. The capital budget is only 63% executed
(BOOST, 2022), which is a separate risk the tool does not yet model (MacroFiscalModels TODO 23).

Wider public exposure: adding the utility's investment (P2, a contingent liability if tariffs
fall short) gives 0.95% of GDP; everything needing public money or a public decision is 1.46%.

### 3. How it is paid for (Macro 1)

The macro models are fed the *incremental* budget need, 0.49% of GDP in 2030 (the baseline would
spend the rest anyway). Of that, 0.39 is spending on land use and non-CO2, and only 0.10 is capital.

PriceFeasibility finds a feasible power-sector carbon tax of USD 88/t; the binding limit is the
loss to fossil-fuel generators. On Kenya's small fossil generation it raises 0.05% of GDP in 2030, falling to
zero as oil-fired plants retire:

> financing gap 2030 = 0.49 − 0.05 = **0.44% of GDP**; pricing covers 5% of the need to 2050

### 4. What it does to growth and debt (Macro 2, Macro sheet)

Kenya runs on **DIGNAD-lite**, the model for lower-middle income countries that rely on
concessional finance (56% of its external public debt is multilateral). The programme is financed
30% by grants, 30% concessional and the rest commercially, less the pricing revenue.

| | 2030 | 2050 |
|---|---|---|
| GDP, programme vs no programme | +0.07% | +0.26% |
| Public debt with the programme (% GDP) | 68.3 | 74.6 |
| Public debt without it (% GDP) | 67.3 | 67.3 |
| **Debt added by the programme (pp)** | **1.0** | **7.2** |

The growth effect is small because the capital part is small. The live simple model on the Macro
sheet (DIGNAD-R, in formulas) gives almost the same answer: +0.25% and 75.0% of GDP. Change its
yellow cells to see what moves it. For example, more grant financing lowers the debt path;
switching off the pricing earmark barely matters.

**Cross-check:** GIMF-lite on the same programme gives **+4.6% of GDP by 2050 and debt of 70.5%**.
GIMF-lite counts energy savings (1.7% of GDP by 2050) and the private investment that comes with
the programme; DIGNAD-lite counts only the public part. Read the two as a range, and treat which
channels to credit as the question to settle with the country team.

### 5. Is the debt sustainable (DSA sheet)

Kenya's debt-carrying capacity comes from the LIC-DSF Composite Indicator:

> CI = 0.385 × CPIA 3.83 + 2.719 × growth 0.045 + 4.052 × cover 0.184 − 3.990 × 0.184² + 2.022 × remittances 0.036 + 13.520 × world growth 0.033 = **2.73 → medium** (cut-offs 2.69 and 3.05)

| Indicator (%) | Threshold | Baseline, highest 2025-34 | With programme, highest 2025-34 |
|---|---|---|---|
| PV of public debt / GDP | 55 | 66.4 (breach) | 66.5 (breach) |
| Ext. debt service / revenue | 18 | 19.3 (breach) | 19.8 (breach) |
| PV ext. debt / GDP | 40 | 30.4 | 31.2 |
| PV ext. debt / exports | 180 | 102.1 | 104.9 |
| Ext. debt service / exports | 15 | 11.4 | 11.6 |

**Mechanical rating: high risk, before and with the programme.** The programme moves the
indicators by less than a point within ten years; the problem is the starting position.

**One assumption matters a lot here.** The tool has no country trade data yet, so reserve cover
uses an assumed import ratio of 40% of GDP. If Kenya's imports are nearer 20% of GDP (to check
against WDI), reserve cover doubles and the CI rises to about 3.07, just over the cut-off for
*strong* capacity, which has higher thresholds. Getting trade data (MacroFiscalModels TODO 11) is the
single most useful data fix for this country.

### 6. The fiscal and institutional backdrop (Dashboard, Appendix E)

- Gross debt 67% of GDP; net interest 5.2% of GDP and 30% of revenue; tax revenue 12.9% of GDP.
- 44% of public debt is external, 74% of it in US dollars; credit rating index 7.3 out of 21.
- Debt maturing within a year: 6.1% of GDP. Short-term external debt: 26% of reserves.
- Climate is built into public investment management, budget tagging and a national climate fund
  (2021 snapshot). The institutions exist; execution (63%) and efficiency (PEFA 2.5) are the weak points.

### 7. A second opinion on the need (Appendix F)

The CDPM-based global method (v1.1 workbook) puts Kenya's 2030 need higher in total (4.7% of GDP)
but lower above trend (1.1%). It applies one flat 40% public share to all developing countries,
which gives a public need of 1.9% of GDP against NMIM's 0.6%. The difference is the value of
NMIM's who-pays split. The two methods are not expected to agree.

## What this means for a policy conversation

1. Kenya's mitigation bill is not mainly an energy bill. The budget question is land use and
   non-CO2, which earn no tariff revenue and build little capital: grant and concessional
   finance, not carbon pricing, is the realistic route.
2. Delivery matters as much as the amount: closing the efficiency gap is worth about 0.2% of GDP
   a year on the same programme.
3. Any new borrowing for it lands on a debt position already at high risk. Financing terms
   (grant share), not the programme's size, decide the debt outcome.
4. Before using these numbers externally: settle the model choice (DIGNAD-lite vs GIMF-lite),
   add trade data, and replace the income-group defaults for the return on public capital, grants
   and the hand-to-mouth share with country values (MacroFiscalModels TODO 38: pilot calibrations
   for Kenya, Sierra Leone, Uganda, Liberia).
