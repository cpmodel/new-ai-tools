"""Short deck (7 slides): what the Macro-Fiscal-Energy-Scenario Tool does, its dashboard subcomponent, and how it
could link to the FSA. Style: AIClimate-FiscalTools_v2.5 (kit.py). Run: python3 build_scenario_tool.py
Numbers: MacFiscEnergyScenarioProto/Macro_Fiscal_Energy_Scenario_Tool_v0.7.xlsx, Kenya (from build/outlook.py --check)."""
import sys
from pathlib import Path
from pptx.util import Emu
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from kit import *                                                     # noqa: E402

d = Deck(str(HERE.parent / "AIClimate-FiscalTools_v2.5.pptx"), "Macro-Fiscal-Energy-Scenario Tool · 20 Oct 2026")

# 1 -----------------------------------------------------------------------------------------
d.title("Follow-up to our meeting · 20 October 2026", "The Macro-Fiscal-Energy-Scenario Tool",
        "One country, one scenario: revenue, public investment, green investment needs and carbon pricing, "
        "through to growth, debt and debt risk to 2050",
        "Stephen Stretton · World Bank · prototype v0.7 (MacFiscEnergyScenarioProto)")

# 2 -----------------------------------------------------------------------------------------
s = d.content("What it does", "A scenario tool: set fiscal and energy policy, see the fiscal path")
d.cards(s, [
    dict(tag="You set", color=ORANGE, title="Four main levers", bullets=[
        ("Revenue:", "increase in pp of GDP a year, up to a ceiling"),
        ("Public investment:", "increase in pp of GDP"),
        ("Climate investment programme:", "on or off (the budget need for net zero)"),
        ("Carbon pricing:", "tax, feebate, or the most effective politically feasible mix")]),
    dict(tag="It computes", color=NAVY, title="Live, for 186 countries", bullets=[
        "Green investment need and who pays (NMIM)",
        "Feasible carbon price and its revenue (PriceFeasibility)",
        "Growth and debt (DIGNAD-type macro model)",
        "Debt risk (LIC-DSF-lite, low and lower-middle income)"]),
    dict(tag="You read", color=TEAL, title="One Dashboard, 2025-2050", bullets=[
        "Public debt; revenue and expenditure; balance",
        "GDP against no new measures",
        "Green investment need, budget need, green public investment",
        "Carbon price in force; debt sustainability"]),
], h=3300000)
d.so(s, "it answers 'what does this package do to the budget and debt?', in Excel, with no code.")

# 3 -----------------------------------------------------------------------------------------
s = d.content("How it works", "Five models joined in one workbook, driven from the Dashboard", bg=MIST)
d.flow(s, [
    ("1  Data", "Tax & Macro-Fiscal Dashboard panel: 217 countries, fiscal, PIMA, debt, energy", "DCE8F5"),
    ("2  Green investment", "NMIM: net zero investment by sector, who pays (P1-P5), budget need 2025-50", "DCE8F5"),
    ("3  Carbon pricing", "PriceFeasibility: tax vs feebate under limits on tariffs, the poor, coal generators", "DCE8F5"),
    ("4  Macro", "DIGNAD-R live; GIMF / DIGNAD / DIGNAR-lite integrated run as cross-check", "DCE8F5"),
    ("5  Debt", "LIC-DSF-lite: five indicators against thresholds; risk rating", "DCE8F5"),
    ("Dashboard", "Debt, revenue, spending, GDP, green investment, carbon price, DSA", None),
], y=1750000, h=2250000, size=10)
d.box(s, M, 4250000, CW, 620000, WHITE)
d.box(s, M, 4250000, 50000, 620000, ORANGE)
d.text(s, M + 200000, 4300000, CW - 300000, 540000, [
    [("Live, not pasted: ", {"bold": True, "color": ORANGE}),
     ("change an investment assumption (e.g. the public share of viability-gap investment) and the budget need, "
      "the fiscal outlook and debt all recalculate. Every formula is checked against a Python copy of the equations.", {})]], size=11)
d.so(s, "the levers are few; the detail sits behind them for anyone who wants it.", y=5050000)

# 4 -----------------------------------------------------------------------------------------
s = d.content("The workbook", "One core Dashboard; the indicators dashboard is a subcomponent")
s.shapes.add_picture(str(HERE / "products" / "scenario_tool_dashboard_v0_7.png"), Emu(M), Emu(1720000), width=Emu(6300000))
x2 = M + 6550000
d.table_at = None
d.text(s, x2, 1720000, CW - 6550000, 300000, "Tabs, left to right", size=12, bold=True, color=NAVY)
d.bullets(s, x2, 2050000, CW - 6550000, 3000000, [
    ("Dashboard:", "the only place you drive it"),
    ("Detail ->>", "Fiscal outlook · Macro & debt · Green investment · Carbon pricing · Policy settings · model sheets"),
    ("Not master ->>", "the Tax & Macro-Fiscal Dashboard's country profile (16 blocks, 217 countries), kept as a "
                       "subcomponent; its master is maintained separately"),
    ("Data:", "the dashboard panel feeds every model's calibration"),
], size=11, spacing=8)
d.so(s, "the dashboard is the data layer and a country profile; the scenario tool is what runs on top of it.")

# 5 -----------------------------------------------------------------------------------------
s = d.content("Example · Kenya", "What a package does to debt: an illustration")
d.kpis(s, [("75%", "public debt in 2050 (% of GDP) with the climate programme and feasible carbon pricing; 67% with no new measures", NAVY),
           ("128%", "if public investment also rises by 1.5pp of GDP, debt-financed", ORANGE),
           ("31%", "if revenue rises 0.5pp of GDP a year (to 3pp) to pay for it; GDP 3.2% higher by 2050", TEAL),
           ("89 USD/t", "feasible power-sector carbon tax: abates 24% of power CO2, within all three limits", BLUE)], h=1500000)
d.table(s, ["Kenya, 2030", "Value", "Source in the tool"], [
    ["Incremental green investment need", "1.6% of GDP", "Green investment (NMIM)"],
    ["of which direct budget need", "0.5% of GDP", "Green investment: P1 + public gap of P4"],
    ["Debt risk (LIC-DSF-lite)", "High, before and with the programme", "Macro & debt"],
], widths=[0.4, 0.3, 0.3], y=3550000, size=11, h_row=330000)
d.text(s, M, 4950000, CW, 300000, "Illustrative calibration: the baseline holds today's position (no WEO projections yet) and several "
       "parameters are priors. For comparing packages, not as a forecast.", size=10, color=MUTED)

# 6 -----------------------------------------------------------------------------------------
s = d.content("Integration", "A natural fit with the FSA", bg=MIST)
d.cards(s, [
    dict(tag="Tool to FSA", color=ORANGE, title="Scenario paths in", bullets=[
        "Revenue and expenditure paths, by item", "Primary balance, growth, interest",
        "Green investment and budget need by year", "Carbon revenue, and how it is used"]),
    dict(tag="FSA to tool", color=NAVY, title="Sustainability out", bullets=[
        "A proper baseline (where ours is a placeholder)", "Debt dynamics and fiscal space",
        "Replaces our simple debt module for all income groups", "One verdict, consistent with country work"]),
    dict(tag="R and energy risks", color=TEAL, title="Two good additions", bullets=[
        ("R tool:", "integrate it; the tool's data and model sheets map to a clear interface"),
        ("Energy fiscal risks:", "fossil revenue loss, utility liabilities (P2 tariff-backed investment), subsidies"),
        "Most inputs already sit in the tool"]),
], h=3300000)
d.so(s, "the tool sets the policy scenario; the FSA judges its sustainability.")

# 7 -----------------------------------------------------------------------------------------
s = d.content("Status and next steps", "Built and checked; next, the link to the FSA")
d.cards(s, [
    dict(tag="Built", color=TEAL, title="Prototype v0.7", bullets=[
        "186 countries; Excel, no code", "Formulas checked against Python to 1e-14",
        "Green investment matches NMIM exactly", "Docs and version notes in the repository"]),
    dict(tag="Limits today", color=ORANGE, title="Illustrative", bullets=[
        "Baseline holds today's ratios; no WEO path yet", "Macro calibration partly income-group defaults",
        "Carbon pricing: power sector only, several priors"]),
    dict(tag="Next", color=NAVY, title="Proposed steps", bullets=[
        "Agree the FSA interface: which variables, which years", "Integrate the R tool",
        "Add the energy fiscal risks module", "WEO baseline; one pilot country end to end"]),
], h=3300000)
d.so(s, "a short working session on the FSA interface would let us run one country through both.")

d.save(str(HERE / "ScenarioTool-FSA-v1.pptx"))
print("slides:", d.n)
