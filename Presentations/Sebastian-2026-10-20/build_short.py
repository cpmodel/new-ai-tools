"""Short Sebastian deck (10 slides): the one-country chain as the PFR proposal. Style: AIClimate-FiscalTools_v2.5.
The full reference appendix is Sebastian-Platforms-v2.pptx. Run: python3 build_short.py"""
import sys
from pathlib import Path
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from kit import *                                                     # noqa: E402

d = Deck(str(HERE.parent / "AIClimate-FiscalTools_v2.5.pptx"), "Climate-fiscal tools · 20 Oct 2026")

# 1 -----------------------------------------------------------------------------------------
d.title("For discussion · 20 October 2026", "One country, all the tools: a PFR process for energy costs and fiscal risks",
        "Meeting with Sebastian: fiscal policy platform, excise, adaptation, and the proposal",
        "Stephen Stretton · World Bank · full reference appendix: Sebastian-Platforms-v2.pptx")

# 2 -----------------------------------------------------------------------------------------
s = d.content("Purpose and agenda", "Today: agree one proposal and three steers")
d.cards(s, [
    dict(tag="The proposal", color=BLUE, title="An end-to-end PFR process", bullets=[
        "Run one country through all our tools", "Energy costs, fiscal risks, efficiency of spending, tax changes", "Results back in the dashboard, ready for a PFR"]),
    dict(tag="Your agenda", color=NAVY, title="Three topics", bullets=[
        ("1  Fiscal policy platform:", "the chain is its core"), ("2  Excise:", "delivered; what next"), ("3  Adaptation:", "DIGNAD-lite already has the blocks")]),
    dict(tag="Decisions", color=ORANGE, title="What I need from you", bullets=[
        "Go-ahead and a pilot country", "Introduction to the macro unit", "What to drop to make room"]),
], h=3300000)
d.so(s, "the pieces were built this month; the proposal is to join them, for one country first.")

# 3 -----------------------------------------------------------------------------------------
s = d.content("The proposal", "One country, run through all the tools", bg=MIST)
d.flow(s, [
    ("1  Dashboard: data in", "Current panel plus missing macro inputs: CPIA, reserves, remittances, debt service, NFA, Findex, public capital, import share", "DCE8F5"),
    ("2  Investment needs", "How much and who pays: the budget need path 2025-2050 (NMIM)", "DCE8F5"),
    ("3  Price feasibility", "Most power-sector pricing the limits allow (F1-F3); which limit binds; tax revenue", "DCE8F5"),
    ("4  Macro model", "GIMF-lite (market access) · DIGNAD-lite (LIC, concessional) · DIGNAR-lite (resource-rich, new)", "DCE8F5"),
    ("5  Debt test", "LIC-DSF-lite: the verdict on the debt path", "DCE8F5"),
    ("6  Dashboard: results", "Budget need, pricing revenue, financing gap, GDP and debt 2030/50, risk rating, binding limit", None),
], y=1750000, h=2250000, size=10)
d.box(s, M, 4250000, CW, 620000, WHITE)
d.box(s, M, 4250000, 50000, 620000, ORANGE)
d.text(s, M + 200000, 4300000, CW - 300000, 540000, [
    [("Prerequisite: ", {"bold": True, "color": ORANGE}),
     ("land-use and non-CO2 spending must stop counting as productive public capital (TODO item 7). Today they are 90% of Liberia's and 53% of Kenya's "
      "incremental investment in 2025-35, so results there are distorted until this is fixed.", {})]], size=11)
d.so(s, "revenue from feasible pricing becomes the tax-financed share of the budget need; the rest is the financing gap.", y=5050000)

# 4 -----------------------------------------------------------------------------------------
s = d.content("The proposal", "What a PFR gets from the chain")
d.table(s, ["PFR question", "Where the chain answers it", "Status today"], [
    ["Energy costs: what must be invested, and who pays?", "Step 2: NMIM budget need path (direct need 14% of the world total; 29% in EMDCs)", "NMIM v0.2 built"],
    ["Tax changes: what can pricing raise, politically?", "Step 3: F1 tariff rise, F2 bottom-40% loss, F3 generator payments", "PriceFeasibility v0.1; priors to replace"],
    ["Efficiency of public spending", "Step 4: PIMA-mapped efficiency; return on public capital", "Calibrated from dashboard; fuller efficiency on the TODO"],
    ["Fiscal risks", "Steps 4-5: debt path, DSF rating; fossil revenue loss (DIGNAR-lite); utility debt as a contingent liability", "GIMF/DIGNAD/DSF-lite v0.1; DIGNAR and utility debt to build"],
    ["One page for the minister", "Step 6: integrated results block in the dashboard", "To build"],
], widths=[0.3, 0.45, 0.25], y=1750000, size=11, h_row=520000)
d.text(s, M, 5000000, CW, 300000, "All macro calibrations and DSF baselines are illustrative today; only the investment programme is country data. Not a substitute for CPAT or a full DSA.",
       size=10, color=MUTED)

# 5 -----------------------------------------------------------------------------------------
s = d.content("Agenda 1 · Fiscal policy platform", "The fiscal policy platform: one engine, one Python code base")
d.cards(s, [
    dict(tag="Engine", color=NAVY, title="CPAT (Excel-AI v6.2 + Python)", bullets=[
        "Prices, revenue, emissions, households", "Feeds the price-feasibility step (to replace priors)"]),
    dict(tag="Built this month", color=ORANGE, title="Macro and debt", bullets=[
        "GIMF-lite, DIGNAD-lite, LIC-DSF-lite (v0.1)", "Shared calibration from the dashboard", "Pitch deck and user guides"]),
    dict(tag="Also in the platform", color=ORANGE, title="Policy tools", bullets=[
        "FTT national (Zoe, 27 Nov); FTT + MFMod 12 Feb", "Crediting, GHG accounting, FFI (5 Mar)", "CBAM tool (Egypt, global MVP)"]),
    dict(tag="Macro unit", color=TEAL, title="Why them", bullets=[
        "Our lite models mirror IMF/WB tools", "MFMod is the Bank's own macro model", "Ask: a first call in November"]),
], h=3200000, size=11)
d.so(s, "the lite models make the chain work now; the macro unit's MFMod is the natural replacement or cross-check later.")

# 6 -----------------------------------------------------------------------------------------
s = d.content("Agenda 2 and 3", "Excise and adaptation")
d.cards(s, [
    dict(tag="Agenda 2 · Excise", color=ORANGE, title="Delivered; what next?", bullets=[
        "Excise Diagnostic: Sierra Leone, Uganda, Liberia; handed to Dirk 16 Oct", "Run with AI, no code; Word report + Excel",
        "Proposal on the table: excises, externalities, pathways; pilots Colombia/Chile, Uganda",
        "Option: fold fuel excise into step 3 of the chain (same F1 logic)"]),
    dict(tag="Agenda 3 · Adaptation", color=TEAL, title="The blocks exist in DIGNAD-lite", bullets=[
        "Adaptation public capital, disasters, grant/concessional/commercial financing", "Dashboard: INFORM climate risk, ND-GAIN adaptation",
        "Not in the FY27 plan: no days budgeted", "Question: part of the PFR chain, or FY28?"]),
], h=3500000, size=11)

# 7 -----------------------------------------------------------------------------------------
s = d.content("The latest prototype", "Excel-AI MVP v6.2: built, checked, usable")
d.kpis(s, [("56 tabs", "about 560,000 formulas; mitigation, power, distribution, GDP, NDC comparison", NAVY),
           ("123", "legacy CPAT bugs fixed, none copied; cell-by-cell fix plans for the CPAT team", ORANGE),
           ("0 errors", "quick check; 12,992 output codes unchanged between releases", TEAL),
           ("15 min", "from download to results: open, test, change a policy, run scenarios", BLUE)], h=1250000)
d.numbered(s, [("Open", "Excel 365, enable macros; run CheckBatchRun once"), ("Change a policy", "MTInputs column K (e.g. carbon price); it recalculates"),
               ("Compare", "RunAllScenarios; ScenarioCompare up to 8 scenarios; every figure traceable by code")], y=3300000, step=500000, size=12)
d.text(s, M, 4900000, CW, 300000, "Next: Excel 365 review on Windows; final Excel 4 Dec; Python 11 Dec; Shiny app 5 Feb.", size=11, color=MUTED)

# 8 -----------------------------------------------------------------------------------------
s = d.content("Timeline", "Plan to March, and where the chain could fit")
WEEKS = ["19 Oct", "", "2 Nov", "", "16", "", "30", "", "14 Dec", "", "", "4 Jan", "", "18", "", "1 Feb", "", "15", "", "1 Mar", "", "15"]
d.gantt(s, [
    ("Excel-AI MVP; final Excel; Python", "core", 0, 7, NAVY, 7),
    ("FTT national (Zoe)", "fiscal", 0, 5, ORANGE, 5),
    ("Pathways with CDPM: integration, pilots", "climate", 11, 12, TEAL, 12),
    ("Shiny app", "core", 13, 15, NAVY, 15),
    ("FTT + MFMod; crediting, GHG, FFI", "fiscal", 16, 19, ORANGE, 19),
    ("Chain: fix item 7; dashboard data block", "proposed", 2, 4, "9DBBDD", None),
    ("Chain: price feasibility link; DIGNAR-lite", "proposed", 4, 7, "9DBBDD", None),
    ("Chain: pilot country run; PFR results block", "proposed", 11, 14, "9DBBDD", 14),
    ("Handover; contract ends 20 Mar", "", 20, 21, "D3DAE4", 21),
], WEEKS, y=1700000, row_h=330000, today=0.2,
    legend=[("Core engine", NAVY), ("Fiscal & policy", ORANGE), ("Climate", TEAL), ("Chain (proposed)", "9DBBDD"), ("Milestone", GOLD)])
d.text(s, M, 5400000, CW, 300000, "The FY27 plan has no buffer: the chain needs days from somewhere (e.g. the Shiny app, or part of crediting). Away 21 Oct to about 2 Nov.",
       size=10, color=MUTED)

# 9 -----------------------------------------------------------------------------------------
s = d.content("Questions for you", "Questions and decisions", bg=MIST)
d.numbered(s, [
    ("The proposal", "Go ahead with the one-country chain as the PFR process for energy costs and fiscal risks?"),
    ("Pilot country", "Which country first? (Uganda is named as the PFR expenditure pilot; Egypt and Morocco are pathways pilots)"),
    ("Macro unit", "Will you introduce us? MFMod as replacement or cross-check for the lite models? MFMod in or out for 12 Feb?"),
    ("Excise", "Who owns it after Dirk's handover; extend, or fold fuel excise into the chain?"),
    ("Adaptation", "Use DIGNAD-lite's adaptation blocks now, or a note for FY28?"),
    ("Room in the plan", "What to drop or defer to make days for the chain; Claude Code ($200)"),
], y=1750000, step=520000, size=12, col=NAVY)

# 10 ----------------------------------------------------------------------------------------
d.divider("For reference", "The full work programme", "Appendix in Sebastian-Platforms-v2.pptx: core engine, fiscal & policy, climate (CDPM, formerly SiSePuede), PFR, other, plan and funding. "
          "Models: new-ai-tools/MacroFiscalModels and PriceFeasibility (pitch deck and user guides in Docs/).")

d.save(str(HERE / "Sebastian-Short-v3.pptx"))
print("slides:", d.n)
