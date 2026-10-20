"""Sebastian meeting deck v2 (20 Oct 2026): core slides on his issues, then a full appendix on the work programme.
Style: AIClimate-FiscalTools_v2.5. Run: python3 build_v2.py  ->  Sebastian-Platforms-v2.pptx
Name rule: CDPM is the model formerly called SiSePuede; say so on every slide that names it."""
import sys
from pathlib import Path
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from kit import *                                                     # noqa: E402
import appendix_bdef                                                  # noqa: E402

CDPM_NOTE = "CDPM = the model formerly called SiSePuede"
d = Deck(str(HERE.parent / "AIClimate-FiscalTools_v2.5.pptx"), "Climate-fiscal tools · 20 Oct 2026")
WEEKS = ["19 Oct", "", "2 Nov", "", "16", "", "30", "", "14 Dec", "", "", "4 Jan", "", "18", "", "1 Feb", "", "15", "", "1 Mar", "", "15"]


# =============================================================================================
# CORE
# =============================================================================================
d.title("For discussion · 20 October 2026",
        "Climate-fiscal tools: one engine, four platforms",
        "Meeting with Sebastian: the fiscal policy platform, excise, adaptation, and a PFR proposal for an end-to-end process on energy costs and fiscal risks",
        "Stephen Stretton · World Bank · core slides 2-15, appendix from slide 16")

s = d.content("Purpose and agenda", "Today: agree the platforms and one PFR proposal")
d.cards(s, [
    dict(tag="Purpose", color=ORANGE, title="Why this meeting", bullets=[
        "Agree how the tools group into platforms, with one shared engine",
        "Agree one PFR proposal: an end-to-end process for energy costs and fiscal risks",
        "Agree the link with the macro unit (MFMod)",
        "Steer on what to drop, and on next steps to March"]),
    dict(tag="Your agenda", color=NAVY, title="Three topics", bullets=[
        ("1  Fiscal policy platform:", "crediting, FTT, MFMod, CBAM, excise"),
        ("2  Excise:", "the Excise Diagnostic and what comes next"),
        ("3  Adaptation platform:", "open question, to discuss"),
    ]),
    dict(tag="Then", color=TEAL, title="Status and decisions", bullets=[
        "The PFR proposal (slides 8-9)",
        "The latest MVP and how people use it (slides 10-11)",
        "Timeline, next steps and questions for you (slides 12-15)",
        "Appendix: the whole work programme, for reference"]),
], h=3300000)
d.so(s, "one engine (CPAT, in Excel and Python) feeds four platforms; the PFR platform is where the new proposal sits.")

# ---- one engine, four platforms (diagram) ---------------------------------------------------
s = d.content("The work programme in one picture", "One engine, four platforms")
ex, ey, ew, eh = M, 2350000, 2500000, 1900000
d.box(s, ex, ey, ew, eh, NAVY)
d.text(s, ex + 160000, ey + 170000, ew - 320000, 300000, "CORE ENGINE", size=9, bold=True, color="E3A36B")
d.text(s, ex + 160000, ey + 420000, ew - 320000, 500000, "CPAT, one set of equations", size=16, bold=True, color=WHITE, font=SERIF)
d.bullets(s, ex + 160000, ey + 1000000, ew - 320000, 900000,
          ["Excel-AI MVP v6.2 (checked reference)", "Coded CPAT (Python + Shiny)"], size=11, color=PALE)
plats = [
    ("fiscal", "Fiscal & policy platform", "Policy crediting, GHG accounting, FFI · FTT national · MFMod (macro unit) · CBAM tool · Excise Diagnostic"),
    ("climate", "Climate policy platform", "Pathways tool: coded CPAT inside CDPM (Edmundo) · pilots Libya, Egypt, Morocco"),
    ("pfr", "PFR platform", "Energy costs and fiscal risks · investment needs (NMIM) · Tax & Macro-Fiscal Dashboard · CPAT-LTS"),
    ("other", "Other", "AI working methods and guides · Energy GP cross-support · parked ideas · handover"),
]
px, pw, ph, pg = M + 3300000, CW - 3300000, 830000, 120000
py0 = 1720000
for i, (k, ttl, txt) in enumerate(plats):
    py = py0 + i * (ph + pg)
    d.box(s, px, py, pw, ph, WHITE)
    d.box(s, px, py, 60000, ph, GROUP[k])
    d.text(s, px + 200000, py + 110000, pw - 300000, 300000, ttl, size=14, bold=True, color=GROUP[k] if k != "other" else "9A6B12")
    d.text(s, px + 200000, py + 430000, pw - 300000, 400000, txt, size=11, color=BODY)
    d.arrow(s, ex + ew + 60000, ey + eh // 2, px - 60000, py + ph // 2, color="9AA8BC")
d.text(s, M, H - 830000, CW, 250000, CDPM_NOTE + ". Groupings follow Stephen's FY27 programme; appendix sections A-F follow the same order.",
       size=10, color=MUTED)

# ---- status at a glance ---------------------------------------------------------------------
s = d.content("Status at 20 October", "Where each platform stands")
d.cards(s, [
    dict(tag="Core engine", color=NAVY, title="Excel-AI MVP v6.2", pct=70, status="MVP complete", bullets=[
        "Mitigation, power, distribution built", "123 legacy bugs fixed, none copied", "Python port next"],
        foot="MVP 23 Oct · final Excel 4 Dec · Python 11 Dec"),
    dict(tag="Fiscal & policy", color=ORANGE, title="Platform being assembled", pct=None, bullets=[
        "CBAM tool: Egypt v3.0, global v1.0", "Zoe's FTT national: due 27 Nov", "Crediting / FFI: Feb-Mar",
        "Excise Diagnostic: handed to Dirk"], foot="FTT + MFMod working 12 Feb"),
    dict(tag="Climate policy", color=TEAL, title="Pathways with CDPM", pct=None, bullets=[
        "Edmundo meeting 19 Oct: on track", "Egypt proof of concept: -7.3% vs -8.0%", "CPAT sent; repo created"],
        foot="Integration 4-8 Jan · pilots 11-15 Jan"),
    dict(tag="PFR", color=BLUE, title="Pieces exist; process to build", pct=None, bullets=[
        "Dashboard: 53 indicators, 217 countries", "Investment needs NMIM v0.2", "Proposal: end-to-end energy costs and fiscal risks"],
        foot="To agree today"),
], h=3500000, size=11)

# ---- agenda 1: fiscal policy platform --------------------------------------------------------
s = d.content("Agenda 1 · Fiscal policy platform", "A fiscal policy platform, shared with the macro unit")
d.flow(s, [
    ("Coded CPAT", "Prices, revenue, emissions, households: the policy engine", "FBE3CE"),
    ("FTT national", "Technology diffusion (power, cars, freight, heat, steel) on CPAT's countries", "FBE3CE"),
    ("MFMod-Python", "Macro unit's model: investment, debt, growth (ModelFlow)", "FBE3CE"),
    ("Policy tools", "Crediting, GHG accounting, FFI · CBAM · excise", "FBE3CE"),
    ("One Python platform", "One repo, one guide; used with the macro unit", None),
], y=1800000, h=1350000)
d.table(s, ["Purpose", "Timeline", "Next steps"], [[
    "Give the macro unit and the fiscal team one code base for climate-related fiscal policy",
    "FTT national 27 Nov · FTT + MFMod working 12 Feb · crediting, GHG accounting, FFI 5 Mar",
    "Introduction to the macro unit; first call in November; decide MFMod in or out for February"]],
    widths=[0.34, 0.33, 0.33], y=3450000, size=11, h_row=380000)
d.so(s, "MFMod already supplies CPAT's fiscal multipliers; the platform makes that link two-way (investment, debt, fiscal space).")

# ---- agenda 2: excise -----------------------------------------------------------------------
appendix_bdef.core_excise(d)

# ---- agenda 3: adaptation -------------------------------------------------------------------
s = d.content("Agenda 3 · Adaptation platform", "Adaptation platform: an open question")
d.cards(s, [
    dict(tag="What we have", color=TEAL, title="Pieces that could serve adaptation", bullets=[
        "CDPM is a mitigation pathways model today", "Investment-needs method covers mitigation; adaptation totals come from UNEP's Adaptation Gap Report",
        "The Dashboard has fiscal space and risk indicators"]),
    dict(tag="Not in the FY27 plan", color=ORANGE, title="No adaptation milestone yet", bullets=[
        "No days budgeted to 20 March", "Would compete with the PFR proposal and the Python platform"]),
    dict(tag="To discuss", color=NAVY, title="Questions for you", bullets=[
        "What should an adaptation platform answer, and for whom?", "Is it part of the PFR process (fiscal risks from climate)?",
        "Who leads, and with which unit?", "Scope now, or a note for FY28?"]),
], h=3300000)
d.so(s, "a short scoping note after the trip, if you want it; otherwise park it to FY28.")

# ---- PFR proposal ---------------------------------------------------------------------------
s = d.content("PFR proposal", "An end-to-end process for energy costs and fiscal risks", bg=MIST)
d.flow(s, [
    ("1  Energy costs", "Investment needs and running costs by sector and country (NMIM, CDPM)", "DCE8F5"),
    ("2  Fiscal risks", "Subsidies, utility and SOE liabilities, falling fuel revenue", "DCE8F5"),
    ("3  Efficiency of spending", "Who pays (public, utility, private); value per dollar of public money", "DCE8F5"),
    ("4  Tax changes", "Carbon price, excise, subsidy reform, CBAM, through CPAT", "DCE8F5"),
    ("5  PFR chapter", "Energy and climate in expenditure policy, with numbers", None),
], y=1800000, h=1500000)
d.cards(s, [
    dict(tag="Why", color=BLUE, title="The gap", bullets=["PFRs lack a dedicated chapter on expenditure policy",
                                                         "Energy is where spending, risk and tax meet"]),
    dict(tag="What exists", color=BLUE, title="Tools already built", bullets=["NMIM v0.2 (who pays, budget need)",
                                                                             "Dashboard; fiscal-risk upgrade started", "CPAT revenue and prices"]),
    dict(tag="What is new", color=BLUE, title="The join", bullets=["One process, country data in, PFR text and tables out",
                                                                  "Same engine as the other platforms"]),
], y=3500000, h=1650000, size=11)

appendix_bdef.core_pfr_plan(d)

# ---- latest MVP in detail --------------------------------------------------------------------
s = d.content("The latest prototype", "Excel-AI MVP v6.2 (20 Oct): what is in it")
d.kpis(s, [("56 tabs", "about 560,000 formulas, 26 charts; Excel 365 with macros", NAVY),
           ("123", "legacy CPAT bugs fixed, none copied (9 high, 41 medium, 73 low)", ORANGE),
           ("12,992", "output codes unchanged v6.0 to v6.1; 0 error values (quick check)", TEAL),
           ("2022-41", "Egypt data, real 2026 USD; any country by ISO code", BLUE)], h=1150000)
d.table(s, ["Module", "What it does"], [
    ["Policies", "Carbon tax, ETS, feebates, fuel price and subsidy reform, methane fee, regulations as shadow prices"],
    ["Prices and fuel use", "Retail prices by fuel and sector; 16 subsectors x 8 fuels; road engineer model (switch)"],
    ["Power", "Engineer model: dispatch, retirement, investment; PPAs, renewable subsidies, excise (built 15-16 Oct, ahead of plan)"],
    ["Revenue, emissions, GDP", "Revenue by policy; CO2 to the inventory; non-CO2; NDC comparison; fiscal space and multipliers"],
    ["Distribution", "Household incidence by decile; recycling, transfers, Gini; matches a Python reference to 5e-13"],
], widths=[0.22, 0.78], y=3150000, size=11, h_row=340000)
d.notes(s, "Version history: v5.0 18 Oct to v6.2 20 Oct (power P1-P5, NDC comparison, distribution as legacy, producer subsidies, methane fee). Details in appendix A.")

s = d.content("Usability", "How people use it: from download to results in 15 minutes")
d.numbered(s, [
    ("Open and enable macros", "CPAT-AI-Mitigation-MVP-v6.2.xlsm in Excel 365; unblock the file"),
    ("Test once", "Alt+F8, run CheckBatchRun: MacroCheck shows PASS"),
    ("Read results", "Mitigation opens rolled up: one summary row per section; '+' expands. MTOutputs lists 71 key results; Charts plots them"),
    ("Change a policy", "MTInputs column K, e.g. CPLevelTarget (carbon price in the target year); it recalculates"),
    ("Run many scenarios", "Copy a definition column, set Run? = Yes, run RunAllScenarios; compare up to 8 in ScenarioCompare"),
    ("Another country", "Settings C4 = ISO code; paste the proprietary tables (Egypt shipped)"),
], y=1750000, step=480000, size=12)
d.box(s, M, 4750000, CW, 640000, WHITE)
d.text(s, M + 200000, 4800000, CW - 400000, 560000, [
    [("Every figure is traceable: ", {"bold": True, "color": NAVY}),
     ("each row has a code (e.g. egy.mit.ener.rod.gso.e.1), an equation note, and a LAMBDA formula; LegacyDiff lists every departure from legacy. ", {}),
     ("Python and Shiny: ", {"bold": True, "color": NAVY}),
     ("coded CPAT runs end to end (57 tests); the Shiny app has six sections, distribution to add; a simple app by 5 Feb.", {})]], size=11)

# ---- timeline ---------------------------------------------------------------------------------
s = d.content("Timeline", "From now to the end of the contract (20 March)")
d.gantt(s, [
    ("Excel-AI MVP", "core", 0, 0, NAVY, 0),
    ("Manila trip, leave", "", 0, 1, "D3DAE4", None),
    ("Final Excel-AI", "core", 2, 6, NAVY, 6),
    ("Final coded CPAT", "core", 2, 7, NAVY, 7),
    ("FTT national (Zoe)", "fiscal", 0, 5, ORANGE, 5),
    ("CDPM: workshop, coded CPAT, Morocco", "climate", 2, 7, TEAL, None),
    ("Pathways: integration, pilots", "climate", 11, 12, TEAL, 12),
    ("Shiny app", "core", 13, 15, NAVY, 15),
    ("FTT + MFMod", "fiscal", 16, 16, ORANGE, 16),
    ("Crediting, GHG accounting, FFI", "fiscal", 17, 19, ORANGE, 19),
    ("PFR proposal (to agree)", "pfr", 4, 9, "9DBBDD", None),
    ("Support to Dirk; handover", "", 20, 21, "D3DAE4", 21),
], WEEKS, y=1700000, row_h=270000, today=0.2,
    legend=[("Core engine", NAVY), ("Fiscal & policy", ORANGE), ("Climate", TEAL), ("PFR (proposed)", "9DBBDD"), ("Milestone", GOLD)])
d.notes(s, "Week-of dates. Christmas leave 23 Dec-1 Jan. PFR bar is a proposal: dates to agree. Buffer about 3 days before 4 Dec.")

s = d.content("Next steps", "Next steps by stage")
d.cards(s, [
    dict(tag="Before the trip (this week)", color=ORANGE, title="Now", bullets=[
        "Excel-AI MVP sent (due 23 Oct); Kacper holds the bug list", "Work programme (WPA) and one paper proposal",
        "Introduction to the macro unit requested"]),
    dict(tag="November-December", color=NAVY, title="Build", bullets=[
        "Final Excel-AI 4 Dec; coded CPAT 11 Dec", "CDPM user-flow workshop; coded CPAT to Edmundo", "Zoe's FTT 27 Nov",
        "PFR proposal: concept note (if agreed)"]),
    dict(tag="January-March", color=TEAL, title="Apply and hand over", bullets=[
        "Pathways pilots 11-15 Jan", "Shiny 5 Feb; FTT + MFMod 12 Feb", "Crediting, GHG accounting, FFI 5 Mar",
        "Handover 15-19 Mar; contract ends 20 Mar"]),
], h=3500000)

s = d.content("Questions for you", "Questions and decisions for today", bg=MIST)
d.numbered(s, [
    ("Platforms", "Do the four platforms, with one engine, match how you want to present the work?"),
    ("Macro unit", "Will you introduce us to the MFMod team? MFMod in or out for the 12 Feb milestone?"),
    ("Excise", "What next for the Excise Diagnostic beyond the handover to Dirk?"),
    ("Adaptation", "What should an adaptation platform answer? Scope now, or FY28?"),
    ("PFR proposal", "Go ahead with the end-to-end energy costs and fiscal risks process? Which pilot country?"),
    ("Resources", "Claude Code ($200); a wider prototype team; what to drop if time is short"),
], y=1750000, step=520000, size=12, col=NAVY)

# =============================================================================================
# APPENDIX
# =============================================================================================
d.divider("Appendix", "The work programme in full", "For reference: A core engine · B fiscal & policy platform · C climate policy platform · D PFR platform · E other · F plan, days and funding")

# ---- A. CORE ENGINE ----------------------------------------------------------------------------
d.divider("Appendix A", "Core engine: CPAT in Excel and Python", "One set of equations, two forms, kept equal by script. Status, the latest MVP, usability, quality, timeline")

s = d.content("A · Purpose", "Two engines, one set of equations")
d.table(s, ["", "Excel-AI-CPAT (CPAT-AI)", "Coded CPAT (Python + Shiny)"], [
    ["What it is", "Cut-down CPAT in Excel, rebuilt with AI from legacy CPAT", "Component-based Python model: mitigation and distribution"],
    ["Why", "An auditable replacement of the mitigation chain; every figure traceable", "Maintainable code with tests and an app; feeds pathways and platforms"],
    ["Users", "Modellers, economists, reviewers; the WB-IMF CPAT process", "Developers; CDPM integration; web users via Shiny"],
    ["Reference", "Legacy CPAT 1.0pre_456 (read only)", "Validated Excel logic, ported with identical codes"],
    ["Strength", "Familiar, transparent cell by cell, easy to audit", "Scales, tests, links to other Python models"],
    ["Limit", "Needs Excel 365 (LAMBDA); harder to scale", "Not open to non-coders without Shiny"],
], widths=[0.16, 0.42, 0.42], y=1800000, size=11, h_row=420000)
d.so(s, "Excel-AI validates the logic first; the validated logic is then ported to Python with the same codes.")

s = d.content("A · Purpose", "Excel-AI-CPAT serves three purposes")
d.cards(s, [
    dict(big="1", color=ORANGE, title="Input to the WB-IMF process", bullets=[
        "Agreed with Dirk: a contributor to future development", "Not yet the full CPAT; a base for one maintainable model"]),
    dict(big="2", color=ORANGE, title="GHG accounting and policy crediting", bullets=[
        "A transparent Excel tool for ex-ante DPF accounting and crediting", "Every figure traceable"]),
    dict(big="3", color=ORANGE, title="Cross-checking", bullets=[
        "An independent check on coded CPAT and legacy CPAT", "Differences point to errors on either side"]),
], h=3300000)

s = d.content("A · Latest MVP", "Version history: v5.0 to v6.2 in three days")
d.table(s, ["Version", "Date", "Highlight"], [
    ["v5.0", "18 Oct", "Consolidated release: subsidy and price-control phase-outs, producer subsidies, regulations as shadow prices"],
    ["v5.1-5.3", "19 Oct", "Residential LPG/kerosene switch, biomass leakage; legacy Manual inputs tab; full check repaired"],
    ["v5.4-5.6", "19 Oct", "Power P1-P5: per-scenario inputs, PPAs, renewable subsidies, rebate, excise, manual capacity, coal cost"],
    ["v5.7", "19 Oct", "NDC comparison (five NDC types); GDP memo rows"],
    ["v6.0", "19 Oct", "Distribution as legacy (ASPIRE, poverty-line transfers, 11 charts); bugs D10-D31 fixed"],
    ["v6.1", "20 Oct", "Producer-subsidy price effect; methane fee in fiscal space; NDC years 2021-24"],
    ["v6.2", "20 Oct", "Producer-subsidy reform as a separate wedge; L67 fixed without circularity; methane fee on own activity"],
], widths=[0.1, 0.1, 0.8], y=1800000, size=11, h_row=380000)
d.text(s, M, 5000000, CW, 400000, "Earlier: v0.1 on 8 Oct ($20/t cut fuel use 10.7%); v1.0 9 Oct; v4.0 16 Oct (49 legacy bugs then); build time cut from about 70 to 2 minutes (v4.9).",
       size=11, color=MUTED)

s = d.content("A · Latest MVP", "The model, section by section (numbered as legacy)")
d.table(s, ["Section", "Content"], [
    ["0-1", "Scenario assumptions; policies: carbon tax, ETS (cap, benchmarks, goal seek), feebates, fuel price reform, methane fee, phase-outs"],
    ["2", "Retail prices by fuel and sector; price indices"],
    ["3", "Power engineer model: costs, prices, demand, dispatch, retirement, investment"],
    ["5-8", "Fuel use: 16 subsectors x 8 fuels; road engineer model; residential leakage"],
    ["11-13", "Energy use; revenues by policy; CO2 calibrated to the inventory"],
    ["14-15", "Non-CO2 GHGs, methane by source, NDC comparison; local air pollutants"],
    ["16", "GDP: fiscal space, fiscal multipliers (global, income group, country, manual), external balance"],
    ["17-18", "Carbon pricing (implicit and effective rates, coverage); welfare (SCC, damages avoided)"],
    ["Distribution", "Incidence by decile in one year; recycling, Gini, compensation, ASPIRE and poverty-line transfers"],
], widths=[0.14, 0.86], y=1750000, size=11, h_row=360000)

s = d.content("A · Usability", "What the user inputs and sees")
d.cards(s, [
    dict(tag="Inputs", color=NAVY, title="MTInputs and Settings", bullets=[
        "Column J baseline; K the live policy; L+ scenario definitions", "Power inputs per scenario (e.g. NewInvCoa, PPAs, PowerRebates)",
        "TrModel switches the road engineer model", "Settings C4: the country ISO code"]),
    dict(tag="Outputs", color=ORANGE, title="Results and charts", bullets=[
        "Mitigation tab rolled up: one row per section", "MTOutputs: 71 key results by code; Charts: one series per scenario",
        "StoredResults; ScenarioCompare: up to 8 scenarios"]),
    dict(tag="Built in", color=TEAL, title="Scenarios and macros", bullets=[
        "Shipped: $20/t from 2027; Package A (to $50 by 2030 + feebates); Package B (ETS + $25)",
        "Macros: RunAllScenarios, CheckBatchRun, SolveETS", "Groups 1-2 plain formulas, so older Excel still calculates to 2040"]),
], h=3400000, size=11)

s = d.content("A · Traceability", "Every figure: equation, parameter, source")
d.flow(s, [
    ("Figure", "A result row with a code: country.mit.variable.sector.fuel.suffix.scenario", None and WHITE),
    ("Equation", "Equations_v6_2.md: the model in mathematics, by section", None),
    ("Formula", "Formulas_v6_2.md lists every formula; a LAMBDA per row; one formula per block", None),
    ("Parameter", "Row parameters in D:G, looked up from Inputs; code, source, confidence colour", None),
    ("Source", "Data tabs keep legacy layouts; LegacyDiff lists each departure", None),
], y=1800000, h=1600000, last_dark=False)
d.text(s, M, 3700000, CW, 900000, [
    [("Formula rules: ", {"bold": True, "color": NAVY}), ("no formula over 100 characters, no numeric literals, nothing country-specific in formulas.", {})],
    [("Independent check: ", {"bold": True, "color": NAVY}), ("a Python recomputation checks every number; checks also run in LibreOffice and openpyxl, so on Linux.", {})]],
    size=12, spacing=6)

s = d.content("A · Quality", "Bugs are never copied; issues are copied and tracked")
d.kpis(s, [("123", "legacy bugs not replicated: fixed, or the feature left out", ORANGE), ("0", "legacy bugs still replicated (since v4.10)", TEAL),
           ("39", "legacy issues replicated on purpose and tracked; 11 exceptions", NAVY), ("5", "Word plans with cell-by-cell fixes for the CPAT team", BLUE)])
d.flow(s, [("Preflight", "Find rows by label", None), ("Snapshot", "Copy, never save", None), ("Fix", "Real drag and copy-paste", None),
           ("Tests T1-T7", "Nothing else moved; past unchanged; no new errors; reconciled", None), ("FixLog", "Report written to the sheet", None)],
       y=3250000, h=1050000, last_dark=True, size=11)
d.text(s, M, 4500000, CW, 500000, "Worked example: legacy bug L1 (services electricity follows the industry multiple), macro FixL1_ServicesElectricity.bas. Other high bugs: L20 population 185m in 2041 vs UN 147m; L24 GDP effect compounds to about +10%.",
       size=11, color=MUTED)

s = d.content("A · Quality", "Checks, what was dropped, and known gaps")
d.cards(s, [
    dict(tag="Checks", color=TEAL, title="Run after every change", bullets=[
        "Quick check (2-5 min): structure, literals, recalculation, output codes: PASS at v6.1", "Full check (2.5 h): Python recomputation, sections 0-4 passed at v6.0",
        "LAMBDA group matches plain formulas to 3e-15", "Not yet: review in Excel 365 (13-step checklist ready)"]),
    dict(tag="Dropped 20 Oct", color=ORANGE, title="By decision", bullets=[
        "Road model follow-ups; partial-adjustment fuel model; real local-currency prices", "Air-pollution co-benefits and transport externalities (plans kept)",
        "Standalone distribution tool (19 Oct)"]),
    dict(tag="Known gaps", color=NAVY, title="To close", bullets=[
        "ETS goal seek does not converge for one definition (6.2%)", "Data requests: GAINS paths, VAT cells, other countries' tables, GHGs 2018-20",
        "Egypt only so far; legacy file 459 not yet reconciled"]),
], h=3500000, size=11)

s = d.content("A · Coded CPAT", "Coded CPAT and the Shiny app")
d.cards(s, [
    dict(tag="Status", color=NAVY, title="Runs end to end", bullets=[
        "About 20,300 lines of Python; 57 tests", "Mitigation (14 components) and distribution (merged 17 Oct)",
        "python run_model.py; shiny run app.py (Python 3.12)"]),
    dict(tag="Shiny app", color=TEAL, title="Six sections", bullets=[
        "Manual inputs, scenarios, raw results, indicators, charts, documentation", "Distribution not yet in the app", "Target: one input page and one results page by 5 Feb"]),
    dict(tag="Legacy export", color=ORANGE, title="Python back to legacy CPAT", bullets=[
        "Writes energy use, prices, household survey, IO_GTAP in legacy layout", "HHSurvey and IO_GTAP match legacy cell for cell (Egypt)",
        "Gap: 48 price columns blank"]),
], h=3300000, size=11)
d.so(s, "next: port the validated MVP logic, fix revenue units (USD m used as LCU), compare Python with Excel by 11 Dec.")

s = d.content("A · Status, timeline, next steps", "Core engine: ahead on build, Python is the work")
d.table(s, ["Milestone", "Window", "Due", "Status"], [
    ["1a_m1 Excel-AI MVP", "5-23 Oct", "23 Oct", "70%: equations list, traceability, known-gaps note"],
    ["1a_m2 Final Excel-AI", "5 Oct-4 Dec", "4 Dec", "50%: power and distribution already built"],
    ["1b_m1 Final Python", "2 Nov-11 Dec", "11 Dec", "10%: port, compare, write up 7-11 Dec"],
    ["1c_m1 Shiny app", "18 Jan-5 Feb", "5 Feb", "0%: one input page, one results page"],
], widths=[0.26, 0.18, 0.12, 0.44], y=1800000, size=11, h_row=380000)
d.cards(s, [
    dict(tag="Next steps", color=NAVY, title="Release and review", bullets=[
        "Release v6.2 properly (log, check, Word registers, deck)", "Excel 365 review on Windows", "Send data requests and legacy fix plans to the CPAT team"]),
    dict(tag="Risks / asks", color=ORANGE, title="Asks", bullets=[
        "Someone on Windows to run the checklist", "CPAT team sign-off on the replication rule", "A developer (e.g. Kacper) to share the Python port"]),
], y=3800000, h=1600000, size=11)

# ---- B. FISCAL & POLICY ------------------------------------------------------------------------
appendix_bdef.section_b(d)

# ---- C. CLIMATE POLICY PLATFORM ----------------------------------------------------------------
d.divider("Appendix C", "Climate policy platform: the pathways project", "Coded CPAT inside CDPM (formerly SiSePuede): one pathways tool and one user flow", accent="8CC7BD")

s = d.content("C · Purpose", "One pathways tool, one user flow", eyecolor=TEAL, source=CDPM_NOTE)
d.cards(s, [
    dict(tag="Purpose", color=TEAL, title="Two halves of one question", bullets=[
        "CDPM builds the physical trajectory", "CPAT adds policy, prices, revenue, households and health", "Neither can answer it alone"]),
    dict(tag="First use", color=TEAL, title="Finance and environment ministries", bullets=[
        "Tax, revenue use, investment needs, emissions", "Phase one policies: carbon pricing and public investment"]),
    dict(tag="Wider programme", color=TEAL, title="One country policy tool", bullets=[
        "With David Groves: CPAT, CDPM, MANAGE and MINDSET", "Pilots: Libya, Egypt, Morocco"]),
], h=3300000)

s = d.content("C · Why join them", "Each model fills the other's gap", eyecolor=TEAL, source=CDPM_NOTE)
d.table(s, ["", "CPAT", "CDPM (formerly SiSePuede)"], [
    ["Strength", "Retail prices, elasticities, carbon-tax revenue, subsidy reform, decile incidence, air-pollution health",
     "Full IPCC inventory, transformation library, least-cost power (NeMo-Mod), cost-benefit, about 1,000 futures"],
    ["Gap", "Energy CO2 only; no AFOLU, IPPU or waste; deterministic; fixed technologies", "No prices inside: no revenue, no fiscal incidence, no deciles"],
    ["Evidence", "-", "A carbon price moved CDPM's power mix by 0.000 pp through three routes"],
], widths=[0.14, 0.43, 0.43], y=1800000, size=11, h_row=520000)
d.so(s, "the tax enters CDPM as a demand response written by CPAT, not as a price.")

s = d.content("C · Partners", "Who does what", eyecolor=TEAL, source=CDPM_NOTE)
d.cards(s, [
    dict(tag="CDPM lead", color=TEAL, title="Edmundo Molina", bullets=["Tec de Monterrey", "Decides the implementation flow behind the user flow"]),
    dict(tag="Policy process", color=TEAL, title="David (Dave) Groves", bullets=["Leads the policy process; stress-tests the user flow", "Runs the wider programme"]),
    dict(tag="User interface", color=TEAL, title="Sonya Sampson", bullets=["User interface and user story", "Article 6 work for Egypt first"]),
    dict(tag="CPAT side", color=NAVY, title="Stephen Stretton", bullets=["Policy process and user flow write-up", "Delivers coded CPAT", "Kacper: power, prices database"]),
], h=3100000, size=11)

s = d.content("C · Architecture", "How coded CPAT plugs into CDPM", eyecolor=TEAL, source=CDPM_NOTE)
d.flow(s, [("Policy", "Carbon tax, revenue use (CPAT inputs)", "E1EFEC"), ("Prices", "CPAT retail prices", "E1EFEC"),
           ("CPAT engine", "Demand response, revenue", "E1EFEC"), ("CDPM pathways", "Emissions, transformations, power", "E1EFEC"),
           ("Welfare", "Air pollution, households (CPAT)", None)], y=1800000, h=1150000)
d.numbered(s, [
    ("Price channel", "Demand ratio R = (1 + k·tax)^ε from CPAT's elasticity table, applied to CDPM's own quantities"),
    ("Income channel", "CPAT's GDP effect for each revenue use scales CDPM's GDP"),
    ("Money channel", "Dollars / cost per dose = a transformation magnitude, capped at the largest dose measured"),
    ("Power channel", "Scale CDPM's minimum gas share by CPAT's ratio; CDPM picks the replacement"),
], y=3200000, step=440000, size=11, col=TEAL)
d.text(s, M, 5050000, CW, 300000, "Rule: prices, elasticities and money cross the bridge; tonnes never do. 218 of 2,638 input columns change in a taxed copy.", size=11, color=MUTED)

s = d.content("C · Data", "Sectoral and data mapping", eyecolor=TEAL, source=CDPM_NOTE)
d.bullets(s, M, 1800000, CW, 3200000, [
    ("Feasible:", "a consistent two-way CPAT-CDPM conversion despite different structures (Dec 2025 study)"),
    ("Structures:", "CDPM 5 sectors, 20 IPCC subsectors; CPAT 16 energy sectors from 23 IEA flows plus 5 non-energy"),
    ("Mapping:", "non-energy sectors map almost directly; energy sectors need proxy-based split factors; cement needs special care"),
    ("Shared inputs:", "GDP aligned on growth (PPP vs market dollars, 3.6x apart); power aligned on fossil share"),
    ("Base year 2022:", "fuel totals agree to 2.6% but the mix does not; it matters ±12% on the response"),
    ("Data pack:", "Egypt, MENA and global files for the cpat_sisepuede repo"),
], size=13, spacing=8)

s = d.content("C · Egypt proof of concept", "CDPM reproduces CPAT's carbon-tax response", eyecolor=TEAL, source="July test, tax $0 (2025) to $30/t (2040), 2022-40 · " + CDPM_NOTE)
d.kpis(s, [("-7.29%", "total cumulative cut in CDPM", TEAL), ("-8.03%", "total cumulative cut in CPAT", NAVY),
           ("exact", "industry, transport, buildings match", ORANGE), ("power", "the whole residual: -3.6% vs -5.4%", BLUE)])
d.bullets(s, M, 3200000, CW, 1800000, [
    "Nothing tuned to match: CPAT's table elasticities reproduce both CPAT runs to within 2.1%; a fitted exponent drifts 17% on average",
    "Grids aligned in the baseline first: the renewable gap went from 27 points to a 6-point overshoot",
    "Relative responses, not tonnes: CDPM's 2040 baseline is 1.57x CPAT's",
    "Egypt's effective carbon price is -$202/t in 2022 (net subsidy): the tool must show headline and effective prices"], size=12, spacing=6)

s = d.content("C · Egypt proof of concept", "The two-track tool: tax, revenue, investment", eyecolor=TEAL, source="Tax $0 (2025) to $50/t (2050) · " + CDPM_NOTE)
d.kpis(s, [("US$146.9bn", "revenue 2022-41; split 30% labour tax, 50% public investment, 20% transfers", TEAL),
           ("-11.5%", "2041 emissions: 810 to 717 MtCO2e with the tax", NAVY),
           ("-19.5%", "cumulative 2025-41 when public investment buys CDPM transformations", ORANGE),
           ("US$6.4/t", "tax in 2040 that pays for a 10% target (track 2)", BLUE)], h=1250000)
d.bullets(s, M, 3300000, CW, 1700000, [
    "Track 2: the pathway reaches 9.6% below baseline, costs US$17.2bn; the tax raises US$34.6bn",
    "GDP +0.10% (2041); 5,301 deaths averted by 2040; poorest 10% gain 0.43% of consumption, richest lose 0.10%",
    "Revenue surface: 35 CPAT runs (7 tax levels x 5 uses); worst error 0.09%"], size=12, spacing=6)

s = d.content("C · User flow", "One flow for the user", eyecolor=TEAL, source=CDPM_NOTE)
d.flow(s, [("1  Pathway", "BAU, NDC or custom", "E1EFEC"), ("2  Policy inputs", "Carbon tax, revenue use, other choices", "E1EFEC"),
           ("3  Run", "CPAT, pathways engine, cost modules", "E1EFEC"), ("4  Save", "The simulated pathway", "E1EFEC"),
           ("5-6  Optimise", "Optional: tax level, investments, allocation", None)], y=1800000, h=1250000)
d.bullets(s, M, 3300000, CW, 1800000, [
    "Every run reports emissions, revenue, investment needs, prices, GDP, households, carbon markets; distribution and CBAM to add",
    "Three requirements: an agreed user flow; convertibility between models; one interface with integrated policy inputs",
    "Early November workshop on user experience, kept separate from the technical integration"], size=12, spacing=6)

s = d.content("C · Methodology", "Open issue: incremental vs capital costs", eyecolor=TEAL, source=CDPM_NOTE)
d.cards(s, [
    dict(tag="Problem", color=ORANGE, title="Two kinds of cost are mixed", bullets=[
        "CDPM: subsidies pay part of private marginal abatement cost", "Upfront public investment is treated like subsidies",
        "CDPM's 'investment' is closer to net incremental cost"]),
    dict(tag="Fix", color=TEAL, title="Separate them", bullets=[
        "Running subsidies vs upfront public investment", "Show how a carbon tax lowers each", "Report dollars and marginal costs"]),
    dict(tag="Options", color=NAVY, title="Where the fix lives", bullets=[
        "Bring some CDPM data into coded CPAT", "Or put demand reductions into CDPM", "Next: a short methodology note"]),
], h=3200000)
d.so(s, "this is also the bridge to the PFR proposal: the same split drives the budget cost of energy investment.")

s = d.content("C · Status, timeline, next steps", "Climate platform: on track to pilots in January", eyecolor=TEAL, source=CDPM_NOTE)
d.table(s, ["When", "Milestone"], [
    ["13 Oct", "Scoping meeting with David Groves and Sonya Sampson"],
    ["19 Oct", "Edmundo meeting went well; CPAT sent; CPAT-CDPM repo created (milestone done)"],
    ["Late Oct", "Egypt mission: Article 6 and carbon markets"],
    ["Early Nov", "User-flow workshop"],
    ["Early Dec", "Final coded CPAT to Edmundo; Morocco first cut and user feedback"],
    ["4-15 Jan", "Integration (prices database check, Kacper); pilot projects; tool complete 15 Jan"],
    ["29 Jan / Feb", "Applied in Morocco; full coded integration end January; mission in February"],
], widths=[0.16, 0.84], y=1750000, size=11, h_row=350000)
d.text(s, M, 4700000, CW, 600000, [
    [("Risks: ", {"bold": True, "color": NAVY}), ("double counting of demand; coded CPAT reproduces about half of Excel's Egypt response so far; CPAT runs to 2041, CDPM to 2050; failed power solves return zeros. ", {}),
     ("Asks: ", {"bold": True, "color": NAVY}), ("agree the cost split before integration; confirm when CDPM switches to coded CPAT.", {})]], size=11)

# ---- D, E, F --------------------------------------------------------------------------------------
appendix_bdef.section_d(d)
appendix_bdef.section_e(d)
appendix_bdef.section_f(d)

d.save(str(HERE / "Sebastian-Platforms-v2.pptx"))
print("slides:", d.n)
