"""Core slides on excise and the PFR plan, and appendix sections B (fiscal & policy), D (PFR), E (other), F (plan)."""
from kit import *

CDPM_NOTE = "CDPM = the model formerly called SiSePuede"
BROWN = "9A6B12"


# =============================================================================================
# CORE: agenda 2 (excise) and the PFR plan
# =============================================================================================
def core_excise(d):
    s = d.content("Agenda 2 · Excise", "Excise Diagnostic: delivered; what next is the question")
    d.kpis(s, [("3", "countries: Sierra Leone, Uganda, Liberia (draft v0.3)", ORANGE),
               ("1.5-1.7%", "of GDP from Uganda's energy, tobacco, alcohol and SSB excises", NAVY),
               ("16 Oct", "handed to Dirk, early (MVP 25 Sep)", TEAL),
               ("no code", "run with Claude or Copilot; Word report + Excel workbook", BLUE)], h=1100000)
    d.cards(s, [
        dict(tag="What it is", color=ORANGE, title="A synthesis tool, not a model", bullets=[
            "Sorts every excise line into seven categories", "Each rate as US$ per unit, share of price, revenue % of GDP",
            "Shows erosion, gaps and incoherence; a four-level report"]),
        dict(tag="Options for next", color=NAVY, title="The proposal on the table", bullets=[
            "Three-part report: excises, externalities, pathways", "28 key questions, 20 from climate and health programmes",
            "Pilots: Colombia or Chile (tier 1); Uganda (tier 2)"]),
        dict(tag="To decide", color=TEAL, title="Questions for you", bullets=[
            "Who owns it after Dirk's handover?", "Extend to more countries, or fold into the PFR process?",
            "Link to CPAT prices (fuel excise is most of the revenue)"]),
    ], y=3100000, h=2100000, size=11)


def core_pfr_plan(d):
    s = d.content("PFR proposal · how and when", "How we would build it: mostly from parts that exist", bg=MIST)
    d.table(s, ["Step", "Question for the finance ministry", "Tool that answers it today", "Gap to close"], [
        ["1 Energy costs", "What must be invested, and what does it cost to run?", "NMIM v0.2; investment-needs v1.1; CDPM pathways",
         "Price response; CCDR data per country"],
        ["2 Fiscal risks", "What could land on the budget?", "Fiscal Risk Dashboard v1.5; CPAT subsidies and fuel revenue",
         "Utility and SOE liabilities; owner and plan for the upgrade"],
        ["3 Efficiency", "Who pays, and what does each public dollar buy?", "NMIM who-pays split (budget need 14% world, 29% EMDCs); Dashboard PIMA block",
         "IMF PIMA and C-PIMA scores; investment approach V3 not yet agreed"],
        ["4 Tax changes", "What do carbon pricing, excise and subsidy reform raise?", "Excel-AI CPAT revenue by policy; Excise Diagnostic; CBAM tool",
         "One documented chain to the debt equation (V2)"],
        ["5 PFR chapter", "What goes in the expenditure-policy chapter?", "-", "Template, pilot country, owner"],
    ], widths=[0.14, 0.27, 0.32, 0.27], y=1750000, size=10, h_row=430000)
    d.cards(s, [
        dict(tag="Phase 1 (proposed)", color=BLUE, title="Concept note", bullets=["Map the five steps to tools and data", "Agree V2-V3 of the investment approach"]),
        dict(tag="Phase 2 (proposed)", color=BLUE, title="One pilot country", bullets=["Candidates: Uganda (PFR expenditure pilot), Egypt or Morocco",
                                                                                       "Same engine as the pathways tool"]),
        dict(tag="Phase 3 (proposed)", color=BLUE, title="Template for PFRs", bullets=["Chapter template and tables", "Hand to PFR teams with a guide"]),
    ], y=4450000, h=1150000, size=10)
    d.notes(s, "Dates to agree: not in the FY27 plan, which is full to 20 March. Something would have to be dropped or the plan extended. Uganda is named in the Fiscal-LTS master doc as the pilot for the PFR expenditure section.")


# =============================================================================================
# B. FISCAL & POLICY PLATFORM
# =============================================================================================
def section_b(d):
    d.divider("Appendix B", "Fiscal & policy platform", "Policy crediting, GHG accounting and FFI · FTT national · MFMod and the macro unit · CBAM tool · Excise Diagnostic (fiscal tools)")

    s = d.content("B · Overview", "Five parts, one Python platform")
    d.cards(s, [
        dict(tag="2a", color=ORANGE, title="Crediting, GHG accounting, FFI", bullets=["Ex-ante GHG figure for policy-based DPFs", "Due 5 Mar 2027"], foot="15 days, TF-TCAF"),
        dict(tag="2b", color=ORANGE, title="FTT national (Zoe)", bullets=["Technology diffusion on CPAT's countries", "Due 27 Nov"], foot="Zoe delivers; Stephen reviews"),
        dict(tag="2b", color=ORANGE, title="FTT + MFMod", bullets=["One Python platform, one guide", "Due 12 Feb 2027; MFMod tbc"], foot="5 days, TF-CSF-LTS"),
        dict(tag="Tool", color=ORANGE, title="CBAM policy tool", bullets=["Egypt v3.0; global MVP v1.0 (59 rows)", "Feeds NMIM process emissions"], foot="No dated milestone"),
        dict(tag="3a", color=ORANGE, title="Excise Diagnostic", bullets=["Three countries; handed to Dirk", "Fiscal tools layer"], foot="Support 8-12 Mar"),
    ], h=3300000, size=10)

    s = d.content("B · Crediting, GHG accounting, FFI (2a)", "Policy crediting: a GHG figure for policy reforms")
    d.cards(s, [
        dict(tag="Why", color=ORANGE, title="No tool for policy-based lending", bullets=[
            "The Climate Change Action Plan update recommends GHG accounting for DPFs", "DPFs turn on policy reforms, not physical assets"]),
        dict(tag="What", color=ORANGE, title="A simple Excel-AI-CPAT tool", bullets=[
            "GHG figure for a fuel-pricing policy, with the Total Carbon Price", "Then policy crediting and the FFI (cross-border externalities)",
            "Parameters and assumptions visible"]),
        dict(tag="When", color=NAVY, title="15 Feb-5 Mar 2027", bullets=[
            "15-19 Feb: GHG accounting MVP", "22 Feb-5 Mar: crediting and FFI", "Depends on final Excel-AI (4 Dec)"]),
    ], h=3000000)
    d.so(s, "asks: confirm the TF-TCAF-Crediting charge code; requirements for accounting, crediting and FFI are still 'to be added'.")

    s = d.content("B · FTT and MFMod (2b)", "Existing models working together")
    d.flow(s, [("Zoe's FTT national", "Data so FTT runs on CPAT's countries: power, cars, freight, heat, steel. Due 27 Nov", "FBE3CE"),
               ("FTT ready to run", "Data aligned to CPAT countries", "FBE3CE"),
               ("MFMod-Python (tbc)", "Macro unit's model: investment, debt, all taxes and spending", "FBE3CE"),
               ("One platform, 8-12 Feb", "CPAT + FTT + MFMod; one guide; due 12 Feb", None)], y=1800000, h=1300000)
    d.bullets(s, M, 3350000, CW, 1800000, [
        ("Macro link today:", "CPAT's GDP block uses MFMod fiscal multipliers; the debt equation does not replace MFMod, FSA or DSA"),
        ("Master doc:", "work with the macro team on MFMod for investment-macro analysis (memo item, TBC); contact Charl Jooste"),
        ("Grant:", "C2 'connection between CPAT and FTT', FY28 USD 25,000; FTT national database possibly with Exeter"),
        ("Gap:", "no ModelFlow material in any repo yet; out of scope: full FTT or MFMod integration into CPAT")], size=12, spacing=6)

    s = d.content("B · CBAM policy tool", "What a domestic carbon price does to the EU border charge")
    d.table(s, ["Egypt v3.0 (2030, $20/t, long run)", "Heavy-industry emissions", "Border charge per tonne", "Of which credit"], [
        ["1A Comprehensive levy", "-12.1%", "-26.7%", "-18.3%"], ["2A/2B Upstream levy", "-7.3%", "-14.7%", "-9.1%"],
        ["3A Downstream price", "-12.1%", "-26.7%", "-18.3%"], ["3B With feebate", "-8.9%", "-8.4%", "0.0%"],
        ["3C With abatement fund", "-19.8%", "-32.8%", "-16.8%"]], widths=[0.34, 0.22, 0.22, 0.22], y=1750000, size=11, h_row=340000)
    d.bullets(s, M, 3950000, CW, 1300000, [
        "Credit for the domestic price does more for the border charge than abatement; a feebate forfeits the credit; a fund roughly doubles abatement",
        "1A nationally: 71.5% of emissions priced, $5.6bn revenue, 40.7 Mt cut, 1,659 deaths avoided a year",
        "66 sourced inputs; EU price $100/t with no phase-in; about 40% of the long-run response by 2030"], size=11, spacing=4)

    s = d.content("B · CBAM policy tool", "Global MVP, dashboard and limits")
    d.cards(s, [
        dict(tag="Global MVP v1.0", color=ORANGE, title="59 country rows", bullets=[
            "Egypt reproduces the study", "Other countries: heavy-industry and border-charge rows", "Placeholders register for unsourced values"]),
        dict(tag="CBAM Dashboard v1.0", color=ORANGE, title="Five countries, 2024-35", bullets=[
            "Egypt, Albania, Uzbekistan, Azerbaijan, India", "Price that captures CBAM revenue: Egypt $94/t, India $92/t (2030)",
            "CBAM is the trigger for a domestic price, not the main revenue"]),
        dict(tag="Limits and next", color=NAVY, title="To decide", bullets=[
            "No trade or leakage model; credit rules may change", "No dated milestone in the FY27 plan", "Ask: CBAM inside CPAT, a priority?"]),
    ], h=3300000, size=11)

    s = d.content("B · Excise Diagnostic (fiscal tools)", "Excise Diagnostic: method and results")
    d.cards(s, [
        dict(tag="Method", color=ORANGE, title="Every line, three measures", bullets=[
            "Seven categories, five fields; 30 standard lines", "US$ per unit, share of price, revenue % of GDP",
            "Four-level report: minister's page down to legal history"]),
        dict(tag="Sierra Leone", color=ORANGE, title="1.25-1.28% of GDP", bullets=[
            "About 90% from fuel", "Cigarettes about US$0.14 a pack (14% of price)", "Peer review: 49 entries, 16 corrected"]),
        dict(tag="Uganda, Liberia", color=ORANGE, title="1.5-1.7% of GDP (Uganda)", bullets=[
            "Energy, tobacco, alcohol, SSBs", "Uganda lines not yet peer reviewed", "Liberia draft v0.3 (not for citation)"]),
        dict(tag="Usability", color=NAVY, title="Run with AI, no code", bullets=[
            "Claude Projects or Copilot agent mode", "Stages A1-A6; human review between stages", "Word report + Excel workbook; peer review mandatory"]),
    ], h=3400000, size=10)

    s = d.content("B · Status, timeline, next steps", "Fiscal & policy platform: timeline")
    d.table(s, ["When", "Milestone", "Next step / ask"], [
        ["16 Oct (done)", "Excise Diagnostic to Dirk (3a_m2)", "Name an owner; decide on the proposal"],
        ["27 Nov", "Zoe's FTT national project (2b_m1)", "Ask Zoe for a review date"],
        ["8-12 Feb", "FTT + MFMod on one platform (2b_m2, due 12 Feb)", "Decide MFMod in or out; macro-unit introduction"],
        ["15 Feb-5 Mar", "GHG accounting, crediting, FFI (2a_m1, due 5 Mar)", "Confirm charge code and requirements"],
        ["8-12 Mar", "Follow-up support to Dirk", "-"],
        ["tbc", "CBAM in CPAT", "Decide priority"],
    ], widths=[0.18, 0.44, 0.38], y=1800000, size=11, h_row=380000)


# =============================================================================================
# D. PFR PLATFORM
# =============================================================================================
def section_d(d):
    d.divider("Appendix D", "PFR platform", "Energy costs and fiscal risks · efficiency of public expenditure · integration of tax changes · Tax & Macro-Fiscal Dashboard · investment needs · CPAT-LTS",
              accent="9DBBDD")

    s = d.content("D · Why", "The gap: PFRs have no expenditure-policy chapter on climate", eyecolor=BLUE)
    d.cards(s, [
        dict(tag="PFRs today", color=BLUE, title="Debt, deficits, risks", bullets=[
            "Climate chapters on fiscal risk (macro) and revenue (e.g. carbon taxes)", "No dedicated chapter on expenditure policy"]),
        dict(tag="CCDRs and LTSs", color=BLUE, title="The 'what', not the 'how'", bullets=[
            "Emissions pathways, sector transformation, co-benefits", "Mitigation is mostly infrastructure: energy, transport, urban"]),
        dict(tag="Fiscal-LTS", color=NAVY, title="'How to pay for it'", bullets=[
            "Less detailed than a CCDR, built on CCDR data", "For PFRs and CCDR macro chapters"]),
    ], h=2700000)
    d.table(s, ["Fiscal-LTS component", "What it does"], [
        ["1 Public investment model", "Public investment to emissions cuts (CCDR data; CDPM pseudo-model)"],
        ["2 Private investment policy model", "Power sector: carbon pricing, procurement, fuel price reform"],
        ["3 Debt equation", "IMF revenue forecasts in forward projections, with gains from tax reform"],
        ["4 Investment needs", "Total mitigation needs (CCDR) against current investment and maintenance spending"],
    ], widths=[0.32, 0.68], y=4600000, size=10, h_row=230000)

    s = d.content("D · Energy costs and fiscal risks", "The investment approach in three versions", eyecolor=BLUE)
    d.cards(s, [
        dict(tag="V1 · in progress", color=BLUE, title="Process emissions and green investment", bullets=[
            "Process emissions for CBAM sectors (with Karly)", "Green public investment from CCDR data (prototyped)"], pct=50, status="partly done"),
        dict(tag="V2 · to be agreed", color=BLUE, title="Debt projections", bullets=[
            "Debt(t+1) = debt·(1+i)/(1+g+π) - primary balance + adjustments + disaster costs", "IMF 10-year projections; not a DSA"], pct=0, status="not started"),
        dict(tag="V3 · to be agreed", color=NAVY, title="Efficiency and energy fiscal risks", bullets=[
            "Efficiency of public investment", "Energy fiscal costs and risks for finance ministries", "= the core of the PFR proposal"], pct=0, status="not started"),
    ], h=3400000, size=11)
    d.so(s, "V2 and V3 have no days or dates; the PFR proposal would give them a home.")

    s = d.content("D · Efficiency of public expenditure", "Efficiency: reform is finance at negative cost", eyecolor=BLUE)
    d.kpis(s, [("1.5 / 4", "median PEFA public investment score (61 economies); only 2 above 3", BLUE),
               ("87%", "median capital budget execution; 11 economies spend under half", ORANGE),
               ("+16%", "EMDC investment needs after the PIMA efficiency adjustment", NAVY),
               ("13/27/40%", "efficiency gaps: advanced, emerging, low-income (IMF 2015)", TEAL)], h=1250000)
    d.bullets(s, M, 3300000, CW, 1800000, [
        "Investment-needs v1.1: the PIMA adjustment lifts world needs from USD 5.2tn to 5.8tn",
        "Dashboard PIMA block: BOOST budget execution, PEFA PI-11 and PI-12.2, UNDP climate-PFM flag",
        ("Gap:", "IMF PIMA and C-PIMA scores not yet in the data; V3 of the investment approach not agreed")], size=12, spacing=6)

    s = d.content("D · Integration of tax changes", "Tax changes: the pieces exist, the chain does not", eyecolor=BLUE)
    d.flow(s, [("Excise Diagnostic", "Standardised excise data for PFRs", "DCE8F5"), ("CPAT prices database", "Fuel prices, taxes, subsidies; feeds all models", "DCE8F5"),
               ("Excel-AI CPAT", "Revenue by policy; fiscal space", "DCE8F5"), ("Debt equation (V2)", "Tax gains in forward projections", "DCE8F5"),
               ("PFR chapter", "Revenue and spending together", None)], y=1800000, h=1300000)
    d.so(s, "no single documented process links excise and tax changes through to debt and the PFR chapter: step 4 of the proposal.", y=3500000)

    s = d.content("D · Investment needs", "How much, and who pays", eyecolor=BLUE, source=CDPM_NOTE)
    d.kpis(s, [("USD 5.2tn", "net zero investment in 2030, of which 1.7tn incremental (v1.1)", BLUE),
               ("USD 38bn", "net cost a year at 6%, about $41/t; saves $43bn with co-benefits", TEAL),
               ("14%", "direct budget need, world 2030: USD 708bn (NMIM v0.2)", ORANGE),
               ("29%", "in other EMDCs (USD 379bn); two-thirds in low-income countries", NAVY)], h=1250000)
    d.table(s, ["Who pays (world 2030)", "Share", "USD bn"], [
        ["P1 pure public", "10%", "511"], ["P4 viability gap", "13%", "-"], ["P5 pure private", "31%", "1,567"],
        ["Direct budget need (P1 + 30% of P4)", "14%", "708"]], widths=[0.6, 0.2, 0.2], y=3300000, size=11, h_row=320000)
    d.text(s, M, 5000000, CW, 300000, "NMIM: 186 countries, 2024-50, CPAT norms; workbook equals Python for 8 test countries. Type shares are judgements; market data end 2013; no price response.",
           size=10, color=MUTED)

    s = d.content("D · Fiscal risks", "Tax & Macro-Fiscal Dashboard and the fiscal-risk upgrade", eyecolor=BLUE)
    d.cards(s, [
        dict(tag="3b · delivered", color=BLUE, title="53 indicators, 217 countries", bullets=[
            "9 blocks, 16 public datasets; scored 1-10, sourced, nothing imputed", "Judging blocks: fiscal space, delivery (PIMA), cost of capital, financing risk",
            "Handed to Dirk; support 8-12 Mar"]),
        dict(tag="Findings (98 CCDR economies)", color=BLUE, title="Where risk sits", bullets=[
            "19 net fuel exporters face a revenue cliff", "23 economies get over half their power from oil", "Official-private borrowing spread median 2.1pp, up to 8.7pp"]),
        dict(tag="Upgrade v1.5", color=NAVY, title="Climate-fiscal risk tool", bullets=[
            "Climate risk (INFORM), adaptation (ND-GAIN), mitigation, revenue, spending", "PFR models catalogue (Climate Scan, CPAT, MFMod)",
            "No owner, version plan or dates yet"]),
    ], h=3500000, size=10)

    s = d.content("D · CPAT-LTS", "CPAT-LTS (CSF grant): components and money", eyecolor=BLUE)
    d.table(s, ["Component", "Timeline", "USD"], [
        ["C1 Securing continuation (data)", "07/2025-06/2028", "50k a year FY26-28"],
        ["C2 Power tipping points; transport technology", "07/2025-06/2027", "25k + 25k a year FY26-27"],
        ["C2 CPAT-FTT connection", "07/2027-06/2028", "25k FY28"],
        ["C3 Policy packages (mixes, costs, financing)", "07/2025-06/2027", "25k a year FY26-27"],
        ["C4 Online tool; capacity building", "07/2025-06/2028", "25k a year; 50k, 50k, 125k"],
    ], widths=[0.5, 0.25, 0.25], y=1750000, size=11, h_row=360000)
    d.bullets(s, M, 4050000, CW, 1300000, [
        ("Status (July 2026):", "C1 data done; C2 transport model done, power in progress; C3 policy mixes in main CPAT; C4 Python tool done"),
        ("Countries:", "Uganda, Egypt, Morocco (TOR); Morocco application in FY27; capacity building done in Uzbekistan and Azerbaijan"),
        ("Partners:", "Tec de Monterrey USD 215k (year one), of which USD 20k for CPAT integration")], size=11, spacing=4)

    s = d.content("D · Status, timeline, next steps", "PFR platform: from parts to a process", eyecolor=BLUE)
    d.table(s, ["Piece", "Status", "Next step"], [
        ["Tax & Macro-Fiscal Dashboard", "Delivered (Dirk); support 8-12 Mar", "Owner; PIMA and C-PIMA data"],
        ["Fiscal Risk Dashboard v1.5", "Prototype", "Owner and plan (proposal step 2)"],
        ["Investment needs (NMIM v0.2)", "Built 20 Oct", "Review market structure; link to MFMod debt"],
        ["Investment approach V1 / V2 / V3", "In progress / to agree / to agree", "Agree V2-V3 (proposal)"],
        ["CPAT-LTS grant", "On track; FY27 Morocco", "Online coded CPAT"],
        ["End-to-end process (proposal)", "Proposed today", "Concept note; pilot country"],
    ], widths=[0.34, 0.3, 0.36], y=1800000, size=11, h_row=380000)


# =============================================================================================
# E. OTHER
# =============================================================================================
def section_e(d):
    d.divider("Appendix E", "Other", "AI working methods and guides · Energy GP cross-support · global coverage · parked and dropped · Egypt and training", accent=GOLD)

    s = d.content("E · AI working methods", "How we use AI: one tool, two modes", eyecolor=BROWN)
    d.cards(s, [
        dict(tag="Main tool", color=GOLD, title="GitHub Copilot (and Claude)", bullets=["Many models; VS Code or desktop app", "Default allowance about $19 a month"]),
        dict(tag="Personal mode", color=GOLD, title="Code, no proprietary data", bullets=["Proprietary data taken out of code and scripts"]),
        dict(tag="Enterprise mode", color=GOLD, title="Confidential data and reports", bullets=["Enterprise AI incl. MAI", "Ask: enterprise access (blocked)"]),
        dict(tag="For others", color=GOLD, title="Guides and scripts", bullets=["Copilot user guide (33 slides)", "Repos in cpmodel so IMF colleagues can work in them"]),
    ], h=2700000, size=10)
    d.numbered(s, [("AI recommends data", "Which datasets and where to find them"), ("Scripts download it", "Source list and download scripts"),
                   ("AI builds the database and tool", "Structured, auditable, with sources"), ("People check", "AI makes many mistakes; about 3 days of checking per tool")],
               y=4550000, step=270000, size=10, col=GOLD)

    s = d.content("E · Energy GP, coverage, Egypt, training", "Other work", eyecolor=BROWN)
    d.cards(s, [
        dict(tag="3c · not funded", color=GOLD, title="Energy GP cross-support", bullets=[
            "Treasury GP finance model with the Energy GP model", "For concessional financing; code help", "16 days unfunded; 0 slotted"]),
        dict(tag="Global coverage", color=GOLD, title="Country reach", bullets=[
            "CPAT about 200 countries", "Dashboard 217; investment needs 195; NMIM 186", "CBAM global MVP 59 rows"]),
        dict(tag="Egypt", color=GOLD, title="First pilot", bullets=[
            "Planning and Egypt work done (28 Sep-2 Oct)", "CBAM study v3.0", "Pilot projects from 11 Jan"]),
        dict(tag="Training", color=GOLD, title="Capacity building", bullets=[
            "Done: Uzbekistan, Azerbaijan", "Uganda requested training", "Grant: USD 50k, 50k, 125k (FY26-28)"]),
    ], h=3400000, size=10)

    s = d.content("E · Materials", "Where to find the decks, guides and models", eyecolor=BROWN)
    d.table(s, ["Material", "Where (repo / folder)"], [
        ["Overview deck (this deck's predecessor): AIClimate-FiscalTools v2.5; FY27 work plan v2.6", "PlanningPPTsAndGeneral and CPAT-AI: 'Presentations - Overarching'"],
        ["Excel-AI MVP v6.2, QuickStart, method notes, registers, overview decks", "CPAT-AI: cpat_excel/ (mitigation_copypaste/)"],
        ["Coded CPAT, Shiny app, developer guide, legacy export", "CPAT-AI: cpat_coded/"],
        ["Pathways: CDPM decks, walkthrough, bridge map, methodology, data mappings", "CPAT-CDPM: MainProjectFolder/"],
        ["CPAT-LTS grant deck and Fiscal-LTS master document", "CPAT-CDPM: MainProjectFolder/9. CPAT-LTS Project, 8. Archive"],
        ["Excise Diagnostic: user guide, proposal, working paper v0.9", "new-ai-tools: ExciseDiagnostic/"],
        ["Tax & Macro-Fiscal Dashboard: user guide, v1.0 deck; fiscal-risk and CBAM dashboards", "new-ai-tools: TaxMacroFiscalDashboard/ (upgrades/)"],
        ["CBAM policy tool: guide v3.0, Egypt study, global MVP", "new-ai-tools: CBAMPolicyTool/"],
        ["Investment needs: v1.1, NMIM v0.2, methodology; GitHub Copilot user guide", "new-ai-tools: InvestmentNeeds/, Docs and PPTs/"],
    ], widths=[0.58, 0.42], y=1700000, size=10, h_row=360000)

    s = d.content("E · Parked and dropped", "What we are not doing this year", eyecolor=BROWN)
    d.table(s, ["Item", "Status"], [
        ["Patching legacy Excel (new tab, LAMBDAs); full CPAT replacement", "Dropped from FY27"],
        ["Merging FTT or MFMod into CPAT", "Out of scope"],
        ["Validating legacy CPAT", "Separate note, out of scope"],
        ["ETS pilot, downstream pricing, CPAT paper (April items)", "Confirm done, handed over or dropped"],
        ["Policy issue reports (e.g. feebates)", "Dropped"],
        ["Air-pollution co-benefits, transport externalities, road model follow-ups", "Dropped 20 Oct (plans kept)"],
        ["Single Python platform; global CPAT; MAC curve for policies; investment framework for PFRs", "Parked ideas"],
    ], widths=[0.68, 0.32], y=1800000, size=11, h_row=380000)


# =============================================================================================
# F. PLAN, DAYS AND FUNDING
# =============================================================================================
def section_f(d):
    d.divider("Appendix F", "Plan, days and funding", "FY27 work plan v2.6 (15 Oct 2026): objectives, days, charge codes, milestones, sequencing, risks")

    s = d.content("F · Objectives", "Three objectives, 125 funded days")
    d.cards(s, [
        dict(tag="Objective one · 50 days", color=NAVY, title="Engine rebuild (1a, 1b, 1c)", bullets=[
            "Excel-AI-CPAT and coded CPAT, one set of equations", "Shiny app", "A clean engine for the new WB-IMF interface"]),
        dict(tag="Objective two · 35 days", color=ORANGE, title="Applications (2a, 2b, 2c)", bullets=[
            "GHG accounting, crediting, FFI", "FTT and MFMod", "Pathways tool and pilots"]),
        dict(tag="Objective three · 40 days", color=TEAL, title="Other AI tools (3a, 3b, 3c)", bullets=[
            "Excise Diagnostic and Dashboard: 4 of 4 milestones done", "Energy GP if funded", "25 days spent, 15 left"]),
    ], h=3000000)
    d.so(s, "the plan is full to 20 March: anything added must replace something.")

    s = d.content("F · Days and funding", "Days and charge codes")
    d.table(s, ["Code", "Sub-objective", "Days", "Funding"], [
        ["1a", "Excel-AI-CPAT", "20", "TF-CSF-LTS"], ["1b", "Coded-Conversion (Python)", "15", "TF-CSF-LTS"], ["1c", "Shiny app", "15", "TF-CSF-LTS"],
        ["2a", "GHG accounting, crediting, FFI", "15", "TF-TCAF-Crediting (code tbc)"], ["2b", "FTT and MFMod (tbc)", "5", "TF-CSF-LTS"],
        ["2c", "Pathways, pilots, planning, Egypt", "15", "TF-CSF-LTS"], ["3a", "Excise Diagnostic", "20", "TF-GTP-NONSECO"],
        ["3b", "Tax and Macro-Fiscal Dashboard", "20", "TF-GTP-SECO"], ["", "Total funded (+16 unfunded, e.g. Energy GP)", "125", "141 in all"],
    ], widths=[0.08, 0.5, 0.1, 0.32], y=1750000, size=11, h_row=330000)
    d.text(s, M, 5150000, CW, 300000, "TF-CSF-LTS AA-P508228-GPLG-TF0C8701 (70 days) · TF-GTP-NONSECO ...TF0C9102 (20) · TF-GTP-SECO ...TF0C7596 (20) · TF-TCAF-Crediting to confirm (15)",
           size=10, color=MUTED)

    s = d.content("F · Milestones", "Twelve milestones, four done")
    d.table(s, ["Code", "Milestone", "Date", "Status"], [
        ["3a_m1 / 3b_m1", "Excise Diagnostic and Dashboard MVPs", "25 Sep", "Done"],
        ["3a_m2 / 3b_m2", "Excise Diagnostic and Dashboard to Dirk", "16 / 23 Oct", "Done (early)"],
        ["1a_m1", "Excel-AI-CPAT mitigation MVP", "23 Oct", "70%"], ["2b_m1", "Zoe's FTT project", "27 Nov", "Open"],
        ["1a_m2", "Final Excel-AI-CPAT", "4 Dec", "50%"], ["1b_m1", "Final coded CPAT (Python)", "11 Dec", "10%"],
        ["2c_m1", "Pathways tool complete", "15 Jan", "Open; Edmundo on track"], ["1c_m1", "Final Shiny app", "5 Feb", "0%"],
        ["2b_m2", "FTT and MFMod working smoothly", "12 Feb", "Open"], ["2a_m1", "GHG accounting, crediting, FFI", "5 Mar", "Open"],
    ], widths=[0.17, 0.45, 0.15, 0.23], y=1700000, size=10, h_row=320000)

    s = d.content("F · Sequencing", "Week by week to the end of the contract")
    d.cards(s, [
        dict(tag="To December", color=NAVY, title="Build", bullets=[
            "26-30 Oct vacation", "2-13 Nov power (already built)", "16-27 Nov remaining mitigation", "30 Nov-4 Dec distribution (already built)",
            "7-11 Dec Python tests and docs", "15-22 Dec buffer; 23 Dec-1 Jan leave"]),
        dict(tag="January to March", color=TEAL, title="Apply", bullets=[
            "4-8 Jan pathways integration; 11-15 Jan pilots", "18 Jan-5 Feb Shiny app", "8-12 Feb FTT and MFMod",
            "15 Feb-5 Mar GHG accounting, crediting, FFI", "8-12 Mar support to Dirk", "15-19 Mar handover; 20 Mar contract ends"]),
        dict(tag="Who", color=ORANGE, title="Division of labour", bullets=[
            "Stephen: engine, GHG accounting, handovers", "Paulina: co-benefits incl. air pollution", "Kacper: air pollution and mitigation integration; bug list",
            "Dirk and others: Excise and Dashboard onward"]),
    ], h=3600000, size=10)

    s = d.content("F · Risks", "Risks and what we need")
    d.table(s, ["Objective", "Risk", "What we need"], [
        ["One", "CPAT not well validated or published; IMF may build its own; LAMBDA needs recent Excel", "CPAT team sign-off; Excel 365 review"],
        ["Two", "Accounting, crediting, FFI requirements not set; TCAF code unconfirmed", "Requirements and charge code"],
        ["Two", "Prices and policies only partly in CDPM; price database unverified; V2-V3 not agreed", "Country checks; agree V2-V3"],
        ["Three", "Handover format and owners not agreed; MVPs not validated", "Named owner per tool"],
        ["All", "No buffer in 125 days; shared AI set-up not in place", "Claude Code; enterprise AI; what to drop"],
    ], widths=[0.12, 0.55, 0.33], y=1800000, size=11, h_row=420000)
    d.text(s, M, 4700000, CW, 300000, CDPM_NOTE + ".", size=10, color=MUTED)
