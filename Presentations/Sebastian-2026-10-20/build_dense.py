"""Dense Sebastian deck (5 slides): single Python platform for fiscal policy; macro-fiscal upgrade to the new CPAT;
CPAT Excel-AI and Python; other pieces; conclusions. Style: AIClimate-FiscalTools_v2.5. Run: python3 build_dense.py"""
import sys
from pathlib import Path
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from kit import *                                                     # noqa: E402

d = Deck(str(HERE.parent / "AIClimate-FiscalTools_v2.5.pptx"), "Stephen Stretton · meeting with Sebastian · 20 Oct 2026")
CDPM = "CDPM = the model formerly called SiSePuede."


def tile(s, x, y, w, h, tag, col, title, lines, status=None, size=10):
    d.box(s, x, y, w, h, WHITE)
    d.box(s, x, y, w, 45000, col)
    d.text(s, x + 140000, y + 110000, w - 280000, 180000, tag.upper(), size=8, bold=True, color=col)
    d.text(s, x + 140000, y + 290000, w - 280000, 300000, title, size=12, bold=True, color=NAVY)
    d.bullets(s, x + 140000, y + 600000, w - 280000, h - 900000, lines, size=size, spacing=2)
    if status:
        d.text(s, x + 140000, y + h - 290000, w - 280000, 220000, status, size=9, bold=True, color=col)


# 1 ---- single Python platform -----------------------------------------------------------------
s = d.content("1 · Proposal", "One Python platform for fiscal policy, built around the new CPAT")
lw = 2900000
d.box(s, M, 1750000, lw, 3500000, NAVY)
d.text(s, M + 160000, 1870000, lw - 320000, 200000, "WHY", size=9, bold=True, color="E3A36B")
d.bullets(s, M + 160000, 2100000, lw - 320000, 3100000, [
    "One code base for climate-related fiscal policy, shared with the macro unit",
    "A PFR process: energy costs, fiscal risks, spending efficiency, tax changes",
    "One country run end to end: data in, results back on one page",
    "Replaces one-off spreadsheets; every number traceable",
    "Lite versions now; MFMod and full IMF tools as cross-checks"], size=11, color=PALE, spacing=6)
gx, gw, gh, gg = M + lw + 150000, (CW - lw - 150000 - 2 * 120000) // 3, 1690000, 120000
tiles = [
    ("Engine", NAVY, "Coded CPAT (new CPAT)", ["Prices, revenue, emissions, households", "Excel-AI v6.2 is the checked reference"], "Python final 11 Dec"),
    ("Data", BLUE, "Tax & Macro-Fiscal Dashboard", ["64 indicators, 11 blocks, 217 countries", "Now with macro-model inputs and results"], "Built; handed to Dirk"),
    ("Needs", TEAL, "Investment needs + CDPM", ["NMIM: budget need path 2025-50, who pays", "CDPM pathways (formerly SiSePuede)"], "NMIM v0.2; CDPM Jan"),
    ("Policy", ORANGE, "Pricing and tax tools", ["Price feasibility (F1-F3): tax vs feebate", "Excise, CBAM, crediting / FFI"], "v0.1 to v3.0"),
    ("Macro and debt", ORANGE, "GIMF / DIGNAD / DIGNAR-lite", ["Chosen by country type", "LIC-DSF-lite debt verdict; MFMod option"], "v0.1, 186 countries run"),
    ("Technology", TEAL, "FTT national", ["Technology diffusion on CPAT's countries", "FTT + MFMod on one platform"], "27 Nov; 12 Feb"),
]
for i, (tag, col, ttl, lines, st) in enumerate(tiles):
    x = gx + (i % 3) * (gw + gg)
    y = 1750000 + (i // 3) * (gh + gg)
    tile(s, x, y, gw, gh, tag, col, ttl, lines, st)
d.so(s, "most components exist; the platform is one repo, one guide, one calibration and one country chain joining them.", y=5420000)

# 2 ---- macro-fiscal upgrade --------------------------------------------------------------------
s = d.content("2 · The upgrade", "Macro-fiscal block: built this month, an upgrade to the new CPAT", source=CDPM)
d.flow(s, [("Dashboard", "Data in", "DCE8F5"), ("Needs (NMIM)", "Budget need", "DCE8F5"), ("Price feasibility", "Pricing revenue", "DCE8F5"),
           ("GIMF / DIGNAD / DIGNAR", "GDP, fiscal path", "DCE8F5"), ("LIC-DSF-lite", "Debt verdict", "DCE8F5"), ("Dashboard", "Results back", None)],
       y=1700000, h=780000, size=9)
d.table(s, ["Piece", "Status (20 Oct)", "Upgrade to the new CPAT"], [
    ["GIMF-lite / DIGNAD-lite", "v0.1; one shared calibration from the dashboard; routed by country type", "Replaces CPAT's fixed fiscal multipliers with a model of GDP, prices and fiscal path"],
    ["DIGNAR-lite (resource-rich)", "Built; fossil rents fall 80% by 2050 as the world decarbonises", "New: stranded fossil revenue as a fiscal risk"],
    ["Debt test (LIC-DSF-lite)", "Built; 5 of 66 LIC/LMIC ratings worsen with the programme; coefficients to check", "Replaces the planned debt equation (investment approach V2)"],
    ["Public spending efficiency", "PEFA/PIMA-mapped efficiency in the calibration; spending split from capital", "Investment approach V3; next: PIMA scores, absorptive capacity, execution"],
    ["CDPM pathways", "Edmundo on track; coded CPAT to CDPM in Dec; pilots Jan", "Physical pathways and costs feeding the needs step"],
    ["Integrated run", "186 countries: 135 GIMF, 35 DIGNAD, 16 DIGNAR; 184 solve, 2 not financeable", "One country chain; results block in the dashboard"],
], widths=[0.21, 0.42, 0.37], y=2650000, size=9, h_row=330000)
d.text(s, M, 5150000, CW, 400000, "Caveats: calibrations illustrative; pricing revenue power-only; CPAT not yet in the chain (next: feed CPAT prices and revenue to the pricing step); not a substitute for a full DSA or MFMod.",
       size=9, color=MUTED)

# 3 ---- CPAT Excel-AI and Python ---------------------------------------------------------------
s = d.content("3 · Core engine", "CPAT in Excel-AI and Python: ahead on build; Python is the work")
d.kpis(s, [("v6.2", "Excel-AI MVP, 20 Oct: mitigation, power, distribution, GDP, NDC; 56 tabs", NAVY),
           ("123", "legacy CPAT bugs fixed, none copied; fix plans for the CPAT team", ORANGE),
           ("0 errors", "quick check; 12,992 output codes stable; every figure traceable", TEAL),
           ("57 tests", "coded CPAT runs end to end; Shiny app has six sections", BLUE)], y=1700000, h=1050000)
half = (CW - 180000) // 2
tile(s, M, 2900000, half, 2250000, "Excel-AI-CPAT", NAVY, "Status and next", [
    "Power and distribution built ahead of plan (Oct, not Nov-Dec)", "Usable in 15 minutes: open, change a policy, run scenarios",
    "Next: Excel 365 review; data requests; send fix plans", "Final 4 Dec (50%); MVP due 23 Oct (70%)"], "Milestones 1a_m1, 1a_m2", size=10)
tile(s, M + half + 180000, 2900000, half, 2250000, "Coded CPAT (Python)", BLUE, "Status and next", [
    "About 20,000 lines; mitigation and distribution; legacy export", "Port the validated Excel logic; fix revenue units",
    "Feeds CDPM pathways and the fiscal platform", "Final 11 Dec (10%); Shiny app 5 Feb"], "Milestones 1b_m1, 1c_m1", size=10)

# 4 ---- other pieces ---------------------------------------------------------------------------
s = d.content("4 · Other pieces", "Everything else, in one table")
d.table(s, ["Piece", "Status", "Next / date", "Ask"], [
    ["Excise Diagnostic", "SLE, UGA, LBR; handed to Dirk 16 Oct", "Support 8-12 Mar", "Owner; extend or fold into platform?"],
    ["Tax & Macro-Fiscal Dashboard", "Handed to Dirk; now 64 indicators", "Support 8-12 Mar", "Owner"],
    ["CBAM policy tool", "Egypt v3.0; global MVP v1.0", "No milestone", "CBAM in CPAT: priority?"],
    ["Crediting, GHG accounting, FFI", "Planned", "15 Feb-5 Mar", "Charge code; requirements"],
    ["FTT national (Zoe); FTT + MFMod", "Zoe delivering", "27 Nov; 12 Feb", "MFMod in or out"],
    ["Pathways with CDPM (Edmundo)", "Meeting went well 19 Oct; repo created", "Pilots 11-15 Jan", "Agree cost split"],
    ["Adaptation", "DIGNAD-lite has adaptation capital and disasters", "Not in plan", "Now or FY28?"],
    ["Energy GP cross-support", "Not funded (16 days)", "-", "Funding"],
], widths=[0.27, 0.31, 0.17, 0.25], y=1700000, size=10, h_row=400000)
d.text(s, M, 5400000, CW, 250000, CDPM + " Full detail: Sebastian-Platforms-v2.pptx (appendix).", size=9, color=MUTED)

# 5 ---- conclusions ----------------------------------------------------------------------------
s = d.content("5 · Conclusions", "Conclusions and decisions", bg=MIST)
d.cards(s, [
    dict(tag="Conclusions", color=NAVY, title="Where we are", bullets=[
        "The new CPAT engine is ahead of plan in Excel; Python is the next push",
        "A macro-fiscal block now runs every country end to end",
        "Together they make a single Python platform for fiscal policy and a PFR process"]),
    dict(tag="Decisions for you", color=ORANGE, title="What I need", bullets=[
        "Go-ahead: the platform as the FY27 frame; the chain as the PFR process",
        "A pilot country (e.g. Uganda, Kenya, Egypt)", "Introduction to the macro unit (MFMod)",
        "What to drop to make days (no buffer to 20 Mar)"]),
    dict(tag="Next 6 weeks", color=TEAL, title="Then", bullets=[
        "Excel-AI MVP 23 Oct; final 4 Dec; Python 11 Dec", "Feed CPAT prices and revenue into the chain",
        "Check DSF coefficients; real calibrations for the pilot", "CDPM user-flow workshop (Nov)"]),
], y=1700000, h=3500000, size=11)

d.save(str(HERE / "Sebastian-Dense-v4.pptx"))
print("slides:", d.n)
