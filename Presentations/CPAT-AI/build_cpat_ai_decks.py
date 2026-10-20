"""CPAT-AI managerial decks (20 Oct 2026): one overview deck plus four short topic decks.

Uses the slide kit and theme of the Sebastian decks (../Sebastian-2026-10-20/kit.py, AIClimate-FiscalTools_v2.5).
Run from anywhere:  python3 build_cpat_ai_decks.py
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "Sebastian-2026-10-20"))
from kit import Deck, M, CW, NAVY, ORANGE, BLUE, TEAL, GOLD, MIST, MUTED, BODY, WHITE  # noqa: E402

TEMPLATE = str(HERE.parent / "AIClimate-FiscalTools_v2.5.pptx")
FOOT = "CPAT-AI · 20 Oct 2026"
SRC = "Source: CPAT-AI repo, release v6.3 and bug/issue register (20 Oct 2026)"


# ---- slides ---------------------------------------------------------------------------------------
def cover(d, title, sub):
    d.title("CPAT-AI · management briefing", title, sub, "Stephen Stretton · 20 October 2026")


def done(d):
    s = d.content("Where we are", "Mitigation: substantial progress. Distribution: done.", source=SRC)
    d.cards(s, [
        dict(tag="Mitigation", color=BLUE, title="Substantial progress",
             bullets=["Prices, demand by 16 sectors and 8 fuels, power, road transport",
                      "Revenues, CO2, other GHGs, GDP effects, carbon rates",
                      "Checked against legacy CPAT and an independent Python recomputation",
                      "Some features still being added"]),
        dict(tag="Distribution", color=TEAL, title="Done",
             bullets=["Household impacts as legacy, by decile",
                      "ASPIRE and poverty-line transfers, production-sector prices, charts",
                      "All 31 legacy distribution bugs fixed"]),
        dict(tag="Around it", color=ORANGE, title="Documentation and hand-over",
             bullets=["Quick start, equations, method notes, every formula listed",
                      "Bug register and cell-by-cell fix plans for legacy CPAT",
                      "Python model (Coded CPAT) shares the same codes"]),
    ], h=3300000)
    d.so(s, "a working, documented CPAT-AI; two co-benefit modules remain (see next steps).")
    return s


def differences(d):
    s = d.content("CPAT-AI vs legacy CPAT", "What is different for the user", source=SRC)
    d.table(s, ["", "Legacy CPAT", "CPAT-AI"], [
        ["Scenarios", "Scenario columns copied in by macro", "A scenario is a column group: copy-paste to add one"],
        ["How many", "Two at a time (baseline + policy)", "Any number: paste more groups, or batch-run and compare up to 8"],
        ["Bugs", "159 known bugs", "No known bugs: every legacy bug fixed, none copied"],
        ["Formulas", "Long, nested formulas", "Runs with LAMBDA formulae, or entirely without (four variants)"],
        ["Data", "Data selected inside the calculation sheets", "Data sheets at legacy addresses: paste a country's data, set its code"],
        ["Speed", "More than 5 s to recalculate a scenario", "Faster: Excel timings to be confirmed (TBC)"],
    ], widths=[0.16, 0.34, 0.50], size=12, h_row=420000)
    d.so(s, "same model as legacy, easier to use, extend and trust.")
    return s


def speed(d):
    s = d.content("Recalculation", "Changing a scenario: recalculation time", source="Timings to be measured in Excel 365 on a standard WB laptop")
    d.kpis(s, [(">5 s", "Legacy CPAT: recalculation after a scenario change", ORANGE),
               ("TBC", "CPAT-AI with LAMBDA formulae: after a scenario change", BLUE),
               ("TBC", "CPAT-AI without LAMBDAs: after a scenario change", TEAL)], h=1300000)
    d.bullets(s, M, 3400000, CW, 1500000, [
        ("Why it is faster:", "short formulas (max 100 characters), one formula per block."),
        ("Only the changed scenario recalculates:", "each scenario is its own column group."),
        ("To do:", "time the three cases in Excel on the same machine and fill in this slide."),
    ], size=13, spacing=8)
    return s


def bugs(d):
    s = d.content("Legacy CPAT: bugs found", "159 bugs found in legacy CPAT", source=SRC)
    d.kpis(s, [("159", "legacy bugs: 128 mitigation, 31 distribution", ORANGE),
               ("9 / 43 / 76", "mitigation bugs by priority: high / medium / low", BLUE),
               ("0", "legacy bugs copied into CPAT-AI", TEAL)], h=1100000)
    d.text(s, M, 3050000, CW, 250000, "Typical types: wrong row or flag referenced · effect applied twice · hard-coded values · inputs that are never read",
           size=12, bold=True, color=NAVY)
    d.table(s, ["Example", "What legacy does", "Effect"], [
        ["GDP effect (L24)", "Adds the GDP effect to the growth rate every year", "GDP effects compound: overstated"],
        ["Flaring (L19)", "Counts flaring CO2 as methane", "Emissions in the wrong gas"],
        ["Regulations (L87)", "Regulation share is always 0", "Regulation scenarios do nothing"],
        ["Transfers (D1)", "Targeted transfer capped at 10 points", "Wrong household compensation"],
    ], widths=[0.2, 0.45, 0.35], y=3400000, size=11, h_row=330000)
    return s


def issues(d):
    s = d.content("Legacy CPAT: issues found", "62 issues found in legacy CPAT", source=SRC)
    d.kpis(s, [("62", "legacy issues: 52 mitigation, 10 distribution", ORANGE),
               ("50", "replicated as legacy and tracked", BLUE),
               ("12", "not replicated, each by explicit decision", TEAL)], h=1100000)
    d.text(s, M, 3050000, CW, 250000, "Typical types: old data · simplifications · questionable design choices",
           size=12, bold=True, color=NAVY)
    d.table(s, ["Example", "What legacy does", "Why it matters"], [
        ["Price data (I4)", "Price data end in 2022", "Recent price changes missed"],
        ["Fuel demand (I71)", "Fuel use grows with total GDP, so population gets the income elasticity", "Demand growth overstated"],
        ["Prices (I63)", "Demand responds to real USD prices, not local-currency prices", "Devaluations have no effect"],
    ], widths=[0.2, 0.5, 0.3], y=3400000, size=11, h_row=360000)
    return s


def approach(d):
    s = d.content("Approach", "Remove bugs; log issues; do not replicate bugs", source=SRC)
    d.cards(s, [
        dict(tag="Bug", color=ORANGE, title="Legacy does not do what it means to",
             bullets=["Wrong row, double counting, a dead reference",
                      "Always fixed in CPAT-AI, never copied",
                      "Fix plans written for the legacy team"]),
        dict(tag="Issue", color=BLUE, title="Legacy does what it means to, but it is open to question",
             bullets=["Old data, simplifications, design choices",
                      "Replicated as legacy, logged and tracked",
                      "Changed only by explicit decision"]),
        dict(tag="Python", color=TEAL, title="Coded CPAT matches Excel-AI",
             bullets=["Like legacy CPAT, without any of the bugs",
                      "Same input and output codes as the Excel",
                      "Same results on the same inputs"]),
    ], h=3300000)
    d.so(s, "results stay comparable with legacy, except where legacy is simply wrong.")
    return s


def process(d):
    s = d.content("How it is built", "AI-assisted development on GitHub", source="Tools: Claude Code (main), GitHub Copilot (occasionally)")
    d.flow(s, [
        ("Repos on GitHub", "Model, scripts, documents and rules in version control", None),
        ("Claude Code", "Writes builders, formulas, checks and documents; cloud sessions on branches", None),
        ("GitHub Copilot", "Used occasionally for smaller edits", None),
        ("Checks", "Independent Python recomputation; comparison with legacy", None),
        ("Release", "Versioned workbook, register, docs; every task logged", None),
    ], y=1850000, h=1500000)
    d.bullets(s, M, 3650000, CW, 1400000, [
        ("Rules in the repo:", "AGENTS.md, NORMS.md and CAVEATS.md tell any AI tool how to work and what was done."),
        ("Reviewable:", "every change is a commit; bugs and issues are generated from one register."),
        ("Human in charge:", "AI proposes and builds; decisions on issues are mine."),
    ], size=13, spacing=6)
    return s


def todo(d):
    s = d.content("What remains", "Four pieces of work remain", source=SRC)
    d.cards(s, [
        dict(tag="Module", color=TEAL, title="Air pollution co-benefits", bullets=["Health and mortality effects of lower local pollution"]),
        dict(tag="Module", color=TEAL, title="Transport co-benefits", bullets=["Congestion, accidents, road damage"]),
        dict(tag="Code", color=BLUE, title="Alignment of code", bullets=["Excel-AI, Python and the master code list use the same codes and results"]),
        dict(tag="Data", color=ORANGE, title="Database back end", bullets=["Data outside the workbook; perhaps a second data table on SharePoint"]),
    ], h=2700000, size=12)
    return s


def repo(d):
    s = d.content("Repository", "How the CPAT-AI repo is organised", source="github.com/cpmodel/CPAT-AI")
    d.table(s, ["Folder", "What it holds", "Who uses it"], [
        ["README, AGENTS, NORMS, CAVEATS", "Who reads what; rules for people and AI; log of every task", "Everyone"],
        ["cpat_excel/", "Excel-AI: the CPAT-AI workbook, its builder, checks and documents", "Modellers, reviewers"],
        ["cpat_excel/cpat_excel_original/", "Legacy CPAT workbook: the reference (read-only)", "Everyone"],
        ["cpat_coded/", "Coded CPAT: the Python model, tests and Shiny app", "Developers"],
        ["legacy_bugfix_plans/", "Cell-by-cell fix plans for legacy CPAT", "CPAT team"],
        ["templates/, _research/", "Input template; source literature", "Everyone"],
    ], widths=[0.3, 0.5, 0.2], size=12, h_row=420000)
    d.so(s, "one repo, two projects (Excel and Python), one reference model and one task log.")
    return s


def discuss(d):
    s = d.content("For discussion", "What next, and how to work together", source=None)
    d.numbered(s, [
        ("What to share with the IMF", "Which version, documents and bug list; when; in what form"),
        ("Involve the rest of the team", "Open the repos and the process to colleagues, not just results"),
        ("Two missing modules", "Proposal: you and the STCs develop air pollution and transport co-benefits"),
        ("User interface", "Proposal: us or the IMF build the UI; or we split the UI in half"),
        ("Way of working", "One shared repo, one register, shared rules: collaborative by design"),
    ], y=1850000, step=640000, size=13)
    d.so(s, "I would like this to be a joint effort: agree owners for the modules and the UI.")
    return s


def close(d):
    s = d._slide(NAVY)
    d.text(s, M, 2600000, CW, 900000, "Thank you", size=44, bold=True, color=WHITE, font="Cambria")
    d.text(s, M, 3500000, CW, 600000, "Questions and next steps", size=16, color="C9D3E0")
    return s


# ---- decks ----------------------------------------------------------------------------------------
DECKS = {
    "CPAT-AI_0_Management_Overview.pptx": ("CPAT-AI: progress and next steps",
        "A bug-free, faster, copy-pasteable CPAT in Excel and Python",
        [done, differences, speed, bugs, issues, approach, process, todo, repo, discuss, close]),
    "CPAT-AI_1_Progress_and_Differences.pptx": ("CPAT-AI: progress and what is new",
        "Where we are, and how CPAT-AI differs from legacy CPAT", [done, differences, speed, close]),
    "CPAT-AI_2_Legacy_Bugs_and_Issues.pptx": ("Legacy CPAT: bugs and issues",
        "What we found, and how CPAT-AI treats it", [bugs, issues, approach, close]),
    "CPAT-AI_3_Process_and_Repository.pptx": ("CPAT-AI: how it is built",
        "AI-assisted development and the repository", [process, repo, close]),
    "CPAT-AI_4_Next_Steps_and_Discussion.pptx": ("CPAT-AI: next steps",
        "What remains, and how to work together", [todo, discuss, close]),
}

if __name__ == "__main__":
    for name, (title, sub, slides) in DECKS.items():
        d = Deck(TEMPLATE, FOOT)
        cover(d, title, sub)
        for f in slides:
            f(d)
        d.save(str(HERE / name))
        print(name, d.n, "slides")
