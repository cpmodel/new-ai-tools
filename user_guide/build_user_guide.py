"""Build the two AI quick-start guides (Word, PowerPoint, Markdown, prompts) and the zip.

One source (GUIDES below) -> every output, so they never drift apart.
Run from anywhere:  python user_guide/build_user_guide.py
Needs: python-docx, python-pptx.
"""
import os
import zipfile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm
from pptx import Presentation
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt as PPt

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
DATE = "October 2026"
ZIP_NAME = "AI_Tools_Quick_Start.zip"
NAVY, BLUE, GOLD, GREY, LIGHT, WARN, CODE = "1F3864", "2E75B6", "C9A227", "404040", "EEF3F9", "FDECEA", "F2F2F2"

# --------------------------------------------------------------------------- content
GUIDES = [
    dict(
        folder="1_Excise_Diagnostic_Claude",
        name="Excise_Diagnostic_Quick_Start",
        title="Excise diagnostic for Indonesia",
        tool="Claude (deep research)",
        goal="Claude researches Indonesia's excise laws and writes the diagnostic: a Word report and an Excel workbook.",
        assumed=[
            "The new-ai-tools repo is downloaded on your laptop and up to date.",
            "You have a Claude account with Projects and Research (deep research) switched on.",
        ],
        warn=("Deep research uses a lot of credits.",
              "Go one prompt at a time. Read each answer before you send the next. "
              "If an answer goes wrong, stop and fix it before going on. "
              "Never ask for the whole report in one go."),
        steps=[
            ("Make a Project", "In Claude: Projects > New project. Call it \"Excise - Indonesia\"."),
            ("Add two files", "excise_diagnostic_process.md (in the repo: ExciseDiagnostic/process_files/) "
                              "and INDONESIA_INPUTS.md (in this folder)."),
            ("Start a chat", "New chat inside the Project. Switch on Research."),
            ("Paste Prompt 1", "Read the answer. Is every rate quoted, with its date and a link? "
                               "Then paste Prompt 2, and so on to Prompt 6."),
            ("Save the results", "Put the Word report and the Excel file in "
                                 "ExciseDiagnostic/country_results/Indonesia/."),
        ],
        prompts=[
            ("Prompt 1 - the laws",
             "Use the excise diagnostic process file and INDONESIA_INPUTS.md. Run stage A1 for "
             "Indonesia, 2020-01-01 to 2026-09-30: find the excise laws and regulations and list "
             "every rate as printed, with its article, date and a link. If you cannot get a source, "
             "give me a list of links to download. Stop after A1."),
            ("Prompt 2 - the schedule", "Run stage A2. Stop after A2."),
            ("Prompt 3 - the numbers and workbook", "Run stage A3. No comparator. Stop after A3."),
            ("Prompt 4 - the findings", "Run stage A4. Stop after A4."),
            ("Prompt 5 - the review", "Run stage A5. Stop after A5."),
            ("Prompt 6 - the final files",
             "Run stage A6. Give me the report as a Word file and the workbook as an Excel file."),
        ],
        extra=("Subsidy diagnostic instead?",
               "Use the subsidy prompt at the end of PROMPTS.md in place of Prompt 1. "
               "There is no subsidy process file yet, so treat the result as a first draft."),
        extra_prompt=("Subsidy - Prompt 1",
             "Using the structure and quality rules of the excise diagnostic process file, start a "
             "fossil fuel SUBSIDY diagnostic for Indonesia at 2026-09-30. Step 1 only: list each "
             "fuel and electricity subsidy and compensation payment, its legal basis, and the "
             "amounts budgeted and spent for the last five years (IDR and % of GDP), with a link "
             "for every source. Mark gaps data_gap; never guess. Stop after step 1 and suggest "
             "the next steps."),
        check_title="What good looks like",
        check=[("Rates", "quoted as printed, with date"), ("Sources", "a working link for each"),
               ("Gaps", "marked data_gap, never guessed"), ("Bills", "pending ones listed, not applied")],
        zip_extra={"ExciseDiagnostic/process_files/excise_diagnostic_process.md": "excise_diagnostic_process.md"},
    ),
    dict(
        folder="2_Tax_Indicators_Copilot",
        name="Tax_Indicators_Quick_Start",
        title="Add tax indicators to the dashboard",
        tool="GitHub Copilot",
        goal="Copilot adds four rows under Tax revenue in the Global Fiscal Indicators dashboard: "
             "personal income tax, corporate income tax, VAT and excises (% of GDP). "
             "The data are already in the repo.",
        assumed=[
            "The new-ai-tools repo is downloaded on your laptop and up to date.",
            "GitHub Copilot is installed and signed in (if not: Presentations/GitHub_Copilot_AI_User_Guide.pptx).",
            "Python is set up for the dashboard (TaxMacroFiscalDashboard/build/0_setup.bat, run once).",
        ],
        warn=None,
        steps=[
            ("Open the repo", "Open the new-ai-tools folder in VS Code. Click Sync to get the latest."),
            ("Open Copilot Chat", "Pick Agent mode and the model Claude Sonnet."),
            ("Paste Prompt 1", "Copilot reads the files and tells you its plan. Nothing changes yet."),
            ("Paste Prompt 2", "Copilot makes the change and rebuilds. Click Keep. Check the numbers (last page)."),
            ("Paste Prompt 3", "Copilot saves version 1.4 and pushes it to GitHub. Done."),
        ],
        prompts=[
            ("Prompt 1 - plan",
             "Read TaxMacroFiscalDashboard/build/README.md. I want four new indicators in the Fiscal "
             "block, indented under tax_revenue_pct_gdp, from the IMF WoRLD file already in raw_data "
             "(world-imf2026.xlsx, already % of GDP): pit_revenue_pct_gdp = TaxIncI, "
             "cit_revenue_pct_gdp = TaxIncC, vat_revenue_pct_gdp = TaxSalG, "
             "excise_revenue_pct_gdp = TaxSalExc. Labels start \"of which:\". "
             "Tell me your plan in a short list. Do not edit yet."),
            ("Prompt 2 - do it",
             "Go ahead: change the Python script, the R script and config/indicators.csv as planned. "
             "Then rebuild (build_global_indicators.py, then build_dashboard.py) and show me "
             "Indonesia's five tax rows."),
            ("Prompt 3 - save",
             "Save this as version 1.4 following VersionNotes.md. Then commit and push."),
        ],
        extra=None, extra_prompt=None,
        check_title="Check: Indonesia, % of GDP (2024)",
        check=[("Tax revenue", "10.08"), ("Personal income tax", "1.79"), ("Corporate income tax", "2.65"),
               ("VAT", "3.52"), ("Excises", "1.11")],
        zip_extra={},
    ),
]


# --------------------------------------------------------------------------- markdown
def build_md(g, d):
    L = [f"# {g['title']} - with {g['tool']}", "", g["goal"], "", "## What you need (we assume you have this)", ""]
    L += [f"- {x}" for x in g["assumed"]]
    if g["warn"]: L += ["", f"> **{g['warn'][0]}** {g['warn'][1]}"]
    L += ["", "## Steps", ""] + [f"{i}. **{a}.** {b}" for i, (a, b) in enumerate(g["steps"], 1)]
    L += ["", "## Prompts (copy and paste, one at a time)", ""]
    for a, b in g["prompts"]: L += [f"**{a}**", "", "```text", b, "```", ""]
    L += [f"## {g['check_title']}", "", "| | |", "|---|---|"] + [f"| {a} | {b} |" for a, b in g["check"]]
    if g["extra"]:
        L += ["", f"## {g['extra'][0]}", "", g["extra"][1], "", f"**{g['extra_prompt'][0]}**", "",
              "```text", g["extra_prompt"][1], "```"]
    with open(os.path.join(d, "README.md"), "w", encoding="utf-8") as fh: fh.write("\n".join(L) + "\n")
    P = [f"# Prompts - {g['title']}", "", "Copy and paste one at a time. Read each answer first.", ""]
    for a, b in g["prompts"] + ([g["extra_prompt"]] if g["extra_prompt"] else []):
        P += [f"## {a}", "", "```text", b, "```", ""]
    with open(os.path.join(d, "PROMPTS.md"), "w", encoding="utf-8") as fh: fh.write("\n".join(P))


# --------------------------------------------------------------------------- word
def _shade(p, fill):
    pr = p._p.get_or_add_pPr(); shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill); pr.append(shd)


def _num(doc, i, a, b, color=None):
    p = doc.add_paragraph(); f = p.paragraph_format
    f.left_indent = Cm(0.7); f.first_line_indent = Cm(-0.7); f.space_after = Pt(4)
    r = p.add_run(f"{i}\t{a}" + (". " if b else "")); r.bold = True
    if color: r.font.color.rgb = RGBColor.from_string(color)
    if b: p.add_run(b)


def build_docx(g, path):
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    for s in ["Title", "Heading 1"]:
        doc.styles[s].font.name = "Arial"; doc.styles[s].font.color.rgb = RGBColor.from_string(NAVY)
    for sec in doc.sections:
        sec.left_margin = sec.right_margin = Cm(2.2); sec.top_margin = sec.bottom_margin = Cm(1.8)
    t = doc.add_paragraph(g["title"], style="Title"); t.runs[0].font.size = Pt(24)
    p = doc.add_paragraph(); r = p.add_run(f"Quick start with {g['tool']}  ·  {DATE}"); r.italic = True
    r.font.color.rgb = RGBColor.from_string(GREY)
    doc.add_paragraph(g["goal"])

    doc.add_heading("What you need (we assume you have this)", 1)
    for x in g["assumed"]:
        p = doc.add_paragraph(); p.add_run("✓  ").bold = True; p.add_run(x)
    if g["warn"]:
        p = doc.add_paragraph(); _shade(p, WARN)
        p.add_run("⚠  " + g["warn"][0] + " ").bold = True; p.add_run(g["warn"][1])

    doc.add_heading("Steps", 1)
    for i, (a, b) in enumerate(g["steps"], 1): _num(doc, f"{i}.", a, b)

    doc.add_heading("Prompts (copy and paste, one at a time)", 1)
    for a, b in g["prompts"]:
        doc.add_paragraph().add_run(a).bold = True
        p = doc.add_paragraph(); _shade(p, CODE); r = p.add_run(b); r.font.name = "Consolas"; r.font.size = Pt(10)

    doc.add_heading(g["check_title"], 1)
    tb = doc.add_table(rows=0, cols=2); tb.style = "Table Grid"
    for a, b in g["check"]:
        c = tb.add_row().cells; c[0].text = a; c[1].text = b; c[0].paragraphs[0].runs[0].bold = True
        c[0].width = Cm(5); c[1].width = Cm(9)
    if g["extra"]:
        doc.add_heading(g["extra"][0], 1); doc.add_paragraph(g["extra"][1])
        p = doc.add_paragraph(); _shade(p, CODE)
        r = p.add_run(g["extra_prompt"][1]); r.font.name = "Consolas"; r.font.size = Pt(10)
    doc.save(path)


# --------------------------------------------------------------------------- powerpoint
def _rgb(h): return PRGB.from_string(h)


def _text(s, x, y, w, h, text, size=14, bold=False, color=NAVY, font="Arial", align=PP_ALIGN.LEFT):
    tf = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
    tf.word_wrap = True; tf.margin_left = tf.margin_right = Inches(0.05)
    for i, ln in enumerate(text if isinstance(text, list) else [text]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); p.alignment = align
        p.space_after = PPt(6)
        r = p.add_run(); r.text = ln; r.font.size = PPt(size); r.font.bold = bold
        r.font.color.rgb = _rgb(color); r.font.name = font


def _rect(s, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
    r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    r.fill.solid(); r.fill.fore_color.rgb = _rgb(fill); r.line.fill.background(); r.shadow.inherit = False
    return r


def _badge(s, x, y, label, fill=BLUE, d=0.5):
    c = _rect(s, x, y, d, d, fill, MSO_SHAPE.OVAL); tf = c.text_frame; tf.margin_left = tf.margin_right = 0
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER; r = p.add_run(); r.text = label
    r.font.size = PPt(15); r.font.bold = True; r.font.color.rgb = _rgb("FFFFFF"); r.font.name = "Arial"


def _slide(prs, g, kicker, title):
    s = prs.slides.add_slide(prs.slide_layouts[6]); _rect(s, 0, 0, 13.333, 0.12, NAVY)
    _text(s, 0.7, 0.4, 12, 0.35, kicker.upper(), 12, True, GOLD)
    _text(s, 0.7, 0.75, 12, 0.8, title, 30, True, NAVY)
    _text(s, 0.7, 7.0, 12, 0.3, f"{g['title']}  ·  quick start with {g['tool']}  ·  {DATE}", 9, False, "808080")
    return s


def build_pptx(g, path):
    prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)

    # 1 cover: goal + what you need (+ credits warning)
    s = prs.slides.add_slide(prs.slide_layouts[6]); _rect(s, 0, 0, 5.2, 7.5, NAVY)
    _text(s, 0.6, 0.8, 4.2, 0.4, "QUICK START", 13, True, GOLD)
    _text(s, 0.6, 1.5, 4.2, 2.2, g["title"], 34, True, "FFFFFF")
    _text(s, 0.6, 3.9, 4.2, 0.5, f"with {g['tool']}", 18, False, "D9E2F3")
    _text(s, 0.6, 4.7, 4.2, 2.0, g["goal"], 13, False, "D9E2F3")
    _text(s, 5.8, 0.8, 7, 0.5, "What you need (we assume you have this)", 20, True, NAVY)
    y = 1.5
    for x in g["assumed"]:
        _badge(s, 5.8, y, "✓", "3A7D44", 0.45); _text(s, 6.45, y + 0.02, 6.3, 0.8, x, 15, False, GREY); y += 0.95
    if g["warn"]:
        _rect(s, 5.8, y + 0.2, 6.9, 2.0, WARN); _rect(s, 5.8, y + 0.2, 0.1, 2.0, "C0392B")
        _text(s, 6.1, y + 0.35, 6.5, 0.5, "⚠  " + g["warn"][0], 18, True, "C0392B")
        _text(s, 6.1, y + 0.9, 6.5, 1.3, g["warn"][1], 13, False, GREY)

    # 2 steps
    s = _slide(prs, g, "Steps", f"{len(g['steps'])} steps")
    for i, (a, b) in enumerate(g["steps"]):
        y = 1.8 + i * 0.98
        _badge(s, 0.7, y, str(i + 1)); _text(s, 1.4, y - 0.02, 11, 0.4, a, 18, True, NAVY)
        _text(s, 1.4, y + 0.38, 11.2, 0.6, b, 13, False, GREY)

    # 3 prompts
    s = _slide(prs, g, "Prompts", "Copy and paste, one at a time")
    long_ = len(g["prompts"]) <= 3
    y = 1.75
    for a, b in g["prompts"]:
        h = (0.5 + 0.25 * (len(b) // 95)) if long_ else (0.42 if len(b) < 90 else 1.25)
        _text(s, 0.7, y, 3.0, 0.4, a, 13, True, BLUE)
        _rect(s, 3.7, y - 0.05, 9.0, h, CODE)
        _text(s, 3.8, y, 8.8, h, b, 12, False, "262626", "Consolas")
        y += h + 0.2

    # 4 check
    s = _slide(prs, g, "Check", g["check_title"])
    for i, (a, b) in enumerate(g["check"]):
        y = 1.8 + i * 0.75
        _rect(s, 0.7, y, 7.4, 0.62, LIGHT)
        _text(s, 0.9, y + 0.1, 3.3, 0.45, a, 16, True, NAVY); _text(s, 4.3, y + 0.1, 3.7, 0.45, b, 16, False, GREY)
    if g["extra"]:
        _rect(s, 8.6, 1.8, 4.1, 3.0, "FFF6DB")
        _text(s, 8.8, 1.95, 3.8, 0.5, g["extra"][0], 16, True, NAVY)
        _text(s, 8.8, 2.55, 3.8, 2.2, g["extra"][1], 13, False, GREY)
    prs.save(path)


# --------------------------------------------------------------------------- top README + zip
def build_index():
    L = ["# AI Tools Quick Start", "", "Two separate guides. Pick the one you need.", "",
         "| Guide | Tool | Folder |", "|---|---|---|"]
    for g in GUIDES:
        L.append(f"| {g['title']} | {g['tool']} | [`{g['folder']}/`]({g['folder']}/) |")
    L += ["", f"Each folder has the guide as Word, PowerPoint and Markdown, plus `PROMPTS.md` to copy from. "
          f"Everything is also in [`{ZIP_NAME}`]({ZIP_NAME}).", "",
          "To change the guides, edit `build_user_guide.py` and run `python user_guide/build_user_guide.py`.", ""]
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write("\n".join(L))


if __name__ == "__main__":
    build_index()
    with zipfile.ZipFile(os.path.join(HERE, ZIP_NAME), "w", zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(HERE, "README.md"), "AI_Tools_Quick_Start/README.md")
        for g in GUIDES:
            d = os.path.join(HERE, g["folder"]); os.makedirs(d, exist_ok=True)
            build_md(g, d)
            build_docx(g, os.path.join(d, g["name"] + ".docx"))
            build_pptx(g, os.path.join(d, g["name"] + ".pptx"))
            for f in sorted(os.listdir(d)):
                z.write(os.path.join(d, f), f"AI_Tools_Quick_Start/{g['folder']}/{f}")
            for src, dst in g["zip_extra"].items():
                z.write(os.path.join(REPO, src), f"AI_Tools_Quick_Start/{g['folder']}/{dst}")
    print("built", ", ".join(g["folder"] for g in GUIDES), "and", ZIP_NAME)
