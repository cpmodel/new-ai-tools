#!/usr/bin/env python3
"""Build the interactive workbook from the global (217-country) panel.
Input : global_data_long.csv  (from build_global_indicators.py)
Output: Global_fiscal_indicators.xlsx
"""
import re
import pandas as pd, numpy as np, openpyxl, warnings
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
warnings.filterwarnings("ignore")

import os, argparse
_here = os.path.dirname(os.path.abspath(__file__))
_ap = argparse.ArgumentParser()
_ap.add_argument("--out", default=os.environ.get("OUT_DIR", os.path.join(_here, "..", "output")))
_ap.add_argument("--config", default=os.environ.get("CONFIG_DIR", os.path.join(_here, "..", "config")))
_args, _ = _ap.parse_known_args()
OUT = os.path.join(os.path.abspath(_args.out), "")
CONFIG = os.path.join(os.path.abspath(_args.config), "")
F = "Arial"; INK = "0E1B2A"; ACC = "C05621"; MUT = "5A6B7B"
HDR = PatternFill("solid", fgColor=INK); TINT = PatternFill("solid", fgColor="F0F3F5")
thin = Side(style="thin", color="C8D0D6"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

# ---------------------------------------------------------------- blocks
# Blocks and indicators are defined ONCE, in config/indicators.csv, shared with
# build_global_indicators.py. Nothing about the indicator set is hardcoded here.
_cfg = pd.read_csv(CONFIG + "indicators.csv").sort_values("order")
QUESTIONS = [tuple(r) for r in _cfg.drop_duplicates("block")[
    ["block","block_theme","block_question","block_subtitle","block_fill","block_font"]].values]
MAGNITUDE = set(_cfg.loc[_cfg.block.str.startswith(("Context","Appendix")), "block"])
APPENDIX  = set(_cfg.loc[_cfg.block.str.startswith("Appendix"), "block"])
spec = _cfg.rename(columns={"block":"Q","label":"Indicator","area":"Area",
        "direction":"Direction","source_label":"Source","headline":"Headline",
        "eval_scale":"Eval","score_flag":"Flag","indent":"Indent"})[
        ["Q","code","Indicator","Area","Direction","Source","Headline","Eval","Flag","Indent"]]
spec["Eval"] = spec.Eval.fillna("")
QORDER = [q[0] for q in QUESTIONS]
QMAP = {q[0]: q for q in QUESTIONS}
ORD = list(spec.code)

# ---------------------------------------------------------------- data
long = pd.read_csv(OUT + "global_data_long.csv")
long = long[long.indicator.isin(ORD)]
META = long.drop_duplicates("iso3").set_index("iso3")[["country","region","income","is_ccdr"]]
wide = long.pivot(index="iso3", columns="indicator", values="value")
wyr  = long.pivot(index="iso3", columns="indicator", values="year").add_suffix("__year")
W = pd.concat([wide, wyr], axis=1).reset_index()
for i, c in enumerate(["country","region","income","is_ccdr"], start=1):
    W.insert(i, c, W.iso3.map(META[c]))
W = W.reindex(columns=["iso3","country","region","income","is_ccdr"] + ORD +
              [c + "__year" for c in ORD if c + "__year" in wyr.columns]).sort_values("country")
W = W.reset_index(drop=True)
spec["Coverage"] = spec.code.map({c: int(W[c].notna().sum()) for c in ORD})

wb = openpyxl.Workbook()
def head(ws, labels, widths, row=1, h=28):
    for j, l in enumerate(labels, 1):
        c = ws.cell(row, j, l); c.font = Font(name=F, size=9, bold=True, color="FFFFFF")
        c.fill = HDR; c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for j, w_ in enumerate(widths, 1): ws.column_dimensions[get_column_letter(j)].width = w_
    ws.row_dimensions[row].height = h; ws.freeze_panes = ws.cell(row + 1, 1)

# ---- README
ws = wb.active; ws.title = "README"
for col, w_ in [("A",3),("B",30),("C",104)]: ws.column_dimensions[col].width = w_
ws["B2"] = "Global Fiscal Indicators — climate and mitigation finance"
ws["B2"].font = Font(name=F, size=16, bold=True, color=INK)
ws["B3"] = f"{W.shape[0]} countries, {len(spec)} indicators, 16 public datasets. Compiled 26 August 2026."
ws["B3"].font = Font(name=F, size=10, italic=True, color=MUT)
r = 5
for k, v in [
("What this is", f"A comparable fiscal and energy layer for all {W.shape[0]} countries in the World Bank country list. The is_ccdr column flags the 98 economies with a published Country Climate and Development Report, so the original CCDR view is one filter away."),
("How it is organised", "Nine blocks, in reading order. Four ASSESSMENT blocks come first and carry the diagnosis: fiscal space, delivery, cost of capital, and financing. Three CONTEXT blocks follow and explain it: how much fossil revenue runs through the budget, whether the country sells or buys fossil fuels, and how it generates power. Two are APPENDIX - fuel pricing and power market regulation - shown in grey as reference detail rather than part of the diagnosis; they carry no block verdict."),
("How to use it", "Dashboard: pick a country in the yellow cell. Everything resolves - evaluation, score, value, observation year, coverage - grouped and coloured by block. Data_wide is the cross-section, Data_long names the source file for every value, Scores holds the 1-10 rescaling, Spec is the indicator list, Sources lists the datasets."),
("Reading the rows", "Indented italic rows are components of the line above: subsidies split into explicit and implicit, net fuel trade into exports and imports, fossil rents into oil, gas and coal, and the headline WACC into its cross-check sources. They are scored individually but they are not independent of their parent."),
("Evaluation and score", "Every indicator carries a plain-word Evaluation beside a 1-10 Score. Good / Medium / Bad where a direction exists. High / Medium / Low where the number is a magnitude rather than a verdict. High or Low Exporter / Importer for net fuel trade, read off the value. Thresholds: 7 and above, 4 to 7, below 4."),
("Context blocks are magnitudes", "The three Context blocks use High / Medium / Low, not Good / Bad. A country reading High on Context 2 has a large extraction and export position - read as opportunity that is revenue it could expand, read as exposure it is revenue the transition puts at risk. The file does not pick one."),
("Do not build an index", "Block averages are given, but do not average across blocks into a single country score. There are no defensible weights between them, and it would erase the triage: a country weak on Fiscal needs different instruments from one weak on PIMA."),
("Non-negotiables", "Nothing is imputed; missing stays missing. Every value carries its observation year, because vintages span five years - fiscal core 2024, capex execution mostly 2022, rents and WACC 2021. Scores are relative to the 217-country sample and move if the sample changes."),
("Appendix blocks", "Appendix A gives retail prices by fuel and sector, all from CPAT and all 2024. Appendix B describes how the power sector is organised. Note the vintages: the PSRT market-structure and regulator flags are 2013 and non-OECD only, so read them as structural context, not current state; RISE scores are 2023 and cover 140 countries."),
("A caveat on electricity prices", "The two electricity price rows come from the same CPAT series behind the cost-recovery indicators that were dropped from this tool on data-quality grounds. Retail prices are closer to observed than the modelled supply costs were, and they are the input CPAT uses to compute subsidies - but treat them as reference values rather than evidence."),
("Not included", "IMF PIMA and C-PIMA scores. Publication is consent-based and the published subsample self-selects toward better performers, so a scored column would understate weakness. PEFA PI-11 is the substitute. Electricity cost recovery was also dropped: the CPAT retail-price basis is not reliable enough to carry a scored column."),
]:
    a = ws.cell(r, 2, k); a.font = Font(name=F, size=10, bold=True, color=INK)
    a.alignment = Alignment(vertical="top")
    b = ws.cell(r, 3, v); b.font = Font(name=F, size=10)
    b.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 62; r += 2

# ---- Data_wide
wd = wb.create_sheet("Data_wide")
cols = list(W.columns); head(wd, cols, [10,28,22,20,9] + [15]*(len(cols)-5), h=50)
for i, row in enumerate(W.itertuples(index=False), 2):
    for j, v in enumerate(row, 1):
        c = wd.cell(i, j, None if (isinstance(v, float) and np.isnan(v)) else v)
        c.font = Font(name=F, size=9); c.border = BOX
        if isinstance(v, (int, float)) and not isinstance(v, bool) and j > 5: c.number_format = "0.00"
    wd.row_dimensions[i].height = 14
LAST = W.shape[0] + 1
CL = {c: get_column_letter(cols.index(c) + 1) for c in cols}

# ---- Scores (live formulas: 5th-95th percentile rescale, direction-aware)
sv = wb.create_sheet("Scores")
head(sv, ["iso3","country"] + ORD, [10,28] + [15]*len(ORD), h=50)
for rr, txt in {2:"P5 bound (low)", 3:"P95 bound (high)", 4:"Direction flag (1 better / -1 worse / 0 none)"}.items():
    c = sv.cell(rr, 2, txt); c.font = Font(name=F, size=8, italic=True, color=MUT); c.border = BOX
sv.cell(5, 2, "— country scores below —").font = Font(name=F, size=8, italic=True, color=ACC)
FLAG = dict(zip(spec.code, spec.Flag))
for k, code in enumerate(ORD):
    L = get_column_letter(k + 3); D = CL[code]
    sv.cell(2, k+3, f'=IFERROR(PERCENTILE(Data_wide!${D}$2:${D}${LAST},0.05),"")').number_format = "0.00"
    sv.cell(3, k+3, f'=IFERROR(PERCENTILE(Data_wide!${D}$2:${D}${LAST},0.95),"")').number_format = "0.00"
    sv.cell(4, k+3, FLAG[code])
    for rr in (2,3,4):
        c = sv.cell(rr, k+3); c.font = Font(name=F, size=8, color=MUT); c.border = BOX
SROW0 = 6
for i in range(W.shape[0]):
    rr = SROW0 + i; dr = 2 + i
    sv.cell(rr, 1, W.iloc[i]["iso3"]).font = Font(name=F, size=9)
    sv.cell(rr, 2, W.iloc[i]["country"]).font = Font(name=F, size=9)
    for k, code in enumerate(ORD):
        L = get_column_letter(k + 3); D = CL[code]
        f = (f'=IF(OR({L}$4=0,Data_wide!${D}${dr}=""),"",'
             f'IFERROR(ROUND(IF({L}$4=1,'
             f'1+9*(MIN(MAX(Data_wide!${D}${dr},{L}$2),{L}$3)-{L}$2)/({L}$3-{L}$2),'
             f'1+9*({L}$3-MIN(MAX(Data_wide!${D}${dr},{L}$2),{L}$3))/({L}$3-{L}$2)),1),""))')
        c = sv.cell(rr, k+3, f); c.font = Font(name=F, size=9); c.number_format = "0.0"; c.border = BOX
    for j in (1,2): sv.cell(rr, j).border = BOX
    sv.row_dimensions[rr].height = 14
SLAST = SROW0 + W.shape[0] - 1

# ---- Dashboard
db = wb.create_sheet("Dashboard", 1)
for col, w_ in [("A",2),("B",3),("C",40),("D",14),("E",6),("F",11),("G",7),
                ("H",18),("I",18),("J",22),("K",5),("L",7)]:
    db.column_dimensions[col].width = w_
db["B2"] = "Country profile"; db["B2"].font = Font(name=F, size=16, bold=True, color=INK)
db["B4"] = "Country"; db["B4"].font = Font(name=F, size=10, bold=True)
db["C4"] = "Kenya"; db["C4"].font = Font(name=F, size=12, bold=True, color="0000FF")
db["C4"].fill = PatternFill("solid", fgColor="FFFF00"); db["C4"].border = BOX
db["C4"].alignment = Alignment(horizontal="center")
dv = DataValidation(type="list", formula1=f"=Data_wide!$B$2:$B${LAST}", allow_blank=False)
db.add_data_validation(dv); dv.add(db["C4"])
for cell, col, lab in [("D4","A","ISO"), ("F4","C","Region")]:
    db[cell] = f'=INDEX(Data_wide!${col}$2:${col}${LAST},MATCH($C$4,Data_wide!$B$2:$B${LAST},0))'
    db[cell].font = Font(name=F, size=10, color=MUT)
db["B6"] = f'="CCDR economy: "&IF(INDEX(Data_wide!$E$2:$E${LAST},MATCH($C$4,Data_wide!$B$2:$B${LAST},0))=TRUE,"yes","no")&"   |   indicators with data: "&COUNT(F9:F{9+len(spec)+len(QUESTIONS)+2})&" of {len(spec)}"'
db["B6"].font = Font(name=F, size=10, bold=True, color=ACC)
db["B7"] = "Pick a country above. \u25cf marks the headline indicator for each block."
db["B7"].font = Font(name=F, size=9, italic=True, color=MUT)

head(db, ["","","Indicator","Evaluation","Score","Value","Year","Direction","Area",
          "Data source","Cov.","Data?"],
     [2,3,40,14,6,11,7,18,18,22,5,7], row=8, h=32)
db.freeze_panes = "A9"
r = 9
for qid, qshort, qtext, qsub, fill, fontcol in QUESTIONS:
    band = db.cell(r, 3, f"{qid}  \u2014  {qtext}")
    band.font = Font(name=F, size=11, bold=True, color=fontcol)
    n_ind = len(spec[spec.Q == qid])
    words = ("High","Medium","Low") if qid in MAGNITUDE else ("Good","Medium","Bad")
    if qid in APPENDIX:
        bev = db.cell(r, 4, "reference")
        bev.font = Font(name=F, size=9, italic=True, color=fontcol)
    else:
        bev = db.cell(r, 4, f'=IFERROR(IF(E{r}>=7,"{words[0]}",IF(E{r}>=4,"{words[1]}","{words[2]}")),"")')
        bev.font = Font(name=F, size=10, bold=True, color=fontcol)
    bev.alignment = Alignment(horizontal="center")
    avg = db.cell(r, 5, f'=IFERROR(ROUND(AVERAGE(E{r+1}:E{r+n_ind}),1),"")')
    avg.font = Font(name=F, size=11, bold=True, color=fontcol); avg.number_format = "0.0"
    avg.alignment = Alignment(horizontal="center")
    db.cell(r, 8, qshort).font = Font(name=F, size=9, bold=True, color=fontcol)
    db.cell(r, 9, qsub).font = Font(name=F, size=8, italic=True, color=fontcol)
    nb = db.cell(r, 12, f'=COUNT(F{r+1}:F{r+n_ind})&"/{n_ind}"')
    nb.font = Font(name=F, size=9, bold=True, color=fontcol)
    nb.alignment = Alignment(horizontal="center")
    for j in range(2, 13):
        db.cell(r, j).fill = PatternFill("solid", fgColor=fill); db.cell(r, j).border = BOX
    db.row_dimensions[r].height = 20; r += 1
    for row in spec[spec.Q == qid].itertuples(index=False):
        code = row.code; vc = CL[code]; yc = CL.get(code + "__year")
        m = db.cell(r, 2, "\u25cf" if row.Headline else "")
        m.font = Font(name=F, size=10, bold=True, color=fontcol)
        m.alignment = Alignment(horizontal="center")
        n = db.cell(r, 3, row.Indicator)
        n.font = Font(name=F, size=9, bold=bool(row.Headline),
                      italic=bool(row.Indent), color=MUT if row.Indent else "000000")
        n.alignment = Alignment(indent=2 if row.Indent else 0)
        sc_ = get_column_letter(ORD.index(code) + 3)
        if row.Eval == "GMB":
            ef = f'=IF(E{r}="","",IF(E{r}>=7,"Good",IF(E{r}>=4,"Medium","Bad")))'
        elif row.Eval == "HML":
            ef = f'=IF(E{r}="","",IF(E{r}>=7,"High",IF(E{r}>=4,"Medium","Low")))'
        elif row.Eval == "STRUCT":
            ef = (f'=IF(F{r}="","",IF(F{r}=3,"Wholesale market",'
                  f'IF(F{r}=2,"Single buyer","Vertically integrated")))')
        elif row.Eval == "YESNO":
            ef = f'=IF(F{r}="","",IF(F{r}=1,"Yes","No"))'
        elif row.Eval == "TRADE":
            ef = (f'=IF(F{r}="","",IF(F{r}>20,"High Exporter",IF(F{r}>0,"Low Exporter",'
                  f'IF(F{r}>=-20,"Low Importer","High Importer"))))')
        else:
            ef = None
        if ef:
            ev = db.cell(r, 4, ef); ev.font = Font(name=F, size=9, bold=True, color=fontcol)
            ev.alignment = Alignment(horizontal="center")
        s_ = db.cell(r, 5, f'=IFERROR(INDEX(Scores!${sc_}${SROW0}:${sc_}${SLAST},MATCH($C$4,Scores!$B${SROW0}:$B${SLAST},0)),"")')
        s_.font = Font(name=F, size=10, color=MUT); s_.number_format = "0.0"
        s_.alignment = Alignment(horizontal="center")
        # INDEX on an EMPTY cell returns 0, not blank - so a missing value would
        # display as 0.00 and read as present. Test for blank first, then fetch.
        _ix = f'INDEX(Data_wide!${vc}$2:${vc}${LAST},MATCH($C$4,Data_wide!$B$2:$B${LAST},0))'
        v = db.cell(r, 6, f'=IFERROR(IF({_ix}="","",{_ix}),"")')
        v.font = Font(name=F, size=10, bold=True); v.number_format = "0.00"
        v.alignment = Alignment(horizontal="center")
        if yc:
            _iy = f'INDEX(Data_wide!${yc}$2:${yc}${LAST},MATCH($C$4,Data_wide!$B$2:$B${LAST},0))'
            y = db.cell(r, 7, f'=IFERROR(IF({_iy}="","",{_iy}),"")')
            y.font = Font(name=F, size=9, color=MUT); y.alignment = Alignment(horizontal="center")
        d = db.cell(r, 8, row.Direction)
        d.font = Font(name=F, size=8, color=ACC if row.Direction.startswith(("+","Not")) else MUT)
        db.cell(r, 9, row.Area).font = Font(name=F, size=8, color=MUT)
        sr = db.cell(r, 10, row.Source); sr.font = Font(name=F, size=8, color=MUT)
        sr.alignment = Alignment(wrap_text=False)
        cc = db.cell(r, 11, row.Coverage); cc.font = Font(name=F, size=8, color=MUT)
        cc.alignment = Alignment(horizontal="center")
        dp = db.cell(r, 12, f'=IF(F{r}="","-","Yes")')
        dp.font = Font(name=F, size=8, bold=True, color=MUT)
        dp.alignment = Alignment(horizontal="center")
        for j in range(2, 13): db.cell(r, j).border = BOX
        for j in (2, 3): db.cell(r, j).fill = PatternFill("solid", fgColor=fill)
        db.row_dimensions[r].height = 14; r += 1
db.cell(r+1, 3, "Data? = whether this country has a value. Cov. = how many of the 217 countries have one. Nothing is imputed.").font = Font(name=F, size=8, italic=True, color=MUT)
db.cell(r+2, 3, "Context blocks read High / Medium / Low (a magnitude). Assessment blocks read Good / Medium / Bad (a verdict).").font = Font(name=F, size=8, italic=True, color=ACC)

# ---- Spec
sp = wb.create_sheet("Spec")
head(sp, ["Block","Theme","Question","Indicator","Column name","Area","Direction","Evaluation scale","Coverage","Source"],
     [11,24,40,44,40,26,28,16,10,38])
EV = {"GMB":"Good / Medium / Bad", "HML":"High / Medium / Low", "TRADE":"Exporter / Importer",
      "STRUCT":"Market model", "YESNO":"Yes / No", "":"not evaluated"}
for i, row in enumerate(spec.itertuples(index=False), 2):
    q = QMAP[row.Q]
    for j, v in enumerate([row.Q, q[1], q[2], row.Indicator, row.code, row.Area, row.Direction,
                           EV[row.Eval], row.Coverage, row.Source], 1):
        c = sp.cell(i, j, v); c.font = Font(name=F, size=9, bold=(j == 4 and bool(row.Headline)),
                                            italic=(j == 4 and bool(row.Indent)))
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="top",
                                indent=2 if (j == 4 and row.Indent) else 0)
        c.fill = PatternFill("solid", fgColor=q[4])
    sp.row_dimensions[i].height = 24

# ---- Sources
# Source metadata lives in config/sources.csv; the coverage column is derived from
# the panel itself rather than typed in, so it cannot drift out of date.
_srcmeta = pd.read_csv(CONFIG + "sources.csv")
_cov = (long.groupby("source").iso3.nunique()
            .rename("n").reset_index())
def _cov_for(key):
    key = "" if pd.isna(key) else str(key).strip()
    if not key: return "not used"
    hits = _cov[_cov.source.astype(str).str.contains(re.escape(key), case=False, regex=True, na=False)]
    return f"{int(hits.n.max())}" if len(hits) else "-"
SRC = [(r.file, r.organisation, r.contents, r.used_for, _cov_for(r.match_key), r.url)
       for r in _srcmeta.itertuples(index=False)]

sc = wb.create_sheet("Sources")
head(sc, ["File used","Source (organisation / publication)","What it contains","Used for","Cov.","URL"],
     [44,40,52,42,10,56])
for i, row in enumerate(SRC, 2):
    for j, v in enumerate(row, 1):
        c = sc.cell(i, j, v); c.font = Font(name=F, size=9); c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if i % 2 == 0: c.fill = TINT
    sc.row_dimensions[i].height = 34

# ---- Data_long
dl = wb.create_sheet("Data_long")
L = long[["iso3","country","region","income","is_ccdr","indicator","value","year","source"]].copy()
L["indicator"] = pd.Categorical(L.indicator, ORD, ordered=True)
L = L.sort_values(["country","indicator"])
head(dl, list(L.columns), [10,26,22,20,9,40,13,16,46])
for i, row in enumerate(L.itertuples(index=False), 2):
    for j, v in enumerate(row, 1):
        c = dl.cell(i, j, None if (isinstance(v, float) and np.isnan(v)) else
                    (str(v) if j == 6 else v))
        c.font = Font(name=F, size=9); c.border = BOX
    dl.row_dimensions[i].height = 13

wb.save(OUT + "Global_fiscal_indicators.xlsx")
print(f"saved | {W.shape[0]} countries | {len(spec)} indicators | {len(QUESTIONS)} blocks | {len(L)} long rows")
